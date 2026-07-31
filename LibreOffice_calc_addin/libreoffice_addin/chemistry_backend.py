"""
Chemistry backend CLI for the LibreOffice Calc "Chemistry" add-in.

Why a separate process instead of importing RDKit inside LibreOffice's
bundled Python: LibreOffice ships its own internal Python interpreter, and
getting a compiled package like RDKit installed into *that* interpreter is
fragile and version-locked. Instead, the Calc-side macro (chemistry_calc.py,
which runs under LibreOffice's Python and has zero third-party dependencies)
shells out to this script, run by a regular system Python that has RDKit
installed -- the same way the Excel version hands off chemistry work from
VBA to Python via xlwings.

This script is intentionally self-contained (no imports from the Excel
add-in folder) so the libreoffice_addin/ folder can be copied/zipped and
handed to someone else on its own.

Usage (all communication after the first two argv tokens happens as JSON on
stdin/stdout, so no argument ever has to survive shell quoting):

    python chemistry_backend.py load_sdf <path-to-sdf>
        -> stdout: {"headers": [...], "rows": [[...], ...], "ctabs": [...]}
        ctabs is a parallel array (same order/length as rows) holding each
        record's original CTAB text, for later use by render_structures and
        view_structure -- it preserves the source file's own 2D layout
        instead of RDKit recomputing one from SMILES.

    python chemistry_backend.py write_sdf <path-to-sdf>
        <- stdin:  {"headers": [...], "rows": [[...], ...]}
        -> stdout: {"written": N, "skipped": N}

    python chemistry_backend.py strip_salts
        <- stdin:  {"smiles": [...]}
        -> stdout: {"parents": [...]}   (same length/order as input; null on failure)

    python chemistry_backend.py add_column <descriptor-name>
        <- stdin:  {"smiles": [...]}
        -> stdout: {"values": [...]}    (same length/order as input; null on failure)

    python chemistry_backend.py list_descriptors
        -> stdout: {"names": [...]}

    python chemistry_backend.py render_structures
        <- stdin:  {"ctabs": [...]}
        -> stdout: {"images": [path-or-null, ...], "tmp_dir": "..."}
        One small 2D structure SVG per input CTAB, same length/order as
        input, null for anything that fails to parse. All SVGs live in a
        single fresh temp directory (tmp_dir); the caller is responsible
        for deleting it once consumed (e.g. inserted into Calc), since this
        process can't know when that's safe to do.

    python chemistry_backend.py view_structure
        <- stdin:  {"ctab": "...", "smiles": "..."}
        -> stdout: {"image": "path", "tmp_dir": "..."}
        Renders one large structure SVG for the "View Structure" popup
        (opened in the system's default SVG viewer, usually a browser).
        Prefers ctab (preserves the original layout); falls back to smiles
        with a freshly computed layout if ctab is missing/unparseable.

On error, prints {"error": "..."} to stdout and exits with status 1, so the
caller never has to parse stderr to show the user something useful.
"""
from __future__ import annotations

import json
import os
import sys
import tempfile
from typing import List, Optional, Tuple

from rdkit import Chem, RDLogger
from rdkit.Chem import Descriptors
from rdkit.Chem.Draw import rdMolDraw2D
from rdkit.Chem.SaltRemover import SaltRemover

RDLogger.DisableLog("rdApp.*")

# ---------------------------------------------------------------------------
# Descriptor registry used by "Add Column" (kept in the same order/names as
# the Excel add-in so sheets produced by either tool line up).
# ---------------------------------------------------------------------------
DESCRIPTORS = {
    "MolWt": Descriptors.MolWt,
    "LogP": Descriptors.MolLogP,
    "TPSA": Descriptors.TPSA,
    "HBD": Descriptors.NumHDonors,
    "HBA": Descriptors.NumHAcceptors,
    "NumRings": Descriptors.RingCount,
    "RotatableBonds": Descriptors.NumRotatableBonds,
    "HeavyAtomCount": Descriptors.HeavyAtomCount,
    "FractionCSP3": Descriptors.FractionCSP3,
}

_SALT_REMOVER = SaltRemover()  # RDKit's default salt/solvent dictionary

# Name of the hidden column chemistry_calc.py uses to stash each row's
# original CTAB (see _read_sdf_to_table). Shared here so write_sdf never
# accidentally reinjects it as a bogus SD property tag.
CTAB_COLUMN_NAME = "_CTAB"


# ---------------------------------------------------------------------------
# Pure chemistry helpers (identical in spirit to excel_addin/chemistry_addin.py)
# ---------------------------------------------------------------------------
def _read_sdf_to_table(filepath: str) -> Tuple[List[str], List[List[str]], List[str]]:
    """Parse an SD file into (headers, rows, ctabs). First column is always
    SMILES. ctabs is a parallel list (same order/length as rows) holding
    each record's original CTAB (molfile block) -- SDMolSupplier preserves
    the file's own atom coordinates rather than recomputing a layout, so
    re-serializing via MolToMolBlock reproduces the original 2D depiction.
    This is kept out of headers/rows (not a chemical "property") so it
    doesn't need special-casing in the generic property-column logic; the
    caller decides where to stash it."""
    supplier = Chem.SDMolSupplier(filepath)
    headers: List[str] = ["SMILES"]
    seen = {"SMILES"}
    records = []
    ctabs: List[str] = []

    for mol in supplier:
        if mol is None:
            continue
        row = {"SMILES": Chem.MolToSmiles(mol)}
        for prop in mol.GetPropNames():
            if prop.startswith("_"):
                continue
            if prop not in seen:
                seen.add(prop)
                headers.append(prop)
            row[prop] = mol.GetProp(prop)
        records.append(row)
        try:
            ctabs.append(Chem.MolToMolBlock(mol))
        except Exception:  # noqa: BLE001 - a bad conformer shouldn't sink the whole load
            ctabs.append("")

    rows = [[rec.get(h, "") for h in headers] for rec in records]
    return headers, rows, ctabs


def _table_to_sdf(headers: List[str], rows: List[List[str]], filepath: str) -> Tuple[int, int]:
    """Write a (headers, rows) table back out as an SD file.

    Column named "SMILES" (case-insensitive) supplies connectivity; every
    other non-empty cell becomes an SD property tag on that record, except
    CTAB_COLUMN_NAME ("_CTAB"), the hidden column chemistry_calc.py uses to
    cache each row's original CTAB for "View Structure" -- that's raw
    structural data, not a property, and reinjecting a multi-line molblock
    as a property tag would produce a malformed/bloated SD file.
    Returns (written_count, skipped_count).
    """
    lower_headers = [str(h).strip().lower() for h in headers]
    if "smiles" not in lower_headers:
        raise ValueError('No "SMILES" column found -- cannot rebuild structures.')
    smiles_idx = lower_headers.index("smiles")
    excluded_idx = {smiles_idx}
    for i, header in enumerate(headers):
        if str(header).strip() == CTAB_COLUMN_NAME:
            excluded_idx.add(i)

    writer = Chem.SDWriter(filepath)
    written, skipped = 0, 0
    for row in rows:
        smi = (row[smiles_idx] or "").strip() if row[smiles_idx] is not None else ""
        if not smi:
            skipped += 1
            continue
        mol = Chem.MolFromSmiles(smi)
        if mol is None:
            skipped += 1
            continue
        for i, header in enumerate(headers):
            if i in excluded_idx:
                continue
            val = row[i]
            if val is None or str(val).strip() == "":
                continue
            mol.SetProp(str(header), str(val))
        writer.write(mol)
        written += 1
    writer.close()
    return written, skipped


def _strip_salt_from_smiles(smiles: str) -> Optional[str]:
    """Return the parent (largest organic fragment) SMILES, or None on failure."""
    mol = Chem.MolFromSmiles(smiles)
    if mol is None:
        return None
    parent = _SALT_REMOVER.StripMol(mol, dontRemoveEverything=True)
    if parent is None or parent.GetNumAtoms() == 0:
        return smiles  # nothing left after stripping -- keep original
    return Chem.MolToSmiles(parent)


def _compute_descriptor(smiles: str, prop_name: str):
    """Compute one named RDKit descriptor for a SMILES string."""
    if prop_name not in DESCRIPTORS:
        valid = ", ".join(sorted(DESCRIPTORS))
        raise ValueError(f'Unknown property "{prop_name}". Valid options: {valid}')
    mol = Chem.MolFromSmiles(smiles)
    if mol is None:
        return None
    return round(DESCRIPTORS[prop_name](mol), 3)


_THUMBNAIL_SVG_SIZE = (220, 160)  # px viewBox; Calc displays these at a fixed mm size regardless
_LARGE_SVG_SIZE = (900, 700)  # px viewBox for the external "View Structure" popup


def _mol_from_ctab_or_smiles(ctab: str, smiles: str):
    """Prefer the original CTAB (preserves the source file's own 2D layout);
    fall back to a freshly-computed layout from SMILES if no CTAB is
    available or it fails to parse. Returns an RDKit Mol or None."""
    if ctab:
        mol = Chem.MolFromMolBlock(ctab)
        if mol is not None:
            return mol
    if smiles:
        mol = Chem.MolFromSmiles(smiles)
        if mol is not None:
            from rdkit.Chem import AllChem

            AllChem.Compute2DCoords(mol)
            return mol
    return None


def _draw_svg(mol, size: Tuple[int, int]) -> str:
    drawer = rdMolDraw2D.MolDraw2DSVG(*size)
    drawer.DrawMolecule(mol)
    drawer.FinishDrawing()
    return drawer.GetDrawingText()


def _render_structure_svgs(ctabs: List[str]) -> Tuple[List[Optional[str]], str]:
    """Render one 2D structure SVG per CTAB into a fresh temp directory,
    preserving each CTAB's own atom coordinates (no fresh layout computed).
    Returns (paths, tmp_dir) where paths has the same length/order as
    ctabs, with None for anything that fails to parse."""
    tmp_dir = tempfile.mkdtemp(prefix="chem_structs_")
    paths: List[Optional[str]] = []
    for i, ctab in enumerate(ctabs):
        mol = _mol_from_ctab_or_smiles(ctab, "")
        if mol is None:
            paths.append(None)
            continue
        path = os.path.join(tmp_dir, f"mol_{i}.svg")
        try:
            with open(path, "w", encoding="utf-8") as fh:
                fh.write(_draw_svg(mol, _THUMBNAIL_SVG_SIZE))
            paths.append(path)
        except Exception:  # noqa: BLE001 - one bad structure shouldn't sink the batch
            paths.append(None)
    return paths, tmp_dir


# ---------------------------------------------------------------------------
# CLI plumbing
# ---------------------------------------------------------------------------
def _read_stdin_json() -> dict:
    raw = sys.stdin.read()
    if not raw.strip():
        return {}
    return json.loads(raw)


def _emit(obj: dict) -> None:
    sys.stdout.write(json.dumps(obj))
    sys.stdout.flush()


def main(argv: List[str]) -> int:
    if not argv:
        _emit({"error": "No command given."})
        return 1

    cmd = argv[0]
    try:
        if cmd == "load_sdf":
            if len(argv) < 2:
                raise ValueError("load_sdf requires a filepath argument.")
            headers, rows, ctabs = _read_sdf_to_table(argv[1])
            _emit({"headers": headers, "rows": rows, "ctabs": ctabs})

        elif cmd == "write_sdf":
            if len(argv) < 2:
                raise ValueError("write_sdf requires a filepath argument.")
            payload = _read_stdin_json()
            headers = payload.get("headers") or []
            rows = payload.get("rows") or []
            written, skipped = _table_to_sdf(headers, rows, argv[1])
            _emit({"written": written, "skipped": skipped})

        elif cmd == "strip_salts":
            payload = _read_stdin_json()
            smiles_list = payload.get("smiles") or []
            parents = [
                _strip_salt_from_smiles(str(s)) if s else None for s in smiles_list
            ]
            _emit({"parents": parents})

        elif cmd == "add_column":
            if len(argv) < 2:
                raise ValueError("add_column requires a descriptor-name argument.")
            prop_name = argv[1]
            payload = _read_stdin_json()
            smiles_list = payload.get("smiles") or []
            values = [
                _compute_descriptor(str(s), prop_name) if s else None
                for s in smiles_list
            ]
            _emit({"values": values})

        elif cmd == "list_descriptors":
            _emit({"names": sorted(DESCRIPTORS)})

        elif cmd == "render_structures":
            payload = _read_stdin_json()
            ctab_list = payload.get("ctabs") or []
            paths, tmp_dir = _render_structure_svgs([str(c) if c else "" for c in ctab_list])
            _emit({"images": paths, "tmp_dir": tmp_dir})

        elif cmd == "view_structure":
            payload = _read_stdin_json()
            ctab = str(payload.get("ctab") or "")
            smiles = str(payload.get("smiles") or "")
            mol = _mol_from_ctab_or_smiles(ctab, smiles)
            if mol is None:
                raise ValueError("Could not parse a structure from the stored CTAB or SMILES.")
            tmp_dir = tempfile.mkdtemp(prefix="chem_view_")
            path = os.path.join(tmp_dir, "structure.svg")
            with open(path, "w", encoding="utf-8") as fh:
                fh.write(_draw_svg(mol, _LARGE_SVG_SIZE))
            _emit({"image": path, "tmp_dir": tmp_dir})

        else:
            raise ValueError(f'Unknown command "{cmd}".')

    except Exception as exc:  # noqa: BLE001 - surface any failure as JSON, not a traceback
        _emit({"error": str(exc)})
        return 1

    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
