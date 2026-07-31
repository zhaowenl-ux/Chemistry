# -*- coding: utf-8 -*-
"""
Chemistry ribbon add-in for LibreOffice Calc.

This module is loaded by LibreOffice's Python script provider (no third-party
packages required here -- it must import cleanly under LO's *bundled*
interpreter). All RDKit work happens in a separate process, backend/
chemistry_backend.py, run with a regular system Python that has RDKit
installed. See _find_python() below and the add-in's README for how that
Python is located.

Buttons wired up in Addons.xcu, all calling functions in this file:
    Load SDF                     -> load_sdf
    Save SDF                      -> save_sdf
    Salt Stripping                 -> strip_salts
    Add Column > ...               -> add_column_<descriptor> (nine of these,
                                       one per RDKit descriptor -- Calc has no
                                       built-in text InputBox, so instead of
                                       prompting for a property name the
                                       ribbon offers a submenu of the fixed
                                       set of descriptors, same list as the
                                       Excel add-in)
    View Structure                 -> sheet_double_click, bound to the
                                       document's "Double click" sheet event
                                       (not a menu item -- double-click any
                                       cell in the "Structure" column to open
                                       that row's structure; see
                                       _bind_sheet_double_click)

Design notes (mirrors excel_addin/chemistry_addin.py where it makes sense):
- After "Load SDF", column A is a rendered structure image, column B is the
  canonical SMILES (header "SMILES"), and every other SD-file property tag
  becomes its own column after that, in the order first seen. A hidden
  "_CTAB" column (see chemistry_backend.CTAB_COLUMN_NAME) caches each row's
  original CTAB so View Structure can redraw it later at full size without
  losing the source file's own 2D layout; Save SDF ignores this column.
- "Salt Stripping" is non-destructive: it adds a new "Parent SMILES" column
  rather than overwriting the original SMILES.
- Every button reports a one-line summary via a message box (LibreOffice
  doesn't expose a simple public "status bar text" API the way Excel does),
  and any failure -- RDKit error, missing SMILES column, backend Python not
  found -- surfaces as a message box instead of failing silently.
"""
from __future__ import annotations

import json
import os
import re
import shutil
import subprocess
import sys

import uno
from com.sun.star.awt.FontWeight import BOLD
from com.sun.star.beans import PropertyValue
from com.sun.star.ui.dialogs.TemplateDescription import FILEOPEN_SIMPLE, FILESAVE_SIMPLE

# ---------------------------------------------------------------------------
# Locating the RDKit-enabled Python + the backend script
# ---------------------------------------------------------------------------
_CONFIG_PATH = os.path.join(os.path.expanduser("~"), ".chemistry_calc_addin.json")


def _find_python() -> str:
    """Return the path to a Python executable that has RDKit installed.

    Checked in order:
      1. CHEMISTRY_PYTHON environment variable
      2. "python_exe" key in ~/.chemistry_calc_addin.json
      3. "python" / "python3" on PATH (works only if that Python has RDKit)
    """
    env = os.environ.get("CHEMISTRY_PYTHON")
    if env:
        return env

    if os.path.isfile(_CONFIG_PATH):
        try:
            with open(_CONFIG_PATH, "r", encoding="utf-8") as fh:
                cfg = json.load(fh)
            configured = cfg.get("python_exe")
            if configured:
                return configured
        except (OSError, ValueError):
            pass  # fall through to PATH guesses

    return "python.exe" if os.name == "nt" else "python3"


def _this_file_path() -> str:
    """Native filesystem path of this module. LibreOffice's Python script
    provider sets __file__ to a file:// URL (e.g.
    "file:///C:/Users/.../ChemistryCalcAddin.oxt/python/chemistry_calc.py")
    rather than a native path, at least on Windows. os.path.abspath() does
    not recognize that as absolute and silently glues it onto the current
    working directory (LibreOffice's install folder), which is exactly the
    bad concatenated path this function exists to avoid."""
    f = __file__
    if f.startswith("file:"):
        return uno.fileUrlToSystemPath(f)
    return os.path.abspath(f)


def _backend_script() -> str:
    """Path to chemistry_backend.py, bundled as a sibling folder inside the
    extension (python/chemistry_calc.py -> ../backend/chemistry_backend.py).
    LibreOffice's script provider loads this module from its real on-disk
    location (inside the extension's install cache), so __file__ (once
    converted from a file:// URL, see _this_file_path) is reliable here."""
    here = os.path.dirname(_this_file_path())
    return os.path.normpath(os.path.join(here, "..", "backend", "chemistry_backend.py"))


def _run_backend(args, stdin_obj=None):
    """Run chemistry_backend.py with the given CLI args, feeding stdin_obj as
    JSON on stdin (if given) and parsing stdout as JSON. Raises RuntimeError
    with a human-readable message on any failure."""
    python_exe = _find_python()
    script = _backend_script()
    if not os.path.isfile(script):
        raise RuntimeError(
            f"Chemistry backend script not found at:\n{script}\n"
            "The extension may not have installed correctly."
        )

    cmd = [python_exe, script] + list(args)

    # LibreOffice runs with PYTHONHOME/PYTHONPATH pointed at its own bundled
    # interpreter (e.g. .../LibreOffice/program/python-core-3.12.12), and a
    # child process inherits the parent's environment by default. If we
    # don't strip those, the external RDKit Python launches correctly but
    # then loads *LibreOffice's* standard library instead of its own,
    # mixing a pure-Python module (e.g. re.py) from one Python build with a
    # compiled extension (_sre) from another -- "SRE module mismatch" and
    # similar failures. Give the child a clean environment instead.
    env = os.environ.copy()
    for var in ("PYTHONHOME", "PYTHONPATH", "PYTHONEXECUTABLE"):
        env.pop(var, None)

    # Note: no stdin=subprocess.PIPE here -- passing `input=` below already
    # implies a stdin pipe, and subprocess.run() raises ValueError
    # ("stdin and input arguments may not both be used") if both are given.
    kwargs = dict(
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        env=env,
    )
    if os.name == "nt":
        kwargs["creationflags"] = 0x08000000  # CREATE_NO_WINDOW

    try:
        proc = subprocess.run(
            cmd,
            input=json.dumps(stdin_obj) if stdin_obj is not None else "",
            **kwargs,
        )
    except FileNotFoundError:
        raise RuntimeError(
            f'Could not launch Python at "{python_exe}".\n\n'
            f"Set CHEMISTRY_PYTHON, or create {_CONFIG_PATH} with:\n"
            '{"python_exe": "C:\\\\path\\\\to\\\\python.exe"}\n\n'
            "Use the same Python you ran `pip install rdkit` into."
        )

    if not proc.stdout.strip():
        raise RuntimeError(
            f"Chemistry backend produced no output.\nstderr:\n{proc.stderr}"
        )

    try:
        result = json.loads(proc.stdout)
    except ValueError:
        raise RuntimeError(f"Chemistry backend returned invalid output:\n{proc.stdout}\n{proc.stderr}")

    if "error" in result:
        raise RuntimeError(result["error"])
    return result


# ---------------------------------------------------------------------------
# UNO helpers (document, sheet, dialogs)
# ---------------------------------------------------------------------------
def _doc():
    return XSCRIPTCONTEXT.getDocument()


def _sheet():
    return _doc().getCurrentController().getActiveSheet()


def _frame():
    return _doc().getCurrentController().getFrame()


def _msgbox(message: str, title: str = "Chemistry Tools", box_type: str = "INFOBOX") -> None:
    # com.sun.star.awt.MessageBoxType enum members are upper-case
    # (INFOBOX, ERRORBOX, WARNINGBOX, QUERYBOX, MESSAGEBOX) -- normalize
    # defensively so a lower/mixed-case caller doesn't blow up with a
    # RuntimeException instead of showing the intended error message.
    ctx = XSCRIPTCONTEXT.getComponentContext()
    smgr = ctx.getServiceManager()
    toolkit = smgr.createInstanceWithContext("com.sun.star.awt.Toolkit", ctx)
    parent = _frame().getContainerWindow()
    box = toolkit.createMessageBox(
        parent,
        uno.Enum("com.sun.star.awt.MessageBoxType", box_type.upper()),
        1,  # BUTTONS_OK
        title,
        message,
    )
    box.execute()


def _pick_file(save: bool, default_name: str = ""):
    """Show a file picker filtered to .sdf files. Returns a system filepath
    or None if the user cancelled."""
    ctx = XSCRIPTCONTEXT.getComponentContext()
    smgr = ctx.getServiceManager()
    picker = smgr.createInstanceWithContext("com.sun.star.ui.dialogs.FilePicker", ctx)
    picker.initialize((FILESAVE_SIMPLE if save else FILEOPEN_SIMPLE,))
    picker.appendFilter("SD Files (*.sdf)", "*.sdf")
    picker.appendFilter("All Files (*.*)", "*.*")
    picker.setCurrentFilter("SD Files (*.sdf)")
    if default_name:
        picker.setDefaultName(default_name)

    if picker.execute() == 1:  # ExecutableDialogResults.OK
        files = picker.getSelectedFiles()
        if files:
            path = uno.fileUrlToSystemPath(files[0])
            if save and not path.lower().endswith(".sdf"):
                path += ".sdf"
            return path
    return None


def _get_used_range_data(sheet=None):
    """Return (headers, rows) for sheet's used area (the active sheet if
    sheet is not given). Every value that came back as a float from Calc
    (e.g. an integer property that Calc auto-typed as a number) is
    converted to a clean string, matching how the Excel add-in round-trips
    numeric-looking cells."""
    if sheet is None:
        sheet = _sheet()
    cursor = sheet.createCursor()
    cursor.gotoStartOfUsedArea(False)
    cursor.gotoEndOfUsedArea(True)
    addr = cursor.RangeAddress
    used_range = sheet.getCellRangeByPosition(
        addr.StartColumn, addr.StartRow, addr.EndColumn, addr.EndRow
    )
    data = used_range.getDataArray()
    if not data:
        raise RuntimeError("Active sheet is empty -- nothing to do.")

    def clean(v):
        if isinstance(v, float):
            if v == int(v):
                return str(int(v))
            return str(v)
        return v

    headers = [str(h) for h in data[0]]
    rows = [[clean(v) for v in row] for row in data[1:]]
    return headers, rows


def _write_table(headers, rows):
    """Clear the active sheet and write headers+rows starting at A1, bold
    the header row, and autofit columns -- mirrors the Excel add-in's
    Load SDF behavior."""
    sheet = _sheet()

    # Clear whatever is currently on the sheet first.
    cursor = sheet.createCursor()
    cursor.gotoStartOfUsedArea(False)
    cursor.gotoEndOfUsedArea(True)
    old_addr = cursor.RangeAddress
    old_range = sheet.getCellRangeByPosition(
        old_addr.StartColumn, old_addr.StartRow, old_addr.EndColumn, old_addr.EndRow
    )
    old_range.clearContents(1023)  # all content flags

    n_rows = len(rows) + 1
    n_cols = len(headers)
    dest = sheet.getCellRangeByPosition(0, 0, max(n_cols - 1, 0), max(n_rows - 1, 0))

    # getDataArray/setDataArray need a rectangular tuple-of-tuples of
    # consistent width; pad short rows defensively.
    table = [headers] + [
        (row + [""] * (n_cols - len(row)))[:n_cols] for row in rows
    ]
    dest.setDataArray(tuple(tuple(r) for r in table))

    header_range = sheet.getCellRangeByPosition(0, 0, max(n_cols - 1, 0), 0)
    header_range.CharWeight = BOLD
    dest.Columns.OptimalWidth = True


def _append_column(header: str, values):
    """Append one new column (header + one value per existing data row) to
    the right of the sheet's current used area."""
    sheet = _sheet()
    cursor = sheet.createCursor()
    cursor.gotoStartOfUsedArea(False)
    cursor.gotoEndOfUsedArea(True)
    addr = cursor.RangeAddress
    new_col = addr.EndColumn + 1

    header_cell = sheet.getCellByPosition(new_col, addr.StartRow)
    header_cell.setString(header)
    header_cell.CharWeight = BOLD

    for i, val in enumerate(values):
        cell = sheet.getCellByPosition(new_col, addr.StartRow + 1 + i)
        if val is None:
            cell.setString("")
        elif isinstance(val, (int, float)):
            cell.setValue(val)
        else:
            cell.setString(str(val))

    col_range = sheet.getCellRangeByPosition(new_col, addr.StartRow, new_col, addr.EndRow)
    col_range.Columns.OptimalWidth = True


# ---------------------------------------------------------------------------
# Structure images (Load SDF's "Structure" column)
# ---------------------------------------------------------------------------
_STRUCT_SHAPE_PREFIX = "ChemStructure_"
_STRUCT_IMG_WIDTH_MM = 30
_STRUCT_IMG_HEIGHT_MM = 22
_STRUCT_CELL_MARGIN_MM = 2  # extra room around the image inside its cell


def _prop(name: str, value) -> PropertyValue:
    p = PropertyValue()
    p.Name = name
    p.Value = value
    return p


def _clear_structure_shapes(sheet) -> None:
    """Remove any structure images this add-in previously placed on this
    sheet (identified by name prefix, so we never touch shapes the user
    added themselves). Called before writing a fresh table so reloading
    into the same sheet doesn't leave orphaned floating images behind."""
    draw_page = sheet.getDrawPage()
    to_remove = []
    for i in range(draw_page.Count):
        shape = draw_page.getByIndex(i)
        try:
            name = shape.Name
        except Exception:  # noqa: BLE001
            name = ""
        if name.startswith(_STRUCT_SHAPE_PREFIX):
            to_remove.append(shape)
    for shape in to_remove:
        draw_page.remove(shape)


def _insert_structure_image(sheet, col: int, row: int, image_path: str) -> None:
    """Place image_path as a floating picture anchored over (col, row),
    sized to a fixed on-screen size regardless of the source PNG's pixel
    dimensions."""
    ctx = XSCRIPTCONTEXT.getComponentContext()
    smgr = ctx.getServiceManager()

    # GraphicProvider is a free-standing global service -- fine to create via
    # the component context.
    provider = smgr.createInstanceWithContext("com.sun.star.graphic.GraphicProvider", ctx)
    graphic = provider.queryGraphic((_prop("URL", uno.systemPathToFileUrl(image_path)),))

    # The shape itself must be created through the *document's* own factory
    # (XMultiServiceFactory.createInstance), not the global service manager,
    # so it's properly bound to this document's model before being added to
    # a draw page. Creating it via the global service manager is a common
    # cause of the shape silently failing to render.
    shape = _doc().createInstance("com.sun.star.drawing.GraphicObjectShape")
    sheet.getDrawPage().add(shape)  # must be added to a page before Graphic/Size/Position stick
    shape.Graphic = graphic
    shape.Name = f"{_STRUCT_SHAPE_PREFIX}{row}"

    cell_pos = sheet.getCellByPosition(col, row).Position
    margin = _STRUCT_CELL_MARGIN_MM * 100
    shape.Size = uno.createUnoStruct(
        "com.sun.star.awt.Size", _STRUCT_IMG_WIDTH_MM * 100, _STRUCT_IMG_HEIGHT_MM * 100
    )
    shape.Position = uno.createUnoStruct(
        "com.sun.star.awt.Point", cell_pos.X + margin, cell_pos.Y + margin
    )


def _insert_structure_images(sheet, image_paths):
    """Insert one structure image per (row, path) pair into column A,
    starting at data row 1 (row 0 is the header). Skips None paths (failed
    renders). Also sizes column A and the data rows to fit. Returns
    (placed_count, last_error_message_or_None) -- the error is surfaced (not
    swallowed) so a total failure shows up in the Load SDF summary instead
    of silently leaving column A blank."""
    # Resize column A and the data rows *before* reading any cell's
    # Position: shape positions below are computed from cell.Position at
    # the moment of insertion, using whatever row heights are in effect
    # then. Resizing rows *after* placing shapes leaves every already-placed
    # shape's absolute Y coordinate stale -- rows below the resized ones
    # shift down on screen, but the shapes don't move with them, so the
    # mismatch grows with every row (invisible for a couple of rows, obvious
    # past ~5). Doing the resize first means every cell.Position read below
    # already reflects the final layout.
    if image_paths:
        last_row = len(image_paths)
        col_range = sheet.getCellRangeByPosition(0, 0, 0, last_row)
        col_range.Columns.Width = (_STRUCT_IMG_WIDTH_MM + 2 * _STRUCT_CELL_MARGIN_MM) * 100
        rows_range = sheet.getCellRangeByPosition(0, 1, 0, last_row)
        rows_range.Rows.Height = (_STRUCT_IMG_HEIGHT_MM + 2 * _STRUCT_CELL_MARGIN_MM) * 100

    placed = 0
    last_error = None
    for i, path in enumerate(image_paths):
        if not path:
            continue
        try:
            _insert_structure_image(sheet, 0, i + 1, path)
            placed += 1
        except Exception as exc:  # noqa: BLE001 - one bad image shouldn't sink the whole batch
            last_error = f"{type(exc).__name__}: {exc}"
            continue

    return placed, last_error


_INVALID_SHEET_CHARS = set("[]*?:/\\")


def _rename_active_sheet(desired_name: str) -> str:
    """Rename the active sheet to desired_name, sanitized for Calc's sheet
    naming rules (no [ ] * ? : / \\, can't start/end with an apostrophe,
    max 31 characters) and de-duplicated against the document's other
    sheets. Returns the name actually applied, or the sheet's unchanged
    name if renaming fails for any reason -- this is a cosmetic step and
    should never block a successful Load SDF."""
    sheet = _sheet()
    try:
        sheets = _doc().getSheets()

        cleaned = "".join(c if c not in _INVALID_SHEET_CHARS else "_" for c in desired_name)
        cleaned = cleaned.strip().strip("'")
        if not cleaned:
            cleaned = "Sheet"
        cleaned = cleaned[:31]

        existing = {
            sheets.getByIndex(i).Name
            for i in range(sheets.Count)
            if sheets.getByIndex(i).Name != sheet.Name
        }

        candidate = cleaned
        suffix = 2
        while candidate in existing:
            trimmed = cleaned[: 31 - len(f" ({suffix})")]
            candidate = f"{trimmed} ({suffix})"
            suffix += 1

        sheet.setName(candidate)
        return candidate
    except Exception:  # noqa: BLE001 - cosmetic step, never fail the load over it
        return sheet.Name


def _smiles_column_index(headers):
    lower = [h.strip().lower() for h in headers]
    if "smiles" not in lower:
        raise RuntimeError('No "SMILES" column found on this sheet.')
    return lower.index("smiles")


# ---------------------------------------------------------------------------
# Hidden "_CTAB" column -- caches each row's original CTAB so View Structure
# can redraw the source file's own layout later without needing to reopen
# the SDF or fall back to a freshly-computed one from SMILES.
# ---------------------------------------------------------------------------
_CTAB_COLUMN_NAME = "_CTAB"  # must match chemistry_backend.CTAB_COLUMN_NAME


def _write_ctab_column(sheet, ctabs) -> None:
    """Append a hidden column holding one CTAB per data row, to the right
    of the sheet's current used area. Never raises -- this is a cache, not
    something Load SDF should fail over."""
    try:
        cursor = sheet.createCursor()
        cursor.gotoStartOfUsedArea(False)
        cursor.gotoEndOfUsedArea(True)
        addr = cursor.RangeAddress
        new_col = addr.EndColumn + 1

        sheet.getCellByPosition(new_col, addr.StartRow).setString(_CTAB_COLUMN_NAME)
        for i, ctab in enumerate(ctabs):
            sheet.getCellByPosition(new_col, addr.StartRow + 1 + i).setString(ctab or "")

        sheet.getColumns().getByIndex(new_col).IsVisible = False
    except Exception:  # noqa: BLE001 - cache only; View Structure falls back to SMILES without it
        pass


def _get_ctab_for_row(sheet, row_index: int) -> str:
    """Look up the CTAB cached for data row row_index (a sheet row index,
    not a 0-based data index). Returns "" if the column or value is
    missing -- callers fall back to SMILES in that case."""
    try:
        headers, rows = _get_used_range_data()
        lower = [str(h).strip() for h in headers]
        if _CTAB_COLUMN_NAME not in lower:
            return ""
        ctab_idx = lower.index(_CTAB_COLUMN_NAME)
        data_idx = row_index - 1
        if data_idx < 0 or data_idx >= len(rows):
            return ""
        return str(rows[data_idx][ctab_idx] or "")
    except Exception:  # noqa: BLE001
        return ""


def _open_svg_externally(svg_path: str) -> None:
    """Hand the rendered SVG off to the OS's default viewer (typically a
    browser, which renders SVG natively and supports pan/zoom -- a "large
    external canvas" for free, with no custom viewer to build/maintain).
    Deliberately does not delete svg_path afterward: the external viewer is
    a separate process with an unknown lifetime, so there's no safe moment
    to clean up: it's left for the OS's normal temp-file housekeeping."""
    if os.name == "nt":
        os.startfile(svg_path)  # noqa: S606 - a path we just wrote ourselves
    else:
        subprocess.run(["xdg-open", svg_path], check=False)


# ---------------------------------------------------------------------------
# Button handlers (bound from Addons.xcu) -- every one is wrapped so a
# failure shows a message box instead of failing silently.
# ---------------------------------------------------------------------------
def load_sdf(*args) -> None:
    tmp_dir = None
    try:
        filepath = _pick_file(save=False)
        if not filepath:
            return
        result = _run_backend(["load_sdf", filepath])
        headers, rows = result["headers"], result["rows"]
        ctabs = result.get("ctabs") or []

        # Prepend a "Structure" column; everything else (Salt Stripping, Add
        # Column, Save SDF) looks up "SMILES" by header name, not position,
        # so this doesn't disturb them.
        display_headers = ["Structure"] + headers
        display_rows = [[""] + row for row in rows]

        sheet = _sheet()
        _clear_structure_shapes(sheet)
        _write_table(display_headers, display_rows)
        base_name = os.path.splitext(os.path.basename(filepath))[0]
        sheet_name = _rename_active_sheet(base_name)
        _write_ctab_column(sheet, ctabs)
        _bind_sheet_double_click(_doc())

        placed = 0
        image_error = None
        try:
            render_result = _run_backend(["render_structures"], {"ctabs": ctabs})
            image_paths, tmp_dir = render_result["images"], render_result.get("tmp_dir")
            placed, image_error = _insert_structure_images(sheet, image_paths)
        except Exception as exc:  # noqa: BLE001 - structures are a nice-to-have, not worth failing the load over
            image_error = str(exc)

        msg = (
            f"Loaded {len(rows)} molecules, {len(headers)} fields from "
            f"{os.path.basename(filepath)} into sheet \"{sheet_name}\""
        )
        if placed:
            msg += f" ({placed} structure images drawn)"
        elif image_error:
            msg += f"\n\nStructure images were not drawn:\n{image_error}"
        _msgbox(msg)
    except Exception as exc:  # noqa: BLE001
        _msgbox(str(exc), title="Load SDF failed", box_type="errorbox")
    finally:
        if tmp_dir:
            shutil.rmtree(tmp_dir, ignore_errors=True)


def save_sdf(*args) -> None:
    try:
        headers, rows = _get_used_range_data()
        filepath = _pick_file(save=True, default_name="molecules.sdf")
        if not filepath:
            return
        result = _run_backend(["write_sdf", filepath], {"headers": headers, "rows": rows})
        written, skipped = result["written"], result["skipped"]
        msg = f"Saved {written} molecules to {os.path.basename(filepath)}"
        if skipped:
            msg += f" ({skipped} rows skipped: blank/invalid SMILES)"
        _msgbox(msg)
    except Exception as exc:  # noqa: BLE001
        _msgbox(str(exc), title="Save SDF failed", box_type="errorbox")


def strip_salts(*args) -> None:
    try:
        headers, rows = _get_used_range_data()
        smiles_idx = _smiles_column_index(headers)
        smiles_list = [row[smiles_idx] for row in rows]
        result = _run_backend(["strip_salts"], {"smiles": smiles_list})
        parents = result["parents"]
        _append_column("Parent SMILES", parents)
        changed = sum(1 for s, p in zip(smiles_list, parents) if p and p != s)
        _msgbox(
            f"Salt stripping complete: {changed} of {len(rows)} rows had a "
            "salt/solvent removed."
        )
    except Exception as exc:  # noqa: BLE001
        _msgbox(str(exc), title="Salt Stripping failed", box_type="errorbox")


def _add_column(prop_name: str) -> None:
    headers, rows = _get_used_range_data()
    smiles_idx = _smiles_column_index(headers)
    smiles_list = [row[smiles_idx] for row in rows]
    result = _run_backend(["add_column", prop_name], {"smiles": smiles_list})
    values = result["values"]
    _append_column(prop_name, values)
    computed = sum(1 for v in values if v is not None)
    _msgbox(f'Added "{prop_name}" column: computed for {computed} of {len(rows)} rows.')


def _add_column_safe(prop_name: str) -> None:
    try:
        _add_column(prop_name)
    except Exception as exc:  # noqa: BLE001
        _msgbox(str(exc), title="Add Column failed", box_type="errorbox")


# One exported wrapper per descriptor (bound individually in Addons.xcu's
# "Add Column" submenu, since Calc has no built-in text-entry InputBox).
def add_column_molwt(*args) -> None:
    _add_column_safe("MolWt")


def add_column_logp(*args) -> None:
    _add_column_safe("LogP")


def add_column_tpsa(*args) -> None:
    _add_column_safe("TPSA")


def add_column_hbd(*args) -> None:
    _add_column_safe("HBD")


def add_column_hba(*args) -> None:
    _add_column_safe("HBA")


def add_column_numrings(*args) -> None:
    _add_column_safe("NumRings")


def add_column_rotatablebonds(*args) -> None:
    _add_column_safe("RotatableBonds")


def add_column_heavyatomcount(*args) -> None:
    _add_column_safe("HeavyAtomCount")


def add_column_fractioncsp3(*args) -> None:
    _add_column_safe("FractionCSP3")


_SHEET_DOUBLE_CLICK_SCRIPT_URI = (
    "vnd.sun.star.script:ChemistryCalcAddin.oxt|python|chemistry_calc.py"
    "$sheet_double_click?language=Python&location=user:uno_packages"
)


def _bind_sheet_double_click(doc) -> None:
    """Assign sheet_double_click to the document's "Double click" sheet
    event (Tools > Customize > Events / Sheet > Sheet Events in the UI).
    This is a genuine Calc event -- distinct from, and more reliable than,
    binding a click handler to each individual floating image shape -- so
    double-clicking any cell in the "Structure" column opens that row's
    structure. Idempotent (safe to call again on every Load SDF); harmless
    everywhere else, since the handler itself ignores clicks outside the
    Structure column and lets Calc's normal double-click-to-edit proceed."""
    try:
        events = doc.getEvents()
        events.replaceByName(
            "OnDoubleClick",
            (_prop("EventType", "Script"), _prop("Script", _SHEET_DOUBLE_CLICK_SCRIPT_URI)),
        )
    except Exception:  # noqa: BLE001 - View Structure just won't auto-trigger; nothing else breaks
        pass


def sheet_double_click(cell_range) -> bool:
    """Bound to the document's "Double click" sheet event (see
    _bind_sheet_double_click). Fires for every double-click on any cell in
    the document; only acts if the cell is in column A ("Structure") of a
    sheet that has one, on a data row. Returns True to tell Calc the
    double-click was handled (skip entering edit mode on that cell), or
    False to let Calc's normal double-click behavior proceed everywhere
    else."""
    try:
        sheet = cell_range.getSpreadsheet()
        addr = cell_range.RangeAddress
        col, row = addr.StartColumn, addr.StartRow
        if col != 0 or row == 0:
            return False  # not the Structure column, or the header row

        headers, rows = _get_used_range_data(sheet)
        if not headers or str(headers[0]).strip().lower() != "structure":
            return False

        smiles_idx = _smiles_column_index(headers)
        data_idx = row - 1
        if data_idx < 0 or data_idx >= len(rows):
            return False

        smiles = str(rows[data_idx][smiles_idx] or "")
        ctab = _get_ctab_for_row(sheet, row)
        result = _run_backend(["view_structure"], {"ctab": ctab, "smiles": smiles})
        _open_svg_externally(result["image"])  # not cleaned up -- see _open_svg_externally
    except Exception as exc:  # noqa: BLE001
        _msgbox(str(exc), title="View Structure failed", box_type="errorbox")
    return True


def _addin_version() -> str:
    """Read the version straight out of description.xml (one directory up
    from this file inside the extension) so the About box can never drift
    out of sync with the packaged version number."""
    try:
        desc_path = os.path.join(os.path.dirname(_this_file_path()), "..", "description.xml")
        with open(desc_path, "r", encoding="utf-8") as fh:
            content = fh.read()
        match = re.search(r'<version\s+value="([^"]+)"', content)
        if match:
            return match.group(1)
    except Exception:  # noqa: BLE001
        pass
    return "unknown"


def show_about(*args) -> None:
    try:
        lines = [
            "Chemistry Tools for LibreOffice Calc",
            f"Version {_addin_version()}",
            "",
            "RDKit-powered tools for SD files: Load SDF, Save SDF, Salt",
            "Stripping, Add Column (nine descriptors), and View Structure",
            "(double-click any cell in the Structure column).",
            "",
            f"Backend Python: {_find_python()}",
            f"Backend script: {_backend_script()}",
        ]
        _msgbox("\n".join(lines), title="About Chemistry Tools")
    except Exception as exc:  # noqa: BLE001
        _msgbox(str(exc), title="About failed", box_type="errorbox")
