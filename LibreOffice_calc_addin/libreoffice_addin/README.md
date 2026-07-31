# Chemistry Tools for LibreOffice Calc (RDKit)

**Version 0.21.**

The LibreOffice Calc counterpart to the Excel `Chemistry` ribbon, plus one
Calc-only extra. **Load SDF**, **Save SDF**, **Salt Stripping**, and **Add
Column** match the Excel add-in; **View Structure** is new here, packaged
as a real LibreOffice extension (`.oxt`), no VBA/RibbonX editor required.
Chemistry still runs via RDKit; LibreOffice's own Python has nothing
installed into it.

> **v0.11 fixed a packaging bug in v1.0.0:** clicking any button raised `A
> Scripting Framework error occurred ... an error occurred during file
> opening`. Cause: the URI LibreOffice uses to locate a Python script
> bundled *inside* an extension has a specific three-part format
> (`<oxt-filename>|<folder>|<module>.py$<function>` with
> `location=user:uno_packages`), which the first build didn't use.
>
> **v0.12 fixed a second bug** uncovered once v0.11 let buttons actually
> run: the error-reporting message box itself used the wrong casing for a
> UNO enum (`errorbox` instead of `ERRORBOX`), so any real error crashed the
> error dialog instead of showing a message.
>
> **v0.13 fixed a third bug**, the one v0.12 finally revealed: on Windows,
> LibreOffice's Python script provider sets a bundled script's `__file__`
> to a `file://` URL (e.g. `file:///C:/Users/.../ChemistryCalcAddin.oxt/
> python/chemistry_calc.py`) rather than a normal Windows path.
> `os.path.abspath()` doesn't recognize a `file://` string as absolute, so
> it silently glued it onto LibreOffice's own install directory when
> looking for `chemistry_backend.py`, producing a mangled path like
> `C:\Program Files\LibreOffice\program\file:\C:\Users\...`. Fixed by
> converting through `uno.fileUrlToSystemPath()` first.
>
> **v0.14 fixed a fourth bug**, the one v0.13 finally revealed once the
> extension could actually find and launch its backend Python: the
> subprocess call passed both `stdin=subprocess.PIPE` and `input=...` to
> `subprocess.run()`, which Python rejects outright ("stdin and input
> arguments may not both be used"). Removed the redundant `stdin=` kwarg --
> `input=` already implies a pipe.
>
> **v0.15 fixed a fifth bug**, the one v0.14 finally revealed once the
> backend Python actually launched: it crashed on `import re` with
> `AssertionError: SRE module mismatch`. Cause: LibreOffice sets
> `PYTHONHOME`/`PYTHONPATH` to point at its own bundled interpreter, and a
> subprocess inherits those by default -- so your separate RDKit Python
> launched, but then loaded *LibreOffice's* standard library instead of
> its own, mixing a pure-Python module from one Python build with a
> compiled extension from another. Fixed by stripping
> `PYTHONHOME`/`PYTHONPATH`/`PYTHONEXECUTABLE` from the environment before
> launching the backend.
>
> **v0.16 adds a feature**: Load SDF now renames the active sheet to the
> loaded file's name (invalid characters replaced with `_`, truncated to
> Calc's 31-character sheet-name limit, and de-duplicated with " (2)",
> " (3)", etc. if a sheet with that name already exists). If renaming fails
> for any reason, the load still succeeds -- it's cosmetic only.
>
> **v0.17 adds two features:**
> 1. Load SDF now draws each molecule's 2D structure as an image in a new
>    "Structure" column A (SMILES and everything else shift one column
>    right -- Salt Stripping/Add Column/Save SDF all look up "SMILES" by
>    header name, so this doesn't break them). Images are rendered by
>    RDKit in the backend process, inserted as floating pictures anchored
>    over column A, and the temp PNGs are deleted once inserted. Reloading
>    into the same sheet clears out the previous batch of images first.
> 2. A new **Chemistry > About Chemistry Tools...** menu item shows the
>    installed version (read live from `description.xml`, so it can't
>    drift out of sync) plus which backend Python/script it's using --
>    handy for confirming a reinstall actually took.
>
> **v0.18 fixed the structure images**: v0.17 added the "Structure" column
> but the images never actually appeared. Cause: the picture shape was
> created via the global UNO service manager instead of the *document's*
> own factory (`doc.createInstance(...)`), which is required for a shape
> to properly attach to that document's model. It likely failed silently
> in a way the old blanket `except: pass` around the image step then
> swallowed -- that catch-all is now gone too, so if images still don't
> appear, the Load SDF summary will show the actual error instead of
> nothing.
>
> **v0.19 fixed image row alignment**: with more than a handful of
> molecules, images landed in the wrong rows, drifting further off with
> each row. Cause: each image's position was read from `cell.Position`
> *before* the data rows were resized to fit the images, then the resize
> ran afterward -- shifting every later row down on screen without moving
> the already-placed shapes to match, so the misalignment compounded row
> by row. Fixed by resizing column A and the data rows first, then placing
> images against the final layout.
>
> **v0.20 adds View Structure and switches to SVG rendering from the
> original CTAB:**
> - Load SDF now keeps each record's original CTAB (molfile block) and
>   renders both the small in-sheet thumbnail and the new full-size view as
>   SVG instead of PNG. This preserves the source file's own 2D layout
>   (RDKit no longer recomputes one from SMILES) and stays crisp at any
>   zoom. The CTABs are cached in a new hidden `_CTAB` column on the sheet
>   (Save SDF ignores it -- it still rebuilds from the `SMILES` column, as
>   before).
> - Clicking a structure image -- where LibreOffice's shape-click event
>   binding takes on your version -- opens a large SVG rendering of that
>   molecule in your system's default SVG viewer (normally a browser),
>   which handles pan/zoom natively. Calc's shape events don't cleanly
>   distinguish a double-click from a single click, so it fires on click.
>   A **Chemistry > View Structure (Selected Image)** menu item is the
>   guaranteed fallback: select an image (always works, native Calc
>   behavior) and run the command.
> - The earlier (unshipped) SMILES-editing dialog has been dropped in favor
>   of this view-only feature, per request.
>
> **v0.21 replaces the shape-click/menu triggers with a real double-click
> event on the cell:**
> - Removed the **Chemistry > View Structure (Selected Image)** menu item
>   and the per-image shape-click binding entirely.
> - View Structure now fires from Calc's own document-level **"Double
>   click"** sheet event (the same one under Sheet > Sheet Events / Tools >
>   Customize > Events in the UI), bound automatically each time you run
>   Load SDF. Double-click any cell in the **Structure** column on a data
>   row to open that row's large SVG; double-clicking anywhere else behaves
>   exactly as Calc normally would (enters edit mode). This is a genuine
>   Calc cell event rather than a shape click, so it's consistent across
>   LibreOffice versions and doesn't require selecting an image first.
>
> If you installed an older `.oxt`, remove it in Extension Manager first
> and install this one.

## Files in this folder

| File | Purpose |
|---|---|
| `ChemistryCalcAddin.oxt` | The extension. Install this in LibreOffice. **Do not rename this file** -- its filename is baked into the menu commands (see "Do not rename" below). |
| `chemistry_calc.py` | Source of the UNO macro bundled inside the `.oxt` (`python/`), for reference/editing. |
| `chemistry_backend.py` | Source of the RDKit CLI backend bundled inside the `.oxt` (`backend/`). |
| `Addons.xcu`, `description.xml`, `META-INF/manifest.xml` | Extension packaging metadata, for reference. |

## Why a two-step architecture

LibreOffice ships its own bundled Python interpreter, and that interpreter
is what runs `chemistry_calc.py`. Installing a compiled package like RDKit
*into* LibreOffice's internal Python is fragile and breaks across LO
updates, so this add-in doesn't try. Instead, the macro shells out to a
plain system Python (the same one you can use for the Excel add-in) that
has RDKit installed, the same way the Excel version hands off from VBA to
Python via xlwings. That's the one manual step below.

## One-time setup

**1. Install RDKit into a regular Python** (skip if you already did this
for the Excel add-in -- reuse the same interpreter):

```
pip install rdkit
```

Note where that Python executable lives, e.g. `C:\Python312\python.exe` or
a venv's `Scripts\python.exe`.

**2. Tell the add-in which Python to use.**

Create a file at `%USERPROFILE%\.chemistry_calc_addin.json` (i.e.
`C:\Users\<you>\.chemistry_calc_addin.json`) containing:

```json
{"python_exe": "C:\\Python312\\python.exe"}
```

(Use the exact path from step 1, with doubled backslashes as shown, or
forward slashes.) You only need to do this once per machine. Alternatively,
set an environment variable `CHEMISTRY_PYTHON` to the same path.

**3. Install the extension.**

- Double-click `ChemistryCalcAddin.oxt`, or in LibreOffice: **Tools >
  Extension Manager > Add...** and pick the file.
- When prompted, choose **"Only for me"** (this build is packaged for a
  per-user install and needs no admin rights). "For all users" would need
  a differently-packaged `.oxt` -- ask if you need that variant instead.
- Restart LibreOffice (Extension Manager will prompt you to).

**Do not rename or copy `ChemistryCalcAddin.oxt`.** LibreOffice looks up
bundled scripts by the exact filename the extension was installed from, so
`ChemistryCalcAddin (1).oxt` or a browser-mangled download name will break
every button with the same "Scripting Framework error" this version fixes.
If that happens: remove the extension in Extension Manager, make sure the
file is named exactly `ChemistryCalcAddin.oxt`, and reinstall.

**4. Open Calc.**

You should see a **Chemistry** menu on the menu bar and a **Chemistry**
toolbar with three buttons (Load SDF, Save SDF, Salt Stripping). The full
set of tools -- including the Add Column submenu, View Structure, and
About -- is under the **Chemistry** menu.

## Using it

- **Load SDF** -- file picker -> parses the SD file with RDKit and writes a
  flat table to the active sheet, then renames that sheet to the file's
  name (e.g. loading `KLK_registration.sdf` renames the sheet to
  `KLK_registration`). Column A is a rendered 2D **Structure** image for
  each row (SVG, drawn from the record's original CTAB so it matches the
  source file's own layout); column B is always canonical `SMILES`; every
  other SD property tag becomes its own column after that, in first-seen
  order. A hidden `_CTAB` column caches the original CTAB per row for View
  Structure -- don't unhide/edit/delete it, and note Save SDF (below)
  ignores it.
- **Save SDF** -- file picker -> rebuilds molecules from the `SMILES`
  column and writes every other non-blank column back out as an SD
  property tag (the hidden `_CTAB` column is skipped, not written out as a
  bogus property).
- **Salt Stripping** -- non-destructive. Adds a new `Parent SMILES` column
  with counter-ions/solvates removed via RDKit's `SaltRemover`; the
  original `SMILES` column is untouched.
- **Add Column** -- a submenu of nine RDKit descriptors (`MolWt`, `LogP`,
  `TPSA`, `HBD`, `HBA`, `NumRings`, `RotatableBonds`, `HeavyAtomCount`,
  `FractionCSP3`); pick one and it's appended as a new column, computed per
  row from the `SMILES` column. (Calc has no built-in text-entry prompt
  like Excel's InputBox, so this is a pick-list instead of free text --
  same nine options as the Excel add-in.)
- **View Structure** -- double-click any cell in the **Structure** column
  on a data row to open a large SVG rendering of that molecule in your
  system's default SVG viewer (usually a browser) -- scalable, pan/zoom for
  free. No menu item or image click needed; this is a real Calc cell
  double-click event, (re-)bound automatically each time you run Load SDF.
  Prefers the row's cached `_CTAB`; falls back to a fresh layout from
  `SMILES` if that's missing.

- **About Chemistry Tools...** -- shows the installed version and which
  backend Python/script the add-in is currently using.

Every action reports a one-line summary in a dialog box when done (Calc
doesn't expose a simple status-bar-text API the way Excel does), and any
error -- RDKit failure, missing `SMILES` column, backend Python not found
-- shows as a message box instead of failing silently.

## Already verified

`chemistry_backend.py`'s core functions were run directly against
`KLK_registration_Jul13_2026_registered.sdf` (same file used to verify the
Excel add-in):

- **Load -> Save round-trip**: 20/20 molecules parsed, rebuilt, and
  re-verified as valid SD records with no data loss.
- **Salt Stripping**: correctly detected and removed a TFA
  (trifluoroacetate) counter-ion from 10 of 20 compounds, leaving parent
  structures intact and already-neutral compounds unchanged.
- **Add Column**: `MolWt` computed correctly on sample rows; an unknown
  descriptor name raises a clear error.
- The `.oxt` package's XML files (`Addons.xcu`, `description.xml`,
  `manifest.xml`) are well-formed, and the zip structure was verified with
  `unzip -l`.

The one part that couldn't be tested here is the live Calc/menu interaction
itself, since that requires an actual LibreOffice instance with a display --
this is how the v1.0.0 URI-format bug slipped through; v0.11's fix was
verified against LibreOffice's own documented Python script URI spec and a
known-working real-world extension, but you're the first real test of the
menu firing end to end.

If a menu item still doesn't fire after installing v0.11: check that the
file is named exactly `ChemistryCalcAddin.oxt` (see "Do not rename" above)
and that it was installed "Only for me". If a button shows "Chemistry
backend script not found," the extension likely didn't install correctly
-- try removing and reinstalling via Extension Manager. If it fires but
shows "Could not launch Python," double check step 2 above.
