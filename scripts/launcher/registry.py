"""
Function registry for the script launcher.

Add a new function = add a new FunctionSpec to FUNCTIONS below.
The web UI rebuilds itself from this list automatically.

Each FunctionSpec has:
  id           Unique slug; used in URLs and DOM ids.
  title        Card heading.
  description  One sentence shown under the title.
  fields       List of FieldSpec (text inputs / selects shown to the user).
  build_command  fn(values: dict[str, str]) -> list[str]
                 Returns the argv list passed to subprocess.
                 For cards with run_actions, the launcher adds __run_action
                 (matching RunActionSpec.id) so the lambda can branch.
  run_actions  Optional tuple of RunActionSpec: replaces the single Run button
               with one labeled button per entry (type button + __run_action).
  background   If True, fire-and-forget (use for GUI/long-running scripts).
  cwd          Optional working directory.
  hidden       If True, card is omitted from the web UI (API run by id still works).

PowerShell scripts are invoked as:
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File <script> -Param value ...
That avoids quoting issues entirely.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Literal


SCRIPTS_DIR = Path(__file__).resolve().parent.parent
WDTS_MCP_DIR = SCRIPTS_DIR.parent / "wdts-altium-mcp"
ALTIUM_BOM_TO_CSV = Path(
    r"C:\Users\mlewis\OneDrive - Walker Digital\Documents\Convert_BOM_to_CSV.ps1"
)


@dataclass
class WorkspaceSpec:
    id: str
    label: str
    workspace_id: str


# Top-bar workspace selector. The chosen workspace_id is forwarded as
# ARENA_WORKSPACE_ID into the subprocess environment so any Arena script
# (that reads ARENA_WORKSPACE_ID) honours it.
WORKSPACES: list[WorkspaceSpec] = [
    WorkspaceSpec(id="real", label="Real workspace", workspace_id="896309030"),
    WorkspaceSpec(id="training", label="Training workspace", workspace_id="896315696"),
]
DEFAULT_WORKSPACE_ID = "real"


@dataclass
class FieldSpec:
    name: str
    label: str
    placeholder: str = ""
    required: bool = False
    default: str = ""
    kind: Literal["text", "select", "textarea", "file", "loadable_select", "checkbox"] = "text"
    options: list[str] = field(default_factory=list)
    accept: str = ""  # for kind="file": HTML accept attribute, e.g. ".pdf,.zip"
    # for kind="loadable_select": which field names from the discovery
    # response should also be exposed to build_command, named
    # "<field>__<meta>", e.g. existing_file__edition.
    discovery_meta: list[str] = field(default_factory=list)


@dataclass(frozen=True)
class RunActionSpec:
    """Optional extra run targets on one card (replaces the default Run label)."""

    id: str
    label: str


@dataclass(frozen=True)
class ProceedAfterFailSpec:
    """When a foreground run fails, the UI may offer a follow-up action."""

    label: str
    target_id: str  # must match another FunctionSpec.id


@dataclass
class DiscoverySpec:
    """Optional discovery for a card.

    The launcher exposes POST /api/discover/{function_id}; on call, it
    runs `build_command(values)` (a script that emits JSON on stdout) and
    parses the JSON. The frontend uses the result to populate the card's
    `loadable_select` fields.

    The script's JSON output should be either a list of options, or an
    object with key `options` whose value is the list. Each option must
    have at least {"value": ..., "label": ...}; any extra keys are
    forwarded to the frontend as data attributes and surface as
    "<field>__<key>" form fields when the form is submitted.
    """

    build_command: Callable[[dict[str, str]], list[str]]


@dataclass
class FunctionSpec:
    id: str
    title: str
    description: str
    build_command: Callable[[dict[str, str]], list[str]]
    fields: list[FieldSpec] = field(default_factory=list)
    run_actions: tuple[RunActionSpec, ...] = ()
    background: bool = False
    cwd: Path | None = None
    group: str = "General"
    discovery: DiscoverySpec | None = None
    icon: str = ""
    proceed_after_fail: ProceedAfterFailSpec | None = None
    # If true, launcher shows Yes/No when script exits 10 (duplicate filename on Arena).
    duplicate_upload_prompt: bool = False
    # If true, launcher offers "Proceed with Zip Overwrite" when script exits 11.
    zip_overwrite_prompt: bool = False
    # With discovery: "Load name & description" applies JSON `prefill` to same-named form fields.
    discovery_prefill: bool = False
    hidden: bool = False

    @property
    def has_file_field(self) -> bool:
        return any(f.kind == "file" for f in self.fields)

    @property
    def has_loadable_select(self) -> bool:
        return any(f.kind == "loadable_select" for f in self.fields)


def _wdts_python(script: str, *args: str) -> list[str]:
    """Run a script under wdts-altium-mcp with that repo's .venv Python."""
    python = WDTS_MCP_DIR / ".venv" / "Scripts" / "python.exe"
    return [str(python), str(WDTS_MCP_DIR / script), *args]


def _ps(script: Path, *args: str) -> list[str]:
    """Invoke a PowerShell script via -Command so we can force UTF-8 stdout.

    Windows PowerShell 5.1 fixes its stdout encoding at process start (before
    any in-script `[Console]::OutputEncoding = …` runs). Routing through a
    `-Command` line that sets the encoding *first* and then invokes the script
    gives us correct UTF-8 bytes for ±, °, etc.
    """
    quoted = [_ps_quote(a) for a in args]
    cmd = (
        "[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding $false; "
        f"& '{_ps_quote_path(script)}' " + " ".join(quoted) + "; exit $LASTEXITCODE"
    )
    return [
        "powershell.exe",
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-Command",
        cmd,
    ]


def _ps_quote_path(p: Path) -> str:
    return str(p).replace("'", "''")


def _ps_quote(s: str) -> str:
    # PowerShell parameter names (-Foo) must be passed unquoted, otherwise
    # PowerShell treats them as positional string values.
    if s.startswith("-") and not s.startswith("--") and " " not in s:
        return s
    return "'" + s.replace("'", "''") + "'"


def _force_upload_args(v: dict[str, str]) -> list[str]:
    fu = (v.get("force_upload") or "").strip().lower()
    if fu in ("1", "true", "yes"):
        return ["-ForceUpload"]
    return []


def _force_zip_overwrite_args(v: dict[str, str]) -> list[str]:
    fz = (v.get("force_zip_overwrite") or "").strip().lower()
    if fz in ("1", "true", "yes", "on"):
        return ["-ForceZipOverwrite"]
    return []


def _checkbox_on(values: dict[str, str], name: str, *, default: bool = False) -> bool:
    raw = (values.get(name) or "").strip().lower()
    if not raw:
        return default
    return raw in ("1", "true", "on", "yes")


def _optional_switch(values: dict[str, str], name: str, flag: str, *, default: bool = False) -> list[str]:
    if _checkbox_on(values, name, default=default):
        return [flag]
    return []


FUNCTIONS: list[FunctionSpec] = [
    # ----- Arena PLM -----
    FunctionSpec(
        id="arena_upload_all_working_rev_files",
        group="Arena PLM",
        icon="📤",
        title="Upload All Working Revision Files (not inc BOM)",
        description=(
            "For a PCBA part number and revision (letter + number, e.g. A1): upload existing "
            "Fabrication zip to the bare board, build/upload Pick & Place zip (revision letter "
            "must match Altium), sync Product Image, set PCBA working rev to A1 and bare-board "
            "working rev to the letter only, then run Check PCBA working-rev files."
        ),
        fields=[
            FieldSpec(
                name="part_number",
                label="PCBA part number",
                placeholder="239-10-001",
                required=True,
            ),
            FieldSpec(
                name="revision",
                label="Revision (letter + number)",
                placeholder="A1",
                required=True,
            ),
        ],
        background=False,
        build_command=lambda v: _ps(
            SCRIPTS_DIR / "Upload-AllArenaWorkingRevisionFiles.ps1",
            "-PartNumber",
            v["part_number"],
            "-Revision",
            (v.get("revision") or "").strip(),
        ),
    ),
    FunctionSpec(
        id="arena_sync_fabrication_gerbers",
        group="Arena PLM",
        icon="📦",
        hidden=True,
        duplicate_upload_prompt=True,
        title="Sync fabrication zip to Arena (Gerbers)",
        description=(
            "One step: resolve Altium outputs for the part, ensure "
            "'Fabrication PPP-20/25-YYY Rev <letter>.zip' exists (build it if "
            "missing), then on the bare-board Arena item's working revision "
            "either upload a new 'Altium Gerber Files' file (edition 1) or "
            "replace the existing Gerber file (edition +1). Enter PCBA or bare "
            "board number; Arena always targets the mapped PPP-20 / PPP-25 part."
        ),
        fields=[
            FieldSpec(
                name="part_number",
                label="Part number (Altium lookup)",
                placeholder="294-10-022 or 294-20-022",
                required=True,
            ),
        ],
        build_command=lambda v: _ps(
            SCRIPTS_DIR / "Sync-ArenaFabricationGerbers.ps1",
            "-PartNumber",
            v["part_number"],
            *_force_upload_args(v),
        ),
    ),
    FunctionSpec(
        id="arena_sync_pick_place_zip",
        group="Arena PLM",
        icon="📌",
        hidden=True,
        duplicate_upload_prompt=True,
        title="Sync Pick & Place zip to Arena",
        description=(
            "One step: resolve Altium outputs, ensure "
            "'Pick Place <canonical> Rev <rev>.zip' exists (build if missing), "
            "then on the PCBA Arena item's working revision (PPP-10 / PPP-15 — "
            "mapped from bare-board canonical if needed) either upload a new "
            "'Pick and Place' file (edition 1) or replace the existing one "
            "(edition +1)."
        ),
        fields=[
            FieldSpec(
                name="part_number",
                label="Part number (Altium lookup)",
                placeholder="294-10-022 or 294-20-022",
                required=True,
            ),
        ],
        build_command=lambda v: _ps(
            SCRIPTS_DIR / "Sync-ArenaPickPlaceZip.ps1",
            "-PartNumber",
            v["part_number"],
            *_force_upload_args(v),
        ),
    ),
    FunctionSpec(
        id="arena_sync_product_image",
        group="Arena PLM",
        icon="🖼️",
        duplicate_upload_prompt=True,
        title="Sync Arena Product Image",
        description=(
            "One step: resolve Altium production outputs, take the newest .png from "
            "the 'PCB 3D Print' folder under the production document folder, then on "
            "the PCBA Arena item's working revision (PPP-10 / PPP-15 — mapped from "
            "bare-board canonical if needed) either upload a new 'Illustrative Image' "
            "(edition 1, title 'Product Image <part> Rev <letter>') or replace the "
            "existing Illustrative Image (edition +1)."
        ),
        fields=[
            FieldSpec(
                name="part_number",
                label="Part number (Altium lookup)",
                placeholder="294-10-022 or 294-20-022",
                required=True,
            ),
        ],
        build_command=lambda v: _ps(
            SCRIPTS_DIR / "Sync-ArenaProductImage.ps1",
            "-PartNumber",
            v["part_number"],
            *_force_upload_args(v),
        ),
    ),
    FunctionSpec(
        id="compare_bom_auto",
        group="Arena PLM",
        icon="🍒",
        title="Compare Arena vs Altium BOM v2",
        description=(
            "Takes only a part number: resolves the newest BOM .xlsx/.xls, "
            "fetches the Arena BOM, and diffs qty/designators plus spec-aware "
            "Arena item names vs Altium descriptions (voltage, current, "
            "footprint, passive values). Also checks sheet 2 (01-PJM Part_Number) "
            "when present."
        ),
        fields=[
            FieldSpec(
                name="part_number",
                label="Part number",
                placeholder="272-15-001",
                required=True,
            ),
            FieldSpec(
                name="sheet",
                label="Altium sheet (name or 1-based index)",
                default="1",
            ),
        ],
        build_command=lambda v: _ps(
            SCRIPTS_DIR / "Compare-ArenaAltiumBomAuto.ps1",
            "-PartNumber", v["part_number"],
            *(["-Sheet", v["sheet"]] if (v.get("sheet") or "").strip() else []),
        ),
    ),
    FunctionSpec(
        id="arena_check_upload_altium_bom_export",
        group="Arena PLM",
        icon="✅",
        hidden=True,
        title="Check and Create CSV Altium BOM Export",
        description=(
            "Uses the newest BOM workbook in the production BOM folder: on the PJM "
            "sheet (default sheet 2), each 01-PJM Part_Number is looked up in Arena as an "
            "item (not via assembly BOM). Arena item description is used when set, "
            "otherwise item name, and compared to 03-Description General with the same "
            "spec-aware rules as BOM compare v2. When Excel Name/Comment matches Arena manufacturer "
            "part number or item name, description checks for that line are skipped. On that same sheet, when a Name "
            "or Comment column exists, Name/Comment vs 03-Description General is checked only for "
            "098-12-xxx and 098-13-xxx part numbers, with the same rules (unless Name/Comment matched identity). "
            "Every BOM line part is checked even if the "
            "assembly has no Arena BOM or that part is missing from the BOM. Then runs "
            "'Convert BOM (Excel -> CSV) v2' only if all lines pass."
        ),
        fields=[
            FieldSpec(
                name="part_number",
                label="Part number",
                placeholder="241-10-001",
                required=True,
            ),
        ],
        background=False,
        proceed_after_fail=ProceedAfterFailSpec(
            label="Proceed with CSV Creation?",
            target_id="convert_bom_to_csv_auto",
        ),
        build_command=lambda v: _ps(
            SCRIPTS_DIR / "Check-AndUploadAltiumBomExport.ps1",
            "-PartNumber",
            v["part_number"],
        ),
    ),
    FunctionSpec(
        id="arena_check_pcba_files",
        group="Arena PLM",
        icon="⚽",
        title="Check PCBA working-rev files",
        description=(
            "For a PCBA part number (PPP-10-XXX or PPP-15-XXX), verify the "
            "working revision has a Pick and Place + Illustrative Image "
            "file, and that the matching bare board (PPP-20-XXX / PPP-25-XXX) "
            "has an Altium Gerber Files file. Each file's title must end "
            "with 'Rev <X>' matching the working revision."
        ),
        fields=[
            FieldSpec(
                name="part_number",
                label="PCBA part number",
                placeholder="251-10-001 or 251-15-001",
                required=True,
            ),
            FieldSpec(
                name="expected_revision",
                label="Expected revision letter (optional override)",
                placeholder="leave blank to auto-detect from working rev",
            ),
        ],
        build_command=lambda v: _ps(
            SCRIPTS_DIR / "Check-ArenaPcbaFiles.ps1",
            "-PartNumber", v["part_number"],
            *(["-ExpectedRevision", v["expected_revision"]] if v.get("expected_revision") else []),
        ),
    ),
    FunctionSpec(
        id="arena_overwrite_file",
        group="Arena PLM",
        icon="🚲",
        title="Overwrite File on Working Revision",
        description=(
            "Replace an existing file on an item's working revision with a "
            "new edition. Click 'Load files' first to populate the picker, "
            "then choose a replacement. The new title is auto-set to the "
            "replacement filename and the edition is bumped by 1."
        ),
        fields=[
            FieldSpec(
                name="part_number",
                label="Part number",
                placeholder="294-10-001",
                required=True,
            ),
            FieldSpec(
                name="existing_file",
                label="Existing file on working rev",
                kind="loadable_select",
                required=True,
                discovery_meta=["edition", "title"],
            ),
            FieldSpec(
                name="replacement",
                label="Replacement file",
                kind="file",
                required=True,
            ),
        ],
        discovery=DiscoverySpec(
            build_command=lambda v: _ps(
                SCRIPTS_DIR / "Get-ArenaItemFiles.ps1",
                "-PartNumber", v["part_number"],
                "-Revision", "working",
            ),
        ),
        build_command=lambda v: _ps(
            SCRIPTS_DIR / "Replace-ArenaFile.ps1",
            "-FileGuid",       v["existing_file"],
            "-FilePath",       v["replacement"],
            "-CurrentEdition", v.get("existing_file__edition", ""),
        ),
    ),
    FunctionSpec(
        id="arena_upload_fab_test",
        group="Arena PLM",
        icon="🍕",
        title="Upload New File",
        description=(
            "Upload a file to an Arena item revision. Resolves the working "
            "revision (default), uploads the file with the chosen category, "
            "and associates it with that revision."
        ),
        fields=[
            FieldSpec(
                name="part_number",
                label="Part number",
                placeholder="294-10-001",
                required=True,
            ),
            FieldSpec(
                name="revision",
                label="Revision to attach to",
                kind="select",
                options=["working", "effective"],
                default="working",
            ),
            FieldSpec(
                name="file",
                label="File to upload",
                kind="file",
                required=True,
            ),
            FieldSpec(
                name="title",
                label="Title (shown in Arena, auto-fills from filename)",
                placeholder="auto-filled when you pick a file",
                required=True,
            ),
            FieldSpec(
                name="category",
                label="File category",
                kind="select",
                options=[
                    "Altium Gerber Files",
                    "Illustrative Image",
                    "Pick and Place",
                    "Board Stuffing Specification",
                ],
                default="Altium Gerber Files",
            ),
            FieldSpec(
                name="edition",
                label="Edition",
                default="1",
            ),
            FieldSpec(
                name="description",
                label="Description",
                kind="textarea",
                placeholder="Optional notes about this upload",
            ),
            FieldSpec(
                name="author",
                label="Author full name (optional override)",
                placeholder="blank = Windows session user (display name or DOMAIN\\user)",
            ),
        ],
        build_command=lambda v: _ps(
            SCRIPTS_DIR / "Upload-ArenaFile.ps1",
            "-PartNumber", v["part_number"],
            "-FilePath",   v["file"],
            "-Title",      v["title"],
            "-Revision",   v.get("revision") or "working",
            "-CategoryName", v.get("category") or "Test Specification",
            "-Edition",    v.get("edition") or "1",
            *(["-Description", v["description"]] if v.get("description") else []),
            *(["-Author", v["author"]] if v.get("author") else []),
        ),
    ),
    FunctionSpec(
        id="arena_check_pick_place_with_bom",
        group="Arena PLM",
        icon="🔍",
        title="Check Pick and Place File with Arena BOM",
        description=(
            "For a part number, reads the single .csv in the production Pick Place folder "
            "and compares each designator's Description (or Comment if Description is empty) "
            "to the Arena assembly BOM on the working or most recent released revision. "
            "Uses the same spec-aware regex checks as BOM export / Arena–Altium compare. "
            "N/F or DNF in P&P is OK when Arena also marks N/F or the line is absent; "
            "the PCB board item may be missing from P&P; fiducials and mounting holes are skipped. "
            "Read-only; does not modify files."
        ),
        fields=[
            FieldSpec(
                name="part_number",
                label="Part number (Altium lookup)",
                placeholder="272-15-001",
                required=True,
            ),
            FieldSpec(
                name="revision_mode",
                label="Arena BOM revision",
                kind="select",
                options=[
                    "Working revision",
                    "Most recent released revision",
                ],
                default="Working revision",
            ),
        ],
        build_command=lambda v: _ps(
            SCRIPTS_DIR / "Check-PickPlaceFileWithArenaBom.ps1",
            "-PartNumber",
            v["part_number"],
            "-RevisionMode",
            (
                "effective_latest"
                if (v.get("revision_mode") or "").strip()
                == "Most recent released revision"
                else "working"
            ),
        ),
    ),
    FunctionSpec(
        id="arena_correct_pick_place_from_bom",
        group="Arena PLM",
        icon="📍",
        title="Correct P&P for Old Project from Arena BOM",
        description=(
            "For a part number, locates the single .csv and .txt in the production "
            "Pick Place folder, fetches the Arena assembly BOM (working or most recent "
            "released revision), and updates both files: "
            "Description from Arena (item description, else name); Comment from "
            "PJM_Component_Database_Lib.xls when 01-PJM Part_Number matches, else keep "
            "existing comment if it reasonably matches Arena text, else use Arena description. "
            "Fiducial lines are not changed. Designators on Arena Not Fitted / N/F / DNF "
            "lines are removed. Original files are saved as .bak under "
            "'Pick Place Archive - DO NOT USE' in the production folder."
        ),
        fields=[
            FieldSpec(
                name="part_number",
                label="Part number (Altium lookup)",
                placeholder="272-15-001",
                required=True,
            ),
            FieldSpec(
                name="revision_mode",
                label="Arena BOM revision",
                kind="select",
                options=[
                    "Working revision",
                    "Most recent released revision",
                ],
                default="Working revision",
            ),
        ],
        build_command=lambda v: _ps(
            SCRIPTS_DIR / "Correct-PickPlaceFromArenaBom.ps1",
            "-PartNumber",
            v["part_number"],
            "-RevisionMode",
            (
                "effective_latest"
                if (v.get("revision_mode") or "").strip()
                == "Most recent released revision"
                else "working"
            ),
        ),
    ),
    FunctionSpec(
        id="arena_find_pcb_assemblies",
        group="Arena PLM",
        icon="🍉",
        title="Find PCB assemblies in BOM",
        description=(
            "Recursively walk the BOM of a top-level system and list every "
            "unique PCB assembly used at any depth. A PCBA is any item "
            "matching PPP-10-YYY or PPP-15-YYY where PPP is not 098."
        ),
        fields=[
            FieldSpec(
                name="part_number",
                label="Top-level part number",
                placeholder="260-00-001",
                required=True,
            ),
            FieldSpec(
                name="format",
                label="Format",
                kind="select",
                options=["table", "json", "csv"],
                default="table",
            ),
            FieldSpec(
                name="max_depth",
                label="Max recursion depth",
                default="20",
            ),
        ],
        build_command=lambda v: _ps(
            SCRIPTS_DIR / "Find-ArenaPcbAssemblies.ps1",
            "-PartNumber", v["part_number"],
            "-Format", v.get("format") or "table",
            *(["-MaxDepth", v["max_depth"]] if (v.get("max_depth") or "").strip() else []),
        ),
    ),
    FunctionSpec(
        id="arena_bom",
        group="Arena PLM",
        icon="🍇",
        hidden=True,
        title="Get Arena BOM",
        description="Print the immediate child items (BOM) of an Arena assembly by part number.",
        fields=[
            FieldSpec(
                name="part_number",
                label="Part number",
                placeholder="294-10-001",
                required=True,
            ),
            FieldSpec(
                name="format",
                label="Format",
                kind="select",
                options=["table", "json", "csv"],
                default="table",
            ),
        ],
        build_command=lambda v: _ps(
            SCRIPTS_DIR / "Get-ArenaBom.ps1",
            "-PartNumber",
            v["part_number"],
            "-Format",
            v.get("format") or "table",
        ),
    ),
    FunctionSpec(
        id="arena_export_supplier_bom",
        group="Arena PLM",
        icon="📤",
        title="Export Supplier BOM",
        description=(
            "Export an Arena assembly BOM to Excel using the SUPPLIER format "
            "(effective latest revision by default). Saves to your Documents folder."
        ),
        fields=[
            FieldSpec(
                name="part_number",
                label="Part number",
                placeholder="294-10-001",
                required=True,
            ),
            FieldSpec(
                name="revision_mode",
                label="Revision",
                kind="select",
                options=["effective_latest", "working"],
                default="effective_latest",
            ),
            FieldSpec(
                name="bom_levels",
                label="BOM levels",
                kind="select",
                options=["FULL", "SINGLE"],
                default="FULL",
            ),
        ],
        build_command=lambda v: _ps(
            SCRIPTS_DIR / "Export-ArenaSupplierBom.ps1",
            "-PartNumber",
            v["part_number"],
            "-RevisionMode",
            v.get("revision_mode") or "effective_latest",
            "-BomLevels",
            v.get("bom_levels") or "FULL",
        ),
    ),
    FunctionSpec(
        id="compare_bom",
        group="Arena PLM",
        icon="🍋",
        hidden=True,
        title="OBSOLETE - Compare Arena vs Altium BOM",
        description=(
            "Legacy manual path: diff an Arena BOM CSV file against an Altium "
            "BOM xlsx (pick sheet). Prefer 'Compare Arena vs Altium BOM v2' or "
            "'BOM export check and CSV' on the Altium tab for part-number-driven flows."
        ),
        fields=[
            FieldSpec(
                name="arena_csv",
                label="Arena CSV",
                placeholder=r"C:\path\to\arena_export.csv",
                required=True,
            ),
            FieldSpec(
                name="altium_xlsx",
                label="Altium xlsx",
                placeholder=r"C:\path\to\altium_bom.xlsx",
                required=True,
            ),
            FieldSpec(
                name="sheet",
                label="Altium sheet (name or 1-based index)",
                default="2",
            ),
        ],
        build_command=lambda v: _ps(
            SCRIPTS_DIR / "Compare-ArenaAltiumBom.ps1",
            "-ArenaCsv",
            v["arena_csv"],
            "-AltiumXlsx",
            v["altium_xlsx"],
            *(["-Sheet", v["sheet"]] if (v.get("sheet") or "").strip() else []),
        ),
    ),
    # ----- Altium -----
    FunctionSpec(
        id="altium_open_explorer",
        group="Altium",
        icon="🍓",
        title="Open Altium part folder",
        description=(
            "Resolve PPP-XX-XXX -> Altium folder. Open Main Folder opens the design "
            "(latest Rev/Ver). Open Prod Folder opens the production outputs area "
            '("Production Document…", "Project Outputs for …", or a child with Gerber/BOM).'
        ),
        fields=[
            FieldSpec(
                name="part_number",
                label="Part number",
                placeholder="272-10-001 or 272-10-001 BJ 5P",
                required=True,
            ),
        ],
        run_actions=(
            RunActionSpec(id="main", label="Open Main Folder"),
            RunActionSpec(id="prod", label="Open Prod Folder"),
        ),
        # Foreground so the launcher captures stdout (you can confirm the
        # script actually ran) and so the new Explorer window can grab focus
        # without being blocked by Windows' detached-process focus rules.
        background=False,
        build_command=lambda v: _ps(
            SCRIPTS_DIR / "Open-AltiumPartFolder.ps1",
            "-PartNumber", v["part_number"],
            "-Open",
            "ProductionExplorer"
            if (v.get("__run_action") or "").strip() == "prod"
            else "Explorer",
        ),
    ),
    FunctionSpec(
        id="altium_open_prjpcb",
        group="Altium",
        icon="🥝",
        title="Open Altium .PrjPcb project",
        description=(
            "Resolve PPP-XX-XXX and open the most recently modified .PrjPcb under the "
            "latest Rev/Ver folder, or use Open PCB Only for the newest .PcbDoc in that "
            "folder (same Archive exclusion)."
        ),
        fields=[
            FieldSpec(
                name="part_number",
                label="Part number",
                placeholder="272-10-001 or 272-10-001 BJ 5P",
                required=True,
            ),
        ],
        run_actions=(
            RunActionSpec(id="prjpcb", label="Open .PrjPcb project"),
            RunActionSpec(id="pcb_only", label="Open PCB Only"),
        ),
        background=False,
        build_command=lambda v: _ps(
            SCRIPTS_DIR / "Open-AltiumPartFolder.ps1",
            "-PartNumber", v["part_number"],
            "-Open",
            "PcbDoc"
            if (v.get("__run_action") or "").strip() == "pcb_only"
            else "PrjPcb",
        ),
    ),
    FunctionSpec(
        id="altium_fabrication_zip",
        group="Altium",
        icon="🚀",
        hidden=True,
        title="Check and Build Fabrication Zip",
        description="(Superseded by combined Fabrication zip card.)",
        fields=[
            FieldSpec(
                name="part_number",
                label="Part number",
                placeholder="272-10-001",
                required=True,
            ),
        ],
        background=False,
        build_command=lambda v: _ps(
            SCRIPTS_DIR / "Open-AltiumPartFolder.ps1",
            "-PartNumber", v["part_number"],
            "-Open", "FabricationZip",
        ),
    ),
    FunctionSpec(
        id="altium_fabrication_zip_workflow",
        group="Altium",
        icon="🚀",
        duplicate_upload_prompt=True,
        zip_overwrite_prompt=True,
        title="Fabrication zip",
        description=(
            "Check fabrication outputs and optionally create "
            "'Fabrication PPP-20/25-YYY Rev <letter>.zip', then optionally upload or "
            "replace 'Altium Gerber Files' on the bare-board Arena working revision. "
            "Use checkboxes to run only the steps you need."
        ),
        fields=[
            FieldSpec(
                name="part_number",
                label="Part number",
                placeholder="294-10-022 or 294-20-022",
                required=True,
            ),
            FieldSpec(
                name="create_zip",
                label="Create Zip",
                kind="checkbox",
                default="1",
            ),
            FieldSpec(
                name="upload_to_arena",
                label="Upload Zip File to Arena",
                kind="checkbox",
                default="",
            ),
            FieldSpec(
                name="use_existing_file",
                label="Use Existing File",
                kind="checkbox",
                default="",
            ),
        ],
        background=False,
        build_command=lambda v: _ps(
            SCRIPTS_DIR / "Invoke-AltiumFabricationZipWorkflow.ps1",
            "-PartNumber",
            v["part_number"],
            *_optional_switch(v, "create_zip", "-CreateZip", default=True),
            *_optional_switch(v, "upload_to_arena", "-UploadToArena"),
            *_optional_switch(v, "use_existing_file", "-UseExistingFile"),
            *_force_upload_args(v),
            *_force_zip_overwrite_args(v),
        ),
    ),
    FunctionSpec(
        id="altium_gerber_zip",
        group="Altium",
        icon="📦",
        hidden=True,
        title="Check and Build Gerber Zip",
        description="(Superseded by combined Gerber zip card.)",
        fields=[
            FieldSpec(
                name="part_number",
                label="Part number",
                placeholder="272-10-001",
                required=True,
            ),
        ],
        background=False,
        build_command=lambda v: _ps(
            SCRIPTS_DIR / "Open-AltiumPartFolder.ps1",
            "-PartNumber", v["part_number"],
            "-Open", "GerberZip",
        ),
    ),
    FunctionSpec(
        id="altium_gerber_zip_workflow",
        group="Altium",
        icon="📦",
        zip_overwrite_prompt=True,
        title="Gerber zip",
        description=(
            "Check Gerber and NC Drill outputs and optionally create "
            "'Gerbers PPP-20/25-YYY Rev <letter>.zip' (no PFS folder). "
            "Same checks as fabrication zip except PFS is not required or included."
        ),
        fields=[
            FieldSpec(
                name="part_number",
                label="Part number",
                placeholder="294-10-022 or 294-20-022",
                required=True,
            ),
            FieldSpec(
                name="create_zip",
                label="Create Zip",
                kind="checkbox",
                default="1",
            ),
        ],
        background=False,
        build_command=lambda v: _ps(
            SCRIPTS_DIR / "Invoke-AltiumGerberZipWorkflow.ps1",
            "-PartNumber",
            v["part_number"],
            *_optional_switch(v, "create_zip", "-CreateZip", default=True),
            *_force_zip_overwrite_args(v),
        ),
    ),
    FunctionSpec(
        id="altium_bom_export_workflow",
        group="Altium",
        icon="✅",
        title="BOM export check and CSV",
        description=(
            "On the PJM sheet (default sheet 2), checks each 01-PJM Part_Number against "
            "Arena (item description or name) and optional Name vs 03-Description General "
            "for 098-12/13 lines — same rules as the former Arena BOM export check. "
            "When Create CSV is enabled, runs Convert BOM (Excel -> CSV) v2 after all "
            "checks pass. If checks fail, you can still choose Proceed with CSV Creation."
        ),
        fields=[
            FieldSpec(
                name="part_number",
                label="Part number",
                placeholder="241-10-001",
                required=True,
            ),
            FieldSpec(
                name="create_csv",
                label="Create CSV",
                kind="checkbox",
                default="1",
            ),
        ],
        background=False,
        proceed_after_fail=ProceedAfterFailSpec(
            label="Proceed with CSV Creation?",
            target_id="convert_bom_to_csv_auto",
        ),
        build_command=lambda v: _ps(
            SCRIPTS_DIR / "Check-AndUploadAltiumBomExport.ps1",
            "-PartNumber",
            v["part_number"],
            *_optional_switch(v, "create_csv", "-CreateCsv", default=True),
        ),
    ),
    FunctionSpec(
        id="altium_pickplace_zip",
        group="Altium",
        icon="🏀",
        hidden=True,
        title="Build Pick & Place zip",
        description="(Superseded by combined Pick & Place zip card.)",
        fields=[
            FieldSpec(
                name="part_number",
                label="Part number",
                placeholder="272-10-001",
                required=True,
            ),
        ],
        background=False,
        build_command=lambda v: _ps(
            SCRIPTS_DIR / "Open-AltiumPartFolder.ps1",
            "-PartNumber", v["part_number"],
            "-Open", "PickPlaceZip",
        ),
    ),
    FunctionSpec(
        id="altium_pickplace_zip_workflow",
        group="Altium",
        icon="🏀",
        duplicate_upload_prompt=True,
        zip_overwrite_prompt=True,
        title="Pick & Place zip",
        description=(
            "Optionally zip the production 'Pick Place' folder into "
            "'Pick Place PPP-TT-XXX Rev <letter><number>.zip', then optionally upload or "
            "replace 'Pick and Place' on the PCBA Arena working revision. Revision entry "
            "must match the Altium revision letter; the number you enter is used in the zip name."
        ),
        fields=[
            FieldSpec(
                name="part_number",
                label="Part number",
                placeholder="294-10-022 or 294-20-022",
                required=True,
            ),
            FieldSpec(
                name="revision",
                label="Revision (letter + number)",
                placeholder="D1",
                required=True,
            ),
            FieldSpec(
                name="create_zip",
                label="Create Zip",
                kind="checkbox",
                default="1",
            ),
            FieldSpec(
                name="upload_to_arena",
                label="Upload Zip File to Arena",
                kind="checkbox",
                default="",
            ),
            FieldSpec(
                name="use_existing_file",
                label="Use Existing File",
                kind="checkbox",
                default="",
            ),
        ],
        background=False,
        build_command=lambda v: _ps(
            SCRIPTS_DIR / "Invoke-AltiumPickPlaceZipWorkflow.ps1",
            "-PartNumber",
            v["part_number"],
            "-RevisionOverride",
            (v.get("revision") or "").strip(),
            *_optional_switch(v, "create_zip", "-CreateZip", default=True),
            *_optional_switch(v, "upload_to_arena", "-UploadToArena"),
            *_optional_switch(v, "use_existing_file", "-UseExistingFile"),
            *_force_upload_args(v),
            *_force_zip_overwrite_args(v),
        ),
    ),
    # ----- Altium MCP (live bridge) -----
    FunctionSpec(
        id="altium_mcp_get_designators",
        group="Altium MCP",
        icon="🏷️",
        title="Get Designators",
        description=(
            "Resolve the part number to the Altium PCB/project, query live component "
            "designators via the WDTS bridge, and print a sorted list. If the bridge "
            "times out, falls back to the newest Pick & Place CSV under the part folder."
        ),
        fields=[
            FieldSpec(
                name="part_number",
                label="Part number",
                placeholder="213-15-007 or 213-25-007",
                required=True,
            ),
        ],
        background=False,
        build_command=lambda v: _ps(
            SCRIPTS_DIR / "Get-AltiumDesignators.ps1",
            "-PartNumber", v["part_number"],
        ),
    ),
    FunctionSpec(
        id="altium_mcp_check_silkscreen_text",
        group="Altium MCP",
        icon="🔍",
        title="Check Silkscreen Text",
        description=(
            "Checks Top and Bottom Overlay silkscreen via the live Altium bridge: "
            "bare PCB and assembly part numbers with Rev/Ver, copyright text (or ©), "
            "and a 25 mm x 10 mm serial number box. Requires compiled Altium_API.PrjScr "
            "(or other_utils » MCP_ProcessRequest if Altium is already open)."
        ),
        fields=[
            FieldSpec(
                name="part_number",
                label="Part number",
                placeholder="213-15-007 or 213-25-007",
                required=True,
            ),
        ],
        background=False,
        cwd=WDTS_MCP_DIR,
        build_command=lambda v: _wdts_python(
            "scripts/check_silkscreen_text.py",
            "--part", v["part_number"],
        ),
    ),
    FunctionSpec(
        id="altium_mcp_check_silkscreen_part_number",
        group="Altium MCP",
        icon="🔍",
        hidden=True,
        title="Check Top Silkscreen Part Number",
        description="(Renamed to Check Silkscreen Text.)",
        fields=[
            FieldSpec(
                name="part_number",
                label="Part number",
                placeholder="213-15-007 or 213-25-007",
                required=True,
            ),
        ],
        background=False,
        cwd=WDTS_MCP_DIR,
        build_command=lambda v: _wdts_python(
            "scripts/check_silkscreen_text.py",
            "--part", v["part_number"],
        ),
    ),
    FunctionSpec(
        id="altium_mcp_find_untented_vias",
        group="Altium MCP",
        icon="◎",
        title="Find Untented Vias",
        description=(
            "Opens the part PCB, finds vias not tented on Top or Bottom solder mask, "
            "deselects everything else, selects those vias, and zooms. Altium shows "
            "\"No untented vias\" if none match. Uses WDTS Altium MCP bridge "
            "(compile Altium_API.PrjScr or run other_utils » MCP_ProcessRequest)."
        ),
        fields=[
            FieldSpec(
                name="part_number",
                label="Part number",
                placeholder="213-15-007 or 213-25-007",
                required=True,
            ),
        ],
        background=False,
        cwd=WDTS_MCP_DIR,
        build_command=lambda v: _wdts_python(
            "scripts/show_untented_vias.py",
            "--part", v["part_number"],
        ),
    ),
    FunctionSpec(
        id="show_untented_vias",
        group="Altium",
        icon="◎",
        hidden=True,
        title="Show untented vias",
        description="(Moved to Altium MCP tab.)",
        fields=[],
        build_command=lambda v: _ps(
            SCRIPTS_DIR / "Show-UntentedVias.ps1",
        ),
    ),
    FunctionSpec(
        id="check_part_number_silkscreen",
        group="Altium",
        icon="🔍",
        hidden=True,
        title="Check part number on silkscreen",
        description="(Moved to Altium MCP tab.)",
        fields=[],
        build_command=lambda v: _ps(
            SCRIPTS_DIR / "Check-PartNumberSilkscreen.ps1",
        ),
    ),
    FunctionSpec(
        id="altium_copy_checklist_into_folder",
        group="Altium",
        icon="📋",
        title="Copy Checklist into Folder",
        description=(
            "Copies the newest 'WDTS - PCB Design Checklist vX.X' template from the "
            "Altium Library into the part's main folder (same root as Open Altium part folder), "
            "then opens the copy in Excel."
        ),
        fields=[
            FieldSpec(
                name="part_number",
                label="Part number",
                placeholder="213-15-007 or 272-10-001 BJ 5P",
                required=True,
            ),
        ],
        background=False,
        build_command=lambda v: _ps(
            SCRIPTS_DIR / "Copy-ChecklistIntoAltiumPartFolder.ps1",
            "-PartNumber",
            v["part_number"],
        ),
    ),
    FunctionSpec(
        id="convert_bom_to_csv",
        group="Altium",
        icon="🍊",
        title="Convert BOM (Excel -> CSV)",
        description="Pick an Altium BOM xlsx via the Excel file dialog; saves cleaned CSV beside it.",
        background=True,
        hidden=True,
        build_command=lambda v: _ps(ALTIUM_BOM_TO_CSV),
    ),
    FunctionSpec(
        id="convert_bom_to_csv_auto",
        group="Altium",
        icon="🍑",
        hidden=True,
        title="Convert BOM (Excel -> CSV) v2",
        description="(Superseded by BOM export check and CSV card.)",
        fields=[
            FieldSpec(
                name="part_number",
                label="Part number",
                placeholder="272-15-001",
                required=True,
            ),
        ],
        build_command=lambda v: _ps(
            SCRIPTS_DIR / "Convert-AltiumBomAuto.ps1",
            "-PartNumber", v["part_number"],
        ),
    ),
    FunctionSpec(
        id="compare_excel_altium_library_arena",
        group="Altium",
        icon="📊",
        title="Compare Excel Altium Library with Arena",
        description=(
            "Upload a multi-sheet Excel library workbook: on each sheet, the 01-PJM Part_Number "
            "column is the key and 03-Description General (or the same fallback description columns "
            "as 'BOM export check and CSV') is compared to the Arena item description "
            "(when set, else item name) using "
            "the same spec-aware checks (voltage, current, footprint, passive values, etc.). "
            "Produces a new workbook with the same sheet names and order; each sheet lists only "
            "Part Number and Discrepancy for failing lines. Sheets without PJM headers appear with "
            "headers only. Default report path: Desktop, ``<file>_ArenaDiscrepancies_<timestamp>.xlsx``."
        ),
        fields=[
            FieldSpec(
                name="library_workbook",
                label="Excel library workbook (.xlsx / .xls)",
                kind="file",
                required=True,
                accept=".xlsx,.xls",
            ),
            FieldSpec(
                name="report_path",
                label="Report save path (.xlsx); leave blank for Desktop + timestamp",
                placeholder=r"C:\Users\you\Desktop\library_ArenaDiscrepancies.xlsx",
            ),
        ],
        build_command=lambda v: _ps(
            SCRIPTS_DIR / "Compare-ExcelAltiumLibraryWithArena.ps1",
            "-ExcelPath",
            v["library_workbook"],
            *(
                ["-OutputPath", v["report_path"].strip()]
                if (v.get("report_path") or "").strip()
                else []
            ),
        ),
    ),
    # Last in launcher list order (Arena PLM tab): infrequent / advanced.
    FunctionSpec(
        id="arena_duplicate_item_from_existing",
        group="Arena PLM",
        icon="📑",
        title="Duplicate Arena Item from Existing Item (untested)",
        description=(
            "Creates a new Arena item from an existing one: enter the old and new part numbers, "
            "use 'Load name & description from Arena' to pull the source item name and description into "
            "editable fields, adjust them, choose whether to copy BOM lines and/or file associations, "
            "and whether to read the source from the working revision, the latest effective (released) "
            "revision, or a specific revision letter in the override box. Set the new item's starting "
            "revision number. Owner defaults to your Windows display name (override optional). "
            "Untested — verify results in Arena."
        ),
        fields=[
            FieldSpec(
                name="old_part_number",
                label="Existing (source) part number",
                placeholder="294-10-001",
                required=True,
            ),
            FieldSpec(
                name="new_part_number",
                label="New part number",
                placeholder="294-10-099",
                required=True,
            ),
            FieldSpec(
                name="item_name",
                label="Item name (edit before run)",
                kind="textarea",
                placeholder="Loaded from Arena or type here",
                required=True,
            ),
            FieldSpec(
                name="item_description",
                label="Item description (edit before run)",
                kind="textarea",
                placeholder="Loaded from Arena or type here",
            ),
            FieldSpec(
                name="copy_bom",
                label="Copy BOM from source",
                kind="select",
                options=["no", "yes"],
                default="no",
            ),
            FieldSpec(
                name="copy_files",
                label="Copy file associations from source revision",
                kind="select",
                options=["no", "yes"],
                default="no",
            ),
            FieldSpec(
                name="source_revision_mode",
                label="Source revision (when override below is empty)",
                kind="select",
                options=["working", "effective_latest"],
                default="working",
            ),
            FieldSpec(
                name="source_revision_override",
                label="Source revision override (e.g. B; leave blank to use mode above)",
                placeholder="",
            ),
            FieldSpec(
                name="new_item_revision",
                label="New item starting revision number",
                default="A",
            ),
            FieldSpec(
                name="owner_full_name",
                label="Owner full name (blank = Windows session display name)",
                placeholder="",
            ),
        ],
        discovery=DiscoverySpec(
            build_command=lambda v: _ps(
                SCRIPTS_DIR / "Duplicate-ArenaItemFromExisting.ps1",
                "-Mode",
                "Preview",
                "-OldPartNumber",
                v["old_part_number"],
            ),
        ),
        discovery_prefill=True,
        build_command=lambda v: _ps(
            SCRIPTS_DIR / "Duplicate-ArenaItemFromExisting.ps1",
            "-OldPartNumber",
            v["old_part_number"],
            "-NewPartNumber",
            v["new_part_number"],
            "-ItemName",
            v["item_name"],
            "-ItemDescription",
            v.get("item_description") or "",
            "-CopyBom",
            v.get("copy_bom") or "no",
            "-CopyFiles",
            v.get("copy_files") or "no",
            "-SourceRevisionMode",
            v.get("source_revision_mode") or "working",
            *(
                ["-SourceRevisionOverride", v["source_revision_override"].strip()]
                if (v.get("source_revision_override") or "").strip()
                else []
            ),
            "-NewItemRevision",
            (v.get("new_item_revision") or "A").strip(),
            *(
                ["-OwnerFullName", v["owner_full_name"].strip()]
                if (v.get("owner_full_name") or "").strip()
                else []
            ),
        ),
    ),
]


def find(function_id: str) -> FunctionSpec | None:
    for fn in FUNCTIONS:
        if fn.id == function_id:
            return fn
    return None
