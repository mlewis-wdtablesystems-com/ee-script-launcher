import {
  Button,
  Callout,
  Card,
  CardBody,
  CardHeader,
  Checkbox,
  Code,
  Divider,
  Grid,
  H1,
  H2,
  H3,
  Pill,
  Row,
  Stack,
  Stat,
  Table,
  Text,
  TextArea,
  TextInput,
  useCanvasAction,
  useCanvasState,
} from "cursor/canvas";
import { type KeyboardEvent, useEffect, useState } from "react";

type WorkbenchPane = "overview" | "altium" | "sharepoint" | "arena" | "bomComparison";

/** Default conversation that compared Arena CSV vs Altium ENG BOM (292-15-001). Override in the canvas if you fork the agent. */
const DEFAULT_BOM_AGENT_ID = "40e0bf4b-5df2-4050-bcc9-19d399fa77dc";

const PANES: { id: WorkbenchPane; label: string }[] = [
  { id: "overview", label: "Overview" },
  { id: "altium", label: "Altium" },
  { id: "sharepoint", label: "SharePoint" },
  { id: "arena", label: "Arena PLM" },
  { id: "bomComparison", label: "Arena vs Altium BOM" },
];

const BOM_WORKFLOW_ROWS: [string, string][] = [
  ["Arena", "Export assembly BOM as CSV (UTF-8). Note internal part number and quantity columns."],
  ["Altium", "Project Outputs: BOM to Excel. Use the engineering sheet (often ENG BOM or sheet 2) that matches how you review parts."],
  ["Align", "Pick a stable key: internal MPN, manufacturer + MPN, or designator set — same as in your agent prompt."],
  ["Agent", "Open the comparison agent and attach or paste paths to both files. Review only-in-Arena, only-in-Altium, and field mismatches."],
];

const ALTIUM_LIBRARY_ROOT =
  "C:\\Users\\mlewis\\OneDrive - Walker Digital\\WTDS-RDAustralia - Documents\\23. Altium";
const SCRIPTS_FOLDER =
  "C:\\Users\\mlewis\\.cursor\\projects\\empty-window\\scripts";
const OPEN_ALTIUM_FOLDER_SCRIPT = `${SCRIPTS_FOLDER}\\Open-AltiumPartFolder.ps1`;
const COMPARE_ARENA_ALTIUM_BOM_SCRIPT = `${SCRIPTS_FOLDER}\\Compare-ArenaAltiumBom.ps1`;
const GET_ARENA_ITEM_SCRIPT = `${SCRIPTS_FOLDER}\\Get-ArenaItem.ps1`;
const SHOW_ARENA_PLM_DASHBOARD_SCRIPT = `${SCRIPTS_FOLDER}\\Show-ArenaPlmDashboard.ps1`;

const CONVERT_BOM_TO_CSV_SCRIPT =
  "C:\\Users\\mlewis\\OneDrive - Walker Digital\\Documents\\Convert_BOM_to_CSV.ps1";

const SCRIPT_LINKS: { path: string; label: string; description: string }[] = [
  {
    path: SHOW_ARENA_PLM_DASHBOARD_SCRIPT,
    label: "Show-ArenaPlmDashboard.ps1",
    description: "Small WinForms window: Item info + Get BOM by part number.",
  },
  {
    path: GET_ARENA_ITEM_SCRIPT,
    label: "Get-ArenaItem.ps1",
    description: "CLI: print full item record for a part number via Arena REST API.",
  },
  {
    path: COMPARE_ARENA_ALTIUM_BOM_SCRIPT,
    label: "Compare-ArenaAltiumBom.ps1",
    description: "Diff Arena BOM CSV against an Altium BOM xlsx sheet.",
  },
  {
    path: OPEN_ALTIUM_FOLDER_SCRIPT,
    label: "Open-AltiumPartFolder.ps1",
    description: "Resolve PPP-XX-XXX → Altium folder; open Explorer / .PrjPcb / build zips.",
  },
  {
    path: CONVERT_BOM_TO_CSV_SCRIPT,
    label: "Convert_BOM_to_CSV.ps1",
    description: "Convert an Altium BOM xlsx to CSV (lives in OneDrive Documents).",
  },
];

const ALT_PART_PATTERN = /^\d{3}-\d{2}-\d{3}(?:\s+.+)?$/;

/** Trim and collapse whitespace; preserves a single descriptive suffix after PPP-XX-XXX. */
function normalizeAltiumPartArg(raw: string): string {
  return raw.trim().replace(/\s+/g, " ");
}

function partDedupeKey(raw: string): string {
  const t = normalizeAltiumPartArg(raw);
  const m = t.match(/^(\d{3}-\d{2}-\d{3})(?:\s+(.*))?$/);
  if (!m) {
    return t.toUpperCase();
  }
  const core = m[1].toUpperCase();
  const tail = (m[2] ?? "").trim().toLowerCase();
  return tail.length > 0 ? `${core}|${tail}` : core;
}

function pushRecentPart(entries: readonly string[], part: string, maxItems = 15): string[] {
  const canon = normalizeAltiumPartArg(part);
  if (!ALT_PART_PATTERN.test(canon)) {
    return [...entries];
  }
  const key = partDedupeKey(canon);
  const next = entries.filter((e) => partDedupeKey(e) !== key);
  return [canon, ...next].slice(0, maxItems);
}

function escapeSingleQuotes(s: string): string {
  return s.replace(/'/g, "''");
}

type AltiumOpenAction = "Explorer" | "PrjPcb" | "FabricationZip" | "PickPlaceZip";

function formatConvertBomToCsvCommand(): string {
  return [
    `powershell.exe -NoProfile -NoExit -ExecutionPolicy Bypass -File`,
    `"${CONVERT_BOM_TO_CSV_SCRIPT}"`,
  ].join(" ");
}

function formatCompareBomCommand(arenaCsv: string, altiumXlsx: string, sheet?: string): string {
  const parts = [
    `powershell.exe -NoProfile -NoExit -ExecutionPolicy Bypass -File`,
    `"${COMPARE_ARENA_ALTIUM_BOM_SCRIPT}"`,
    `-ArenaCsv '${escapeSingleQuotes(arenaCsv.trim())}'`,
    `-AltiumXlsx '${escapeSingleQuotes(altiumXlsx.trim())}'`,
  ];
  const trimmedSheet = (sheet ?? "").trim();
  if (trimmedSheet && trimmedSheet !== "2") {
    parts.push(`-Sheet '${escapeSingleQuotes(trimmedSheet)}'`);
  }
  return parts.join(" ");
}

function formatOpenAltiumCommand(partArgument: string, open: AltiumOpenAction = "Explorer"): string {
  const safe = escapeSingleQuotes(normalizeAltiumPartArg(partArgument));
  const parts = [
    `powershell -NoProfile -NoExit -ExecutionPolicy Bypass -File`,
    `"${OPEN_ALTIUM_FOLDER_SCRIPT}"`,
    `-PartNumber '${safe}'`,
  ];
  if (open !== "Explorer") {
    parts.push("-Open", open);
  }
  return parts.join(" ");
}

export default function PlmFileWorkbench() {
  const dispatch = useCanvasAction();
  const [pane, setPane] = useCanvasState<WorkbenchPane>("activePane", "overview");
  const [bomAgentId, setBomAgentId] = useCanvasState("bomComparisonAgentId", DEFAULT_BOM_AGENT_ID);
  const [arenaBomNote, setArenaBomNote] = useCanvasState("noteArenaBomPath", "");
  const [altiumBomNote, setAltiumBomNote] = useCanvasState("noteAltiumBomPath", "");
  const [transferNotes, setTransferNotes] = useCanvasState("transferScratchNotes", "");
  const [chkArenaExport, setChkArenaExport] = useCanvasState("chkArenaExport", false);
  const [chkAltiumBom, setChkAltiumBom] = useCanvasState("chkAltiumBom", false);
  const [chkSharePoint, setChkSharePoint] = useCanvasState("chkSharePoint", false);
  const [chkCompare, setChkCompare] = useCanvasState("chkCompare", false);
  const [altiumPartCommitted, setAltiumPartCommitted] = useCanvasState("altiumPartNumberEntry", "");
  const [altiumPartDraft, setAltiumPartDraft] = useState(altiumPartCommitted);

  useEffect(() => {
    setAltiumPartDraft(altiumPartCommitted);
  }, [altiumPartCommitted]);

  const [altiumRecentParts, setAltiumRecentParts] = useCanvasState<string[]>("altiumRecentPartNumbers", []);
  const [altiumExplorerNote, setAltiumExplorerNote] = useCanvasState("altiumExplorerLastNote", "");
  const [bomCsvNote, setBomCsvNote] = useCanvasState("bomCsvConversionLastNote", "");
  const [bomCompareSheet, setBomCompareSheet] = useCanvasState("bomCompareSheet", "2");
  const [bomCompareNote, setBomCompareNote] = useCanvasState("bomCompareLastNote", "");
  const [scriptCopyNote, setScriptCopyNote] = useCanvasState("scriptCopyLastNote", "");

  const copyToClipboard = async (text: string, successLabel: string) => {
    try {
      if (typeof navigator !== "undefined" && navigator.clipboard && typeof navigator.clipboard.writeText === "function") {
        await navigator.clipboard.writeText(text);
        setScriptCopyNote(`${successLabel} copied — paste into Cursor (Ctrl+P), File Explorer, or Win+R.`);
      } else {
        setScriptCopyNote(text);
      }
    } catch {
      setScriptCopyNote(`Copy failed — use manually: ${text}`);
    }
  };

  const openBomAgent = () => {
    const id = bomAgentId.trim() || DEFAULT_BOM_AGENT_ID;
    dispatch({ type: "openAgent", agentId: id });
  };

  const commitAltiumDraft = () => {
    const normalized = normalizeAltiumPartArg(altiumPartDraft);
    if (!ALT_PART_PATTERN.test(normalized)) {
      setAltiumExplorerNote(
        "Use PPP-XX-XXX, optionally followed by space and folder text (e.g. 272-10-001 BJ 5P), then press Enter to apply.",
      );
      return;
    }
    setAltiumPartCommitted(normalized);
    setAltiumExplorerNote("Part number applied — previews and launcher commands use this value.");
  };

  const handleAltiumPartFieldKeyDown = (e: KeyboardEvent<HTMLDivElement>) => {
    if (e.key !== "Enter") {
      return;
    }
    e.preventDefault();
    commitAltiumDraft();
  };

  const copyAltiumLauncher = async (open: AltiumOpenAction) => {
    const pn = normalizeAltiumPartArg(altiumPartCommitted);
    if (!ALT_PART_PATTERN.test(pn)) {
      setAltiumExplorerNote(
        "Type PPP-XX-XXX (optional suffix), then press Enter to apply before copying. Or pick a recent pill.",
      );
      return;
    }
    setAltiumRecentParts((prev) => pushRecentPart(prev ?? [], pn));
    setAltiumPartCommitted(pn);
    const cmd = formatOpenAltiumCommand(pn, open);
    const successMsg =
      open === "Explorer"
        ? "Copied the PowerShell command. Paste into Cursor's terminal (or Win+R Run) and press Enter — Explorer opens the latest Rev / Ver folder when present."
        : open === "PrjPcb"
          ? "Copied the PowerShell command. Paste into Cursor's terminal (or Win+R Run) and press Enter — the newest .PrjPcb under that folder (recursive, after latest Rev / Ver when present) opens in Altium."
          : open === "FabricationZip"
            ? "Copied the PowerShell command. Paste into a terminal and press Enter — Gerber / NC Drill / PFS are zipped into the production document folder as Fabrication PPP-YY-XXX Rev ?.zip (missing folders print a warning; orange warnings flag duplicate Gerber file types)."
            : "Copied the PowerShell command. Paste into a terminal and press Enter — the Pick Place folder is zipped into the production document folder as Pick Place PPP-TT-XXX Rev ?.zip. Errors if the folder has more than one .csv or .txt.";
    try {
      if (typeof navigator !== "undefined" && navigator.clipboard && typeof navigator.clipboard.writeText === "function") {
        await navigator.clipboard.writeText(cmd);
        setAltiumExplorerNote(successMsg);
      } else {
        setAltiumExplorerNote(cmd);
      }
    } catch {
      setAltiumExplorerNote(`Copy failed — run manually: ${cmd}`);
    }
  };

  const selectRecentAltiumPart = (part: string) => {
    setAltiumPartCommitted(part);
    setAltiumExplorerNote("");
  };

  const copyCompareBomLauncher = async () => {
    const arenaPath = arenaBomNote.trim();
    const altiumPath = altiumBomNote.trim();
    if (!arenaPath || !altiumPath) {
      setBomCompareNote(
        "Fill in both Arena BOM (CSV) and Altium BOM (xlsx) paths above before copying — the launcher needs both linked files.",
      );
      return;
    }
    const cmd = formatCompareBomCommand(arenaPath, altiumPath, bomCompareSheet);
    const successMsg =
      "Copied. Paste into a terminal and press Enter — the script reads the named/indexed Altium sheet (default 2) and prints only-in-Arena, only-in-Altium, qty / designator diffs on common parts, plus voltage-rating mismatches in descriptions.";
    try {
      if (typeof navigator !== "undefined" && navigator.clipboard && typeof navigator.clipboard.writeText === "function") {
        await navigator.clipboard.writeText(cmd);
        setBomCompareNote(successMsg);
      } else {
        setBomCompareNote(cmd);
      }
    } catch {
      setBomCompareNote(`Copy failed — run manually: ${cmd}`);
    }
  };

  const copyConvertBomLauncher = async () => {
    const cmd = formatConvertBomToCsvCommand();
    const successMsg =
      "Copied. Paste into a terminal and press Enter — choose the Altium BOM .xlsx; CSV is saved beside it after cleanup. Blank cells in column C on sheet 1 abort with an error.";
    try {
      if (typeof navigator !== "undefined" && navigator.clipboard && typeof navigator.clipboard.writeText === "function") {
        await navigator.clipboard.writeText(cmd);
        setBomCsvNote(successMsg);
      } else {
        setBomCsvNote(cmd);
      }
    } catch {
      setBomCsvNote(`Copy failed — run manually: ${cmd}`);
    }
  };

  return (
    <Stack gap={20}>
      <Stack gap={6}>
        <H1>PLM file workbench</H1>
        <Text tone="secondary" size="small">
          Hub for file hand-offs between Altium Designer, SharePoint, and Arena PLM. This canvas stores your notes and checklist
          state beside the chat; it does not call network APIs.
        </Text>
      </Stack>

      <Row gap={8} wrap style={{ rowGap: 8 }}>
        {PANES.map((p) => (
          <Pill
            key={p.id}
            active={pane === p.id}
            onClick={() => setPane(p.id)}
            title={p.label}
          >
            {p.label}
          </Pill>
        ))}
      </Row>

      <Divider />

      {pane === "overview" && (
        <Stack gap={16}>
          <Grid columns={3} gap={16}>
            <Stat value="Altium" label="Design + outputs" />
            <Stat value="SharePoint" label="Libraries + documents" />
            <Stat value="Arena" label="PLM BOM + lifecycle" />
          </Grid>
          <Text>
            Use the tabs above for system-specific checklists. On the Altium tab you can resolve a part folder and jump to Explorer, open the
            newest <Code>.PrjPcb</Code>, or build fabrication / pick &amp; place zips via the canvas launchers, or workspace chords —{" "}
            <Text weight="semibold">Ctrl+Shift+Alt+O</Text> /{" "}
            <Text weight="semibold">Ctrl+Shift+Alt+P</Text> /{" "}
            <Text weight="semibold">Ctrl+Shift+Alt+F</Text> /{" "}
            <Text weight="semibold">Ctrl+Shift+Alt+K</Text> /{" "}
            <Text weight="semibold">Ctrl+Shift+Alt+B</Text> (Explorer / PCB project / Fabrication / Pick &amp; place / BOM→CSV). For structured
            BOM reconciliation between Arena exports and Altium BOM Excel, use <Text weight="semibold">Arena vs Altium BOM</Text>.
          </Text>
          <Card>
            <CardHeader>Quick path</CardHeader>
            <CardBody>
              <Stack gap={10}>
                <Row gap={12} align="center" wrap>
                  <Checkbox checked={chkArenaExport} onChange={setChkArenaExport} label="Arena BOM exported for this rev" />
                  <Checkbox checked={chkAltiumBom} onChange={setChkAltiumBom} label="Altium BOM generated (xlsx)" />
                  <Checkbox checked={chkSharePoint} onChange={setChkSharePoint} label="Artifacts in SharePoint if required" />
                  <Checkbox checked={chkCompare} onChange={setChkCompare} label="BOM comparison reviewed" />
                </Row>
                <Text tone="secondary" size="small">
                  Checkboxes persist in this canvas so you can track a release without leaving the IDE.
                </Text>
              </Stack>
            </CardBody>
          </Card>

          <Card>
            <CardHeader>Scripts</CardHeader>
            <CardBody>
              <Stack gap={12}>
                <Row gap={8} align="center" wrap>
                  <Text size="small">Folder:</Text>
                  <Code style={{ wordBreak: "break-word" }}>{SCRIPTS_FOLDER}</Code>
                  <Button
                    variant="ghost"
                    onClick={() => copyToClipboard(SCRIPTS_FOLDER, "Scripts folder path")}
                  >
                    Copy folder
                  </Button>
                  <Button
                    variant="ghost"
                    onClick={() => copyToClipboard(`explorer "${SCRIPTS_FOLDER}"`, "Open-folder command")}
                  >
                    Copy "open folder" command
                  </Button>
                </Row>
                <Stack gap={10}>
                  {SCRIPT_LINKS.map((s) => (
                    <Stack key={s.path} gap={4}>
                      <Row gap={8} align="center" wrap>
                        <Text weight="semibold" size="small">
                          {s.label}
                        </Text>
                        <Button variant="ghost" onClick={() => copyToClipboard(s.path, s.label)}>
                          Copy path
                        </Button>
                      </Row>
                      <Text tone="secondary" size="small">
                        {s.description}
                      </Text>
                      <Code style={{ display: "block", whiteSpace: "pre-wrap", wordBreak: "break-word" }}>
                        {s.path}
                      </Code>
                    </Stack>
                  ))}
                </Stack>
                {scriptCopyNote ? (
                  <Callout
                    tone={scriptCopyNote.includes("failed") ? "danger" : scriptCopyNote.includes("copied") ? "success" : "neutral"}
                    title="Clipboard"
                  >
                    {scriptCopyNote}
                  </Callout>
                ) : null}
                <Text tone="secondary" size="small">
                  Paste a script path into Cursor's Quick Open (Ctrl+P) to view or edit, or the open-folder command into a terminal / Win+R to launch Explorer.
                </Text>
              </Stack>
            </CardBody>
          </Card>
        </Stack>
      )}

      {pane === "altium" && (
        <Stack gap={14}>
          <H2>Altium Designer</H2>

          <Card>
            <CardHeader>Open Altium design folder or .PrjPcb</CardHeader>
            <CardBody>
              <Stack gap={12}>
                <Text tone="secondary" size="small">
                  Enter <Text weight="semibold">PPP-XX-XXX</Text>, optionally followed by text that matches the folder name (for example{" "}
                  <Code>272-10-001 BJ 5P</Code>), then press <Text weight="semibold">Enter</Text> so the canvas applies it &mdash; previews
                  and copy buttons only use the applied value (typing alone does not persist). The helper looks under{" "}
                  <Code>{ALTIUM_LIBRARY_ROOT}</Code> &mdash; same resolution rules as before (Rev / Ver folders, nested layouts, suffix
                  matching). Opening <Code>.PrjPcb</Code> still searches recursively for the newest project file after that folder resolves.
                </Text>
                <Stack gap={6}>
                  <Text weight="semibold" size="small">
                    Part number
                  </Text>
                  <div onKeyDown={handleAltiumPartFieldKeyDown}>
                    <TextInput
                      value={altiumPartDraft}
                      onChange={setAltiumPartDraft}
                      placeholder="e.g. 272-10-001 — press Enter to apply"
                      style={{ minWidth: 200, flex: "1 1 200px" }}
                    />
                  </div>
                  <Row gap={10} align="center" wrap style={{ rowGap: 10 }}>
                    <Button variant="primary" onClick={() => copyAltiumLauncher("Explorer")}>
                      Copy launcher — open in Explorer
                    </Button>
                    <Button variant="secondary" onClick={() => copyAltiumLauncher("PrjPcb")}>
                      Copy launcher — open .PrjPcb in Altium
                    </Button>
                    <Button variant="primary" onClick={() => copyAltiumLauncher("FabricationZip")}>
                      Copy launcher — Create Fabrication zip
                    </Button>
                    <Button variant="primary" onClick={() => copyAltiumLauncher("PickPlaceZip")}>
                      Copy launcher — Create Pick &amp; Place zip
                    </Button>
                  </Row>
                </Stack>
                <Text tone="secondary" size="small">
                  This canvas cannot run shell commands by itself. Each button copies a one-line PowerShell command. Explorer launches{" "}
                  <Code>explorer.exe</Code>; PCB uses <Code>-Open PrjPcb</Code> to find <Code>*.PrjPcb</Code> recursively and open the newest;
                  fabrication zip (<Code>-Open FabricationZip</Code>) bundles <Code>Gerber</Code>, <Code>NC Drill</Code>, and{" "}
                  <Code>PFS</Code> into <Code>Fabrication PPP-YY-XXX Rev ?.zip</Code> inside the production document folder
                  (TT&nbsp;10→YY&nbsp;20, TT&nbsp;15→YY&nbsp;25), with an orange warning per duplicate Gerber file type and red errors that
                  abort the zip if <Code>PFS</Code> doesn’t hold exactly one <Code>.pdf</Code>, <Code>Gerber</Code> has fewer than 8 files, or{" "}
                  <Code>NC Drill</Code> has fewer than 3 files; pick &amp; place zip (<Code>-Open PickPlaceZip</Code>) zips the{" "}
                  <Code>Pick Place</Code> folder into <Code>Pick Place PPP-TT-XXX Rev ?.zip</Code> in the same location, and errors if there is more
                  than one <Code>.csv</Code> or <Code>.txt</Code> file in it. Revision letter is read from the folder name or any{" "}
                  <Code>.PrjPcb</Code> file under it.
                </Text>
                <Callout tone="info" title="Open without pasting">
                  Workspace tasks <Code>Open Altium part folder (Explorer)</Code>, <Code>Open Altium PCB project (.PrjPcb)</Code>,{" "}
                  <Code>Build Altium fabrication zip</Code>, and <Code>Build Altium pick &amp; place zip</Code> share the same part-number prompt.
                  Run <Code>Tasks: Run Task</Code>, pick the task name, or use{" "}
                  <Text weight="semibold">Ctrl+Shift+Alt+O</Text> /{" "}
                  <Text weight="semibold">Ctrl+Shift+Alt+P</Text> /{" "}
                  <Text weight="semibold">Ctrl+Shift+Alt+F</Text> /{" "}
                  <Text weight="semibold">Ctrl+Shift+Alt+K</Text> /{" "}
                  <Text weight="semibold">Ctrl+Shift+Alt+B</Text> (BOM-to-CSV) — requires this workspace open.
                </Callout>
                {altiumExplorerNote ? (
                  <Callout
                    tone={
                      altiumExplorerNote.includes("Expected") ||
                      altiumExplorerNote.includes("Use format") ||
                      altiumExplorerNote.includes("Use PPP") ||
                      altiumExplorerNote.includes("Enter to apply") ||
                      altiumExplorerNote.includes("Enter to apply before")
                        ? "warning"
                        : altiumExplorerNote.includes("failed")
                          ? "danger"
                          : "success"
                    }
                    title="Launcher"
                  >
                    {altiumExplorerNote}
                  </Callout>
                ) : null}
                <Stack gap={6}>
                  <Text weight="semibold" size="small">
                    Preview commands
                  </Text>
                  <Text size="small" tone="secondary">
                    Explorer (default script)
                  </Text>
                  <Code style={{ display: "block", whiteSpace: "pre-wrap", wordBreak: "break-word" }}>
                    {formatOpenAltiumCommand(
                      ALT_PART_PATTERN.test(normalizeAltiumPartArg(altiumPartCommitted))
                        ? normalizeAltiumPartArg(altiumPartCommitted)
                        : "296-10-001",
                      "Explorer",
                    )}
                  </Code>
                  <Text size="small" tone="secondary">
                    PCB project (.PrjPcb, newest LastWriteTime in resolved folder subtree)
                  </Text>
                  <Code style={{ display: "block", whiteSpace: "pre-wrap", wordBreak: "break-word" }}>
                    {formatOpenAltiumCommand(
                      ALT_PART_PATTERN.test(normalizeAltiumPartArg(altiumPartCommitted))
                        ? normalizeAltiumPartArg(altiumPartCommitted)
                        : "296-10-001",
                      "PrjPcb",
                    )}
                  </Code>
                  <Text size="small" tone="secondary">
                    Fabrication zip (Gerber + NC Drill + PFS → Fabrication PPP-YY-XXX Rev ?.zip)
                  </Text>
                  <Code style={{ display: "block", whiteSpace: "pre-wrap", wordBreak: "break-word" }}>
                    {formatOpenAltiumCommand(
                      ALT_PART_PATTERN.test(normalizeAltiumPartArg(altiumPartCommitted))
                        ? normalizeAltiumPartArg(altiumPartCommitted)
                        : "296-10-001",
                      "FabricationZip",
                    )}
                  </Code>
                  <Text size="small" tone="secondary">
                    Pick &amp; place zip (Pick Place folder → Pick Place PPP-TT-XXX Rev ?.zip; errors on multiple .csv or .txt)
                  </Text>
                  <Code style={{ display: "block", whiteSpace: "pre-wrap", wordBreak: "break-word" }}>
                    {formatOpenAltiumCommand(
                      ALT_PART_PATTERN.test(normalizeAltiumPartArg(altiumPartCommitted))
                        ? normalizeAltiumPartArg(altiumPartCommitted)
                        : "296-10-001",
                      "PickPlaceZip",
                    )}
                  </Code>
                </Stack>
                {altiumRecentParts.length > 0 ? (
                  <Stack gap={8}>
                    <Text weight="semibold" size="small">
                      Recent part numbers
                    </Text>
                    <Row gap={8} wrap style={{ rowGap: 8 }}>
                      {altiumRecentParts.map((p) => (
                        <Pill key={p} onClick={() => selectRecentAltiumPart(p)} title={`Fill ${p}`}>
                          {p}
                        </Pill>
                      ))}
                    </Row>
                    <Text tone="secondary" size="small">
                      Click a pill to apply immediately, then copy. Or edit the field and press Enter before copying.
                    </Text>
                  </Stack>
                ) : null}
              </Stack>
            </CardBody>
          </Card>

          <Card>
            <CardHeader>Convert BOM (Excel → CSV)</CardHeader>
            <CardBody>
              <Stack gap={10}>
                <Text tone="secondary" size="small">
                  Launcher for <Code>Convert_BOM_to_CSV.ps1</Code>: picks an Altium BOM spreadsheet, saves a CSV beside it, applies your
                  line trims and checks. Sheet 1 must have every used-range row filled in{" "}
                  <Text weight="semibold">column C</Text> — empty cells terminate with exit code 1 before export.
                </Text>
                <Row gap={8} wrap style={{ rowGap: 8 }}>
                  <Button variant="secondary" onClick={() => copyConvertBomLauncher()}>
                    Copy launcher — BOM to CSV
                  </Button>
                </Row>
                <Callout tone="info" title="Workspace task">
                  Run <Code>Tasks: Run Task</Code>, then choose <Text weight="semibold">Convert BOM Excel to CSV</Text> — same script,
                  no clipboard step (Excel GUI file dialog opens).
                </Callout>
                <Code style={{ display: "block", whiteSpace: "pre-wrap", wordBreak: "break-word" }}>
                  {formatConvertBomToCsvCommand()}
                </Code>
                {bomCsvNote ? (
                  <Callout
                    tone={
                      bomCsvNote.includes("failed") ? "danger" : bomCsvNote.includes("Copied") ? "success" : "neutral"
                    }
                    title="BOM conversion"
                  >
                    {bomCsvNote}
                  </Callout>
                ) : null}
              </Stack>
            </CardBody>
          </Card>
        </Stack>
      )}

      {pane === "sharepoint" && (
        <Stack gap={14}>
          <H2>SharePoint</H2>
          <Text>
            Use SharePoint for document libraries, released packages, and team-visible file drops. This canvas does not connect to
            Microsoft Graph; track locations in the scratch pad below or in your agent prompt when moving files.
          </Text>
          <H3>Scratch — paths and library links</H3>
          <TextArea
            value={transferNotes}
            onChange={setTransferNotes}
            placeholder="e.g. site / library / folder paths, naming convention for rev folders…"
            rows={5}
          />
        </Stack>
      )}

      {pane === "arena" && (
        <Stack gap={14}>
          <H2>Arena PLM</H2>
          <Text>
            Export the assembly BOM from Arena as CSV for offline comparison. Column names vary by export profile; the comparison
            agent needs the same field mapping you care about (internal part number, quantity, reference designators, description).
          </Text>
          <H3>Reminder — Arena BOM file</H3>
          <TextInput
            value={arenaBomNote}
            onChange={setArenaBomNote}
            placeholder="Optional: paste or describe latest Arena CSV path for this program"
          />
        </Stack>
      )}

      {pane === "bomComparison" && (
        <Stack gap={16}>
          <H2>Arena vs Altium BOM comparison</H2>
          <Text>
            This implements the <Text weight="semibold">Comparison of Arena and Altium BOMs</Text> workflow as a first-class pane:
            prepare both files, then open the agent conversation that performs the CSV vs Excel engineering-BOM diff.
          </Text>

          <Card>
            <CardHeader>Open agent</CardHeader>
            <CardBody>
              <Stack gap={12}>
                <Text tone="secondary" size="small">
                  Opens the Cursor agent transcript by UUID. Defaults to your existing BOM comparison chat; change only if you
                  duplicate or replace that agent.
                </Text>
                <H3 style={{ margin: 0 }}>Agent transcript ID</H3>
                <TextInput
                  value={bomAgentId}
                  onChange={setBomAgentId}
                  placeholder={DEFAULT_BOM_AGENT_ID}
                />
                <Row gap={8} wrap style={{ rowGap: 8 }}>
                  <Button variant="primary" onClick={openBomAgent}>
                    Open BOM comparison agent
                  </Button>
                  <Button
                    variant="ghost"
                    onClick={() => setBomAgentId(DEFAULT_BOM_AGENT_ID)}
                  >
                    Reset to default ID
                  </Button>
                </Row>
              </Stack>
            </CardBody>
          </Card>

          <H3>File hints (persisted)</H3>
          <Grid columns="1fr 1fr" gap={12}>
            <Stack gap={6}>
              <Text weight="semibold" size="small">
                Arena BOM (CSV)
              </Text>
              <TextInput
                value={arenaBomNote}
                onChange={setArenaBomNote}
                placeholder="Path or filename for Arena export"
              />
            </Stack>
            <Stack gap={6}>
              <Text weight="semibold" size="small">
                Altium BOM (xlsx)
              </Text>
              <TextInput
                value={altiumBomNote}
                onChange={setAltiumBomNote}
                placeholder="Path to Project Outputs BOM xlsx"
              />
            </Stack>
          </Grid>

          <Callout tone="warning" title="Sheet selection">
            Altium BOM workbooks often have multiple sheets. Confirm whether your source of truth matches the Arena export (for
            example an engineering BOM sheet versus a flattened procurement view).
          </Callout>

          <Card>
            <CardHeader>Compare linked BOMs</CardHeader>
            <CardBody>
              <Stack gap={12}>
                <Text tone="secondary" size="small">
                  Runs <Code>Compare-ArenaAltiumBom.ps1</Code> against the two paths above. Matches line items by internal part
                  number and reports only-in-Arena, only-in-Altium, Altium rows without a part number (mounting holes, test
                  points), and quantity / designator differences across common parts. Also flags voltage-rating mismatches
                  inside descriptions when the part numbers agree.
                </Text>
                <Stack gap={6}>
                  <Text weight="semibold" size="small">
                    Altium sheet (name or 1-based index)
                  </Text>
                  <TextInput
                    value={bomCompareSheet}
                    onChange={setBomCompareSheet}
                    placeholder="Default: 2 (engineering BOM)"
                    style={{ minWidth: 200, flex: "0 1 240px" }}
                  />
                </Stack>
                <Row gap={8} wrap style={{ rowGap: 8 }}>
                  <Button variant="primary" onClick={copyCompareBomLauncher}>
                    Copy launcher — Compare Arena vs Altium BOM
                  </Button>
                </Row>
                {bomCompareNote ? (
                  <Callout
                    tone={
                      bomCompareNote.includes("failed")
                        ? "danger"
                        : bomCompareNote.includes("Fill in both")
                          ? "warning"
                          : bomCompareNote.includes("Copied")
                            ? "success"
                            : "neutral"
                    }
                    title="Compare BOMs"
                  >
                    {bomCompareNote}
                  </Callout>
                ) : null}
                <Stack gap={6}>
                  <Text weight="semibold" size="small">
                    Preview command
                  </Text>
                  <Code style={{ display: "block", whiteSpace: "pre-wrap", wordBreak: "break-word" }}>
                    {formatCompareBomCommand(
                      arenaBomNote.trim() || "C:\\path\\to\\arena_export.csv",
                      altiumBomNote.trim() || "C:\\path\\to\\altium_bom.xlsx",
                      bomCompareSheet,
                    )}
                  </Code>
                </Stack>
              </Stack>
            </CardBody>
          </Card>

          <H3>Workflow</H3>
          <Table headers={["Source", "Action"]} rows={BOM_WORKFLOW_ROWS} />

          <H3>What the agent compares</H3>
          <Text tone="secondary" size="small">
            Example from a prior run: match line items by part number or designator scope, list only-in-Arena / only-in-Altium
            parts, and call out description or internal-PN mismatches on otherwise aligned rows.
          </Text>
        </Stack>
      )}
    </Stack>
  );
}
