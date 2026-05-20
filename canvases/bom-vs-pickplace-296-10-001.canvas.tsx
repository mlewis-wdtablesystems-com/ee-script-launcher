import {
  Callout,
  Card,
  CardBody,
  CardHeader,
  Checkbox,
  Divider,
  Grid,
  H1,
  H2,
  Pill,
  Row,
  Stack,
  Stat,
  Table,
  Text,
  useCanvasState,
  useHostTheme,
} from "cursor/canvas";

const PROJECT = "296-10-001 6Pos 2MB 9SB Rev A";
const PP_FILE =
  "Pick Place for 296-10-001 6Pos 2MB 9SB Rev A.csv";
const BOM_FILE =
  "Bill of Materials-296-10-001 6Pos 2MB 9SB Rev A.csv";
const PP_DATE = "08/04/26 10:52";
const VARIANT =
  "Variant of 296-10-001 6Pos 2MB 9SB Rev A-Big_and_Small";

type Reason = "fiducial" | "dnp" | "bare-pcb" | "missing";

interface OnlyInPP {
  designator: string;
  comment: string;
  layer: string;
  footprint: string;
  description: string;
  reason: Reason;
}

interface OnlyInBOM {
  designator: string;
  item: string;
  reason: Reason;
}

const ONLY_IN_PP: OnlyInPP[] = [
  { designator: "M2", comment: "Fiducial", layer: "TopLayer", footprint: "FIDUCIAL_MARK_1MM", description: "Fiducial", reason: "fiducial" },
  { designator: "M3", comment: "Fiducial", layer: "BottomLayer", footprint: "FIDUCIAL_MARK_1MM", description: "Fiducial", reason: "fiducial" },
  { designator: "M4", comment: "Fiducial", layer: "BottomLayer", footprint: "FIDUCIAL_MARK_1MM", description: "Fiducial", reason: "fiducial" },
  { designator: "M5", comment: "Fiducial", layer: "BottomLayer", footprint: "FIDUCIAL_MARK_1MM", description: "Fiducial", reason: "fiducial" },
  { designator: "M6", comment: "Fiducial", layer: "TopLayer", footprint: "FIDUCIAL_MARK_1MM", description: "Fiducial", reason: "fiducial" },
  { designator: "M7", comment: "Fiducial", layer: "TopLayer", footprint: "FIDUCIAL_MARK_1MM", description: "Fiducial", reason: "fiducial" },
  { designator: "VC1_A0", comment: "DNP", layer: "BottomLayer", footprint: "TRIM_CAP1", description: "CAP TRIMMER 5.5-65PF 150V TH", reason: "dnp" },
  { designator: "VC1_A1", comment: "DNP", layer: "BottomLayer", footprint: "TRIM_CAP1", description: "CAP TRIMMER 5.5-65PF 150V TH", reason: "dnp" },
  { designator: "VC1_A2", comment: "DNP", layer: "BottomLayer", footprint: "TRIM_CAP1", description: "CAP TRIMMER 5.5-65PF 150V TH", reason: "dnp" },
  { designator: "VC1_A3", comment: "DNP", layer: "BottomLayer", footprint: "TRIM_CAP1", description: "CAP TRIMMER 5.5-65PF 150V TH", reason: "dnp" },
  { designator: "VC1_A4", comment: "DNP", layer: "BottomLayer", footprint: "TRIM_CAP1", description: "CAP TRIMMER 5.5-65PF 150V TH", reason: "dnp" },
  { designator: "VC1_A5", comment: "DNP", layer: "BottomLayer", footprint: "TRIM_CAP1", description: "CAP TRIMMER 5.5-65PF 150V TH", reason: "dnp" },
  { designator: "VC1_A6", comment: "DNP", layer: "BottomLayer", footprint: "TRIM_CAP1", description: "CAP TRIMMER 5.5-65PF 150V TH", reason: "dnp" },
  { designator: "VC1_A7", comment: "DNP", layer: "BottomLayer", footprint: "TRIM_CAP1", description: "CAP TRIMMER 5.5-65PF 150V TH", reason: "dnp" },
  { designator: "VC1_A8", comment: "DNP", layer: "BottomLayer", footprint: "TRIM_CAP1", description: "CAP TRIMMER 5.5-65PF 150V TH", reason: "dnp" },
  { designator: "VC1_A9", comment: "DNP", layer: "BottomLayer", footprint: "TRIM_CAP1", description: "CAP TRIMMER 5.5-65PF 150V TH", reason: "dnp" },
  { designator: "VC1_A10", comment: "DNP", layer: "BottomLayer", footprint: "TRIM_CAP1", description: "CAP TRIMMER 5.5-65PF 150V TH", reason: "dnp" },
];

const ONLY_IN_BOM: OnlyInBOM[] = [
  { designator: "PCB", item: "296-20-001", reason: "bare-pcb" },
];

const COUNTS = {
  pickPlace: 150,
  bom: 134,
  bomLessPcb: 133,
  matched: 133,
  onlyPP: 17,
  onlyBOM: 1,
  fiducials: 6,
  dnp: 11,
  barePcb: 1,
};

interface ReasonStyle {
  label: string;
  tone: "neutral" | "warning" | "info" | "success";
  short: string;
}

const REASON_STYLE: Record<Reason, ReasonStyle> = {
  fiducial: { label: "Fiducial — not a part", tone: "neutral", short: "Fiducial" },
  dnp: { label: "DNP — Do Not Place", tone: "warning", short: "DNP" },
  "bare-pcb": { label: "Bare PCB itself", tone: "neutral", short: "Bare PCB" },
  missing: { label: "Missing on the other side", tone: "warning", short: "Missing" },
};

export default function BomVsPickPlace() {
  const theme = useHostTheme();
  const [hideFiducials, setHideFiducials] = useCanvasState("hideFiducials", false);
  const [hideDnp, setHideDnp] = useCanvasState("hideDnp", false);
  const [hideBarePcb, setHideBarePcb] = useCanvasState("hideBarePcb", false);

  const filteredPP = ONLY_IN_PP.filter((row) => {
    if (hideFiducials && row.reason === "fiducial") return false;
    if (hideDnp && row.reason === "dnp") return false;
    return true;
  });

  const filteredBOM = ONLY_IN_BOM.filter((row) => {
    if (hideBarePcb && row.reason === "bare-pcb") return false;
    return true;
  });

  const ppHeaders = ["Designator", "Comment", "Footprint", "Layer", "Reason"];
  const ppRows = filteredPP.map((row) => [
    <Text weight="medium">{row.designator}</Text>,
    <Text>{row.comment}</Text>,
    <Text tone="secondary">{row.footprint}</Text>,
    <Text tone="secondary" size="small">{row.layer}</Text>,
    <Pill size="sm" tone={REASON_STYLE[row.reason].tone}>{REASON_STYLE[row.reason].short}</Pill>,
  ]);
  const ppTone = filteredPP.map((row) =>
    row.reason === "dnp" ? ("warning" as const) : undefined,
  );

  const bomHeaders = ["Designator", "Item Number", "Reason"];
  const bomRows = filteredBOM.map((row) => [
    <Text weight="medium">{row.designator}</Text>,
    <Text>{row.item}</Text>,
    <Pill size="sm" tone={REASON_STYLE[row.reason].tone}>{REASON_STYLE[row.reason].short}</Pill>,
  ]);

  const consistencyRows: Array<[string, "ok" | "warn", string]> = [
    [
      "BOM row quantities match designator counts",
      "ok",
      "Every Quantity column matches the count of comma-separated designators in that row.",
    ],
    [
      "BOM groups have consistent component values",
      "ok",
      "Every BOM line item shares one Comment value across all its designators in the Pick & Place file.",
    ],
    [
      "BOM groups have consistent footprints",
      "ok",
      "Every BOM line item shares one Footprint across all its designators in the Pick & Place file.",
    ],
  ];

  return (
    <Stack gap={20} style={{ padding: 20, maxWidth: 1100 }}>
      <Stack gap={4}>
        <H1>BOM vs Pick &amp; Place</H1>
        <Text tone="secondary">
          Project <Text weight="semibold">{PROJECT}</Text>{" "}
          <Text tone="tertiary">· P&amp;P generated {PP_DATE}</Text>
        </Text>
        <Text tone="tertiary" size="small" italic>
          {VARIANT}
        </Text>
      </Stack>

      <Grid columns={4} gap={12}>
        <Stat value={COUNTS.pickPlace} label="Pick & Place designators" />
        <Stat value={COUNTS.bom} label="BOM designators" />
        <Stat value={COUNTS.matched} label="Matched on both sides" tone="success" />
        <Stat
          value={COUNTS.onlyPP + COUNTS.onlyBOM}
          label="Different (one side only)"
          tone="warning"
        />
      </Grid>

      <Callout tone="success" title="No unexpected mismatches">
        All {COUNTS.onlyPP + COUNTS.onlyBOM} differences are accounted for: {COUNTS.fiducials} fiducial marks
        (P&amp;P only — fiducials aren't BOM line items), {COUNTS.dnp} VC1 trim caps marked
        <Text weight="semibold"> DNP</Text> in the P&amp;P (intentionally omitted from the BOM), and {COUNTS.barePcb} bare
        PCB row (BOM only — the board itself isn't placed). Every other designator
        appears in both files.
      </Callout>

      <Divider />

      <Stack gap={12}>
        <Row align="center" justify="space-between" wrap>
          <Stack gap={2}>
            <H2>Designators only in Pick &amp; Place</H2>
            <Text tone="secondary" size="small">
              Showing {filteredPP.length} of {ONLY_IN_PP.length}
            </Text>
          </Stack>
          <Row gap={16} align="center" wrap>
            <Checkbox
              checked={hideFiducials}
              onChange={setHideFiducials}
              label={`Hide fiducials (${COUNTS.fiducials})`}
            />
            <Checkbox
              checked={hideDnp}
              onChange={setHideDnp}
              label={`Hide DNP parts (${COUNTS.dnp})`}
            />
          </Row>
        </Row>
        <Table
          headers={ppHeaders}
          rows={ppRows}
          rowTone={ppTone}
          columnAlign={["left", "left", "left", "left", "right"]}
          striped
          emptyMessage={
            <Text tone="secondary">All hidden by current filters.</Text>
          }
        />
      </Stack>

      <Stack gap={12}>
        <Row align="center" justify="space-between" wrap>
          <Stack gap={2}>
            <H2>Designators only in BOM</H2>
            <Text tone="secondary" size="small">
              Showing {filteredBOM.length} of {ONLY_IN_BOM.length}
            </Text>
          </Stack>
          <Checkbox
            checked={hideBarePcb}
            onChange={setHideBarePcb}
            label="Hide bare PCB row"
          />
        </Row>
        <Table
          headers={bomHeaders}
          rows={bomRows}
          columnAlign={["left", "left", "right"]}
          emptyMessage={
            <Text tone="secondary">All hidden by current filters.</Text>
          }
        />
      </Stack>

      <Divider />

      <H2>Internal BOM consistency</H2>
      <Card>
        <CardHeader>Cross-checks against the Pick &amp; Place file</CardHeader>
        <CardBody>
          <Stack gap={10}>
            {consistencyRows.map(([title, status, detail]) => (
              <Row key={title} gap={12} align="start">
                <Pill
                  size="sm"
                  tone={status === "ok" ? "success" : "warning"}
                  active
                  style={{ marginTop: 2, minWidth: 36, justifyContent: "center" }}
                >
                  {status === "ok" ? "OK" : "Check"}
                </Pill>
                <Stack gap={2} style={{ flex: 1, minWidth: 0 }}>
                  <Text weight="semibold">{title}</Text>
                  <Text tone="secondary" size="small">{detail}</Text>
                </Stack>
              </Row>
            ))}
          </Stack>
        </CardBody>
      </Card>

      <Divider />

      <H2>How designators were matched</H2>
      <Stack gap={6}>
        <Text>
          Each row of the BOM lists one or more comma-separated designators in the
          <Text weight="semibold"> Reference Designator</Text> column. Those were
          expanded into the BOM designator set and compared against the
          <Text weight="semibold"> Designator</Text> column of the Pick &amp; Place file.
          A designator counts as a match if the exact string (case-sensitive)
          appears in both sets.
        </Text>
        <Text tone="secondary" size="small">
          Sources: <Text as="span" italic>{PP_FILE}</Text> · <Text as="span" italic>{BOM_FILE}</Text>
        </Text>
        <Text
          tone="tertiary"
          size="small"
          style={{ borderLeft: `2px solid ${theme.stroke.tertiary}`, paddingLeft: 8 }}
        >
          To re-run this comparison on a different revision, ask the agent to
          regenerate this canvas with the new file paths.
        </Text>
      </Stack>
    </Stack>
  );
}
