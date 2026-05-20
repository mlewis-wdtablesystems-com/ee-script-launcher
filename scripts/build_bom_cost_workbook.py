"""Build BOM cost workbook: columns M,N,I,Q,R,S with S1 = SUM(S2:S*)."""
from pathlib import Path

from openpyxl import Workbook
from openpyxl.utils import get_column_letter

# 1-based Excel column indices
COL_M = 13
COL_N = 14
COL_I = 9
COL_Q = 17
COL_R = 18
COL_S = 19
COL_NOTE = 20

OUT = Path(__file__).resolve().parent / "263-70-001_BOM_cost_export.xlsx"

# Lines identifiable from BOM image; Q not visible — default Q = I for tier lookup.
LINES = [
    {
        "vendor": "Wagner Electronics",
        "vpn": "SMP12V12A-25P",
        "qty_i": 1,
        "qty_q": 1,
        "unit_r": 50.7,
        "note": "12V 12A SMP; AUD ex-GST proxy from OHM Electronics May-2026 (verify on Wagner).",
    },
    {
        "vendor": "RS",
        "vpn": "224-0471",
        "qty_i": 28,
        "qty_q": 28,
        "unit_r": 4.478,
        "note": "RS AU 224-0471, 10–40 ea tier, AUD ex-GST per each (pack of 10).",
    },
    {
        "vendor": "Digikey",
        "vpn": "118-SRP3212-2R2MCT-ND",
        "qty_i": 1,
        "qty_q": 1,
        "unit_r": None,
        "note": "DigiKey unit price not retrieved here; enter AUD in R.",
    },
    {
        "vendor": "Digikey",
        "vpn": "WM19641-ND",
        "qty_i": 1,
        "qty_q": 1,
        "unit_r": None,
        "note": "DigiKey unit price not retrieved here; enter AUD in R.",
    },
]


def main() -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "BOM Cost"

    # Row 1: cumulative total in S1 (per request); labels in M/N area
    ws.cell(1, COL_M, "BOM line-total sum (S1)")
    ws.cell(1, COL_N, "Add all rows below; partial sample only.")

    # Row 2: column headers
    ws.cell(2, COL_M, "Vendor (M)")
    ws.cell(2, COL_N, "Vendor item (N)")
    ws.cell(2, COL_I, "Qty (I)")
    ws.cell(2, COL_Q, "Price lookup qty (Q)")
    ws.cell(2, COL_R, "Unit cost (R)")
    ws.cell(2, COL_S, "Line total (S)")
    ws.cell(2, COL_NOTE, "Notes")

    start = 3
    for i, row in enumerate(LINES, start=start):
        ws.cell(i, COL_M, row["vendor"])
        ws.cell(i, COL_N, row["vpn"])
        ws.cell(i, COL_I, row["qty_i"])
        ws.cell(i, COL_Q, row["qty_q"])
        if row["unit_r"] is not None:
            ws.cell(i, COL_R, row["unit_r"])
        ws.cell(i, COL_S, f"={get_column_letter(COL_I)}{i}*{get_column_letter(COL_R)}{i}")
        ws.cell(i, COL_NOTE, row["note"])

    last = start + len(LINES) - 1
    ws.cell(1, COL_S, f"=SUM({get_column_letter(COL_S)}{start}:{get_column_letter(COL_S)}{last})")

    ws.freeze_panes = "A3"
    wb.save(OUT)
    print(f"Wrote {OUT}")


if __name__ == "__main__":
    main()
