"""Generate synthetic test data for Employee Score Card.

Designs the test scenarios from the implementation plan:
- Two xlsx files with overlapping wildcards (dedup test)
- Multi-week cumulative → incremental conversion (Bob on WO-400)
- Split operation flagged (WO-100/0010, Alice + Bob, 60/40 of charges)
- ACTIVE op with no Hours Earned row (WO-100/0020) — unmatched
- Whole new WO with no Hours Earned (WO-300) — surfaces in unmatched report at top
"""
from openpyxl import Workbook
from datetime import datetime
from pathlib import Path

OUT = Path(__file__).parent

HEADERS = ["Order No", "Week Ending", "Employee Name", "Badge", "Part No.",
           "Part Description", "Operation", "Total Hours", "Status"]

def make_xlsx(filename, rows, export_label):
    """Write an .xlsx with a 3-row export-metadata block, then headers on row 4, then data."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Charges"
    ws["A1"] = f"Charge Export — {export_label}"
    ws["A2"] = f"Generated: {datetime.now().isoformat()}"
    ws["A3"] = "Source: synthetic test data"
    for c, h in enumerate(HEADERS, start=1):
        ws.cell(row=4, column=c, value=h)
    for r, row in enumerate(rows, start=5):
        for c, val in enumerate(row, start=1):
            ws.cell(row=r, column=c, value=val)
    wb.save(OUT / filename)
    print(f"Wrote {filename} ({len(rows)} data rows)")

# All weeks formatted MM/DD/YYYY HH:MM:SS AM/PM as strings
def w(month, day, year=2026):
    return f"{month}/{day}/{year} 12:00:00 AM"

# ----------------------------------------------------------------------------
# Scenario 1: WO-100, Op 0010, "Wing Spar Machining", part 555-01
#   Alice charges cumulative: w1=10, w2=30 (incr 10, 20)
#   Bob charges cumulative:    w1=0,  w2=20 (incr 0, 20)  → only one row at w2
#   Closes Mar 2026, Hours Earned = 60. Expected: Alice 36 earned, Bob 24 earned. SPLIT.
# ----------------------------------------------------------------------------
# Scenario 2: WO-100, Op 0020 — ACTIVE (no Hours Earned row)
#   Alice charges 5h. Unmatched.
# Scenario 3: WO-200, Op 0010, "Fuselage Drill", part 777-12
#   Charlie charges 50h cumulative. Closes Apr 2026, Hours Earned = 40.
#   Expected: Charlie 40 earned, 50 actual, CPI 0.8.
# Scenario 4: WO-300, Op 0030, "Inspect", part 999-05
#   Alice 10h, Charlie 8h. ACTIVE, never closes. Unmatched (will surface in report).
# Scenario 5: WO-400, Op 0010, "Sheet Metal Cut", part 555-02
#   Bob multi-week:
#     w 3/7  cum 5h
#     w 3/14 cum 12h
#     w 3/21 cum 20h
#   Closes Mar 2026, Hours Earned 25. Bob earned = 25, actual 20, CPI 1.25.
# Scenario 6: WO-500, Op 0010, "Test", part 555-01
#   Bob solo 30h, closes Apr 2026, Hours Earned 25. CPI 0.83.

# File A: covers parts 555-* and 777-*
file_a_rows = [
    # WO-100/0010 split op
    ["WO-100", w(2, 28), "Alice Johnson", "B001", "555-01", "Wing Spar Machining", "0010", 10.0, "ACTIVE"],
    ["WO-100", w(3, 7),  "Alice Johnson", "B001", "555-01", "Wing Spar Machining", "0010", 30.0, "CLOSE"],
    ["WO-100", w(3, 7),  "Bob Lee",       "B002", "555-01", "Wing Spar Machining", "0010", 20.0, "CLOSE"],

    # WO-200/0010 Charlie alone
    ["WO-200", w(4, 4),  "Charlie Mendez","B003", "777-12", "Fuselage Drill",      "0010", 50.0, "CLOSE"],

    # WO-400/0010 Bob multi-week (cumulative)
    ["WO-400", w(3, 7),  "Bob Lee",       "B002", "555-02", "Sheet Metal Cut",     "0010", 5.0,  "ACTIVE"],
    ["WO-400", w(3, 14), "Bob Lee",       "B002", "555-02", "Sheet Metal Cut",     "0010", 12.0, "ACTIVE"],
    ["WO-400", w(3, 21), "Bob Lee",       "B002", "555-02", "Sheet Metal Cut",     "0010", 20.0, "CLOSE"],

    # WO-500/0010 Bob alone (slow tech: actual > earned)
    ["WO-500", w(4, 11), "Bob Lee",       "B002", "555-01", "Wing Spar Machining", "0010", 30.0, "CLOSE"],
]

# File B: covers parts 5*-* (overlaps 555-*) and 999-* (no earned coverage)
file_b_rows = [
    # WO-100/0010 — duplicates of File A's rows; dedup should keep them once
    ["WO-100", w(2, 28), "Alice Johnson", "B001", "555-01", "Wing Spar Machining", "0010", 10.0, "ACTIVE"],
    ["WO-100", w(3, 7),  "Alice Johnson", "B001", "555-01", "Wing Spar Machining", "0010", 30.0, "CLOSE"],
    ["WO-100", w(3, 7),  "Bob Lee",       "B002", "555-01", "Wing Spar Machining", "0010", 20.0, "CLOSE"],

    # WO-100/0020 — Alice charges, never closes (unmatched)
    ["WO-100", w(3, 7),  "Alice Johnson", "B001", "555-01", "Wing Spar Machining", "0020", 5.0,  "ACTIVE"],

    # WO-300/0030 — entirely unmatched (no earned), part 999-05
    ["WO-300", w(3, 14), "Alice Johnson", "B001", "999-05", "Final Inspection",    "0030", 10.0, "ACTIVE"],
    ["WO-300", w(3, 14), "Charlie Mendez","B003", "999-05", "Final Inspection",    "0030", 8.0,  "ACTIVE"],

    # WO-400/0010 again (duplicate of file A's last row)
    ["WO-400", w(3, 21), "Bob Lee",       "B002", "555-02", "Sheet Metal Cut",     "0010", 20.0, "CLOSE"],

    # WO-500/0010 again
    ["WO-500", w(4, 11), "Bob Lee",       "B002", "555-01", "Wing Spar Machining", "0010", 30.0, "CLOSE"],
]

make_xlsx("test-charges-555-777.xlsx", file_a_rows, "555-* and 777-* parts")
make_xlsx("test-charges-5x-and-999.xlsx", file_b_rows, "5*-* wildcard + 999-* parts")

# ----------------------------------------------------------------------------
# Hours Earned CSV
# ----------------------------------------------------------------------------
earned_lines = [
    "Order No,Year of Actual End Date,Month of Actual End Date,Oper No,Actual End Date,Hours Earned",
    "WO-100,2026,March,0010,3/7/2026,60",
    "WO-200,2026,April,0010,4/4/2026,40",
    "WO-400,2026,March,0010,3/21/2026,25",
    "WO-500,2026,April,0010,4/11/2026,25",
    # Note: WO-100/0020 and WO-300/0030 deliberately absent → unmatched test
]
(OUT / "test-hours-earned.csv").write_text("\n".join(earned_lines), encoding="utf-8")
print("Wrote test-hours-earned.csv")

# ----------------------------------------------------------------------------
# EV Details CSV (optional) — exposes Program filter
# ----------------------------------------------------------------------------
ev_lines = [
    "Network,Order No,Part No,Plan Title,Program",
    "NET-A,WO-100,555-01,Wing Spar Machining,Aero-Alpha",
    "NET-A,WO-200,777-12,Fuselage Drill,Aero-Alpha",
    "NET-B,WO-300,999-05,Final Inspection,Aero-Beta",
    "NET-A,WO-400,555-02,Sheet Metal Cut,Aero-Alpha",
    "NET-A,WO-500,555-01,Wing Spar Machining,Aero-Alpha",
]
(OUT / "test-ev-details.csv").write_text("\n".join(ev_lines), encoding="utf-8")
print("Wrote test-ev-details.csv")

# ----------------------------------------------------------------------------
# Expected results summary (for hand-verification)
# ----------------------------------------------------------------------------
print("\n=== Expected results (March 2026 + April 2026) ===")
print("Charges actuals (incremental):")
print("  Alice: WO-100/0010 30h (Mar)  +  WO-100/0020 5h (Mar)  +  WO-300/0030 10h (Mar) = 45h")
print("  Bob:   WO-100/0010 20h (Mar)  +  WO-400/0010 5+7+8=20h (Mar)  +  WO-500/0010 30h (Apr) = 70h")
print("  Charlie: WO-200/0010 50h (Apr)  +  WO-300/0030 8h (Mar) = 58h")
print()
print("Earned (split + close month):")
print("  Alice: WO-100/0010 60×30/50 = 36h (Mar)  → 36h earned")
print("  Bob:   WO-100/0010 60×20/50 = 24h (Mar) + WO-400/0010 25h (Mar) + WO-500/0010 25h (Apr) = 74h earned")
print("  Charlie: WO-200/0010 40h (Apr) → 40h earned")
print()
print("CPI (filtered, all months):")
print("  Alice:   36 / 45 = 0.80")
print("  Bob:     74 / 70 = 1.057")
print("  Charlie: 40 / 58 = 0.69")
print()
print("Split operations: WO-100/0010 (Alice 30/Bob 20)")
print("Unmatched charges: WO-100/0020 (5h, part 555-01), WO-300/0030 (18h, part 999-05)")
print("  Ranked: 999-05: 18h (top), 555-01: 5h")
