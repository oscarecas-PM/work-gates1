# Architecture Proposal: Gate Registry as Single Source of Truth

**Date:** March 20, 2026
**Status:** Proposal — no files changed

---

## The Problem You Identified

You said it well: the Gate Number ↔ Gate Name pairing never changes. Gate 7 is always "Pressure Test." Gate 7.1 is always "Pressure Test - Data Review." But right now that pairing lives in two places — and in the Decoder Ring, it's written once per *row*, not once per gate.

### Where gate names live today

| File | Where gate names appear | How many times? |
|---|---|---|
| Gate Technicians CSV | One row per gate (with blank tech rows) | Once per gate ✓ |
| Decoder Ring CSV | Every single row that belongs to that gate | Repeated for every operation × every part number |

In the test data, Gate 4 "Welding" appears 4 times in the decoder ring (rows 6, 7, 18 for ABC-100 Welding ops, row 18 for DEF-200). In a real decoder ring with 10+ parts and 50+ operations per part, that's hundreds of repetitions of the same gate names.

### Consequences of the current design

1. **Rename a gate → touch two files.** If "Pressure Test" gets renamed to "Hydrostatic Test," you update the Gate Technicians file *and* every row in the Decoder Ring for every part that visits that gate. Miss one and the Detail View and Gate View show different names for the same gate.

2. **The app already resolves this inconsistency silently.** Gate View names come from the Tech file first, falling back to the Decoder Ring. Detail View names come *only* from the Decoder Ring (line 1035 in gate-tracker-24.html). If these disagree, users see different names in the two views — a bug waiting to happen.

3. **The Decoder Ring Builder must ask for gate names** even though they're already defined. It shows them because it loads the Gate Technicians file, but it still exports them into the CSV redundantly.

---

## The Fix: One File Owns Gate Names

### Core principle

> The Gate Registry (currently called "Gate Technicians") is the single source of truth for all gate metadata — number, name, type, and technician assignments.
>
> The Decoder Ring references gates by *number only*. It never needs to spell out the name.

This is already conceptually true — the PRD (Section 3.3) calls the Gate Technicians file "the authoritative list of physical gate locations." The proposed change makes the architecture match that stated intent.

---

## Proposed File Schemas

### Gate Registry (rename from "Gate Technicians")

No schema changes. The file already does this job correctly.

```
Gate Number, Gate Name, Tech Name, Priority
HOLD-1, Hold Rack A, ,
HOLD-2, Hold Rack B, ,
1, Receiving, Mike Johnson, 1
2, CNC Machining, Sarah Chen, 1
2, CNC Machining, Tom Rivera, 2
3, Deburr / QA, John Smith, 1
7, Pressure Test, ,
7.1, Pressure Test - Data Review, ,
...
```

**What changes:** The file is renamed conceptually from "Gate Technicians" to "Gate Registry." Tech assignments are a feature of the file, not its purpose. Rows with no Tech Name are now first-class — they're not edge cases, they're the standard way to define a gate location before assigning technicians.

**What doesn't change:** Column names, row format, anything else.

---

### Gate Decoder Ring (remove Gate Name column)

**Current:**
```
Part No., Part Name, Gate Number, Gate Name, Oper No., Work Dept, Modifier
ABC-100, Phoenix Assembly, 1, Receiving, 10, RECEIVING,
ABC-100, Phoenix Assembly, 2, CNC Machining, 20, SMART,
ABC-100, Phoenix Assembly, 2, CNC Machining, 30, SMART,
...
```

**Proposed:**
```
Part No., Part Name, Gate Number, Oper No., Work Dept, Modifier
ABC-100, Phoenix Assembly, 1, 10, RECEIVING,
ABC-100, Phoenix Assembly, 2, 20, SMART,
ABC-100, Phoenix Assembly, 2, 30, SMART,
...
```

The `Gate Name` column is removed. The app resolves the display name by looking up `Gate Number` in the Gate Registry at processing time.

**What this prevents:** You can no longer write an inconsistency into the Decoder Ring. "Gate 7" is always resolved to whatever the Gate Registry says, regardless of what any decoder ring row claims.

---

## How the App Would Change

### Processing order

1. Parse Gate Registry → build `gateNameMap: { "1": "Receiving", "2": "CNC Machining", "7": "Pressure Test", "7.1": "Pressure Test - Data Review", "HOLD-1": "Hold Rack A", ... }`
2. Parse Decoder Ring → build `gateLookup` and `gatesByPart` with gate numbers only (no names embedded)
3. Whenever a gate name is needed for display, look it up in `gateNameMap`

### Name resolution hierarchy (simplified from two conflicting sources to one)

| Today | Proposed |
|---|---|
| Detail View: uses Decoder Ring gate name | Detail View: looks up gateNameMap |
| Gate View: uses Tech file first, Decoder Ring fallback | Gate View: looks up gateNameMap |
| Mobile Audit: uses Decoder Ring gate name | Mobile Audit: looks up gateNameMap |

All three views now use the same source. Renaming a gate is a one-file operation.

### Backward compatibility

The app can keep reading `Gate Name` from the Decoder Ring if the column is present — it just ignores it. Existing decoder ring files continue to work without modification. The change is purely additive.

### What if Gate Registry isn't loaded?

The Gate Registry is currently optional. Under this proposal, it should become **strongly recommended**, with a visible warning if absent (not a hard error). If not loaded:

- Gate names fall back to `"Gate {number}"` labels everywhere
- A warning banner appears: *"Gate names are unavailable. Load the Gate Registry to see gate names in all views."*
- All functionality still works — it's a usability degradation, not a failure

---

## Secondary Observation: Part Name Redundancy

The same pattern exists for Part Name in the Decoder Ring — `ABC-100, Phoenix Assembly` repeats on every row for that part. The current code already handles this correctly (`if (!partNames[partNo]) partNames[partNo] = ...`), so it causes no bugs. But it's still extra typing.

**Recommendation:** Leave Part Name in the Decoder Ring for now. It's lightweight, and a dedicated "Part Registry" file would be over-engineering for the current scale. The redundancy here causes no inconsistency risk because the code takes the first occurrence and ignores subsequent ones. If Part Name ever needed to appear in a different context (e.g., a label on the Gate View card), this would be worth revisiting.

---

## Impact on the Decoder Ring Builder

The Decoder Ring Builder already loads the Gate Technicians file and uses it to display gate names in the workspace. The only change needed to the builder:

1. Remove `Gate Name` from the exported CSV columns
2. No other changes — the builder already has gate names from the loaded tech file

This is a small, clean change. The exported file gets simpler (6 columns instead of 7), and the gate names that were being exported as redundant data now live entirely in the Gate Registry where they belong.

---

## Summary

| | Today | Proposed |
|---|---|---|
| Where does "Gate 7 = Pressure Test" live? | Gate Technicians file + every Decoder Ring row for gate 7 | Gate Registry only |
| How many times is each gate name written? | 1 + (number of ops at that gate × number of parts through that gate) | Once |
| Can the two views show different names for the same gate? | Yes (if files disagree) | No |
| What breaks if you rename a gate? | Must update two files, many rows | Update one row |
| Gate Registry required? | No | No, but strongly recommended (warning if absent) |
| Backward compatibility | — | Old decoder rings still work (Gate Name column is ignored) |
