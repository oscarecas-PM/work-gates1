# Other Architectural Observations

**Date:** March 20, 2026
**Status:** Notes — not proposals, no priority assigned

These are patterns that emerged while reviewing the codebase in context of the gate name redundancy question. None are urgent. Listed here for the record.

---

## 1. Gate Display Order Has Two Different Authorities

**What:** Gate ordering is inconsistent between views.

- **Detail View** gate column order: defined by row order in the Decoder Ring (per-part). Each part can have a different gate sequence, which is intentional and correct (F-20, F-21 support gate re-entry).
- **Gate View** gate header order: defined by the Gate Registry / ascending gate number sort.

These are fundamentally different things — Detail View is "this part's operational sequence," Gate View is "physical floor layout." The inconsistency is arguably *correct*. Calling it out here because the gate-name architecture proposal affects both, and it's worth being intentional that the Display Order authority stays split rather than being accidentally unified.

**Verdict:** Leave as-is. The split is logical.

---

## 2. The `gatesByPart` Structure Embeds Gate Names That No Longer Need to Be There

**What:** The `gatesByPart` object (built while parsing the Decoder Ring) currently stores `{ gateNumber, gateName, opers[] }` per part per gate. Gate names are embedded in this structure at parse time.

Under the gate-registry proposal, `gatesByPart` would only need `{ gateNumber, opers[] }`. Names would be resolved on demand from `gateNameMap`.

This simplifies the data structure and eliminates the risk of a stale name surviving into a display path. Minor code cleanup, not a separate feature.

---

## 3. The Mobile Audit Gets Gate Names from the Decoder Ring, Not the Tech File

**What:** When generating the Mobile Audit HTML (the tool that runs on the phone), the list of gates and their names is built from `gatesByPart`, which comes from the Decoder Ring (line ~845 in gate-tracker-24.html). The Gate Technicians file is used for *technician buttons* but not for gate name labels.

This means if a gate is defined in the Gate Technicians file but has no parts in the current Decoder Ring, it won't appear in the Mobile Audit. It also means the Mobile Audit gate names can diverge from Gate View gate names.

Under the gate-registry proposal, both would use `gateNameMap` and this inconsistency disappears.

---

## 4. "Gate Technicians" Is a Misleading Filename

**What:** The file is described in the PRD as "the authoritative list of physical gate locations." But its name ("Gate Technicians") implies its primary purpose is to list technicians. This leads users to think of it as an optional add-on rather than the gate master list.

**Suggestion:** The user-facing label in the HTML file input could be changed from "Gate Technicians (optional)" to "Gate Registry (optional)" or "Gate Master File (optional)." The file format doesn't change — just the framing in the UI. This reinforces the architectural intent.

---

## 5. Open Question: Should the Gate Registry Become Required?

Currently optional. The main reason it's optional is backward compatibility with the earliest versions of the app. At this point, most users are expected to have one.

**Arguments for making it required:**
- Enforces the single source of truth for gate names
- Removes the fallback code paths that resolve names from the Decoder Ring
- Users already have it — it's how hold racks and sub-gates are defined

**Arguments for keeping it optional:**
- "Required" means the Process button won't activate without it
- Someone doing a quick one-off analysis shouldn't need to prepare a registry file
- Gate numbers as fallback labels are good enough for understanding the data

**Suggested middle ground:** Keep it optional, but display a persistent warning banner (not a modal) when it's absent: *"No Gate Registry loaded — gate names will not be shown. Load the Gate Registry file for full labels in all views."* This nudges users without blocking them.

---

## 6. The Decoder Ring Builder Exports Data the New Architecture Doesn't Need

**What:** The Builder currently exports 7 columns: `Part No., Part Name, Gate Number, Gate Name, Oper No., Work Dept, Modifier`. Under the gate-registry proposal, `Gate Name` becomes unnecessary in the export (it's now owned by the Gate Registry).

The Builder already *has* the gate name at export time (it loaded the Gate Technicians file to populate the workspace). It just wouldn't write it to the output CSV.

This is a 5-line change to the Builder's export function.

---

## 7. Potential Future: A Part Registry

If Part Name redundancy ever becomes a maintenance concern (e.g., a part gets a new program name and dozens of Decoder Ring rows need updating), a lightweight "Part Registry" file could be introduced:

```
Part No., Part Name
ABC-100, Phoenix Assembly
DEF-200, Neptune Housing
```

The Decoder Ring would then contain only `Part No., Gate Number, Oper No., Work Dept, Modifier` — pure operation-to-gate mappings with no display metadata at all.

This is probably over-engineering for the current scale. The current code already deduplicates Part Names at parse time (taking only the first occurrence), so there's no consistency risk. Flagged here only as a natural extension of the gate-registry pattern if the tool grows significantly.
