# Gate Tracker Infographic — NotebookLM Source Document

---

## Prompt for NotebookLM

Create a professional infographic describing the Gate Tracker manufacturing workflow end-to-end.

**Visual style:** Clean, modern manufacturing/industrial aesthetic. Flat design with bold icons and clear typography. Feels like a process guide for a factory floor or onboarding deck. Avoid clutter — keep bullet points short (3 per step max) and let icons do the heavy lifting.

**Format:** Horizontal flow diagram with 8 numbered steps, grouped into 4 color-coded phases. Show data flow arrows between steps. Include a feedback loop arrow from Step 8 back to Step 3 labeled "Repeats Daily."

**Color coding by phase:**
- Phase 1 (Setup) — Teal/dark cyan
- Phase 2 (Daily Preparation) — Blue
- Phase 3 (Shop Floor Work) — Orange/amber
- Phase 4 (Review & Action) — Green

**Key elements to highlight visually:**
- The transition from desk work to floor work and back to desk
- The three discovery alerts: split-location warnings, not-found-on-rack, and missing-from-MES
- The daily cycle nature of Phases 2 through 4
- The technician mapping step and its role in workload balancing

**Visual density guidance:** Aim for clean and scannable. Each step should have a large icon, a short title, and no more than 3 short bullet points. Save detail for the source material below — the infographic should communicate flow and purpose at a glance.

---

## Source Material: Gate Tracker End-to-End Process

### What is Gate Tracker?

Gate Tracker is a browser-based tool for manufacturing that visualizes where work orders sit in a gate sequence. Physical racks on the shop floor correspond to "gates" — stages in the manufacturing process. The tool shows which rack each piece of hardware belongs on, based on its current MES (Manufacturing Execution System) operation status.

It runs entirely in the browser — no server, no network calls, no installed software. Users open a single HTML file, load CSV data files, and get an interactive visual map of the shop floor.

---

### The Problem It Solves

Manufacturing floors have dozens of racks (gates) with hardware at various stages. Without a visual tool:

- It is hard to know which gate a work order belongs at based on its current MES operation.
- Hardware gets misplaced or left on a hold rack indefinitely.
- A single work order's hardware can end up separated across multiple racks — requiring engineering intervention.
- There is no easy way to reconcile the MES data with what is physically on the floor.
- Supervisors lack a clear, printable view for daily standups.

Gate Tracker bridges the gap between the MES database and the physical shop floor.

---

### Key Concepts

**Gate** — A physical rack on the shop floor corresponding to a manufacturing stage (e.g., Gate 2 = Machining, Gate 5 = Final Inspection).

**Work Order** — A manufacturing job tracked in the MES. Each work order moves through a sequence of operations.

**Operation** — A single step within a work order (e.g., CNC Rough Cut, Deburr). Statuses: Active, In Queue, Pending, or Complete.

**Decoder Ring** — A mapping file that assigns each part number's operations to specific gates. Answers: "For part X, which gate does operation Y belong to?"

**Program** — A product line or contract (e.g., PHOENIX, NEPTUNE). Used for color coding and filtering.

**Daily Upload** — A daily MES snapshot (~7,500 rows) covering all active work orders, operations, statuses, and hold flags.

**Mobile Audit** — A physical walk-through of the shop floor using a mobile-friendly HTML file to record which work orders are on which racks.

**Gate Technicians** — A CSV mapping technician names to gates with priority ordering. Displays who is assigned to each gate and is updated when workload shifts require reassignment.

---

### The End-to-End Process

The workflow has 8 steps organized into 4 phases. Phase 1 is setup (infrequent). Phases 2 through 4 repeat daily.

---

#### Phase 1: Setup

##### Step 1 — Create Decoder Ring

The Decoder Ring maps a part number's MES operations to physical gates. It is the foundation of the system.

- Open the Decoder Ring Builder tool (a separate single-file HTML app).
- Upload the MES operations list for the part number, then drag and drop operation cards into gate outlines.
- Gates can repeat (e.g., Gate 3 → 5 → 3) to support re-entry patterns.
- Export as a CSV. Reuse daily — only rebuild when the part's operation sequence changes.

##### Step 2 — Create / Update Technician Mapping

The Gate Technicians file maps technicians to gates so the team knows who is responsible for each rack.

- Build a CSV with columns: Gate Number, Gate Name, Technician Name, Priority.
- One row per technician per gate. Priority sets display order in gate headers.
- **Update this file whenever workload shifts.** During surge weeks — when one gate has heavy volume — reassign technicians from lighter gates to support the bottleneck. This is a manual but critical workforce-balancing step that happens as often as weekly.
- The same file feeds into both Gate Tracker (gate header display) and the Decoder Ring Builder (gate dropdown pre-population).

---

#### Phase 2: Daily Preparation (At Your Desk)

##### Step 3 — Download Daily Data

- Export the MES snapshot from the internal manufacturing system.
- Tab-delimited CSV, ~7,500 rows: all active work orders with operation statuses, hold flags, programs, part numbers, serial numbers, and average days at current operation.

##### Step 4 — Load Files into Gate Tracker

- Open Gate Tracker (single HTML file) in Chrome or Edge.
- Load: Decoder Ring CSV, Daily Upload CSV, and Gate Technicians CSV.
- Click Process. Gate Tracker normalizes operation numbers and generates the Detail View — per-order progress bars grouped by Program → Part Number → Order Number, with color-coded operation chips.

##### Step 5 — Generate Mobile Audit

- Click "Generate Mobile Audit" in Gate Tracker.
- Downloads a self-contained HTML file with all work order data embedded (no network needed).
- Email the file to yourself to open on your phone.

---

#### Phase 3: Shop Floor Work (Walking the Floor)

##### Step 6 — Conduct Floor Audit

- Open the mobile audit HTML on your phone and walk the floor gate by gate.
- At each rack: tap the gate button, search orders by last 4 digits, and tap to record.
- "Already on Gate X" warnings flag split-location situations — the same order's hardware is on multiple racks.
- Use manual entry for orders visible on the rack but not found in MES data.
- When done, tap "Copy to Clipboard" and email the results to yourself.
- A typical audit takes 30–60 minutes.

---

#### Phase 4: Review and Action (Back at Desk)

##### Step 7 — Upload Audit Results

- Save the emailed audit data as a CSV (columns: Gate Number, Order No., Part No., Serial No., Technician).
- Load it as the fourth input in Gate Tracker and click Process again.
- Gate Tracker generates the Gate View: audit-based physical placement cross-referenced with MES status data.

##### Step 8 — Review and Act

The Gate View is the primary output — where each work order was physically found, enriched with MES data.

**What you see:**
- Work order cards arranged by gate, color-coded by program.
- Technician names in gate headers (from the technician CSV).
- DR Operations (red border) for disposition/rework ops. QA Operations (blue border) for quality ops.

**Three discovery alerts for action:**

1. **Split Locations** — Order found at multiple racks. Reunite hardware with manufacturing engineer.
2. **Not Found on Rack** — Active MES orders not found during audit. Investigate: in transit, mislabeled, or missing?
3. **Missing from MES** — Found on floor but not in the MES snapshot. Check for data lag or recent creation.

**Sharing:** Print/PDF for standup meetings, or save a self-contained interactive HTML snapshot to email to colleagues.

---

### Daily Cycle

Phases 2–4 repeat every workday. Today's audit flags inform tomorrow's focus: "not found" orders get hunted, split-location items get verified, hold rack items get checked for release.

The Decoder Ring (Step 1) is rebuilt only when a part number's operation sequence changes. The Technician Mapping (Step 2) is updated whenever workload shifts require tech reassignment — as often as weekly during surge periods.

---

### Users

- **Production planners** — track order progress and gate placement for daily work assignments.
- **Manufacturing engineers** — investigate split-location alerts and resolve hardware placement issues.
- **QA leads** — monitor QA-flagged orders and inspect for gate bottlenecks.
- **Shop floor supervisors** — use printed Gate View reports in daily standups.

---

### Summary of Data Flow

```
MES Database
    |
    v
Daily Upload CSV (~7,500 rows)  +  Decoder Ring CSV  +  Gate Technicians CSV
    |                                   |                       |
    +-----------------------------------+-----------------------+
    |
    v
Gate Tracker (browser) — Process
    |
    +---> Detail View (operation-level progress)
    +---> Mobile Audit HTML (generated for phone)
              |
              v
         Floor Audit (walking the racks, 30-60 min)
              |
              v
         Audit CSV (emailed back to desk)
              |
              v
         Gate Tracker (browser) — Process again with audit file
              |
              +---> Gate View (audit-based placement map)
              +---> Split Location Warnings
              +---> Not Found on Rack list
              +---> Missing from Daily Upload flags
              +---> Print/PDF for standup meetings
              +---> Shareable Snapshot HTML
```

---

## Suggested Infographic Layout

### Overall Structure
- Horizontal flow moving left to right.
- 8 numbered steps in a gentle S-curve or straight line.
- Steps grouped into 4 color-coded phase bands.
- Data flow arrows connecting each step.
- Dashed feedback loop arrow from Step 8 back to Step 3 labeled "Repeats Daily."

### Phase Groupings and Colors

| Phase | Steps | Color | Label |
|-------|-------|-------|-------|
| Phase 1: Setup | Steps 1–2 | Teal / dark cyan | Setup |
| Phase 2: Daily Prep | Steps 3–5 | Blue | At Your Desk |
| Phase 3: Floor Work | Step 6 | Orange / amber | On the Shop Floor |
| Phase 4: Review | Steps 7–8 | Green | Back at Desk |

### Suggested Icons per Step

1. **Create Decoder Ring** — Puzzle piece or grid/matrix icon
2. **Technician Mapping** — People/team icon or clipboard with person
3. **Download Daily Data** — Download arrow with spreadsheet
4. **Load Files & Process** — Upload/import icon with gear
5. **Generate Mobile Audit** — Smartphone icon
6. **Conduct Floor Audit** — Walking person with clipboard
7. **Upload Audit Results** — Upload arrow with checkmark
8. **Review & Act** — Dashboard/chart icon with magnifying glass

### Key Callouts to Feature

Three discovery alert boxes branching from Step 8:

- **Split Locations** (orange) — "Hardware separated across racks"
- **Not Found on Rack** (amber) — "Active orders missing from floor"
- **Missing from MES** (amber) — "Found on floor but not in system"

### Additional Visual Elements

- "Desk" icon above Phases 2 and 4; "factory floor" icon above Phase 3.
- File flow arrows: Decoder Ring CSV and Technicians CSV flowing into Step 4, Daily Upload CSV flowing into Step 4, Mobile Audit HTML flowing from Step 5 to Step 6, Audit CSV flowing from Step 6 to Step 7.
- Subtitle: "Bridging the gap between MES data and the physical shop floor."
- Footer: "Steps 1–2 are setup. Steps 3–8 repeat daily. Update technician assignments as workload shifts."

### Tool / File Quick Reference (optional small table)

| Tool / Feature | File Type | Frequency |
|---|---|---|
| Decoder Ring Builder | .CSV (Mapping) | One-time per Part |
| Technician Mapping | .CSV (People) | As needed (weekly during surges) |
| Gate Tracker Main | .CSV (MES Data) | Daily |
| Mobile Audit App | .HTML (Offline) | Daily |
