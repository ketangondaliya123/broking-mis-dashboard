# Broking MIS Dashboard — Reusable Regeneration Prompt

Paste this whole document into a new conversation (Project: Broking MIS Dashboard), with the
updated data file(s) uploaded, to regenerate the full analytics suite.

---

## 1. Context

I run a Broking MIS Dashboard analytics suite for Marwadi Shares and Finance Ltd. (financial
broking firm). Each period, I upload a new operational MIS data file and need the full suite of
deliverables rebuilt against it, compared against the standing full-fiscal-year (FY) file.

**Base/reference file (rarely changes):** the full-fiscal-year file, e.g.
`Opr_MIS_Data_Apr25_to_Mar26.xlsx`. Treat whichever full-FY file is present as the fixed base
period unless I say otherwise. Always label it explicitly as **"FY 2025-26"** (or whatever fiscal
year applies) — never just "FY" or "base period" — in every header, tab title and dashboard label.

**Rolling/new file (changes each time):** a shorter-period file, e.g.
`Opr_MIS_Data_Apr26_to_Jul26.xlsx`. **Read the period length from the filename/date range
yourself — do not assume it is a calendar quarter.** Label it explicitly too, e.g. "Apr'26–Jul'26
(FY 2026-27, 4 months actual)".

**Comparison methodology — pro-rata the NEW period up, FY stays actual:** scale the new period's
actual figures *up* to a full-year equivalent (multiply by `12 / months covered`), and compare
that against the FY figure **as reported, unscaled**. E.g. for a 4-month new period:
`New Period pro-rata = New Period actual × 3`, compared against `FY actual` (never divided). This
was reversed once already (an earlier version divided FY down instead) — if in doubt, confirm
which direction I want before rebuilding, since getting it backwards means redoing every
deliverable.

**Three-layer structure — required in every workbook and the dashboard, for both Channels and
Clients:**
  1. **FY actual** — the full base-period figure, unscaled, shown standalone.
  2. **New Period actual** — the new period's figure, unscaled, shown standalone.
  3. **Comparison** — New Period actual × pro-rata factor, set against FY actual, with growth %.

Layers 1 and 2 are not just hidden helper columns feeding layer 3 — they must be full, independent,
browsable views (their own sheet/tab in Excel, their own sub-tab in the dashboard), shown before
the comparison layer.

## 2. Source file structure (confirm this still holds — it has been stable across three files)

Both the FY file and the new-period file have two sheets:

- **Channelwise** — branch/sub-broker level. Columns: Branch Code, First Date of Trading,
  Channel ID, Channel Type (Branch/SB), Group, Channel Name, City, State, Turnover, Gross
  Brokerage, Passout, Net Brokerage, Yield. One grand-total row at the bottom (Branch Code is
  blank there) — exclude it.
- **Clientwise** — end-client level. Columns: Branch Code, Branch Name, Client Code, Client
  Name, City, Turnover, Gross Brokerage, Yield. Also has a trailing grand-total row — exclude it.

**Critical join fact:** Clientwise's "Branch Code" column is actually the **Channel ID** from
Channelwise, not the true parent branch. To resolve the real parent branch, Group, and Channel
Type for a client row, join Clientwise.`Branch Code` → Channelwise.`Channel ID`, and read
Channelwise's own `Branch Code` column as the true parent.

**Channel Name mapper (every client-level table):** every client row must carry a "Channel Name"
column showing which channel the client trades through — the same join as above. A small number
of clients trade through more than one channel in a period; use the channel where they generated
the most brokerage as their "primary" one rather than trying to show all of them.

**House/proprietary account:** Marwadi Shares and Finance Ltd.'s own book appears as 4 client
codes under Channel ID `A1H`: `A11`, `A10287`, `A111111`, and `ERROR`. Flag these
(`Is House Account`) and **exclude them from all client rankings and client-level group
summaries** — they carry disproportionately large turnover with zero brokerage and will distort
any ranking if left in.

**Groups (Channelwise column E):** MSFL, MSFL-Sharing, Arbitrage, P-Sec. Confirm this list hasn't
changed before hardcoding it into group-rollup formulas.

**Column scope for channel-level analysis/comparison tables (not raw data dumps):** exclude
Turnover as a visible figure — keep it only as a hidden helper column where needed to compute a
correctly-weighted Yield %. Show **Passout** instead. Client-level tables keep Turnover (no
Passout exists at client grain in the source data). **Yield % must appear in both channel-level
and client-level analysis tables.**

**Reconciliation:** Check the turnover/brokerage sum between Channelwise and Clientwise on each
file. The FY file has historically shown a ~1% gap (Channelwise is the source of truth when it
does); newer shorter-period files have reconciled exactly. Report whatever gap you find, don't
assume it matches prior runs.

**Client identity across branches:** A given Client Code can appear under more than one Channel
ID/branch within a single period (same client, same name, trading through two channels). When
building client-level rollups, group by Client Code + Client Name and sum Turnover/Gross
Brokerage across branches first.

**Branch-level drill-down (grouping):** every channel- and client-level table should let me
collapse individual branches down to a branch-total row, or collapse everything at once, then
expand back out — each branch typically has ~35 channels/clients nested under it.
  - **In Excel:** native row Group/Outline (not a PivotTable — pivot caches built
    programmatically are fragile). Detail rows carry `outlineLevel = 1`; the branch's subtotal
    row carries `outlineLevel = 0`. **Put the subtotal row *after* its detail rows, not before** —
    a custom `summaryBelow = False` setting does not survive a LibreOffice recalc round-trip and
    silently reverts to Excel's true default (`summaryBelow = True`), which breaks a
    header-first layout. Build for `summaryBelow = True` from the start.
  - **In the dashboard:** a "Grouped by Branch" view (toggle alongside the flat/search view) with
    collapsible branch headers showing a running total, "Expand All"/"Collapse All" controls, and
    each branch's detail table lazy-rendered only the first time it's expanded (don't build all
    ~120 branches' worth of detail HTML upfront — the flat/searchable view is the fast default,
    grouped is opt-in).

## 3. Deliverables (5 files, every time)

1. **New Period Analysis workbook** (standalone, for the new file only)
   - README, branch-grouped Channelwise, branch-grouped Clientwise (with Channel Name mapper),
     Group Summary - Channel, Group Summary - Client, Top 100 / Top 500 Clients, Lists (house
     account reference).
   - Channel-level figures: Turnover excluded (kept only as a hidden helper for weighted Yield),
     Passout shown, Yield % shown. Client-level: Turnover kept, Yield % shown, Channel Name shown.
   - Top 100/500 Clients: pre-rank in pandas and write a flat list rather than live in-sheet
     AGGREGATE ranking — the branch-grouped Clientwise sheet has subtotal rows interspersed
     through the range, which would corrupt a range-based ranking formula.

2. **Comparative Analysis workbook** (channel-level, 3-layer)
   - README, "FY 2025-26 Actual" (branch-grouped, standalone), "Apr-Jul 2026 Actual"
     (branch-grouped, standalone), "Channel Comparison" (branch-grouped; New Period actual × pro-
     rata factor vs. FY actual, unscaled; growth % per metric), "Group Comparison" (SUMIF rollup
     of Channel Comparison — SUMIF naturally skips the branch subtotal rows since they carry a
     blank Group value, so no special-casing needed there).

3. **Client-Level Comparison workbook(s), split into two files** — combining the full 3-layer
   client structure in one workbook exceeded this container's memory during LibreOffice recalc
   (confirmed by monitoring memory climb to the ~4GB ceiling, not a timeout — verify with
   `free -h` during a retry before assuming it's just slow). Splitting resolved it:
   - **File A — actuals:** "FY 2025-26 Actual" and "Apr-Jul 2026 Actual" standalone client lists
     (branch-grouped, Channel Name mapper, full universe for each period — tens of thousands of
     rows each).
   - **File B — comparison:** README, Matched Clients (branch-grouped comparison), New Clients,
     Lost Clients, All / Top 100 / Top 500 Clients (flat, pre-ranked in pandas — same reasoning
     as WB1's Top 100/500), Top 20 Gainers / Top 20 Decliners (ranked by **absolute ₹ change**,
     New Period pro-rata − FY actual, not growth %, since a tiny FY base produces misleading
     percentage swings; growth % still shown for context).
   - Client-level figures keep Turnover (no Passout at client grain in the source) and show
     Yield % and Channel Name throughout.

4. **Self-contained HTML dashboard**
   - SheetJS (xlsx) installed via `npm install` and inlined directly into the HTML (no CDN
     references — must work fully offline). **No charting library — charts/graphs are explicitly
     not wanted; use data tables for everything**, including gainers/decliners (a small in-cell
     CSS magnitude bar is fine for visual interest, but not a canvas chart).
   - Inject data via a large embedded JSON blob assembled by a Python script (don't hand-type
     megabytes of data through the document-writing tool).
   - Sections: Overview (KPI cards + group-wise summary table + a channel-type breakdown table +
     a top-10-branches table — favor **more data tables** over fewer, richer visuals), Channels
     (3-layer sub-nav: FY Actual / New Period Actual / Comparison, in that order), Clients (same
     3-layer sub-nav, every row carries Channel Name), Gainers & Decliners (tables only), Lost
     Clients.
   - Channels and Clients layers each need a "Flat / Search" vs. "Grouped by Branch" view toggle
     — grouped view shows collapsible branch headers with a running total and Expand All/Collapse
     All controls; lazy-render each branch's detail rows only the first time it's expanded so the
     page stays responsive with 500+ branches × tens of thousands of rows underneath.
   - Every table header must be **sticky** (stays pinned while scrolling that table's rows) —
     wrap tables in a `max-height` + `overflow:auto` container with `thead th{position:sticky;
     top:0}`, not just `freeze_panes` in Excel. Headers should wrap to two lines rather than
     truncate or force the column wider than its content needs.
   - For a client-level "actual" layer that spans both matched and lost/new clients, don't
     duplicate a full raw dataset in the JSON if it can be composed client-side from data you
     already have (e.g. FY-actual-all-clients = matched-clients' FY columns ∪ lost-clients) —
     keeps the embedded JSON smaller.
   - Export to Excel (SheetJS, respects current filter/sort) and Export to PDF (print-styled
     `window.print()`, no extra library needed) on every table view.
   - Every KPI card, table header, and section label must name the actual period it represents
     (e.g. "FY 2025-26", "Apr'26–Jul'26 Actual") — never a bare "FY" or "New Period" with no year.
   - Design: avoid generic AI-dashboard defaults. A "financial ledger/statement" aesthetic has
     worked well here — deep navy letterhead, ivory/paper content area, serif headers, monospace
     figures, gold accent, green/red for gains/losses.
   - **Before shipping:** render it headlessly with Playwright, check for console/page errors,
     click through every tab and sub-tab, test search/sort/filter/export interactions and the
     sticky-header behavior (scroll a table and screenshot to confirm the header stayed put), and
     screenshot each view. Fix anything found (a real bug caught this way before: display-column
     order didn't match raw-data-array order, so clicking "sort by Status" actually sorted by the
     wrong column — always map sort clicks through the column's key, not its display position).

## 4. Excel technical patterns (carried over from prior sessions — keep using these)

- Prefix `AGGREGATE` as `_xlfn.AGGREGATE` (plain `AGGREGATE` → `#NAME?` in LibreOffice recalc).
- Array formulas via `openpyxl.worksheet.formula.ArrayFormula`.
- LARGE/SMALL ranking needs a tie-break helper column (`value + ROW()*1e-10`) to avoid duplicate-
  value MATCH collisions.
- Exclude rows from Top-N ranking with a **sentinel value**, not a forced division error.
- **Freeze panes (title + header rows, and any leading ID columns) on every data sheet** so
  headers stay visible when scrolling — set this explicitly on each sheet, don't rely on a
  default.
- Avoid O(n²) SUMPRODUCT/COUNTIFS/SUMIFS on large per-entity ranges (tens of thousands of unique
  keys against tens of thousands of raw rows) — pre-aggregate once in pandas instead, and reserve
  in-sheet formulas for small dimension tables (Group-level, a few rows) or cheap same-row
  arithmetic.
- Display units throughout: Turnover in ₹ Crore, Brokerage/Payout/Net Revenue in ₹ Lacs, Yield as
  plain percentages.

## 5. Validation checklist before delivering

- [ ] Every workbook recalculated headlessly (LibreOffice, 280–300s timeout for large files) with
      **zero formula errors** — not "errors we expect and are fine with."
- [ ] Spot-check computed totals against a raw pandas aggregation of the source files.
- [ ] Confirm the annualization factor matches the actual period length in the new file, and that
      this is stated in the response and in every README tab.
- [ ] HTML dashboard rendered and interacted with via Playwright — no console errors, search/sort/
      filter/export all verified against known values.
- [ ] Flag anything unusual found in the data (large swings, new reconciliation behavior, changed
      group list, etc.) rather than silently reproducing it.

## 6. Output file naming

- `1_New_Period_Analysis_<period>.xlsx`
- `2_Comparative_Analysis_FY_vs_NewPeriod_Annualized.xlsx`
- `3_Client_Level_Comparison_FY_vs_NewPeriod.xlsx`
- `Broking_MIS_Dashboard_<period>.html`

Note the dashboard file will likely be several MB (data volume, not bloat) — recommend sharing
via cloud storage rather than email, since HTML attachments are often blocked.
