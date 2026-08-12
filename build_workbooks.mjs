import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const ROOT = path.dirname(decodeURIComponent(new URL(import.meta.url).pathname)).replace(/^\/(?:[A-Za-z]:)/, (m) => m.slice(1));
const WORK = process.argv[2] || "all";
const dataFile = { new: "data_new.json", channel: "data_channel_comp.json", actuals: "data_client_actuals.json", comp: "data_client_comp.json", all: "data.json" }[WORK] || "data.json";
const data = JSON.parse(await fs.readFile(path.join(ROOT, "build", dataFile), "utf8"));
const outDir = path.join(ROOT, "outputs");
const previewDir = path.join(ROOT, "build", "previews");
await fs.mkdir(outDir, { recursive: true });
await fs.mkdir(previewDir, { recursive: true });

const C = {
  navy: "#10253F", navy2: "#173A5E", ivory: "#FAF7F0", paper: "#F3EFE6", gold: "#C9A227",
  ink: "#1C2833", muted: "#617182", line: "#D9D4C8", green: "#166534", red: "#9F1239",
};

function col(n) {
  let s = "";
  for (let x = n + 1; x > 0; x = Math.floor((x - 1) / 26)) s = String.fromCharCode(65 + ((x - 1) % 26)) + s;
  return s;
}
function endCol(headers) { return col(headers.length - 1); }
function rangeRef(headers, startRow, endRow) { return `A${startRow}:${endCol(headers)}${endRow}`; }
function safeSheetName(name) { return name.slice(0, 31); }
function fmtHeader(sheet, headers, row = 4) {
  const r = sheet.getRange(`A${row}:${endCol(headers)}${row}`);
  r.values = [headers];
  r.format = { fill: C.navy2, font: { bold: true, color: "#FFFFFF", size: 10 }, wrapText: true, horizontalAlignment: "center", verticalAlignment: "center", borders: { preset: "all", style: "thin", color: C.navy2 } };
  r.format.rowHeight = 34;
}
function writeRows(sheet, headers, rows, startRow = 5, chunk = 5000) {
  for (let i = 0; i < rows.length; i += chunk) {
    const part = rows.slice(i, i + chunk);
    sheet.getRange(`A${startRow + i}:${endCol(headers)}${startRow + i + part.length - 1}`).values = part;
  }
}
function applyWidths(sheet, headers, endRow, options = {}) {
  const defaults = {
    "Branch Code": 12, "Channel ID": 12, "Channel Type": 12, "Group": 16, "Channel Name": 34, "City": 18, "State": 16,
    "Branch Name": 30, "Client Code": 16, "Client Name": 32, "Is House Account": 14, "Row Type": 14, "Status": 12,
  };
  for (let i = 0; i < headers.length; i++) {
    const h = headers[i];
    let w = options[h] ?? defaults[h] ?? 15;
    if (h.includes("[helper]")) w = 0.2;
    const rr = sheet.getRange(`${col(i)}1:${col(i)}${Math.max(5, endRow)}`);
    rr.format.columnWidth = w;
  }
}
function formatNumeric(sheet, headers, endRow) {
  const formats = {
    "Turnover (₹ Crore)": "#,##0.00", "Turnover (₹ Crore) [helper]": "#,##0.00", "FY Turnover (₹ Crore) [helper]": "#,##0.00", "New Turnover Actual (₹ Crore) [helper]": "#,##0.00", "Annualized New Turnover (₹ Crore) [helper]": "#,##0.00",
    "FY Turnover (₹ Crore)": "#,##0.00", "New Turnover Actual (₹ Crore)": "#,##0.00", "Annualized New Turnover (₹ Crore)": "#,##0.00",
    "Gross Brokerage (₹ Lacs)": "#,##0.00", "Passout (₹ Lacs)": "#,##0.00", "Net Brokerage (₹ Lacs)": "#,##0.00",
    "FY Gross (₹ Lacs)": "#,##0.00", "FY Passout (₹ Lacs)": "#,##0.00", "FY Net (₹ Lacs)": "#,##0.00",
    "New Gross Actual (₹ Lacs)": "#,##0.00", "New Passout Actual (₹ Lacs)": "#,##0.00", "New Net Actual (₹ Lacs)": "#,##0.00",
    "Annualized New Gross (₹ Lacs)": "#,##0.00", "Annualized New Passout (₹ Lacs)": "#,##0.00", "Annualized New Net (₹ Lacs)": "#,##0.00",
    "Turnover (₹ Crore) [helper]": "#,##0.00", "FY Turnover (₹ Crore) [helper]": "#,##0.00", "New Turnover Actual (₹ Crore) [helper]": "#,##0.00", "Annualized New Turnover (₹ Crore) [helper]": "#,##0.00",
    "FY Yield %": "0.00", "New Yield Actual %": "0.00", "Annualized New Yield %": "0.00", "Yield %": "0.00",
    "Gross Growth %": "0.0%", "Passout Growth %": "0.0%", "Net Growth %": "0.0%", "Yield Growth %": "0.0%",
    "Gross Growth %": "0.0%", "Absolute Gross Change (₹ Lacs)": "#,##0.00",
  };
  for (let i = 0; i < headers.length; i++) {
    const f = formats[headers[i]];
    if (f) sheet.getRange(`${col(i)}${5}:${col(i)}${endRow}`).format.numberFormat = f;
  }
}
function baseSheet(wb, name, title, subtitle, headers, rows, opts = {}) {
  const sheet = wb.worksheets.add(safeSheetName(name));
  sheet.showGridLines = false;
  const last = endCol(headers);
  sheet.getRange(`A1:${last}1`).merge();
  sheet.getRange("A1").values = [[title]];
  sheet.getRange(`A1:${last}1`).format = { fill: C.navy, font: { bold: true, color: "#FFFFFF", size: 15 }, horizontalAlignment: "left", verticalAlignment: "center" };
  sheet.getRange("A1").format.rowHeight = 28;
  sheet.getRange(`A2:${last}2`).merge();
  sheet.getRange("A2").values = [[subtitle]];
  sheet.getRange(`A2:${last}2`).format = { fill: C.paper, font: { italic: true, color: C.muted, size: 10 }, wrapText: true, verticalAlignment: "center" };
  sheet.getRange("A2").format.rowHeight = 28;
  fmtHeader(sheet, headers, 4);
  writeRows(sheet, headers, rows, 5, opts.chunk ?? 5000);
  const endRow = Math.max(4, 4 + rows.length);
  const body = sheet.getRange(`A5:${last}${endRow}`);
  body.format.font = { color: C.ink, size: 9 };
  body.format.borders = { insideHorizontal: { style: "thin", color: C.line } };
  applyWidths(sheet, headers, endRow, opts.widths ?? {});
  formatNumeric(sheet, headers, endRow);
  sheet.freezePanes.freezeRows(4);
  sheet.freezePanes.freezeColumns(Math.min(opts.freezeColumns ?? 2, headers.length));
  if (opts.table && rows.length) {
    try { sheet.tables.add(`A4:${last}${endRow}`, true, opts.table); } catch {}
  }
  return { sheet, endRow };
}
function readme(wb, title, notes, reconText) {
  const s = wb.worksheets.add("README");
  s.showGridLines = false;
  s.getRange("A1:H1").merge(); s.getRange("A1").values = [[title]];
  s.getRange("A1:H1").format = { fill: C.navy, font: { bold: true, color: "#FFFFFF", size: 16 }, verticalAlignment: "center" }; s.getRange("A1").format.rowHeight = 30;
  s.getRange("A3").values = [["Method & controls"]]; s.getRange("A3:H3").merge();
  s.getRange("A3:H3").format = { fill: C.gold, font: { bold: true, color: C.ink, size: 11 } };
  const rows = notes.map((x) => [x]);
  if (reconText) rows.push(["Reconciliation: " + reconText]);
  s.getRange(`A4:H${3 + rows.length}`).merge(true); s.getRange(`A4:A${3 + rows.length}`).values = rows;
  s.getRange(`A4:H${3 + rows.length}`).format = { fill: C.ivory, font: { color: C.ink, size: 10 }, wrapText: true, verticalAlignment: "top", borders: { preset: "outside", style: "thin", color: C.line } };
  s.getRange(`A4:H${3 + rows.length}`).format.rowHeight = 30;
  s.getRange("A:A").format.columnWidth = 95;
  s.freezePanes.freezeRows(3);
  return s;
}
function addSubtotals(rows, branchIdx, metricIdxs, yieldIdx, labelIdx, labelPrefix) {
  const out = [];
  let i = 0;
  while (i < rows.length) {
    const branch = rows[i][branchIdx];
    let j = i + 1;
    while (j < rows.length && rows[j][branchIdx] === branch) j++;
    const details = rows.slice(i, j);
    out.push(...details);
    const total = Array(details[0].length).fill("");
    total[branchIdx] = "";
    total[labelIdx] = `${labelPrefix} ${branch}`;
    for (const idx of metricIdxs) total[idx] = details.reduce((a, r) => a + (Number(r[idx]) || 0), 0);
    total[yieldIdx] = total[metricIdxs[1]] && total[metricIdxs[0]] ? total[metricIdxs[1]] / total[metricIdxs[0]] * 100 : 0;
    total[total.length - 1] = "Branch Total";
    out.push(total);
    i = j;
  }
  return out;
}
function addComparisonSubtotals(rows) {
  const out = [];
  let i = 0;
  while (i < rows.length) {
    const branch = rows[i][0]; let j = i + 1; while (j < rows.length && rows[j][0] === branch) j++;
    const details = rows.slice(i, j); out.push(...details);
    const t = Array(details[0].length).fill("");
    t[4] = `BRANCH TOTAL — ${branch}`;
    const sumIdx = [7,8,9,10,12,13,14,15,18,19,20,21];
    for (const idx of sumIdx) t[idx] = details.reduce((a, r) => a + (Number(r[idx]) || 0), 0);
    t[11] = t[7] ? t[8] / t[7] * 100 : 0; t[16] = t[12] ? t[13] / t[12] * 100 : 0; t[22] = t[18] ? t[19] / t[18] * 100 : 0;
    t[23] = t[8] ? t[19] / t[8] - 1 : null; t[24] = t[9] ? t[20] / t[9] - 1 : null; t[25] = t[10] ? t[21] / t[10] - 1 : null; t[26] = t[11] ? t[22] / t[11] - 1 : null; t[27] = "Branch Total";
    out.push(t); i = j;
  }
  return out;
}
function groupSummarySheet(wb, name, title, subtitle, headers, rows, tableName) {
  return baseSheet(wb, name, title, subtitle, headers, rows, { table: tableName, freezeColumns: 1 });
}
async function exportWorkbook(wb, filename, renderSheets) {
  for (const [sheetName, range] of renderSheets) {
    try {
      const blob = await wb.render({ sheetName, range, scale: 1, format: "png" });
      await fs.writeFile(path.join(previewDir, filename.replace(/\.xlsx$/i, "") + "__" + sheetName.replace(/[^A-Za-z0-9]+/g, "_") + ".png"), new Uint8Array(await blob.arrayBuffer()));
    } catch (e) { console.warn("preview failed", filename, sheetName, String(e)); }
  }
  const xlsx = await SpreadsheetFile.exportXlsx(wb);
  await xlsx.save(path.join(outDir, filename));
}

const A = data.schemas.channelActual;
const CClient = data.schemas.clientActual;
const CC = data.schemas.channelComparison;
const CL = data.schemas.clientComparison;
const meta = data.meta;
// Workbook 1: rolling/new period analysis.
if (WORK === "all" || WORK === "new") {
  const channelNew = addSubtotals(data.channels.newDetails, 0, [7,8,9,10], 11, 4, "BRANCH TOTAL —");
  const clientNew = addSubtotals(data.clients.newDetails, 0, [9,10], 11, 2, "BRANCH TOTAL —");
  const wb = Workbook.create();
  readme(wb, "Broking MIS — New Period Analysis", [
    `${meta.newLabel}. This workbook is standalone for the new period; FY 2025-26 is used only as a reference in the companion comparative workbook.`,
    `Channelwise figures show Gross Brokerage, Passout, Net Brokerage and Yield %. Turnover is retained only as a narrow weighted-yield helper column.`,
    `Clientwise rows are collapsed to Client Code + Client Name across channels. Channel Name is the primary trading channel, selected by highest brokerage.`,
    `House accounts ${meta.houseCodes.join(", ")} are flagged in actual client views and excluded from client group summaries and rankings.`,
    `Branch totals follow detail rows so the workbook is ready for Excel outline/grouping.`,
  ], `New-period Channelwise vs Clientwise turnover gap ${meta.reconciliation[meta.newLabel].turnover_gap_pct.toFixed(4)}%; gross brokerage gap ${meta.reconciliation[meta.newLabel].gross_gap_pct.toFixed(4)}%.`);
  baseSheet(wb, "Channelwise Apr-Jul26", `Channelwise — ${meta.newShortLabel}`, `Standalone actuals: ${meta.newLabel}. Branch detail rows precede branch total rows.`, A, channelNew, { table: "NewChannelwise", freezeColumns: 2 });
  baseSheet(wb, "Clientwise Apr-Jul26", `Clientwise — ${meta.newShortLabel}`, `Standalone actuals: ${meta.newLabel}. Channel Name is mapped through Channel ID; house accounts are flagged.`, CClient, clientNew, { table: "NewClientwise", freezeColumns: 2 });
  const gch = [["Group","Channels","Turnover (₹ Crore) [helper]","Gross Brokerage (₹ Lacs)","Passout (₹ Lacs)","Net Brokerage (₹ Lacs)","Yield %"], ...data.channels.groupNew];
  groupSummarySheet(wb, "Group Summary - Channel", `Group Summary — Channel (${meta.newShortLabel})`, `Group totals for ${meta.newLabel}; Turnover is a narrow weighted-yield helper.`, gch[0], gch.slice(1), "NewChannelGroupSummary");
  const gcl = [["Group","Clients","Turnover (₹ Crore)","Gross Brokerage (₹ Lacs)","Yield %"], ...data.clients.groupNew];
  groupSummarySheet(wb, "Group Summary - Client", `Group Summary — Client (${meta.newShortLabel})`, `House accounts excluded from client group summaries.`, gcl[0], gcl.slice(1), "NewClientGroupSummary");
  baseSheet(wb, "Top 100 Clients", `Top 100 Clients — ${meta.newShortLabel}`, `Pre-ranked by Gross Brokerage for ${meta.newLabel}; house accounts excluded.`, CL, data.clients.top100, { table: "NewTop100Clients", freezeColumns: 2 });
  baseSheet(wb, "Top 500 Clients", `Top 500 Clients — ${meta.newShortLabel}`, `Pre-ranked by Gross Brokerage for ${meta.newLabel}; house accounts excluded.`, CL, data.clients.top500, { table: "NewTop500Clients", freezeColumns: 2 });
  const lists = [["House account code","Reason"], ...meta.houseCodes.map((x) => [x, "Marwadi Shares and Finance Ltd. proprietary / house book — exclude from client rankings and group summaries"])];
  baseSheet(wb, "Lists", "Lists — House Account Reference", "Reference list used by the client-level analysis rules.", lists[0], lists.slice(1), { table: "HouseAccounts", freezeColumns: 1 });
  await exportWorkbook(wb, "1_New_Period_Analysis_Apr26-Jul26.xlsx", [["README","A1:H12"],["Channelwise Apr-Jul26","A1:M18"],["Group Summary - Channel","A1:G12"],["Top 100 Clients","A1:U18"]]);
}

// Workbook 2: channel comparative analysis.
if (WORK === "all" || WORK === "channel") {
  const channelFY = addSubtotals(data.channels.fyDetails, 0, [7,8,9,10], 11, 4, "BRANCH TOTAL —");
  const channelNew = addSubtotals(data.channels.newDetails, 0, [7,8,9,10], 11, 4, "BRANCH TOTAL —");
  const wb = Workbook.create();
  readme(wb, "Broking MIS — Comparative Channel Analysis", [
    `Three-layer structure: FY 2025-26 actual, ${meta.newShortLabel} actual, then the comparison. FY stays actual and the new period's additive figures are annualized by ${meta.factor}× (${meta.newLabel} covers 4 months).`,
    `Yield % is a ratio and is not multiplied; annualized Yield % is recomputed from annualized turnover and brokerage.`,
    `Channel comparison includes matched, new, and lost channels. Growth % is blank where the FY denominator is zero.`,
    `Channel-level visible metrics exclude Turnover; narrow helper columns are retained for weighted Yield %.`,
  ], `FY reconciliation gap: turnover ${meta.reconciliation["FY 2025-26"].turnover_gap_pct.toFixed(4)}%; gross brokerage ${meta.reconciliation["FY 2025-26"].gross_gap_pct.toFixed(4)}%. New-period reconciliation is effectively zero.`);
  baseSheet(wb, "FY 2025-26 Actual", "Channelwise — FY 2025-26 Actual", "Layer 1 of 3: FY actual, unscaled. Branch detail rows precede branch total rows.", A, channelFY, { table: "FYChannelActual", freezeColumns: 2 });
  baseSheet(wb, "Apr-Jul 2026 Actual", `Channelwise — ${meta.newShortLabel}`, "Layer 2 of 3: new-period actual, unscaled.", A, channelNew, { table: "NewChannelActual", freezeColumns: 2 });
  const compGrouped = addComparisonSubtotals(data.channels.comparison);
  baseSheet(wb, "Channel Comparison", `Channel Comparison — FY 2025-26 vs ${meta.newShortLabel}`, `Layer 3 of 3: ${meta.newShortLabel} actual × ${meta.factor} annualization factor vs FY actual. FY remains unscaled.`, CC, compGrouped, { table: "ChannelComparison", freezeColumns: 2 });
  const groupHeaders = ["Group","Channels","FY Gross (₹ Lacs)","FY Passout (₹ Lacs)","FY Net (₹ Lacs)","FY Yield %","New Gross Actual (₹ Lacs)","New Passout Actual (₹ Lacs)","New Net Actual (₹ Lacs)","New Yield Actual %","Annualized New Gross (₹ Lacs)","Annualized New Passout (₹ Lacs)","Annualized New Net (₹ Lacs)","Annualized New Yield %","Gross Growth %","Passout Growth %","Net Growth %","Yield Growth %"];
  const gr = [["MSFL"],["MSFL-Sharing"],["Arbitrage"],["P-Sec"]].map((r) => r);
  const groupRows = gr.map(([g], i) => [g, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null]);
  baseSheet(wb, "Group Comparison", `Group Comparison — FY 2025-26 vs ${meta.newShortLabel}`, `SUMIF rollup from Channel Comparison; branch subtotal rows have blank Group values and are naturally skipped.`, groupHeaders, groupRows, { table: "ChannelGroupComparison", freezeColumns: 1 });
  const gs = wb.worksheets.getItem("Group Comparison");
  const compEnd = 4 + compGrouped.length;
  for (let r = 5; r <= 8; r++) {
    gs.getRange(`B${r}`).formulas = [[`=COUNTIF('Channel Comparison'!$D$5:$D$${compEnd},A${r})`]];
    const map = { C: 8, D: 9, E: 10, G: 13, H: 14, I: 15, K: 19, L: 20, M: 21 };
    for (const [letter, idx] of Object.entries(map)) gs.getRange(`${letter}${r}`).formulas = [[`=SUMIF('Channel Comparison'!$D$5:$D$${compEnd},A${r},'Channel Comparison'!$${col(idx)}$5:$${col(idx)}$${compEnd})`]];
    gs.getRange(`F${r}`).formulas = [[`=IF(SUMIF('Channel Comparison'!$D$5:$D$${compEnd},A${r},'Channel Comparison'!$H$5:$H$${compEnd})=0,0,C${r}/SUMIF('Channel Comparison'!$D$5:$D$${compEnd},A${r},'Channel Comparison'!$H$5:$H$${compEnd})*100)`]];
    gs.getRange(`J${r}`).formulas = [[`=IF(SUMIF('Channel Comparison'!$D$5:$D$${compEnd},A${r},'Channel Comparison'!$M$5:$M$${compEnd})=0,0,G${r}/SUMIF('Channel Comparison'!$D$5:$D$${compEnd},A${r},'Channel Comparison'!$M$5:$M$${compEnd})*100)`]];
    gs.getRange(`N${r}`).formulas = [[`=IF(SUMIF('Channel Comparison'!$D$5:$D$${compEnd},A${r},'Channel Comparison'!$S$5:$S$${compEnd})=0,0,K${r}/SUMIF('Channel Comparison'!$D$5:$D$${compEnd},A${r},'Channel Comparison'!$S$5:$S$${compEnd})*100)`]];
    gs.getRange(`O${r}`).formulas = [[`=IF(C${r}=0,0,K${r}/C${r}-1)`]]; gs.getRange(`P${r}`).formulas = [[`=IF(D${r}=0,0,L${r}/D${r}-1)`]]; gs.getRange(`Q${r}`).formulas = [[`=IF(E${r}=0,0,M${r}/E${r}-1)`]]; gs.getRange(`R${r}`).formulas = [[`=IF(F${r}=0,0,N${r}/F${r}-1)`]];
  }
  gs.getRange("F5:F8").format.numberFormat = "0.00"; gs.getRange("J5:J8").format.numberFormat = "0.00"; gs.getRange("N5:N8").format.numberFormat = "0.00"; gs.getRange("O5:R8").format.numberFormat = "0.0%";
  await exportWorkbook(wb, "2_Comparative_Analysis_FY_vs_NewPeriod_Annualized.xlsx", [["README","A1:H10"],["FY 2025-26 Actual","A1:M18"],["Channel Comparison","A1:AB18"],["Group Comparison","A1:R10"]]);
}

// Workbook 3A: actual client layers.
if (WORK === "all" || WORK === "actuals") {
  const clientFY = addSubtotals(data.clients.fyDetails, 0, [9,10], 11, 2, "BRANCH TOTAL —");
  const clientNew = addSubtotals(data.clients.newDetails, 0, [9,10], 11, 2, "BRANCH TOTAL —");
  const wb = Workbook.create();
  readme(wb, "Broking MIS — Client Actuals", [
    `Two standalone actual layers: FY 2025-26 actual and ${meta.newShortLabel} actual. This file keeps the full client universe, including flagged house accounts.`,
    `Client Code + Client Name are collapsed across channels. Channel Name is the primary channel by highest brokerage in the period.`,
    `The companion client comparison file excludes house accounts from matched/new/lost lists and all client rankings.`,
  ], `FY turnover gap ${meta.reconciliation["FY 2025-26"].turnover_gap_pct.toFixed(4)}%; FY gross brokerage gap ${meta.reconciliation["FY 2025-26"].gross_gap_pct.toFixed(4)}%. New-period gap is effectively zero.`);
  baseSheet(wb, "FY 2025-26 Actual", "Clientwise — FY 2025-26 Actual", "Layer 1 of 3: FY client actuals, unscaled; house accounts flagged.", CClient, clientFY, { table: "FYClientActual", freezeColumns: 2, chunk: 3000 });
  baseSheet(wb, "Apr-Jul 2026 Actual", `Clientwise — ${meta.newShortLabel}`, "Layer 2 of 3: new-period client actuals, unscaled; house accounts flagged.", CClient, clientNew, { table: "NewClientActual", freezeColumns: 2, chunk: 3000 });
  await exportWorkbook(wb, "3A_Client_Level_Actuals_FY_vs_NewPeriod.xlsx", [["README","A1:H10"],["FY 2025-26 Actual","A1:M18"],["Apr-Jul 2026 Actual","A1:M18"]]);
}

// Workbook 3B: comparison client layers and rankings.
if (WORK === "all" || WORK === "comp") {
  const wb = Workbook.create();
  readme(wb, "Broking MIS — Client Comparison", [
    `Matched, new, and lost client comparisons use FY 2025-26 actual vs ${meta.newShortLabel} actual × ${meta.factor}. FY is never scaled.`,
    `Client comparison excludes the four house account codes ${meta.houseCodes.join(", ")} from all lists and rankings. The actuals workbook retains them as flagged rows.`,
    `Top 100 / Top 500 lists are pre-ranked by annualized new gross brokerage. Gainers and decliners are ranked by absolute ₹ Lacs change, not percentage growth.`,
    `Yield % is shown as a percentage-point ratio and is not multiplied by the annualization factor.`,
  ], `Matched ${data.clients.matched.length.toLocaleString()} | New ${data.clients.new.length.toLocaleString()} | Lost ${data.clients.lost.length.toLocaleString()} non-house client identities.`);
  const groupedComp = (rows) => addSubtotals(rows, 0, [8,9,11,12,15,16], 10, 2, "BRANCH TOTAL —");
  baseSheet(wb, "Matched Clients", `Matched Clients — FY 2025-26 vs ${meta.newShortLabel}`, `New-period actual × ${meta.factor} annualization factor compared with FY actual.`, CL, groupedComp(data.clients.matched), { table: "MatchedClients", freezeColumns: 2, chunk: 3000 });
  baseSheet(wb, "New Clients", `New Clients — ${meta.newShortLabel}`, `Clients absent in FY 2025-26 and present in the new period; house accounts excluded.`, CL, groupedComp(data.clients.new), { table: "NewClients", freezeColumns: 2, chunk: 3000 });
  baseSheet(wb, "Lost Clients", `Lost Clients — FY 2025-26`, `Clients present in FY 2025-26 and absent in the new period; house accounts excluded.`, CL, groupedComp(data.clients.lost), { table: "LostClients", freezeColumns: 2, chunk: 3000 });
  const all = [...data.clients.matched, ...data.clients.new, ...data.clients.lost].sort((a,b) => (Number(b[16])||0) - (Number(a[16])||0));
  baseSheet(wb, "All Clients", `All Clients — FY 2025-26 vs ${meta.newShortLabel}`, `Flat, pre-ranked comparison universe; house accounts excluded.`, CL, all, { table: "AllClientComparison", freezeColumns: 2, chunk: 3000 });
  baseSheet(wb, "Top 100 Clients", `Top 100 Clients — ${meta.newShortLabel} Annualized`, `Pre-ranked by annualized new gross brokerage; house accounts excluded.`, CL, data.clients.top100, { table: "ClientCompTop100", freezeColumns: 2 });
  baseSheet(wb, "Top 500 Clients", `Top 500 Clients — ${meta.newShortLabel} Annualized`, `Pre-ranked by annualized new gross brokerage; house accounts excluded.`, CL, data.clients.top500, { table: "ClientCompTop500", freezeColumns: 2 });
  baseSheet(wb, "Top 20 Gainers", `Top 20 Gainers — Absolute Gross Change`, `Ranked by absolute ₹ Lacs change: annualized new gross brokerage minus FY gross brokerage.`, CL, data.clients.gainers, { table: "Top20Gainers", freezeColumns: 2 });
  baseSheet(wb, "Top 20 Decliners", `Top 20 Decliners — Absolute Gross Change`, `Ranked by absolute ₹ Lacs change: annualized new gross brokerage minus FY gross brokerage.`, CL, data.clients.decliners, { table: "Top20Decliners", freezeColumns: 2 });
  await exportWorkbook(wb, "3B_Client_Level_Comparison_FY_vs_NewPeriod.xlsx", [["README","A1:H10"],["Matched Clients","A1:U18"],["All Clients","A1:U18"],["Top 20 Gainers","A1:U28"]]);
}

console.log("workbooks exported to", outDir);
