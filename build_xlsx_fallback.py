import json
import math
import sys
from pathlib import Path

import xlsxwriter

ROOT = Path(__file__).resolve().parent
OUT = ROOT / "outputs"
OUT.mkdir(exist_ok=True)
DATA = json.loads((ROOT / "build" / "data.json").read_text(encoding="utf-8"))
META = DATA["meta"]
S = DATA["schemas"]

NAVY = "#10253F"; NAVY2 = "#173A5E"; IVORY = "#FAF7F0"; PAPER = "#F3EFE6"; GOLD = "#C9A227"; INK = "#1C2833"; MUTED = "#617182"; LINE = "#D9D4C8"

def add_formats(wb):
    return {
        "title": wb.add_format({"bold": True, "font_color": "white", "bg_color": NAVY, "font_size": 15, "align": "left", "valign": "vcenter"}),
        "subtitle": wb.add_format({"italic": True, "font_color": MUTED, "bg_color": PAPER, "text_wrap": True, "valign": "vcenter"}),
        "section": wb.add_format({"bold": True, "font_color": INK, "bg_color": GOLD}),
        "note": wb.add_format({"font_color": INK, "bg_color": IVORY, "text_wrap": True, "valign": "top", "border": 1, "border_color": LINE}),
        "header": wb.add_format({"bold": True, "font_color": "white", "bg_color": NAVY2, "text_wrap": True, "align": "center", "valign": "vcenter", "border": 1, "border_color": NAVY2}),
        "text": wb.add_format({"font_color": INK, "font_size": 9}),
        "text_total": wb.add_format({"bold": True, "font_color": INK, "bg_color": "#EDE7D9", "top": 1}),
        "num": wb.add_format({"font_color": INK, "font_size": 9, "num_format": "#,##0.00"}),
        "num_total": wb.add_format({"bold": True, "font_color": INK, "bg_color": "#EDE7D9", "num_format": "#,##0.00", "top": 1}),
        "pct": wb.add_format({"font_color": INK, "font_size": 9, "num_format": "0.00"}),
        "pct_total": wb.add_format({"bold": True, "font_color": INK, "bg_color": "#EDE7D9", "num_format": "0.00", "top": 1}),
        "growth": wb.add_format({"font_color": INK, "font_size": 9, "num_format": "0.0%"}),
        "growth_total": wb.add_format({"bold": True, "font_color": INK, "bg_color": "#EDE7D9", "num_format": "0.0%", "top": 1}),
        "formula_num": wb.add_format({"font_color": INK, "font_size": 9, "num_format": "#,##0.00"}),
        "formula_pct": wb.add_format({"font_color": INK, "font_size": 9, "num_format": "0.00"}),
        "formula_growth": wb.add_format({"font_color": INK, "font_size": 9, "num_format": "0.0%"}),
    }

def col_letter(i):
    s = ""
    while i >= 0:
        s = chr(65 + i % 26) + s; i = i // 26 - 1
    return s

def widths(headers):
    d = {"Branch Code": 12, "Channel ID": 12, "Channel Type": 12, "Group": 16, "Channel Name": 34, "City": 18, "State": 16, "Client Code": 16, "Client Name": 32, "Is House Account": 14, "Row Type": 14, "Status": 12}
    return [0.2 if "[helper]" in h else d.get(h, 15) for h in headers]

def num_kind(header):
    if "Growth %" in header: return "growth"
    if "Yield %" in header: return "pct"
    if "Clients" == header or "Channels" == header: return "int"
    if any(x in header for x in ["Turnover", "Gross", "Passout", "Net", "Change"]): return "num"
    if header == "Annualization Factor": return "int"
    return "text"

def write_readme(wb, f, title, notes, recon):
    ws = wb.add_worksheet("README"); ws.hide_gridlines(2); ws.set_column("A:A", 95)
    ws.merge_range("A1:H1", title, f["title"]); ws.set_row(0, 30)
    ws.merge_range("A3:H3", "Method & controls", f["section"])
    all_notes = list(notes) + (["Reconciliation: " + recon] if recon else [])
    for i, note in enumerate(all_notes, start=3):
        ws.merge_range(i, 0, i, 7, note, f["note"]); ws.set_row(i, 30)
    ws.freeze_panes(3, 0)
    return ws

def cell_fmt(f, header, total=False, formula=False):
    kind = num_kind(header)
    if kind == "growth": return f["growth_total" if total else ("formula_growth" if formula else "growth")]
    if kind == "pct": return f["pct_total" if total else ("formula_pct" if formula else "pct")]
    if kind == "num": return f["num_total" if total else ("formula_num" if formula else "num")]
    return f["text_total" if total else "text"]

def write_data_sheet(wb, f, name, title, subtitle, headers, rows, table_name=None, freeze_cols=2, helper_names=True):
    ws = wb.add_worksheet(name[:31]); ws.hide_gridlines(2); ws.set_row(0, 28); ws.set_row(1, 28); ws.set_row(3, 34)
    last = len(headers) - 1
    ws.merge_range(0, 0, 0, last, title, f["title"])
    ws.merge_range(1, 0, 1, last, subtitle, f["subtitle"])
    for c, h in enumerate(headers): ws.write(3, c, h, f["header"])
    widths_ = widths(headers)
    for c, w in enumerate(widths_):
        opts = {"hidden": True} if "[helper]" in headers[c] else {}
        ws.set_column(c, c, w, None, opts)
    for r_i, row in enumerate(rows, start=4):
        total = row[-1] in ("Branch Total",)
        for c, v in enumerate(row):
            header = headers[c]
            if v is None: ws.write_blank(r_i, c, None, cell_fmt(f, header, total))
            elif isinstance(v, bool): ws.write_boolean(r_i, c, v, cell_fmt(f, header, total))
            elif isinstance(v, (int, float)) and not isinstance(v, bool): ws.write_number(r_i, c, float(v), cell_fmt(f, header, total))
            else: ws.write(r_i, c, v, cell_fmt(f, header, total))
        if not total: ws.set_row(r_i, None, None, {"level": 1})
    end = max(3, 4 + len(rows) - 1)
    ws.freeze_panes(4, freeze_cols); ws.autofilter(3, 0, end, last); ws.outline_settings(True, True, True, False)
    return ws

def add_subtotals(rows, branch_idx, metric_idxs, yield_idx, label_idx, label_prefix):
    out=[]; i=0
    while i < len(rows):
        b=rows[i][branch_idx]; j=i+1
        while j < len(rows) and rows[j][branch_idx] == b: j+=1
        details=rows[i:j]; out.extend(details); t=["" for _ in details[0]]; t[label_idx]=f"{label_prefix} {b}"
        for idx in metric_idxs: t[idx]=sum(float(r[idx] or 0) for r in details)
        t[yield_idx] = t[metric_idxs[1]] / t[metric_idxs[0]] * 100 if t[metric_idxs[0]] else 0
        t[-1]="Branch Total"; out.append(t); i=j
    return out

def add_comp_subtotals(rows):
    out=[]; i=0
    while i < len(rows):
        b=rows[i][0]; j=i+1
        while j < len(rows) and rows[j][0] == b: j+=1
        details=rows[i:j]; out.extend(details); t=["" for _ in details[0]]; t[4]=f"BRANCH TOTAL — {b}"
        for idx in [7,8,9,10,12,13,14,15,18,19,20,21]: t[idx]=sum(float(r[idx] or 0) for r in details)
        t[11]=t[8]/t[7]*100 if t[7] else 0; t[16]=t[13]/t[12]*100 if t[12] else 0; t[22]=t[19]/t[18]*100 if t[18] else 0
        t[23]=t[19]/t[8]-1 if t[8] else None; t[24]=t[20]/t[9]-1 if t[9] else None; t[25]=t[21]/t[10]-1 if t[10] else None; t[26]=t[22]/t[11]-1 if t[11] else None; t[27]="Branch Total"
        out.append(t); i=j
    return out

def source_recon(label):
    r=META["reconciliation"][label]
    return f"{label}: turnover gap {r['turnover_gap_pct']:.4f}%; gross brokerage gap {r['gross_gap_pct']:.4f}%"

def build_new():
    fn=OUT/"1_New_Period_Analysis_Apr26-Jul26.xlsx"; wb=xlsxwriter.Workbook(fn, {"constant_memory": True}); f=add_formats(wb)
    write_readme(wb,f,"Broking MIS — New Period Analysis",[
        f"{META['newLabel']}. Standalone new-period actuals; FY 2025-26 is used only in the companion comparative workbook.",
        "Channelwise visible metrics show Gross Brokerage, Passout, Net Brokerage and Yield %. Turnover is hidden as a weighted-yield helper.",
        "Clientwise is collapsed to Client Code + Client Name; Channel Name is the primary trading channel selected by highest brokerage.",
        f"House accounts {', '.join(META['houseCodes'])} are flagged in actuals and excluded from client group summaries and rankings.",
        "Detail rows are followed by branch total rows and are assigned native Excel outline level 1.",
    ], source_recon(META["newLabel"]))
    a=S["channelActual"]; c=S["clientActual"]; cc=S["clientComparison"]
    write_data_sheet(wb,f,"Channelwise Apr-Jul26",f"Channelwise — {META['newShortLabel']}",f"Standalone actuals: {META['newLabel']}.",a,add_subtotals(DATA["channels"]["newDetails"],0,[7,8,9,10],11,4,"BRANCH TOTAL —"),"NewChannelwise")
    write_data_sheet(wb,f,"Clientwise Apr-Jul26",f"Clientwise — {META['newShortLabel']}",f"Standalone actuals: {META['newLabel']}; Channel Name is mapped through Channel ID.",c,add_subtotals(DATA["clients"]["newDetails"],0,[9,10],11,2,"BRANCH TOTAL —"),"NewClientwise",2)
    gch=[["Group","Channels","Turnover (₹ Crore) [helper]","Gross Brokerage (₹ Lacs)","Passout (₹ Lacs)","Net Brokerage (₹ Lacs)","Yield %"]]+DATA["channels"]["groupNew"]
    write_data_sheet(wb,f,"Group Summary - Channel",f"Group Summary — Channel ({META['newShortLabel']})","Group totals; Turnover is a hidden weighted-yield helper.",gch[0],gch[1:],"NewChannelGroupSummary",1)
    gcl=[["Group","Clients","Turnover (₹ Crore)","Gross Brokerage (₹ Lacs)","Yield %"]]+DATA["clients"]["groupNew"]
    write_data_sheet(wb,f,"Group Summary - Client",f"Group Summary — Client ({META['newShortLabel']})","House accounts excluded from client group summaries.",gcl[0],gcl[1:],"NewClientGroupSummary",1)
    write_data_sheet(wb,f,"Top 100 Clients",f"Top 100 Clients — {META['newShortLabel']}",f"Pre-ranked by Gross Brokerage for {META['newLabel']}; house accounts excluded.",cc,DATA["clients"]["top100"],"NewTop100Clients")
    write_data_sheet(wb,f,"Top 500 Clients",f"Top 500 Clients — {META['newShortLabel']}",f"Pre-ranked by Gross Brokerage for {META['newLabel']}; house accounts excluded.",cc,DATA["clients"]["top500"],"NewTop500Clients")
    lists=[["House account code","Reason"]]+[[x,"Marwadi Shares and Finance Ltd. proprietary / house book — exclude from client rankings and group summaries"] for x in META["houseCodes"]]
    write_data_sheet(wb,f,"Lists","Lists — House Account Reference","Reference list used by the client-level analysis rules.",lists[0],lists[1:],"HouseAccounts",1)
    wb.close(); print(fn)

def build_channel():
    fn=OUT/"2_Comparative_Analysis_FY_vs_NewPeriod_Annualized.xlsx"; wb=xlsxwriter.Workbook(fn); f=add_formats(wb)
    write_readme(wb,f,"Broking MIS — Comparative Channel Analysis",[
        f"Three-layer structure: FY 2025-26 actual, {META['newShortLabel']} actual, then comparison.",
        f"FY stays actual and additive new-period figures are annualized by {META['factor']}×; {META['newLabel']} covers 4 months.",
        "Yield % is a ratio and is not multiplied; annualized Yield % is recomputed from annualized turnover and brokerage.",
        "Channel comparison includes matched, new, and lost channels; growth is blank where the FY denominator is zero.",
    ], source_recon("FY 2025-26")+"; "+source_recon(META["newLabel"]))
    a=S["channelActual"]; comp=S["channelComparison"]
    write_data_sheet(wb,f,"FY 2025-26 Actual","Channelwise — FY 2025-26 Actual","Layer 1 of 3: FY actual, unscaled.",a,add_subtotals(DATA["channels"]["fyDetails"],0,[7,8,9,10],11,4,"BRANCH TOTAL —"),"FYChannelActual")
    write_data_sheet(wb,f,"Apr-Jul 2026 Actual",f"Channelwise — {META['newShortLabel']}","Layer 2 of 3: new-period actual, unscaled.",a,add_subtotals(DATA["channels"]["newDetails"],0,[7,8,9,10],11,4,"BRANCH TOTAL —"),"NewChannelActual")
    write_data_sheet(wb,f,"Channel Comparison",f"Channel Comparison — FY 2025-26 vs {META['newShortLabel']}",f"Layer 3 of 3: new-period actual × {META['factor']} annualization factor vs FY actual.",comp,add_comp_subtotals(DATA["channels"]["comparison"]),"ChannelComparison")
    gh=["Group","Channels","FY Gross (₹ Lacs)","FY Passout (₹ Lacs)","FY Net (₹ Lacs)","FY Yield %","New Gross Actual (₹ Lacs)","New Passout Actual (₹ Lacs)","New Net Actual (₹ Lacs)","New Yield Actual %","Annualized New Gross (₹ Lacs)","Annualized New Passout (₹ Lacs)","Annualized New Net (₹ Lacs)","Annualized New Yield %","Gross Growth %","Passout Growth %","Net Growth %","Yield Growth %"]
    groups=META["groups"]; comp_rows=DATA["channels"]["comparison"]; end=4+len(add_comp_subtotals(comp_rows));
    vals={g:{"channels":0,"fy_gross":0,"fy_pass":0,"fy_net":0,"fy_turn":0,"new_gross":0,"new_pass":0,"new_net":0,"new_turn":0,"ann_gross":0,"ann_pass":0,"ann_net":0,"ann_turn":0} for g in groups}
    for r in comp_rows:
        g=r[3]
        if g not in vals: continue
        v=vals[g]; v["channels"]+=1; v["fy_turn"]+=r[7]; v["fy_gross"]+=r[8]; v["fy_pass"]+=r[9]; v["fy_net"]+=r[10]; v["new_turn"]+=r[12]; v["new_gross"]+=r[13]; v["new_pass"]+=r[14]; v["new_net"]+=r[15]; v["ann_turn"]+=r[18]; v["ann_gross"]+=r[19]; v["ann_pass"]+=r[20]; v["ann_net"]+=r[21]
    gr=[]
    for g in groups:
        v=vals[g]; fy_y=v["fy_gross"]/v["fy_turn"]*100 if v["fy_turn"] else 0; new_y=v["new_gross"]/v["new_turn"]*100 if v["new_turn"] else 0; ann_y=v["ann_gross"]/v["ann_turn"]*100 if v["ann_turn"] else 0
        gr.append([g,v["channels"],v["fy_gross"],v["fy_pass"],v["fy_net"],fy_y,v["new_gross"],v["new_pass"],v["new_net"],new_y,v["ann_gross"],v["ann_pass"],v["ann_net"],ann_y,v["ann_gross"]/v["fy_gross"]-1 if v["fy_gross"] else None,v["ann_pass"]/v["fy_pass"]-1 if v["fy_pass"] else None,v["ann_net"]/v["fy_net"]-1 if v["fy_net"] else None,ann_y/fy_y-1 if fy_y else None])
    ws=write_data_sheet(wb,f,"Group Comparison",f"Group Comparison — FY 2025-26 vs {META['newShortLabel']}","SUMIF rollup from Channel Comparison; branch subtotal rows are excluded by blank Group.",gh,gr,"ChannelGroupComparison",1)
    # Keep the small rollup formula-driven and include cached values for immediate display.
    comp_sheet_end = 4 + len(add_comp_subtotals(comp_rows))
    for i, g in enumerate(groups, start=4):
        excel_row = i + 1; v = vals[g]; fy_y=v["fy_gross"]/v["fy_turn"]*100 if v["fy_turn"] else 0; new_y=v["new_gross"]/v["new_turn"]*100 if v["new_turn"] else 0; ann_y=v["ann_gross"]/v["ann_turn"]*100 if v["ann_turn"] else 0
        ws.write_formula(i, 1, f"=COUNTIF('Channel Comparison'!$D$5:$D${comp_sheet_end},A{excel_row})", cell_fmt(f, gh[1], formula=True), v["channels"])
        sum_map = {2:(8,v["fy_gross"]),3:(9,v["fy_pass"]),4:(10,v["fy_net"]),6:(13,v["new_gross"]),7:(14,v["new_pass"]),8:(15,v["new_net"]),10:(19,v["ann_gross"]),11:(20,v["ann_pass"]),12:(21,v["ann_net"])}
        for c,(src_idx,cached) in sum_map.items():
            src_col=col_letter(src_idx); ws.write_formula(i,c,f"=SUMIF('Channel Comparison'!$D$5:$D${comp_sheet_end},A{excel_row},'Channel Comparison'!${src_col}$5:${src_col}${comp_sheet_end})",cell_fmt(f,gh[c],formula=True),cached)
        formulas={5:(f"=IF(SUMIF('Channel Comparison'!$D$5:$D${comp_sheet_end},A{excel_row},'Channel Comparison'!$H$5:$H${comp_sheet_end})=0,0,C{excel_row}/SUMIF('Channel Comparison'!$D$5:$D${comp_sheet_end},A{excel_row},'Channel Comparison'!$H$5:$H${comp_sheet_end})*100)",fy_y),9:(f"=IF(SUMIF('Channel Comparison'!$D$5:$D${comp_sheet_end},A{excel_row},'Channel Comparison'!$M$5:$M${comp_sheet_end})=0,0,G{excel_row}/SUMIF('Channel Comparison'!$D$5:$D${comp_sheet_end},A{excel_row},'Channel Comparison'!$M$5:$M${comp_sheet_end})*100)",new_y),13:(f"=IF(SUMIF('Channel Comparison'!$D$5:$D${comp_sheet_end},A{excel_row},'Channel Comparison'!$S$5:$S${comp_sheet_end})=0,0,K{excel_row}/SUMIF('Channel Comparison'!$D$5:$D${comp_sheet_end},A{excel_row},'Channel Comparison'!$S$5:$S${comp_sheet_end})*100)",ann_y),14:(f"=IF(C{excel_row}=0,0,K{excel_row}/C{excel_row}-1)",gr[i-4][14]),15:(f"=IF(D{excel_row}=0,0,L{excel_row}/D{excel_row}-1)",gr[i-4][15]),16:(f"=IF(E{excel_row}=0,0,M{excel_row}/E{excel_row}-1)",gr[i-4][16]),17:(f"=IF(F{excel_row}=0,0,N{excel_row}/F{excel_row}-1)",gr[i-4][17])}
        for c,(formula,cached) in formulas.items(): ws.write_formula(i,c,formula,cell_fmt(f,gh[c],formula=True),cached if cached is not None else "")
    wb.close(); print(fn)

def build_client_actuals():
    fn=OUT/"3A_Client_Level_Actuals_FY_vs_NewPeriod.xlsx"; wb=xlsxwriter.Workbook(fn, {"constant_memory": True}); f=add_formats(wb)
    write_readme(wb,f,"Broking MIS — Client Actuals",[
        f"Two standalone actual layers: FY 2025-26 actual and {META['newShortLabel']} actual. Full client universe is retained, including flagged house accounts.",
        "Client Code + Client Name are collapsed across channels; Channel Name is primary by highest brokerage in the period.",
        f"The companion comparison file excludes house accounts {', '.join(META['houseCodes'])} from matched/new/lost lists and rankings.",
    ], source_recon("FY 2025-26")+"; "+source_recon(META["newLabel"]))
    c=S["clientActual"]
    write_data_sheet(wb,f,"FY 2025-26 Actual","Clientwise — FY 2025-26 Actual","Layer 1 of 3: FY client actuals, unscaled; house accounts flagged.",c,add_subtotals(DATA["clients"]["fyDetails"],0,[9,10],11,2,"BRANCH TOTAL —"),"FYClientActual",2)
    write_data_sheet(wb,f,"Apr-Jul 2026 Actual",f"Clientwise — {META['newShortLabel']}","Layer 2 of 3: new-period client actuals, unscaled; house accounts flagged.",c,add_subtotals(DATA["clients"]["newDetails"],0,[9,10],11,2,"BRANCH TOTAL —"),"NewClientActual",2)
    wb.close(); print(fn)

def build_client_comp():
    fn=OUT/"3B_Client_Level_Comparison_FY_vs_NewPeriod.xlsx"; wb=xlsxwriter.Workbook(fn, {"constant_memory": True}); f=add_formats(wb)
    write_readme(wb,f,"Broking MIS — Client Comparison",[
        f"Matched, new, and lost client comparisons use FY 2025-26 actual vs {META['newShortLabel']} actual × {META['factor']}. FY is never scaled.",
        f"Client comparison excludes house accounts {', '.join(META['houseCodes'])} from all lists and rankings; the actuals workbook retains them as flagged rows.",
        "Top 100 / Top 500 are pre-ranked by annualized new gross brokerage. Gainers and decliners are ranked by absolute ₹ Lacs change.",
        "Yield % is a percentage-point ratio and is not multiplied by the annualization factor.",
    ], f"Matched {len(DATA['clients']['matched']):,} | New {len(DATA['clients']['new']):,} | Lost {len(DATA['clients']['lost']):,} non-house client identities.")
    c=S["clientComparison"]
    for name, title, subtitle, key in [
        ("Matched Clients",f"Matched Clients — FY 2025-26 vs {META['newShortLabel']}",f"New-period actual × {META['factor']} annualization factor compared with FY actual.","matched"),
        ("New Clients",f"New Clients — {META['newShortLabel']}","Clients absent in FY 2025-26 and present in the new period; house accounts excluded.","new"),
        ("Lost Clients","Lost Clients — FY 2025-26","Clients present in FY 2025-26 and absent in the new period; house accounts excluded.","lost"),
    ]:
        write_data_sheet(wb,f,name,title,subtitle,c,add_subtotals(DATA["clients"][key],0,[8,9,11,12,15,16],10,2,"BRANCH TOTAL —"),key.title().replace(" ","")+"Clients",2)
    all_rows=sorted(DATA["clients"]["matched"]+DATA["clients"]["new"]+DATA["clients"]["lost"],key=lambda r: float(r[16] or 0),reverse=True)
    write_data_sheet(wb,f,"All Clients",f"All Clients — FY 2025-26 vs {META['newShortLabel']}","Flat, pre-ranked comparison universe; house accounts excluded.",c,all_rows,"AllClientComparison",2)
    write_data_sheet(wb,f,"Top 100 Clients",f"Top 100 Clients — {META['newShortLabel']} Annualized","Pre-ranked by annualized new gross brokerage; house accounts excluded.",c,DATA["clients"]["top100"],"ClientCompTop100",2)
    write_data_sheet(wb,f,"Top 500 Clients",f"Top 500 Clients — {META['newShortLabel']} Annualized","Pre-ranked by annualized new gross brokerage; house accounts excluded.",c,DATA["clients"]["top500"],"ClientCompTop500",2)
    write_data_sheet(wb,f,"Top 20 Gainers","Top 20 Gainers — Absolute Gross Change","Ranked by absolute ₹ Lacs change: annualized new gross brokerage minus FY gross brokerage.",c,DATA["clients"]["gainers"],"Top20Gainers",2)
    write_data_sheet(wb,f,"Top 20 Decliners","Top 20 Decliners — Absolute Gross Change","Ranked by absolute ₹ Lacs change: annualized new gross brokerage minus FY gross brokerage.",c,DATA["clients"]["decliners"],"Top20Decliners",2)
    wb.close(); print(fn)

if __name__ == "__main__":
    mode = sys.argv[1] if len(sys.argv) > 1 else "all"
    if mode in ("all", "new"): build_new()
    if mode in ("all", "channel"): build_channel()
    if mode in ("all", "actuals"): build_client_actuals()
    if mode in ("all", "comp"): build_client_comp()
