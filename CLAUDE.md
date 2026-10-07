# Pre-MSR workbook project

Monthly Pre-MSR (pre-Management Sales Review) workbook for Syria: `Pre-MSR  Oct.xlsm` (note two spaces in the file name). 60 sheets, about 147k formulas, 51 pivot tables, VBA macros (`Refresh_all`, category filter buttons on Monthly Gap).

## Goals
- Keep a specification of the sheets and how they relate: `spec/Pre-MSR Workbook Spec.html`, also published at https://claude.ai/code/artifact/81a00bea-f61b-41cf-abb7-871dddf0ece9 (update that artifact in place; read it first).
- Recheck the workbook after the user's fixes and keep the Open / Fixed issue lists current.
- Write justifications for management on the black-tab sheets (for example, explain each Gap OL / Monthly Gap movement between M-1 and M by category, F-Group and SKU).

## Management sheets (black tab, theme colour "Text 1")
2026 DF , NNS ,Tons,Cases · Gap OL · Monthly Gap · TTS · IMS Stock · 2027_Comparision · 2027_by Quarter · 2027 OL · 2027 TTS · CYvsNY

## Data flow
- Inputs: Sales Vs Activity-RAW and NNSBreakDown-RAW (month M), the same two sheets with "M-1" suffix (prior month), Fayrouz-Arrival / -2027.
- Masters: Assist (last actual month, exchange rates, period split), SKU mapping-Pricing (SKU master and prices).
- Drivers: FCST and FCST Y+1 (volume in cases, fed by IMS Stock), Price Chnge, DSCR / DSCR Y+1, BDA Allowance Assist (+Y+1), Distribution by Month (+Y+1), Target Split (+Y+1).
- Engines: Consolidation RAW = `Table1` (2026), Y+1 = `Table19` (2027). Almost every report is a pivot or SUMIFS over these.
- IMS Stock is upstream: it drives FCST / FCST Y+1, so edits there flow into everything.
- "M" = this month's Pre-MSR (Oct), "M-1" = last month's (Sep).

## Status at last check (file saved 7 Oct 2026 22:05)
Fixed: external link in Gap OL; 2027-Mosaic / 2027 Allowance pivots now use local Table19; #REF! named ranges deleted; CYvsNY #N/A (empty rows 81-90) now blank.
Open:
- `_FilterDatabase` names on Gap OL and Summary for email point to #REF!.
- #DIV/0! cells: 2027_Comparision 108, IMS Stock 64, 2027 TTS 20, 2026 DF 18, Gap OL 3, 2027 Allowance 241.
- 2027 OL cell U11 header says "2026 FY".
- Slicers Business / Category / SKU Status not connected; Sheet5 looks like scratch.

## How to read the workbook
- Never open or run the macros. Read the file as data only.
- In a cloud / Linux session, use Python: `openpyxl.load_workbook(path, data_only=False)` for formulas, `data_only=True` for cached values, `keep_vba=True` if writing. Tab colours: `ws.sheet_properties.tabColor`.
- On the user's Windows laptop there is no Python: copy the xlsm to a .zip, `Expand-Archive`, then run `spec/recheck.ps1 <xlsm> <unzip-dir>`, `spec/peek.ps1`, `spec/deps.ps1`.
- `spec/pre-msr-review.html` is a browser page for the black-tab sheets: open it, pick the xlsm, and each management sheet shows as a tab with the workbook's table, autofilter and pivot-label filters, a Category quick filter, frozen panes and hidden rows kept. It reads the file locally (SheetJS + JSZip from cdnjs), never runs macros, and "Save report" writes a standalone HTML copy of the view. Report areas per sheet are set in `MAIN_COLS` at the top of its script; update them if a sheet's layout moves.
- `spec/baseline-2026-10-07-2205.txt` is the recheck output of the last checked version. Diff the new output against it to find what changed, then refresh the baseline.

## Working rules
- Do not modify the .xlsm unless the user asks. Put outputs (spec updates, justification notes) in `spec/` or a new markdown file.
- Workbook data is internal company data. Do not send it to any outside service.
