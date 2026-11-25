# Adding a new SKU to the Pre-MSR workbook

Checked against `Pre-MSR  Oct.xlsm` (saved 7 Oct 2026). The reference was the last SKU added by hand, NIDO 1+ GUM 12x900g N4 XA (12625767). This guide traces every place that SKU appears.

## In short

- A SKU appears in about 40 sheets. Twenty-five of them need rows added. The rest are pivots that update on **Refresh_all**, or system exports that fill in by themselves.
- `spec/NewSKU.bas` is a VBA module. It builds a **New SKU** sheet that holds every variable a SKU needs. It then copies an existing "template" SKU's rows in each of those 25 sheets and writes in your values.
- The New SKU sheet has a live checklist of 35 checks. It also works for SKUs added by hand.
- Run **Widen lookup ranges** once before the first new SKU. Several lookups stop before the end of their lists already (see [Link problems](#link-problems-found)).

The macro was written from the workbook's structure. It was not run in Excel here, because this environment has no Excel. Use it on a copy first. Run **Preview**, which changes nothing, before **Add SKU**.

## Updating from an earlier version

Checked against `Pre-MSR  Nov - BSE1-Enhance -Add Kitkat-Monaf 2.xlsm` (November). That file added the sheets below, and the macro now handles them:
- **Commission by Month**, which Table1's Commission% now reads;
- **2027 M vs M-1** with `Table12`;
- an M-1 half in Table19.

It also deleted Ramadan.

`Pre-MSR  Nov - BSE2-Mosaic Submission.xlsm` (version 2026-11-22 d) brought:
- a new Status value, **IR**;
- TTS `Table4` now reads its pivot list (AA:AB) by position, as Monthly Gap `Table612` does (CD:CE).

Those two tables now follow their pivots after every refresh (see [Report tables that follow a pivot](#report-tables-that-follow-a-pivot)).

1. In the VBA editor, remove the old NewSKU module and import the new `spec/NewSKU.bas`. Check the `' VERSION` line at the top, then run Debug > Compile VBAProject.
2. Run `Setup_RefreshNewSKUSheet` (Alt+F8). It rebuilds the checklist (rows 56–93) for the current sheets, updates the Status note, adds the **Match report tables to pivots** button (O1:P2) and adds the Delete / Change boxes if they are missing. What you typed in the form stays.
3. Click **Match report tables to pivots** once (or run `NewSKU_SyncReportTables`) to repair the tables as they are today.

## Pivot room (why a refresh can fail)

Every pivot on Table1 or Table19 that lists SKUs grows by one row for each new SKU. If the row under a pivot holds anything, Excel cannot grow it. Excel then cancels the refresh of **every** pivot on that table, without an error the macro can catch. Examples are a title like "Case" or the filter cells of the next pivot.

The November test add (code 2, NIDO Fiber 700g) hit exactly this:
- On **2026 DF , NNS ,Tons,Cases**, PivotTable2 (A73:U118) ends right above the "Case" label in A119.
- The Table1 pivots (2026 DF, Monthly Gap, Data Check, Gap OL, TTS ...) stayed at the old data, while the Table19 pivots refreshed.
- PivotTable1 (A17:U61) has one free row before the "Tons" label in A63.

Since version 2026-11-26 f:
- **Preview** lists every pivot with fewer than 2 free rows below it ("Pivot room" lines).
- After the refresh, **Add / Delete / Change** check that each Table1 / Table19 pivot cache holds all the table's rows. If not, they report "Pivots NOT refreshed" as a problem and skip the report-table matching. Matching against out-of-date pivots would give wrong rows.

Fix: give each pivot its own room. In the November BSE2 file the Tons and Cases pivots were copied to their own sheets, **2026 DF ,Tons** (PivotTable2, A21) and **2026 DF ,Cases** (PivotTable5, A17).

The original sheet still holds all three pivots, and its Tons pivot is still blocked by "Case" in A119, so the refresh still fails. Two ways to fix it:
- Delete the Tons and Cases pivots (rows 63–171) from **2026 DF , NNS ,Tons,Cases**, keeping only the NNS pivot.
- Or insert 10–20 empty rows above "Tons" and above "Case".

If you delete them, repoint the NNS-per-case cells in column W from the old Cases pivot (`=S21/S130`) to the new sheet: `=S21/'2026 DF ,Cases'!S21` (same row number, because the new Cases pivot starts 109 rows higher).

Then click Refresh All and **Match report tables to pivots**. The checklist line "2026 DF reports" counts the SKU on every sheet whose name starts with "2026 DF", and expects one row per pivot on those sheets.

## Report tables that follow a pivot

Monthly Gap `Table612` and TTS `Table4` read a pivot's row list by position: table row 1 reads pivot row 1, and so on. They need exactly one table row per pivot row. The pivot lists also keep SKUs that only have last month's (M-1) rows, so the count cannot be worked out from the SKUs added or deleted.

After the pivots refresh, Add SKU and Delete SKU (and the **Match report tables to pivots** button) therefore:
1. find the table's link columns (first row is a plain link such as `=CD35`);
2. count the pivot list down to the first blank, "(blank)" or "Grand Total";
3. add or remove table rows to match;
4. set every link to the first row's pattern, so row n reads list row n.

The preview shows the counts without changing anything.

## Install (once)

1. Save a copy of the workbook.
2. Press Alt+F11, then choose File > Import File and pick `spec/NewSKU.bas`.
3. Press Alt+F8 and run `Setup_NewSKU_Sheet`. A yellow-tab sheet **New SKU** appears after "Guide line for new SKU update".
4. Press Alt+F8 and run `NewSKU_WidenLookupRanges` (only once). It saves a backup, then lists every formula it changed in column L of the New SKU sheet.
5. Save as .xlsm.

## Add a SKU

1. On **New SKU**, fill the yellow cells. Blank cells take the template SKU's value, which is shown in grey. The "Will use" column shows what will be written.
2. Click **Preview**. Column L lists every sheet and row it would add, plus any problems. Nothing changes.
3. Click **Add SKU**. A backup copy is saved next to the workbook (`... - before SKU 12345678 2026-10-19 1030.xlsm`). Then the rows are added and the pivots are refreshed.
4. Read the checklist under the inputs. Every line should say **OK**.
5. Actual sales arrive by themselves: once the SKU sells, Sales Vs Activity-RAW and NNSBreakDown-RAW carry it, and Table1 picks it up by code.

Choose a template SKU in the **same F-Group** as the new one, for example the pack it replaces. Its rows give the new SKU its price-change months, distribution months, DSCR months and report formulas. Its trade terms and customer split are used unless you type others.

## Delete a SKU

The reverse of Add SKU: it removes the SKU's rows from every sheet Add SKU writes to.

1. If your New SKU sheet was built before this feature, run `Setup_DeleteSKU_Section` once (Alt+F8). It adds a **Delete a SKU** box under the checklist and keeps everything you typed. A new `Setup_NewSKU_Sheet` includes the box already.
2. Type the code in **SKU code to delete**. The box shows its description and status, so you can check it is the right SKU.
3. Click **Preview delete**. Column L lists every row it would remove, sheet by sheet. Nothing changes. It also warns if the SKU already has actual sales in Sales Vs Activity-RAW, because those actuals disappear from the reports with it.
4. Click **Delete SKU**. A backup copy is saved first (`... - before deleting SKU 12345678 ....xlsm`). Then the rows are removed and the pivots are refreshed.
5. Type the code in **New SKU code** (C5). Every "Found" in the checklist should now be 0.

What it removes:
- the master row;
- the 2026 and 2027 Price Chnge rows;
- BDA Allowance Assist and Distribution by Month (+Y+1);
- FCST and FCST Y+1;
- the 8-row IMS Stock block;
- the four Fayrouz rows;
- DSCR, DSCR -LS, DSCR Y+1 and Target Split (+Y+1), 2 rows each;
- the Table1 and Table19 rows;
- the 2027_Comparision month block and the 2027_by Quarter helper rows;
- the Commission by Month row (same row number as Distribution by Month);
- the report rows in 2027_Comparision and 2027 VS ICP;
- the 2027 M vs M-1 row for its short description and the Gap OL row for its report F-Group, but only when no other SKU still uses them (the preview says which);
- (after the refresh) Monthly Gap and TTS tables are matched to their pivot lists.

The order keeps the row-by-row links intact:
- Distribution by Month and Commission by Month go before BDA, and BDA before SKU mapping-Pricing.
- The 2027 price row goes before its 2026 row.
- 2027_by Quarter goes before the 2027_Comparision block.

Excel then moves the remaining links with the rows that shift up. Three sheets need care because of what sits beside the SKU's rows:
- **SKU mapping-Pricing:** only the SKU's own cells move (A:V, Z:AH). The unified-name list in W:Y and the category list in S2:S6 stay put.
- **2027_Comparision and 2027_by Quarter:** only the block columns move (V:AH, AD:AN), because pivots sit beside them.
- **CYvsNY:** its table keeps its rows. The spare pair stays blank.

Do not delete a SKU that other rows were copied from while you still need it as a template.

## Change a SKU's code or description

Use this when a SKU already in the workbook gets a new code or a new name. It replaces the old code and description in every sheet, so nothing has to be retyped.

1. If your New SKU sheet was built before this feature, run `Setup_ChangeSKU_Section` once (Alt+F8). It adds a **Change a SKU** box under the delete box and keeps everything you typed. A new `Setup_NewSKU_Sheet` includes the box already (row 100).
2. Type the code as it is today in **SKU code now**. The "Now" column shows its current description, short description and 1C description.
3. Fill only what changes:
   - **New SKU code.** "Now" says *free* or *already used by another SKU*.
   - **New SKU description**, the full name (SKU mapping-Pricing column C).
   - **New short description**, SKU mapping-Pricing column E. Table1 / Table19 `Unifed` read it. The 2027 M vs M-1 row is renamed too when no other SKU shares the old short description; otherwise the preview says to add a row by hand.
   - **New 1C description**, SKU mapping-Pricing column Z.
4. Two options:
   - **Rename unified names too** (default *No*). If the old description is also used as a unified name (SKU mapping A and W:X, Table1 / Table19 SKUDESC and Unifed, FCST Unifed), those cells keep the old text, because a unified name can group several SKUs. The preview lists them. Choose *Yes* to rename them as well.
   - **Change the pasted data sheets too** (default *Yes*). This covers Sales Vs Activity-RAW, NNSBreakDown-RAW, their M-1 copies, Sales vs Activity 2027 and From Target Split. Table1 reads actuals by product code, so with *No* the past actuals under the old code drop out of the reports.
5. Click **Preview change**. Column L lists, sheet by sheet, every cell it would change, with addresses. Nothing changes. Problems stop the change:
   - the old code is not in SKU mapping-Pricing;
   - the new code is already used;
   - the new description is already another SKU's;
   - another SKU has the same old description.
6. Click **Change SKU**. A backup copy is saved first (`... - before changing SKU 12345678 ....xlsm`). Then the cells are changed and the pivots are refreshed.

How it works:
- **Typed cells.** It looks at every typed cell in every sheet, except the New SKU sheet. A cell that is exactly the old code is replaced, keeping the cell's type (text stays text, a number stays a number). A text cell that is exactly the old description is replaced too.
- **Formulas.** A formula that has the old code or the quoted old description written into it is edited. Formulas that only look the SKU up (VLOOKUP, INDEX/MATCH, SUMIFS on the code column) are left as they are, and follow the typed cells.
- **Pivots.** Pivot cells are skipped. Before the change it notes which pivot fields show the old code or description, and whether it was ticked or was the page filter. After the refresh it gives the new name the same state, so manual filters (for example "for 3 year files") still show the SKU.
- **Protected sheets** such as IMS Stock are unprotected for the change and protected again. Array formulas are listed, not changed.
- **Linked files** (the Target Split workbooks behind FCST column S) are not changed. The preview reminds you to update them.

Not changed:
- SKU mapping-Pricing Status and the other master columns;
- the Mosaic code;
- defined names, data validation lists and conditional formatting rules that have the code written in. None were found in "Pre-MSR  Oct - After Shift - SKU Update3-Extra.xlsm".

## Variables on the New SKU sheet

| Variable | Example | Written to |
|---|---|---|
| **New SKU code** | 12625767 | Every sheet's key column |
| **Copy settings from SKU** (template) | 12596137 | Picks the rows to copy |
| **SKU description** | NIDO 1+ GUM 12x900g N4 XA | SKU mapping C, FCST D, Table1/Table19 Product Desc, report tables, 2027_Comparision AC. Reports match on this text, so type it exactly as the system shows it |
| Short description | NIDO 1+ 900g | SKU mapping E (Table1 / Table19 `Unifed` read it); 2027 M vs M-1 `Table12` row, added when no SKU has this short description yet |
| Unified description (SKUDESC) | NIDO 1+ GUM 12X900g XA | SKU mapping A, FCST E, Table1/Table19 SKUDESC and Unifed |
| Mosaic product code | 20000000009 | SKU mapping D. CYvsNY and the Mosaic sheets find SKUs by this code |
| Category (Main_Cat) | Dairy | SKU mapping L, BDA A, Target Split D, Table1/Table19, report tables |
| CATDESC | 01 Ambient Dairy | FCST A, Table1/Table19 CATDESC |
| F-Group ID | E7J | SKU mapping F, Table1/Table19 Fgroup Code |
| F-Group | Nido 1+ | SKU mapping G, BDA B, Target Split E, Table1/Table19 FGROUP_DESC |
| Update F-Group | Nido 1+ | Table1/Table19, 2027 report tables and data block; Gap OL `Table6` row, added when the F-Group is new |
| Status / Active | Current / 1 | SKU mapping M / N. Status is Current, New or IR. The Status slicers on 2027 OL and the 2027 reports show Current and IR in the BSE2 file, so a New SKU stays hidden there until its status changes (Preview warns) |
| **Tons per case** | 0.0108 | SKU mapping P, IMS Stock "Format" column. Drives every tons figure |
| 1C / Arabic description | | SKU mapping Z / AA |
| IMS Stock label | | IMS Stock "SKU" column |
| **WSL price, retail price** per case (SYP) | 16431 / 17088 | SKU mapping I / J. Price Chnge builds the monthly prices from these |
| Trade terms 2026 and 2027: Progressive, Coverage, Commission, Distribution, each for Sub-Distributor (SD) and Medical WH (MDD) | 3% / 4% ... | BDA Allowance Assist (+Y+1) E:J and M:N; Distribution by Month (+Y+1) E:F; Commission by Month E:F (2026 commission: month 1 links to BDA I:J, the other months repeat it and can be typed over) |
| Customer split 2026 and 2027 | 30% SD / 70% Medical WH | Target Split (+Y+1) H, applied to every month. Must add up to 100% |
| DSCR % per customer, 2026 and 2027 | blank | DSCR, DSCR -LS, DSCR Y+1 F:Q. Blank copies the template's monthly % |
| Volume plan by month, 2026 and 2027 (cases) | 1,500 | IMS Stock "Nestle Syria Target" row. The forecast row follows it, and FCST / FCST Y+1 read it |
| Arrivals by month, 2026 and 2027 | | Fayrouz-Arrival / Fayrouz-Arrival-2027. IMS Stock reads them |
| TPS target by month, 2026 and 2027 | | Fayrouz Target / Fayrouz 2027-Target |
| Stock on hand before the first month | 0 | IMS Stock Closing Stock, first month column (Dec 2024) |
| Refresh pivots when done | Yes | Runs Refresh_all after adding |

The macro does not ask for monthly price-change % (Price Chnge), monthly distribution % or the Table1 / Table19 row layout. These come from the template SKU. In 2026, distribution uses the BDA rate in every month.

## Where a SKU lives

### Sheets you add rows to (the macro does all of these)

| Sheet | Rows per SKU | How it is linked |
|---|---|---|
| SKU mapping-Pricing | 1 | The master. Code in B. Read by code almost everywhere |
| Price Chnge | 2 | 2026 row in rows 6-55. The 2027 row sits 56 rows lower and links back (`=A44`). Table1 reads `$A$2:$AL$55`; Y+1 reads `$A$58:$AL$103` |
| BDA Allowance Assist (+Y+1) | 1 each | C:D link to the master **by row position** (BDA row = master row + 2) |
| Distribution by Month (+Y+1) | 1 each | A:D link to BDA **on the same row number**. Formula rows are ready to row 45, which is why the old guide says "added automatically" |
| Commission by Month | 1 | 2026 commission % by month and customer. A:D link to BDA on the same row number. Table1 Commission% finds the row by matching the code in **Distribution by Month** column C, so this row must sit on the Distribution by Month row number |
| FCST / FCST Y+1 | 1 each | Read IMS Stock by code (the `=VALUE(code)` column, AU since the F-Group column was added) |
| IMS Stock | 8-row block | Sales, Forecast, Nestle Syria Target, Expected Arrivals, Total / Opening / Closing Stock, Cover Days. Arrivals are looked up from the Fayrouz sheets by code **as text**. The sheet is protected (no password). The macro finds its columns from the header row ("code", "SKU", "Stock Cover Days", "F-Group", "Format" and the month dates), so inserted columns do not break it |
| Fayrouz-Arrival / Fayrouz-Arrival-2027 | 1 each | Feed IMS Stock arrivals |
| Fayrouz Target / Fayrouz 2027-Target | 1 each | Check sheets. Only FCST Y+1 column R counts the 2027 one |
| DSCR, DSCR Y+1 (and DSCR -LS) | 2 each | One row per customer; key = customer code & SKU code. DSCR -LS feeds nothing; it is kept in step anyway |
| Target Split (+Y+1) | 2 each | One row per customer; key = customer name & SKU code |
| Consolidation RAW, `Table1` | 24 for a new SKU | 2 customers x 12 months. Existing SKUs also have 24 M-1 rows (last month's Pre-MSR); a new SKU has no M-1 history, so the macro copies only the template's M rows. Columns A:Q are typed keys; everything else is formulas that read the sheets above |
| Y+1, `Table19` | 24 for a new SKU | Same as Table1: M and M-1 since November, M only for a new SKU |
| 2027_Comparision data block W:AH | 24 (12 for a 2027-only SKU) | Source of the pivots on 2027_Comparision, 2027 VS ICP and 2027_by Quarter |
| 2027_by Quarter AD:AN | same as above | Links to 2027_Comparision row by row (6 rows apart) |
| 2027_Comparision `Table20`, 2027 VS ICP `Table2011` | 1 each | Row inside the category section, so the "Total Dairy" subtotal includes it |
| 2027 M vs M-1 `Table12` | 1 per short description | Category / F-Group / short description; sums its pivot (X:AI) by short description. Shared by SKUs with the same short description |
| Gap OL `Table6` | 1 per report F-Group | Category / Update F-Group, inside the category section. Shared by the SKUs of that F-Group |
| Monthly Gap `Table612` | 1 per pivot row | Reads the pivot list in CD:CE by position. Matched to the list after each refresh |
| TTS `Table4` | 1 per F-Group pivot row | Reads the pivot list in AA:AB by position. Matched to the list after each refresh, so a new F-Group gets its row |
| CYvsNY `Table9` | 2 | 2026 and 2027 rows; has room for 41 SKUs (39 after the next one) |

### Sheets that update by themselves

- **After Refresh_all:** 2026 DF, Gap OL pivots, 2027 M vs M-1 pivot, for 3 year files (+ Next Year), Monthly Gap pivots, TTS, 2027 OL, 2027 TTS, CYvsNY pivots, DF, Data Check, Check, Check Mosaic, Target Check, FCST_Mosaic Submission, 2027-Mosaic, Fico Assist pivots, 2027 Allowance, OTAllowance %. All are pivots over Table1, Table19 or the 2027_Comparision block.
- **System exports each month:** Sales Vs Activity-RAW, NNSBreakDown-RAW and their M-1 copies.

### Not done by the macro

- **A new F-Group:** 2027 TTS is pivots only and updates on refresh; TTS `Table4` and Gap OL's table are handled by the macro. Preview warns when the F-Group differs from the template's.
- **The Fico Assist top block and For-Fico** (pasted values), **2027 check** columns K:AA and **Sales vs Activity 2027**. Update these by hand if you use them; nothing else reads them.

## Link problems found

These exist today, before any new SKU is added. `NewSKU_WidenLookupRanges` fixes items 1–4.

1. **Lookups stop at row 25 of SKU mapping-Pricing, but the master runs to row 39.** Y+1 column AP ("Active") is `COUNTIF('SKU mapping-Pricing'!$B$2:$B$25, code)`, so all 14 SKUs in rows 26–39 show Active = 0 in 2027. That includes four Current SKUs: both Nido N4 packs, Maggi Powder Stock and KIT KAT 2 Finger. Consolidation RAW uses the whole column, so in 2026 they are Active = 1. No report filters on this today (2027 OL shows "Active = (All)"). A filter on Active = 1 would drop these SKUs from 2027. FCST and FCST Y+1 use the same short range in columns C, D and T. If rows 26+ were meant to mean "not active", delete that one line from the macro before running it.
2. **Distribution by Month (+Y+1) is read only to row 45.** The list ends at row 41, so four more SKUs fit. A fifth would get #N/A in Table1 / Table19. The macro widens it to row 200.
3. **The Price Chnge 2027 block is read only to row 103**, and the block ends at row 99. The macro widens it to row 200. The 2026 block (rows 6–55) has room for 12 more SKUs and cannot be widened, because the 2027 block starts at row 57. Preview stops if it is full.
4. **Smaller short ranges:**
   - `'FCST Y+1'!$C$4:$Q$51` in Fayrouz 2027-Target.
   - `'Fayrouz 2027-Target'!$A$4:$A$44` in FCST Y+1.
   - One Fayrouz Target formula that reads IMS Stock to row 223.
   - Master lookups ending at rows 59–100 in Consolidation RAW, Y+1, CYvsNY, the 2027 reports and Sales vs Activity 2027.

   All are widened.
5. **Rows linked by position, not by code:**
   - SKU mapping-Pricing → BDA (row + 2) → Distribution by Month and Commission by Month (same row).
   - Price Chnge 2026 → 2027 (+56 rows).
   - 2027_Comparision → 2027_by Quarter (−6 rows).
   - The Monthly Gap pivot list → Table612 (+3 rows), and the TTS pivot list → Table4 (+3 rows). Both are matched after each refresh.

   Adding at the end of each list keeps them aligned. Inserting in the middle of only one of these sheets breaks the link silently, and nothing shows an error. The macro adds at the end, links BDA C:D to the master row explicitly, and warns if the rows do not line up.
6. **Codes stored as text in some places.** IMS Stock's "code" column (C since the F-Group column was added) and the Fayrouz sheets' column A hold the code as text, and the arrivals lookup matches text. Elsewhere the code is a number. Typing the code as a number there makes arrivals show 0. The macro writes text there.
7. **IMS Stock filter.** IMS Stock has a saved filter (A7:AS335 in the After Shift file). New blocks go below it, so they always show. Extend the filter range if you filter that sheet.

## Files

- `spec/NewSKU.bas`: the VBA module (setup, preview, add, widen ranges, delete a SKU).
- `spec/new-sku-guide.md`: this guide.
