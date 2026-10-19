# Adding a new SKU to the Pre-MSR workbook

Checked against `Pre-MSR  Oct.xlsm` (saved 7 Oct 2026). The reference was the last SKU added by hand, NIDO 1+ GUM 12x900g N4 XA (12625767). This guide traces every place that SKU appears.

## In short

- A SKU appears in about 40 sheets. Twenty-five of them need rows added. The rest are pivots that update on **Refresh_all**, or system exports that fill in by themselves.
- `spec/NewSKU.bas` is a VBA module. It builds a **New SKU** sheet that holds every variable a SKU needs. It then copies an existing "template" SKU's rows in each of those 25 sheets and writes in your values.
- The New SKU sheet has a live checklist of 35 checks. It also works for SKUs added by hand.
- Run **Widen lookup ranges** once before the first new SKU. Several lookups stop before the end of their lists already (see [Link problems](#link-problems-found)).

The macro was written from the workbook's structure. It was not run in Excel here, because this environment has no Excel. Use it on a copy first. Run **Preview**, which changes nothing, before **Add SKU**.

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

## Variables on the New SKU sheet

| Variable | Example | Written to |
|---|---|---|
| **New SKU code** | 12625767 | Every sheet's key column |
| **Copy settings from SKU** (template) | 12596137 | Picks the rows to copy |
| **SKU description** | NIDO 1+ GUM 12x900g N4 XA | SKU mapping C, FCST D, Table1/Table19 Product Desc, report tables, 2027_Comparision AC. Reports match on this text, so type it exactly as the system shows it |
| Short description | NIDO 1+ 900g | SKU mapping E |
| Unified description (SKUDESC) | NIDO 1+ GUM 12X900g XA | SKU mapping A, FCST E, Table1/Table19 SKUDESC and Unifed |
| Mosaic product code | 20000000009 | SKU mapping D. CYvsNY and the Mosaic sheets find SKUs by this code |
| Category (Main_Cat) | Dairy | SKU mapping L, BDA A, Target Split D, Table1/Table19, report tables |
| CATDESC | 01 Ambient Dairy | FCST A, Table1/Table19 CATDESC |
| F-Group ID | E7J | SKU mapping F, Table1/Table19 Fgroup Code |
| F-Group | Nido 1+ | SKU mapping G, BDA B, Target Split E, Table1/Table19 FGROUP_DESC |
| Update F-Group | Nido 1+ | Table1/Table19, 2027 report tables and data block |
| Status / Active | Current / 1 | SKU mapping M / N |
| **Tons per case** | 0.0108 | SKU mapping P, IMS Stock BE. Drives every tons figure |
| 1C / Arabic description | | SKU mapping Z / AA |
| IMS Stock label | | IMS Stock D |
| **WSL price, retail price** per case (SYP) | 16431 / 17088 | SKU mapping I / J. Price Chnge builds the monthly prices from these |
| Trade terms 2026 and 2027: Progressive, Coverage, Commission, Distribution, each for Sub-Distributor (SD) and Medical WH (MDD) | 3% / 4% ... | BDA Allowance Assist (+Y+1) E:J and M:N; Distribution by Month (+Y+1) E:F |
| Customer split 2026 and 2027 | 30% SD / 70% Medical WH | Target Split (+Y+1) H, applied to every month. Must add up to 100% |
| DSCR % per customer, 2026 and 2027 | blank | DSCR, DSCR -LS, DSCR Y+1 F:Q. Blank copies the template's monthly % |
| Volume plan by month, 2026 and 2027 (cases) | 1,500 | IMS Stock "Nestle Syria Target" row. The forecast row follows it, and FCST / FCST Y+1 read it |
| Arrivals by month, 2026 and 2027 | | Fayrouz-Arrival / Fayrouz-Arrival-2027. IMS Stock reads them |
| TPS target by month, 2026 and 2027 | | Fayrouz Target / Fayrouz 2027-Target |
| Stock on hand before the first month | 0 | IMS Stock Closing Stock, column F |
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
| FCST / FCST Y+1 | 1 each | Read IMS Stock by code (IMS column AT) |
| IMS Stock | 8-row block | Sales, Forecast, Nestle Syria Target, Expected Arrivals, Total / Opening / Closing Stock, Cover Days. Arrivals are looked up from the Fayrouz sheets by code **as text**. The sheet is protected (no password) |
| Fayrouz-Arrival / Fayrouz-Arrival-2027 | 1 each | Feed IMS Stock arrivals |
| Fayrouz Target / Fayrouz 2027-Target | 1 each | Check sheets. Only FCST Y+1 column R counts the 2027 one |
| DSCR, DSCR Y+1 (and DSCR -LS) | 2 each | One row per customer; key = customer code & SKU code. DSCR -LS feeds nothing; it is kept in step anyway |
| Target Split (+Y+1) | 2 each | One row per customer; key = customer name & SKU code |
| Consolidation RAW, `Table1` | 48 | 2 customers x 12 months x M and M-1. Columns A:Q are typed keys; everything else is formulas that read the sheets above |
| Y+1, `Table19` | 24 | 2 customers x 12 months, M only |
| 2027_Comparision data block W:AH | 24 (12 for a 2027-only SKU) | Source of the pivots on 2027_Comparision, 2027 VS ICP, Ramadan and 2027_by Quarter |
| 2027_by Quarter AD:AN | same as above | Links to 2027_Comparision row by row (6 rows apart) |
| 2027_Comparision `Table20`, Ramadan `Table2014`, 2027 VS ICP `Table2011` | 1 each | Row inside the category section, so the "Total Dairy" subtotal includes it |
| Monthly Gap `Table612` | 1 | Reads the pivot list in CD:CE by position, so one more SKU needs one more table row |
| CYvsNY `Table9` | 2 | 2026 and 2027 rows; has room for 41 SKUs (39 after the next one) |

### Sheets that update by themselves

- **After Refresh_all:** 2026 DF, Gap OL, Monthly Gap pivots, TTS, 2027 OL, 2027 TTS, CYvsNY pivots, DF, Data Check, Check, Check Mosaic, Target Check, FCST_Mosaic Submission, 2027-Mosaic, Fico Assist pivots, 2027 Allowance, OTAllowance %. All are pivots over Table1, Table19 or the 2027_Comparision block.
- **System exports each month:** Sales Vs Activity-RAW, NNSBreakDown-RAW and their M-1 copies.

### Not done by the macro

- **A new F-Group:** Gap OL, TTS and 2027 TTS have one row per F-Group. A SKU in a new F-Group needs rows there by hand. Preview warns when the F-Group differs from the template's.
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
   - SKU mapping-Pricing → BDA (row + 2) → Distribution by Month (same row).
   - Price Chnge 2026 → 2027 (+56 rows).
   - 2027_Comparision → 2027_by Quarter (−6 rows).
   - The Monthly Gap pivot list → Table612 (+3 rows).

   Adding at the end of each list keeps them aligned. Inserting in the middle of only one of these sheets breaks the link silently, and nothing shows an error. The macro adds at the end, links BDA C:D to the master row explicitly, and warns if the rows do not line up.
6. **Codes stored as text in some places.** IMS Stock column B and the Fayrouz sheets' column A hold the code as text, and the arrivals lookup matches text. Elsewhere the code is a number. Typing the code as a number in IMS Stock B makes arrivals show 0. The macro writes text there.
7. **IMS Stock filter.** IMS Stock has a saved filter on B7:AR311. New blocks go below it, so they always show. Extend the filter range if you filter that sheet.

## Files

- `spec/NewSKU.bas`: the VBA module (setup, preview, add, widen ranges).
- `spec/new-sku-guide.md`: this guide.
