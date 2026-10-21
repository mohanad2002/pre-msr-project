Attribute VB_Name = "NewSKU"
'==============================================================================
' NewSKU - add a new SKU to the Pre-MSR workbook in one step.
'
' Macros (Alt+F8):
'   Setup_NewSKU_Sheet        Builds the "New SKU" input sheet. Run once.
'   NewSKU_Preview            Checks the inputs and lists every row it would add.
'                             Changes nothing.
'   NewSKU_Add                Saves a backup copy, then adds the SKU everywhere.
'   NewSKU_WidenLookupRanges  One-time fix: widens lookup ranges that stop
'                             before the end of their lists.
'
' How it works: every sheet that keeps one row (or block of rows) per SKU gets
' a copy of the rows of a "template" SKU that you name, normally a SKU in the
' same F-Group. The copy keeps the formulas. The macro then writes the new
' code, description and the values from the New SKU sheet. New rows go at the
' end of each list, because SKU mapping-Pricing, BDA Allowance Assist and
' Distribution by Month are linked row by row.
'
' Guide: spec/new-sku-guide.md
'==============================================================================
Option Explicit

Private Const FORM_NAME As String = "New SKU"
Private Const SD_CODE As String = "SY010005"
Private Const SD_NAME As String = "Sub-Distributor"
Private Const MW_CODE As String = "SY010006"
Private Const MW_NAME As String = "Medical WH"
Private Const LOG_COL As Long = 12              ' log starts in column L of the New SKU sheet

Private dryRun As Boolean
Private logRow As Long
Private nProblems As Long
Private nWarnings As Long

' ---- inputs, read from the named cells on the New SKU sheet ----
Private sCode As Double
Private sTpl As Double
Private sTplDesc As String
Private sDesc As String
Private sShort As String
Private sUnified As String
Private sMosaic As String
Private sCat As String
Private sCatDesc As String
Private sFgId As String
Private sFg As String
Private sUpdFg As String
Private sStatus As String
Private sActive As Variant
Private sFormat As Double
Private s1C As String
Private sArabic As String
Private sImsLabel As String
Private sWsl As Double
Private sRet As Double
Private sRefresh As Boolean
Private sOpenStock As Variant
Private terms(1 To 8, 1 To 2) As Variant        ' Progressive, Coverage, Commission, Distribution (SD, MDD) x 2026, 2027
Private custSplit(1 To 2, 1 To 2) As Variant    ' Sub-Distributor, Medical WH x 2026, 2027
Private dscrPct(1 To 2, 1 To 2) As Variant      ' blank = copy the template's monthly DSCR %
Private plan(1 To 6, 1 To 12) As Variant        ' volume 2026, volume 2027, arrivals 2026, arrivals 2027, TPS 2026, TPS 2027

' ---- positions found while running ----
Private mRow As Long                            ' new row in SKU mapping-Pricing
Private bdaRow(1 To 2) As Long                  ' new row in BDA Allowance Assist (2026, 2027)

' ---- IMS Stock columns, read from its header row ----
Private imsHdr As Long, imsCode As Long, imsLabel As Long, imsSku As Long
Private imsFg As Long, imsFormat As Long, imsFirst As Long
Private imsM26(1 To 12) As Long, imsM27(1 To 12) As Long

'==============================================================================
' Public macros
'==============================================================================

Public Sub NewSKU_Preview()
    RunNewSKU True
End Sub

Public Sub NewSKU_Add()
    If MsgBox("This adds the new SKU to about 25 sheets." & vbLf & _
              "A backup copy of the workbook is saved first." & vbLf & vbLf & _
              "Run Preview first if you have not. Continue?", _
              vbYesNo + vbQuestion, "Add new SKU") <> vbYes Then Exit Sub
    RunNewSKU False
End Sub

Public Sub Setup_NewSKU_Sheet()
    Dim f As Worksheet, afterSheet As Worksheet
    Set f = Sh(FORM_NAME)
    If Not f Is Nothing Then
        If MsgBox("The New SKU sheet already exists. Rebuild it?" & vbLf & _
                  "What you typed on it will be cleared.", vbYesNo + vbQuestion, "New SKU") <> vbYes Then Exit Sub
        Application.DisplayAlerts = False
        f.Delete
        Application.DisplayAlerts = True
    End If
    Set afterSheet = Sh("Guide line for new SKU update")
    If afterSheet Is Nothing Then Set afterSheet = ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    Set f = ThisWorkbook.Worksheets.Add(After:=afterSheet)
    f.Name = FORM_NAME
    f.Tab.Color = RGB(255, 192, 0)              ' yellow tab = input sheet, as in Guideline.docx
    ActiveWindow.DisplayGridlines = False
    BuildForm f
    f.Activate
    f.Range("C5").Select
    MsgBox "The New SKU sheet is ready." & vbLf & _
           "Fill the yellow cells, then click Preview.", vbInformation, "New SKU"
End Sub

Public Sub NewSKU_WidenLookupRanges()
    Dim calcMode As Long, n As Long
    If MsgBox("This widens lookup ranges that stop before the end of their lists" & vbLf & _
              "(for example 'SKU mapping-Pricing'!$B$2:$B$25 and 'Distribution by Month'!$A$1:$AB$45)." & vbLf & _
              "A backup copy is saved first. The changes are listed in column L of the New SKU sheet." & vbLf & vbLf & _
              "Continue?", vbYesNo + vbQuestion, "Widen lookup ranges") <> vbYes Then Exit Sub

    PrepareLog "Widened lookup ranges on " & Format(Now, "dd mmm yyyy hh:nn")
    calcMode = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    sCode = 0
    SaveBackup "before widening ranges"

    ' SKU master read only to row 25 (the master now runs past row 39)
    n = n + Widen("FCST", "'SKU mapping-Pricing'!$A$2:$D$25", "'SKU mapping-Pricing'!$A$2:$D$500")
    n = n + Widen("FCST", "'SKU mapping-Pricing'!$B$2:$C$25", "'SKU mapping-Pricing'!$B$2:$C$500")
    n = n + Widen("FCST", "'SKU mapping-Pricing'!$B$2:$B$25", "'SKU mapping-Pricing'!$B$2:$B$500")
    n = n + Widen("FCST Y+1", "'SKU mapping-Pricing'!$A$2:$D$25", "'SKU mapping-Pricing'!$A$2:$D$500")
    n = n + Widen("FCST Y+1", "'SKU mapping-Pricing'!$B$2:$C$25", "'SKU mapping-Pricing'!$B$2:$C$500")
    n = n + Widen("FCST Y+1", "'SKU mapping-Pricing'!$B$2:$B$25", "'SKU mapping-Pricing'!$B$2:$B$500")
    n = n + Widen("Y+1", "'SKU mapping-Pricing'!$B$2:$B$25", "'SKU mapping-Pricing'!$B$2:$B$500")
    n = n + Widen("SKU mapping-Pricing", "$A$2:$A$25", "$A$2:$A$500")

    ' Distribution by Month read only to row 45 (the list ends at row 41)
    n = n + Widen("Consolidation RAW", "'Distribution by Month'!$A$1:$AB$45", "'Distribution by Month'!$A$1:$AB$200")
    n = n + Widen("Consolidation RAW", "'Distribution by Month'!$C$1:$C$45", "'Distribution by Month'!$C$1:$C$200")
    n = n + Widen("Y+1", "'Distribution by Month Y+1'!$A$1:$AB$45", "'Distribution by Month Y+1'!$A$1:$AB$200")
    n = n + Widen("Y+1", "'Distribution by Month Y+1'!$C$1:$C$45", "'Distribution by Month Y+1'!$C$1:$C$200")

    ' 2027 price block read only to row 103 (the block ends at row 99)
    n = n + Widen("Y+1", "'Price Chnge'!$A$58:$AL$103", "'Price Chnge'!$A$58:$AL$200")
    n = n + Widen("Y+1", "'Price Chnge'!$A$58:$A$103", "'Price Chnge'!$A$58:$A$200")

    ' Other short ranges
    n = n + Widen("Fayrouz 2027-Target", "'FCST Y+1'!$C$4:$Q$51", "'FCST Y+1'!$C$4:$Q$200")
    n = n + Widen("FCST Y+1", "'Fayrouz 2027-Target'!$A$4:$A$44", "'Fayrouz 2027-Target'!$A$4:$A$200")
    n = n + Widen("Fayrouz Target", "$223,", "$10016,")              ' 'IMS Stock'!..$8:..$223 ranges, whatever the column
    n = n + Widen("Fayrouz Target", "$223)", "$10016)")
    n = n + Widen("Consolidation RAW", "'SKU mapping-Pricing'!$B$2:$F$60", "'SKU mapping-Pricing'!$B$2:$F$500")
    n = n + Widen("Consolidation RAW", "'SKU mapping-Pricing'!$B$2:$D$94", "'SKU mapping-Pricing'!$B$2:$D$500")
    n = n + Widen("Consolidation RAW", "'SKU mapping-Pricing'!$B$2:$P$100", "'SKU mapping-Pricing'!$B$2:$P$500")
    n = n + Widen("Y+1", "'SKU mapping-Pricing'!$B$2:$F$60", "'SKU mapping-Pricing'!$B$2:$F$500")
    n = n + Widen("Y+1", "'SKU mapping-Pricing'!$B$2:$D$94", "'SKU mapping-Pricing'!$B$2:$D$500")
    n = n + Widen("Y+1", "'SKU mapping-Pricing'!$B$2:$P$100", "'SKU mapping-Pricing'!$B$2:$P$500")
    n = n + Widen("Y+1", "'SKU mapping-Pricing'!$B$2:$M$85", "'SKU mapping-Pricing'!$B$2:$M$500")
    n = n + Widen("2027_Comparision", "'SKU mapping-Pricing'!$C$2:$M$97", "'SKU mapping-Pricing'!$C$2:$M$500")
    n = n + Widen("2027 VS ICP", "'SKU mapping-Pricing'!$C$2:$M$97", "'SKU mapping-Pricing'!$C$2:$M$500")
    n = n + Widen("Ramadan", "'SKU mapping-Pricing'!$C$2:$M$97", "'SKU mapping-Pricing'!$C$2:$M$500")
    n = n + Widen("CYvsNY", "'SKU mapping-Pricing'!$D$2:$N$100", "'SKU mapping-Pricing'!$D$2:$N$500")
    n = n + Widen("CYvsNY", "'SKU mapping-Pricing'!$B$1:$N$100", "'SKU mapping-Pricing'!$B$1:$N$500")
    n = n + Widen("CYvsNY", "'SKU mapping-Pricing'!$D$1:$D$100", "'SKU mapping-Pricing'!$D$1:$D$500")
    n = n + Widen("Sales vs Activity 2027", "'SKU mapping-Pricing'!$B$2:$B$59", "'SKU mapping-Pricing'!$B$2:$B$500")

    Application.Calculation = calcMode
    Application.Calculate
    Application.ScreenUpdating = True
    MsgBox n & " formulas updated. The list is in column L of the New SKU sheet." & vbLf & _
           "Price Chnge rows 6-55 (2026 prices) cannot be widened: the 2027 block starts at row 57.", _
           vbInformation, "Widen lookup ranges"
End Sub

'==============================================================================
' Main run
'==============================================================================

Private Sub RunNewSKU(ByVal isDry As Boolean)
    Dim calcMode As Long, stepName As String
    If Sh(FORM_NAME) Is Nothing Then
        MsgBox "Run Setup_NewSKU_Sheet first.", vbExclamation, "New SKU"
        Exit Sub
    End If
    dryRun = isDry
    PrepareLog IIf(isDry, "Preview on " & Format(Now, "dd mmm yyyy hh:nn") & ": nothing was changed", _
                          "Added on " & Format(Now, "dd mmm yyyy hh:nn"))
    ReadInputs
    Validate

    If nProblems > 0 And Not isDry Then
        MsgBox nProblems & " problem(s) found. Nothing was changed." & vbLf & _
               "See the log in column L of the New SKU sheet.", vbExclamation, "New SKU"
        Exit Sub
    End If

    calcMode = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    On Error GoTo Failed

    If Not isDry Then
        stepName = "backup"
        If Not SaveBackup("before SKU " & Format(sCode, "0")) Then
            If MsgBox("The backup copy could not be saved. Continue without a backup?", _
                      vbYesNo + vbExclamation, "New SKU") <> vbYes Then
                Application.Calculation = calcMode
                Application.EnableEvents = True
                Application.ScreenUpdating = True
                Exit Sub
            End If
        End If
    End If

    stepName = "SKU mapping-Pricing": Step_SkuMapping
    stepName = "Price Chnge": Step_PriceChange
    stepName = "BDA Allowance Assist": Step_BDA "BDA Allowance Assist", 1
    stepName = "BDA Allowance Assist Y+1": Step_BDA "BDA Allowance Assist Y+1", 2
    stepName = "Distribution by Month": Step_Distribution "Distribution by Month", "BDA Allowance Assist", 1
    stepName = "Distribution by Month Y+1": Step_Distribution "Distribution by Month Y+1", "BDA Allowance Assist Y+1", 2
    stepName = "FCST": Step_FCST "FCST"
    stepName = "FCST Y+1": Step_FCST "FCST Y+1"
    stepName = "IMS Stock": Step_IMS
    stepName = "Fayrouz-Arrival": Step_Fayrouz "Fayrouz-Arrival", 5, 8, 3
    stepName = "Fayrouz-Arrival-2027": Step_Fayrouz "Fayrouz-Arrival-2027", 5, 8, 4
    stepName = "Fayrouz Target": Step_Fayrouz "Fayrouz Target", 4, 7, 5
    stepName = "Fayrouz 2027-Target": Step_Fayrouz "Fayrouz 2027-Target", 5, 8, 6
    stepName = "DSCR": Step_DSCR "DSCR", 1
    stepName = "DSCR -LS": Step_DSCR "DSCR -LS", 1
    stepName = "DSCR Y+1": Step_DSCR "DSCR Y+1", 2
    stepName = "Target Split": Step_TargetSplit "Target Split", 1
    stepName = "Target Split Y+1": Step_TargetSplit "Target Split Y+1", 2
    stepName = "Consolidation RAW": Step_EngineTable "Consolidation RAW", "Table1"
    stepName = "Y+1": Step_EngineTable "Y+1", "Table19"
    stepName = "2027_Comparision data block": Step_CompBlock
    stepName = "2027_Comparision table": Step_ReportTable "2027_Comparision", "Table20"
    stepName = "Ramadan table": Step_ReportTable "Ramadan", "Table2014"
    stepName = "2027 VS ICP table": Step_ReportTable "2027 VS ICP", "Table2011"
    stepName = "Monthly Gap table": Step_MonthlyGap
    stepName = "CYvsNY table": Step_CYvsNY

    Application.Calculation = calcMode
    Application.Calculate
    If Not isDry And sRefresh Then
        stepName = "refresh pivots"
        RefreshPivots
    End If

    Application.Calculation = calcMode
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Sh(FORM_NAME).Activate
    If isDry Then
        MsgBox "Preview done: " & nProblems & " problem(s), " & nWarnings & " point(s) to check." & vbLf & _
               "Nothing was changed. The planned rows are listed in column L.", _
               IIf(nProblems > 0, vbExclamation, vbInformation), "New SKU"
    Else
        MsgBox "SKU " & Format(sCode, "0") & " added (" & nWarnings & " point(s) to check)." & vbLf & _
               "Look at the checklist on the New SKU sheet: every line should say OK.", vbInformation, "New SKU"
    End If
    Exit Sub

Failed:
    Note "ERROR", stepName, Err.Description
    Application.Calculation = calcMode
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Stopped at step '" & stepName & "': " & Err.Description & vbLf & vbLf & _
           "Steps before it were done. The backup copy saved at the start has the workbook as it was.", _
           vbCritical, "New SKU"
End Sub

Private Sub ReadInputs()
    Dim i As Long, j As Long
    sCode = InNum("NS_Code")
    sTpl = InNum("NS_Template")
    sDesc = InText("NS_Desc")
    sShort = InText("NS_Short")
    sUnified = InText("NS_Unified")
    sMosaic = InText("NS_Mosaic")
    sCat = InText("NS_Category")
    sCatDesc = InText("NS_CatDesc")
    sFgId = InText("NS_FGroupID")
    sFg = InText("NS_FGroup")
    sUpdFg = InText("NS_UpdFGroup")
    sStatus = InText("NS_Status")
    sActive = InNum("NS_Active")
    sFormat = InNum("NS_Format")
    s1C = InText("NS_1C")
    sArabic = InText("NS_Arabic")
    sImsLabel = InText("NS_IMSLabel")
    sWsl = InNum("NS_WSL")
    sRet = InNum("NS_Ret")
    sRefresh = (UCase$(InText("NS_Refresh")) <> "NO")
    sOpenStock = InVal("NS_OpenStock")
    For i = 1 To 8
        For j = 1 To 2
            terms(i, j) = GridVal("NS_Terms", i, j)
        Next j
    Next i
    For i = 1 To 2
        For j = 1 To 2
            custSplit(i, j) = GridVal("NS_Split", i, j)
            dscrPct(i, j) = GridVal("NS_DSCR", i, j)
        Next j
    Next i
    For i = 1 To 6
        For j = 1 To 12
            plan(i, j) = GridVal("NS_Plan", i, j)
        Next j
    Next i
    If sShort = "" Then sShort = sDesc
    If sImsLabel = "" Then sImsLabel = sDesc
End Sub

Private Sub Validate()
    Dim w As Worksheet, r As Long, i As Long, tot As Double
    Set w = Sh("SKU mapping-Pricing")
    If sCode <= 0 Or sCode <> Int(sCode) Then Problem "New SKU", "Enter the new SKU code (a whole number, as in Sales Vs Activity-RAW)."
    r = FindKey(w, 2, sCode, 2, 2000)
    If sCode > 0 And r > 0 Then Problem w.Name, "SKU " & Format(sCode, "0") & " already exists (row " & r & ")."
    r = FindKey(w, 2, sTpl, 2, 2000)
    If r = 0 Then
        Problem w.Name, "Template SKU " & Format(sTpl, "0") & " was not found in column B."
    Else
        sTplDesc = Trim$(CStr(w.Cells(r, 3).Value))
    End If
    If sDesc = "" Then Problem "New SKU", "Enter the SKU description."
    If sDesc <> "" And FindText(w, 3, sDesc, 2, 2000) > 0 Then Problem w.Name, "The description is already used by another SKU: " & sDesc
    If sCat = "" Or sCatDesc = "" Or sFg = "" Or sUpdFg = "" Or sUnified = "" Then _
        Problem "New SKU", "Category, CATDESC, F-Group, Update F-Group and Unified description need a value (or a template SKU that has them)."
    If sFormat <= 0 Then Problem "New SKU", "Tons per case must be more than 0 (12 x 900 g = 0.0108)."
    If sWsl <= 0 Or sRet <= 0 Then Problem "New SKU", "WSL price and retail price must be more than 0."
    If sMosaic = "" Then Warn "New SKU", "No Mosaic product code: the SKU will not show in CYvsNY or the Mosaic sheets until one is added."
    For i = 1 To 2
        tot = Num(custSplit(1, i)) + Num(custSplit(2, i))
        If Abs(tot - 1) > 0.0005 Then Problem "New SKU", "The customer split for " & (2025 + i) & " adds up to " & Format(tot, "0.0%") & ", not 100%."
    Next i
    If r > 0 Then
        If Not IsText(w.Cells(r, 7).Value, sFg) Then Warn "New SKU", "The F-Group differs from the template's. A new F-Group also needs rows in Gap OL, TTS and 2027 TTS (not done by this macro)."
    End If
    CheckRoom
End Sub

' Room left in each list, read from the lookup ranges the engine tables really use.
Private Sub CheckRoom()
    Dim m As Long, b1 As Long, b2 As Long, p As Long, lim As Long, r As Long
    m = LastFilled(Sh("SKU mapping-Pricing"), 2, 2, 2000) + 1
    b1 = LastFilled(Sh("BDA Allowance Assist"), 3, 4, 2000) + 1
    b2 = LastFilled(Sh("BDA Allowance Assist Y+1"), 3, 4, 2000) + 1
    p = LastFilled(Sh("Price Chnge"), 1, 6, 55) + 1

    If p > 55 Then Problem "Price Chnge", "The 2026 price block (rows 6-55) is full."
    lim = RefLimit("Consolidation RAW", "Table1", "WSL Price", "Price Chnge")
    If lim > 0 And p > lim Then Problem "Price Chnge", "New 2026 price row " & p & " is below the range Consolidation RAW reads (row " & lim & ")."
    lim = RefLimit("Y+1", "Table19", "WSL Price", "Price Chnge")
    If lim > 0 And p + 56 > lim Then Problem "Price Chnge", "New 2027 price row " & (p + 56) & " is below the range Y+1 reads (row " & lim & "). Run NewSKU_WidenLookupRanges first."

    lim = RefLimit("Consolidation RAW", "Table1", "Distribution%", "Distribution by Month")
    If lim > 0 And b1 > lim Then Problem "Distribution by Month", "The new SKU would be row " & b1 & " but Consolidation RAW reads only to row " & lim & ". Run NewSKU_WidenLookupRanges first."
    lim = RefLimit("Y+1", "Table19", "Distribution%", "Distribution by Month Y+1")
    If lim > 0 And b2 > lim Then Problem "Distribution by Month Y+1", "The new SKU would be row " & b2 & " but Y+1 reads only to row " & lim & ". Run NewSKU_WidenLookupRanges first."

    lim = RefLimit("Consolidation RAW", "Table1", "BDA TTS%", "BDA Allowance Assist")
    If lim > 0 And b1 > lim Then Problem "BDA Allowance Assist", "New row " & b1 & " is below the range Consolidation RAW reads (row " & lim & ")."
    lim = RefLimit("Y+1", "Table19", "BDA TTS%", "BDA Allowance Assist Y+1")
    If lim > 0 And b2 > lim Then Problem "BDA Allowance Assist Y+1", "New row " & b2 & " is below the range Y+1 reads (row " & lim & ")."

    r = LastFilled(Sh("FCST"), 3, 4, 2000) + 1
    lim = RefLimit("Consolidation RAW", "Table1", "Initial Qty", "FCST")
    If lim > 0 And r > lim Then Problem "FCST", "New row " & r & " is below the range Consolidation RAW reads (row " & lim & ")."
    r = LastFilled(Sh("FCST Y+1"), 3, 4, 2000) + 1
    lim = RefLimit("Y+1", "Table19", "Initial Qty", "FCST Y+1")
    If lim > 0 And r > lim Then Problem "FCST Y+1", "New row " & r & " is below the range Y+1 reads (row " & lim & ")."

    r = LastFilled(Sh("Target Split"), 6, 3, 5000) + 2
    lim = RefLimit("Consolidation RAW", "Table1", "Contribution%", "Target Split")
    If lim > 0 And r > lim Then Problem "Target Split", "New rows end at " & r & ", below the range Consolidation RAW reads (row " & lim & ")."
    r = LastFilled(Sh("Target Split Y+1"), 6, 3, 5000) + 2
    lim = RefLimit("Y+1", "Table19", "Contribution%", "Target Split Y+1")
    If lim > 0 And r > lim Then Problem "Target Split Y+1", "New rows end at " & r & ", below the range Y+1 reads (row " & lim & ")."

    r = LastFilled(Sh("DSCR"), 4, 2, 5000) + 2
    lim = RefLimit("Consolidation RAW", "Table1", "DSCR%", "DSCR")
    If lim > 0 And r > lim Then Problem "DSCR", "New rows end at " & r & ", below the range Consolidation RAW reads (row " & lim & ")."
    r = LastFilled(Sh("DSCR Y+1"), 4, 2, 5000) + 2
    lim = RefLimit("Y+1", "Table19", "DSCR%", "DSCR Y+1")
    If lim > 0 And r > lim Then Problem "DSCR Y+1", "New rows end at " & r & ", below the range Y+1 reads (row " & lim & ")."

    lim = RefLimit("Y+1", "Table19", "Active", "SKU mapping-Pricing")
    If lim > 0 And m > lim Then Warn "Y+1", "The Active flag in Y+1 reads SKU mapping-Pricing only to row " & lim & ", so it will show 0 for this SKU. NewSKU_WidenLookupRanges fixes this."
End Sub

'==============================================================================
' Steps. Each one finds its rows, writes a log line, and stops there on Preview.
'==============================================================================

Private Sub Step_SkuMapping()
    Dim w As Worksheet, t As Long
    Set w = Sh("SKU mapping-Pricing")
    t = FindKey(w, 2, sTpl, 2, 2000)
    mRow = LastFilled(w, 2, 2, 2000) + 1
    Note w.Name, "row " & mRow, "Master row (copied from template row " & t & ")"
    If dryRun Or t = 0 Then Exit Sub
    CopyCells w.Range("A" & t & ":P" & t), w.Range("A" & mRow & ":P" & mRow)
    CopyCells w.Range("Z" & t & ":AB" & t), w.Range("Z" & mRow & ":AB" & mRow)
    CopyCells w.Range("AH" & t), w.Range("AH" & mRow)
    w.Range("A" & mRow).Value = sUnified
    w.Range("B" & mRow).Value = sCode
    w.Range("C" & mRow).Value = sDesc
    If IsNumeric(sMosaic) And sMosaic <> "" Then
        w.Range("D" & mRow).Value = CDbl(sMosaic)      ' CYvsNY matches Mosaic codes as numbers
    Else
        w.Range("D" & mRow).Value = sMosaic
    End If
    w.Range("E" & mRow).Value = sShort
    w.Range("F" & mRow).Value = sFgId
    w.Range("G" & mRow).Value = sFg
    w.Range("I" & mRow).Value = sWsl
    w.Range("J" & mRow).Value = sRet
    w.Range("L" & mRow).Value = sCat
    w.Range("M" & mRow).Value = sStatus
    w.Range("N" & mRow).Value = sActive
    w.Range("P" & mRow).Value = sFormat
    If s1C <> "" Then w.Range("Z" & mRow).Value = s1C Else w.Range("Z" & mRow).Value = sDesc
    w.Range("AA" & mRow).Value = sArabic
End Sub

Private Sub Step_PriceChange()
    Dim w As Worksheet, t As Long, r As Long, t2 As Long, r2 As Long, i As Long, fml As String
    Set w = Sh("Price Chnge")
    t = FindKey(w, 1, sTpl, 6, 55)
    r = LastFilled(w, 1, 6, 55) + 1
    ' The 2027 block (rows 57+) links back row by row: =A6, =A7, ...
    For i = 57 To 400
        fml = w.Cells(i, 1).Formula
        If t > 0 And fml = "=A" & t Then t2 = i
        If fml = "=A" & r Then r2 = i
    Next i
    If r2 = 0 And t2 > 0 Then r2 = t2 + (r - t)
    Note w.Name, "row " & r, "2026 prices (copied from row " & t & ")"
    If t2 = 0 Then
        Warn w.Name, "The template's 2027 price row was not found; add the 2027 row by hand."
    Else
        Note w.Name, "row " & r2, "2027 prices (copied from row " & t2 & ")"
    End If
    If t = 0 Then Problem w.Name, "Template SKU not found in rows 6-55."
    If dryRun Or t = 0 Then Exit Sub
    CopyCells w.Range("A" & t & ":AL" & t), w.Range("A" & r & ":AL" & r)
    w.Range("A" & r).Value = sCode
    w.Range("B" & r).Value = sDesc
    If t2 > 0 Then
        CopyCells w.Range("A" & t2 & ":AL" & t2), w.Range("A" & r2 & ":AL" & r2)
        w.Range("A" & r2).Formula = "=A" & r
        w.Range("B" & r2).Formula = "=B" & r
    End If
End Sub

Private Sub Step_BDA(ByVal sheetName As String, ByVal yr As Long)
    Dim w As Worksheet, t As Long, r As Long, i As Long, cols As Variant
    Set w = Sh(sheetName)
    t = FindKey(w, 3, sTpl, 4, 2000)
    r = LastFilled(w, 3, 4, 2000) + 1
    bdaRow(yr) = r
    If t = 0 Then
        t = r - 1
        Warn w.Name, "Template SKU not found here; layout copied from row " & t & "."
    End If
    If r <> mRow + 2 Then Warn w.Name, "Row " & r & " does not line up with SKU mapping-Pricing row " & mRow & " (usually row + 2). Columns C:D are linked to row " & mRow & " directly."
    Note w.Name, "row " & r, "Trade terms (copied from row " & t & ")"
    If dryRun Then Exit Sub
    CopyCells w.Range("A" & t & ":P" & t), w.Range("A" & r & ":P" & r)
    w.Range("A" & r).Value = sCat
    w.Range("B" & r).Value = sFg
    w.Range("C" & r).Formula = "='SKU mapping-Pricing'!B" & mRow
    w.Range("D" & r).Formula = "='SKU mapping-Pricing'!C" & mRow
    cols = Array(5, 6, 7, 8, 9, 10, 13, 14)     ' E:J progressive, coverage, commission; M:N distribution
    For i = 0 To 7
        w.Cells(r, cols(i)).Value = Num(terms(i + 1, yr))
    Next i
    w.Range("K" & r).Formula = "=E" & r & "+G" & r
    w.Range("L" & r).Formula = "=F" & r & "+H" & r
    w.Range("O" & r).Formula = "=M" & r & "+I" & r
    w.Range("P" & r).Formula = "=N" & r & "+J" & r
End Sub

Private Sub Step_Distribution(ByVal sheetName As String, ByVal bdaSheet As String, ByVal yr As Long)
    Dim w As Worksheet, t As Long, r As Long, c As Long
    Set w = Sh(sheetName)
    r = bdaRow(yr)                              ' same row number as BDA Allowance Assist
    t = FindKey(w, 3, sTpl, 4, 2000)
    If t = 0 Then t = r - 1
    Note w.Name, "row " & r, "Distribution % by month, linked to " & bdaSheet & " row " & r
    If dryRun Then Exit Sub
    CopyCells w.Range("A" & t & ":AB" & t), w.Range("A" & r & ":AB" & r)
    For c = 1 To 4
        w.Cells(r, c).Formula = "='" & bdaSheet & "'!" & ColL(c) & r
    Next c
    If yr = 1 Then
        w.Cells(r, 5).Formula = "='" & bdaSheet & "'!M" & r
        w.Cells(r, 6).Formula = "='" & bdaSheet & "'!N" & r
        For c = 7 To 28                         ' G:AB, Sub-Distributor and Medical WH by turns
            If c Mod 2 = 1 Then
                w.Cells(r, c).Formula = "=$E" & r
            Else
                w.Cells(r, c).Formula = "=$F" & r
            End If
        Next c
    Else
        w.Cells(r, 5).Value = Num(terms(7, 2))
        w.Cells(r, 6).Value = Num(terms(8, 2))
        For c = 7 To 28
            w.Cells(r, c).Formula = "=" & ColL(c - 2) & r
        Next c
    End If
End Sub

Private Sub Step_FCST(ByVal sheetName As String)
    Dim w As Worksheet, t As Long, r As Long
    Set w = Sh(sheetName)
    t = FindKey(w, 3, sTpl, 4, 2000)
    r = LastFilled(w, 3, 4, 2000) + 1
    If t = 0 Then
        t = r - 1
        Warn w.Name, "Template SKU not found here; layout copied from row " & t & "."
    End If
    Note w.Name, "row " & r, "Volume row, reads IMS Stock (copied from row " & t & ")"
    If dryRun Then Exit Sub
    CopyCells w.Range("A" & t & ":T" & t), w.Range("A" & r & ":T" & r)
    w.Range("A" & r).Value = sCatDesc
    w.Range("C" & r).Value = sCode
    w.Range("D" & r).Value = sDesc
    w.Range("E" & r).Value = sUnified
End Sub

' IMS Stock: one 8-row block per SKU. Columns are found from the header row ("code", "SKU",
' "Stock Cover Days", "Format" and the month dates), so inserted or moved columns do not matter.
Private Sub Step_IMS()
    Dim w As Worksheet, t As Long, r As Long, lastE As Long, i As Long, m As Long, c As Long
    Dim wasProt As Boolean, lastCol As Long, tplMaster As Long
    Dim oFc As Long, oTgt As Long, oArr As Long, oClose As Long
    Set w = Sh("IMS Stock")
    If Not ReadImsLayout(w) Then Exit Sub
    lastCol = w.UsedRange.Column + w.UsedRange.Columns.Count - 1
    lastE = LastFilled(w, imsLabel, imsHdr + 1, LastUsedRow(w))        ' not End(xlUp): it skips rows hidden by the filter
    For i = imsHdr + 1 To lastE
        If IsKey(w.Cells(i, imsCode).Value, sTpl) Then
            If IsText(w.Cells(i, imsLabel).Value, "IMS Actual ( Sales)") Then
                t = i
                Exit For
            End If
        End If
    Next i
    r = lastE + 1
    If t = 0 Then
        Problem w.Name, "The template SKU's 8-row block was not found (column " & ColL(imsCode) & " = template code, column " & ColL(imsLabel) & " = 'IMS Actual ( Sales)')."
        Exit Sub
    End If
    oFc = BlockRow(w, t, "IMS Actual ( Sales)-C")
    oTgt = BlockRow(w, t, "Nestle Syria Target")
    oArr = BlockRow(w, t, "Expected Arrivales in WH")
    oClose = BlockRow(w, t, "Closing Stock")
    If oFc < 0 Or oTgt < 0 Or oArr < 0 Or oClose < 0 Then
        Problem w.Name, "The template block (rows " & t & "-" & (t + 7) & ") does not have the usual 8 row labels in column " & ColL(imsLabel) & "."
        Exit Sub
    End If
    Note w.Name, "rows " & r & "-" & (r + 7), "8-row stock block (copied from rows " & t & "-" & (t + 7) & "; code in column " & ColL(imsCode) & ")"
    If Not RowsEmpty(w, r, 8, 1, lastCol) Then Note w.Name, "rows " & r & "-" & (r + 7), "Not empty: 8 new rows are inserted here instead, existing rows move down"
    If w.AutoFilterMode Then
        If w.AutoFilter.Range.Row + w.AutoFilter.Range.Rows.Count - 1 < r Then _
            Warn w.Name, "The new rows are below the filter range " & w.AutoFilter.Range.Address(False, False) & "; extend it if you filter this sheet."
    End If
    If dryRun Then Exit Sub

    wasProt = w.ProtectContents
    If wasProt Then w.Unprotect
    If Not RowsEmpty(w, r, 8, 1, lastCol) Then w.Rows(r & ":" & (r + 7)).Insert Shift:=xlDown
    For i = 0 To 7
        CopyCells w.Range(w.Cells(t + i, 1), w.Cells(t + i, lastCol)), w.Range(w.Cells(r + i, 1), w.Cells(r + i, lastCol))
    Next i
    w.Cells(r, imsCode).NumberFormat = "@"                              ' code as text, like the other blocks
    w.Cells(r, imsCode).Value = CStr(Format(sCode, "0"))
    If imsSku > 0 Then w.Cells(r, imsSku).Value = sImsLabel
    ' F-Group label column (if the sheet has one): keep the template's label unless the F-Group changes
    If imsFg > 0 Then
        tplMaster = FindKey(Sh("SKU mapping-Pricing"), 2, sTpl, 2, 2000)
        If tplMaster > 0 Then
            If Not IsText(Sh("SKU mapping-Pricing").Cells(tplMaster, 7).Value, sFg) Then
                For i = 0 To 7
                    If Not w.Cells(r + i, imsFg).HasFormula Then w.Cells(r + i, imsFg).Value = sFg
                Next i
            End If
        End If
    End If
    ' remove the template's numbers in the month columns; keep formulas and labels
    For i = 0 To 7
        For c = imsFirst To imsM27(12)
            With w.Cells(r + i, c)
                If Not .HasFormula Then
                    If Not IsEmpty(.Value) Then
                        If IsNumeric(.Value) Then .ClearContents
                    End If
                End If
            End With
        Next c
        If imsFormat > 0 Then
            If Not IsEmpty(w.Cells(t + i, imsFormat).Value) And Not w.Cells(t + i, imsFormat).HasFormula Then w.Cells(r + i, imsFormat).Value = sFormat
        End If
    Next i
    For m = 1 To 12
        ' forecast row follows the volume plan for future months (as in the other blocks)
        w.Cells(r + oFc, imsM26(m)).Formula = "=" & ColL(imsM26(m)) & (r + oTgt)
        w.Cells(r + oFc, imsM27(m)).Formula = "=" & ColL(imsM27(m)) & (r + oTgt)
        ' volume plan (Nestle Syria Target)
        If Given(plan(1, m)) Then w.Cells(r + oTgt, imsM26(m)).Value = Num(plan(1, m))
        If Given(plan(2, m)) Then w.Cells(r + oTgt, imsM27(m)).Value = Num(plan(2, m))
        ' expected arrivals come from the Fayrouz sheets; the row above the header holds each month's column number there
        w.Cells(r + oArr, imsM26(m)).Formula = "=IFERROR(VLOOKUP($" & ColL(imsCode) & (r + oArr) & ",'Fayrouz-Arrival'!$A$3:$S$146," & ColL(imsM26(m)) & "$" & (imsHdr - 1) & ",0),0)"
        w.Cells(r + oArr, imsM27(m)).Formula = "=IFERROR(VLOOKUP($" & ColL(imsCode) & (r + oArr) & ",'Fayrouz-Arrival-2027'!$A$3:$S$75," & ColL(imsM27(m)) & "$" & (imsHdr - 1) & ",0),0)"
    Next m
    w.Cells(r + oClose, imsFirst).Value = Num(sOpenStock)               ' starting stock: Closing Stock in the first month column
    If imsSku > 0 Then
        If w.Cells(t, imsSku).MergeCells Then
            Application.DisplayAlerts = False
            w.Range(w.Cells(r, imsSku), w.Cells(r + 7, imsSku)).Merge
            Application.DisplayAlerts = True
        End If
    End If
    If wasProt Then
        w.Protect DrawingObjects:=False, Contents:=True, Scenarios:=False, _
            AllowFormattingCells:=True, AllowFormattingColumns:=True, AllowFormattingRows:=True, _
            AllowInsertingColumns:=True, AllowInsertingRows:=True, AllowInsertingHyperlinks:=True, _
            AllowDeletingColumns:=True, AllowDeletingRows:=True, AllowSorting:=True, _
            AllowFiltering:=True, AllowUsingPivotTables:=True
    End If
End Sub

' Finds the IMS Stock columns from its header row (the row with "Stock Cover Days").
Private Function ReadImsLayout(w As Worksheet) As Boolean
    Dim r As Long, c As Long, lastCol As Long, v As Variant, m As Long, missing As String
    imsHdr = 0: imsCode = 0: imsLabel = 0: imsSku = 0: imsFg = 0: imsFormat = 0: imsFirst = 0
    For m = 1 To 12
        imsM26(m) = 0
        imsM27(m) = 0
    Next m
    If w Is Nothing Then Exit Function
    lastCol = w.UsedRange.Column + w.UsedRange.Columns.Count - 1
    For r = 1 To 30
        For c = 1 To lastCol
            If IsText(w.Cells(r, c).Value, "Stock Cover Days") Then
                imsHdr = r
                Exit For
            End If
        Next c
        If imsHdr > 0 Then Exit For
    Next r
    If imsHdr = 0 Then
        Problem w.Name, "Header row not found (no cell 'Stock Cover Days' in rows 1-30)."
        Exit Function
    End If
    For c = 1 To lastCol
        v = w.Cells(imsHdr, c).Value
        If IsText(v, "code") Then imsCode = c
        If IsText(v, "SKU") Then imsSku = c
        If IsText(v, "Stock Cover Days") Then imsLabel = c
        If IsText(v, "F-Group") Then imsFg = c
        If IsText(v, "Format") Then imsFormat = c
        If VarType(v) = vbDate Then
            If imsFirst = 0 Then imsFirst = c
            If Year(v) = 2026 Then imsM26(Month(v)) = c
            If Year(v) = 2027 Then imsM27(Month(v)) = c
        End If
    Next c
    If imsCode = 0 Then missing = missing & " 'code'"
    For m = 1 To 12
        If imsM26(m) = 0 Then missing = missing & " " & MonthName(m, True) & "-2026"
        If imsM27(m) = 0 Then missing = missing & " " & MonthName(m, True) & "-2027"
    Next m
    If Len(missing) > 0 Then
        Problem w.Name, "Header row " & imsHdr & " is missing:" & missing
        Exit Function
    End If
    ReadImsLayout = True
End Function

' Offset (0-7) of the row with this label in the block that starts at row t; -1 if absent
Private Function BlockRow(w As Worksheet, ByVal t As Long, ByVal label As String) As Long
    Dim i As Long
    BlockRow = -1
    For i = 0 To 7
        If IsText(w.Cells(t + i, imsLabel).Value, label) Then
            BlockRow = i
            Exit Function
        End If
    Next i
End Function

' codeCol: column of the SKU code; monthCol: column of January; planIdx: row of the plan grid
Private Sub Step_Fayrouz(ByVal sheetName As String, ByVal codeCol As Long, ByVal monthCol As Long, ByVal planIdx As Long)
    Dim w As Worksheet, t As Long, r As Long, src As Long, lc As Long, m As Long, c As Long
    Set w = Sh(sheetName)
    t = FindKey(w, codeCol, sTpl, 3, 2000)
    r = LastFilled(w, codeCol, 3, 2000) + 1
    src = t
    If src = 0 Then
        src = r - 1
        Warn w.Name, "Template SKU not found here; layout copied from row " & src & "."
    End If
    Note w.Name, "row " & r, "Monthly figures from the plan grid (copied from row " & src & ")"
    If dryRun Then Exit Sub
    lc = w.Cells(src, w.Columns.Count).End(xlToLeft).Column
    If lc < monthCol + 11 Then lc = monthCol + 11
    CopyCells w.Range(w.Cells(src, 1), w.Cells(src, lc)), w.Range(w.Cells(r, 1), w.Cells(r, lc))
    w.Cells(r, codeCol).Value = sCode
    w.Cells(r, codeCol + 1).Value = sDesc
    If t = 0 Then
        w.Cells(r, codeCol - 3).Value = sCat
        w.Cells(r, codeCol - 2).Value = sFg
    End If
    For m = 1 To 12
        c = monthCol + m - 1
        If Not w.Cells(r, c).HasFormula Then
            If Given(plan(planIdx, m)) Then
                w.Cells(r, c).Value = Num(plan(planIdx, m))
            Else
                w.Cells(r, c).ClearContents
            End If
        End If
    Next m
End Sub

Private Sub Step_DSCR(ByVal sheetName As String, ByVal yr As Long)
    Dim w As Worksheet, t As Long, r As Long, k As Long, row As Long, src As Long, c As Long
    Dim custCodes As Variant, custNames As Variant
    Set w = Sh(sheetName)
    If w Is Nothing Then Exit Sub
    custCodes = Array(SD_CODE, MW_CODE)
    custNames = Array(SD_NAME, MW_NAME)
    r = LastFilled(w, 4, 2, 5000) + 1
    For k = 0 To 1
        row = r + k
        t = FindText(w, 1, custCodes(k) & Format(sTpl, "0"), 2, 5000)
        src = t
        If src = 0 Then
            src = row - 1
            Warn w.Name, custNames(k) & ": template row not found; DSCR % left blank."
        End If
        Note w.Name, "row " & row, custNames(k) & " DSCR % (copied from row " & src & ")"
        If Not dryRun Then
            CopyCells w.Range("A" & src & ":Q" & src), w.Range("A" & row & ":Q" & row)
            w.Range("A" & row).Formula = "=B" & row & "&D" & row
            w.Range("B" & row).Value = custCodes(k)
            w.Range("C" & row).Value = custNames(k)
            w.Range("D" & row).Value = sCode
            w.Range("E" & row).Value = sDesc
            For c = 6 To 17                                             ' F:Q = Jan..Dec
                If Given(dscrPct(k + 1, yr)) Then
                    w.Cells(row, c).Value = Num(dscrPct(k + 1, yr))
                ElseIf t = 0 Then
                    w.Cells(row, c).ClearContents
                ElseIf w.Cells(row, c).HasFormula Then
                    w.Cells(row, c).Value = w.Cells(t, c).Value          ' template linked to another row: keep its value
                End If
            Next c
        End If
    Next k
End Sub

Private Sub Step_TargetSplit(ByVal sheetName As String, ByVal yr As Long)
    Dim w As Worksheet, t As Long, r As Long, k As Long, row As Long, src As Long, c As Long
    Dim custCodes As Variant, custNames As Variant
    Set w = Sh(sheetName)
    custCodes = Array(SD_CODE, MW_CODE)
    custNames = Array(SD_NAME, MW_NAME)
    r = LastFilled(w, 6, 3, 5000) + 1
    For k = 0 To 1
        row = r + k
        t = FindText(w, 1, custNames(k) & Format(sTpl, "0"), 3, 5000)
        src = t
        If src = 0 Then src = row - 1
        Note w.Name, "row " & row, custNames(k) & " share " & Format(Num(custSplit(k + 1, yr)), "0%") & " (copied from row " & src & ")"
        If Not dryRun Then
            CopyCells w.Range("A" & src & ":AH" & src), w.Range("A" & row & ":AH" & row)
            w.Range("A" & row).Formula = "=B" & row & "&F" & row
            w.Range("B" & row).Value = custNames(k)
            w.Range("C" & row).Value = custCodes(k)
            w.Range("D" & row).Value = sCat
            w.Range("E" & row).Value = sFg
            w.Range("F" & row).Value = sCode
            w.Range("G" & row).Value = sDesc
            w.Range("H" & row).Value = Num(custSplit(k + 1, yr))
            If yr = 1 Then
                w.Range("I" & row).Value = Num(custSplit(k + 1, yr))
            Else
                w.Range("I" & row).Formula = "=H" & row
            End If
            For c = 10 To 21                                            ' J:U = months 1..12
                w.Cells(row, c).Formula = "=$H" & row
            Next c
        End If
    Next k
End Sub

Private Sub Step_EngineTable(ByVal sheetName As String, ByVal tableName As String)
    Dim lo As ListObject, keyIdx As Long, a As Variant, i As Long, n As Long, k As Long
    Dim tplRows As New Collection, nr As ListRow, src As Range
    Set lo = Sh(sheetName).ListObjects(tableName)
    keyIdx = lo.ListColumns("product code").Index
    a = lo.ListColumns(keyIdx).DataBodyRange.Value
    For i = 1 To UBound(a, 1)
        If IsKey(a(i, 1), sTpl) Then tplRows.Add i
    Next i
    n = tplRows.Count
    If n = 0 Then
        Problem sheetName, "The template SKU has no rows in " & tableName & "."
        Exit Sub
    End If
    If Not RowBelowEmpty(lo) Then
        Problem sheetName, "The row under " & tableName & " is not empty, so rows cannot be added safely."
        Exit Sub
    End If
    Note sheetName, tableName & " +" & n & " rows", "Customer x month rows (copied from the template SKU's " & n & " rows)"
    If dryRun Then Exit Sub
    For k = 1 To n
        Set src = lo.ListRows(tplRows(k)).Range
        Set nr = lo.ListRows.Add(AlwaysInsert:=False)
        CopyCells src, nr.Range
        SetCol lo, nr, "Main_Cat", sCat
        SetCol lo, nr, "CATDESC", sCatDesc
        SetCol lo, nr, "Fgroup Code", sFgId
        SetCol lo, nr, "FGROUP_DESC", sFg
        SetCol lo, nr, "Update F-Group", sUpdFg
        SetCol lo, nr, "SKUDESC", sUnified
        SetCol lo, nr, "Unifed", sUnified
        SetCol lo, nr, "product code", sCode
        SetCol lo, nr, "Product Desc", sDesc
    Next k
End Sub

' 2027_Comparision W:AH is the source of the pivots on 2027_Comparision, 2027 VS ICP,
' Ramadan and 2027_by Quarter. 2027_by Quarter AD:AN links to it row by row.
Private Sub Step_CompBlock()
    Dim w As Worksheet, q As Worksheet, lastC As Long, qLast As Long, i As Long, n As Long, k As Long, r As Long
    Dim tplRows() As Long
    Set w = Sh("2027_Comparision")
    Set q = Sh("2027_by Quarter")
    lastC = LastFilled(w, 29, 16, LastUsedRow(w))                     ' AC; not End(xlUp), the sheet is filtered
    ReDim tplRows(1 To 1)
    For i = 16 To lastC
        If IsText(w.Cells(i, 29).Value, sTplDesc) Then
            If IsText(w.Cells(i, 23).Value, "M") Or IsText(w.Cells(i, 23).Value, "M-1") Then
                n = n + 1
                ReDim Preserve tplRows(1 To n)
                tplRows(n) = i
            End If
        End If
    Next i
    If n = 0 Then
        Warn w.Name, "The template SKU is not in the data block W:AH; add its month rows by hand."
        Exit Sub
    End If
    qLast = LastFilled(q, 36, 10, LastUsedRow(q))                     ' AJ
    If Not RowsEmpty(w, lastC + 1, n, 22, 34) Then
        Problem w.Name, "Rows " & (lastC + 1) & "-" & (lastC + n) & " (V:AH) are not empty; nothing was written there."
        Exit Sub
    End If
    If Not RowsEmpty(q, qLast + 1, n, 30, 40) Then
        Problem q.Name, "Rows " & (qLast + 1) & "-" & (qLast + n) & " (AD:AN) are not empty; nothing was written there."
        Exit Sub
    End If
    Note w.Name, "rows " & (lastC + 1) & "-" & (lastC + n), n & " month rows in the data block W:AH (pivot source)"
    Note q.Name, "rows " & (qLast + 1) & "-" & (qLast + n), "Helper rows linked to the new 2027_Comparision rows"
    If dryRun Then Exit Sub
    For k = 1 To n
        r = lastC + k
        If w.Cells(tplRows(k), 22).ListObject Is Nothing Then
            CopyCells w.Range(w.Cells(tplRows(k), 22), w.Cells(tplRows(k), 34)), w.Range(w.Cells(r, 22), w.Cells(r, 34))
        Else                                                            ' column V belongs to Table20 in the top rows
            CopyCells w.Range(w.Cells(tplRows(k), 23), w.Cells(tplRows(k), 34)), w.Range(w.Cells(r, 23), w.Cells(r, 34))
        End If
        w.Cells(r, 27).Value = sCat
        w.Cells(r, 28).Value = sUpdFg
        w.Cells(r, 29).Value = sDesc
        CopyCells q.Range(q.Cells(qLast, 30), q.Cells(qLast, 40)), q.Range(q.Cells(qLast + k, 30), q.Cells(qLast + k, 40))
        q.Cells(qLast + k, 31).Formula = "='2027_Comparision'!X" & r
        q.Cells(qLast + k, 32).Formula = "='2027_Comparision'!Y" & r
        q.Cells(qLast + k, 34).Formula = "='2027_Comparision'!AA" & r
        q.Cells(qLast + k, 35).Formula = "='2027_Comparision'!AB" & r
        q.Cells(qLast + k, 36).Formula = "='2027_Comparision'!AC" & r
    Next k
End Sub

' Report tables with category sections and subtotal rows: insert inside the section,
' above its last SKU row, so the section's SUM range grows to include the new row.
Private Sub Step_ReportTable(ByVal sheetName As String, ByVal tableName As String)
    Dim lo As ListObject, a As Variant, i As Long, idx As Long, nr As ListRow
    Set lo = Sh(sheetName).ListObjects(tableName)
    a = lo.ListColumns(1).DataBodyRange.Value
    For i = 1 To UBound(a, 1)
        If IsText(a(i, 1), sCat) Then idx = i
    Next i
    If idx = 0 Then
        Problem sheetName, "Category '" & sCat & "' was not found in " & tableName & "."
        Exit Sub
    End If
    If Not RowBelowEmpty(lo) Then
        Problem sheetName, "The row under " & tableName & " is not empty, so a row cannot be added safely. Add the SKU by hand."
        Exit Sub
    End If
    Note sheetName, tableName & " row " & (lo.HeaderRowRange.Row + idx), "Report row in the " & sCat & " section"
    If dryRun Then Exit Sub
    Set nr = lo.ListRows.Add(Position:=idx, AlwaysInsert:=False)
    CopyCells lo.ListRows(idx + 1).Range, nr.Range
    nr.Range.Cells(1, 1).Value = sCat
    nr.Range.Cells(1, 2).Value = sUpdFg
    nr.Range.Cells(1, 3).Value = sDesc
End Sub

' Monthly Gap Table612 reads the pivot list in CD:CE row by row; one more SKU needs one more row.
Private Sub Step_MonthlyGap()
    Dim lo As ListObject, nr As ListRow
    Set lo = Sh("Monthly Gap").ListObjects("Table612")
    If Not RowBelowEmpty(lo) Then
        Problem "Monthly Gap", "The row under Table612 is not empty; add one table row by hand."
        Exit Sub
    End If
    Note "Monthly Gap", "Table612 +1 row", "One more SKU row (reads the pivot list in CD:CE after refresh)"
    If dryRun Then Exit Sub
    Set nr = lo.ListRows.Add(AlwaysInsert:=False)
    CopyCells lo.ListRows(lo.ListRows.Count - 1).Range, nr.Range
End Sub

' CYvsNY Table9 has two rows (2026, 2027) per SKU, read from the pivot in AK:AN.
Private Sub Step_CYvsNY()
    Dim lo As ListObject, need As Long, k As Long, nr As ListRow
    Set lo = Sh("CYvsNY").ListObjects("Table9")
    need = 2 * (mRow - 1)
    If lo.ListRows.Count >= need Then
        Note "CYvsNY", "Table9", "Has room for " & (lo.ListRows.Count \ 2) & " SKUs; nothing to add"
        Exit Sub
    End If
    If Not RowBelowEmpty(lo) Then
        Problem "CYvsNY", "Table9 needs 2 more rows but the row under it is not empty; add them by hand."
        Exit Sub
    End If
    Note "CYvsNY", "Table9 +2 rows", "Room for one more SKU (2026 and 2027 rows)"
    If dryRun Then Exit Sub
    For k = 1 To 2
        Set nr = lo.ListRows.Add(AlwaysInsert:=False)
        CopyCells lo.ListRows(lo.ListRows.Count - 2).Range, nr.Range
    Next k
End Sub

'==============================================================================
' The New SKU sheet
'==============================================================================

Private Sub BuildForm(f As Worksheet)
    Dim r As Long, i As Long, labels As Variant, bdaCols As Variant, planLabels As Variant
    Dim head As Long
    head = RGB(221, 235, 247)

    f.Cells.Font.Name = "Calibri"
    f.Cells.Font.Size = 10
    f.Columns("A").ColumnWidth = 2
    f.Columns("B").ColumnWidth = 48
    f.Columns("C:D").ColumnWidth = 16
    f.Columns("E").ColumnWidth = 30
    f.Columns("F").ColumnWidth = 14
    f.Columns("G:H").ColumnWidth = 12
    f.Columns("I:K").ColumnWidth = 10
    f.Columns("L").ColumnWidth = 26
    f.Columns("M").ColumnWidth = 22
    f.Columns("N").ColumnWidth = 80

    f.Range("B1").Value = "New SKU"
    f.Range("B1").Font.Size = 16
    f.Range("B1").Font.Bold = True
    f.Range("B2").Value = "1. Fill the yellow cells. Blank cells use the template SKU's value (grey column)."
    f.Range("B3").Value = "2. Click Preview and read the log (column L).  3. Click Add SKU.  4. Check that every line of the checklist says OK."

    ' ---- identity and prices ----
    HeaderRow f, 4, Array("Field", "Your input", "Template SKU", "Will use", "")
    f.Range("F4").Value = "Notes"
    Field f, 5, "NS_Code", "New SKU code", "", "=C5", "8-digit code, as in Sales Vs Activity-RAW column E."
    Field f, 6, "NS_Template", "Copy settings from SKU code", TplMaster("C"), "=C6", "An existing SKU in the same F-Group. Its rows are copied in every sheet."
    Field f, 7, "NS_Desc", "SKU description", "", "=C7", "Exactly as the system will show it. Reports match on this text."
    Field f, 8, "NS_Short", "Short description", "=C7", "", "SKU mapping-Pricing column E."
    Field f, 9, "NS_Unified", "Unified description (SKUDESC)", TplMaster("A"), "", "Groups old and new versions of a product (FCST column E)."
    Field f, 10, "NS_Mosaic", "Mosaic product code", "", "", "Needed for CYvsNY and the Mosaic sheets."
    Field f, 11, "NS_Category", "Category (Main_Cat)", TplMaster("L"), "", "Coffee, Confectionary, Dairy, Food or Nutrition."
    Field f, 12, "NS_CatDesc", "CATDESC", "=IFERROR(INDEX(FCST!$A:$A,MATCH($C$6,FCST!$C:$C,0)),"""")", "", "For example 01 Ambient Dairy."
    Field f, 13, "NS_FGroupID", "F-Group ID", TplMaster("F"), "", "For example E7J."
    Field f, 14, "NS_FGroup", "F-Group", TplMaster("G"), "", ""
    Field f, 15, "NS_UpdFGroup", "Update F-Group (report group)", "=IFERROR(INDEX('Consolidation RAW'!$M:$M,MATCH($C$6,'Consolidation RAW'!$P:$P,0)),"""")", "", "Used by Table1 / Table19 and the 2027 reports."
    Field f, 16, "NS_Status", "Status", "=""New""", "", "New or Current."
    Field f, 17, "NS_Active", "Active", "=1", "", "1 = active, 0 = not active."
    Field f, 18, "NS_Format", "Tons per case", TplMaster("P"), "=C18", "Units x grams / 1,000,000. 12 x 900 g = 0.0108. Must be typed."
    Field f, 19, "NS_1C", "1C description", "=C7", "", "SKU mapping-Pricing column Z."
    Field f, 20, "NS_Arabic", "Arabic description", "", "", "SKU mapping-Pricing column AA."
    Field f, 21, "NS_IMSLabel", "IMS Stock label", "=C7", "", "IMS Stock 'SKU' column."
    Field f, 22, "NS_WSL", "WSL price per case (SYP)", TplMaster("I"), "=C22", "Must be typed. Template price shown for reference."
    Field f, 23, "NS_Ret", "Retail price per case (SYP)", TplMaster("J"), "=C23", "Must be typed. Template price shown for reference."
    Field f, 24, "NS_Refresh", "Refresh pivots when done", "=""Yes""", "", "Yes or No. Runs Refresh_all."
    f.Range("C18,D18,E18").NumberFormat = "0.0000"
    f.Range("C22:E23").NumberFormat = "#,##0"
    AddList f.Range("C6"), "='SKU mapping-Pricing'!$B$2:$B$500"
    AddList f.Range("C11"), "Coffee,Confectionary,Dairy,Food,Nutrition"
    AddList f.Range("C16"), "New,Current"
    AddList f.Range("C17"), "1,0"
    AddList f.Range("C24"), "Yes,No"

    ' ---- trade terms ----
    HeaderRow f, 26, Array("Trade terms (share of NPS)", "2026", "2027", "Template 2026", "Template 2027", "Will use 2026", "Will use 2027")
    labels = Array("Progressive, Sub-Distributor (SD)", "Progressive, Medical WH (MDD)", _
                   "Coverage, Sub-Distributor", "Coverage, Medical WH", _
                   "Commission, Sub-Distributor", "Commission, Medical WH", _
                   "Distribution, Sub-Distributor", "Distribution, Medical WH")
    bdaCols = Array("E", "F", "G", "H", "I", "J", "M", "N")
    For i = 0 To 7
        r = 27 + i
        f.Cells(r, 2).Value = labels(i)
        InputCell f.Range("C" & r & ":D" & r)
        f.Cells(r, 5).Formula = "=IFERROR(INDEX('BDA Allowance Assist'!$" & bdaCols(i) & ":$" & bdaCols(i) & ",MATCH($C$6,'BDA Allowance Assist'!$C:$C,0)),0)"
        f.Cells(r, 6).Formula = "=IFERROR(INDEX('BDA Allowance Assist Y+1'!$" & bdaCols(i) & ":$" & bdaCols(i) & ",MATCH($C$6,'BDA Allowance Assist Y+1'!$C:$C,0)),0)"
        f.Cells(r, 7).Formula = "=IF(C" & r & "<>"""",C" & r & ",E" & r & ")"
        f.Cells(r, 8).Formula = "=IF(D" & r & "<>"""",D" & r & ",F" & r & ")"
    Next i
    f.Range("C27:H34").NumberFormat = "0.0%"
    f.Range("E27:F34").Font.Color = RGB(118, 118, 118)
    ThisWorkbook.Names.Add Name:="NS_Terms", RefersTo:="='" & FORM_NAME & "'!$G$27:$H$34"

    ' ---- customer split ----
    HeaderRow f, 36, Array("Customer split of volume", "2026", "2027", "Template 2026", "Template 2027", "Will use 2026", "Will use 2027")
    SplitRow f, 37, SD_NAME
    SplitRow f, 38, MW_NAME
    f.Range("B39").Value = "Total (must be 100%)"
    f.Range("G39").Formula = "=SUM(G37:G38)"
    f.Range("H39").Formula = "=SUM(H37:H38)"
    f.Range("B39:H39").Font.Bold = True
    f.Range("C37:H39").NumberFormat = "0%"
    ThisWorkbook.Names.Add Name:="NS_Split", RefersTo:="='" & FORM_NAME & "'!$G$37:$H$38"

    ' ---- DSCR ----
    HeaderRow f, 41, Array("DSCR % (blank = copy the template's monthly %)", "2026", "2027", "Template 2026 avg", "Template 2027 avg")
    DscrRow f, 42, SD_NAME, SD_CODE
    DscrRow f, 43, MW_NAME, MW_CODE
    f.Range("C42:F43").NumberFormat = "0.0%"
    ThisWorkbook.Names.Add Name:="NS_DSCR", RefersTo:="='" & FORM_NAME & "'!$C$42:$D$43"

    ' ---- monthly plan ----
    HeaderRow f, 45, Array("Monthly plan (cases)", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
    f.Range("O45").Value = "Total"
    f.Range("O45").Font.Bold = True
    planLabels = Array("Volume plan 2026 (IMS Stock 'Nestle Syria Target'; drives FCST)", _
                       "Volume plan 2027 (drives FCST Y+1)", _
                       "Arrivals 2026 (Fayrouz-Arrival)", _
                       "Arrivals 2027 (Fayrouz-Arrival-2027)", _
                       "TPS target 2026 (Fayrouz Target)", _
                       "TPS target 2027 (Fayrouz 2027-Target)")
    For i = 0 To 5
        r = 46 + i
        f.Cells(r, 2).Value = planLabels(i)
        InputCell f.Range("C" & r & ":N" & r)
        f.Cells(r, 15).Formula = "=SUM(C" & r & ":N" & r & ")"
    Next i
    f.Range("C46:O51").NumberFormat = "#,##0"
    ThisWorkbook.Names.Add Name:="NS_Plan", RefersTo:="='" & FORM_NAME & "'!$C$46:$N$51"
    f.Range("B53").Value = "Stock on hand before the first month (cases)"
    InputCell f.Range("C53")
    f.Range("C53").NumberFormat = "#,##0"
    ThisWorkbook.Names.Add Name:="NS_OpenStock", RefersTo:="='" & FORM_NAME & "'!$C$53"

    BuildChecks f, 56

    ' ---- log ----
    f.Range("L4").Value = "Sheet"
    f.Range("M4").Value = "Where"
    f.Range("N4").Value = "What"
    f.Range("L4:N4").Font.Bold = True
    f.Range("L4:N4").Interior.Color = head
    f.Range("L3").Font.Bold = True

    ' ---- buttons ----
    AddButton f, f.Range("G1:H2"), "Preview", "NewSKU_Preview"
    AddButton f, f.Range("I1:J2"), "Add SKU", "NewSKU_Add"
    AddButton f, f.Range("L1:M2"), "Widen lookup ranges (one time)", "NewSKU_WidenLookupRanges"
End Sub

Private Sub BuildChecks(f As Worksheet, ByVal r0 As Long)
    Dim r As Long, items As Collection, it As Variant, imsCol As String
    Set items = New Collection
    ' label | found formula | expected formula | note
    items.Add Array("SKU mapping-Pricing", "=COUNTIF('SKU mapping-Pricing'!$B:$B,$E$5)", "1", "")
    items.Add Array("Price Chnge 2026 (rows 6-55)", "=COUNTIF('Price Chnge'!$A$6:$A$55,$E$5)", "1", "")
    items.Add Array("Price Chnge 2027 (rows 58+)", "=COUNTIF('Price Chnge'!$A$58:$A$400,$E$5)", "1", "")
    items.Add Array("BDA Allowance Assist", "=COUNTIF('BDA Allowance Assist'!$C$4:$C$500,$E$5)", "1", "")
    items.Add Array("BDA Allowance Assist Y+1", "=COUNTIF('BDA Allowance Assist Y+1'!$C$4:$C$500,$E$5)", "1", "")
    items.Add Array("Distribution by Month", "=COUNTIF('Distribution by Month'!$C$4:$C$500,$E$5)", "1", "")
    items.Add Array("Distribution by Month Y+1", "=COUNTIF('Distribution by Month Y+1'!$C$4:$C$500,$E$5)", "1", "")
    items.Add Array("FCST", "=COUNTIF(FCST!$C$4:$C$500,$E$5)", "1", "")
    items.Add Array("FCST Y+1", "=COUNTIF('FCST Y+1'!$C$4:$C$500,$E$5)", "1", "")
    If ReadImsLayout(Sh("IMS Stock")) Then imsCol = ColL(imsCode) Else imsCol = "C"
    items.Add Array("IMS Stock (8-row block)", "=SUMPRODUCT(--('IMS Stock'!$" & imsCol & "$" & (imsHdr + 1) & ":$" & imsCol & "$5000&""""=$E$5&""""))", "8", "")
    items.Add Array("Fayrouz-Arrival", "=SUMPRODUCT(--('Fayrouz-Arrival'!$E$3:$E$500&""""=$E$5&""""))", "1", "Feeds IMS Stock arrivals 2026")
    items.Add Array("Fayrouz-Arrival-2027", "=SUMPRODUCT(--('Fayrouz-Arrival-2027'!$E$3:$E$500&""""=$E$5&""""))", "1", "Feeds IMS Stock arrivals 2027")
    items.Add Array("Fayrouz Target", "=SUMPRODUCT(--('Fayrouz Target'!$D$3:$D$500&""""=$E$5&""""))", "1", "")
    items.Add Array("Fayrouz 2027-Target", "=SUMPRODUCT(--('Fayrouz 2027-Target'!$E$3:$E$500&""""=$E$5&""""))", "1", "")
    items.Add Array("DSCR (2 customers)", "=COUNTIF(DSCR!$D:$D,$E$5)", "2", "")
    items.Add Array("DSCR -LS (2 customers)", "=COUNTIF('DSCR -LS'!$D:$D,$E$5)", "2", "")
    items.Add Array("DSCR Y+1 (2 customers)", "=COUNTIF('DSCR Y+1'!$D:$D,$E$5)", "2", "")
    items.Add Array("Target Split (2 customers)", "=COUNTIF('Target Split'!$F:$F,$E$5)", "2", "")
    items.Add Array("Target Split shares add up to 100%", "=SUMIF('Target Split'!$F:$F,$E$5,'Target Split'!$H:$H)", "1", "")
    items.Add Array("Target Split Y+1 (2 customers)", "=COUNTIF('Target Split Y+1'!$F:$F,$E$5)", "2", "")
    items.Add Array("Target Split Y+1 shares add up to 100%", "=SUMIF('Target Split Y+1'!$F:$F,$E$5,'Target Split Y+1'!$H:$H)", "1", "")
    items.Add Array("Consolidation RAW (Table1) rows", "=COUNTIF('Consolidation RAW'!$P:$P,$E$5)", "=COUNTIF('Consolidation RAW'!$P:$P,$E$6)", "Same number of rows as the template SKU")
    items.Add Array("Y+1 (Table19) rows", "=COUNTIF('Y+1'!$P:$P,$E$5)", "=COUNTIF('Y+1'!$P:$P,$E$6)", "Same number of rows as the template SKU")
    items.Add Array("Errors in this SKU's Table1 rows", "=SUMPRODUCT(('Consolidation RAW'!$P$4:$P$5000=$E$5)*ISERROR('Consolidation RAW'!$R$4:$AN$5000))", "0", "A lookup does not reach the new row: run Widen lookup ranges")
    items.Add Array("Errors in this SKU's Table19 rows", "=SUMPRODUCT(('Y+1'!$P$4:$P$5000=$E$5)*ISERROR('Y+1'!$R$4:$AN$5000))", "0", "A lookup does not reach the new row: run Widen lookup ranges")
    items.Add Array("Y+1 Active flag", "=SUMIF('Y+1'!$P:$P,$E$5,'Y+1'!$AP:$AP)", "=COUNTIF('Y+1'!$P:$P,$E$5)", "0 means the Active lookup stops at row 25: run Widen lookup ranges")
    items.Add Array("2027_Comparision data block (W:AH)", "=COUNTIF('2027_Comparision'!$AC$16:$AC$5000,$E$7)", "=COUNTIF('2027_Comparision'!$AC$16:$AC$5000,$D$6)", "Same number of month rows as the template SKU")
    items.Add Array("2027_by Quarter helper rows", "=COUNTIF('2027_by Quarter'!$AJ$10:$AJ$5000,$E$7)", "=COUNTIF('2027_by Quarter'!$AJ$10:$AJ$5000,$D$6)", "")
    items.Add Array("2027_Comparision table", "=COUNTIF('2027_Comparision'!$C$16:$C$300,$E$7)", "1", "")
    items.Add Array("Ramadan table", "=COUNTIF(Ramadan!$C$16:$C$300,$E$7)", "1", "")
    items.Add Array("2027 VS ICP table", "=COUNTIF('2027 VS ICP'!$C$16:$C$300,$E$7)", "1", "")
    items.Add Array("Monthly Gap table (after refresh)", "=COUNTIF('Monthly Gap'!$C$32:$C$300,$E$7)", "1", "Appears once the pivots are refreshed")
    items.Add Array("2026 DF report (after refresh)", "=COUNTIF('2026 DF , NNS ,Tons,Cases '!$C:$C,$E$5)", "3", "Three pivots: NNS, Tons, Cases")
    items.Add Array("2027 OL report (after refresh)", "=COUNTIF('2027 OL'!$C:$C,$E$5)", "3", "Three pivots: NNS, Tons, Cases")
    items.Add Array("CYvsNY (after refresh)", "=IF($E$10="""",0,COUNTIF(CYvsNY!$E:$E,$E$10))", "2", "Needs the Mosaic product code")

    HeaderRow f, r0, Array("Checklist (live; works for SKUs added by hand too)", "Found", "Expected", "Status", "")
    r = r0 + 1
    For Each it In items
        f.Cells(r, 2).Value = it(0)
        f.Cells(r, 3).Formula = it(1)
        If Left$(it(2), 1) = "=" Then f.Cells(r, 4).Formula = it(2) Else f.Cells(r, 4).Value = CDbl(it(2))
        f.Cells(r, 5).Formula = "=IF(ISERROR(C" & r & "),""Error"",IF(ROUND(C" & r & "-D" & r & ",4)=0,""OK"",""Missing""))"
        f.Cells(r, 6).Value = it(3)
        r = r + 1
    Next it
    f.Cells(r0, 6).Formula = "=COUNTIF(E" & (r0 + 1) & ":E" & (r - 1) & ",""OK"")&"" of ""&ROWS(E" & (r0 + 1) & ":E" & (r - 1) & ")&"" OK"""
    With f.Range("E" & (r0 + 1) & ":E" & (r - 1))
        .FormatConditions.Delete
        .FormatConditions.Add(xlCellValue, xlEqual, "=""OK""").Interior.Color = RGB(198, 239, 206)
        .FormatConditions.Add(xlCellValue, xlNotEqual, "=""OK""").Interior.Color = RGB(255, 199, 206)
    End With
End Sub

Private Sub Field(f As Worksheet, ByVal r As Long, ByVal nm As String, ByVal label As String, _
                  ByVal tplFormula As String, ByVal useFormula As String, ByVal note As String)
    f.Cells(r, 2).Value = label
    InputCell f.Cells(r, 3)
    If tplFormula <> "" Then f.Cells(r, 4).Formula = tplFormula
    f.Cells(r, 4).Font.Color = RGB(118, 118, 118)
    If useFormula = "" Then useFormula = "=IF(C" & r & "<>"""",C" & r & ",IF(D" & r & "="""","""",D" & r & "))"
    f.Cells(r, 5).Formula = useFormula
    f.Cells(r, 6).Value = note
    f.Cells(r, 6).Font.Color = RGB(89, 89, 89)
    ThisWorkbook.Names.Add Name:=nm, RefersTo:="='" & FORM_NAME & "'!$E$" & r
End Sub

Private Sub SplitRow(f As Worksheet, ByVal r As Long, ByVal custName As String)
    f.Cells(r, 2).Value = custName
    InputCell f.Range("C" & r & ":D" & r)
    f.Cells(r, 5).Formula = "=IFERROR(INDEX('Target Split'!$H:$H,MATCH(""" & custName & """&$C$6,'Target Split'!$A:$A,0)),0)"
    f.Cells(r, 6).Formula = "=IFERROR(INDEX('Target Split Y+1'!$H:$H,MATCH(""" & custName & """&$C$6,'Target Split Y+1'!$A:$A,0)),0)"
    f.Cells(r, 7).Formula = "=IF(C" & r & "<>"""",C" & r & ",E" & r & ")"
    f.Cells(r, 8).Formula = "=IF(D" & r & "<>"""",D" & r & ",F" & r & ")"
    f.Range("E" & r & ":F" & r).Font.Color = RGB(118, 118, 118)
End Sub

Private Sub DscrRow(f As Worksheet, ByVal r As Long, ByVal custName As String, ByVal custCode As String)
    f.Cells(r, 2).Value = custName
    InputCell f.Range("C" & r & ":D" & r)
    f.Cells(r, 5).Formula = "=IFERROR(AVERAGE(INDEX(DSCR!$F:$Q,MATCH(""" & custCode & """&$C$6,DSCR!$A:$A,0),0)),0)"
    f.Cells(r, 6).Formula = "=IFERROR(AVERAGE(INDEX('DSCR Y+1'!$F:$Q,MATCH(""" & custCode & """&$C$6,'DSCR Y+1'!$A:$A,0),0)),0)"
    f.Range("E" & r & ":F" & r).Font.Color = RGB(118, 118, 118)
End Sub

Private Function TplMaster(ByVal col As String) As String
    TplMaster = "=IFERROR(INDEX('SKU mapping-Pricing'!$" & col & ":$" & col & ",MATCH($C$6,'SKU mapping-Pricing'!$B:$B,0)),"""")"
End Function

Private Sub HeaderRow(f As Worksheet, ByVal r As Long, ByVal headers As Variant)
    Dim i As Long
    For i = 0 To UBound(headers)
        f.Cells(r, 2 + i).Value = headers(i)
    Next i
    With f.Range(f.Cells(r, 2), f.Cells(r, 2 + UBound(headers)))
        .Font.Bold = True
        .Interior.Color = RGB(221, 235, 247)
    End With
End Sub

Private Sub InputCell(rng As Range)
    rng.Interior.Color = RGB(255, 242, 204)
    rng.Locked = False
    rng.Borders.Color = RGB(191, 191, 191)
End Sub

Private Sub AddList(rng As Range, ByVal listFormula As String)
    With rng.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=listFormula
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
End Sub

Private Sub AddButton(f As Worksheet, where As Range, ByVal caption As String, ByVal macroName As String)
    Dim b As Object
    Set b = f.Buttons.Add(where.Left, where.Top, where.Width, where.Height)
    b.Caption = caption
    b.OnAction = macroName
End Sub

'==============================================================================
' Helpers
'==============================================================================

Private Function Sh(ByVal nm As String) As Worksheet
    On Error Resume Next
    Set Sh = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
End Function

Private Function ColL(ByVal c As Long) As String
    Dim s As String
    Do While c > 0
        s = Chr$(65 + (c - 1) Mod 26) & s
        c = (c - 1) \ 26
    Loop
    ColL = s
End Function

Private Function IsKey(ByVal v As Variant, ByVal key As Double) As Boolean
    Dim s As String
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    s = Trim$(CStr(v))
    If Len(s) = 0 Then Exit Function
    If IsNumeric(s) Then IsKey = (CDbl(s) = key)
End Function

Private Function IsText(ByVal v As Variant, ByVal t As String) As Boolean
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    IsText = (StrComp(Trim$(CStr(v)), Trim$(t), vbTextCompare) = 0)
End Function

Private Function HasValue(ByVal v As Variant) As Boolean
    Dim s As String
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    s = Trim$(CStr(v))
    HasValue = (Len(s) > 0 And s <> "0")
End Function

' True when the user typed something (0 counts as typed)
Private Function Given(ByVal v As Variant) As Boolean
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    Given = (Len(Trim$(CStr(v))) > 0)
End Function

Private Function Num(ByVal v As Variant) As Double
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    If IsNumeric(v) Then Num = CDbl(v)
End Function

Private Function ColValues(w As Worksheet, ByVal col As Long, ByVal r1 As Long, ByVal r2 As Long) As Variant
    Dim a As Variant
    If r2 <= r1 Then
        ReDim a(1 To 1, 1 To 1)
        a(1, 1) = w.Cells(r1, col).Value
        ColValues = a
    Else
        ColValues = w.Range(w.Cells(r1, col), w.Cells(r2, col)).Value
    End If
End Function

Private Function FindKey(w As Worksheet, ByVal col As Long, ByVal key As Double, ByVal r1 As Long, ByVal r2 As Long) As Long
    Dim a As Variant, i As Long
    If w Is Nothing Or key = 0 Then Exit Function
    a = ColValues(w, col, r1, r2)
    For i = 1 To UBound(a, 1)
        If IsKey(a(i, 1), key) Then
            FindKey = r1 + i - 1
            Exit Function
        End If
    Next i
End Function

Private Function FindText(w As Worksheet, ByVal col As Long, ByVal t As String, ByVal r1 As Long, ByVal r2 As Long) As Long
    Dim a As Variant, i As Long
    If w Is Nothing Then Exit Function
    a = ColValues(w, col, r1, r2)
    For i = 1 To UBound(a, 1)
        If IsText(a(i, 1), t) Then
            FindText = r1 + i - 1
            Exit Function
        End If
    Next i
End Function

' Last row in r1..r2 of a column holding a value other than blank or 0
Private Function LastFilled(w As Worksheet, ByVal col As Long, ByVal r1 As Long, ByVal r2 As Long) As Long
    Dim a As Variant, i As Long
    a = ColValues(w, col, r1, r2)
    For i = UBound(a, 1) To 1 Step -1
        If HasValue(a(i, 1)) Then
            LastFilled = r1 + i - 1
            Exit Function
        End If
    Next i
    LastFilled = r1 - 1
End Function

' Copies cells one by one: formulas as relative R1C1 (so they point at the new row),
' constants as values, and number formats. Works on filtered sheets, unlike Copy/Paste.
Private Sub CopyCells(src As Range, dst As Range)
    Dim i As Long, s As Range, d As Range
    For i = 1 To src.Cells.Count
        Set s = src.Cells(i)
        Set d = dst.Cells(i)
        If s.HasFormula Then
            d.FormulaR1C1 = s.FormulaR1C1
        Else
            d.Value = s.Value
        End If
        d.NumberFormat = s.NumberFormat
    Next i
End Sub

Private Sub SetCol(lo As ListObject, nr As ListRow, ByVal header As String, ByVal v As Variant)
    Dim idx As Long, c As Range
    On Error Resume Next
    idx = lo.ListColumns(header).Index
    On Error GoTo 0
    If idx = 0 Then Exit Sub
    Set c = nr.Range.Cells(1, idx)
    If Not c.HasFormula Then c.Value = v                ' columns that are formulas work the value out themselves
End Sub

' Last row of the used range. Unlike End(xlUp), it does not depend on filters or hidden rows.
Private Function LastUsedRow(w As Worksheet) As Long
    LastUsedRow = w.UsedRange.Row + w.UsedRange.Rows.Count - 1
End Function

' True when rows r..r+n-1, columns c1..c2, hold nothing (no values and no formulas)
Private Function RowsEmpty(w As Worksheet, ByVal r As Long, ByVal n As Long, ByVal c1 As Long, ByVal c2 As Long) As Boolean
    RowsEmpty = (Application.WorksheetFunction.CountA(w.Range(w.Cells(r, c1), w.Cells(r + n - 1, c2))) = 0)
End Function

Private Function RowBelowEmpty(lo As ListObject) As Boolean
    RowBelowEmpty = (Application.WorksheetFunction.CountA(lo.Range.Rows(lo.Range.Rows.Count).Offset(1, 0)) = 0)
End Function

' Last row of the range that a table column's first formula uses on refSheet ($C$1:$C$45 -> 45).
' 0 when the reference is a whole column or is not found.
Private Function RefLimit(ByVal sheetName As String, ByVal tableName As String, _
                          ByVal header As String, ByVal refSheet As String) As Long
    Dim fml As String, k As Long, tail As String, p As Long, ch As String, digits As String
    On Error GoTo Done
    fml = Sh(sheetName).ListObjects(tableName).ListColumns(header).DataBodyRange.Cells(1, 1).Formula
    k = InStr(1, fml, "'" & refSheet & "'!", vbTextCompare)
    If k > 0 Then
        k = k + Len(refSheet) + 3
    Else
        k = InStr(1, fml, refSheet & "!", vbTextCompare)
        If k = 0 Then Exit Function
        k = k + Len(refSheet) + 1
    End If
    tail = Mid$(fml, k)
    p = 1
    Do While p <= Len(tail)
        ch = Mid$(tail, p, 1)
        If Not (ch Like "[A-Za-z0-9$:]") Then Exit Do
        p = p + 1
    Loop
    tail = Left$(tail, p - 1)
    If InStr(tail, ":") = 0 Then Exit Function
    tail = Mid$(tail, InStr(tail, ":") + 1)
    For p = Len(tail) To 1 Step -1
        ch = Mid$(tail, p, 1)
        If ch Like "[0-9]" Then digits = ch & digits Else Exit For
    Next p
    If Len(digits) > 0 Then RefLimit = CLng(digits)
Done:
End Function

Private Function Widen(ByVal sheetName As String, ByVal findText As String, ByVal replText As String) As Long
    Dim w As Worksheet, fcells As Range, c As Range, fml As String, n As Long, wasProt As Boolean
    Set w = Sh(sheetName)
    If w Is Nothing Then Exit Function
    On Error Resume Next
    Set fcells = w.UsedRange.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0
    If fcells Is Nothing Then Exit Function
    wasProt = w.ProtectContents
    If wasProt Then w.Unprotect
    For Each c In fcells.Cells
        fml = c.Formula
        If InStr(1, fml, findText, vbTextCompare) > 0 Then
            If c.HasArray Then
                Note sheetName, c.Address(False, False), "Array formula not changed; widen by hand: " & findText
            Else
                c.Formula = Replace(fml, findText, replText, 1, -1, vbTextCompare)
                n = n + 1
            End If
        End If
    Next c
    If wasProt Then w.Protect Contents:=True, AllowFormattingCells:=True, AllowFormattingColumns:=True, _
        AllowFormattingRows:=True, AllowInsertingColumns:=True, AllowInsertingRows:=True, _
        AllowInsertingHyperlinks:=True, AllowDeletingColumns:=True, AllowDeletingRows:=True, _
        AllowSorting:=True, AllowFiltering:=True, AllowUsingPivotTables:=True
    If n > 0 Then Note sheetName, n & " formulas", findText & "  ->  " & replText
    Widen = n
End Function

Private Function SaveBackup(ByVal reason As String) As Boolean
    Dim p As String, base As String, ext As String
    On Error GoTo Failed
    If ThisWorkbook.Path = "" Then Exit Function
    base = ThisWorkbook.Name
    ext = ".xlsm"
    If InStrRev(base, ".") > 0 Then
        ext = Mid$(base, InStrRev(base, "."))
        base = Left$(base, InStrRev(base, ".") - 1)
    End If
    p = ThisWorkbook.Path & Application.PathSeparator & base & " - " & reason & " " & Format(Now, "yyyy-mm-dd hhnn") & ext
    ThisWorkbook.SaveCopyAs p
    Note "Backup", "", p
    SaveBackup = True
    Exit Function
Failed:
    Warn "Backup", "Could not save a backup copy: " & Err.Description
End Function

Private Sub RefreshPivots()
    On Error Resume Next
    Application.Run "'" & ThisWorkbook.Name & "'!Refresh_all"
    If Err.Number <> 0 Then
        Err.Clear
        ThisWorkbook.RefreshAll
    End If
    On Error GoTo 0
    Note "Pivots", "", "Refreshed"
End Sub

Private Function InVal(ByVal nm As String) As Variant
    On Error GoTo NoName
    InVal = ThisWorkbook.Names(nm).RefersToRange.Value
    Exit Function
NoName:
    InVal = Empty
End Function

Private Function InText(ByVal nm As String) As String
    Dim v As Variant
    v = InVal(nm)
    If IsError(v) Or IsEmpty(v) Then Exit Function
    InText = Trim$(CStr(v))
End Function

Private Function InNum(ByVal nm As String) As Double
    InNum = Num(InVal(nm))
End Function

Private Function GridVal(ByVal nm As String, ByVal r As Long, ByVal c As Long) As Variant
    Dim v As Variant
    On Error GoTo NoName
    v = ThisWorkbook.Names(nm).RefersToRange.Cells(r, c).Value
    If IsError(v) Then v = Empty
    GridVal = v
    Exit Function
NoName:
    GridVal = Empty
End Function

Private Sub PrepareLog(ByVal title As String)
    Dim f As Worksheet
    nProblems = 0
    nWarnings = 0
    logRow = 5
    Set f = Sh(FORM_NAME)
    If f Is Nothing Then Exit Sub
    f.Range(f.Cells(5, LOG_COL), f.Cells(2000, LOG_COL + 2)).ClearContents
    f.Range(f.Cells(5, LOG_COL), f.Cells(2000, LOG_COL + 2)).Font.Color = RGB(0, 0, 0)
    f.Cells(3, LOG_COL).Value = title
End Sub

Private Sub Note(ByVal sheetName As String, ByVal where As String, ByVal what As String)
    Dim f As Worksheet
    Set f = Sh(FORM_NAME)
    If f Is Nothing Then
        Debug.Print sheetName, where, what
        Exit Sub
    End If
    f.Cells(logRow, LOG_COL).Value = sheetName
    f.Cells(logRow, LOG_COL + 1).Value = where
    f.Cells(logRow, LOG_COL + 2).Value = what
    If where = "PROBLEM" Then f.Range(f.Cells(logRow, LOG_COL), f.Cells(logRow, LOG_COL + 2)).Font.Color = RGB(192, 0, 0)
    If where = "check" Then f.Range(f.Cells(logRow, LOG_COL), f.Cells(logRow, LOG_COL + 2)).Font.Color = RGB(156, 87, 0)
    logRow = logRow + 1
End Sub

Private Sub Problem(ByVal sheetName As String, ByVal what As String)
    nProblems = nProblems + 1
    Note sheetName, "PROBLEM", what
End Sub

Private Sub Warn(ByVal sheetName As String, ByVal what As String)
    nWarnings = nWarnings + 1
    Note sheetName, "check", what
End Sub
