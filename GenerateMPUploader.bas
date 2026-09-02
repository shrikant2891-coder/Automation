Attribute VB_Name = "GenerateMPUploader"
Option Explicit

'===============================================================================
' MP Summary - Dynamic Uploader Generator (VBA)
' Same framework as RetailX GenerateUploader.bas
' Sheets: Summary, GL Backup, Uploader Format
'===============================================================================

Private Const OI_FILE As String = "MEC-FKMP-OPEN-INVOICE-FLOW.csv"
Private Const CREDITOR_FILE As String = "FKMP-CREDITOR-REPORT (1).csv"
Private Const IGST_GL_DL As Long = 142067
Private Const IGST_GL_OTHER As Long = 142013
Private Const COL_IGST As String = "sum(igst_total_amount)"
Private Const COL_TDS As String = "sum(invoice_tds_income_tax_amount)"
Private Const FN_OTHERS As String = "Others"

Private gSl As Long
Private gDate As Date
Private gMonthLabel As String
Private gPriorMonth As String
Private gCompany As String
Private gInvType As String
Private gLocation As String
Private gLastCol As Long

Public Sub GenerateMPUploader()
    Dim wsSum As Worksheet, wsGL As Worksheet, wsUp As Worksheet
    Dim wsCtrl As Worksheet
    Dim glMap As Object
    Dim lastSum As Long, lastGL As Long

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    On Error GoTo Fail

    Set wsSum = ThisWorkbook.Worksheets("Summary")
    Set wsGL = ThisWorkbook.Worksheets("GL Backup")
    Set wsUp = ThisWorkbook.Worksheets("Uploader Format")

    gCompany = "HRFK"
    gInvType = "B2B"
    gLocation = "Business Operations"
    gSl = 0

    lastSum = wsSum.Cells(wsSum.Rows.Count, 1).End(xlUp).Row
    lastGL = wsGL.Cells(wsGL.Rows.Count, 1).End(xlUp).Row
    gLastCol = wsSum.Cells(1, wsSum.Columns.Count).End(xlToLeft).Column
    If lastSum < 2 Then Err.Raise vbObjectError + 1, , "Summary has no data"

    LoadConfig gMonthLabel, gDate, gCompany, gInvType
    gPriorMonth = PriorMonthLabel(gMonthLabel)

    Set glMap = CreateObject("Scripting.Dictionary")
    LoadGLMap wsGL, lastGL, glMap

    ClearSheetKeepHeader wsUp, 16
    EnsureSheet "Control"
    Set wsCtrl = ThisWorkbook.Worksheets("Control")
    ClearSheetKeepHeader wsCtrl, 5
    wsCtrl.Range("A1:E1").Value = Array("Voucher No", "Narration", "Debit", "Credit", "Difference")

    BuildExpenseVoucher wsSum, lastSum, glMap, wsUp, 28, gMonthLabel, OI_FILE, "prepaid", False, "AR-Journal", "prepaid", _
        "MP_Charges_OI closing-Prepaid for the month of " & gMonthLabel, "No"
    BuildExpenseVoucher wsSum, lastSum, glMap, wsUp, 29, gMonthLabel, OI_FILE, "postpaid", False, "AR-Journal", "postpaid", _
        "MP_Charges_OI closing-Postpaid for the month of " & gMonthLabel, "No"
    BuildExpenseVoucher wsSum, lastSum, glMap, wsUp, 30, gMonthLabel, CREDITOR_FILE, "postpaid", False, "AR-Journal", "postpaid", _
        "MP Charges creditor for the month of " & gMonthLabel & " - Postpaid", "No"
    BuildExpenseVoucher wsSum, lastSum, glMap, wsUp, 31, gMonthLabel, CREDITOR_FILE, "prepaid", False, "AR-Journal", "prepaid", _
        "MP Charges creditor for the month of " & gMonthLabel & " - Prepaid", "No"
    BuildExpenseVoucher wsSum, lastSum, glMap, wsUp, 32, gPriorMonth, OI_FILE, "prepaid", True, "AR-Journal", "prepaid", _
        "MP_Charges_OI opening-Prepaid for the month of " & gMonthLabel, "No"
    BuildExpenseVoucher wsSum, lastSum, glMap, wsUp, 33, gPriorMonth, OI_FILE, "postpaid", True, "AR-Journal", "postpaid", _
        "MP_Charges_OI opening-postpaid for the month of " & gMonthLabel, "No"
    BuildExpenseVoucher wsSum, lastSum, glMap, wsUp, 34, gMonthLabel, "Provision", "Provision", False, "AR-Provision", "provision", _
        "MP charges Provision - subsequent return & undelivered for the month of " & gMonthLabel, "Yes"
    BuildExpenseVoucher wsSum, lastSum, glMap, wsUp, 35, gMonthLabel, "VD", "VD", False, "AR-Journal", "vd", _
        "MP  charges volume discount for the month of " & gMonthLabel, "No"
    BuildExpenseVoucher wsSum, lastSum, glMap, wsUp, 36, gMonthLabel, "PBO VD", "VD", False, "AR-Journal", "vd", _
        "MP charges volume discount on PBO adjustment for the month  of " & gMonthLabel, "No"

    BuildTcsTdsVoucher wsSum, lastSum, glMap, wsUp, 41, gMonthLabel, "postpaid", "tcs", _
        "TCS GST Receivable ( Postpaid) For the month of " & gMonthLabel
    BuildTcsTdsVoucher wsSum, lastSum, glMap, wsUp, 42, gMonthLabel, "postpaid", "tds", _
        "TDS Receivable ( Postpaid) For the month of " & gMonthLabel
    BuildTcsTdsVoucher wsSum, lastSum, glMap, wsUp, 43, gMonthLabel, "prepaid", "tcs", _
        "TCS GST Receivable ( Prepaid) For the month of " & gMonthLabel
    BuildTcsTdsVoucher wsSum, lastSum, glMap, wsUp, 44, gMonthLabel, "prepaid", "tds", _
        "TDS Receivable ( Prepaid) For the month of " & gMonthLabel

    WriteControl wsUp, wsCtrl

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "MP Uploader rebuilt with " & gSl & " lines." & vbCrLf & _
           "Month: " & gMonthLabel & " | Check Control sheet for balance.", vbInformation, "MP Summary"
    Exit Sub

Fail:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "GenerateMPUploader failed: " & Err.Description, vbCritical, "MP Summary"
End Sub


'===== GL Backup ==============================================================

Private Sub LoadGLMap(wsGL As Worksheet, lastRow As Long, glMap As Object)
    Dim r As Long, field As String
    For r = 2 To lastRow
        field = NormHeader(CStr(wsGL.Cells(r, 1).Value & ""))
        If Len(field) > 0 Then
            If Not glMap.Exists(field) Then glMap.Add field, wsGL.Cells(r, 2).Value
        End If
    Next r
End Sub

Private Function GLForColumn(glMap As Object, ByVal header As String, ByVal state As String) As Variant
    Dim h As String
    h = NormHeader(header)
    If LCase$(h) = COL_IGST Then
        If state = "IN-DL" Then
            GLForColumn = IGST_GL_DL
        Else
            GLForColumn = IGST_GL_OTHER
        End If
        Exit Function
    End If
    If glMap.Exists(h) Then GLForColumn = glMap(h) Else GLForColumn = Empty
End Function

Private Function DebtorGL(glMap As Object, ByVal scope As String) As Variant
    Select Case LCase$(scope)
        Case "prepaid"
            If glMap.Exists("Prepaid Debtor") Then DebtorGL = glMap("Prepaid Debtor")
        Case "postpaid"
            If glMap.Exists("Postpaid Debtor") Then DebtorGL = glMap("Postpaid Debtor")
        Case "provision"
            If glMap.Exists("Provision Ledger") Then DebtorGL = glMap("Provision Ledger")
        Case "vd"
            If glMap.Exists("VD Debtor") Then DebtorGL = glMap("VD Debtor")
    End Select
End Function


'===== Voucher builders =======================================================

Private Sub BuildExpenseVoucher(ws As Worksheet, lastRow As Long, glMap As Object, wsUp As Worksheet, _
    voucherNo As Long, monthLbl As String, filename As String, orderType As String, reverse As Boolean, _
    vType As String, debtorScope As String, narr As String, provRev As String)

    Dim agg As Object
    Dim r As Long, colIdx As Long, h As String, st As String, amt As Double, gl As Variant
    Dim key As Variant, parts() As String, dr As Double, cr As Double
    Dim totalDr As Double, totalCr As Double, diff As Double, dGL As Variant

    Set agg = CreateObject("Scripting.Dictionary")

    For r = 2 To lastRow
        If Not RowMatches(ws, r, monthLbl, orderType, filename) Then GoTo NextR
        st = NormalizeState(CStr(ws.Cells(r, 4).Value & ""))
        If Len(st) = 0 Then GoTo NextR
        For colIdx = 1 To gLastCol
            h = HeaderAt(ws, colIdx)
            If IsExpenseCol(h) Or IsGstCol(h) Then
                amt = Nz(ws.Cells(r, colIdx).Value)
                If Abs(amt) >= 0.005 Then
                    gl = GLForColumn(glMap, h, st)
                    If Not IsValidGL(gl) Then GoTo NextCol
                    key = CStr(gl) & "|" & st
                    Acc agg, key, amt
                End If
            End If
NextCol:
        Next colIdx
NextR:
    Next r

    If agg.Count = 0 Then Exit Sub

    totalDr = 0: totalCr = 0
    For Each key In agg.Keys
        parts = Split(CStr(key), "|")
        gl = parts(0)
        st = parts(1)
        SignToDrCr CDbl(agg(key)), reverse, dr, cr
        If dr > 0.005 Then
            AddLine wsUp, voucherNo, CLng(gl), st, FN_OTHERS, dr, 0, narr, vType, provRev
            totalDr = totalDr + dr
        End If
        If cr > 0.005 Then
            AddLine wsUp, voucherNo, CLng(gl), st, FN_OTHERS, 0, cr, narr, vType, provRev
            totalCr = totalCr + cr
        End If
    Next key

    dGL = DebtorGL(glMap, debtorScope)
    If Not IsValidGL(dGL) Then Exit Sub
    diff = Round(totalDr - totalCr, 2)
    If diff > 0 Then
        AddLine wsUp, voucherNo, CLng(dGL), "IN-OTH", FN_OTHERS, 0, diff, narr, vType, provRev
    ElseIf diff < 0 Then
        AddLine wsUp, voucherNo, CLng(dGL), "IN-OTH", FN_OTHERS, -diff, 0, narr, vType, provRev
    End If
End Sub

Private Sub BuildTcsTdsVoucher(ws As Worksheet, lastRow As Long, glMap As Object, wsUp As Worksheet, _
    voucherNo As Long, monthLbl As String, orderType As String, mode As String, narr As String)

    Dim agg As Object
    Dim r As Long, colIdx As Long, h As String, st As String, amt As Double, gl As Variant
    Dim key As Variant, parts() As String, dr As Double, cr As Double, useCol As Boolean
    Dim totalDr As Double, totalCr As Double, diff As Double, dGL As Variant

    Set agg = CreateObject("Scripting.Dictionary")

    For r = 2 To lastRow
        If Trim$(CStr(ws.Cells(r, 1).Value & "")) <> monthLbl Then GoTo NextR2
        If Trim$(CStr(ws.Cells(r, 3).Value & "")) <> orderType Then GoTo NextR2
        st = NormalizeState(CStr(ws.Cells(r, 5).Value & ""))
        If Len(st) = 0 Then GoTo NextR2
        For colIdx = 1 To gLastCol
            h = HeaderAt(ws, colIdx)
            useCol = False
            If mode = "tcs" Then useCol = IsTcsCol(h)
            If mode = "tds" Then useCol = IsTdsCol(h)
            If useCol Then
                amt = Nz(ws.Cells(r, colIdx).Value)
                If Abs(amt) >= 0.005 Then
                    gl = GLForColumn(glMap, h, st)
                    If Not IsValidGL(gl) Then GoTo NextCol2
                    key = CStr(gl) & "|" & st
                    Acc agg, key, amt
                End If
            End If
NextCol2:
        Next colIdx
NextR2:
    Next r

    If agg.Count = 0 Then Exit Sub

    totalDr = 0: totalCr = 0
    For Each key In agg.Keys
        parts = Split(CStr(key), "|")
        gl = parts(0)
        st = parts(1)
        SignToDrCr CDbl(agg(key)), False, dr, cr
        If dr > 0.005 Then
            AddLine wsUp, voucherNo, CLng(gl), st, FN_OTHERS, dr, 0, narr, "AR-Journal", "No"
            totalDr = totalDr + dr
        End If
        If cr > 0.005 Then
            AddLine wsUp, voucherNo, CLng(gl), st, FN_OTHERS, 0, cr, narr, "AR-Journal", "No"
            totalCr = totalCr + cr
        End If
    Next key

    dGL = DebtorGL(glMap, orderType)
    If Not IsValidGL(dGL) Then Exit Sub
    diff = Round(totalDr - totalCr, 2)
    If diff > 0 Then
        AddLine wsUp, voucherNo, CLng(dGL), "IN-OTH", FN_OTHERS, 0, diff, narr, "AR-Journal", "No"
    ElseIf diff < 0 Then
        AddLine wsUp, voucherNo, CLng(dGL), "IN-OTH", FN_OTHERS, -diff, 0, narr, "AR-Journal", "No"
    End If
End Sub


'===== Aggregators / writers (RetailX style) ==================================

Private Sub Acc(d As Object, ByVal key As String, ByVal value As Double)
    If Len(key) = 0 Then Exit Sub
    If d.Exists(key) Then d(key) = CDbl(d(key)) + value Else d.Add key, value
End Sub

Private Sub AddLine(wsUp As Worksheet, voucherNo As Long, account As Variant, _
                    state As String, fn As String, debit As Double, credit As Double, narr As String, _
                    vType As String, provRev As String)
    Dim rr As Long
    debit = Round(debit, 2)
    credit = Round(credit, 2)
    If Abs(debit) < 0.005 And Abs(credit) < 0.005 Then Exit Sub
    If debit < 0 And credit = 0 Then credit = -debit: debit = 0
    If credit < 0 And debit = 0 Then debit = -credit: credit = 0

    gSl = gSl + 1
    rr = gSl + 1
    wsUp.Cells(rr, 1).Value = vType
    wsUp.Cells(rr, 2).Value = account
    wsUp.Cells(rr, 3).Value = gDate
    wsUp.Cells(rr, 6).Value = voucherNo
    wsUp.Cells(rr, 7).Value = state
    wsUp.Cells(rr, 8).Value = fn
    wsUp.Cells(rr, 9).Value = gLocation
    wsUp.Cells(rr, 10).Value = debit
    wsUp.Cells(rr, 11).Value = credit
    wsUp.Cells(rr, 12).Value = narr
    wsUp.Cells(rr, 13).Value = gSl
    wsUp.Cells(rr, 14).Value = gInvType
    wsUp.Cells(rr, 15).Value = gCompany
    wsUp.Cells(rr, 16).Value = provRev
End Sub

Private Sub WriteControl(wsUp As Worksheet, wsCtrl As Worksheet)
    Dim last As Long, r As Long, v As Variant
    Dim d As Object, c As Object, n As Object
    Dim key As Variant, rr As Long, diff As Double
    Set d = CreateObject("Scripting.Dictionary")
    Set c = CreateObject("Scripting.Dictionary")
    Set n = CreateObject("Scripting.Dictionary")
    last = wsUp.Cells(wsUp.Rows.Count, 1).End(xlUp).Row
    For r = 2 To last
        v = wsUp.Cells(r, 6).Value
        Acc d, CStr(v), Nz(wsUp.Cells(r, 10).Value)
        Acc c, CStr(v), Nz(wsUp.Cells(r, 11).Value)
        If Not n.Exists(CStr(v)) Then n.Add CStr(v), wsUp.Cells(r, 12).Value
    Next r
    rr = 2
    For Each key In d.Keys
        wsCtrl.Cells(rr, 1).Value = CLng(key)
        wsCtrl.Cells(rr, 2).Value = n(key)
        wsCtrl.Cells(rr, 3).Value = Round(CDbl(d(key)), 2)
        wsCtrl.Cells(rr, 4).Value = Round(CDbl(c(key)), 2)
        diff = Round(CDbl(d(key)) - CDbl(c(key)), 2)
        wsCtrl.Cells(rr, 5).Value = diff
        rr = rr + 1
    Next key
End Sub


'===== Utilities ==============================================================

Private Function HeaderAt(ws As Worksheet, ByVal colIdx As Long) As String
    HeaderAt = NormHeader(CStr(ws.Cells(1, colIdx).Value & ""))
End Function

Private Function NormHeader(ByVal h As String) As String
    NormHeader = Trim$(Replace(Replace(h, vbLf, ""), vbCr, ""))
End Function

Private Function NormalizeState(ByVal st As String) As String
    Dim s As String
    s = UCase$(Trim$(st))
    If Len(s) = 0 Or s = "NA" Or s = "NONE" Then Exit Function
    If Left$(s, 3) = "IN-" Then
        NormalizeState = s
    ElseIf Len(s) = 2 Then
        NormalizeState = "IN-" & s
    Else
        NormalizeState = s
    End If
End Function

Private Function IsExpenseCol(ByVal h As String) As Boolean
    Dim u As String
    u = LCase$(NormHeader(h))
    If u = "sum(due_amount)" Then Exit Function
    If InStr(u, "invoice_tcs") > 0 Or InStr(u, "invoice_tds") > 0 Then Exit Function
    If InStr(u, "sgst_utgst") > 0 Or InStr(u, "cgst_total") > 0 Or u = COL_IGST Then Exit Function
    IsExpenseCol = (Left$(u, 4) = "sum(")
End Function

Private Function IsGstCol(ByVal h As String) As Boolean
    Dim u As String
    u = LCase$(NormHeader(h))
    IsGstCol = (InStr(u, "sgst_utgst") > 0 Or u = "sum(cgst_total_amount)" Or u = COL_IGST)
End Function

Private Function IsTcsCol(ByVal h As String) As Boolean
    Dim u As String
    u = LCase$(NormHeader(h))
    IsTcsCol = (InStr(u, "invoice_tcs_cgst") > 0 Or InStr(u, "invoice_tcs_sgst") > 0 Or InStr(u, "invoice_tcs_igst") > 0)
End Function

Private Function IsTdsCol(ByVal h As String) As Boolean
    IsTdsCol = (LCase$(NormHeader(h)) = COL_TDS)
End Function

Private Function RowMatches(ws As Worksheet, ByVal r As Long, ByVal monthLbl As String, _
    ByVal orderType As String, ByVal filename As String) As Boolean
    If Trim$(CStr(ws.Cells(r, 1).Value & "")) <> monthLbl Then Exit Function
    If Trim$(CStr(ws.Cells(r, 3).Value & "")) <> orderType Then Exit Function
    If Len(filename) > 0 Then
        If Trim$(CStr(ws.Cells(r, 2).Value & "")) <> filename Then Exit Function
    End If
    RowMatches = True
End Function

Private Sub SignToDrCr(ByVal amount As Double, ByVal reverse As Boolean, ByRef dr As Double, ByRef cr As Double)
    dr = 0: cr = 0
    If Abs(amount) < 0.005 Then Exit Sub
    If amount < 0 Then
        dr = Abs(amount)
    Else
        cr = Abs(amount)
    End If
    If reverse Then
        Dim t As Double
        t = dr: dr = cr: cr = t
    End If
End Sub

Private Function IsValidGL(gl As Variant) As Boolean
    If IsEmpty(gl) Then Exit Function
    If UCase$(Trim$(CStr(gl))) = "NA" Then Exit Function
    If Not IsNumeric(gl) Then Exit Function
    IsValidGL = True
End Function

Private Function Nz(ByVal v As Variant) As Double
    If IsError(v) Or IsEmpty(v) Or v = "" Then Nz = 0 Else Nz = CDbl(v)
End Function

Private Function DetectMonth(ws As Worksheet, lastRow As Long) As String
    Dim r As Long, m As String, best As String
    best = ""
    For r = 2 To lastRow
        m = Trim$(CStr(ws.Cells(r, 1).Value & ""))
        If Len(m) > 0 Then
            If Len(best) = 0 Or MonthSortKey(m) >= MonthSortKey(best) Then best = m
        End If
    Next r
    DetectMonth = best
End Function

Private Function MonthSortKey(ByVal lbl As String) As Long
    Dim dt As Date
    dt = MonthEndFromLabel(lbl)
    MonthSortKey = Year(dt) * 100 + Month(dt)
End Function

Private Function PriorMonthLabel(ByVal lbl As String) As String
    Dim dt As Date
    dt = MonthEndFromLabel(lbl)
    dt = DateSerial(Year(dt), Month(dt), 1)
    dt = DateAdd("m", -1, dt)
    PriorMonthLabel = Format(dt, "Mmm") & "'" & Right$(CStr(Year(dt)), 2)
End Function

Private Sub LoadConfig(ByRef monthLabel As String, ByRef bookDate As Date, _
                       ByRef companyCode As String, ByRef invoiceType As String)
    Dim wsCfg As Worksheet
    On Error Resume Next
    Set wsCfg = ThisWorkbook.Worksheets("Config")
    On Error GoTo 0
    If Not wsCfg Is Nothing Then
        If Len(Trim$(CStr(wsCfg.Range("B2").Value & ""))) > 0 Then
            monthLabel = CStr(wsCfg.Range("B2").Value)
        End If
        If IsDate(wsCfg.Range("B3").Value) Then
            bookDate = CDate(wsCfg.Range("B3").Value)
        End If
        If Len(Trim$(CStr(wsCfg.Range("B4").Value & ""))) > 0 Then
            companyCode = CStr(wsCfg.Range("B4").Value)
        End If
        If Len(Trim$(CStr(wsCfg.Range("B5").Value & ""))) > 0 Then
            invoiceType = CStr(wsCfg.Range("B5").Value)
        End If
    End If
    If Len(Trim$(monthLabel)) = 0 Then
        monthLabel = DetectMonth(ThisWorkbook.Worksheets("Summary"), _
            ThisWorkbook.Worksheets("Summary").Cells(ThisWorkbook.Worksheets("Summary").Rows.Count, 1).End(xlUp).Row)
    End If
    If bookDate = 0 Then bookDate = MonthEndFromLabel(monthLabel)
End Sub

Private Function MonthEndFromLabel(ByVal label As String) As Date
    Dim s As String, dt As Date
    s = Replace(Replace(Trim$(label), "'", ""), ChrW(&H2019), "")
    On Error Resume Next
    dt = DateValue("1 " & Left$(s, Len(s) - 2) & " 20" & Right$(s, 2))
    If Err.Number <> 0 Then
        Err.Clear
        dt = DateSerial(Year(Date), Month(Date), 0)
    End If
    On Error GoTo 0
    MonthEndFromLabel = DateSerial(Year(dt), Month(dt) + 1, 0)
End Function

Private Sub ClearSheetKeepHeader(ws As Worksheet, headerCols As Long)
    Dim last As Long
    last = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If last >= 2 Then ws.Rows("2:" & last).Delete
    If (ws.Name = "Uploader" Or ws.Name = "Uploader Format") And Len(ws.Cells(1, 1).Value & "") = 0 Then
        ws.Range("A1:P1").Value = Array( _
            "VoucherType", "Account Name", "Date", "Ref New Field", "Ledger Narration", _
            "Voucher No", " State Name", "Function", "Location", "Debit Amount", _
            "Credit Amount", "Narration", "Sl no", "Invoice  Type", "Company Code", _
            "Is provision reverse")
    End If
End Sub

Private Function EnsureSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set EnsureSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If EnsureSheet Is Nothing Then
        Set EnsureSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        EnsureSheet.Name = sheetName
    End If
End Function
