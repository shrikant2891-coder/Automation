Attribute VB_Name = "GenerateMPUploader"
Option Explicit

'===============================================================================
' MP Summary - Dynamic Uploader Generator (VBA)
' Run: Alt+F8 -> GenerateMPUploader
'===============================================================================

Private Const OI_FILE As String = "MEC-FKMP-OPEN-INVOICE-FLOW.csv"
Private Const CREDITOR_FILE As String = "FKMP-CREDITOR-REPORT (1).csv"
Private Const IGST_GL_DL As Long = 142067
Private Const IGST_GL_OTHER As Long = 142013
Private Const COL_IGST As String = "sum(igst_total_amount)"
Private Const COL_TDS As String = "sum(invoice_tds_income_tax_amount)"

Private gSl As Long
Private gVNo As Long
Private gDate As Date
Private gMonth As String
Private gPriorMonth As String
Private gCompany As String
Private gInvType As String
Private gLocation As String
Private gFuncName As String

Public Sub GenerateMPUploader()
    Dim wsSum As Worksheet, wsGL As Worksheet, wsUp As Worksheet
    Dim wsCtrl As Worksheet
    Dim glMap As Object, headers As Object
    Dim lastSum As Long, lastGL As Long

    On Error GoTo Fail
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Set wsSum = ThisWorkbook.Worksheets("Summary")
    Set wsGL = ThisWorkbook.Worksheets("GL Backup")
    Set wsUp = ThisWorkbook.Worksheets("Uploader Format")
    Set wsCtrl = EnsureSheet("Control")

    gSl = 0
    gCompany = "HRFK"
    gInvType = "B2B"
    gLocation = "Business Operations"
    gFuncName = "Others"

    lastSum = wsSum.Cells(wsSum.Rows.Count, 1).End(xlUp).Row
    lastGL = wsGL.Cells(wsGL.Rows.Count, 1).End(xlUp).Row
    If lastSum < 2 Then Err.Raise vbObjectError + 1, , "Summary has no data"

    Set glMap = CreateObject("Scripting.Dictionary")
    Set headers = CreateObject("Scripting.Dictionary")
    LoadGLMap wsGL, lastGL, glMap
    LoadHeaders wsSum, headers

    gMonth = DetectMonth(wsSum, lastSum)
    gPriorMonth = PriorMonthLabel(gMonth)
    gDate = MonthEndDate(gMonth)

    ClearData wsUp
    ClearData wsCtrl
    WriteUploaderHeader wsUp
    wsCtrl.Range("A1:E1").Value = Array("Voucher No", "Narration", "Debit", "Credit", "Difference")

    BuildExpenseVoucher wsSum, lastSum, headers, glMap, wsUp, 28, gMonth, OI_FILE, "prepaid", False, "AR-Journal", "prepaid", _
        "MP_Charges_OI closing-Prepaid for the month of " & gMonth, "No"
    BuildExpenseVoucher wsSum, lastSum, headers, glMap, wsUp, 29, gMonth, OI_FILE, "postpaid", False, "AR-Journal", "postpaid", _
        "MP_Charges_OI closing-Postpaid for the month of " & gMonth, "No"
    BuildExpenseVoucher wsSum, lastSum, headers, glMap, wsUp, 30, gMonth, CREDITOR_FILE, "postpaid", False, "AR-Journal", "postpaid", _
        "MP Charges creditor for the month of " & gMonth & " - Postpaid", "No"
    BuildExpenseVoucher wsSum, lastSum, headers, glMap, wsUp, 31, gMonth, CREDITOR_FILE, "prepaid", False, "AR-Journal", "prepaid", _
        "MP Charges creditor for the month of " & gMonth & " - Prepaid", "No"
    BuildExpenseVoucher wsSum, lastSum, headers, glMap, wsUp, 32, gPriorMonth, OI_FILE, "prepaid", True, "AR-Journal", "prepaid", _
        "MP_Charges_OI opening-Prepaid for the month of " & gMonth, "No"
    BuildExpenseVoucher wsSum, lastSum, headers, glMap, wsUp, 33, gPriorMonth, OI_FILE, "postpaid", True, "AR-Journal", "postpaid", _
        "MP_Charges_OI opening-postpaid for the month of " & gMonth, "No"
    BuildExpenseVoucher wsSum, lastSum, headers, glMap, wsUp, 34, gMonth, "Provision", "Provision", False, "AR-Provision", "provision", _
        "MP charges Provision - subsequent return & undelivered for the month of " & gMonth, "Yes"
    BuildExpenseVoucher wsSum, lastSum, headers, glMap, wsUp, 35, gMonth, "VD", "VD", False, "AR-Journal", "vd", _
        "MP  charges volume discount for the month of " & gMonth, "No"
    BuildExpenseVoucher wsSum, lastSum, headers, glMap, wsUp, 36, gMonth, "PBO VD", "VD", False, "AR-Journal", "vd", _
        "MP charges volume discount on PBO adjustment for the month  of " & gMonth, "No"

    BuildTcsTdsVoucher wsSum, lastSum, headers, glMap, wsUp, 41, gMonth, "postpaid", "tcs", _
        "TCS GST Receivable ( Postpaid) For the month of " & gMonth
    BuildTcsTdsVoucher wsSum, lastSum, headers, glMap, wsUp, 42, gMonth, "postpaid", "tds", _
        "TDS Receivable ( Postpaid) For the month of " & gMonth
    BuildTcsTdsVoucher wsSum, lastSum, headers, glMap, wsUp, 43, gMonth, "prepaid", "tcs", _
        "TCS GST Receivable ( Prepaid) For the month of " & gMonth
    BuildTcsTdsVoucher wsSum, lastSum, headers, glMap, wsUp, 44, gMonth, "prepaid", "tds", _
        "TDS Receivable ( Prepaid) For the month of " & gMonth

    WriteControl wsUp, wsCtrl

    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "MP Uploader rebuilt with " & gSl & " lines." & vbCrLf & _
           "Month: " & gMonth & " | Check Control sheet for balance.", vbInformation, "MP Summary"
    Exit Sub

Fail:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "GenerateMPUploader failed (" & Err.Number & "): " & Err.Description, vbCritical, "MP Summary"
End Sub

'===== Loaders ================================================================

Private Sub LoadGLMap(ws As Worksheet, lastRow As Long, glMap As Object)
    Dim r As Long, field As String
    For r = 2 To lastRow
        field = NormHeader(CStr(ws.Cells(r, 1).Value))
        If Len(field) > 0 Then glMap(field) = ws.Cells(r, 2).Value
    Next r
End Sub

Private Sub LoadHeaders(ws As Worksheet, headers As Object)
    Dim c As Long, h As String, lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        h = NormHeader(CStr(ws.Cells(1, c).Value))
        If Len(h) > 0 Then headers(CStr(c)) = h
    Next c
End Sub

Private Function DetectMonth(ws As Worksheet, lastRow As Long) As String
    Dim r As Long, m As String, best As String
    best = ""
    For r = 2 To lastRow
        m = Trim$(CStr(ws.Cells(r, 1).Value))
        If Len(m) > 0 Then
            If Len(best) = 0 Or MonthSortKey(m) >= MonthSortKey(best) Then best = m
        End If
    Next r
    DetectMonth = best
End Function

Private Function MonthSortKey(ByVal lbl As String) As Long
    Dim dt As Date
    dt = MonthEndDate(lbl)
    MonthSortKey = Year(dt) * 100 + Month(dt)
End Function

Private Function PriorMonthLabel(ByVal lbl As String) As String
    Dim dt As Date
    dt = MonthEndDate(lbl)
    dt = DateSerial(Year(dt), Month(dt), 1)
    dt = DateAdd("m", -1, dt)
    PriorMonthLabel = Format(dt, "Mmm") & "'" & Right$(CStr(Year(dt)), 2)
End Function

Private Function MonthEndDate(ByVal lbl As String) As Date
    Dim s As String, dt As Date
    s = Replace(Replace(Trim$(lbl), "'", ""), ChrW(&H2019), "")
    On Error Resume Next
    dt = DateValue("1 " & Left$(s, 3) & " 20" & Right$(s, 2))
    If Err.Number <> 0 Then
        Err.Clear
        dt = Date
    End If
    On Error GoTo 0
    MonthEndDate = DateSerial(Year(dt), Month(dt) + 1, 0)
End Function

Private Function NormHeader(ByVal h As String) As String
    NormHeader = Trim$(Replace(Replace(h, vbLf, ""), vbCr, ""))
End Function

Private Function NormalizeState(ByVal st As Variant) As String
    Dim s As String
    s = UCase$(Trim$(CStr(st & "")))
    If Len(s) = 0 Or s = "NA" Or s = "NONE" Then Exit Function
    If Left$(s, 3) = "IN-" Then
        NormalizeState = s
    ElseIf Len(s) = 2 Then
        NormalizeState = "IN-" & s
    Else
        NormalizeState = s
    End If
End Function

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
    If glMap.Exists(h) Then
        GLForColumn = glMap(h)
    Else
        GLForColumn = Empty
    End If
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
    If Trim$(CStr(ws.Cells(r, 1).Value)) <> monthLbl Then Exit Function
    If Trim$(CStr(ws.Cells(r, 3).Value)) <> orderType Then Exit Function
    If Len(filename) > 0 Then
        If Trim$(CStr(ws.Cells(r, 2).Value)) <> filename Then Exit Function
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

'===== Voucher builders =========================================================

Private Sub BuildExpenseVoucher(ws As Worksheet, lastRow As Long, headers As Object, glMap As Object, _
    wsUp As Worksheet, ByVal vNo As Long, ByVal monthLbl As String, ByVal filename As String, _
    ByVal orderType As String, ByVal reverse As Boolean, ByVal vType As String, ByVal debtorScope As String, _
    ByVal narration As String, ByVal provRev As String)

    Dim agg As Object, key As Variant, parts() As String
    Dim r As Long, c As Variant, h As String, st As String, amt As Double, gl As Variant
    Dim totalDr As Double, totalCr As Double, dr As Double, cr As Double
    Dim dGL As Variant, diff As Double

    gVNo = vNo
    Set agg = CreateObject("Scripting.Dictionary")
    For r = 2 To lastRow
        If Not RowMatches(ws, r, monthLbl, orderType, filename) Then GoTo NextR
        st = NormalizeState(ws.Cells(r, 4).Value)
        If Len(st) = 0 Then GoTo NextR
        For Each c In headers.Keys
            h = headers(c)
            If IsExpenseCol(h) Or IsGstCol(h) Then
                amt = NzDbl(ws.Cells(r, CLng(c)).Value)
                If Abs(amt) >= 0.005 Then
                    gl = GLForColumn(glMap, h, st)
                    If Not IsValidGL(gl) Then GoTo NextCol
                    key = st & "|" & h
                    If agg.Exists(key) Then
                        agg(key) = CDbl(agg(key)) + amt
                    Else
                        agg(key) = amt
                    End If
                End If
            End If
NextCol:
        Next c
NextR:
    Next r

    If agg.Count = 0 Then Exit Sub

    totalDr = 0: totalCr = 0
    For Each key In agg.Keys
        parts = Split(CStr(key), "|")
        st = parts(0)
        h = parts(1)
        gl = GLForColumn(glMap, h, st)
        If Not IsValidGL(gl) Then GoTo NextKey
        SignToDrCr CDbl(agg(key)), reverse, dr, cr
        AddLine wsUp, vType, CLng(gl), gDate, st, dr, cr, narration, provRev
        totalDr = totalDr + dr
        totalCr = totalCr + cr
NextKey:
    Next key

    dGL = DebtorGL(glMap, debtorScope)
    If Not IsValidGL(dGL) Then Exit Sub
    diff = Round(totalDr - totalCr, 2)
    If diff > 0 Then
        AddLine wsUp, vType, CLng(dGL), gDate, "IN-OTH", 0, diff, narration, provRev
    ElseIf diff < 0 Then
        AddLine wsUp, vType, CLng(dGL), gDate, "IN-OTH", -diff, 0, narration, provRev
    End If
End Sub

Private Sub BuildTcsTdsVoucher(ws As Worksheet, lastRow As Long, headers As Object, glMap As Object, _
    wsUp As Worksheet, ByVal vNo As Long, ByVal monthLbl As String, ByVal orderType As String, _
    ByVal mode As String, ByVal narration As String)

    Dim agg As Object, key As Variant, parts() As String
    Dim r As Long, c As Variant, h As String, st As String, amt As Double, gl As Variant
    Dim totalDr As Double, totalCr As Double, dr As Double, cr As Double
    Dim dGL As Variant, diff As Double, useCol As Boolean

    gVNo = vNo
    Set agg = CreateObject("Scripting.Dictionary")
    For r = 2 To lastRow
        If Trim$(CStr(ws.Cells(r, 1).Value)) <> monthLbl Then GoTo NextR2
        If Trim$(CStr(ws.Cells(r, 3).Value)) <> orderType Then GoTo NextR2
        st = NormalizeState(ws.Cells(r, 5).Value)
        If Len(st) = 0 Then GoTo NextR2
        For Each c In headers.Keys
            h = headers(c)
            useCol = False
            If mode = "tcs" Then useCol = IsTcsCol(h)
            If mode = "tds" Then useCol = IsTdsCol(h)
            If useCol Then
                amt = NzDbl(ws.Cells(r, CLng(c)).Value)
                If Abs(amt) >= 0.005 Then
                    gl = GLForColumn(glMap, h, st)
                    If Not IsValidGL(gl) Then GoTo NextCol2
                    key = st & "|" & h
                    If agg.Exists(key) Then
                        agg(key) = CDbl(agg(key)) + amt
                    Else
                        agg(key) = amt
                    End If
                End If
            End If
NextCol2:
        Next c
NextR2:
    Next r

    If agg.Count = 0 Then Exit Sub

    totalDr = 0: totalCr = 0
    For Each key In agg.Keys
        parts = Split(CStr(key), "|")
        st = parts(0)
        h = parts(1)
        gl = GLForColumn(glMap, h, st)
        If Not IsValidGL(gl) Then GoTo NextKey2
        SignToDrCr CDbl(agg(key)), False, dr, cr
        AddLine wsUp, "AR-Journal", CLng(gl), gDate, st, dr, cr, narration, "No"
        totalDr = totalDr + dr
        totalCr = totalCr + cr
NextKey2:
    Next key

    dGL = DebtorGL(glMap, orderType)
    If Not IsValidGL(dGL) Then Exit Sub
    diff = Round(totalDr - totalCr, 2)
    If diff > 0 Then
        AddLine wsUp, "AR-Journal", CLng(dGL), gDate, "IN-OTH", 0, diff, narration, "No"
    ElseIf diff < 0 Then
        AddLine wsUp, "AR-Journal", CLng(dGL), gDate, "IN-OTH", -diff, 0, narration, "No"
    End If
End Sub

'===== Output helpers =========================================================

Private Sub AddLine(ws As Worksheet, ByVal vType As String, ByVal account As Long, ByVal dt As Date, _
    ByVal state As String, ByVal dr As Double, ByVal cr As Double, ByVal narration As String, ByVal provRev As String)
    Dim rr As Long
    dr = Round(dr, 2)
    cr = Round(cr, 2)
    If Abs(dr) < 0.005 And Abs(cr) < 0.005 Then Exit Sub
    If dr < 0 And cr = 0 Then cr = -dr: dr = 0
    If cr < 0 And dr = 0 Then dr = -cr: cr = 0

    gSl = gSl + 1
    rr = gSl + 1
    ws.Cells(rr, 1).Value = vType
    ws.Cells(rr, 2).Value = account
    ws.Cells(rr, 3).Value = dt
    ws.Cells(rr, 4).Value = ""
    ws.Cells(rr, 5).Value = ""
    ws.Cells(rr, 6).Value = gVNo
    ws.Cells(rr, 7).Value = state
    ws.Cells(rr, 8).Value = gFuncName
    ws.Cells(rr, 9).Value = gLocation
    ws.Cells(rr, 10).Value = dr
    ws.Cells(rr, 11).Value = cr
    ws.Cells(rr, 12).Value = narration
    ws.Cells(rr, 13).Value = gSl
    ws.Cells(rr, 14).Value = gInvType
    ws.Cells(rr, 15).Value = gCompany
    ws.Cells(rr, 16).Value = provRev
End Sub

Private Sub WriteUploaderHeader(ws As Worksheet)
    ws.Cells(1, 1).Value = "VoucherType"
    ws.Cells(1, 2).Value = "Account Name"
    ws.Cells(1, 3).Value = "Date"
    ws.Cells(1, 4).Value = "Ref New Field"
    ws.Cells(1, 5).Value = "Ledger Narration"
    ws.Cells(1, 6).Value = "Voucher No"
    ws.Cells(1, 7).Value = " State Name"
    ws.Cells(1, 8).Value = "Function"
    ws.Cells(1, 9).Value = "Location"
    ws.Cells(1, 10).Value = "Debit Amount"
    ws.Cells(1, 11).Value = "Credit Amount"
    ws.Cells(1, 12).Value = "Narration"
    ws.Cells(1, 13).Value = "Sl no"
    ws.Cells(1, 14).Value = "Invoice  Type"
    ws.Cells(1, 15).Value = "Company Code"
    ws.Cells(1, 16).Value = "Is provision reverse"
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
        If IsEmpty(v) Or v = 0 Then GoTo NextR3
        Acc d, CStr(v), NzDbl(wsUp.Cells(r, 10).Value)
        Acc c, CStr(v), NzDbl(wsUp.Cells(r, 11).Value)
        If Not n.Exists(CStr(v)) Then n.Add CStr(v), wsUp.Cells(r, 12).Value
NextR3:
    Next r

    rr = 2
    For Each key In d.Keys
        wsCtrl.Cells(rr, 1).Value = key
        wsCtrl.Cells(rr, 2).Value = n(key)
        wsCtrl.Cells(rr, 3).Value = Round(CDbl(d(key)), 2)
        wsCtrl.Cells(rr, 4).Value = Round(CDbl(c(key)), 2)
        diff = Round(CDbl(d(key)) - CDbl(c(key)), 2)
        wsCtrl.Cells(rr, 5).Value = diff
        rr = rr + 1
    Next key
End Sub

Private Sub Acc(dict As Object, ByVal k As String, ByVal v As Double)
    If dict.Exists(k) Then
        dict(k) = CDbl(dict(k)) + v
    Else
        dict.Add k, v
    End If
End Sub

Private Function NzDbl(v As Variant) As Double
    If IsError(v) Or IsEmpty(v) Or v = "" Then
        NzDbl = 0
    ElseIf IsNumeric(v) Then
        NzDbl = CDbl(v)
    Else
        NzDbl = 0
    End If
End Function

Private Function IsValidGL(gl As Variant) As Boolean
    If IsEmpty(gl) Then Exit Function
    If UCase$(Trim$(CStr(gl))) = "NA" Then Exit Function
    If Not IsNumeric(gl) Then Exit Function
    IsValidGL = True
End Function

Private Function EnsureSheet(ByVal name As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = name
    End If
    Set EnsureSheet = ws
End Function

Private Sub ClearData(ws As Worksheet)
    On Error Resume Next
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    On Error GoTo 0
    ws.Cells.ClearContents
End Sub
