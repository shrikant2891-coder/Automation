Attribute VB_Name = "GenerateUploader"
Option Explicit

'===============================================================================
' RetailX — Dynamic Uploader Generator (VBA)
'-------------------------------------------------------------------------------
' Import: Alt+F11 → File → Import File → select this .bas → Run GenerateUploader
'
' What it fixes vs old SUMIFS/VLOOKUP grid:
'   • New states in Extract are posted automatically
'   • New analytics_category values are posted if present in GL Master
'   • One voucher series per revenue type (report_name driven)
'
' Sheets required: Extract, GL Master, Uploader
' Optional:       Config (voucher numbers / narrations)
'===============================================================================

Private Const COL_MONTH As Long = 1
Private Const COL_REPORT As Long = 2
Private Const COL_SCTYPE As Long = 3
Private Const COL_ST_FROM As Long = 5
Private Const COL_ST_TO As Long = 6
Private Const COL_CAT As Long = 7
Private Const COL_SALES_REV As Long = 10
Private Const COL_SALES_DISC As Long = 11
Private Const COL_SALES_TAX As Long = 12
Private Const COL_PBO_REV As Long = 13
Private Const COL_PBO_TAX As Long = 14
Private Const COL_PICKUP_REV As Long = 16
Private Const COL_PICKUP_TAX As Long = 17
Private Const COL_SALES_RET As Long = 18
Private Const COL_RET_DISC As Long = 19
Private Const COL_RET_TAX As Long = 20
Private Const COL_SHIP As Long = 21
Private Const COL_SHIP_TAX As Long = 22
Private Const COL_PD As Long = 23
Private Const COL_PD_TAX As Long = 24
Private Const COL_BUYER As Long = 25
Private Const COL_BUYER_TAX As Long = 26
Private Const COL_BUMP As Long = 27
Private Const COL_BUMP_TAX As Long = 28

Private Const GL_POSTPAID As Long = 131144
Private Const GL_PREPAID As Long = 131102
Private Const GL_PBO_DEBTOR As Long = 131126

Private gSl As Long
Private gDate As Date
Private gMonthLabel As String
Private gCompany As String
Private gInvType As String
Private gLocation As String

Public Sub GenerateUploader()
    Dim wsEx As Worksheet, wsGL As Worksheet, wsUp As Worksheet
    Dim wsCtrl As Worksheet, wsUn As Worksheet
    Dim revGL As Object, discGL As Object, retGL As Object
    Dim lastEx As Long, r As Long
    Dim unmapped As Object

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    On Error GoTo Fail

    Set wsEx = ThisWorkbook.Worksheets("Extract")
    Set wsGL = ThisWorkbook.Worksheets("GL Master")
    Set wsUp = ThisWorkbook.Worksheets("Uploader")

    gCompany = "UTSPLFK"
    gInvType = "B2C"
    gLocation = "Business Operations"
    gSl = 0

    lastEx = wsEx.Cells(wsEx.Rows.Count, COL_REPORT).End(xlUp).Row
    If lastEx < 2 Then Err.Raise vbObjectError + 1, , "Extract has no data"

    gMonthLabel = CStr(wsEx.Cells(2, COL_MONTH).Value)
    gDate = MonthEndFromLabel(gMonthLabel)

    Set revGL = CreateObject("Scripting.Dictionary")
    Set discGL = CreateObject("Scripting.Dictionary")
    Set retGL = CreateObject("Scripting.Dictionary")
    Set unmapped = CreateObject("Scripting.Dictionary")
    LoadGLMaps wsGL, revGL, discGL, retGL

    ClearSheetKeepHeader wsUp, 16
    EnsureSheet("Control")
    EnsureSheet("Unmapped")
    Set wsCtrl = ThisWorkbook.Worksheets("Control")
    Set wsUn = ThisWorkbook.Worksheets("Unmapped")
    ClearSheetKeepHeader wsCtrl, 5
    ClearSheetKeepHeader wsUn, 3
    wsCtrl.Range("A1:E1").Value = Array("Voucher No", "Narration", "Debit", "Credit", "Difference")
    wsUn.Range("A1:C1").Value = Array("Type", "Key", "Detail")

    '--- Voucher series (edit here or mirror Config sheet) ---
    BuildSales wsUp, wsEx, lastEx, revGL, discGL, unmapped, 11, _
        "Being sales booked for the month of " & gMonthLabel
    BuildSalesReturn wsUp, wsEx, lastEx, retGL, discGL, unmapped, 12, _
        "Being Sales Return booked for the month of " & gMonthLabel
    BuildShipping wsUp, wsEx, lastEx, 13, _
        "Being Shipping Revenue booked for the month of " & gMonthLabel
    BuildPBO wsUp, wsEx, lastEx, 14, _
        "Being PBO Revenue booked for the month of " & gMonthLabel
    BuildPriceDrop wsUp, wsEx, lastEx, revGL, 15, _
        "Price drop for the month of " & gMonthLabel
    BuildBuyerFee wsUp, wsEx, lastEx, 16, _
        "Being Secure Packaging revenue booked for the month " & gMonthLabel
    BuildPrexoBumpup wsUp, wsEx, lastEx, 17, _
        "Being PREXO BUMPUP Revenue booked for the month of " & gMonthLabel

    WriteControl wsUp, wsCtrl
    WriteUnmapped wsUn, unmapped

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Uploader rebuilt with " & gSl & " lines." & vbCrLf & _
           "Check Control sheet for debit/credit balance and Unmapped for missing GLs.", _
           vbInformation, "RetailX"
    Exit Sub

Fail:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "GenerateUploader failed: " & Err.Description, vbCritical
End Sub


'===== GL Master ===============================================================

Private Sub LoadGLMaps(wsGL As Worksheet, revGL As Object, discGL As Object, retGL As Object)
    Dim last As Long, r As Long
    Dim clas As String, prod As String, rec As String, gl As Variant
    last = wsGL.Cells(wsGL.Rows.Count, 1).End(xlUp).Row
    For r = 2 To last
        clas = Trim$(CStr(wsGL.Cells(r, 1).Value & ""))
        prod = Trim$(CStr(wsGL.Cells(r, 2).Value & ""))
        rec = Trim$(CStr(wsGL.Cells(r, 3).Value & ""))
        gl = wsGL.Cells(r, 5).Value
        If Len(clas) = 0 Or IsEmpty(gl) Or Len(prod) = 0 Then GoTo NextR
        If IsTaxOrDebtorKey(prod) Then GoTo NextR

        If clas = "Sales" Then
            If InStr(1, rec, "discount", vbTextCompare) > 0 Then
                If Not discGL.Exists(prod) Then discGL.Add prod, gl
            ElseIf Len(rec) = 0 Then
                If Not revGL.Exists(prod) Then revGL.Add prod, gl
            End If
        ElseIf clas = "Sales Return" Then
            If InStr(1, rec, "discount", vbTextCompare) > 0 Then
                If Not discGL.Exists(prod) Then discGL.Add prod, gl
            ElseIf Len(rec) = 0 Then
                If Not retGL.Exists(prod) Then retGL.Add prod, gl
            End If
        End If
NextR:
    Next r
End Sub

Private Function IsTaxOrDebtorKey(ByVal prod As String) As Boolean
    Dim p As String
    p = UCase$(Trim$(prod))
    If p = "" Or p = "POSTPAID" Or p = "PREPAID" Or p = "BNPL" Or p = "DEBTOR" Then
        IsTaxOrDebtorKey = True
        Exit Function
    End If
    IsTaxOrDebtorKey = (Left$(p, 4) = "CGST" Or Left$(p, 4) = "SGST" Or Left$(p, 4) = "IGST")
End Function


'===== Voucher builders ========================================================

Private Sub BuildSales(wsUp As Worksheet, wsEx As Worksheet, lastEx As Long, _
                       revGL As Object, discGL As Object, unmapped As Object, _
                       voucherNo As Long, narr As String)
    Dim debtor As Object, revAgg As Object, discAgg As Object, taxAgg As Object
    Dim r As Long, cat As String, sc As String, stFrom As String, stTo As String
    Dim rev As Double, disc As Double, tax As Double
    Dim k As Variant

    Set debtor = CreateObject("Scripting.Dictionary")
    Set revAgg = CreateObject("Scripting.Dictionary")
    Set discAgg = CreateObject("Scripting.Dictionary")
    Set taxAgg = CreateObject("Scripting.Dictionary")

    For r = 2 To lastEx
        If CStr(wsEx.Cells(r, COL_REPORT).Value) <> "NONDIGITAL" Then GoTo NextSales
        cat = Trim$(CStr(wsEx.Cells(r, COL_CAT).Value & ""))
        sc = LCase$(Trim$(CStr(wsEx.Cells(r, COL_SCTYPE).Value & "")))
        stFrom = CStr(wsEx.Cells(r, COL_ST_FROM).Value & "")
        stTo = CStr(wsEx.Cells(r, COL_ST_TO).Value & "")
        rev = Nz(wsEx.Cells(r, COL_SALES_REV).Value)
        disc = Nz(wsEx.Cells(r, COL_SALES_DISC).Value)
        tax = Nz(wsEx.Cells(r, COL_SALES_TAX).Value)

        Acc debtor, sc, rev + disc + tax
        If revGL.Exists(cat) Then
            Acc revAgg, CStr(revGL(cat)) & "|" & AliasState(stFrom), rev
        ElseIf Abs(rev) > 0.005 Then
            NoteUnmapped unmapped, "Sales Revenue GL", cat
        End If
        If discGL.Exists(cat) Then
            Acc discAgg, CStr(discGL(cat)) & "|" & AliasState(stFrom), Abs(disc)
        ElseIf Abs(disc) > 0.005 Then
            NoteUnmapped unmapped, "Sales Discount GL", cat
        End If
        AccTax taxAgg, tax, rev, stFrom, stTo
NextSales:
    Next r

    AddLine wsUp, voucherNo, GL_POSTPAID, "IN-OTH", "Sales", Nz(DictGet(debtor, "postpaid")), 0, narr
    AddLine wsUp, voucherNo, GL_PREPAID, "IN-OTH", "Sales", Nz(DictGet(debtor, "prepaid")), 0, narr
    DumpAggCredit wsUp, voucherNo, revAgg, "Sales", narr
    DumpAggDebit wsUp, voucherNo, discAgg, "Sales", narr
    DumpAggCredit wsUp, voucherNo, taxAgg, "Sales", narr
End Sub

Private Sub BuildSalesReturn(wsUp As Worksheet, wsEx As Worksheet, lastEx As Long, _
                             retGL As Object, discGL As Object, unmapped As Object, _
                             voucherNo As Long, narr As String)
    Dim debtor As Object, retAgg As Object, discAgg As Object, taxAgg As Object
    Dim r As Long, cat As String, sc As String, stFrom As String, stTo As String
    Dim sret As Double, rdisc As Double, rtax As Double, net As Double

    Set debtor = CreateObject("Scripting.Dictionary")
    Set retAgg = CreateObject("Scripting.Dictionary")
    Set discAgg = CreateObject("Scripting.Dictionary")
    Set taxAgg = CreateObject("Scripting.Dictionary")

    For r = 2 To lastEx
        If CStr(wsEx.Cells(r, COL_REPORT).Value) <> "RETURN_CREATED" Then GoTo NextRet
        cat = Trim$(CStr(wsEx.Cells(r, COL_CAT).Value & ""))
        sc = LCase$(Trim$(CStr(wsEx.Cells(r, COL_SCTYPE).Value & "")))
        stFrom = CStr(wsEx.Cells(r, COL_ST_FROM).Value & "")
        stTo = CStr(wsEx.Cells(r, COL_ST_TO).Value & "")
        sret = Nz(wsEx.Cells(r, COL_SALES_RET).Value)
        rdisc = Nz(wsEx.Cells(r, COL_RET_DISC).Value)
        rtax = Nz(wsEx.Cells(r, COL_RET_TAX).Value)

        Acc debtor, sc, sret + rdisc + rtax
        If retGL.Exists(cat) Then
            Acc retAgg, CStr(retGL(cat)) & "|" & AliasState(stFrom), Abs(sret)
        ElseIf Abs(sret) > 0.005 Then
            NoteUnmapped unmapped, "Sales Return GL", cat
        End If
        If discGL.Exists(cat) Then
            Acc discAgg, CStr(discGL(cat)) & "|" & AliasState(stFrom), Abs(rdisc)
        End If
        AccTax taxAgg, rtax, sret, stFrom, stTo
NextRet:
    Next r

    net = Nz(DictGet(debtor, "postpaid"))
    If net < 0 Then AddLine wsUp, voucherNo, GL_POSTPAID, "IN-OTH", "Sales", 0, -net, narr _
    Else: AddLine wsUp, voucherNo, GL_POSTPAID, "IN-OTH", "Sales", net, 0, narr

    net = Nz(DictGet(debtor, "prepaid"))
    If net < 0 Then AddLine wsUp, voucherNo, GL_PREPAID, "IN-OTH", "Sales", 0, -net, narr _
    Else: AddLine wsUp, voucherNo, GL_PREPAID, "IN-OTH", "Sales", net, 0, narr

    DumpAggDebitAbs wsUp, voucherNo, retAgg, "Sales", narr
    DumpAggCredit wsUp, voucherNo, discAgg, "Sales", narr
    DumpAggDebitAbs wsUp, voucherNo, taxAgg, "Sales", narr
End Sub

Private Sub BuildShipping(wsUp As Worksheet, wsEx As Worksheet, lastEx As Long, _
                          voucherNo As Long, narr As String)
    Const SHIP_GL As Long = 401051
    Dim debtor As Object, revCr As Object, revDr As Object, taxCr As Object, taxDr As Object
    Dim r As Long, report As String, sc As String, stFrom As String, stTo As String
    Dim ship As Double, stax As Double
    Dim tmpTax As Object, k As Variant, parts() As String, amt As Double

    Set debtor = CreateObject("Scripting.Dictionary")
    Set revCr = CreateObject("Scripting.Dictionary")
    Set revDr = CreateObject("Scripting.Dictionary")
    Set taxCr = CreateObject("Scripting.Dictionary")
    Set taxDr = CreateObject("Scripting.Dictionary")

    For r = 2 To lastEx
        report = CStr(wsEx.Cells(r, COL_REPORT).Value)
        If report <> "NONDIGITAL" And report <> "RETURN_CREATED" Then GoTo NextShip
        ship = Nz(wsEx.Cells(r, COL_SHIP).Value)
        stax = Nz(wsEx.Cells(r, COL_SHIP_TAX).Value)
        If Abs(ship) < 0.0000001 And Abs(stax) < 0.0000001 Then GoTo NextShip
        sc = LCase$(Trim$(CStr(wsEx.Cells(r, COL_SCTYPE).Value & "")))
        stFrom = CStr(wsEx.Cells(r, COL_ST_FROM).Value & "")
        stTo = CStr(wsEx.Cells(r, COL_ST_TO).Value & "")
        Acc debtor, sc, ship + stax
        If ship >= 0 Then
            Acc revCr, CStr(SHIP_GL) & "|" & AliasState(stFrom), ship
        Else
            Acc revDr, CStr(SHIP_GL) & "|" & AliasState(stFrom), Abs(ship)
        End If
        Set tmpTax = CreateObject("Scripting.Dictionary")
        AccTax tmpTax, stax, ship, stFrom, stTo
        For Each k In tmpTax.Keys
            amt = CDbl(tmpTax(k))
            If amt >= 0 Then Acc taxCr, CStr(k), amt Else Acc taxDr, CStr(k), Abs(amt)
        Next k
NextShip:
    Next r

    AddLine wsUp, voucherNo, GL_POSTPAID, "IN-OTH", "Sales", Nz(DictGet(debtor, "postpaid")), 0, narr
    AddLine wsUp, voucherNo, GL_PREPAID, "IN-OTH", "Sales", Nz(DictGet(debtor, "prepaid")), 0, narr
    DumpAggCredit wsUp, voucherNo, revCr, "Sales", narr
    DumpAggCredit wsUp, voucherNo, taxCr, "Sales", narr
    DumpAggDebit wsUp, voucherNo, revDr, "Sales", narr
    DumpAggDebit wsUp, voucherNo, taxDr, "Sales", narr
End Sub

Private Sub BuildPBO(wsUp As Worksheet, wsEx As Worksheet, lastEx As Long, _
                     voucherNo As Long, narr As String)
    Const REV_GL As Long = 400169
    Dim debtor As Double, r As Long, report As String
    Dim stFrom As String, stTo As String, rev As Double, tax As Double
    Dim revCr As Object, revDr As Object, taxCr As Object, taxDr As Object
    Dim tmpTax As Object, k As Variant, amt As Double

    Set revCr = CreateObject("Scripting.Dictionary")
    Set revDr = CreateObject("Scripting.Dictionary")
    Set taxCr = CreateObject("Scripting.Dictionary")
    Set taxDr = CreateObject("Scripting.Dictionary")
    debtor = 0

    For r = 2 To lastEx
        report = CStr(wsEx.Cells(r, COL_REPORT).Value)
        If report <> "PBO_SALES" And report <> "PBO_RETURN" Then GoTo NextPBO
        rev = Nz(wsEx.Cells(r, COL_PBO_REV).Value)
        tax = Nz(wsEx.Cells(r, COL_PBO_TAX).Value)
        If Abs(rev) < 0.0000001 And Abs(tax) < 0.0000001 Then GoTo NextPBO
        stFrom = CStr(wsEx.Cells(r, COL_ST_FROM).Value & "")
        stTo = CStr(wsEx.Cells(r, COL_ST_TO).Value & "")
        debtor = debtor + rev + tax
        If rev >= 0 Then
            Acc revCr, CStr(REV_GL) & "|" & AliasState(stFrom), rev
        Else
            Acc revDr, CStr(REV_GL) & "|" & AliasState(stFrom), Abs(rev)
        End If
        Set tmpTax = CreateObject("Scripting.Dictionary")
        AccTax tmpTax, tax, rev, stFrom, stTo
        For Each k In tmpTax.Keys
            amt = CDbl(tmpTax(k))
            If amt >= 0 Then Acc taxCr, CStr(k), amt Else Acc taxDr, CStr(k), Abs(amt)
        Next k
NextPBO:
    Next r

    If debtor >= 0 Then
        AddLine wsUp, voucherNo, GL_PBO_DEBTOR, "IN-OTH", "Sales", debtor, 0, narr
    Else
        AddLine wsUp, voucherNo, GL_PBO_DEBTOR, "IN-OTH", "Sales", 0, -debtor, narr
    End If
    DumpAggCredit wsUp, voucherNo, revCr, "Sales", narr
    DumpAggCredit wsUp, voucherNo, taxCr, "Sales", narr
    DumpAggDebit wsUp, voucherNo, revDr, "Sales", narr
    DumpAggDebit wsUp, voucherNo, taxDr, "Sales", narr
End Sub

Private Sub BuildPriceDrop(wsUp As Worksheet, wsEx As Worksheet, lastEx As Long, _
                           revGL As Object, voucherNo As Long, narr As String)
    Dim debtor As Object, revAgg As Object, taxAgg As Object
    Dim r As Long, cat As String, sc As String, stFrom As String, stTo As String
    Dim pd As Double, pdt As Double, gl As Variant, net As Double

    Set debtor = CreateObject("Scripting.Dictionary")
    Set revAgg = CreateObject("Scripting.Dictionary")
    Set taxAgg = CreateObject("Scripting.Dictionary")

    For r = 2 To lastEx
        If CStr(wsEx.Cells(r, COL_REPORT).Value) <> "PRICE_DROP" Then GoTo NextPD
        cat = Trim$(CStr(wsEx.Cells(r, COL_CAT).Value & ""))
        sc = LCase$(Trim$(CStr(wsEx.Cells(r, COL_SCTYPE).Value & "")))
        stFrom = CStr(wsEx.Cells(r, COL_ST_FROM).Value & "")
        stTo = CStr(wsEx.Cells(r, COL_ST_TO).Value & "")
        pd = Nz(wsEx.Cells(r, COL_PD).Value)
        pdt = Nz(wsEx.Cells(r, COL_PD_TAX).Value)
        If revGL.Exists(cat) Then gl = revGL(cat) ElseIf revGL.Exists("Mobile") Then gl = revGL("Mobile") Else gl = 401121
        Acc debtor, sc, pd + pdt
        Acc revAgg, CStr(gl) & "|" & AliasState(stFrom), pd
        AccTax taxAgg, pdt, pd, stFrom, stTo
NextPD:
    Next r

    net = Nz(DictGet(debtor, "postpaid"))
    If Abs(net) > 0.005 Then
        If net < 0 Then AddLine wsUp, voucherNo, GL_POSTPAID, "IN-OTH", "Sales", 0, -net, narr _
        Else: AddLine wsUp, voucherNo, GL_POSTPAID, "IN-OTH", "Sales", net, 0, narr
    End If
    net = Nz(DictGet(debtor, "prepaid"))
    If Abs(net) > 0.005 Then
        If net < 0 Then AddLine wsUp, voucherNo, GL_PREPAID, "IN-OTH", "Sales", 0, -net, narr _
        Else: AddLine wsUp, voucherNo, GL_PREPAID, "IN-OTH", "Sales", net, 0, narr
    End If
    DumpAggDebitAbs wsUp, voucherNo, revAgg, "Sales", narr
    DumpAggDebitAbs wsUp, voucherNo, taxAgg, "Sales", narr
End Sub

Private Sub BuildBuyerFee(wsUp As Worksheet, wsEx As Worksheet, lastEx As Long, _
                          voucherNo As Long, narr As String)
    Const REV_GL As Long = 401056
    Dim debtor As Object, revCr As Object, revDr As Object, taxCr As Object, taxDr As Object
    Dim r As Long, sc As String, stFrom As String, stTo As String
    Dim fee As Double, tax As Double, tmpTax As Object, k As Variant, amt As Double

    Set debtor = CreateObject("Scripting.Dictionary")
    Set revCr = CreateObject("Scripting.Dictionary")
    Set revDr = CreateObject("Scripting.Dictionary")
    Set taxCr = CreateObject("Scripting.Dictionary")
    Set taxDr = CreateObject("Scripting.Dictionary")

    For r = 2 To lastEx
        If CStr(wsEx.Cells(r, COL_REPORT).Value) <> "BUYER_FEE" Then GoTo NextBF
        sc = LCase$(Trim$(CStr(wsEx.Cells(r, COL_SCTYPE).Value & "")))
        stFrom = CStr(wsEx.Cells(r, COL_ST_FROM).Value & "")
        stTo = CStr(wsEx.Cells(r, COL_ST_TO).Value & "")
        fee = Nz(wsEx.Cells(r, COL_BUYER).Value)
        tax = Nz(wsEx.Cells(r, COL_BUYER_TAX).Value)
        Acc debtor, sc, fee + tax
        If fee >= 0 Then Acc revCr, CStr(REV_GL) & "|" & AliasState(stFrom), fee _
        Else: Acc revDr, CStr(REV_GL) & "|" & AliasState(stFrom), Abs(fee)
        Set tmpTax = CreateObject("Scripting.Dictionary")
        AccTax tmpTax, tax, fee, stFrom, stTo
        For Each k In tmpTax.Keys
            amt = CDbl(tmpTax(k))
            If amt >= 0 Then Acc taxCr, CStr(k), amt Else Acc taxDr, CStr(k), Abs(amt)
        Next k
NextBF:
    Next r

    AddLine wsUp, voucherNo, GL_POSTPAID, "IN-OTH", "Sales", Nz(DictGet(debtor, "postpaid")), 0, narr
    AddLine wsUp, voucherNo, GL_PREPAID, "IN-OTH", "Sales", Nz(DictGet(debtor, "prepaid")), 0, narr
    DumpAggCredit wsUp, voucherNo, revCr, "Sales", narr
    DumpAggCredit wsUp, voucherNo, taxCr, "Sales", narr
    DumpAggDebit wsUp, voucherNo, revDr, "Sales", narr
    DumpAggDebit wsUp, voucherNo, taxDr, "Sales", narr
End Sub

Private Sub BuildPrexoBumpup(wsUp As Worksheet, wsEx As Worksheet, lastEx As Long, _
                             voucherNo As Long, narr As String)
    Const REV_GL As Long = 400169
    Dim debtor As Double, r As Long, stFrom As String, stTo As String
    Dim rev As Double, tax As Double, revAgg As Object, taxAgg As Object

    Set revAgg = CreateObject("Scripting.Dictionary")
    Set taxAgg = CreateObject("Scripting.Dictionary")
    debtor = 0

    For r = 2 To lastEx
        If CStr(wsEx.Cells(r, COL_REPORT).Value) <> "PREXO_BUMPUP" Then GoTo NextBU
        stFrom = CStr(wsEx.Cells(r, COL_ST_FROM).Value & "")
        stTo = CStr(wsEx.Cells(r, COL_ST_TO).Value & "")
        rev = Nz(wsEx.Cells(r, COL_BUMP).Value)
        tax = Nz(wsEx.Cells(r, COL_BUMP_TAX).Value)
        debtor = debtor + rev + tax
        Acc revAgg, CStr(REV_GL) & "|" & AliasState(stFrom), rev
        AccTax taxAgg, tax, rev, stFrom, stTo
NextBU:
    Next r

    AddLine wsUp, voucherNo, GL_PBO_DEBTOR, "IN-OTH", "Sales", debtor, 0, narr
    DumpAggCredit wsUp, voucherNo, revAgg, "Sales", narr
    DumpAggCredit wsUp, voucherNo, taxAgg, "Sales", narr
End Sub


'===== Tax helpers =============================================================

Private Sub AccTax(agg As Object, ByVal tax As Double, ByVal baseAmt As Double, _
                   ByVal stFrom As String, ByVal stTo As String)
    Dim rate As Long, st As String
    If Abs(tax) < 0.0000001 Then Exit Sub
    rate = SnapRate(tax, baseAmt)
    st = AliasState(stFrom)
    If stFrom = stTo Then
        Acc agg, CStr(CgstGl(rate)) & "|" & st, tax / 2#
        Acc agg, CStr(SgstGl(rate)) & "|" & st, tax / 2#
    Else
        Acc agg, CStr(IgstGl(rate)) & "|" & st, tax
    End If
End Sub

Private Function SnapRate(ByVal tax As Double, ByVal baseAmt As Double) As Long
    Dim pct As Double
    If Abs(baseAmt) < 0.0000001 Then
        SnapRate = 18
        Exit Function
    End If
    pct = Abs(tax / baseAmt) * 100#
    If pct < 8# Then SnapRate = 5 Else SnapRate = 18
End Function

Private Function IgstGl(ByVal rate As Long) As Long
    Select Case rate
        Case 5: IgstGl = 225001
        Case 12: IgstGl = 225002
        Case 28: IgstGl = 225004
        Case Else: IgstGl = 225003
    End Select
End Function

Private Function CgstGl(ByVal rate As Long) As Long
    Select Case rate
        Case 5: CgstGl = 225006
        Case 12: CgstGl = 225007
        Case 28: CgstGl = 225009
        Case Else: CgstGl = 225008
    End Select
End Function

Private Function SgstGl(ByVal rate As Long) As Long
    Select Case rate
        Case 5: SgstGl = 225011
        Case 12: SgstGl = 225012
        Case 28: SgstGl = 225014
        Case Else: SgstGl = 225013
    End Select
End Function


'===== Aggregators / writers ===================================================

Private Sub Acc(d As Object, ByVal key As String, ByVal value As Double)
    If Len(key) = 0 Then Exit Sub
    If d.Exists(key) Then d(key) = CDbl(d(key)) + value Else d.Add key, value
End Sub

Private Function DictGet(d As Object, ByVal key As String) As Variant
    If d.Exists(key) Then DictGet = d(key) Else DictGet = 0
End Function

Private Sub DumpAggCredit(wsUp As Worksheet, voucherNo As Long, agg As Object, _
                          fn As String, narr As String)
    Dim k As Variant, parts() As String, amt As Double
    For Each k In agg.Keys
        amt = Round(CDbl(agg(k)), 2)
        If Abs(amt) < 0.005 Then GoTo NextK
        parts = Split(CStr(k), "|")
        AddLine wsUp, voucherNo, CLng(parts(0)), parts(1), fn, 0, amt, narr
NextK:
    Next k
End Sub

Private Sub DumpAggDebit(wsUp As Worksheet, voucherNo As Long, agg As Object, _
                         fn As String, narr As String)
    Dim k As Variant, parts() As String, amt As Double
    For Each k In agg.Keys
        amt = Round(CDbl(agg(k)), 2)
        If Abs(amt) < 0.005 Then GoTo NextK
        parts = Split(CStr(k), "|")
        AddLine wsUp, voucherNo, CLng(parts(0)), parts(1), fn, amt, 0, narr
NextK:
    Next k
End Sub

Private Sub DumpAggDebitAbs(wsUp As Worksheet, voucherNo As Long, agg As Object, _
                            fn As String, narr As String)
    Dim k As Variant, parts() As String, amt As Double
    For Each k In agg.Keys
        amt = Round(Abs(CDbl(agg(k))), 2)
        If Abs(amt) < 0.005 Then GoTo NextK
        parts = Split(CStr(k), "|")
        AddLine wsUp, voucherNo, CLng(parts(0)), parts(1), fn, amt, 0, narr
NextK:
    Next k
End Sub

Private Sub AddLine(wsUp As Worksheet, voucherNo As Long, account As Variant, _
                    state As String, fn As String, debit As Double, credit As Double, narr As String)
    Dim rr As Long
    debit = Round(debit, 2)
    credit = Round(credit, 2)
    If Abs(debit) < 0.005 And Abs(credit) < 0.005 Then Exit Sub
    If debit < 0 And credit = 0 Then credit = -debit: debit = 0
    If credit < 0 And debit = 0 Then debit = -credit: credit = 0

    gSl = gSl + 1
    rr = gSl + 1 ' header on row 1
    wsUp.Cells(rr, 1).Value = "AR-JV Sale"
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
    wsUp.Cells(rr, 16).Value = "No"
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

Private Sub WriteUnmapped(wsUn As Worksheet, unmapped As Object)
    Dim k As Variant, rr As Long, parts() As String
    rr = 2
    For Each k In unmapped.Keys
        parts = Split(CStr(k), "||")
        wsUn.Cells(rr, 1).Value = parts(0)
        wsUn.Cells(rr, 2).Value = parts(1)
        wsUn.Cells(rr, 3).Value = "Add GL Master mapping then re-run"
        rr = rr + 1
    Next k
End Sub

Private Sub NoteUnmapped(unmapped As Object, ByVal kind As String, ByVal key As String)
    Dim k As String
    k = kind & "||" & key
    If Not unmapped.Exists(k) Then unmapped.Add k, True
End Sub


'===== Utilities ===============================================================

Private Function AliasState(ByVal st As String) As String
    Select Case UCase$(Trim$(st))
        Case "IN-GJ": AliasState = "IN-GR"
        Case "IN-OR": AliasState = "IN-OS"
        Case Else: AliasState = Trim$(st)
    End Select
    If AliasState = "" Then AliasState = "IN-OTH"
End Function

Private Function Nz(ByVal v As Variant) As Double
    If IsError(v) Or IsEmpty(v) Or v = "" Then Nz = 0 Else Nz = CDbl(v)
End Function

Private Function MonthEndFromLabel(ByVal label As String) As Date
    Dim s As String, dt As Date
    s = Replace(Replace(Trim$(label), "'", ""), "’", "")
    On Error Resume Next
    dt = DateValue("1 " & Left$(s, Len(s) - 2) & " 20" & Right$(s, 2))
    If Err.Number <> 0 Then
        Err.Clear
        dt = DateSerial(Year(Date), Month(Date), 0) ' previous month end fallback
    End If
    On Error GoTo 0
    MonthEndFromLabel = DateSerial(Year(dt), Month(dt) + 1, 0)
End Function

Private Sub ClearSheetKeepHeader(ws As Worksheet, headerCols As Long)
    Dim last As Long
    last = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If last >= 2 Then ws.Rows("2:" & last).Delete
    ' Ensure standard Uploader headers if blank
    If ws.Name = "Uploader" And Len(ws.Cells(1, 1).Value & "") = 0 Then
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
