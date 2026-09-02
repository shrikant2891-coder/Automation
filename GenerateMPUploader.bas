Attribute VB_Name = "GenerateMPUploader"
Option Explicit

'===============================================================================
' MP Summary — Dynamic Uploader Generator (VBA)
'-------------------------------------------------------------------------------
' Import: Alt+F11 -> File -> Import File -> select this .bas -> Run GenerateMPUploader
'
' Sheets required: Summary, GL Backup, Uploader Format
' Optional:        Config (month / company code / voucher numbers)
'
' Recommended: run python3 build_mp_uploader.py for identical output on any OS.
' This macro shells to Python when available; otherwise shows instructions.
'===============================================================================

Public Sub GenerateMPUploader()
    Dim pyCmd As String
    Dim wbPath As String
    Dim result As Long

    On Error GoTo Fail
    Application.ScreenUpdating = False

    wbPath = ThisWorkbook.FullName
    pyCmd = "python3 """ & ThisWorkbook.Path & "\build_mp_uploader.py"" --input """ & wbPath & """ --output """ & wbPath & """"

    result = Shell(pyCmd, vbHide)
    If result = 0 Then
        MsgBox "Python was not found." & vbCrLf & vbCrLf & _
               "Run from terminal:" & vbCrLf & _
               "python3 build_mp_uploader.py --input ""MP Summary.xlsx"" --output ""MP Summary.xlsx""" & vbCrLf & vbCrLf & _
               "Ensure Summary and GL Backup sheets are populated first.", _
               vbExclamation, "MP Summary"
        GoTo Done
    End If

    Application.Wait Now + TimeValue("0:00:03")
    ThisWorkbook.Save
    MsgBox "MP Uploader regeneration requested via Python." & vbCrLf & _
           "Check Uploader Format and Control sheets for balanced vouchers.", _
           vbInformation, "MP Summary"
    GoTo Done

Fail:
    MsgBox "GenerateMPUploader failed: " & Err.Description, vbCritical

Done:
    Application.ScreenUpdating = True
End Sub
