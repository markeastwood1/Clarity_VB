'
' Macros for the sheet named "DB Service"
' Copyright FICO (Fair Isaac Inc) 2023, 2024, 2025
'
Option Explicit

Private Sub Worksheet_Activate()
    On Error Resume Next

    Dim Cell As Range
    For Each Cell In Application.ActiveSheet.UsedRange
        ThisWorkbook.CheckRequiredCell Cell
    Next

    ' check this value so we can auto-select the dedicated DB
    ThisWorkbook.CheckEPC

    Dim result As Boolean
    result = ThisWorkbook.checkCapabilityInSolution("FICO Applications Workbench - Cloud Edition")
    IF result = True then
        ThisWorkbook.CrossCheckDedicated
        ThisWorkbook.CrossCheckSize
    End If

    Range("A1").Select

    '
    ' I'm using a Sub at the ThisWorkbook level so I can toggle
    ' this on/off easily for mass edits.
    '
    ThisWorkbook.Protect_This_Sheet

End Sub
Sub Worksheet_Change(ByVal Target As Range)
    On Error Resume Next
    
    ThisWorkbook.UnProtect_This_Sheet
    
    Dim Cell As Range
    For Each Cell In Application.ActiveSheet.UsedRange
        ThisWorkbook.CheckRequiredCell Cell
    Next
    sheet_autofit

    ThisWorkbook.Protect_This_Sheet
End Sub
Sub sheet_autofit()
    On Error Resume Next
    Application.ActiveSheet.UsedRange.EntireRow.AutoFit
End Sub
