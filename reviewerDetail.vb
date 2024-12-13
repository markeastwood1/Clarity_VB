'
' Macros for the sheet named "Reviewer Details"
' Copyright FICO (Fair Isaac Inc) 2023, 2024, 2025
'

Option Explicit

Private Sub Worksheet_Activate()
    On Error Resume Next

    ThisWorkbook.Worksheets("Reviewer Detail").Protect _
        password:=ThisWorkbook.myPassword, _
        AllowFormattingCells:=True, _
        AllowFormattingColumns:=True, _
        AllowFormattingRows:=True, _
        AllowInsertingColumns:=True, _
        AllowInsertingRows:=True

    Range("A1").Select
    '
    ' I'm using a Sub at the ThisWorkbook level so I can toggle
    ' this on/off easily for mass edits.
    '
    ThisWorkbook.Protect_This_Sheet
End Sub
Sub sheet_autofit()
    Application.ActiveSheet.UsedRange.EntireRow.AutoFit
End Sub
Sub Worksheet_Change(ByVal Target As Range)
    ThisWorkbook.UnProtect_This_Sheet
    sheet_autofit
    ThisWorkbook.Protect_This_Sheet
End Sub