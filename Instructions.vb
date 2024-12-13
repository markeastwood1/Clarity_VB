'
' Macros for the sheet named "opti models"
' Copyright FICO (Fair Isaac Inc) 2023, 2024, 2025
'

Option Explicit

Private Sub Worksheet_Activate()
    On Error Resume Next

    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    Range("A1").Select

    '
    ' I'm using a Sub at the ThisWorkbook level so I can toggle
    ' this on/off easily for mass edits.
    '
    ThisWorkbook.Protect_This_Sheet

End Sub
