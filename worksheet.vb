'
' Macros for all worksheets unless they have one with additional functionality.
' MOST sheets have this content and nothing else.
' Copyright FICO (Fair Isaac Inc) 2023, 2024, 2025
'

Option Explicit

Private Sub Worksheet_Activate()
    On Error Resume Next

    Dim Cell As Range
    For Each Cell In Application.ActiveSheet.UsedRange
        ThisWorkbook.CheckRequiredCell Cell
    Next

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
    
    Target.AutoFit

    ThisWorkbook.Protect_This_Sheet
End Sub