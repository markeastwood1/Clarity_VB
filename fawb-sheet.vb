Option Explicit

Private Sub Worksheet_Activate()
    On Error Resume Next
    
    Dim cell As Range
    For Each cell In Application.ActiveSheet.UsedRange
        ThisWorkbook.CheckRequiredCell cell
    Next

    ' check this value so we can auto-select the dedicated DB
    ThisWorkbook.CheckEPC
    
    Dim result As Boolean
    result = ThisWorkbook.checkCapabilityInSolution("FICO Platform - Database Service")
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
    
    sheet_autofit

    ThisWorkbook.Protect_This_Sheet
End Sub
Sub sheet_autofit()
    On Error Resume Next
    Application.ActiveSheet.UsedRange.EntireRow.AutoFit
End Sub