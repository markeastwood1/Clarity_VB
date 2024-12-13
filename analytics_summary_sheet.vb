'
' Macros for the sheet named "analytics summary"
' Copyright FICO (Fair Isaac Inc) 2023, 2024, 2025
'
Option Explicit

Const CHAR_LIB_IN_SOLN_QUESTION As String = "D16"

Private Sub Worksheet_Activate()
    On Error Resume Next

    Dim Cell As Range
    For Each Cell In Application.ActiveSheet.UsedRange
        ThisWorkbook.CheckRequiredCell Cell
    Next

    Range("A1").Select

    checkLibraryInSolution

    '
    ' I'm using a Sub at the "ThisWorkbook" level so I can toggle
    ' this on/off easily for mass edits.
    '
    ThisWorkbook.Protect_This_Sheet

End Sub
Sub Worksheet_Change(ByVal Target As Range)
    On Error Resume Next

    Target.AutoFit

End Sub
Private Sub checkLibraryInSolution()
    On Error Resume Next

    Dim hasMFCL As Boolean
    Dim hasOMFCL As Boolean
    Dim hasCBCL As Boolean
    Dim solutionContains As String

    solutionContains = ""

    hasCBCL = ThisWorkbook.checkCapabilityInSolution("FICO Credit Bureau Characteristic Library")
    hasMFCL = ThisWorkbook.checkCapabilityInSolution("FICO Platform - Master File Characteristics Library - Cloud Edition")
    'this won't actually work until it canactually be selected
    hasOMFCL = ThisWorkbook.checkCapabilityInSolution("OMFCL")

    If hasMFCL Then
        solutionContains = solutionContains & "MFCL "
    End If

    If hasOMFCL Then
        solutionContains = solutionContains & " O-MFCL "
    End If

    If hasCBCL Then
        solutionContains = solutionContains & " CBCL "
    End If

    ThisWorkbook.UnProtect_This_Sheet

    If hasMFCL Or hasOMFCL Or hasCBCL Then
        Range(CHAR_LIB_IN_SOLN_QUESTION).Value2 = solutionContains
    Else
        Range(CHAR_LIB_IN_SOLN_QUESTION).Value2 = "None licensed in this opportunity."
    End If

    ThisWorkbook.Protect_This_Sheet
End Sub
