'
' Macros for all worksheets unless they have one with additional functionality.
' MOST sheets have this content and nothing else.
' Copyright FICO (Fair Isaac Inc) 2023, 2024, 2025
'

Option Explicit

Const DIS_ADVANCED_QUESTION As String = "C10"
Const DIS_UI_QUESTION As String = "C11"


Private Sub Worksheet_Activate()
    On Error Resume Next
    
    ShowHideDIS_UI_ADVANCED_Question Range(DIS_ADVANCED_QUESTION)

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

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error Resume Next
    ThisWorkbook.UnProtect_This_Sheet

    Target.AutoFit
    ShowHideDIS_UI_ADVANCED_Question Target

    ThisWorkbook.Protect_This_Sheet
End Sub

Private Sub ShowHideDIS_UI_ADVANCED_Question(Target As Range)
    On Error Resume Next
    ThisWorkbook.UnProtect_This_Sheet

    Dim disAdvanced As Range

    'did the change happen where we are interested?
    Set disAdvanced = Intersect(Target, Range(DIS_ADVANCED_QUESTION))

  ' if the change happens outside of the range then ignore it
    If disAdvanced Is Nothing Then
        Exit Sub
    End If

    Dim r As Range
    If disAdvanced.Value2 = "No" Or disAdvanced.Value2 = "" Or disAdvanced.Value2 = "Uncertain" Then
    ' hide the rows
         For Each r In Range(DIS_UI_QUESTION).Rows
             r.EntireRow.Hidden = True
         Next r
         'Worksheets("DIS").Visible = False
    Else
        ' show the rows
        For Each r In Range(DIS_UI_QUESTION).Rows
             r.EntireRow.Hidden = False
         Next r
    End If

    ThisWorkbook.Protect_This_Sheet
End Sub
