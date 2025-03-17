Option Explicit

Const PS_MANAGING_QUESTION As String = "D10"
Const RANGE_PS_MANAGING As String = "D11"

Const SELF_SUPPORT_QUESTION As String = "D12"
Const RANGE_SS_QUESTIONS_RANGE As String = "D13:D14"
Const YES_SS_CASE_MAP As String = "D13"

Const NO_SS_CASE_MAP As String = "D14"

Const PREFERRED_VENDOR_CASE As String = "D15"

Private Sub Worksheet_Activate()
    On Error Resume Next

    ShowHideModelsQuestions Range(PS_MANAGING_QUESTION)
    ShowHideSelfSupportQuestions Range(SELF_SUPPORT_QUESTION)
    ShowHidePreferredVendorQuestion Range(NO_SS_CASE_MAP)
    
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
    
    ShowHideModelsQuestions Range(PS_MANAGING_QUESTION)
    ShowHideSelfSupportQuestions Range(SELF_SUPPORT_QUESTION)
    ShowHidePreferredVendorQuestion Range(NO_SS_CASE_MAP)
    
    ThisWorkbook.Protect_This_Sheet
End Sub
Sub ShowHideModelsQuestions(Target As Range)
    On Error Resume Next
    ThisWorkbook.UnProtect_This_Sheet

    Dim modelsAnswer As Range

    'did the change happen where we are interested?
    Set modelsAnswer = Intersect(Target, Range(PS_MANAGING_QUESTION))

  ' if the change happens outside of the range then ignore it
    If modelsAnswer Is Nothing Then
        Exit Sub
    End If
    
    Dim r As Range
    If modelsAnswer.Value2 <> "Mix of these" Or modelsAnswer.Value2 = "" Then
    ' hide the rows
         For Each r In Range(RANGE_PS_MANAGING).Rows
             r.EntireRow.Hidden = True
         Next r
         'Worksheets("Analytics").Visible = False
    
    ElseIf modelsAnswer.Value2 = "Mix of these" Then
        ' show the rows
        For Each r In Range(RANGE_PS_MANAGING).Rows
             r.EntireRow.Hidden = False
         Next r
    End If

    ThisWorkbook.Protect_This_Sheet
End Sub

Sub ShowHideSelfSupportQuestions(Target As Range)
    On Error Resume Next
    ThisWorkbook.UnProtect_This_Sheet

    Dim modelsAnswer As Range

    'did the change happen where we are interested?
    Set modelsAnswer = Intersect(Target, Range(SELF_SUPPORT_QUESTION))

  ' if the change happens outside of the range then ignore it
    If modelsAnswer Is Nothing Then
        Exit Sub
    End If

    If modelsAnswer.Value2 = "" Then
    ' hide the rows
         Dim r As Range
         For Each r In Range(RANGE_SS_QUESTIONS_RANGE).Rows
             r.EntireRow.Hidden = True
         Next r
    
    ElseIf modelsAnswer.Value2 = "Yes" Then
        ' show the rows
        For Each r In Range(YES_SS_CASE_MAP).Rows
             r.EntireRow.Hidden = False
         Next r
         
         For Each r In Range(NO_SS_CASE_MAP).Rows
            r.EntireRow.Hidden = True
         Next r
         
    ElseIf modelsAnswer.Value2 = "No" Then
        For Each r In Range(NO_SS_CASE_MAP).Rows
             r.EntireRow.Hidden = False
         Next r
         
         For Each r In Range(YES_SS_CASE_MAP).Rows
            r.EntireRow.Hidden = True
         Next r
         
    End If

    ThisWorkbook.Protect_This_Sheet
End Sub

Sub ShowHidePreferredVendorQuestion(Target As Range)
    On Error Resume Next
    ThisWorkbook.UnProtect_This_Sheet

    Dim modelsAnswer As Range

    'did the change happen where we are interested?
    Set modelsAnswer = Intersect(Target, Range(NO_SS_CASE_MAP))

  ' if the change happens outside of the range then ignore it
    If modelsAnswer Is Nothing Then
        Exit Sub
    End If

    If modelsAnswer.Value2 = "Internal Resources" Or modelsAnswer.Value2 = "" Then
    ' hide the rows
         Dim r As Range
         For Each r In Range(PREFERRED_VENDOR_CASE).Rows
             r.EntireRow.Hidden = True
         Next r
    
    Else
        ' show the rows
        For Each r In Range(PREFERRED_VENDOR_CASE).Rows
             r.EntireRow.Hidden = False
         Next r
    End If

    ThisWorkbook.Protect_This_Sheet
End Sub
