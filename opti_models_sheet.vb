'
' Macros for the sheet named "opti models"
' Copyright FICO (Fair Isaac Inc) 2023, 2024, 2025
'

Option Explicit
' string compare case-insensitive
Option Compare Text

Const SOLUTION_TYPE_QUESTION As String = "C4"
Const REQUIREMENTS_QUESTION As String = "C5"
Const DECISION_AREA_QUESTION As String = "C6"

Const ALL_REQUESTS As String = "B2:B5"
Const REQUIREMENTS_BOS_NOTES As String = "B8:B10"
Const REQUIREMENTS_CONSULTING_SUPPORT As String = "B12:B14"
Const REQUIREMENTS_TRAINING_SUPPORT As String = "B16:B18"
Const SOLUTION_TYPE_NON_FS_OPTI As String = "B20:B50"
Const SOLUTION_TYPE_FS_OPTI_DO As String = "B53:B106"
Const DECISION_AREA_PERSONAL_LOANS As String = "B109:B128"
Const DECISION_AREA_CARDS As String = "B130:B149"
Const DECISION_AREA_MORTGAGE As String = "B151:B169"
Const DECISION_AREA_DEPOSIT As String = "B171:B193"
Const DECISION_AREA_COLLECTIONS As String = "B195:B214"
Const DECISION_AREA_ADS As String = "B216:B236"
Const DECISION_AREA_OTHER As String = "B238:B250"
Const SOLUTION_TYPE_OTHER = "B252:B258"
Const REQUIREMENTS_OTHER As String = "B260:B262"

Private Sub Worksheet_Activate()
    On Error Resume Next

    ThisWorkbook.UnProtect_This_Sheet

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

    Dim solnTypeQuestion As Range
    Dim requirementsTypeQuestion As Range
    Dim decisionAreaQuestion As Range

    Dim solnValue As String
    Dim requirementsValue As String
    Dim decisionAreaValue As String

    ThisWorkbook.UnProtect_This_Sheet

    ' check to see if one of these values has changed
    Set solnTypeQuestion = Intersect(Target, Range(SOLUTION_TYPE_QUESTION))
    Set requirementsTypeQuestion = Intersect(Target, Range(REQUIREMENTS_QUESTION))
    Set decisionAreaQuestion = Intersect(Target, Range(DECISION_AREA_QUESTION))

    solnValue = Range(SOLUTION_TYPE_QUESTION).Value2
    requirementsValue = Range(REQUIREMENTS_QUESTION).Value2
    decisionAreaValue = Range(DECISION_AREA_QUESTION).Value2

    ' if the user changed either the solution type, the analytic requirements or the decision area
    ' see what we need to show for the user... but check we don't have stale data in one of the inputs.
    If Not solnTypeQuestion Is Nothing Or Not requirementsTypeQuestion Is Nothing Or Not decisionAreaQuestion Is Nothing Then

        ' clear this out if it doesn't make sense to have a value but don't create an event loop!
        If IsEmpty(solnValue) Or Len(solnValue) = 0 Then
            Application.EnableEvents = False
            Range(REQUIREMENTS_QUESTION).Value = ""
            Application.EnableEvents = True
        End If

        If IsEmpty(requirementsValue) Or Len(requirementsValue) = 0 Then
            Application.EnableEvents = False
            Range(DECISION_AREA_QUESTION).Value = ""
            Application.EnableEvents = True
        End If

        ShowSection solnValue, requirementsValue
    Else
        ' the change was likely some other cells and we'll just ensure that we autoFit in case the text is large
        Target.AutoFit
    End If

    ThisWorkbook.Protect_This_Sheet
End Sub
Sub ShowSection(ByVal SolutionType As String, SolnRequirements As String)
    On Error Resume Next

    ' We need values or we can't proceed
    If IsEmpty(SolutionType) Then
        MsgBox "Solution Type is empty, nothing to show"
        Exit Sub
    End If

    Dim decisionAreaValue As String
    decisionAreaValue = Range(DECISION_AREA_QUESTION).Value2

    ThisWorkbook.UnProtect_This_Sheet
    'Application.ScreenUpdating = False

    ToggleQuestions True, ALL_REQUESTS

    ToggleQuestions False, DECISION_AREA_QUESTION
    ToggleQuestions False, REQUIREMENTS_QUESTION

    ToggleQuestions False, SOLUTION_TYPE_NON_FS_OPTI
    ToggleQuestions False, SOLUTION_TYPE_FS_OPTI_DO
    ToggleQuestions False, SOLUTION_TYPE_OTHER

    ToggleQuestions False, REQUIREMENTS_BOS_NOTES
    ToggleQuestions False, REQUIREMENTS_CONSULTING_SUPPORT
    ToggleQuestions False, REQUIREMENTS_TRAINING_SUPPORT
    ToggleQuestions False, DECISION_AREA_PERSONAL_LOANS
    ToggleQuestions False, DECISION_AREA_CARDS
    ToggleQuestions False, DECISION_AREA_MORTGAGE
    ToggleQuestions False, DECISION_AREA_DEPOSIT
    ToggleQuestions False, DECISION_AREA_COLLECTIONS
    ToggleQuestions False, DECISION_AREA_ADS
    ToggleQuestions False, DECISION_AREA_OTHER
    ToggleQuestions False, REQUIREMENTS_OTHER

    If SolutionType = "BOS" Then
        ToggleQuestions True, REQUIREMENTS_BOS_NOTES
    End If

    If SolutionType = "Other" Then
        ToggleQuestions True, SOLUTION_TYPE_OTHER
    End If

    If SolutionType = "DO" Or SolutionType = "Financial Services" Or SolutionType = "NOT Financial Services" Then
        ToggleQuestions True, REQUIREMENTS_QUESTION

        If SolnRequirements = "Analytic Consulting Support Existing Client Solution" Then
            ToggleQuestions True, REQUIREMENTS_CONSULTING_SUPPORT
        End If

        If SolnRequirements = "Software Training" Then
            ToggleQuestions True, REQUIREMENTS_TRAINING_SUPPORT
        End If

    End If

    If SolutionType = "NOT Financial Services" And SolnRequirements = "Development of Client Solution" Then
        ToggleQuestions True, SOLUTION_TYPE_NON_FS_OPTI
    End If

    ' ask for decision area
    If (SolutionType = "DO" Or SolutionType = "Financial Services") And SolnRequirements = "Development of Client Solution" Then
        ToggleQuestions True, SOLUTION_TYPE_FS_OPTI_DO
        ToggleQuestions True, DECISION_AREA_QUESTION
    End If

    If (SolutionType = "DO" Or SolutionType = "Financial Services") And SolnRequirements = "Development of Client Solution" And Not IsEmpty(decisionAreaValue) Then
        Select Case Range(DECISION_AREA_QUESTION).Value2
            Case ""
            Case "Personal Loans Pricing"
                ToggleQuestions True, DECISION_AREA_PERSONAL_LOANS
            Case "Personal Loans Amount"
                ToggleQuestions True, DECISION_AREA_PERSONAL_LOANS
            Case "Personal Loans Pricing & Amount"
                ToggleQuestions True, DECISION_AREA_PERSONAL_LOANS
            Case "Credit Cards Initial Credit Limit"
                ToggleQuestions True, DECISION_AREA_CARDS
            Case "Credit Cards Credit Limit Increase"
                ToggleQuestions True, DECISION_AREA_CARDS
            Case "Mortgage Pricing"
                ToggleQuestions True, DECISION_AREA_MORTGAGE
            Case "Deposit Pricing"
                ToggleQuestions True, DECISION_AREA_DEPOSIT
            Case "Collections"
                ToggleQuestions True, DECISION_AREA_COLLECTIONS
            Case "ADS"
                ToggleQuestions True, DECISION_AREA_ADS
            Case "Other"
                ToggleQuestions True, DECISION_AREA_OTHER
        End Select
    End If

    If (SolutionType = "DO" Or SolutionType = "Financial Services" Or SolutionType = "NOT Financial Services") And SolnRequirements = "Other" Then
        ToggleQuestions True, REQUIREMENTS_OTHER
    End If

    'Application.ScreenUpdating = True
    'Application.CalculateFullRebuild

    ThisWorkbook.Protect_This_Sheet
End Sub
Sub ToggleQuestions(visibility As Boolean, sectionName As String)
    On Error Resume Next
    ThisWorkbook.UnProtect_This_Sheet

    If visibility = False Then
    ' hide the rows
         Range(sectionName).EntireRow.Hidden = True
    Else
         Range(sectionName).EntireRow.Hidden = False
    End If

    ThisWorkbook.Protect_This_Sheet
End Sub