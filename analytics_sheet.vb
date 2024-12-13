'
' Macros for the sheet named "analytics"
' Copyright FICO (Fair Isaac Inc) 2023, 2024, 2025
'
Option Explicit
' this essentially makes string compare case-insensitive ... use binary for case-sensitive
Option Compare Text

Const SUPPORT_QUESTION As String = "C17"
Const BIZ_AREA_QUESTION As String = "C30"
Const PII_QUESTION_CELL As String = "C39"
Const CHAR_LIB_IN_SOLN_QUESTION As String = "C252"

Const ENGAGEMENT_RANGE As String = "B2:B17"
Const SUPPORT_RANGE As String = "B19:B26"
Const BUILD_RANGE As String = "B28:B41"
Const TRANSACTION_FRAUD_FIRST_RANGE As String = "C43:C67"
Const TRANSACTION_FRAUD_THIRD_RANGE As String = "C69:C93"
Const TRANSACTION_FRAUD_RETAIL_BANK_RANGE As String = "C95:C106"
Const LIFECYCLE_RANGE_ACCT As String = "C108:C136"
Const LIFECYCLE_RANGE_COLLECT As String = "C138:C166"
Const LIFECYCLE_RANGE_ORIGINATIONS As String = "C168:C201"
Const APP_FRAUD_RANGE As String = "C203:C235"
Const OTHER_RANGE As String = "C237:C248"
Const CHARS_REQUEST As String = "C250:C256"

' We need to know which cells should allow "multi-select" as its been implemented.
' We can not allow this for all drop-downs/pick-lists
Const FRAUD_FIRST_PORTFOLIOS As String = "C46"
Const FRAUD_FIRST_CARD_TYPES As String = "C47"
Const FRAUD_FIRST_DATA_SOURCES As String = "C60"
Const FRAUD_FIRST_TAGS As String = "C67"

Const FRAUD_THIRD_PORTFOLIOS As String = "C72"
Const FRAUD_THIRD_CARD_TYPES As String = "C73"
Const FRAUD_THIRD_DATA_SOURCES As String = "C86"
Const FRAUD_THIRD_TAGS As String = "C93"

Const RETAIL_FRAUD_PORTFOLIOS As String = "C97"

Const LIFECYCLE_ACCT_PORTFOLIOS As String = "C111"
Const LIFECYCLE_ACCT_INDUSTRY_TYPE As String = "C112"
Const LIFECYCLE_ACCT_RISK_CATEGORY As String = "C113"
Const LIFECYCLE_ACCT_PRODUCT_TYPE As String = "C114"
Const LIFECYCLE_ACCT_PRODUCT_SUBTYPE As String = "C115"
Const LIFECYCLE_ACCT_COLLATERALIZE As String = "C116"
Const LIFECYCLE_ACCT_TYPES_FEATURE As String = "C129"
Const LIFECYCLE_ACCT_DATA_SRC_FEATURE As String = "C130"

Const LIFECYCLE_COLLECT_PORTFOLIOS As String = "C141"
Const LIFECYCLE_COLLECT_INDUSTRY_TYPE As String = "C142"
Const LIFECYCLE_COLLECT_RISK_CATEGORY As String = "C143"
Const LIFECYCLE_COLLECT_PRODUCT_TYPE As String = "C144"
Const LIFECYCLE_COLLECT_PRODUCT_SUBTYPE As String = "C145"
Const LIFECYCLE_COLLECT_COLLATERALIZE As String = "C146"
Const LIFECYCLE_COLLECT_TYPES_FEATURE As String = "C159"
Const LIFECYCLE_COLLECT_DATA_SRC_FEATURE As String = "C160"

Const LIFECYCLE_ORIG_PORTFOLIOS As String = "C171"
Const LIFECYCLE_ORIG_INDUSTRY_TYPE As String = "C172"
Const LIFECYCLE_ORIG_RISK_CATEGORY As String = "C173"
Const LIFECYCLE_ORIG_PRODUCT_TYPE As String = "C174"
Const LIFECYCLE_ORIG_PRODUCT_SUBTYPE As String = "C175"
Const LIFECYCLE_ORIG_COLLATERALIZE As String = "C176"
Const LIFECYCLE_ORIG_PROD_STRUCT As String = "C177"
Const LIFECYCLE_ORIG_ACCT_TYPES_FEATURE As String = "C194"
Const LIFECYCLE_ORIG_DATA_SRC_FEATURE As String = "C195"
Const LIFECYCLE_ORIG_TAGS As String = "C201"

Const LIFECYCLE_APP_FRAUD_PORTFOLIOS As String = "C206"
Const LIFECYCLE_APP_FRAUD_INDUSTRY_TYPE As String = "C207"
Const LIFECYCLE_APP_FRAUD_RISK_CATEGORY As String = "C208"
Const LIFECYCLE_APP_FRAUD_PRODUCT_TYPE As String = "C209"
Const LIFECYCLE_APP_FRAUD_PRODUCT_SUBTYPE As String = "C210"
Const LIFECYCLE_APP_FRAUD_COLLATERALIZE As String = "C211"
Const LIFECYCLE_APP_FRAUD_DATA_SRC_FEATURE As String = "C228"
Const LIFECYCLE_APP_FRAUD_TAGS As String = "C235"

Private Sub Worksheet_Activate()
    On Error Resume Next

    'ActiveSheet.Rows.EntireRow.Hidden = True
    ToggleQuestions True, ENGAGEMENT_RANGE

    ' show the main rows
    CheckSupportQuestion
    CheckBizAreaQuestion ActiveSheet.Range(BIZ_AREA_QUESTION).Value2

    Dim cell As Range
    For Each cell In Application.ActiveSheet.UsedRange
        ThisWorkbook.CheckRequiredCell cell
    Next

    Range("A1").Select

    SetPIIValue ' using a formula was not working as expected so we have this function

    ThisWorkbook.Protect_This_Sheet
End Sub
Sub Worksheet_Change(ByVal Target As Range)
    On Error Resume Next

    Dim bizAreaChange As Range
    Dim supportChange As Range
    Dim rngValidatedCell As Range

    Target.AutoFit

    Set supportChange = Intersect(Target, Range(SUPPORT_QUESTION))
    If Not supportChange Is Nothing Then
        CheckSupportQuestion
    End If

    Set bizAreaChange = Intersect(Target, Range(BIZ_AREA_QUESTION))

    If Not bizAreaChange Is Nothing Then
        CheckBizAreaQuestion Range(BIZ_AREA_QUESTION).Value2
    End If

    'this gets a range that has ALL CELLS with VALIDATIONS (of any type)
    Set rngValidatedCell = Cells.SpecialCells(xlCellTypeAllValidation)

    ' if we have cells with validation on this sheet and the one that changed is a cell with validation
    'If Not rngValidatedCell Is Nothing Then
        ' if cell that changed has a validation
        If Not Intersect(Target, rngValidatedCell) Is Nothing Then
            ' if the validation type is list... then do the multi-select logic...
            If Target.Validation.Type = xlValidateList Then
                Call DropdownMultiSelect(Target)
            End If
        End If
    'End If

    ThisWorkbook.Protect_This_Sheet
    Application.ScreenUpdating = True

End Sub
Sub CheckSupportQuestion()
    On Error Resume Next
    ThisWorkbook.UnProtect_This_Sheet

    Select Case Range(SUPPORT_QUESTION).Value2

    Case "Yes"
        ToggleQuestions True, SUPPORT_RANGE
        ToggleQuestions False, BUILD_RANGE
        ToggleQuestions False, TRANSACTION_FRAUD_FIRST_RANGE
        ToggleQuestions False, TRANSACTION_FRAUD_THIRD_RANGE
        ToggleQuestions False, TRANSACTION_FRAUD_RETAIL_BANK_RANGE
        ToggleQuestions False, LIFECYCLE_RANGE_ACCT
        ToggleQuestions False, LIFECYCLE_RANGE_COLLECT
        ToggleQuestions False, LIFECYCLE_RANGE_ORIGINATIONS
        ToggleQuestions False, APP_FRAUD_RANGE
        ToggleQuestions False, OTHER_RANGE
        ToggleQuestions False, CHARS_REQUEST

    Case "No"
        ToggleQuestions False, SUPPORT_RANGE
        ToggleQuestions True, BUILD_RANGE
        ToggleQuestions False, TRANSACTION_FRAUD_FIRST_RANGE
        ToggleQuestions False, TRANSACTION_FRAUD_THIRD_RANGE
        ToggleQuestions False, TRANSACTION_FRAUD_RETAIL_BANK_RANGE
        ToggleQuestions False, LIFECYCLE_RANGE_ACCT
        ToggleQuestions False, LIFECYCLE_RANGE_COLLECT
        ToggleQuestions False, LIFECYCLE_RANGE_ORIGINATIONS
        ToggleQuestions False, APP_FRAUD_RANGE
        ToggleQuestions False, OTHER_RANGE
        ToggleQuestions False, CHARS_REQUEST
    Case ""
        ' for a new form we have no value for this to force them to answer
        ToggleQuestions False, SUPPORT_RANGE
        ToggleQuestions False, BUILD_RANGE
        ToggleQuestions False, TRANSACTION_FRAUD_FIRST_RANGE
        ToggleQuestions False, TRANSACTION_FRAUD_THIRD_RANGE
        ToggleQuestions False, TRANSACTION_FRAUD_RETAIL_BANK_RANGE
        ToggleQuestions False, LIFECYCLE_RANGE_ACCT
        ToggleQuestions False, LIFECYCLE_RANGE_COLLECT
        ToggleQuestions False, LIFECYCLE_RANGE_ORIGINATIONS
        ToggleQuestions False, APP_FRAUD_RANGE
        ToggleQuestions False, OTHER_RANGE
        ToggleQuestions False, CHARS_REQUEST
    Case Else
        MsgBox "Invalid Support selection - this shouldn't happen"
    End Select

    ThisWorkbook.Protect_This_Sheet
End Sub
Sub CheckBizAreaQuestion(ByVal businessArea As String)
    On Error Resume Next
    ThisWorkbook.UnProtect_This_Sheet

    'If Range(SUPPORT_QUESTION).Value2 = "Yes" Then
    '    Exit Sub
    'End If

    ' ensure that we're all hidden so that we only show the appropriate range after selection
    ToggleQuestions False, TRANSACTION_FRAUD_FIRST_RANGE
    ToggleQuestions False, TRANSACTION_FRAUD_THIRD_RANGE
    ToggleQuestions False, TRANSACTION_FRAUD_RETAIL_BANK_RANGE
    ToggleQuestions False, LIFECYCLE_RANGE_ACCT
    ToggleQuestions False, LIFECYCLE_RANGE_COLLECT
    ToggleQuestions False, LIFECYCLE_RANGE_ORIGINATIONS
    ToggleQuestions False, APP_FRAUD_RANGE
    ToggleQuestions False, OTHER_RANGE
    ToggleQuestions False, CHARS_REQUEST


    Application.ScreenUpdating = False
    'Application.ScreenUpdating = True

    Select Case Range(BIZ_AREA_QUESTION).Value2

    Case "Payment Fraud - 1st Party Card Fraud"
        ToggleQuestions True, TRANSACTION_FRAUD_FIRST_RANGE
        ActiveWindow.ScrollRow = Range(TRANSACTION_FRAUD_FIRST_RANGE).Row - 5

    Case "Payment Fraud - 3rd Party Card Fraud"
        ToggleQuestions True, TRANSACTION_FRAUD_THIRD_RANGE
        ActiveWindow.ScrollRow = Range(TRANSACTION_FRAUD_THIRD_RANGE).Row - 5

    Case "Payment Fraud - Retail Banking/DDA"
        ToggleQuestions True, TRANSACTION_FRAUD_RETAIL_BANK_RANGE
        ActiveWindow.ScrollRow = Range(TRANSACTION_FRAUD_RETAIL_BANK_RANGE).Row - 5

    Case "Application Fraud"
        ToggleQuestions True, APP_FRAUD_RANGE
        ActiveWindow.ScrollRow = Range(APP_FRAUD_RANGE).Row - 5

    Case "Lifecycle - Account Management"
        ToggleQuestions True, LIFECYCLE_RANGE_ACCT
        ActiveWindow.ScrollRow = Range(LIFECYCLE_RANGE_ACCT).Row - 5

    Case "Lifecycle - Originations"
        ToggleQuestions True, LIFECYCLE_RANGE_ORIGINATIONS
        ActiveWindow.ScrollRow = Range(LIFECYCLE_RANGE_ORIGINATIONS).Row - 5

    Case "Lifecycle - Collections"
        ToggleQuestions True, LIFECYCLE_RANGE_COLLECT
        ActiveWindow.ScrollRow = Range(LIFECYCLE_RANGE_COLLECT).Row - 5

    Case "Other"
        ToggleQuestions True, OTHER_RANGE
        ActiveWindow.ScrollRow = Range(OTHER_RANGE).Row - 8

    Case "Lifecycle Libraries - New Chars Request"
        ToggleQuestions True, CHARS_REQUEST
        checkLibraryInSolution
        ActiveWindow.ScrollRow = Range(CHARS_REQUEST).Row - 8

    Case ""
        'Need to handle the case where we have no selection - its not an error
        ActiveWindow.ScrollRow = Range("A1").Row
    Case Else
        MsgBox "Invalid selection - this shouldn't happen"
    End Select

    Application.ScreenUpdating = True
    Application.CalculateFullRebuild

    ThisWorkbook.Protect_This_Sheet
End Sub
Sub DropdownMultiSelect(Target As Range)
    On Error Resume Next

    Dim oldValue As String
    Dim newValue As String

    Application.EnableEvents = False

    With Target
        newValue = .Value2
        Application.Undo
        oldValue = .Value2
        .Value2 = newValue
    End With

    Application.EnableEvents = True

    'ThisWorkbook.UnProtect_This_Sheet
    Dim DelimiterType As String
    DelimiterType = ", "

    ' there are only certain cells that I want to allow multi-select...
    ' we've already filtered cells with validation (in general) and cells (more specifically) that have the "list validation" metadata
    ' but now we want to allow this only for specific cells... which is where the below is used
    Dim MULTI_SELECT_FRAUD_FIRST As Range
    Dim MULTI_SELECT_FRAUD_THIRD As Range
    Dim MULTI_SELECT_LIFE_ACCT As Range
    Dim MULTI_SELECT_LIFE_COLLECT As Range
    Dim MULTI_SELECT_LIFE_ORIG As Range
    Dim MULTI_SELECT_LIFE_APP_FRAUD As Range

    Dim MULTI_SELECT_CELLS As Range

    ' besides hoping to make maintenance simpler
    ' there are limits to the number of unions and the number of "line continuation"
    Set MULTI_SELECT_FRAUD_FIRST = Union(Range(FRAUD_FIRST_PORTFOLIOS), Range(FRAUD_FIRST_CARD_TYPES), Range(FRAUD_FIRST_DATA_SOURCES), Range(FRAUD_FIRST_TAGS), Range(RETAIL_FRAUD_PORTFOLIOS))
    Set MULTI_SELECT_FRAUD_THIRD = Union(Range(FRAUD_THIRD_PORTFOLIOS), Range(FRAUD_THIRD_CARD_TYPES), Range(FRAUD_THIRD_DATA_SOURCES), Range(FRAUD_THIRD_TAGS))
    Set MULTI_SELECT_LIFE_ACCT = Union(Range(LIFECYCLE_ACCT_PORTFOLIOS), Range(LIFECYCLE_ACCT_INDUSTRY_TYPE), Range(LIFECYCLE_ACCT_RISK_CATEGORY), Range(LIFECYCLE_ACCT_PRODUCT_TYPE), Range(LIFECYCLE_ACCT_PRODUCT_SUBTYPE), Range(LIFECYCLE_ACCT_COLLATERALIZE), Range(LIFECYCLE_ACCT_TYPES_FEATURE), Range(LIFECYCLE_ACCT_DATA_SRC_FEATURE))
    Set MULTI_SELECT_LIFE_COLLECT = Union(Range(LIFECYCLE_COLLECT_PORTFOLIOS), Range(LIFECYCLE_COLLECT_INDUSTRY_TYPE), Range(LIFECYCLE_COLLECT_RISK_CATEGORY), Range(LIFECYCLE_COLLECT_PRODUCT_TYPE), Range(LIFECYCLE_COLLECT_PRODUCT_SUBTYPE), Range(LIFECYCLE_COLLECT_COLLATERALIZE), Range(LIFECYCLE_COLLECT_TYPES_FEATURE), Range(LIFECYCLE_COLLECT_DATA_SRC_FEATURE))
    Set MULTI_SELECT_LIFE_ORIG = Union(Range(LIFECYCLE_ORIG_PORTFOLIOS), Range(LIFECYCLE_ORIG_INDUSTRY_TYPE), Range(LIFECYCLE_ORIG_RISK_CATEGORY), Range(LIFECYCLE_ORIG_PRODUCT_TYPE), Range(LIFECYCLE_ORIG_PRODUCT_SUBTYPE), Range(LIFECYCLE_ORIG_COLLATERALIZE), Range(LIFECYCLE_ORIG_PROD_STRUCT), Range(LIFECYCLE_ORIG_ACCT_TYPES_FEATURE), Range(LIFECYCLE_ORIG_DATA_SRC_FEATURE), Range(LIFECYCLE_ORIG_TAGS))
    Set MULTI_SELECT_LIFE_APP_FRAUD = Union(Range(LIFECYCLE_APP_FRAUD_PORTFOLIOS), Range(LIFECYCLE_APP_FRAUD_INDUSTRY_TYPE), Range(LIFECYCLE_APP_FRAUD_RISK_CATEGORY), Range(LIFECYCLE_APP_FRAUD_PRODUCT_TYPE), Range(LIFECYCLE_APP_FRAUD_PRODUCT_SUBTYPE), Range(LIFECYCLE_APP_FRAUD_COLLATERALIZE), Range(LIFECYCLE_APP_FRAUD_DATA_SRC_FEATURE), Range(LIFECYCLE_APP_FRAUD_TAGS))

    Set MULTI_SELECT_CELLS = Union(MULTI_SELECT_FRAUD_FIRST, MULTI_SELECT_FRAUD_THIRD, MULTI_SELECT_LIFE_ACCT, MULTI_SELECT_LIFE_COLLECT, MULTI_SELECT_LIFE_ORIG, MULTI_SELECT_LIFE_APP_FRAUD)

    On Error Resume Next

   'did the change happen where we are interested?
    Dim inMultiSelectCell As Range
    Set inMultiSelectCell = Intersect(Target, MULTI_SELECT_CELLS)

    If inMultiSelectCell Is Nothing Then
        Exit Sub
    End If

    If oldValue = "" Then
         Exit Sub
    Else
        If newValue = "" Then
            Exit Sub
        Else
            Application.EnableEvents = False
            Target.Value2 = oldValue & DelimiterType & newValue
            Application.EnableEvents = True
         End If
    End If

    'ThisWorkbook.Protect_This_Sheet
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
Sub ShowAnalyticsSummary()
    On Error Resume Next
    ThisWorkbook.UnProtect_This_Sheet
    ToggleTabByName buttonName:="AnalyticsSummary", tabname:="Analytics Summary"
    ThisWorkbook.Protect_This_Sheet
End Sub
Sub ToggleTabByName(buttonName As String, tabname As String)
    ' refactoring some code to do this in 1 place rather than in several places.
    ' we're only using buttons from the active sheet...
    On Error Resume Next
    ThisWorkbook.UnProtect_This_Sheet

    Dim btn As Shape
    Set btn = ActiveSheet.Shapes(buttonName)

    ThisWorkbook.Protect Structure:=False

    ' for any buttons that don't have an associated tab to show/hide
    If tabname = "" Then
        If Not ButtonStatus(btn) Then
            ButtonOn button:=btn
        Else
            ButtonOff button:=btn
        End If
        Exit Sub
    End If

    ' toggle visibility
    If Worksheets(tabname).Visible = xlSheetVisible Then
        ButtonOff button:=btn
        Worksheets(tabname).Visible = xlVeryHidden
    Else
        Worksheets(tabname).Visible = xlSheetVisible
        ButtonOn button:=btn
    End If

    ThisWorkbook.Protect_This_Sheet
    ThisWorkbook.Protect Structure:=True
End Sub
Private Sub ButtonOff(button As Shape)
    ' turn off the button color highlight
    On Error Resume Next
    ThisWorkbook.UnProtect_This_Sheet

    Dim GrayColor As Long
    Dim BlackColor As Long
    GrayColor = RGB(192, 192, 192)
    BlackColor = RGB(0, 0, 0)

    button.Fill.ForeColor.RGB = GrayColor
    button.TextFrame.Characters.Font.Color = BlackColor

    ThisWorkbook.Protect_This_Sheet
End Sub
Private Sub ButtonOn(button As Shape)
    On Error Resume Next
    ' turn ON the button color highlight
    ThisWorkbook.UnProtect_This_Sheet

    Dim GreenColor As Long
    Dim WhiteColor As Long
    GreenColor = RGB(51, 153, 102)
    WhiteColor = RGB(255, 255, 255)

    button.Fill.ForeColor.RGB = GreenColor
    button.TextFrame.Characters.Font.Color = WhiteColor

    ThisWorkbook.Protect_This_Sheet
End Sub
Private Function ButtonStatus(button As Shape) As Boolean
    ' return a boolean indicating if a button is highlighted or not
    On Error Resume Next
    ThisWorkbook.UnProtect_This_Sheet

    Dim GrayColor As Long
    GrayColor = RGB(192, 192, 192)

    If button.Fill.ForeColor.RGB = GrayColor Then
        ButtonStatus = False
    Else
        ButtonStatus = True
    End If

    ThisWorkbook.Protect_This_Sheet
End Function
Private Sub SetPIIValue()
    On Error Resume Next
    ThisWorkbook.UnProtect_This_Sheet

    Dim PII_Question As String
    Dim PII_Ext_Question As String

    PII_Question = Worksheets("Clarity").Range("C68").Value2
    PII_Ext_Question = Worksheets("Clarity").Range("C69").Value2

    If PII_Question = "Yes" Or PII_Ext_Question = "Yes" Then
        Worksheets("Analytics").Range(PII_QUESTION_CELL).Value2 = "Yes"
    End If

    If PII_Question = "No" And PII_Ext_Question = "No" Then
        Worksheets("Analytics").Range(PII_QUESTION_CELL).Value2 = "No"
    End If

    If PII_Question = "Uncertain" Or PII_Ext_Question = "Uncertain" Then
        Worksheets("Analytics").Range(PII_QUESTION_CELL).Value2 = "Uncertain"
    End If

    If PII_Question = "" And PII_Ext_Question = "" Then
        Worksheets("Analytics").Range(PII_QUESTION_CELL).Value2 = "Question wasn't answered on the main Clarity Tab."
    End If

    ThisWorkbook.Protect_This_Sheet
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
        solutionContains = solutionContains & "MFCL;"
    End If

    If hasOMFCL Then
        solutionContains = solutionContains & "O-MFCL;"
    End If

    If hasCBCL Then
        solutionContains = solutionContains & "CBCL;"
    End If

    ThisWorkbook.UnProtect_This_Sheet

    If hasMFCL Or hasOMFCL Or hasCBCL Then
        Range(CHAR_LIB_IN_SOLN_QUESTION).Value2 = solutionContains
    Else
        Range(CHAR_LIB_IN_SOLN_QUESTION).Value2 = "No FICO Libraries (cloud-edition) are currently specified in the solution. If needed for a model, update the Clarity tab."
    End If

    ThisWorkbook.Protect_This_Sheet
End Sub