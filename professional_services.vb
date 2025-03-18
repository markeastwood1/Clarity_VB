Option Explicit

Const CUST_APP_MGMT_QUESTION As String = "D10"
Const CUST_APP_MGMT_HOW As String = "D11"

Const SELF_SUPPORT_QUESTION As String = "D12"

Const YES_SS_CASE_MAP As String = "D13"
Const NO_SS_CASE_MAP As String = "D14"

Const PREFERRED_VENDOR_CASE As String = "D15"

Private Sub Worksheet_Activate()
    On Error Resume Next

    ShowHideManagingQuestion Range(CUST_APP_MGMT_QUESTION)
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

    ShowHideManagingQuestion Range(CUST_APP_MGMT_QUESTION)
    ShowHideSelfSupportQuestions Range(SELF_SUPPORT_QUESTION)
    ShowHidePreferredVendorQuestion Range(NO_SS_CASE_MAP)

    ThisWorkbook.Protect_This_Sheet
End Sub
Sub ShowHideManagingQuestion(Target As Range)
    On Error Resume Next
    ThisWorkbook.UnProtect_This_Sheet

    Dim appManagingAnswer As Range

    'did the change happen where we are interested?
    Set appManagingAnswer = Intersect(Target, Range(CUST_APP_MGMT_QUESTION))

  ' if the change happens outside of the range then ignore it
    If appManagingAnswer Is Nothing Then
        Exit Sub
    End If

    If appManagingAnswer.Value2 = "Internal Resources" Or appManagingAnswer.Value2 = "Third-party Consultants" Or appManagingAnswer.Value2 = "" Then
    ' hide the rows
        Range(CUST_APP_MGMT_HOW).EntireRow.Hidden = True
    End If

    If appManagingAnswer.Value2 = "Mix of these" Then
        ' show the rows
        Range(CUST_APP_MGMT_HOW).EntireRow.Hidden = False
    End If

    ThisWorkbook.Protect_This_Sheet
End Sub

Sub ShowHideSelfSupportQuestions(Target As Range)
    On Error Resume Next
    ThisWorkbook.UnProtect_This_Sheet

    Dim customerSelfSupport As Range

    'did the change happen where we are interested?
    Set customerSelfSupport = Intersect(Target, Range(SELF_SUPPORT_QUESTION))

  ' if the change happens outside of the range then ignore it
    If customerSelfSupport Is Nothing Then
        Exit Sub
    End If
    
    If customerSelfSupport.Value2 = "" Or customerSelfSupport.Value2 = "Uncertain" Then
    ' hide the rows
        Range(YES_SS_CASE_MAP).EntireRow.Hidden = True
        Range(NO_SS_CASE_MAP).EntireRow.Hidden = True
    End If
    
    Dim r As Range
    If customerSelfSupport.Value2 = "Yes" Then
        ' show the rows
         Range(YES_SS_CASE_MAP).EntireRow.Hidden = False
         Range(NO_SS_CASE_MAP).EntireRow.Hidden = True

    End If

    If customerSelfSupport.Value2 = "No" Then
        Range(YES_SS_CASE_MAP).EntireRow.Hidden = True
        Range(NO_SS_CASE_MAP).EntireRow.Hidden = False
    End If

    ThisWorkbook.Protect_This_Sheet
End Sub

Sub ShowHidePreferredVendorQuestion(Target As Range)
    On Error Resume Next
    ThisWorkbook.UnProtect_This_Sheet

    Dim preferredVendorAnswer As Range

    'did the change happen where we are interested?
    Set preferredVendorAnswer = Intersect(Target, Range(NO_SS_CASE_MAP))

  ' if the change happens outside of the range then ignore it
    If preferredVendorAnswer Is Nothing Then
        Exit Sub
    End If

    If preferredVendorAnswer.Value2 = "Internal Resources" Or preferredVendorAnswer.Value2 = "" Then
    ' hide the rows
        Range(PREFERRED_VENDOR_CASE).EntireRow.Hidden = True
    Else
        ' show the rows
        Range(PREFERRED_VENDOR_CASE).EntireRow.Hidden = False
    End If
    
    ThisWorkbook.Protect_This_Sheet
End Sub
