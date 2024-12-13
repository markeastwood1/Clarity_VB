'
' Macros for the sheet named "Clarity" -- This is the main sheet of the whole thing.
' Copyright FICO (Fair Isaac Inc) 2023, 2024, 2025
'
Option Explicit

'a bastardized way to have a global of this range so I can edit it in only 1 place
Private Function RangeCapability() As Range
    Set RangeCapability = ThisWorkbook.Sheets("Clarity").Range("A4:A18")
End Function
Private Function RangeOptional() As Range
    Set RangeOptional = ThisWorkbook.Sheets("Clarity").Range("B4:B18")
End Function
Private Function RangeNew() As Range
    Set RangeNew = ThisWorkbook.Sheets("Clarity").Range("C4:C18")
End Function
Private Function RangeUseCase() As Range
    Set RangeUseCase = ThisWorkbook.Sheets("Clarity").Range("C25")
End Function
Private Function RangeModels() As Range
    Set RangeModels = ThisWorkbook.Sheets("Clarity").Range("C36")
End Function
Private Function RangeModelsFICO() As Range
    Set RangeModelsFICO = ThisWorkbook.Sheets("Clarity").Range("C37")
End Function
Private Function RangeModelQuestions() As Range
    Set RangeModelQuestions = ThisWorkbook.Sheets("Clarity").Range("A37:A39")
End Function
Private Function RangeProducts() As Range ' if you add or remove products from the list update the range here.
    Set RangeProducts = ThisWorkbook.Sheets("ProductsMap").Range("A2:B80")
End Function
Private Function RangeDependencies() As Range ' if you add or remove products from the list update the range here.
    Set RangeDependencies = ThisWorkbook.Sheets("ProductsMap").Range("A2:D80")
End Function
Private Function RangeUseCases() As Range
    Set RangeUseCases = ThisWorkbook.Sheets("UseCaseMap").Range("A2:B47")
End Function
Private Function RangeLastUser() As Range
    Set RangeLastUser = ThisWorkbook.Sheets("Clarity").Range("C2")
End Function
Private Function RangePSMigration() As Range
    Set RangePSMigration = ThisWorkbook.Sheets("Clarity").Range("A32")
End Function
Private Function RangeAutoFit() As Range
    Set RangeAutoFit = ThisWorkbook.Sheets("Clarity").Range("C20:D78")
End Function
Sub Worksheet_Activate()
    ' process the cells marked as required

    Application.EnableEvents = False
    ThisWorkbook.Protect Structure:=False
    ThisWorkbook.UnProtect_This_Sheet

    Dim Cell As Range
    For Each Cell In Application.ActiveSheet.UsedRange
        ThisWorkbook.CheckRequiredCell Cell
    Next

    Worksheets("Clarity").Range("A1").ColumnWidth = 50
    Worksheets("Clarity").Range("B1").ColumnWidth = 60
    Worksheets("Clarity").Range("C1").ColumnWidth = 65
    Worksheets("Clarity").Range("D1").ColumnWidth = 65

    ThisWorkbook.Protect Structure:=True

    Application.EnableEvents = True
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    ThisWorkbook.Worksheets("CLARITY").Range("A4").Activate
    ThisWorkbook.Protect_This_Sheet
End Sub
Sub Worksheet_Change(ByVal Target As Range)
' Worksheet_Change() gets in-cell changes and Worksheet_SelectionChange only fires when the selection changes
'
    On Error Resume Next

    ' attempt to handle the case where someone deletes multiple capabilities from the
    ' at the same time
    If Target.Count = 1 Then
        ShowCapabilityTabs Target
    Else

        Dim iCell As Range
        For Each iCell In Target.Cells
            ShowCapabilityTabs iCell
        Next iCell

    End If

    ShowHideModelsQuestions Target
    ShowHideModelsTab Target
    ShowUseCaseTabs Target

End Sub
Sub ShowCapabilityTabs(Target As Range)
    ' a hidden mapping tab associates the Salesforce product name to the simpler capability tab name
    ' need to check if they have special value for EOL or NA
    ' We also allow up to 2 dependencies PER capability such as Decision Modeler and ADM

    On Error Resume Next
    ThisWorkbook.Protect Structure:=False
    Dim capabilityUpdated As Range
    Dim aDependency As String

    ' show or hide the tabs

    Set capabilityUpdated = Intersect(Target, RangeCapability())

  ' if the change happens outside of the range then ignore it
    If capabilityUpdated Is Nothing Then
        Exit Sub
    End If

    Dim oldValue As String, newValue As String, tabname As String

    ' this is can grab the current value, then undo to get the old value and then put the new value back
    Application.EnableEvents = False

    With Target
        newValue = .Value2
        Application.Undo
        oldValue = .Value2
        .Value2 = newValue
    End With

    Application.EnableEvents = True
   ' this is can grab the current value, then undo to get the old value and then put the new value back

    ThisWorkbook.Protect Structure:=False

    ' if NewValue <> OldValue then we potentially need to hide the old tab.
    ' also need to handle dependencies
    If Not newValue = oldValue Then

        ' hide the one we don't need any more...
        ' if the oldValue is blank then skip this part
        If Not oldValue = "" And Not IsNull(oldValue) Then
            tabname = ProductSelectMapping(oldValue)

            ' if we need to hide something check the dependencies too
            If Not tabname = "EOL" Then
                Worksheets(tabname).Visible = xlVeryHidden

                aDependency = ProductDependency_1(oldValue)
                DeleteDependency (aDependency)

                aDependency = ProductDependency_2(oldValue)
                DeleteDependency (aDependency)
            Else
                Target.Interior.ColorIndex = xlColorIndexNone
                Target.Font.Color = RGB(0, 0, 0) ' Black
            End If
        End If
    End If

    ' if we get here we're not empty/NULL (as in someone deleted a selection made earlier)
    ' we've finished with hiding something that was deleted now we deal with showing something that's
    ' a new selection
    If Not newValue = "" And Not IsNull(newValue) Then
        ' need to grab the tab name not the capability name because we have the capability --> tab mapping
        ' I am assuming that we get back something... anything - blanks in the lookup table are not acceptable

        tabname = ProductSelectMapping(newValue) ' this means the dependencies have to use the ProductsMap from Column A

        ' Added this behavior for when we have selectable items that are actually EOL.
        If tabname = "EOL" Then
            Target.Interior.ColorIndex = 3 '(which is red)
            Target.Font.Color = RGB(255, 255, 255) ' White
            MsgBox "That Selection is EOL"
        Else
            Target.Interior.ColorIndex = xlColorIndexNone ' no shading for the cell
            Target.Font.Color = RGB(0, 0, 0) ' Black
            Worksheets(tabname).Visible = xlSheetVisible

            ' now go look for a dependency to add to the list
            aDependency = ProductDependency_1(newValue)

            ' Need to check to see if a dependency exists ...
            If Not aDependency = "" And Not IsNull(aDependency) Then
                InsertDependency (aDependency)
            End If

            aDependency = ProductDependency_2(newValue)

            If Not aDependency = "" And Not IsNull(aDependency) Then
                InsertDependency (aDependency)
            End If

        End If
    Else
        ' ensure we don't leave a cell red that is now empty and was previously a EOL selection
        Target.Interior.ColorIndex = xlColorIndexNone ' no shading for the cell
        Target.Font.Color = RGB(0, 0, 0) ' Black
    End If


     ThisWorkbook.Protect Structure:=True

End Sub
Sub ShowHideModelsQuestions(Target As Range)
    On Error Resume Next
    ThisWorkbook.UnProtect_This_Sheet

    Dim modelsAnswer As Range

    'did the change happen where we are interested?
    Set modelsAnswer = Intersect(Target, RangeModels())

  ' if the change happens outside of the range then ignore it
    If modelsAnswer Is Nothing Then
        Exit Sub
    End If

    If modelsAnswer.Value2 = "No" Or modelsAnswer.Value2 = "" Then
    ' hide the rows
         Dim r As Range

         For Each r In RangeModelQuestions().Rows
             r.EntireRow.Hidden = True
         Next r
         Worksheets("Analytics").Visible = False
    Else
        ' show the rows
        For Each r In RangeModelQuestions().Rows
             r.EntireRow.Hidden = False
         Next r
    End If

    ThisWorkbook.Protect_This_Sheet
End Sub
Sub ShowHideModelsTab(Target As Range)

    On Error Resume Next
    ThisWorkbook.Protect Structure:=False

    Select Case RangeModelsFICO().Value2

    Case "FICO Models"
        Worksheets("Analytics").Visible = xlSheetVisible
        Worksheets("Opti Models").Visible = xlVeryHidden
    Case "FICO Models + FICO Optimization"
        Worksheets("Analytics").Visible = xlSheetVisible
        Worksheets("Opti Models").Visible = xlSheetVisible
    Case "Optimization Only"
        Worksheets("Analytics").Visible = xlVeryHidden
        Worksheets("Opti Models").Visible = xlSheetVisible
    Case Else
        ' hide them
        Worksheets("Analytics").Visible = xlVeryHidden
        Worksheets("Opti Models").Visible = xlVeryHidden
    End Select

    ThisWorkbook.Protect Structure:=True

End Sub
Sub ShowUseCaseTabs(Target As Range)
    ' a hidden mapping tab associates the Salesforce Use Case name to the simpler tab name
    ' need to check if the cell changes from empty to a value or between values or
    ' from having a value to empty

    On Error Resume Next

    Dim useCaseSelection As Range
    Set useCaseSelection = Intersect(Target, RangeUseCase())

    ' if the change happens outside of the range then ignore it
    If useCaseSelection Is Nothing Then
        Exit Sub
    End If

    Dim c As Range
    Dim oldValue As String, newValue As String, tabname As String
    Dim NewTabValue As String, OldTabValue As String

    ' this grabs the current cell value, then using undo gets the old value and then puts the new value back
    ' switching off events prevents the event handlers from firing and prevents the user
    ' seeing what's happening.
    Application.EnableEvents = False

    With Target
        newValue = .Value2
        Application.Undo
        oldValue = .Value2
        .Value2 = newValue
    End With

    Application.EnableEvents = True

    ' do the mapping here so its done
    NewTabValue = UseCaseSelectMapping(newValue)
    OldTabValue = UseCaseSelectMapping(oldValue)

    ' now we have the old and new values of the selection
    If Not NewTabValue = OldTabValue Then

        If OldTabValue <> "" And Not IsNull(OldTabValue) Then
            ThisWorkbook.Protect Structure:=False
            Worksheets(OldTabValue).Visible = xlVeryHidden
            ThisWorkbook.Protect Structure:=True
        End If

        If NewTabValue <> "" And Not IsNull(NewTabValue) Then
            ThisWorkbook.Protect Structure:=False
            Worksheets(NewTabValue).Visible = xlSheetVisible
            ThisWorkbook.Protect Structure:=True
        End If

        If NewTabValue = "" Then
            Target.Interior.ColorIndex = 0
        Else
            Target.Interior.ColorIndex = 37
        End If
    End If
End Sub
Private Sub ShowAllTabs()
    ' this is a helper function I'm using as I test the other things I'm implementing here
    On Error Resume Next

    Dim password As String

    ThisWorkbook.Protect Structure:=False

    password = InputBox("Please enter password", "Password needed!")
    If password = ThisWorkbook.myPassword() Then

        Dim ws As Worksheet
        For Each ws In ActiveWorkbook.Worksheets
            ws.Visible = xlSheetVisible
        Next

    End If

    ThisWorkbook.Worksheets("CLARITY").Activate

    ThisWorkbook.Protect Structure:=True
End Sub
Private Sub ResetButton()

    ThisWorkbook.UnProtect_This_Sheet

    ButtonOn button:=ActiveSheet.Shapes("Reset")

    ' hide all the tabs except Clarity and Triggers
    On Error Resume Next
    ThisWorkbook.Protect Structure:=False

    Dim userResponse As VbMsgBoxResult
    userResponse = MsgBox("SOME data on this tab will clear and all the capability tabs will be hidden but WILL RETAIN ALL DATA previously input.", vbOKCancel, "Please Note!")

    If userResponse = vbCancel Then
        ButtonOff button:=ActiveSheet.Shapes("Reset")
        Exit Sub
    End If

    ThisWorkbook.Worksheets("CLARITY").Visible = xlSheetVisible

    ' if we have the review tabs open, leave them open
    Dim theSheet As Worksheet
    For Each theSheet In ThisWorkbook.Worksheets
       If Not theSheet.Name = "CLARITY" And _
       Not theSheet.Name = "Triggers for Sol Arc" And _
       Not theSheet.Name = "Arc Diagram" And _
       Not theSheet.Name = "Risks" And _
       Not theSheet.Name = "Reviewer Detail" And _
       Not theSheet.Name = "Explanation and Assumptions" And _
       Not theSheet.Name = "Review Summary" And _
       Not theSheet.Name = "Ref_Arch_Diagram" And _
       Not theSheet.Name = "Executive Review" Then
           theSheet.Visible = xlVeryHidden
       End If
    Next

    ThisWorkbook.ClearActiveSheetNames

    ' clear out the drop-downs
    Dim iCell As Range
    For Each iCell In RangeCapability()
        iCell.ClearContents
    Next iCell

    For Each iCell In RangeOptional()
        iCell.ClearContents
    Next iCell

    For Each iCell In RangeNew()
        iCell.ClearContents
    Next iCell

    'simply move the selection to the first row where we can select a capability
    RangeCapability().Cells(1, 1).Select

    'reset the toggles too
    ButtonOff button:=ActiveSheet.Shapes("ProfessionalServices")
    Worksheets("Professional Services").Visible = xlVeryHidden
    ButtonOff button:=ActiveSheet.Shapes("BusinessConsulting")
    Worksheets("Business Consulting").Visible = xlVeryHidden
    ButtonOff button:=ActiveSheet.Shapes("SolutionAccelerators")
    Worksheets("SolutionAccelerators").Visible = xlVeryHidden
    ButtonOff button:=ActiveSheet.Shapes("MultiUseCase")
    ButtonOff button:=ActiveSheet.Shapes("Migration")
    togglePSMigration (0)

    'reset the use case too
    RangeUseCase().Value2 = ""
    RangeUseCase().Interior.ColorIndex = 0

    DeletePicture1
    DeletePicture2

    ButtonOff button:=ActiveSheet.Shapes("Reset")

    ThisWorkbook.Protect_This_Sheet
    ThisWorkbook.Protect Structure:=True

End Sub
Sub ShowSolutionAccelerator()
    On Error Resume Next
    ToggleTabByName buttonName:="SolutionAccelerators", tabname:="SolutionAccelerators"
End Sub
Sub ShowProfessionalServices()
    On Error Resume Next
    ToggleTabByName buttonName:="ProfessionalServices", tabname:="Professional Services"
End Sub
Sub ShowBusinessConsulting()
    On Error Resume Next
    ToggleTabByName buttonName:="BusinessConsulting", tabname:="Business Consulting"
End Sub
Sub ShowReviewTabs()
    On Error Resume Next

    Dim password As String

    password = InputBox("Please enter password", "Password needed!")
    If password <> ThisWorkbook.myPassword() Then
        Exit Sub
    End If

    ThisWorkbook.Protect Structure:=False

   If Not Worksheets("Arc Diagram").Visible = xlSheetVisible Then
        ButtonOn button:=ActiveSheet.Shapes("ReviewRequired")
        Worksheets("Arc Diagram").Visible = xlSheetVisible
        Worksheets("Risks").Visible = xlSheetVisible
        Worksheets("Explanation and Assumptions").Visible = xlSheetVisible
        'Worksheets("Reviewer Detail").Visible = xlSheetVisible
        'Worksheets("Review Summary").Visible = xlSheetVisible
        Worksheets("Clarity").Activate
    Else
        ButtonOff button:=ActiveSheet.Shapes("ReviewRequired")
        Worksheets("Arc Diagram").Visible = xlVeryHidden
        Worksheets("Risks").Visible = xlVeryHidden
        Worksheets("Explanation and Assumptions").Visible = xlVeryHidden
        'Worksheets("Reviewer Detail").Visible = xlVeryHidden
        'Worksheets("Review Summary").Visible = xlVeryHidden
        Worksheets("Clarity").Activate
    End If

    ThisWorkbook.Protect Structure:=True
End Sub
Sub ShowExecutiveTab()
    On Error Resume Next

    Dim password As String

    password = InputBox("Please enter password", "Password needed!")
    If password <> ThisWorkbook.myPassword() Then
        Exit Sub
    End If

    ThisWorkbook.Protect Structure:=False

    If Not Worksheets("Executive Review").Visible = xlSheetVisible Then
        ButtonOn button:=ActiveSheet.Shapes("ExecutiveReview")
        ThisWorkbook.Worksheets("Executive Review").Visible = xlSheetVisible
        ThisWorkbook.Worksheets("Reviewer Detail").Visible = xlSheetVisible
        ThisWorkbook.Worksheets("Instructions").Visible = xlVeryHidden
        ThisWorkbook.Worksheets("Triggers for Sol Arc").Visible = xlVeryHidden
        ThisWorkbook.Worksheets("Clarity").Activate
    Else
        ButtonOff button:=ActiveSheet.Shapes("ExecutiveReview")
        ThisWorkbook.Worksheets("Executive Review").Visible = xlVeryHidden
        ThisWorkbook.Worksheets("Reviewer Detail").Visible = xlVeryHidden
        ThisWorkbook.Worksheets("Clarity").Activate
    End If

    ThisWorkbook.Protect Structure:=True
End Sub
Function ProductSelectMapping(selectedName As String)
    On Error Resume Next

    ' I am mapping Sales Force Product/Capability names to "tab" names in this clarity see ProductMap hidden tab
    ' if we are deleting the contents of the selection then simply clear this cell too

    If selectedName = "" Or IsNull(selectedName) Then
        ProductSelectMapping = ""
        Exit Function
    End If

    ' get the tab name from the mapping tab
    ' the 2 in the statement below is why we get the TAB Name and not the SF Product name from the mapping table

    Dim selectedTab As String

    ' name we're looking for, the place to go looking, the column to return and FALSE means to do an exact match
    selectedTab = Application.VLookup(selectedName, RangeProducts(), 2, False)

    If Not IsError(selectedTab) Then
        ProductSelectMapping = selectedTab
    End If

End Function
Function ProductDependency_1(selectedName As String) As String
    On Error Resume Next
    ' Find the selected capability in column A and this time return the column labeled dependency 1

    If selectedName = "" Or IsNull(selectedName) Then
        ProductDependency_1 = ""
        Exit Function
    End If

    Dim dependency_1 As String

    ' name we're looking for, the place to go looking, the column to return and FALSE means to do an exact match
    dependency_1 = Application.VLookup(selectedName, RangeDependencies(), 3, False)

    If Not IsError(dependency_1) Then
        ProductDependency_1 = dependency_1
    Else
        ProductDependency_1 = ""
    End If

End Function
Function ProductDependency_2(selectedName As String) As String
    On Error Resume Next
    ' Find the selected capability in column A and this time return the column labeled dependency 1

    If selectedName = "" Or IsNull(selectedName) Then
        ProductDependency_2 = ""
        Exit Function
    End If

    Dim dependency_2 As String

    ' name we're looking for, the place to go looking, the column to return and FALSE means to do an exact match
    dependency_2 = Application.VLookup(selectedName, RangeDependencies(), 4, False)

    If Not IsError(dependency_2) Then
        ProductDependency_2 = dependency_2
    Else
        ProductDependency_2 = ""
    End If

End Function
Sub InsertDependency(dependentName As String)
    On Error Resume Next
    ' go look at the list of capabilities in the list and find the first open row
    ' need to also check to see if its already in the list...

    ' maybe its already there because they added it before
    If isListedCapability(dependentName) Then
        Exit Sub
    End If

     Dim numRows As Integer

     numRows = RangeCapability().Rows.Count
     RangeCapability().Select

     Application.ScreenUpdating = False

     Do Until IsEmpty(ActiveCell)
          ActiveCell.Offset(1, 0).Select
     Loop

     Application.ScreenUpdating = True

     ' now add in the dependency --- how do we get it to show theTAB?
     ActiveCell.Value = dependentName
     Worksheets(dependentName).Visible = xlSheetVisible

     Application.EnableEvents = True

End Sub
Sub DeleteDependency(dependentName As String)
    On Error Resume Next
    ' search for the one that we need to delete and delete it

    If dependentName = "" Or IsNull(dependentName) Then
        Exit Sub
    End If

    Dim tabname As String
    Dim numRows, loopCounter As Integer
    numRows = RangeCapability().Rows.Count
    loopCounter = 0

    RangeCapability().Select

    Application.ScreenUpdating = False

    Do Until (ActiveCell.Value = dependentName) Or (loopCounter = numRows)
        ActiveCell.Offset(1, 0).Select
        loopCounter = loopCounter + 1
    Loop

    Application.ScreenUpdating = True

    ' the "active cell" now MAYBE has the value that we want to delete
    If ActiveCell.Value = dependentName Then
        ActiveCell.ClearContents
        ' get the tab to hide
        tabname = ProductSelectMapping(dependentName)
        ThisWorkbook.Protect Structure:=False
        Worksheets(tabname).Visible = xlVeryHidden
        ThisWorkbook.Protect Structure:=True
    End If
End Sub
Function isListedCapability(selectedName As String) As Boolean
    On Error Resume Next
    isListedCapability = False

    If selectedName = "" Or IsNull(selectedName) Then
        Exit Function
    End If

    Dim numRows, loopCounter As Integer
    numRows = RangeCapability().Rows.Count
    loopCounter = 0

    RangeCapability().Select

    Application.ScreenUpdating = False

    Do Until (ActiveCell.Value = selectedName) Or (loopCounter = numRows)
        ActiveCell.Offset(1, 0).Select
        loopCounter = loopCounter + 1
    Loop

    Application.ScreenUpdating = True

    If ActiveCell.Value = selectedName Then
        isListedCapability = True
    End If
End Function
Function UseCaseSelectMapping(selectedName As String)
   On Error Resume Next
   ' I am mapping Salesforce Use Case names to "tab" names in this clarity see ProductMap hidden tabs
   ' if we are deleting the contents of the selection then simply clear this cell too

    If selectedName = "" Or IsNull(selectedName) Then
        UseCaseSelectMapping = ""
        Exit Function
    End If

    ' get the tab name from the mapping tab
    ' the 2 in the statement below is why we get the TAB Name and not the SF Product name from the mapping table

    Dim selectedTab As String

    selectedTab = Application.VLookup(selectedName, RangeUseCases(), 2, False)

    If Not IsError(selectedTab) Then
        UseCaseSelectMapping = selectedTab
    End If

End Function
Sub showCaption()
    On Error Resume Next
    ' show the full name in the title bar
    ThisWorkbook.Protect Structure:=False
    ActiveWindow.Caption = ActiveWorkbook.FullName
    ThisWorkbook.Protect Structure:=True

End Sub
Sub showMultiUseCase()
    On Error Resume Next

    Dim btn As Shape
    Set btn = ActiveSheet.Shapes("MultiUseCase")

    If Not ButtonStatus(btn) Then
        ToggleTabByName buttonName:="MultiUseCase", tabname:=""

        Dim Msg, Title, Style

        Msg = "Use a SEPARATE CLARITY Form FOR EACH Use Case"
        Title = "Multiple Use Case Message"
        Style = vbOKOnly Or vbExclamation Or vbSystemModal Or vbMsgBoxSetForeground

        MsgBox Msg, Style, Title
    Else
        ToggleTabByName buttonName:="MultiUseCase", tabname:=""
    End If
End Sub

' so the worksheets are locked meaning you can't just import images and
' put them where you want... I had to invent a way to allow this to happen
' while keeping the protections.
' The the Clarity tab itself, I allow for 2 images ...
Sub ImportPictureFromFile1()
    On Error Resume Next
    addImageByName "A88", "Picture 1"

End Sub
Sub ImportPictureFromFile2()
    On Error Resume Next
    addImageByName "D88", "Picture 2"

End Sub
Sub DeletePicture1()
    On Error Resume Next
    ThisWorkbook.UnProtect_This_Sheet

    deleteImageByName "Picture 1"
    ThisWorkbook.Protect_This_Sheet

End Sub
Sub DeletePicture2()
    On Error Resume Next

    ThisWorkbook.UnProtect_This_Sheet
    deleteImageByName "Picture 2"
    ThisWorkbook.Protect_This_Sheet

End Sub
Private Sub addImageByName(ByVal ImageLocation As String, ByVal ImageName As String)
    On Error Resume Next

    ThisWorkbook.UnProtect_This_Sheet
    Dim fNameAndPath As Variant

    fNameAndPath = Application.GetOpenFilename(Title:="Select picture to be imported")

    If fNameAndPath = False Then
        Exit Sub
    End If

    Dim s As Shape
    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets("Clarity")
    Set s = ws.Shapes.AddPicture2(fNameAndPath, msoFalse, msoTrue, ThisWorkbook.Sheets("Clarity").Range(ImageLocation).Left, ThisWorkbook.Sheets("Clarity").Range(ImageLocation).Top, -1, -1, msoPictureCompressDocDefault)

    s.Name = ImageName
    s.LockAspectRatio = msoTrue
    s.Width = GetWidthABC()

    ThisWorkbook.Protect_This_Sheet
End Sub
Private Sub deleteImageByName(ByVal arg1 As String)
    On Error Resume Next

    Dim pic As Shape

    For Each pic In ThisWorkbook.Worksheets("Clarity").Shapes
        If InStr(1, pic.Name, arg1, vbTextCompare) <> 0 Then
            pic.Delete
        End If
    Next pic

End Sub
Function GetWidthABC()
    On Error Resume Next

    Dim i As Integer

    For i = 1 To 3 ' Columns A to E
        GetWidthABC = GetWidthABC + Columns(i).Width
    Next i

End Function
Sub SuperReset()
    On Error Resume Next

    Dim userResponse As VbMsgBoxResult
    userResponse = MsgBox("Secret close all tabs button", vbOKCancel, "Super Reset!")

    If userResponse = vbCancel Then
        Exit Sub
    End If

    Dim password As String
    password = InputBox("Please enter password", "Password needed!")

    If password <> ThisWorkbook.myPassword() Then
        Exit Sub
    End If

    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    Range("A1").Select
    ThisWorkbook.UnProtect_This_Sheet
    ' hide all the tabs except Instructions and Triggers

    ThisWorkbook.Protect Structure:=False

    ThisWorkbook.Worksheets("CLARITY").Visible = xlVeryHidden
    ThisWorkbook.Worksheets("Instructions").Visible = xlSheetVisible
    ThisWorkbook.Worksheets("Triggers for Sol Arc").Visible = xlSheetVisible

    Dim theSheet As Worksheet
    For Each theSheet In ThisWorkbook.Worksheets
       If Not theSheet.Name = "Instructions" And _
          Not theSheet.Name = "Triggers for Sol Arc" Then
           theSheet.Visible = xlVeryHidden
       End If
    Next

    ThisWorkbook.SetFirstWorkbookOpen
    ThisWorkbook.ClearActiveSheetNames

    ThisWorkbook.Worksheets("CLARITY").Activate
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1

    ' clear out the drop-downs
    Dim iCell As Range
    For Each iCell In RangeCapability()
        iCell.ClearContents
    Next iCell

    For Each iCell In RangeOptional()
        iCell.ClearContents
    Next iCell

    For Each iCell In RangeNew()
        iCell.ClearContents
    Next iCell

    'reset the toggles too
    ButtonOff button:=ActiveSheet.Shapes("ProfessionalServices")
    Worksheets("Professional Services").Visible = xlVeryHidden
    ButtonOff button:=ActiveSheet.Shapes("BusinessConsulting")
    Worksheets("Business Consulting").Visible = xlVeryHidden
    ButtonOff button:=ActiveSheet.Shapes("SolutionAccelerators")
    Worksheets("SolutionAccelerators").Visible = xlVeryHidden
    ButtonOff button:=ActiveSheet.Shapes("MultiUseCase")
    ButtonOff button:=ActiveSheet.Shapes("Migration")
    togglePSMigration (0)

    Dim GrayColor As Long
    Dim BlackColor As Long
    'Dim btn As Shape
    GrayColor = RGB(192, 192, 192)
    BlackColor = RGB(0, 0, 0)

    ButtonOff button:=ActiveSheet.Shapes("ShowReference")
    Worksheets("Ref_Arch_Diagram").Visible = xlVeryHidden
    Sheet43.DeleteDiagram

    ButtonOff button:=ActiveSheet.Shapes("ReviewRequired")
    ButtonOff button:=ActiveSheet.Shapes("ReviewRequired")
    Worksheets("Arc Diagram").Visible = xlVeryHidden
    Worksheets("Risks").Visible = xlVeryHidden
    Worksheets("Explanation and Assumptions").Visible = xlVeryHidden
    Worksheets("Reviewer Detail").Visible = xlVeryHidden
    Worksheets("Review Summary").Visible = xlVeryHidden

    ButtonOff button:=ActiveSheet.Shapes("ExecutiveReview")
    Worksheets("Executive Review").Visible = xlVeryHidden

    'reset the Use Case too
    RangeUseCase().Value2 = ""
    RangeUseCase().Interior.ColorIndex = 0

    'reset the Models Question too
    RangeModels().Value2 = "No"

    'simply move the selection to the first row where we can select a capability
    RangeCapability().Cells(1, 1).Select

    DeletePicture1
    DeletePicture2

    ThisWorkbook.Protect_This_Sheet
    ThisWorkbook.Protect Structure:=True
End Sub
Sub Toggle_Workbook_Protection()
    ' this is a helper function I'm using as I test the other things I'm implementing here
    On Error Resume Next

    Dim password As String

    password = InputBox("Please enter password", "Password needed!")
    If password = ThisWorkbook.myPassword() Then
        If ActiveWorkbook.ProtectStructure = True Then
            ThisWorkbook.Protect Structure:=False
            MsgBox "Workbook Structure Unprotected"
        Else
            ThisWorkbook.Protect Structure:=True
            MsgBox "Workbook Structure Protected"
        End If
    Else
        MsgBox "Incorrect Password, status not changed."
    End If
End Sub
Sub AddReferenceArchitecture()
    On Error Resume Next

    ' if already visible just hide it... otherwise ask for PW to enable it...
    If Worksheets("Ref_Arch_Diagram").Visible = xlSheetVisible Then
        ThisWorkbook.Protect Structure:=False
        Worksheets("Ref_Arch_Diagram").Visible = xlVeryHidden
        ButtonOff button:=ActiveSheet.Shapes("ShowReference")
    Else
        Dim password As String
        password = InputBox("Please enter password", "Password needed!")
        If password <> ThisWorkbook.myPassword() Then
           Exit Sub
        End If

        ThisWorkbook.Protect Structure:=False
        Worksheets("Ref_Arch_Diagram").Visible = xlSheetVisible
        ButtonOn button:=ActiveSheet.Shapes("ShowReference")
        Worksheets("Ref_Arch_Diagram").Activate
    End If

    ThisWorkbook.Protect Structure:=True
End Sub
Sub ShowPSMigration()
    On Error Resume Next

    Dim btn As Shape
    Set btn = ActiveSheet.Shapes("Migration")

    If Not ButtonStatus(btn) Then
        ToggleTabByName buttonName:="Migration", tabname:="LegacyMigration"
        Dim Msg, Title, Style

        Msg = "Please document the name of the PS Lead for the project. Add migration details to the Migration Tab"
        Title = "This deal includes a PS Migration"
        Style = vbOKOnly Or vbExclamation Or vbSystemModal Or vbMsgBoxSetForeground

        MsgBox Msg, Style, Title
        togglePSMigration (1)
        Worksheets("Clarity").Activate
        ' i hate hard coding this but what else can I do because using the helper function at the top of this
        ' files doesn't work the same way.
        Worksheets("CLARITY").Range("C31").Select

    Else
        ToggleTabByName buttonName:="Migration", tabname:="LegacyMigration"
        togglePSMigration (0)
    End If
End Sub
Sub ToggleLookupTabs()
    On Error Resume Next

    Dim password As String

    password = InputBox("Please enter password", "Password needed!")
    If password <> ThisWorkbook.myPassword() Then
        Exit Sub
    End If

    ThisWorkbook.Protect Structure:=False

    If Worksheets("Lookups").Visible = xlSheetVisible Then
        Worksheets("Lookups").Visible = xlVeryHidden
    Else
        Worksheets("Lookups").Visible = xlSheetVisible
    End If

    If Worksheets("Analytics Lookups").Visible = xlSheetVisible Then
        Worksheets("Analytics Lookups").Visible = xlVeryHidden
    Else
        Worksheets("Analytics Lookups").Visible = xlSheetVisible
    End If

    If Worksheets("ProductsMap").Visible = xlSheetVisible Then
        Worksheets("ProductsMap").Visible = xlVeryHidden
    Else
        Worksheets("ProductsMap").Visible = xlSheetVisible
    End If

    If Worksheets("UseCaseMap").Visible = xlSheetVisible Then
        Worksheets("UseCaseMap").Visible = xlVeryHidden
    Else
        Worksheets("UseCaseMap").Visible = xlSheetVisible
    End If

    ThisWorkbook.Protect Structure:=True
End Sub
Sub togglePSMigration(theStatus As Integer)
    On Error Resume Next

    ThisWorkbook.UnProtect_This_Sheet
    Dim r

    If theStatus = 1 Then
        For Each r In RangePSMigration().Rows
             r.EntireRow.Hidden = False
         Next r
    Else
        For Each r In RangePSMigration().Rows
             r.EntireRow.Hidden = True
         Next r
    End If

    ThisWorkbook.Protect_This_Sheet
End Sub
Sub ToggleTabByName(buttonName As String, tabname As String)
    ' refactoring some code to do this in 1 place rather than in several places.
    On Error Resume Next

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

    ' toggle viability
    If Worksheets(tabname).Visible = xlSheetVisible Then
        ButtonOff button:=btn
        Worksheets(tabname).Visible = xlVeryHidden
    Else
        Worksheets(tabname).Visible = xlSheetVisible
        ButtonOn button:=btn
    End If

    ThisWorkbook.Protect Structure:=True
End Sub
Private Sub ButtonOff(button As Shape)
    On Error Resume Next

    Dim GrayColor As Long
    Dim BlackColor As Long
    GrayColor = RGB(192, 192, 192)
    BlackColor = RGB(0, 0, 0)

    button.Fill.ForeColor.RGB = GrayColor
    button.TextFrame.Characters.Font.Color = BlackColor
End Sub
Private Sub ButtonOn(button As Shape)
    On Error Resume Next

    Dim GreenColor As Long
    Dim WhiteColor As Long
    GreenColor = RGB(51, 153, 102)
    WhiteColor = RGB(255, 255, 255)

    button.Fill.ForeColor.RGB = GreenColor
    button.TextFrame.Characters.Font.Color = WhiteColor
End Sub
Private Function ButtonStatus(button As Shape) As Boolean
    On Error Resume Next

    Dim GrayColor As Long
    GrayColor = RGB(192, 192, 192)

    If button.Fill.ForeColor.RGB = GrayColor Then
        ButtonStatus = False
    Else
        ButtonStatus = True
    End If
End Function