'
' Macros for the sheet named "thisworkbook", this is a special GLOBAL
' set of VB for the whole workbook
' Copyright FICO (Fair Isaac Inc) 2023, 2024, 2025
'

Option Explicit
Option Compare Text

Const CAPABILITY_LIST As String = "A4:A18"

Const EPC_QUESTION As String = "C23"

Const DBS_DEDICATED_VALUE As String = "C4"
Const DBS_DEDICATED_COMMENT As String = "D4"
Const DBS_SIZE_VALUE As String = "C10"
Const DBS_SIZE_COMMENT As String = "D10"

Const FAWB_DEDICATED_VALUE As String = "C6"
Const FAWB_DEDICATED_COMMENT As String = "D6"
Const FAWB_SIZE_VALUE As String = "C12"
Const FAWB_SIZE_COMMENT As String = "D12"

Private Sub Workbook_Open()
    ' this runs when the workbook opens
    On Error Resume Next

    ' this is all sheets in the active workbook because Excel doesn't offer this at the sheet level
    ActiveWindow.Zoom = 100

    ' disallow drag and drop
    Application.CellDragAndDrop = False

    ThisWorkbook.Protect Structure:=False
    If CheckFirstOpen = True Then

        ' we've been edited so we can skip the Instructions.
        ThisWorkbook.Worksheets("CLARITY").Visible = xlVeryHidden
        ThisWorkbook.Worksheets("Instructions").Visible = xlSheetVisible
        ThisWorkbook.Worksheets("Triggers for Sol Arc").Visible = xlSheetVisible
        ThisWorkbook.Worksheets("Instructions").Activate
    Else

        ThisWorkbook.Worksheets("CLARITY").Visible = xlSheetVisible
        'ThisWorkbook.Worksheets("Instructions").Visible = xlSheetVisible
        ThisWorkbook.Worksheets("CLARITY").Activate
        ThisWorkbook.Worksheets("CLARITY").Range("A1").Activate
    End If

    ' protect the sheets from changes except where we want people to make changes
    Dim ws As Worksheet

    ThisWorkbook.ProtectAllSheets

    ' Now make the instructions visible and activated so that's what you see when opening

    ThisWorkbook.Protect Structure:=True

End Sub
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' this runs just before the workbook closes
    On Error Resume Next

    ThisWorkbook.Protect Structure:=False

    ThisWorkbook.Worksheets("Instructions").Activate
    ThisWorkbook.Worksheets("Triggers for Sol Arc").Visible = xlVeryHidden
    ' If I'm here, I've been open so the new logic says if I've been open go to the Clarity w/o
    ' going to the Instructions ...
    ClearFirstWorkbookOpen

    SaveActiveSheetNames

    UpdateLastAuthor
    UpdateLastModifiedDate

    ThisWorkbook.Protect Structure:=True

    ThisWorkbook.Worksheets("Instructions").Activate
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1

End Sub
Sub InstructionButton()
    ' I want to ensure the user sees the instructions when they open the file.
    ' and then we show them the Clarity sheet

    On Error Resume Next

    ThisWorkbook.Protect Structure:=False

    'ThisWorkbook.Worksheets("Triggers for Sol Arc").Visible = xlSheetVisible

    ' now that we've been past here we can skip the instructions in the future
    ClearFirstWorkbookOpen
    ThisWorkbook.Worksheets("CLARITY").Visible = xlSheetVisible

    ' this is how I saved off the names of sheets open when the WB was closed and then restore them on re-open
    ReOpenActiveSheets
    ThisWorkbook.ClearActiveSheetNames

    Worksheets("Clarity").Activate
    ThisWorkbook.Worksheets("Instructions").Visible = xlVeryHidden

    ThisWorkbook.Protect Structure:=True

End Sub
Private Sub SaveActiveSheetNames()
    ' Loop through the tabs and make a note of those that are open
    On Error Resume Next
    ThisWorkbook.Worksheets("Instructions").Activate

    ThisWorkbook.ClearActiveSheetNames

    ThisWorkbook.Protect Structure:=False

    ThisWorkbook.UnProtect_This_Sheet
    'List all VISIBLE sheet names in column AA of sheet name = Instructions when closing the file so we can restore later

    ThisWorkbook.Worksheets("Instructions").Visible = xlSheetVisible
    ThisWorkbook.Worksheets("Instructions").Activate
    ThisWorkbook.Worksheets("Instructions").Range("AZ1").Select

    Dim sh As Worksheet

    For Each sh In Worksheets
        If sh.Visible = xlSheetVisible Then
            Selection = sh.name
            Selection.Font.name = "Calibri"
            Selection.Font.Size = 12
            Selection.Offset(1, 0).Select  'Move down a row.
        End If
    Next

    ActiveCell.EntireColumn.AutoFit
    ThisWorkbook.Worksheets("Instructions").Range("A1").Select

    ThisWorkbook.Protect_This_Sheet
    ThisWorkbook.Protect Structure:=True
End Sub
Sub ClearActiveSheetNames()
   ' clear out the list that was created by SaveActiveSheetNames()
    On Error Resume Next

    ThisWorkbook.Worksheets("Instructions").Activate
    Dim theInstructionsSheet As Worksheet
    Set theInstructionsSheet = ThisWorkbook.Worksheets("Instructions")

    ThisWorkbook.UnProtect_This_Sheet

    ' 52 is the number of the column also labeled AZ1
    ' clear all the cells that have a value in this range
    theInstructionsSheet.Columns(52).ClearContents

    ThisWorkbook.Protect_This_Sheet

End Sub
Private Sub ReOpenActiveSheets()
    ' Use the list that was created by SaveActiveSheetNames() to ensure we restore visibility to those tabs
    On Error Resume Next

    ThisWorkbook.Worksheets("Instructions").Activate
    ThisWorkbook.Protect Structure:=False

    Dim theInstructionsSheet As Worksheet
    Set theInstructionsSheet = ThisWorkbook.Worksheets("Instructions")

    Dim LastRow As Long
    Dim StartCell As Range
    Dim ws As Worksheet
    Dim cell As Range

    Set StartCell = Range("AZ1")
    LastRow = theInstructionsSheet.Cells(theInstructionsSheet.Rows.Count, StartCell.Column).End(xlUp).Row

    ThisWorkbook.Protect Structure:=False

    For Each cell In Range(StartCell, StartCell.Offset(LastRow - 1))
        ThisWorkbook.Worksheets(cell.Value).Visible = xlSheetVisible
    Next cell

    ThisWorkbook.Protect Structure:=True
End Sub
Sub CheckRequiredCell(cell As Range)
    ' check that cells with the style "Input Required" have a value or mark
    ' them with a red border
    On Error Resume Next

    If cell.Style = "Input Required" Then
        With cell.Borders
            .LineStyle = xlContinuous

            If IsEmpty(cell) Then
                .Color = vbRed
                .Weight = xlWide
            Else
                 .Color = vbBlack
                 .Weight = xlThin
            End If
        End With
    End If

End Sub
Function myPassword() As String
    ' this gets reused everywhere we want to lock and unlock so we can maintain in 1 place.
    myPassword = "12345"
End Function
Private Sub UpdateLastAuthor()
    On Error Resume Next

    ThisWorkbook.Protect Structure:=False
    ThisWorkbook.Worksheets("Clarity").Unprotect ThisWorkbook.myPassword()

    ThisWorkbook.Worksheets("Clarity").Range("C2").Value = ThisWorkbook.BuiltinDocumentProperties("Last Author")

    ThisWorkbook.Worksheets("Clarity").Protect ThisWorkbook.myPassword()
    ThisWorkbook.Protect Structure:=True
End Sub
Private Sub UpdateLastModifiedDate()
    On Error Resume Next

    ThisWorkbook.Protect Structure:=False
    ThisWorkbook.Worksheets("Clarity").Unprotect ThisWorkbook.myPassword()

    ThisWorkbook.Worksheets("Clarity").Range("B2").Value = Now()

    ThisWorkbook.Worksheets("Clarity").Protect ThisWorkbook.myPassword()
    ThisWorkbook.Protect Structure:=False

End Sub
Sub ProtectAllSheets()
    ' as the name implies loop all the tabs and set protection
    On Error Resume Next

    ThisWorkbook.Protect Structure:=False
    Dim ws As Worksheet

    For Each ws In ActiveWorkbook.Worksheets
        ws.Protect password:=ThisWorkbook.myPassword
    Next ws

    ' this allows CSA to import an image here... while locking everything else
    ThisWorkbook.Worksheets("Review Summary").Protect password:=ThisWorkbook.myPassword, DrawingObjects:=False

    ThisWorkbook.Protect Structure:=True

End Sub
Private Sub UnprotectAllSheets()
    ' as the name implies loop all the tabs and remove protection

    On Error Resume Next
    ThisWorkbook.Protect Structure:=False

    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
        ws.Unprotect password:=ThisWorkbook.myPassword
    Next ws

    ThisWorkbook.Protect Structure:=True

End Sub
Function getOS() As String
    Dim OSname As String

    OSname = Application.OperatingSystem

    If InStr(1, OSname, "Windows", vbTextCompare) Then
        getOS = "Windows"
    Else
        getOS = "Other"
    End If
End Function
'
' This used to be done at the sheet level but when I want to make mass changes I have to
' change things in 50 places. If all sheets reference this Sub then I can temporarily tun this off
' in one place.
Sub Protect_This_Sheet()
    'ActiveSheet.Protect ThisWorkbook.myPassword()
End Sub
Sub UnProtect_This_Sheet()
    ActiveSheet.Unprotect ThisWorkbook.myPassword()
End Sub
Sub SetFirstWorkbookOpen()
    On Error Resume Next

    ThisWorkbook.UnProtect_This_Sheet
    Worksheets("Instructions").Activate

    Range("AY1").Select
    ActiveCell = "FirstOpen"

    ThisWorkbook.Protect_This_Sheet
    'MsgBox "First Open should be TRUE here...."
End Sub
Sub ClearFirstWorkbookOpen()
    On Error Resume Next

    ThisWorkbook.UnProtect_This_Sheet
    Worksheets("Instructions").Activate

    Range("AY1").Select
    Selection.Value = "FALSE"

    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1

    ThisWorkbook.Protect_This_Sheet
End Sub
Private Function CheckFirstOpen() As Boolean
    On Error Resume Next

    Worksheets("Instructions").Activate

    If Range("AY1").Value2 = "FirstOpen" Then
        CheckFirstOpen = True
    Else
        CheckFirstOpen = False
    End If

End Function
Sub AllSheet_Protection_Off()
    ' this is a helper function for when I'm making changes
    On Error Resume Next

    Dim password As String

    password = InputBox("Please enter password", "Password needed!")

    If password = ThisWorkbook.myPassword() Then
       UnprotectAllSheets
    Else
        MsgBox "Incorrect Password, status not changed."
    End If
End Sub
Sub AllSheet_Protection_On()
    ' this is a helper function for when I'm making changes
    On Error Resume Next

    Dim password As String

    password = InputBox("Please enter password", "Password needed!")

    If password = ThisWorkbook.myPassword() Then
       ThisWorkbook.ProtectAllSheets
    Else
        MsgBox "Incorrect Password, status not changed."
    End If
End Sub
Function checkCapabilityInSolution(ByVal name As String) As Boolean
    On Error Resume Next

    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("Clarity")
    Dim rng As Range: Set rng = ws.Range(CAPABILITY_LIST)
    checkCapabilityInSolution = False

    Dim cell As Range
    Dim cellValue As Variant

    For Each cell In rng.Cells
        cellValue = cell.Value
        If InStr(1, cellValue, name) > 0 Then
            ' Found the substring, perform action
            'Debug.Print "Found " & name & " : in cell " & cell.Address
            checkCapabilityInSolution = True
            Exit For
        End If
    Next cell
End Function
Sub CrossCheckDedicated()
    ' cross check values here with the DB Service & FAWB tabs
    On Error Resume Next

    Dim FAWB_DEDICATED As Variant
    Dim DBS_DEDICATED As Variant

    FAWB_DEDICATED = Worksheets("FAWB").Range(FAWB_DEDICATED_VALUE).Value2
    DBS_DEDICATED = Worksheets("DB Service").Range(DBS_DEDICATED_VALUE).Value2

    ' allow blank and "No" as being the same thing....
    If (FAWB_DEDICATED = "" And DBS_DEDICATED = "No") Or (FAWB_DEDICATED = "No" And DBS_DEDICATED = "")  Or (IsEmpty(FAWB_DEDICATED) And IsEmpty(DBS_DEDICATED)) Then
        Exit Sub
    End If

    ' otherwise we need a match - including "uncertain"
    If Not FAWB_DEDICATED = DBS_DEDICATED Then
        MsgBox "Mismatch of dedicated DB property between DB Service and FAWB", , "Error"
    End If
End Sub
Sub CrossCheckSize()
    ' cross check values here with the DB Service & FAWB tabs
    On Error Resume Next

    Dim FAWB_SIZE As Variant
    Dim DBS_SIZE As Variant

    FAWB_SIZE = Worksheets("FAWB").Range(FAWB_SIZE_VALUE).Value2
    DBS_SIZE = Worksheets("DB Service").Range(DBS_SIZE_VALUE).Value2

    ' if both tabs have no size specified that's ok
    If IsEmpty(FAWB_SIZE) and IsEmpty(DBS_SIZE) Then
        Exit Sub
    End IF

    IF  FAWB_SIZE = DBS_SIZE Then
        Exit Sub
    Else
        MsgBox "FAWB DB Size = " & FAWB_SIZE & "\n DB Service Size = " & DBS_SIZE, , "Error"
    End If
End Sub
Sub CheckEPC()
    On Error Resume Next

    ' check status of the EPC flag
    Dim epcValue As Variant
    epcValue = Worksheets("Clarity").Range(EPC_QUESTION).Value2

    If epcValue = "Yes" Then
        Worksheets("DB Service").Range(DBS_DEDICATED_VALUE).Value2 = "Yes"
        Worksheets("DB Service").Range(DBS_DEDICATED_COMMENT).Value2 = "EPC Clients get a dedicated DB"
        Worksheets("FAWB").Range(FAWB_DEDICATED_VALUE).Value2 = "Yes"
        Worksheets("FAWB").Range(FAWB_DEDICATED_COMMENT).Value2 = "EPC Clients get a dedicated DB"

        ' force the size up to large if its not already large or x-large
        If Not Worksheets("DB Service").Range(DBS_SIZE_VALUE).Value2 = "Large" And Not Worksheets("DB Service").Range(DBS_SIZE_VALUE).Value2 = "X-Large" Then
            Worksheets("DB Service").Range(DBS_SIZE_VALUE).Value2 = "Large"
            Worksheets("DB Service").Range(DBS_SIZE_COMMENT).Value2 = "EPC Clients default to LARGE"
        End If

        ' force the size up to large if its not already large or x-large
        If Not Worksheets("FAWB").Range(FAWB_SIZE_VALUE).Value2 = "Large" And Not Worksheets("FAWB").Range(FAWB_SIZE_VALUE).Value2 = "X-Large" Then
            Worksheets("FAWB").Range(FAWB_SIZE_VALUE).Value2 = "Large"
            Worksheets("FAWB").Range(FAWB_SIZE_COMMENT).Value2 = "EPC Clients default to LARGE"
        End If
    Else
        'If IsEmpty(Worksheets("DB Service").Range(DBS_DB_SIZE).Value2) Then
        'End If
    End If
End Sub