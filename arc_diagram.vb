'
' Macros for the sheet named "arc diagram"
' Copyright FICO (Fair Isaac Inc) 2023, 2024, 2025
'
Option Explicit

Const ARC_DEDICATED_DB_VALUE As String = "G40"
Const ARC_DEDICATED_DB_COMMENT As String = "H40"
Const ARC_DB_SIZE_VALUE As String = "G41"
Const ARC_DB_SIZE_COMMENT As String = "H41"

Const ARC_DEDICATED_SERVICES_VALUE As String = "G42"
Const ARC_DEDICATED_SERVICES_COMMENT As String = "H42"

Const FAWB_DEDICATED_DB_VALUE As String = "C6"
Const FAWB_DEDICATED_DB_COMMENT As String = "D6"
Const FAWB_SIZE_VALUE As String = "C12"
Const FAWB_SIZE_COMMENT As String = "D12"

Const DBS_DEDICATED_DB_VALUE As String = "C4"
Const DBS_DEDICATED_DB_COMMENT As String = "D4"
Const DBS_SIZE_VALUE As String = "C10"
Const DBS_SIZE_COMMENT As String = "D10"

Const EPC_QUESTION As String = "C23"

Private Sub Worksheet_Activate()
    On Error Resume Next

    Dim Cell As Range
    For Each Cell In Application.ActiveSheet.UsedRange
        ThisWorkbook.CheckRequiredCell Cell
    Next

    ThisWorkbook.UnProtect_This_Sheet

    Range("B13").EntireRow.AutoFit
    Range("A1").Select

    CheckDB
    CheckServices

    '
    ' I'm using a Sub at the ThisWorkbook level so I can toggle
    ' this on/off easily for mass edits.
    '
    ThisWorkbook.Protect_This_Sheet
End Sub
Sub ImportArchPicture()
' the imported picture is always named ArchPicture

    ThisWorkbook.UnProtect_This_Sheet
    Dim fNameAndPath As Variant

    fNameAndPath = Application.GetOpenFilename("Image Files (*.jpg;*.jpeg;*.png;*.gif), *.jpg;*.jpeg;*.png;*.gif", _
                                              Title:="Select picture to be imported")
    If fNameAndPath = False Then
        Exit Sub
    End If

    Dim s As Shape
    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets("Arc Diagram")
    Set s = ws.Shapes.AddPicture(fNameAndPath, False, True, ThisWorkbook.Sheets("Arc Diagram").Range("A21").Left, ThisWorkbook.Sheets("Arc Diagram").Range("A21").Top, -1, -1)

    s.name = "ArchPicture"
    s.LockAspectRatio = msoTrue

    ' why do we need to do this? I don't know
    If ThisWorkbook.getOS = "Windows" Then
        s.Width = GetColumnWidths()
    Else
        s.Width = GetColumnWidths() '+ 400
    End If

    s.Line.Weight = 1
    s.Line.ForeColor.RGB = RGB(0, 0, 0) ' black border around the imported image
    s.Fill.Visible = msoFalse ' don't fill in, just a border

    ThisWorkbook.Protect_This_Sheet
End Sub
Sub DeleteArchPicture()

    ThisWorkbook.UnProtect_This_Sheet
        deleteImageByName "ArchPicture"
    ThisWorkbook.Protect_This_Sheet

End Sub
Private Sub deleteImageByName(ByVal arg1 As String)
' pass in the name of the image to delete

    Dim pic As Shape

    For Each pic In ThisWorkbook.Worksheets("Arc Diagram").Shapes
        If InStr(1, pic.name, arg1, vbTextCompare) <> 0 Then
            pic.Delete
        End If
    Next pic

End Sub
Function GetColumnWidths()

    Dim i As Integer

    For i = 1 To 5 ' Columns A to E
        GetColumnWidths = GetColumnWidths + Columns(i).Width
    Next i

End Function
Sub ReviewerTabsOn()
    On Error Resume Next

    Dim password As String

    password = InputBox("Please enter password", "Password needed!")
    If password <> ThisWorkbook.myPassword() Then
        Exit Sub
    End If

    ThisWorkbook.Protect Structure:=False

    Worksheets("Reviewer Detail").Visible = xlSheetVisible
    Worksheets("Review Summary").Visible = xlSheetVisible

    Worksheets("Arc Diagram").Activate
    Range("A1").Select
    ActiveWindow.ScrollRow = ActiveCell.Row
    ActiveWindow.ScrollColumn = ActiveCell.Column
    ThisWorkbook.Protect Structure:=True
End Sub
Sub ReviewerTabsOff()
    On Error Resume Next

    Dim password As String

    password = InputBox("Please enter password", "Password needed!")
    If password <> ThisWorkbook.myPassword() Then
        Exit Sub
    End If

    ThisWorkbook.Protect Structure:=False

    Worksheets("Reviewer Detail").Visible = xlVeryHidden
    Worksheets("Review Summary").Visible = xlVeryHidden

    Worksheets("Arc Diagram").Activate
    Range("A1").Select
    ActiveWindow.ScrollRow = ActiveCell.Row
    ActiveWindow.ScrollColumn = ActiveCell.Column
    ThisWorkbook.Protect Structure:=True
End Sub
Sub sheet_autofit()
    ThisWorkbook.UnProtect_This_Sheet

    Range("B14").EntireRow.AutoFit
    Range("G31").EntireRow.AutoFit
    Range("G32").EntireRow.AutoFit
    Range("G36").EntireRow.AutoFit
    Range("G37").EntireRow.AutoFit

    ThisWorkbook.Protect_This_Sheet
End Sub
Sub Worksheet_Change(ByVal Target As Range)
    
    ThisWorkbook.UnProtect_This_Sheet
    Dim Cell As Range
    For Each Cell In Application.ActiveSheet.UsedRange
        ThisWorkbook.CheckRequiredCell Cell
    Next
    sheet_autofit
    ThisWorkbook.Protect_This_Sheet
End Sub
Sub CheckDB()
    On Error Resume Next

    Dim IsEPC As Variant
    IsEPC = Worksheets("Clarity").Range(EPC_QUESTION).Value2

    Dim FAWBInSolution As Boolean
    Dim DBSInSolution As Boolean
    FAWBInSolution = ThisWorkbook.checkCapabilityInSolution("FICO Applications Workbench - Cloud Edition")
    DBSInSolution = ThisWorkbook.checkCapabilityInSolution("FICO Platform - Database Service")

    ThisWorkbook.UnProtect_This_Sheet

    If FAWBInSolution Then
        Worksheets("Arc Diagram").Range(ARC_DEDICATED_DB_VALUE).Value2 = Worksheets("FAWB").Range(FAWB_DEDICATED_DB_VALUE).Value2
        Worksheets("Arc Diagram").Range(ARC_DEDICATED_DB_COMMENT).Value2 = Worksheets("FAWB").Range(FAWB_DEDICATED_DB_COMMENT).Value2

        Worksheets("Arc Diagram").Range(ARC_DB_SIZE_VALUE).Value2 = Worksheets("FAWB").Range(FAWB_SIZE_VALUE).Value2
        Worksheets("Arc Diagram").Range(ARC_DB_SIZE_COMMENT).Value2 = Worksheets("FAWB").Range(FAWB_SIZE_COMMENT).Value2
    Else
        If DBSInSolution Then
            Worksheets("Arc Diagram").Range(ARC_DEDICATED_DB_VALUE).Value2 = Worksheets("DB Service").Range(DBS_DEDICATED_DB_VALUE).Value2
            Worksheets("Arc Diagram").Range(ARC_DEDICATED_DB_COMMENT).Value2 = Worksheets("DB Service").Range(DBS_DEDICATED_DB_COMMENT).Value2

            Worksheets("Arc Diagram").Range(ARC_DB_SIZE_VALUE).Value2 = Worksheets("DB Service").Range(DBS_SIZE_VALUE).Value2
            Worksheets("Arc Diagram").Range(ARC_DB_SIZE_COMMENT).Value2 = Worksheets("DB Service").Range(DBS_SIZE_COMMENT).Value2
        Else
            ' if we get here we may be an EPC but we don't have DB Service or FAWB in this solution
            If IsEPC = "Yes" Then
                Worksheets("Arc Diagram").Range(ARC_DEDICATED_DB_VALUE).Value2 = "Yes"
                Worksheets("Arc Diagram").Range(ARC_DEDICATED_DB_COMMENT).Value2 = "EPC Default"

                Worksheets("Arc Diagram").Range(ARC_DB_SIZE_VALUE).Value2 = "Large"
                Worksheets("Arc Diagram").Range(ARC_DB_SIZE_COMMENT).Value2 = "EPC Default"
            Else
                Worksheets("Arc Diagram").Range(ARC_DEDICATED_DB_VALUE).Value2 = ""
                Worksheets("Arc Diagram").Range(ARC_DEDICATED_DB_COMMENT).Value2 = ""

                Worksheets("Arc Diagram").Range(ARC_DB_SIZE_VALUE).Value2 = ""
                Worksheets("Arc Diagram").Range(ARC_DB_SIZE_COMMENT).Value2 = ""
            End If
        End If
    End If

    ThisWorkbook.Protect_This_Sheet
End Sub
Sub CheckServices()
    On Error Resume Next

    Dim IsEPC As Variant
    IsEPC = Worksheets("Clarity").Range(EPC_QUESTION).Value2

    ThisWorkbook.UnProtect_This_Sheet

    If IsEPC = "Yes" Then
        Worksheets("Arc Diagram").Range(ARC_DEDICATED_SERVICES_VALUE).Value2 = "Yes"
        ThisWorkbook.UnProtect_This_Sheet
        Worksheets("Arc Diagram").Range(ARC_DEDICATED_SERVICES_COMMENT).Value2 = "EPC Customers get dedicated services to prevent noisy neighbor issues."
    Else
        Worksheets("Arc Diagram").Range(ARC_DEDICATED_SERVICES_VALUE).Value2 = "No"
        ThisWorkbook.UnProtect_This_Sheet
        Worksheets("Arc Diagram").Range(ARC_DEDICATED_SERVICES_COMMENT).Value2 = "Default for non-EPC"
    End If

    ThisWorkbook.Protect_This_Sheet
End Sub
