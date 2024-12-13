'
' Macros for the sheet named "Exec Review"
' Copyright FICO (Fair Isaac Inc) 2023, 2024, 2025
'
Option Explicit

Private Sub Worksheet_Activate()
    On Error Resume Next

    Dim Cell As Range
    For Each Cell In Application.ActiveSheet.usedRange
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

    ThisWorkbook.Sheets("Executive Review").usedRange.EntireRow.AutoFit

    ThisWorkbook.Protect_This_Sheet
End Sub
Sub CloneFromArchReview()
    On Error Resume Next

    ThisWorkbook.UnProtect_This_Sheet
    ButtonOn button:=ActiveSheet.Shapes("Clone")

    CopyImageFromAcDiagram

    ButtonOff button:=ActiveSheet.Shapes("Clone")

   ThisWorkbook.Protect_This_Sheet
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
Sub CopyImageFromAcDiagram()
    On Error Resume Next

    Dim sourceSheet As Worksheet
    Dim destinationSheet As Worksheet
    Dim origShape As Shape
    Dim clonedShape As Shape

    ThisWorkbook.UnProtect_This_Sheet
    ThisWorkbook.Sheets("Arc Diagram").Unprotect ThisWorkbook.myPassword()

    sourceSheet = ThisWorkbook.Worksheets("Arc Diagram")
    destinationSheet = ThisWorkbook.Worksheets("Executive Review")

    Dim shp As Shape
    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets("Arc Diagram")

    For Each shp In ws.Shapes
        If shp.Type = msoPicture Then
           MsgBox shp.Name & " is a picture"
           Set clonedShape = shp
           clonedShape.CopyPicture
        End If
    Next shp

    'Set clonedShape = sourceSheet.Shapes("ArchPicture")
    destinationSheet.PasteSpecial (xlPasteAll)

    'ThisWorkbook.Sheets("Arc Diagram").Protect ThisWorkbook.myPassword()
    'ThisWorkbook.Protect_This_Sheet
End Sub
Sub ImportPictureFromFile1()
    On Error Resume Next
    ButtonOn button:=ActiveSheet.Shapes("Import")
    addImageByName "E1", "Diagram"
    ButtonOff button:=ActiveSheet.Shapes("Import")

End Sub
Sub DeletePicture()
    On Error Resume Next
    ThisWorkbook.UnProtect_This_Sheet

    ButtonOn button:=ActiveSheet.Shapes("Delete")
    deleteImageByName "Diagram"
    ButtonOff button:=ActiveSheet.Shapes("Delete")
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

    Set ws = ThisWorkbook.Worksheets("Executive Review")
    Set s = ws.Shapes.AddPicture2(fNameAndPath, False, True, ThisWorkbook.Sheets("Executive Review").Range(ImageLocation).Left, _
    ThisWorkbook.Sheets("Executive Review").Range(ImageLocation).Top, -1, -1, msoPictureCompressFalse)

    s.Name = ImageName
    s.LockAspectRatio = msoTrue
    s.Width = GetWidth()
    s.Height = GetHeight()

    ThisWorkbook.Protect_This_Sheet
End Sub
Private Sub deleteImageByName(ByVal arg1 As String)
    On Error Resume Next

    Dim pic As Shape

    For Each pic In ThisWorkbook.Worksheets("Executive Review").Shapes
        If InStr(1, pic.Name, arg1, vbTextCompare) <> 0 Then
            pic.Delete
        End If
    Next pic

End Sub
Function GetWidth()
    On Error Resume Next

    Dim i As Integer

    For i = 4 To 7 ' Columns A to E
        GetWidth = GetWidth + Columns(i).Width
    Next i

End Function
Function GetHeight()
    On Error Resume Next

    Dim totalHeight As Double
    totalHeight = 0
    Dim i As Integer

    For i = 1 To 23 ' Columns A to E
        totalHeight = totalHeight + Rows(i).RowHeight
    Next i

    GetHeight = totalHeight

End Function
