'
' Macros for the sheet named "Ref Arch"
' Copyright FICO (Fair Isaac Inc) 2023, 2024, 2025
'

Option Explicit

Private Sub Worksheet_Activate()
    On Error Resume Next

    Range("A1").Select

    '
    ' I'm using a Sub at the ThisWorkbook level so I can toggle
    ' this on/off easily for mass edits.
    '
    ThisWorkbook.Protect_This_Sheet

End Sub
Sub AddDiagram
    ' add a diagram to this page
    On Error Resume Next
    addImageByName "B4", "Picture 1"
End Sub
Sub DeleteDiagram
    ' delete the diagram on this page
    On Error Resume Next
    deleteImageByName "Picture 1"
End Sub
Private Sub addImageByName(ByVal ImageLocation As String, ByVal ImageName As String)

    ThisWorkbook.UnProtect_This_Sheet
    Dim fNameAndPath As Variant

    fNameAndPath = Application.GetOpenFilename(Title:="Select image file to be imported")

    If fNameAndPath = False Then
        Exit Sub
    End If

    Dim s As Shape
    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets("Ref_Arch_Diagram")
    Set s = ws.Shapes.AddPicture(fNameAndPath, False, True, ThisWorkbook.Sheets("Ref_Arch_Diagram").Range(ImageLocation).Left, ThisWorkbook.Sheets("Ref_Arch_Diagram").Range(ImageLocation).Top, -1, -1)

    s.Name = ImageName
    s.LockAspectRatio = msoTrue

    ThisWorkbook.Protect_This_Sheet
End Sub
Private Sub deleteImageByName(ByVal arg1 As String)

    Dim pic As Shape

    For Each pic In ThisWorkbook.Worksheets("Ref_Arch_Diagram").Shapes
        If InStr(1, pic.Name, arg1, vbTextCompare) <> 0 Then
            pic.Delete
        End If
    Next pic
End Sub