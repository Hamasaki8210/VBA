Sub shape()

    Dim x As Integer
    Dim y As Integer
    Dim x1 As Integer
    Dim y1 As Integer
    
    Dim rectangle As Object
    
    x = ActiveCell.Column
    y = ActiveCell.Row
    
    x1 = Range(Cells(1, 1), Cells(y, x)).Width - 52
    y1 = Range(Cells(1, 1), Cells(y, x)).Height - 17
    
    'Set rectangle = Selection
    
    With ActiveSheet.shapes.AddShape(msoShapeRectangle, x1, y1, 100, 50).Fill
        .Transparency = 1
    End With
    
    
     '   rectangle.Transparency = 1
    
    'rectangle.Fill.Visible = msoFalse
    
    'MsgBox Selection.Name

End Sub
