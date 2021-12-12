Sub addShapeSpace()
    
    Dim aaa As String
    Dim bbb As String
    Dim ccc As String
    Dim ddd As Integer
    Dim eee() As String
    ReDim eee(ActiveSheet.shapes.Count)
    
    For i = 1 To ActiveSheet.shapes.Count
        
        If i > 1 Then
        
            Dim fff As String
            fff = "A" + eee(i - 1)
            
            With ActiveSheet.shapes(i)
            .Left = Range(fff).Left
            .Top = Range(fff).Top
            End With
            
            aaa = ActiveSheet.shapes(i).BottomRightCell.Address
            bbb = Mid(aaa, InStr(aaa, "$") + 1)
            ccc = Mid(bbb, InStr(bbb, "$") + 1)
            ddd = CInt(ccc) + 4
            eee(i) = CStr(ddd)
            
        ElseIf i = 1 Then
            With ActiveSheet.shapes(i)
                .Left = Range("A2").Left
                .Top = Range("A2").Top
                
            End With
        
            aaa = ActiveSheet.shapes(i).BottomRightCell.Address
            bbb = Mid(aaa, InStr(aaa, "$") + 1)
            ccc = Mid(bbb, InStr(bbb, "$") + 1)
            ddd = CInt(ccc) + 4
            eee(i) = CStr(ddd)
        
        End If
    Next i
End Sub
