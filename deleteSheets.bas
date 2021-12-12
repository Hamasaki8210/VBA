Sub DeleteMultipleSheets()

    Dim msgResult As Integer
    msgResult = MsgBox("削除対象のシート名をご確認ください。削除処理を開始してよいですか？", vbYesNo)
    If msgResult = vbNo Then
        Exit Sub
    End If

    Application.DisplayAlerts = False
    
    Dim ws As Worksheet
    'コピー開始番号
    Dim startNum As Integer
    'コピー終了番号
    Dim endNum As Integer
    '重複確認するシート名が存在するかを確認するフラグ
    Dim notExistedFlg As Boolean
    
    Dim sheetName As String
    
    Dim i As Integer
    
    startNum = Worksheets("マクロ").Range("B7").Value
    endNum = Worksheets("マクロ").Range("B8").Value
    
    If startNum = 0 Or endNum = 0 Then
    
        MsgBox "必要なデータが入力されていません。"
    
    Else
    
        '開始番号から終了番号までシート作成
        For i = startNum To endNum
        
            sheetName = CStr(i)
            
            'シート名重複確認
            For Each ws In Worksheets
                If Not ws.Name = sheetName Then
                    notExistedFlg = True
                Else
                    notExistedFlg = False
                    Exit For
                End If
            Next ws
            
            'もし重複がなかったらシート追加
            If notExistedFlg = False Then
                Worksheets(sheetName).Delete
            End If
            
            notExistedFlg = False
        Next i

    End If

End Sub
