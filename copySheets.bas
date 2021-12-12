Attribute VB_Name = "Module1"
Sub テスト()
 
  MsgBox "ボタンクリック"
 
End Sub

Sub SheetCopy()

    Dim sheetNum As Integer
    sheetNum = 10
    
    For i = 1 To sheetNum
        Worksheets("Sheet1").Copy After:=Worksheets("Sheet1")
        Worksheets("Sheet1").Copy = i + "-2"
    Next

End Sub

Sub CopyMultipleSheets()

    Dim ws As Worksheet
    'コピー開始番号
    Dim startNum As Integer
    'コピー終了番号
    Dim endNum As Integer
    '重複確認するシート名が存在するかを確認するフラグ
    Dim duplicatedFlg As Boolean
    '付けるシート名
    Dim sheetName As String
    'コピー元シート名
    Dim copiedSheet As String
    
    startNum = Worksheets("マクロ").Range("B2").Value
    endNum = Worksheets("マクロ").Range("B3").Value
    copiedSheet = Worksheets("マクロ").Range("B4").Value
    
    If startNum = 0 Or endNum = 0 Or copiedSheet = "" Then
    
        MsgBox "必要なデータが入力されていません。"
    
    Else
    
        '開始番号から終了番号までシート作成
        For i = startNum To endNum
        
            sheetName = CStr(i)
            
            'シート名重複確認
            For Each ws In Worksheets
                If ws.Name = sheetName Then
                    duplicatedFlg = True
                End If
            Next ws
            
            'もし重複がなかったらシート追加
            If duplicatedFlg = False Then
                Worksheets(copiedSheet).Copy After:=Worksheets(Worksheets.Count)
                ActiveSheet.Name = sheetName
            End If
            
            duplicatedFlg = False
        Next

    End If
End Sub
