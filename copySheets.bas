Attribute VB_Name = "Module1"
Sub �e�X�g()
 
  MsgBox "�{�^���N���b�N"
 
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
    '�R�s�[�J�n�ԍ�
    Dim startNum As Integer
    '�R�s�[�I���ԍ�
    Dim endNum As Integer
    '�d���m�F����V�[�g�������݂��邩���m�F����t���O
    Dim duplicatedFlg As Boolean
    '�t����V�[�g��
    Dim sheetName As String
    '�R�s�[���V�[�g��
    Dim copiedSheet As String
    
    startNum = Worksheets("�}�N��").Range("B2").Value
    endNum = Worksheets("�}�N��").Range("B3").Value
    copiedSheet = Worksheets("�}�N��").Range("B4").Value
    
    If startNum = 0 Or endNum = 0 Or copiedSheet = "" Then
    
        MsgBox "�K�v�ȃf�[�^�����͂���Ă��܂���B"
    
    Else
    
        '�J�n�ԍ�����I���ԍ��܂ŃV�[�g�쐬
        For i = startNum To endNum
        
            sheetName = CStr(i)
            
            '�V�[�g���d���m�F
            For Each ws In Worksheets
                If ws.Name = sheetName Then
                    duplicatedFlg = True
                End If
            Next ws
            
            '�����d�����Ȃ�������V�[�g�ǉ�
            If duplicatedFlg = False Then
                Worksheets(copiedSheet).Copy After:=Worksheets(Worksheets.Count)
                ActiveSheet.Name = sheetName
            End If
            
            duplicatedFlg = False
        Next

    End If
End Sub
