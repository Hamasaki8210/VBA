Sub DeleteMultipleSheets()

    Dim msgResult As Integer
    msgResult = MsgBox("�폜�Ώۂ̃V�[�g�������m�F���������B�폜�������J�n���Ă悢�ł����H", vbYesNo)
    If msgResult = vbNo Then
        Exit Sub
    End If

    Application.DisplayAlerts = False
    
    Dim ws As Worksheet
    '�R�s�[�J�n�ԍ�
    Dim startNum As Integer
    '�R�s�[�I���ԍ�
    Dim endNum As Integer
    '�d���m�F����V�[�g�������݂��邩���m�F����t���O
    Dim notExistedFlg As Boolean
    
    Dim sheetName As String
    
    Dim i As Integer
    
    startNum = Worksheets("�}�N��").Range("B7").Value
    endNum = Worksheets("�}�N��").Range("B8").Value
    
    If startNum = 0 Or endNum = 0 Then
    
        MsgBox "�K�v�ȃf�[�^�����͂���Ă��܂���B"
    
    Else
    
        '�J�n�ԍ�����I���ԍ��܂ŃV�[�g�쐬
        For i = startNum To endNum
        
            sheetName = CStr(i)
            
            '�V�[�g���d���m�F
            For Each ws In Worksheets
                If Not ws.Name = sheetName Then
                    notExistedFlg = True
                Else
                    notExistedFlg = False
                    Exit For
                End If
            Next ws
            
            '�����d�����Ȃ�������V�[�g�ǉ�
            If notExistedFlg = False Then
                Worksheets(sheetName).Delete
            End If
            
            notExistedFlg = False
        Next i

    End If

End Sub
