Attribute VB_Name = "Logger"
Sub writeLog()

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    '���O�t�@�C���p�X
    Dim logPath As String
    logPath = ThisWorkbook.Path & "\SAMPLE.log"
    
    With fso
        '���O�t�@�C�����Ȃ�������쐬����
        If Not .FileExists(logPath) Then
            .CreateTextFile (logPath)
        End If
        
        '���O��ǋL���Ă���
        'With .OpenTextFile(logPath, 1) '�ǂݎ���p�Ƃ��ĊJ��
        'With .OpenTextFile(logPath, 2) '�������ݐ�p�Ƃ��ĊJ��
        With .OpenTextFile(logPath, 8) '�t�@�C���̍Ō�ɒǋL����i�������ݐ�p�j
            .WriteLine "���s�����F" & Now
            .Close
        End With
    End With
    
    Set fso = Nothing
End Sub
