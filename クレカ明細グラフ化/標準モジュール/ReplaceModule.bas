Attribute VB_Name = "ReplaceModule"
'�_�u���N�H�[�e�[�V�����ň͂܂�Ă���J���}�͒u�������Ȃ�
Function replaceColon(ByVal str As String) As String

    Dim strTemp As String
    Dim quotCount As Long
    
    Dim l As Long
    For l = 1 To Len(str)  'str�̒��������J��Ԃ�
    
        strTemp = Mid(str, l, 1) 'str���猻�݂�1������؂�o��
    
        If strTemp = """" Then   'strTemp���_�u���N�H�[�e�[�V�����Ȃ�
    
            quotCount = quotCount + 1   '�_�u���N�H�[�e�[�V�����̃J�E���g��1���₷
    
        ElseIf strTemp = "," Then   'strTemp���J���}�Ȃ�
    
            If quotCount Mod 2 = 0 Then   'quotCount��2�̔{���Ȃ�
    
                str = Left(str, l - 1) & ":" & Right(str, Len(str) - l)   '���݂�1�������R�����ɒu��������
    
            End If
    
        End If
    
    Next l
    
    replaceColon = str

End Function
