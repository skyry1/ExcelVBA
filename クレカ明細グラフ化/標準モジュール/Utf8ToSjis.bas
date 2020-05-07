Attribute VB_Name = "Utf8ToSjis"
'�ǉ��Q�Ɛݒ�
'Microsoft ActiveX Data Objects x.x Library

Sub Utf8ToSjis(a_sFrom, a_sTo)
    Dim streamRead  As New ADODB.Stream '// �ǂݍ��݃f�[�^
    Dim streamWrite As New ADODB.Stream '// �������݃f�[�^
    Dim sText                           '// �t�@�C���f�[�^
    
    '// �t�@�C���ǂݍ���
    streamRead.Type = adTypeText
    streamRead.Charset = "UTF-8"
    streamRead.Open
    Call streamRead.LoadFromFile(a_sFrom)
    
    '// ���s�R�[�hLF��CRLF�ɕϊ�
    sText = streamRead.ReadText
    'sText = Replace(sText, vbLf, vbCrLf)
    sText = Replace(sText, vbCr & vbCr, vbCr)
    
    '// �t�@�C����������
    streamWrite.Type = adTypeText
    streamWrite.Charset = "Shift-JIS"
    streamWrite.Open
    
    '// �f�[�^��������
    Call streamWrite.WriteText(sText)
    
    '// �ۑ�
    Call streamWrite.SaveToFile(a_sTo, adSaveCreateOverWrite)
    
    '// �N���[�Y
    streamRead.Close
    streamWrite.Close
End Sub
