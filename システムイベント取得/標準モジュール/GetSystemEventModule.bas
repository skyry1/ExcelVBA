Attribute VB_Name = "GetSystemEventModule"
'�ǉ��Q�Ɛݒ�
'Microsoft WMI Scripting V1.2 Library

Sub getEvent()

    '��ʕ\���X�V���I�t�ɂ���
    Application.StatusBar = "���s��"
    Application.ScreenUpdating = False

    '�Z��������������
    Dim startRow As Integer
    startRow = 3
    Range("A" & startRow & ":G1000").Clear

    'SWbemLocator�N���X�I�u�W�F�N�g
    Dim oLocator As New SWbemLocator

    'WMI�T�[�r�X�I�u�W�F�N�g
    Dim oWMI As SWbemServicesEx
    Set oWMI = oLocator.ConnectServer
    
    
    '�C�x���g�����p�N�G��
    Dim query As String
    query = "SELECT * FROM Win32_NTLogEvent WHERE Logfile = 'System' AND "
    query = query & "(EventCode = '6005' OR EventCode = '6006' OR EventCode = '7001' OR EventCode = '7002')"
    
    
    '�����Œ��o�����S�ẴC�x���g
    Dim eventList As SWbemObjectSet
    Set eventList = oWMI.ExecQuery(query)
    
    '���݂̍s�ʒu��ݒ�
    Dim recordNum As Integer
    recordNum = startRow
    
    '�擾�����C�x���g���O�����[�v
    Dim dateTime, dtUtc As String
    For Each oEvent In eventList
    
        '��������{���Ԃɕϊ�
        dateTime = Left(oEvent.timeWritten, 14)
        dtUtc = CDate( _
                Mid(dateTime, 1, 4) & "/" & Mid(dateTime, 5, 2) & "/" & Mid(dateTime, 7, 2) & " " & _
                Mid(dateTime, 9, 2) & ":" & Mid(dateTime, 11, 2) & ":" & Mid(dateTime, 13, 2))
        '�W�����Ԃ�9���ԉ��Z
        dataTime = DateAdd("h", 9, dtUtc)
        
        '�Z���ɒl���i�[
        Range("A" & recordNum).Value = oEvent.Type
        Range("B" & recordNum).Value = conversionEventId(oEvent.EventCode)
        Range("C" & recordNum).Value = oEvent.EventCode
        Range("D" & recordNum).Value = Format(dataTime, "yyyy/mm/dd")
        Range("E" & recordNum).Value = Format(dataTime, "hh:mm:ss")
        Range("F" & recordNum).Value = oEvent.SourceName
        Range("G" & recordNum).Value = oEvent.Category
        
        recordNum = recordNum + 1
    Next
    
    '��ʕ\���X�V���I���ɂ���
    Application.ScreenUpdating = True
    Application.StatusBar = "���s����"
End Sub



'�C�x���gID���C�x���g���ɕύX
Private Function conversionEventId(eventId As String)
    If eventId = "6005" Then
        conversionEventId = "PC�N��"
    ElseIf eventId = "6006" Then
        conversionEventId = "PC�I��"
    ElseIf eventId = "7001" Then
        conversionEventId = "�X���[�v�J�n"
    ElseIf eventId = "7002" Then
        conversionEventId = "�X���[�v�I��"
    Else
        conversionEventId = ""
    End If
End Function



