Attribute VB_Name = "ReadCSVDetails"
'�ǉ��Q�Ɛݒ�
'Microsoft Scripting Runtime

'�O���t�^�C�g��
Dim graphTitle As String

'�J�e�S��
Dim category As Dictionary
Dim cntr As Integer

Dim startRow As Integer
Dim endRow As Integer

Sub Main()
    
    '������
    Call init
    
    Dim myFile As Variant
    myFile = Application.GetOpenFilename("CSV�t�@�C��(*.csv),*.csv", Title:="���ׁiCSV�t�@�C���j�I��")

    If VarType(myFile) = vbBoolean Then
        '�������Ȃ�
    Else
        '
        Dim pos As Long
        pos = InStrRev(myFile, "\")
        graphTitle = Mid(myFile, pos + 1)
        graphTitle = Replace(graphTitle, ".csv", "")
        Call DetailsSelection(myFile)
    End If
End Sub

'����������
Sub init()
    graphTitle = ""
    Set category = New Dictionary
    cntr = 0
    startRow = 0
    endRow = 0
    MsgBoxCustom_Reset 0
End Sub


'���ׂ�I��
Sub DetailsSelection(myFile As Variant)
    
    '�{�^�����쐬
    MsgBoxCustom_Set vbYes, "View�J�[�h"
    MsgBoxCustom_Set vbNo, "�y�V�J�[�h"
    MsgBoxCustom_Set vbCancel, "�L�����Z��"
    
    '�_�C�A���O�\��
    Dim ans
    MsgBoxCustom.MsgBoxCustom ans, "���ׂ̎�ނ�I�����Ă��������B", vbYesNoCancel
    
    '���x���̕���������Z�b�g
    MsgBoxCustom_Reset 0
    
    '���דǍ�
    If ans = vbYes Then
        Call ReadViewCardDetails(myFile)
    ElseIf ans = vbNo Then
        Call ReadRakutenCardDetails(myFile)
    Else
        Debug.Print "�L�����Z��"
    End If
End Sub

'View�J�[�h���דǍ�
Sub ReadViewCardDetails(myFile As Variant)

    '�J�e�S���쐬
    Set category = New Dictionary
    Call CreateCategory(cntr)
        
    'LF�i���C���t�B�[���h�j�R�[�h�ŉ��s����CSV�t�@�C�����J��
    Dim csvData As Variant
    csvData = ReadCsvLFCode(myFile)
        
    'CSV�t�@�C���̓��e��ǂݍ���
    Dim columns As Variant
    For i = 7 To UBound(csvData) - 1
        columns = Split(Replace(ReplaceModule.replaceColon(csvData(i)), """", ""), ":")
        
        If UBound(columns) < 1 Then
            Exit For
        End If
        
        Dim j As Integer
        Dim sonota As Boolean
        sonota = True
        For j = 0 To cntr
            Dim key As String
            key = category.Keys(j)
            
            '�J�e�S���ƈ�v��������z�𑫂��Ă����i�啶������������ʂ��Ȃ��j
            If InStr(1, columns(1), key, vbTextCompare) > 0 Then
                category.Item(key) = category.Item(key) + Val(Replace(columns(2), ",", ""))
                sonota = False
            End If
        Next j
        '�J�e�S���ɂȂ������炻�̑��ɑ����Ă���
        If sonota Then
            category.Item("���̑�") = category.Item("���̑�") + Val(Replace(columns(2), ",", ""))
        End If
    Next i
    
    '�V�[�g�쐬
    Worksheets.Add
    ActiveSheet.Name = graphTitle
    
    '�������ʂ��V�[�g�ɏ�������
    Call WriteData(category)
    
    '�������ʂ����ƂɃO���t���쐬����
    Call CreateGraph
    
End Sub

'�y�V�J�[�h���דǍ�
Sub ReadRakutenCardDetails(myFile As Variant)

    '�J�e�S���쐬
    Set category = New Dictionary
    Call CreateCategory(cntr)
    
    'Shift-JIS�ɕϊ�
    Dim shiftJisFile As String
    shiftJisFile = Replace(myFile, ".csv", "") & "_Shift-JIS" & ".csv"
    Call Utf8ToSjis.Utf8ToSjis(myFile, shiftJisFile)
        
    'LF�i���C���t�B�[���h�j�R�[�h�ŉ��s����CSV�t�@�C�����J��
    Dim csvData As Variant
    csvData = ReadCsvLFCode(shiftJisFile)
    
    'CSV�t�@�C��(Shift-JIS�ϊ���)�폜
    Kill shiftJisFile
        
    'CSV�t�@�C���̓��e��ǂݍ���
    Dim columns As Variant
    For i = 1 To UBound(csvData) - 1
        columns = Split(Replace(ReplaceModule.replaceColon(csvData(i)), """", ""), ":")
        
        If columns(0) <> "" Then
            Dim j As Integer
            Dim sonota As Boolean
            sonota = True
            For j = 0 To cntr
                Dim key As String
                key = category.Keys(j)
                
                Dim amount As String
                amount = columns(4)
                
                '�J�e�S���ƈ�v��������z�𑫂��Ă����i�啶������������ʂ��Ȃ��j
                If InStr(1, columns(1), key, vbTextCompare) > 0 Then
                    category.Item(key) = category.Item(key) + Val(Replace(amount, ",", ""))
                    sonota = False
                End If
            Next j
            '�J�e�S���ɂȂ������炻�̑��ɑ����Ă���
            If sonota Then
                category.Item("���̑�") = category.Item("���̑�") + Val(Replace(amount, ",", ""))
            End If
        End If
    Next i
    
    '�V�[�g�쐬
    Worksheets.Add
    ActiveSheet.Name = graphTitle
    
    '�������ʂ��V�[�g�ɏ�������
    Call WriteData(category)
    
    '�������ʂ����ƂɃO���t���쐬����
    Call CreateGraph
End Sub

'�V�[�g�ɋL�������J�e�S����ǂݍ���
Sub CreateCategory(cntr As Integer)

    '�J�n�s
    startRow = 2
    
    '�v�f�����擾
    cntr = Cells(startRow, 1).End(xlDown).Row - Cells(startRow, 1).Row

    '�v�f�ݒ�
    Dim i As Integer
    For i = 0 To cntr
        category.Add Cells(i + startRow, 1).Value, 0
    Next i
    If category.Exists("���̑�") = True Then
        category.Item("���̑�") = 0
        endRow = startRow + cntr
    Else
        category.Add "���̑�", 0
        endRow = startRow + cntr + 1
    End If
End Sub

'�V�[�g�ɏW�v���ʂ���������
Sub WriteData(category As Dictionary)

    '�w�b�_�쐬
    Cells(1, 1).Value = "�J�e�S��"
    Cells(1, 2).Value = "���z"
    
    Dim i As Integer
    '�������ʂ��V�[�g�ɏ�������
    For i = 0 To category.Count - 1
        key = category.Keys(i)
        Cells(i + startRow, 1).Value = key
        Cells(i + startRow, 2).Value = category.Item(key)
    Next i
End Sub

'�V�[�g�ɃO���t������
Sub CreateGraph()
    With ActiveSheet.Shapes.AddChart.Chart
        '�~�O���t
        .ChartType = xlPie
        '�f�[�^�͈�
        .SetSourceData Source:=Range(Cells(startRow, 1), Cells(endRow, 2))
        '�O���t�^�C�g��
        .HasTitle = True
        .ChartTitle.Text = graphTitle
        .SeriesCollection(1).HasDataLabels = True
    End With
End Sub


'CSV�t�@�C���ǂݍ���
Function ReadCsvLFCode(myFile As Variant)

    'LF�i���C���t�B�[���h�j�R�[�h�ŉ��s����CSV�t�@�C�����J��
    Dim buf As String
    Dim csvData As Variant
    Open myFile For Input As #1
        Line Input #1, buf
        ReadCsvLFCode = Split(buf, vbLf)
    Close #1
End Function





