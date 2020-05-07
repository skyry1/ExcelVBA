Attribute VB_Name = "ReadCSVDetails"
'追加参照設定
'Microsoft Scripting Runtime

'グラフタイトル
Dim graphTitle As String

'カテゴリ
Dim category As Dictionary
Dim cntr As Integer

Dim startRow As Integer
Dim endRow As Integer

Sub Main()
    
    '初期化
    Call init
    
    Dim myFile As Variant
    myFile = Application.GetOpenFilename("CSVファイル(*.csv),*.csv", Title:="明細（CSVファイル）選択")

    If VarType(myFile) = vbBoolean Then
        '何もしない
    Else
        '
        Dim pos As Long
        pos = InStrRev(myFile, "\")
        graphTitle = Mid(myFile, pos + 1)
        graphTitle = Replace(graphTitle, ".csv", "")
        Call DetailsSelection(myFile)
    End If
End Sub

'初期化処理
Sub init()
    graphTitle = ""
    Set category = New Dictionary
    cntr = 0
    startRow = 0
    endRow = 0
    MsgBoxCustom_Reset 0
End Sub


'明細を選択
Sub DetailsSelection(myFile As Variant)
    
    'ボタンを作成
    MsgBoxCustom_Set vbYes, "Viewカード"
    MsgBoxCustom_Set vbNo, "楽天カード"
    MsgBoxCustom_Set vbCancel, "キャンセル"
    
    'ダイアログ表示
    Dim ans
    MsgBoxCustom.MsgBoxCustom ans, "明細の種類を選択してください。", vbYesNoCancel
    
    'ラベルの文字列をリセット
    MsgBoxCustom_Reset 0
    
    '明細読込
    If ans = vbYes Then
        Call ReadViewCardDetails(myFile)
    ElseIf ans = vbNo Then
        Call ReadRakutenCardDetails(myFile)
    Else
        Debug.Print "キャンセル"
    End If
End Sub

'Viewカード明細読込
Sub ReadViewCardDetails(myFile As Variant)

    'カテゴリ作成
    Set category = New Dictionary
    Call CreateCategory(cntr)
        
    'LF（ラインフィールド）コードで改行したCSVファイルを開く
    Dim csvData As Variant
    csvData = ReadCsvLFCode(myFile)
        
    'CSVファイルの内容を読み込む
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
            
            'カテゴリと一致したら金額を足していく（大文字小文字を区別しない）
            If InStr(1, columns(1), key, vbTextCompare) > 0 Then
                category.Item(key) = category.Item(key) + Val(Replace(columns(2), ",", ""))
                sonota = False
            End If
        Next j
        'カテゴリになかったらその他に足していく
        If sonota Then
            category.Item("その他") = category.Item("その他") + Val(Replace(columns(2), ",", ""))
        End If
    Next i
    
    'シート作成
    Worksheets.Add
    ActiveSheet.Name = graphTitle
    
    '処理結果をシートに書き込む
    Call WriteData(category)
    
    '処理結果をもとにグラフを作成する
    Call CreateGraph
    
End Sub

'楽天カード明細読込
Sub ReadRakutenCardDetails(myFile As Variant)

    'カテゴリ作成
    Set category = New Dictionary
    Call CreateCategory(cntr)
    
    'Shift-JISに変換
    Dim shiftJisFile As String
    shiftJisFile = Replace(myFile, ".csv", "") & "_Shift-JIS" & ".csv"
    Call Utf8ToSjis.Utf8ToSjis(myFile, shiftJisFile)
        
    'LF（ラインフィールド）コードで改行したCSVファイルを開く
    Dim csvData As Variant
    csvData = ReadCsvLFCode(shiftJisFile)
    
    'CSVファイル(Shift-JIS変換版)削除
    Kill shiftJisFile
        
    'CSVファイルの内容を読み込む
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
                
                'カテゴリと一致したら金額を足していく（大文字小文字を区別しない）
                If InStr(1, columns(1), key, vbTextCompare) > 0 Then
                    category.Item(key) = category.Item(key) + Val(Replace(amount, ",", ""))
                    sonota = False
                End If
            Next j
            'カテゴリになかったらその他に足していく
            If sonota Then
                category.Item("その他") = category.Item("その他") + Val(Replace(amount, ",", ""))
            End If
        End If
    Next i
    
    'シート作成
    Worksheets.Add
    ActiveSheet.Name = graphTitle
    
    '処理結果をシートに書き込む
    Call WriteData(category)
    
    '処理結果をもとにグラフを作成する
    Call CreateGraph
End Sub

'シートに記入したカテゴリを読み込む
Sub CreateCategory(cntr As Integer)

    '開始行
    startRow = 2
    
    '要素数を取得
    cntr = Cells(startRow, 1).End(xlDown).Row - Cells(startRow, 1).Row

    '要素設定
    Dim i As Integer
    For i = 0 To cntr
        category.Add Cells(i + startRow, 1).Value, 0
    Next i
    If category.Exists("その他") = True Then
        category.Item("その他") = 0
        endRow = startRow + cntr
    Else
        category.Add "その他", 0
        endRow = startRow + cntr + 1
    End If
End Sub

'シートに集計結果を書き込む
Sub WriteData(category As Dictionary)

    'ヘッダ作成
    Cells(1, 1).Value = "カテゴリ"
    Cells(1, 2).Value = "金額"
    
    Dim i As Integer
    '処理結果をシートに書き込む
    For i = 0 To category.Count - 1
        key = category.Keys(i)
        Cells(i + startRow, 1).Value = key
        Cells(i + startRow, 2).Value = category.Item(key)
    Next i
End Sub

'シートにグラフを書く
Sub CreateGraph()
    With ActiveSheet.Shapes.AddChart.Chart
        '円グラフ
        .ChartType = xlPie
        'データ範囲
        .SetSourceData Source:=Range(Cells(startRow, 1), Cells(endRow, 2))
        'グラフタイトル
        .HasTitle = True
        .ChartTitle.Text = graphTitle
        .SeriesCollection(1).HasDataLabels = True
    End With
End Sub


'CSVファイル読み込み
Function ReadCsvLFCode(myFile As Variant)

    'LF（ラインフィールド）コードで改行したCSVファイルを開く
    Dim buf As String
    Dim csvData As Variant
    Open myFile For Input As #1
        Line Input #1, buf
        ReadCsvLFCode = Split(buf, vbLf)
    Close #1
End Function





