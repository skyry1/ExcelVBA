Attribute VB_Name = "Utf8ToSjis"
'追加参照設定
'Microsoft ActiveX Data Objects x.x Library

Sub Utf8ToSjis(a_sFrom, a_sTo)
    Dim streamRead  As New ADODB.Stream '// 読み込みデータ
    Dim streamWrite As New ADODB.Stream '// 書き込みデータ
    Dim sText                           '// ファイルデータ
    
    '// ファイル読み込み
    streamRead.Type = adTypeText
    streamRead.Charset = "UTF-8"
    streamRead.Open
    Call streamRead.LoadFromFile(a_sFrom)
    
    '// 改行コードLFをCRLFに変換
    sText = streamRead.ReadText
    'sText = Replace(sText, vbLf, vbCrLf)
    sText = Replace(sText, vbCr & vbCr, vbCr)
    
    '// ファイル書き込み
    streamWrite.Type = adTypeText
    streamWrite.Charset = "Shift-JIS"
    streamWrite.Open
    
    '// データ書き込み
    Call streamWrite.WriteText(sText)
    
    '// 保存
    Call streamWrite.SaveToFile(a_sTo, adSaveCreateOverWrite)
    
    '// クローズ
    streamRead.Close
    streamWrite.Close
End Sub
