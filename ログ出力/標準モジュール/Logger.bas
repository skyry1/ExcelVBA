Attribute VB_Name = "Logger"
Sub writeLog()

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    'ログファイルパス
    Dim logPath As String
    logPath = ThisWorkbook.Path & "\SAMPLE.log"
    
    With fso
        'ログファイルがなかったら作成する
        If Not .FileExists(logPath) Then
            .CreateTextFile (logPath)
        End If
        
        'ログを追記していく
        'With .OpenTextFile(logPath, 1) '読み取り専用として開く
        'With .OpenTextFile(logPath, 2) '書き込み専用として開く
        With .OpenTextFile(logPath, 8) 'ファイルの最後に追記する（書き込み専用）
            .WriteLine "実行時刻：" & Now
            .Close
        End With
    End With
    
    Set fso = Nothing
End Sub
