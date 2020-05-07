Attribute VB_Name = "GetSystemEventModule"
'追加参照設定
'Microsoft WMI Scripting V1.2 Library

Sub getEvent()

    '画面表示更新をオフにする
    Application.StatusBar = "実行中"
    Application.ScreenUpdating = False

    'セルを初期化する
    Dim startRow As Integer
    startRow = 3
    Range("A" & startRow & ":G1000").Clear

    'SWbemLocatorクラスオブジェクト
    Dim oLocator As New SWbemLocator

    'WMIサービスオブジェクト
    Dim oWMI As SWbemServicesEx
    Set oWMI = oLocator.ConnectServer
    
    
    'イベント検索用クエリ
    Dim query As String
    query = "SELECT * FROM Win32_NTLogEvent WHERE Logfile = 'System' AND "
    query = query & "(EventCode = '6005' OR EventCode = '6006' OR EventCode = '7001' OR EventCode = '7002')"
    
    
    '検索で抽出した全てのイベント
    Dim eventList As SWbemObjectSet
    Set eventList = oWMI.ExecQuery(query)
    
    '現在の行位置を設定
    Dim recordNum As Integer
    recordNum = startRow
    
    '取得したイベントログ数ループ
    Dim dateTime, dtUtc As String
    For Each oEvent In eventList
    
        '時刻を日本時間に変換
        dateTime = Left(oEvent.timeWritten, 14)
        dtUtc = CDate( _
                Mid(dateTime, 1, 4) & "/" & Mid(dateTime, 5, 2) & "/" & Mid(dateTime, 7, 2) & " " & _
                Mid(dateTime, 9, 2) & ":" & Mid(dateTime, 11, 2) & ":" & Mid(dateTime, 13, 2))
        '標準時間に9時間加算
        dataTime = DateAdd("h", 9, dtUtc)
        
        'セルに値を格納
        Range("A" & recordNum).Value = oEvent.Type
        Range("B" & recordNum).Value = conversionEventId(oEvent.EventCode)
        Range("C" & recordNum).Value = oEvent.EventCode
        Range("D" & recordNum).Value = Format(dataTime, "yyyy/mm/dd")
        Range("E" & recordNum).Value = Format(dataTime, "hh:mm:ss")
        Range("F" & recordNum).Value = oEvent.SourceName
        Range("G" & recordNum).Value = oEvent.Category
        
        recordNum = recordNum + 1
    Next
    
    '画面表示更新をオンにする
    Application.ScreenUpdating = True
    Application.StatusBar = "実行完了"
End Sub



'イベントIDをイベント名に変更
Private Function conversionEventId(eventId As String)
    If eventId = "6005" Then
        conversionEventId = "PC起動"
    ElseIf eventId = "6006" Then
        conversionEventId = "PC終了"
    ElseIf eventId = "7001" Then
        conversionEventId = "スリープ開始"
    ElseIf eventId = "7002" Then
        conversionEventId = "スリープ終了"
    Else
        conversionEventId = ""
    End If
End Function



