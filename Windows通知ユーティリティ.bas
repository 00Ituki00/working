' Windows通知ユーティリティ v2
' PowerShell経由でWindowsトレイ通知（バルーン通知）を表示
' 問題修正: 通知が表示されない問題を解決

' === 通知表示関数 ===
' 引数:
'   Title    - 通知タイトル
'   Message  - 通知本文
'   Duration - 表示時間（秒、デフォルト10秒）
'   IconType - アイコン種別（"Info", "Warning", "Error", "None"、デフォルト"Info"）
Public Sub ShowToast(ByVal Title As String, ByVal Message As String, _
                     Optional Duration As Long = 10, _
                     Optional IconType As String = "Info")
    
    On Error GoTo ErrorHandler
    
    Dim WsShell As Object
    Set WsShell = CreateObject("WScript.Shell")
    
    ' アイコン種別の変換
    Dim psIcon As String
    Select Case LCase(IconType)
        Case "info", "information"
            psIcon = "Information"
        Case "warning", "warn"
            psIcon = "Warning"
        Case "error", "err"
            psIcon = "Error"
        Case Else
            psIcon = "None"
    End Select
    
    ' スクリプトファイルに出力して実行（長いコマンド対策・確実な実行のため）
    Dim scriptPath As String
    scriptPath = Environ("TEMP") & "\toast_notification.ps1"
    
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open scriptPath For Output As #fileNum
    Print #fileNum, "Add-Type -AssemblyName System.Windows.Forms"
    Print #fileNum, "$notify = New-Object System.Windows.Forms.NotifyIcon"
    Print #fileNum, "$notify.Icon = [System.Drawing.SystemIcons]::" & psIcon
    Print #fileNum, "$notify.Visible = $true"
    Print #fileNum, "$notify.ShowBalloonTip(" & Duration & ", '" & EscapeForPowerShell(Title) & "', '" & EscapeForPowerShell(Message) & "', [System.Windows.Forms.ToolTipIcon]::" & psIcon & ")"
    Print #fileNum, "Start-Sleep -Seconds " & Duration + 2
    Print #fileNum, "$notify.Dispose()"
    Close #fileNum
    
    ' PowerShellスクリプト実行
    Dim psCommand As String
    psCommand = "powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & scriptPath & """"
    
    ' デバッグ用: コマンドを確認
    Debug.Print "PS Command: " & psCommand
    
    ' 実行（非同期）
    WsShell.Run psCommand, 0, False
    
    Set WsShell = Nothing
    Exit Sub
    
ErrorHandler:
    Debug.Print "ShowToast Error: " & Err.Description
    Set WsShell = Nothing
End Sub

' === PowerShell用文字列エスケープ ===
Private Function EscapeForPowerShell(ByVal str As String) As String
    ' シングルクォートをエスケープ（PowerShellでは''で'をエスケープ）
    str = Replace(str, "'", "''")
    ' その他の特殊文字を処理
    str = Replace(str, "`", "``")
    EscapeForPowerShell = str
End Function

' === 簡易通知関数（タイトル・メッセージのみ） ===
Public Sub Toast(ByVal Title As String, ByVal Message As String)
    ShowToast Title, Message, 10, "Info"
End Sub

' === 完了通知 ===
Public Sub ToastComplete(Optional ByVal TaskName As String = "処理")
    ShowToast TaskName & "完了", "正常に完了しました", 5, "Info"
End Sub

' === エラー通知 ===
Public Sub ToastError(ByVal ErrorMessage As String)
    ShowToast "エラー", ErrorMessage, 15, "Error"
End Sub

' === 警告通知 ===
Public Sub ToastWarning(ByVal WarningMessage As String)
    ShowToast "警告", WarningMessage, 10, "Warning"
End Sub

' === テスト用診断関数 ===
Public Sub TestNotification()
    ' PowerShellが機能するか簡易テスト
    Dim WsShell As Object
    Set WsShell = CreateObject("WScript.Shell")
    
    ' メッセージボックスでPowerShell連携確認
    Dim testCmd As String
    testCmd = "powershell.exe -Command ""[System.Windows.Forms.MessageBox]::Show('PowerShell連携テスト')"""
    
    WsShell.Run testCmd, 1, True
    
    Set WsShell = Nothing
End Sub

' === Windows 10/11 トースト通知（代替案） ===
Public Sub ShowToastWin10(ByVal Title As String, ByVal Message As String)
    ' Windows 10/11 のモダントースト通知
    ' 注意: この方法はWindows 10/11でしか動作しません
    
    On Error GoTo ErrorHandler
    
    Dim WsShell As Object
    Set WsShell = CreateObject("WScript.Shell")
    
    ' スクリプトファイル作成
    Dim scriptPath As String
    scriptPath = Environ("TEMP") & "\win10_toast.ps1"
    
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open scriptPath For Output As #fileNum
    Print #fileNum, "[Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType=WindowsRuntime] | Out-Null"
    Print #fileNum, "$template = [Windows.UI.Notifications.ToastNotificationManager]::GetTemplateContent([Windows.UI.Notifications.ToastTemplateType]::ToastText02)"
    Print #fileNum, "$template.SelectSingleNode('//text[@id=""1""]').AppendChild($template.CreateTextNode('" & EscapeForPowerShell(Title) & "'))"
    Print #fileNum, "$template.SelectSingleNode('//text[@id=""2""]').AppendChild($template.CreateTextNode('" & EscapeForPowerShell(Message) & "'))"
    Print #fileNum, "$toast = [Windows.UI.Notifications.ToastNotification]::new($template)"
    Print #fileNum, "[Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier('Excel').Show($toast)"
    Close #fileNum
    
    Dim psCommand As String
    psCommand = "powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & scriptPath & """"
    
    WsShell.Run psCommand, 0, False
    Set WsShell = Nothing
    Exit Sub
    
ErrorHandler:
    Debug.Print "ShowToastWin10 Error: " & Err.Description
    Set WsShell = Nothing
End Sub
