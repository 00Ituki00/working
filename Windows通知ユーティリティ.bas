' Windows通知ユーティリティ v3
' PowerShell経由でWindowsトレイ通知（バルーン通知）を表示
' 修正履歴:
'   v2: スクリプトファイル方式に変更
'   v3: PowerShellウィンドウが表示される問題を修正
'       - WScript.Shell.Run から Shell.Application へ変更
'       - 同期的実行でエラーをキャッチ
'       - フォールバック方法を追加

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
    
    ' スクリプトファイルに出力
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
    
    ' 方法1: Shell.Application を使用（ウィンドウを表示しない）
    On Error Resume Next
    Dim shellApp As Object
    Set shellApp = CreateObject("Shell.Application")
    shellApp.ShellExecute "powershell.exe", "-ExecutionPolicy Bypass -WindowStyle Hidden -File """ & scriptPath & """"", "", "open", 0
    
    If Err.Number <> 0 Then
        ' 方法2: WScript.Shell を使用（フォールバック）
        Err.Clear
        Dim WsShell As Object
        Set WsShell = CreateObject("WScript.Shell")
        ' 0 = 非表示ウィンドウ, False = 非同期実行
        WsShell.Run "powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & scriptPath & """"", 0, False
        Set WsShell = Nothing
    End If
    
    Set shellApp = Nothing
    Exit Sub
    
ErrorHandler:
    Debug.Print "ShowToast Error: " & Err.Description
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

' === 診断テスト1: PowerShell連携確認 ===
Public Sub TestPowerShell()
    ' PowerShellが機能するか簡易テスト
    Dim WsShell As Object
    Set WsShell = CreateObject("WScript.Shell")
    
    ' メッセージボックスでPowerShell連携確認
    Dim testCmd As String
    testCmd = "powershell.exe -Command ""[System.Windows.Forms.MessageBox]::Show('PowerShell連携テスト')"""
    
    WsShell.Run testCmd, 1, True
    Set WsShell = Nothing
End Sub

' === 診断テスト2: 通知設定確認 ===
Public Sub TestNotificationSettings()
    ' Windowsの通知設定を確認するPowerShellスクリプトを実行
    Dim WsShell As Object
    Set WsShell = CreateObject("WScript.Shell")
    
    Dim psCommand As String
    psCommand = "powershell.exe -Command """ & _
        "Get-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\PushNotifications' -Name 'ToastEnabled' -ErrorAction SilentlyContinue | Select-Object ToastEnabled" & _
        """"
    
    WsShell.Run psCommand, 1, True
    Set WsShell = Nothing
End Sub

' === 診断テスト3: 基本的なバルーン通知テスト ===
Public Sub TestBasicBalloon()
    ' 最も基本的な方法でバルーン通知をテスト
    Dim WsShell As Object
    Set WsShell = CreateObject("WScript.Shell")
    
    ' 一時的に表示時間を長くしてテスト
    Dim psCommand As String
    psCommand = "powershell.exe -WindowStyle Hidden -Command """ & _
        "Add-Type -AssemblyName System.Windows.Forms; " & _
        "$n = New-Object System.Windows.Forms.NotifyIcon; " & _
        "$n.Icon = [System.Drawing.SystemIcons]::Information; " & _
        "$n.Visible = $true; " & _
        "$n.ShowBalloonTip(10000, 'テスト', 'これはテスト通知です', [System.Windows.Forms.ToolTipIcon]::Information); " & _
        "Start-Sleep -Seconds 12; " & _
        "$n.Dispose()" & _
        """"
    
    ' デバッグ出力
    Debug.Print "Command: " & psCommand
    
    ' 実行（同期的にして結果を確認）
    Dim ret As Long
    ret = WsShell.Run(psCommand, 0, True)
    Debug.Print "Return code: " & ret
    
    Set WsShell = Nothing
End Sub

' === Windows 10/11 トースト通知 ===
Public Sub ShowToastWin10(ByVal Title As String, ByVal Message As String)
    ' Windows 10/11 のモダントースト通知
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

' === 最終手段: MsgBoxを使用した通知 ===
Public Sub ShowMessageBox(ByVal Title As String, ByVal Message As String)
    ' 通知が表示されない場合のフォールバック
    MsgBox Message, vbInformation, Title
End Sub
