' Windows通知ユーティリティ v4
' PowerShell経由でWindowsトレイ通知（バルーン通知）を表示
' 修正履歴:
'   v2: スクリプトファイル方式に変更
'   v3: PowerShellウィンドウが表示される問題を修正
'   v4: Add-Typeの問題を修正、確実な通知表示を実現

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
    Print #fileNum, "Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop"
    Print #fileNum, "Add-Type -AssemblyName System.Drawing -ErrorAction Stop"
    Print #fileNum, "$notify = New-Object System.Windows.Forms.NotifyIcon"
    Print #fileNum, "$notify.Icon = [System.Drawing.SystemIcons]::" & psIcon
    Print #fileNum, "$notify.Visible = $true"
    Print #fileNum, "$notify.Text = '" & EscapeForPowerShell(Title) & "'"
    Print #fileNum, "$notify.ShowBalloonTip(" & Duration & ", '" & EscapeForPowerShell(Title) & "', '" & EscapeForPowerShell(Message) & "', [System.Windows.Forms.ToolTipIcon]::" & psIcon & ")"
    Print #fileNum, "Start-Sleep -Seconds " & Duration + 2
    Print #fileNum, "$notify.Visible = $false"
    Print #fileNum, "$notify.Dispose()"
    Close #fileNum
    
    ' Shell.Application を使用してウィンドウを非表示
    Dim shellApp As Object
    Set shellApp = CreateObject("Shell.Application")
    shellApp.ShellExecute "powershell.exe", "-ExecutionPolicy Bypass -WindowStyle Hidden -File """ & scriptPath & """"", "", "open", 0
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

' === 診断テスト1: PowerShell連携確認（修正版） ===
Public Sub TestPowerShell()
    ' PowerShellが機能するか簡易テスト
    ' 注意: このテストはPowerShellウィンドウが表示されます
    Dim WsShell As Object
    Set WsShell = CreateObject("WScript.Shell")
    
    ' -Command ではなく -File を使用して確実に実行
    Dim scriptPath As String
    scriptPath = Environ("TEMP") & "\test_ps.ps1"
    
    Dim fileNum As Integer
    fileNum = FreeFile
    Open scriptPath For Output As #fileNum
    Print #fileNum, "Add-Type -AssemblyName System.Windows.Forms"
    Print #fileNum, "[System.Windows.Forms.MessageBox]::Show('PowerShell連携テスト')"
    Close #fileNum
    
    WsShell.Run "powershell.exe -ExecutionPolicy Bypass -File """ & scriptPath & """"", 1, True
    Set WsShell = Nothing
End Sub

' === 診断テスト2: 基本的なバルーン通知テスト（修正版） ===
Public Sub TestBasicBalloon()
    ' 最も基本的な方法でバルーン通知をテスト
    Dim scriptPath As String
    scriptPath = Environ("TEMP") & "\test_balloon.ps1"
    
    Dim fileNum As Integer
    fileNum = FreeFile
    Open scriptPath For Output As #fileNum
    Print #fileNum, "Add-Type -AssemblyName System.Windows.Forms"
    Print #fileNum, "Add-Type -AssemblyName System.Drawing"
    Print #fileNum, "$n = New-Object System.Windows.Forms.NotifyIcon"
    Print #fileNum, "$n.Icon = [System.Drawing.SystemIcons]::Information"
    Print #fileNum, "$n.Visible = $true"
    Print #fileNum, "$n.Text = 'テスト'"
    Print #fileNum, "$n.ShowBalloonTip(10000, 'テスト', 'これはテスト通知です', [System.Windows.Forms.ToolTipIcon]::Information)"
    Print #fileNum, "Start-Sleep -Seconds 12"
    Print #fileNum, "$n.Visible = $false"
    Print #fileNum, "$n.Dispose()"
    Close #fileNum
    
    ' 同期的に実行して結果を確認
    Dim WsShell As Object
    Set WsShell = CreateObject("WScript.Shell")
    
    Dim ret As Long
    ret = WsShell.Run("powershell.exe -ExecutionPolicy Bypass -File """ & scriptPath & """"", 0, True)
    Debug.Print "Return code: " & ret
    
    Set WsShell = Nothing
End Sub

' === 診断テスト3: Windows通知設定確認 ===
Public Sub TestNotificationSettings()
    ' Windowsの通知設定を確認
    Dim WsShell As Object
    Set WsShell = CreateObject("WScript.Shell")
    
    Dim scriptPath As String
    scriptPath = Environ("TEMP") & "\test_settings.ps1"
    
    Dim fileNum As Integer
    fileNum = FreeFile
    Open scriptPath For Output As #fileNum
    Print #fileNum, "$regPath = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\PushNotifications'"
    Print #fileNum, "$value = Get-ItemProperty -Path $regPath -Name 'ToastEnabled' -ErrorAction SilentlyContinue"
    Print #fileNum, "if ($value) {"
    Print #fileNum, "    Write-Host 'ToastEnabled: ' $value.ToastEnabled"
    Print #fileNum, "} else {"
    Print #fileNum, "    Write-Host 'ToastEnabled: Not found (default: enabled)'"
    Print #fileNum, "}"
    Print #fileNum, ""
    Print #fileNum, "$regPath2 = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Notifications\Settings'"
    Print #fileNum, "$apps = Get-ChildItem -Path $regPath2 -ErrorAction SilentlyContinue"
    Print #fileNum, "Write-Host 'Notification apps count: ' $apps.Count"
    Close #fileNum
    
    ' 結果を表示
    WsShell.Run "powershell.exe -ExecutionPolicy Bypass -File """ & scriptPath & """"", 1, True
    Set WsShell = Nothing
End Sub

' === Windows 10/11 トースト通知 ===
Public Sub ShowToastWin10(ByVal Title As String, ByVal Message As String)
    ' Windows 10/11 のモダントースト通知
    On Error GoTo ErrorHandler
    
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
    
    Dim shellApp As Object
    Set shellApp = CreateObject("Shell.Application")
    shellApp.ShellExecute "powershell.exe", "-ExecutionPolicy Bypass -WindowStyle Hidden -File """ & scriptPath & """"", "", "open", 0
    Set shellApp = Nothing
    Exit Sub
    
ErrorHandler:
    Debug.Print "ShowToastWin10 Error: " & Err.Description
End Sub

' === 最終手段: MsgBoxを使用した通知 ===
Public Sub ShowMessageBox(ByVal Title As String, ByVal Message As String)
    ' 通知が表示されない場合のフォールバック
    MsgBox Message, vbInformation, Title
End Sub
