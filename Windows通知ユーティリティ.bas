' Windows通知ユーティリティ v5
' PowerShell経由でWindowsトレイ通知（バルーン通知）を表示
' 修正履歴:
'   v2: スクリプトファイル方式に変更
'   v3: PowerShellウィンドウが表示される問題を修正
'   v4: Add-Typeの問題を修正
'   v5: ShowBalloonTipの第4引数を数値に修正、パス指定を修正

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
    
    ' アイコン種別を数値に変換（ShowBalloonTipの第4引数は0-3）
    Dim iconValue As Integer
    Select Case LCase(IconType)
        Case "info", "information"
            iconValue = 1  ' ToolTipIcon.Info
        Case "warning", "warn"
            iconValue = 2  ' ToolTipIcon.Warning
        Case "error", "err"
            iconValue = 3  ' ToolTipIcon.Error
        Case Else
            iconValue = 0  ' ToolTipIcon.None
    End Select
    
    ' スクリプトファイルに出力（長いパス名対策：\Environ("TEMP")を使用）
    Dim scriptPath As String
    scriptPath = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%TEMP%") & "\toast_notification.ps1"
    
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open scriptPath For Output As #fileNum
    Print #fileNum, "Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop"
    Print #fileNum, "Add-Type -AssemblyName System.Drawing -ErrorAction Stop"
    Print #fileNum, "$notify = New-Object System.Windows.Forms.NotifyIcon"
    Print #fileNum, "$notify.Icon = [System.Drawing.SystemIcons]::Information"
    Print #fileNum, "$notify.Visible = $true"
    Print #fileNum, "$notify.Text = '" & EscapeForPowerShell(Title) & "'"
    Print #fileNum, "$notify.ShowBalloonTip(" & Duration & ", '" & EscapeForPowerShell(Title) & "', '" & EscapeForPowerShell(Message) & "', " & iconValue & ")"
    Print #fileNum, "Start-Sleep -Seconds " & Duration + 2
    Print #fileNum, "$notify.Visible = $false"
    Print #fileNum, "$notify.Dispose()"
    Close #fileNum
    
    ' PowerShell実行（長いパス名対策）
    Dim WsShell As Object
    Set WsShell = CreateObject("WScript.Shell")
    
    ' パスに日本語が含まれる場合の対策：短縮パスを取得
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim shortPath As String
    shortPath = fso.GetFile(scriptPath).ShortPath
    
    WsShell.Run "powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & shortPath & """"", 0, False
    Set WsShell = Nothing
    Set fso = Nothing
    
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
    Dim WsShell As Object
    Set WsShell = CreateObject("WScript.Shell")
    
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

' === 診断テスト2: 基本的なバルーン通知テスト ===
Public Sub TestBasicBalloon()
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
    Print #fileNum, "$n.ShowBalloonTip(10000, 'テスト', 'これはテスト通知です', 1)"
    Print #fileNum, "Start-Sleep -Seconds 12"
    Print #fileNum, "$n.Visible = $false"
    Print #fileNum, "$n.Dispose()"
    Close #fileNum
    
    Dim WsShell As Object
    Set WsShell = CreateObject("WScript.Shell")
    WsShell.Run "powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & scriptPath & """"", 0, True
    Set WsShell = Nothing
End Sub

' === 診断テスト3: Windows通知設定確認 ===
Public Sub TestNotificationSettings()
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
    Close #fileNum
    
    WsShell.Run "powershell.exe -ExecutionPolicy Bypass -File """ & scriptPath & """"", 1, True
    Set WsShell = Nothing
End Sub

' === 最終手段: MsgBoxを使用した通知 ===
Public Sub ShowMessageBox(ByVal Title As String, ByVal Message As String)
    MsgBox Message, vbInformation, Title
End Sub
