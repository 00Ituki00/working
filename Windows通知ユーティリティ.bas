' Windows通知ユーティリティ v6
' PowerShell経由でWindowsトレイ通知（バルーン通知）を表示
' 修正履歴:
'   v2: スクリプトファイル方式に変更
'   v3: PowerShellウィンドウが表示される問題を修正
'   v4: Add-Typeの問題を修正
'   v5: ShowBalloonTipの第4引数を数値に修正
'   v6: パス取得方法を統一（WScript.Shell.ExpandEnvironmentStrings）

' === 共通：一時パス取得関数 ===
Private Function GetTempPath() As String
    ' WScript.Shellを使用して確実に一時パスを取得
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    GetTempPath = wsh.ExpandEnvironmentStrings("%TEMP%")
    Set wsh = Nothing
End Function

' === 共通：短縮パス取得関数 ===
Private Function GetShortPath(ByVal longPath As String) As String
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' ファイルが存在する場合は短縮パスを取得
    If Dir(longPath) <> "" Then
        GetShortPath = fso.GetFile(longPath).ShortPath
    Else
        ' ファイルが存在しない場合は元のパスを返す
        GetShortPath = longPath
    End If
    
    Set fso = Nothing
End Function

' === 通知表示関数 ===
Public Sub ShowToast(ByVal Title As String, ByVal Message As String, _
                     Optional Duration As Long = 10, _
                     Optional IconType As String = "Info")
    
    On Error GoTo ErrorHandler
    
    ' アイコン種別を数値に変換（ShowBalloonTipの第4引数は0-3）
    Dim iconValue As Integer
    Select Case LCase(IconType)
        Case "info", "information"
            iconValue = 1
        Case "warning", "warn"
            iconValue = 2
        Case "error", "err"
            iconValue = 3
        Case Else
            iconValue = 0
    End Select
    
    ' スクリプトファイルパス
    Dim scriptPath As String
    scriptPath = GetTempPath() & "\toast_notification.ps1"
    
    ' スクリプトファイル作成
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
    
    ' 短縮パスを取得してPowerShell実行
    Dim shortPath As String
    shortPath = GetShortPath(scriptPath)
    
    Dim WsShell As Object
    Set WsShell = CreateObject("WScript.Shell")
    WsShell.Run "powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & shortPath & """"", 0, False
    Set WsShell = Nothing
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "ShowToast Error: " & Err.Description
End Sub

' === PowerShell用文字列エスケープ ===
Private Function EscapeForPowerShell(ByVal str As String) As String
    str = Replace(str, "'", "''")
    str = Replace(str, "`", "``")
    EscapeForPowerShell = str
End Function

' === 簡易通知関数 ===
Public Sub Toast(ByVal Title As String, ByVal Message As String)
    ShowToast Title, Message, 10, "Info"
End Sub

Public Sub ToastComplete(Optional ByVal TaskName As String = "処理")
    ShowToast TaskName & "完了", "正常に完了しました", 5, "Info"
End Sub

Public Sub ToastError(ByVal ErrorMessage As String)
    ShowToast "エラー", ErrorMessage, 15, "Error"
End Sub

Public Sub ToastWarning(ByVal WarningMessage As String)
    ShowToast "警告", WarningMessage, 10, "Warning"
End Sub

' === 診断テスト：基本的なバルーン通知 ===
Public Sub TestBasicBalloon()
    Dim scriptPath As String
    scriptPath = GetTempPath() & "\test_balloon.ps1"
    
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
    
    Dim shortPath As String
    shortPath = GetShortPath(scriptPath)
    
    Dim WsShell As Object
    Set WsShell = CreateObject("WScript.Shell")
    WsShell.Run "powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & shortPath & """"", 0, True
    Set WsShell = Nothing
End Sub

' === 最終手段：MsgBox ===
Public Sub ShowMessageBox(ByVal Title As String, ByVal Message As String)
    MsgBox Message, vbInformation, Title
End Sub
