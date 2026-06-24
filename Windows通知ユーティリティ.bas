' Windows通知ユーティリティ
' PowerShell経由でWindowsトレイ通知（バルーン通知）を表示

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
    
    ' PowerShellコマンド構築
    Dim psCommand As String
    psCommand = "powershell.exe -WindowStyle Hidden -Command " & _
        """Add-Type -AssemblyName System.Windows.Forms; " & _
        "$notify = New-Object System.Windows.Forms.NotifyIcon; " & _
        "$notify.Icon = [System.Drawing.SystemIcons]::" & psIcon & "; " & _
        "$notify.Visible = $true; " & _
        "$notify.ShowBalloonTip(" & Duration & ",'" & _
        EscapeForPowerShell(Title) & "','" & _
        EscapeForPowerShell(Message) & "'," & _
        "[System.Windows.Forms.ToolTipIcon]::" & psIcon & "); " & _
        "Start-Sleep -Seconds " & Duration + 1 & "; " & _
        "$notify.Dispose()"""
    
    ' 非同期実行（VBA処理をブロックしない）
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
