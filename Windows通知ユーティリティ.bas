' Windows通知ユーティリティ v7
' PowerShell経由でWindowsタスクトレイにトースト通知を表示
' 参考: https://enjoy-tech.net/45754/
' 修正履歴:
'   v7: 記事のコードを参考に全面書き換え

' === 通知表示関数 ===
' 引数:
'   msg_title - 通知タイトル
'   msg_text  - 通知本文
' 戻り値:
'   0 - 正常終了
'   1 - 異常終了
Public Function MakeToastNotification(ByVal msg_title As String, ByVal msg_text As String) As Integer
    
    ' 成功コードと失敗コードを定数として設定
    Const C_SUCCESS As Integer = 0
    Const C_FAILURE As Integer = 1
    
    Dim cmd As String
    
    ' PowerShellコマンドの一部を定数として設定
    Const C_CMD1 = "powershell -Command ""Add-Type -AssemblyName System.Windows.Forms;" & _
        "$toast = New-Object System.Windows.Forms.NotifyIcon;" & _
        "$toast.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon('C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe'); " & _
        "$toast.BalloonTipTitle = '"
    
    Const C_CMD2 = "'; " & _
        "$toast.BalloonTipText = '"
    
    Const C_CMD3 = "'; " & _
        "$toast.Visible = $True; " & _
        "$toast.ShowBalloonTip(0)"""
    
    ' PowerShellコマンドを構築
    cmd = C_CMD1 & EscapeForShell(msg_title) & C_CMD2 & EscapeForShell(msg_text) & C_CMD3
    
    ' PowerShellスクリプトを実行
    On Error Resume Next
    Call VBA.Shell(cmd, vbHide)
    
    If Err.Number <> 0 Then
        MakeToastNotification = C_FAILURE
        Err.Clear
    Else
        MakeToastNotification = C_SUCCESS
    End If
    
End Function

' === シェル用文字列エスケープ ===
Private Function EscapeForShell(ByVal str As String) As String
    ' シングルクォートをエスケープ
    str = Replace(str, "'", "''")
    EscapeForShell = str
End Function

' === 簡易通知関数 ===
Public Sub Toast(ByVal Title As String, ByVal Message As String)
    MakeToastNotification Title, Message
End Sub

Public Sub ToastComplete(Optional ByVal TaskName As String = "処理")
    MakeToastNotification TaskName & "完了", "正常に完了しました"
End Sub

Public Sub ToastError(ByVal ErrorMessage As String)
    MakeToastNotification "エラー", ErrorMessage
End Sub

Public Sub ToastWarning(ByVal WarningMessage As String)
    MakeToastNotification "警告", WarningMessage
End Sub

' === 最終手段：MsgBox ===
Public Sub ShowMessageBox(ByVal Title As String, ByVal Message As String)
    MsgBox Message, vbInformation, Title
End Sub
