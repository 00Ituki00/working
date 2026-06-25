# Windows通知ユーティリティ開発記録

## 概要
Excel VBAからPowerShell経由でWindowsタスクトレイにトースト通知を表示する機能の開発。

## 最終的な動作コード

```vba
Function MakeToastNotification(ByVal msg_title As String, ByVal msg_text As String) As Integer
    Const C_SUCCESS As Integer = 0
    Const C_FAILURE As Integer = 1
    
    Dim cmd As String
    
    Const C_CMD1 = "powershell -Command ""Add-Type -AssemblyName System.Windows.Forms;" & _
        "$toast = New-Object System.Windows.Forms.NotifyIcon;" & _
        "$toast.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon('C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe'); " & _
        "$toast.BalloonTipTitle = '"
    
    Const C_CMD2 = "'; " & _
        "$toast.BalloonTipText = '"
    
    Const C_CMD3 = "'; " & _
        "$toast.Visible = $True; " & _
        "$toast.ShowBalloonTip(0)"""
    
    cmd = C_CMD1 & msg_title & C_CMD2 & msg_text & C_CMD3
    
    On Error Resume Next
    Call VBA.Shell(cmd, vbHide)
    
    If Err.Number <> 0 Then
        MakeToastNotification = C_FAILURE
        Err.Clear
    Else
        MakeToastNotification = C_SUCCESS
    End If
End Function
```

## 問題点と解決方法

### 問題1: PowerShellウィンドウが表示される
**原因:** `WScript.Shell.Run` のウィンドウスタイル指定が効かない場合がある
**解決:** `VBA.Shell(cmd, vbHide)` を使用

### 問題2: パスに無効な文字が含まれるエラー
**原因:** `Environ("TEMP")` で取得したパスに日本語や特殊文字が含まれる場合、8.3形式（短縮名）に変換されるが、その過程で問題が発生
**解決:** パスを使用せず、PowerShellコマンドを直接インラインで実行

### 問題3: ShowBalloonTipのアイコン指定
**原因:** `[System.Windows.Forms.ToolTipIcon]::Information` などの列挙型指定が正しく動作しない場合がある
**解決:** `ExtractAssociatedIcon` で確実にアイコンを取得し、ShowBalloonTip(0)でデフォルト表示

### 問題4: Add-Typeが見つからないエラー
**原因:** PowerShellの実行ポリシーやアセンブリ読み込みの問題
**解決:** インラインコマンドで `-Command` を使用し、単一のセッションで実行

## 重要な教訓

1. **VBAからPowerShellを呼び出す際は、`VBA.Shell` を使用する**
   - `WScript.Shell` より確実にウィンドウを非表示にできる

2. **パス問題を回避するには、インラインコマンドを使用する**
   - 一時ファイルにスクリプトを書き出す方式は、パスに問題があると失敗する
   - 直接コマンドラインに組み込む方が確実

3. **参考サイトの確認**
   - https://enjoy-tech.net/45754/ のコードが最も確実に動作
   - 自分で組み立てるより、動作確認済みのコードを参考にする方が効率的

## 参考URL
- https://enjoy-tech.net/45754/
- https://learn.microsoft.com/ja-jp/windows/apps/design/shell/tiles-and-notifications/toast-notifications-overview
