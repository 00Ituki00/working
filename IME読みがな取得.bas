' IME読みがな取得ユーティリティ
' Windows IME APIを使用して高精度な読みがなを取得
' 参考: https://learn.microsoft.com/en-us/windows/win32/api/imm/

Option Explicit

' === Windows API宣言 ===
Private Declare PtrSafe Function ImmGetContext Lib "imm32.dll" (ByVal hwnd As LongPtr) As LongPtr
Private Declare PtrSafe Function ImmReleaseContext Lib "imm32.dll" (ByVal hwnd As LongPtr, ByVal hIMC As LongPtr) As Long
Private Declare PtrSafe Function ImmGetConversionList Lib "imm32.dll" Alias "ImmGetConversionListW" (ByVal hKL As LongPtr, ByVal hIMC As LongPtr, ByVal lpszSrc As LongPtr, ByRef lpDst As Any, ByVal dwBufLen As Long, ByVal uFlag As Long) As Long
Private Declare PtrSafe Function GetKeyboardLayout Lib "user32.dll" (ByVal dwLayout As Long) As LongPtr

' === 定数 ===
Private Const GCL_CONVERSION As Long = 1
Private Const IME_CMODE_NATIVE As Long = 1

' === 構造体 ===
Private Type CANDIDATELIST
    dwSize As Long
    dwStyle As Long
    dwCount As Long
    dwSelection As Long
    dwPageStart As Long
    dwPageSize As Long
    dwOffset(0 To 9) As Long  ' 最大10候補
End Type

' === メイン関数：読みがな取得 ===
Public Function GetReadingIME(ByVal text As String) As String
    On Error GoTo ErrorHandler
    
    Dim hWnd As LongPtr
    Dim hIMC As LongPtr
    Dim hKL As LongPtr
    Dim result As String
    
    ' Excelのウィンドウハンドルを取得
    hWnd = Application.hWnd
    
    ' IMEコンテキストを取得
    hIMC = ImmGetContext(hWnd)
    
    If hIMC = 0 Then
        ' IMEコンテキスト取得失敗時はフォールバック
        GetReadingIME = GetReadingFallback(text)
        Exit Function
    End If
    
    ' キーボードレイアウトを取得
    hKL = GetKeyboardLayout(0)
    
    ' 変換候補リストを取得
    Dim candidateList As CANDIDATELIST
    Dim bufLen As Long
    Dim srcStr As String
    
    ' Unicode文字列を準備
    srcStr = text & vbNullChar
    
    ' 必要なバッファサイズを取得
    bufLen = ImmGetConversionList(hKL, hIMC, StrPtr(srcStr), ByVal vbNullString, 0, GCL_CONVERSION)
    
    If bufLen > 0 Then
        ' バッファを確保して再取得
        Dim buffer() As Byte
        ReDim buffer(0 To bufLen - 1)
        
        If ImmGetConversionList(hKL, hIMC, StrPtr(srcStr), buffer(0), bufLen, GCL_CONVERSION) > 0 Then
            ' 結果を解析
            result = ParseCandidateList(buffer)
        End If
    End If
    
    ' IMEコンテキストを解放
    ImmReleaseContext hWnd, hIMC
    
    ' 結果が空の場合はフォールバック
    If result = "" Then
        result = GetReadingFallback(text)
    End If
    
    GetReadingIME = result
    Exit Function
    
ErrorHandler:
    Debug.Print "GetReadingIME Error: " & Err.Description
    GetReadingIME = GetReadingFallback(text)
End Function

' === 候補リスト解析 ===
Private Function ParseCandidateList(buffer() As Byte) As String
    On Error Resume Next
    
    Dim dwCount As Long
    Dim dwSelection As Long
    Dim offset As Long
    Dim result As String
    
    ' 候補数を取得
    CopyMemory dwCount, buffer(8), 4
    CopyMemory dwSelection, buffer(12), 4
    
    If dwCount > 0 Then
        ' 最初の候補のオフセットを取得
        CopyMemory offset, buffer(24), 4
        
        ' 文字列を取得（Unicode）
        If offset > 0 And offset < UBound(buffer) Then
            Dim strLen As Long
            strLen = 0
            
            ' 文字列長を計算
            Do While offset + strLen * 2 < UBound(buffer) - 1
                If buffer(offset + strLen * 2) = 0 And buffer(offset + strLen * 2 + 1) = 0 Then
                    Exit Do
                End If
                strLen = strLen + 1
            Loop
            
            ' Unicode文字列を変換
            If strLen > 0 Then
                result = MidB(StrConv(buffer, vbUnicode), offset \ 2 + 1, strLen)
            End If
        End If
    End If
    
    ParseCandidateList = result
End Function

' === フォールバック：GetPhonetic使用 ===
Private Function GetReadingFallback(ByVal text As String) As String
    On Error Resume Next
    
    Dim result As String
    result = Application.GetPhonetic(text)
    
    ' GetPhoneticが失敗した場合は元のテキストを返す
    If result = "" Then
        result = text
    End If
    
    ' 半角カタカナをひらがなに変換
    result = StrConv(result, vbHiragana)
    
    GetReadingFallback = result
End Function

' === 簡易版：セルを使用した読み取得 ===
Public Function GetReadingViaCell(ByVal text As String) As String
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim tempCell As Range
    Dim result As String
    
    ' 一時シートを作成
    Set ws = ThisWorkbook.Worksheets.Add(Visible:=xlSheetVeryHidden)
    Set tempCell = ws.Range("A1")
    
    ' セルにテキストを設定
    tempCell.Value = text
    
    ' ふりがなを設定
    tempCell.SetPhonetic
    
    ' ふりがなを取得
    If tempCell.HasPhonetic Then
        result = tempCell.Phonetics(1).Text
    Else
        result = Application.GetPhonetic(text)
    End If
    
    ' ひらがなに変換
    result = StrConv(result, vbHiragana)
    
    ' 一時シートを削除
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True
    
    GetReadingViaCell = result
    Exit Function
    
ErrorHandler:
    Debug.Print "GetReadingViaCell Error: " & Err.Description
    GetReadingViaCell = GetReadingFallback(text)
End Function

' === テスト関数 ===
Public Sub TestReading()
    Dim testCases As Variant
    testCases = Array("山田太郎", "東京", "日本語", "漢字変換")
    
    Dim i As Long
    For i = LBound(testCases) To UBound(testCases)
        Dim text As String
        text = testCases(i)
        
        Debug.Print "テキスト: " & text
        Debug.Print "IME方式: " & GetReadingIME(text)
        Debug.Print "セル方式: " & GetReadingViaCell(text)
        Debug.Print "フォールバック: " & GetReadingFallback(text)
        Debug.Print "---"
    Next i
End Sub
