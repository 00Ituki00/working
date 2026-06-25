' IME読みがな取得ユーティリティ
' Windows IME APIを使用して高精度な読みがなを取得
' 参考: https://learn.microsoft.com/en-us/windows/win32/api/imm/
' 追加: ローマ字変換対応（ヘボン式）

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

' === ローマ字変換テーブル（ヘボン式） ===
Private Function GetRomanTable() As Object
    Dim table As Object
    Set table = CreateObject("Scripting.Dictionary")
    
    ' あ行
    table.Add "あ", "a": table.Add "い", "i": table.Add "う", "u"
    table.Add "え", "e": table.Add "お", "o"
    
    ' か行
    table.Add "か", "ka": table.Add "き", "ki": table.Add "く", "ku"
    table.Add "け", "ke": table.Add "こ", "ko"
    table.Add "が", "ga": table.Add "ぎ", "gi": table.Add "ぐ", "gu"
    table.Add "げ", "ge": table.Add "ご", "go"
    
    ' さ行
    table.Add "さ", "sa": table.Add "し", "shi": table.Add "す", "su"
    table.Add "せ", "se": table.Add "そ", "so"
    table.Add "ざ", "za": table.Add "じ", "ji": table.Add "ず", "zu"
    table.Add "ぜ", "ze": table.Add "ぞ", "zo"
    
    ' た行
    table.Add "た", "ta": table.Add "ち", "chi": table.Add "つ", "tsu"
    table.Add "て", "te": table.Add "と", "to"
    table.Add "だ", "da": table.Add "ぢ", "ji": table.Add "づ", "zu"
    table.Add "で", "de": table.Add "ど", "do"
    
    ' な行
    table.Add "な", "na": table.Add "に", "ni": table.Add "ぬ", "nu"
    table.Add "ね", "ne": table.Add "の", "no"
    
    ' は行
    table.Add "は", "ha": table.Add "ひ", "hi": table.Add "ふ", "fu"
    table.Add "へ", "he": table.Add "ほ", "ho"
    table.Add "ば", "ba": table.Add "び", "bi": table.Add "ぶ", "bu"
    table.Add "べ", "be": table.Add "ぼ", "bo"
    table.Add "ぱ", "pa": table.Add "ぴ", "pi": table.Add "ぷ", "pu"
    table.Add "ぺ", "pe": table.Add "ぽ", "po"
    
    ' ま行
    table.Add "ま", "ma": table.Add "み", "mi": table.Add "む", "mu"
    table.Add "め", "me": table.Add "も", "mo"
    
    ' や行
    table.Add "や", "ya": table.Add "ゆ", "yu": table.Add "よ", "yo"
    
    ' ら行
    table.Add "ら", "ra": table.Add "り", "ri": table.Add "る", "ru"
    table.Add "れ", "re": table.Add "ろ", "ro"
    
    ' わ行
    table.Add "わ", "wa": table.Add "を", "wo": table.Add "ん", "n"
    
    ' 拗音
    table.Add "きゃ", "kya": table.Add "きゅ", "kyu": table.Add "きょ", "kyo"
    table.Add "しゃ", "sha": table.Add "しゅ", "shu": table.Add "しょ", "sho"
    table.Add "ちゃ", "cha": table.Add "ちゅ", "chu": table.Add "ちょ", "cho"
    table.Add "にゃ", "nya": table.Add "にゅ", "nyu": table.Add "にょ", "nyo"
    table.Add "ひゃ", "hya": table.Add "ひゅ", "hyu": table.Add "ひょ", "hyo"
    table.Add "みゃ", "mya": table.Add "みゅ", "myu": table.Add "みょ", "myo"
    table.Add "りゃ", "rya": table.Add "りゅ", "ryu": table.Add "りょ", "ryo"
    table.Add "ぎゃ", "gya": table.Add "ぎゅ", "gyu": table.Add "ぎょ", "gyo"
    table.Add "じゃ", "ja": table.Add "じゅ", "ju": table.Add "じょ", "jo"
    table.Add "びゃ", "bya": table.Add "びゅ", "byu": table.Add "びょ", "byo"
    table.Add "ぴゃ", "pya": table.Add "ぴゅ", "pyu": table.Add "ぴょ", "pyo"
    
    ' 促音・撥音
    table.Add "っ", "xtsu"
    table.Add "ー", "-"
    
    Set GetRomanTable = table
End Function

' === ひらがなをローマ字に変換（ヘボン式） ===
Public Function HiraganaToRoman(ByVal hiragana As String) As String
    Dim table As Object
    Set table = GetRomanTable()
    
    Dim result As String
    result = ""
    
    Dim i As Long
    i = 1
    
    Do While i <= Len(hiragana)
        Dim current As String
        Dim nextChar As String
        Dim combined As String
        
        current = Mid(hiragana, i, 1)
        
        ' 2文字先を確認（拗音・促音など）
        If i < Len(hiragana) Then
            nextChar = Mid(hiragana, i + 1, 1)
            combined = current & nextChar
            
            ' 2文字の組み合わせをチェック
            If table.Exists(combined) Then
                result = result & table(combined)
                i = i + 2
                GoTo ContinueLoop
            End If
        End If
        
        ' 1文字をチェック
        If table.Exists(current) Then
            result = result & table(current)
        Else
            ' 変換できない文字はそのまま
            result = result & current
        End If
        
        i = i + 1
ContinueLoop:
    Loop
    
    HiraganaToRoman = result
End Function

' === ひらがなをローマ字に変換（改良版：促音対応） ===
Public Function HiraganaToRomanEx(ByVal hiragana As String) As String
    Dim table As Object
    Set table = GetRomanTable()
    
    Dim result As String
    result = ""
    
    Dim i As Long
    i = 1
    
    Do While i <= Len(hiragana)
        Dim current As String
        Dim nextChar As String
        Dim combined As String
        
        current = Mid(hiragana, i, 1)
        
        ' 促音「っ」の処理
        If current = "っ" And i < Len(hiragana) Then
            Dim nextConsonant As String
            nextConsonant = Mid(hiragana, i + 1, 1)
            
            ' 次の文字の子音を重ねる
            If table.Exists(nextConsonant) Then
                Dim romanNext As String
                romanNext = table(nextConsonant)
                If Len(romanNext) > 0 Then
                    result = result & Left(romanNext, 1)
                End If
            End If
            i = i + 1
            GoTo ContinueLoop
        End If
        
        ' 2文字先を確認（拗音など）
        If i < Len(hiragana) Then
            nextChar = Mid(hiragana, i + 1, 1)
            combined = current & nextChar
            
            If table.Exists(combined) Then
                result = result & table(combined)
                i = i + 2
                GoTo ContinueLoop
            End If
        End If
        
        ' 1文字をチェック
        If table.Exists(current) Then
            result = result & table(current)
        Else
            result = result & current
        End If
        
        i = i + 1
ContinueLoop:
    Loop
    
    HiraganaToRomanEx = result
End Function

' === 検索用に文字列を正規化（読み→ローマ字） ===
Public Function ToSearchReading(ByVal text As String) As String
    ' 1. 読みがなを取得
    Dim reading As String
    reading = GetReadingIME(text)
    
    ' 2. ローマ字に変換
    Dim roman As String
    roman = HiraganaToRomanEx(reading)
    
    ' 3. 小文字に統一
    ToSearchReading = LCase(roman)
End Function

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
    testCases = Array("山田太郎", "東京", "日本語", "漢字変換", "株式会社")
    
    Dim i As Long
    For i = LBound(testCases) To UBound(testCases)
        Dim text As String
        text = testCases(i)
        
        Dim reading As String
        reading = GetReadingIME(text)
        
        Dim roman As String
        roman = HiraganaToRomanEx(reading)
        
        Dim searchKey As String
        searchKey = ToSearchReading(text)
        
        Debug.Print "テキスト: " & text
        Debug.Print "読み: " & reading
        Debug.Print "ローマ字: " & roman
        Debug.Print "検索キー: " & searchKey
        Debug.Print "---"
    Next i
End Sub

' === 検索関数（部分一致） ===
Public Function SearchByReading(ByVal searchText As String, ByVal targetText As String) As Boolean
    ' 検索文字列を正規化
    Dim searchKey As String
    searchKey = ToSearchReading(searchText)
    
    ' 対象文字列を正規化
    Dim targetKey As String
    targetKey = ToSearchReading(targetText)
    
    ' 部分一致判定
    SearchByReading = (InStr(targetKey, searchKey) > 0)
End Function
