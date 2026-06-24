Public Sub 空白削除()

    Dim rng As Range
    Dim arr As Variant
    Dim fmtArr() As Variant
    Dim r As Long, c As Long
    Dim v As Variant

    If TypeName(Selection) <> "Range" Then Exit Sub
    Set rng = Selection
    arr = rng.Value2

    ' 列ごとに元の書式を保存
    ReDim fmtArr(1 To rng.Columns.Count)
    For c = 1 To rng.Columns.Count
        fmtArr(c) = rng.Columns(c).NumberFormat
    Next c

    ' 配列内で文字列の空白を削除（数値は触らない）
    For r = 1 To UBound(arr, 1)
        For c = 1 To UBound(arr, 2)
            v = arr(r, c)
            
            If IsEmpty(v) Or IsError(v) Then
                arr(r, c) = ""
            ElseIf VarType(v) = vbString Then
                Select Case v
                    Case "(空白)", "-", " ", "　", ""
                        arr(r, c) = ""
                End Select
            End If
            
        Next c
    Next r

    ' 一括書き込み：文字列として書き込んで桁落ち防止
    rng.NumberFormat = "@"
    rng.Value2 = arr
    
    ' 列ごとに元の書式に戻す
    For c = 1 To rng.Columns.Count
        rng.Columns(c).NumberFormat = fmtArr(c)
    Next c

End Sub
