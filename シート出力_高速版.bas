' シート出力高速化版
' ピボットテーブル・テーブルを事前に値化し、配列一括処理で高速化
' 書式（値書式、セル色、罫線）を正確に再現

Global maked As New Collection

' === メイン関数：高速版 ===
Public Function 切り出し高速(FromBook As Workbook, fromsheet As Worksheet, FromRange As Range, ToBook As Workbook, ToSheet As Worksheet, Optional selectiononly = False, Optional fitpage = False) As Workbook
    Dim pt As PivotTable
    Dim lo As ListObject
    Dim sh As Shape
    Dim chartObj As ChartObject
    Dim dataArr As Variant
    Dim fmtArr() As Variant
    Dim colorArr() As Variant
    Dim borderArr() As String
    Dim r As Long, c As Long
    Dim firstRow As Long, lastRow As Long
    Dim tempPath As String
    
    STVBA
    tempPath = Environ("TEMP") & "\SheetExport_"
    
    ' テーマ適用
    ToBook.ApplyTheme "C:\Users\h_ikegami\AppData\Roaming\Microsoft\Templates\Document Themes\default.thmx"
    DoEvents
    
    ' === Step 1: ピボットテーブルを値化 ===
    For Each pt In fromsheet.PivotTables
        If Not Intersect(pt.TableRange2, FromRange) Is Nothing Then
            On Error Resume Next
            ' ピボットテーブルを値化（データ本体のみ）
            Dim ptDataRange As Range
            If pt.ShowValuesRow Then
                Set ptDataRange = pt.TableRange1.Offset(1, 0).Resize(pt.TableRange1.Rows.Count - 1)
            Else
                Set ptDataRange = pt.TableRange1
            End If
            ' 値を上書き（ピボット機能を失わせる）
            ptDataRange.Value = ptDataRange.Value
            ' ページフィールドも値化
            If pt.PageFields.Count > 0 Then
                pt.PageRange.Value = pt.PageRange.Value
            End If
            On Error GoTo 0
        End If
    Next pt
    DoEvents
    
    ' === Step 2: テーブル（リストオブジェクト）を値化 ===
    For Each lo In fromsheet.ListObjects
        If Not Intersect(lo.Range, FromRange) Is Nothing Then
            On Error Resume Next
            lo.Range.Value = lo.Range.Value
            On Error GoTo 0
        End If
    Next lo
    DoEvents
    
    ' === Step 3: 書式情報を事前収集 ===
    firstRow = FromRange.Row
    lastRow = FromRange.Row + FromRange.Rows.Count - 1
    
    ' 列書式、セル色、罫線を保存
    ReDim fmtArr(1 To FromRange.Rows.Count, 1 To FromRange.Columns.Count)
    ReDim colorArr(1 To FromRange.Rows.Count, 1 To FromRange.Columns.Count)
    ReDim borderArr(1 To FromRange.Rows.Count, 1 To FromRange.Columns.Count)
    
    For r = 1 To FromRange.Rows.Count
        For c = 1 To FromRange.Columns.Count
            Dim srcCell As Range
            Set srcCell = fromsheet.Cells(firstRow + r - 1, FromRange.Column + c - 1)
            fmtArr(r, c) = srcCell.NumberFormat
            colorArr(r, c) = srcCell.Interior.Color
            ' 罫線情報を文字列として保存
            borderArr(r, c) = GetBorderKey(srcCell)
        Next c
    Next r
    DoEvents
    
    ' === Step 4: セルデータを配列で一括コピー ===
    dataArr = FromRange.Value2
    
    ' 配列内で空白処理
    For r = 1 To UBound(dataArr, 1)
        For c = 1 To UBound(dataArr, 2)
            If IsEmpty(dataArr(r, c)) Or IsError(dataArr(r, c)) Then
                dataArr(r, c) = ""
            End If
        Next c
    Next r
    DoEvents
    
    ' === Step 5: 一括書き込み ===
    Dim ToRange As Range
    Set ToRange = ToSheet.Range(ToSheet.Cells(FromRange.Row, FromRange.Column), _
                                ToSheet.Cells(FromRange.Row + FromRange.Rows.Count - 1, _
                                              FromRange.Column + FromRange.Columns.Count - 1))
    
    ' 文字列として書き込んで桁落ち防止
    ToRange.NumberFormat = "@"
    ToRange.Value2 = dataArr
    DoEvents
    
    ' === Step 6: 書式を個別適用 ===
    For r = 1 To FromRange.Rows.Count
        For c = 1 To FromRange.Columns.Count
            Dim dstCell As Range
            Set dstCell = ToSheet.Cells(FromRange.Row + r - 1, FromRange.Column + c - 1)
            ' 値書式
            dstCell.NumberFormat = fmtArr(r, c)
            ' セル色
            If colorArr(r, c) <> 16777215 Then ' 白色以外
                dstCell.Interior.Color = colorArr(r, c)
            End If
            ' 罫線
            ApplyBorderKey dstCell, borderArr(r, c)
        Next c
    Next r
    DoEvents
    
    ' === Step 7: 条件書式を正規化 ===
    書式を正規化 ToRange
    DoEvents
    
    ' === Step 8: 図形を画像としてコピー ===
    Dim imgIndex As Long: imgIndex = 0
    
    ' グラフ（ChartObjects）を画像化
    For Each chartObj In fromsheet.ChartObjects
        If Not Intersect(Range(chartObj.TopLeftCell, chartObj.BottomRightCell), FromRange) Is Nothing Then
            On Error Resume Next
            imgIndex = imgIndex + 1
            Dim chartFile As String
            chartFile = tempPath & "chart_" & imgIndex & ".png"
            chartObj.Chart.Export Filename:=chartFile, FilterName:="PNG"
            
            ' 出力先に画像として挿入
            ToSheet.Activate
            ToSheet.Pictures.Insert(chartFile).Select
            With ToSheet.Shapes(ToSheet.Shapes.Count)
                .Top = ToSheet.Range(chartObj.TopLeftCell.Address).Top
                .Left = ToSheet.Range(chartObj.TopLeftCell.Address).Left
                .Width = chartObj.Width
                .Height = chartObj.Height
            End With
            
            ' 一時ファイル削除
            Kill chartFile
            On Error GoTo 0
        End If
    Next chartObj
    DoEvents
    
    ' その他の図形
    For Each sh In fromsheet.Shapes
        If Not Intersect(Range(sh.TopLeftCell, sh.BottomRightCell), FromRange) Is Nothing Then
            If sh.Type <> msoChart And sh.Type <> 17 Then
                On Error Resume Next
                imgIndex = imgIndex + 1
                Dim shapeFile As String
                shapeFile = tempPath & "shape_" & imgIndex & ".png"
                
                sh.Copy
                ToSheet.Paste
                With ToSheet.Shapes(ToSheet.Shapes.Count)
                    .Top = ToSheet.Range(sh.TopLeftCell.Address).Top
                    .Left = ToSheet.Range(sh.TopLeftCell.Address).Left
                End With
                On Error GoTo 0
            End If
        End If
    Next sh
    DoEvents
    
    ' === Step 9: ウィンドウ設定のコピー ===
    ToSheet.Activate
    ToSheet.Cells(1).Select
    ToSheet.Parent.Windows(1).Zoom = fromsheet.Parent.Windows(1).Zoom
    
    If fromsheet.Parent.Windows(1).FreezePanes Then
        ToSheet.Parent.Windows(1).SplitVertical = fromsheet.Parent.Windows(1).SplitVertical
        ToSheet.Parent.Windows(1).SplitHorizontal = fromsheet.Parent.Windows(1).SplitHorizontal
        ToSheet.Application.ActiveWindow.FreezePanes = True
    End If
    DoEvents
    
    ' === Step 10: 名前定義をコピー ===
    On Error Resume Next
    Dim n As Name, nr As Range, ton As Name
    For Each n In FromBook.Names
        If n.Visible Then
            Set nr = Nothing
            Set nr = n.RefersToRange
            If Not nr Is Nothing Then
                If nr.Worksheet Is fromsheet Then
                    If selectiononly = False Or Not Intersect(nr, FromRange) Is Nothing Then
                        Set ton = ToBook.Names.Add(n.Name, n.RefersTo)
                        ton.Comment = n.Comment
                    End If
                End If
            End If
        End If
    Next n
    On Error GoTo 0
    DoEvents
    
    ' === Step 11: 選択範囲外の削除 ===
    If selectiononly Then
        If FromRange.Column + FromRange.Columns.Count <= ToSheet.Columns.Count Then
            ToSheet.Range(ToSheet.Cells(1, FromRange.Column + FromRange.Columns.Count), _
                         ToSheet.Cells(1, ToSheet.Columns.Count)).EntireColumn.Delete
        End If
        If FromRange.Row + FromRange.Rows.Count <= ToSheet.Rows.Count Then
            ToSheet.Range(ToSheet.Cells(FromRange.Row + FromRange.Rows.Count, 1), _
                         ToSheet.Cells(ToSheet.Rows.Count, 1)).EntireRow.Delete
        End If
        If FromRange.Cells(1, 1).Column > 1 Then
            ToSheet.Range(ToSheet.Columns(1), ToSheet.Columns(FromRange.Cells(1, 1).Column - 1)).Delete
        End If
        If FromRange.Cells(1, 1).Row > 1 Then
            ToSheet.Range(ToSheet.Rows(1), ToSheet.Rows(FromRange.Cells(1, 1).Row - 1)).Delete
        End If
    End If
    
    ToSheet.Cells(1, 1).Select
    ToBook.Sheets(1).Activate
    maked.Add ToBook
    RFVBA
    
    Set 切り出し高速 = ToBook
End Function

' === 罫線情報を文字列として取得 ===
Private Function GetBorderKey(cell As Range) As String
    Dim key As String
    Dim b As Variant
    Dim borderTypes As Variant
    borderTypes = Array(xlEdgeTop, xlEdgeBottom, xlEdgeLeft, xlEdgeRight, xlInsideHorizontal, xlInsideVertical)
    
    For Each b In borderTypes
        With cell.Borders(b)
            key = key & b & ":" & .LineStyle & "," & .Color & "," & .Weight & ";"
        End With
    Next b
    GetBorderKey = key
End Function

' === 罫線情報を適用 ===
Private Sub ApplyBorderKey(cell As Range, key As String)
    Dim parts() As String
    Dim i As Long
    Dim borderTypes As Variant
    borderTypes = Array(xlEdgeTop, xlEdgeBottom, xlEdgeLeft, xlEdgeRight, xlInsideHorizontal, xlInsideVertical)
    
    parts = Split(key, ";")
    On Error Resume Next
    For i = 0 To 5
        If parts(i) <> "" Then
            Dim vals() As String
            vals = Split(parts(i), ":")
            If UBound(vals) >= 1 Then
                Dim borderVals() As String
                borderVals = Split(vals(1), ",")
                With cell.Borders(borderTypes(i))
                    If CLng(borderVals(0)) <> -4142 Then
                        .LineStyle = CLng(borderVals(0))
                        .Color = CLng(borderVals(1))
                        .Weight = CLng(borderVals(2))
                    End If
                End With
            End If
        End If
    Next i
    On Error GoTo 0
End Sub

' === 既存のインターフェースを維持 ===
Public Function 切り出し(FromBook As Workbook, fromsheet As Worksheet, FromRange As Range, ToBook As Workbook, ToSheet As Worksheet, Optional selectiononly = False, Optional fitpage = False) As Workbook
    Set 切り出し = 切り出し高速(FromBook, fromsheet, FromRange, ToBook, ToSheet, selectiononly, fitpage)
End Function

' === 以下、既存のサブルーチン（省略）===
' 切り出し_全体、切り出し_選択、切り出し_定義、
' 切り出し_定義_全体、切り出し_定義_シート内、
' 切り出し_定義とスライサー、切り出し_定義とスライサー_シート内、
' 雛形適応、DirectCall、条件書式の正規化、書式を正規化、
' GetFormatKey、ApplyFormatFromKey、MergeRanges、シートコピー
' は既存のまま維持
