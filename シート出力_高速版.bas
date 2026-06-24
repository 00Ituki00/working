' シート出力高速化版
' ピボットテーブル・テーブルを事前に値化し、配列一括処理で高速化

Global maked As New Collection

' === メイン関数：高速版 ===
Public Function 切り出し高速(FromBook As Workbook, fromsheet As Worksheet, FromRange As Range, ToBook As Workbook, ToSheet As Worksheet, Optional selectiononly = False, Optional fitpage = False) As Workbook
    Dim pt As PivotTable
    Dim lo As ListObject
    Dim sh As Shape
    Dim chartObj As ChartObject
    Dim dataArr As Variant
    Dim fmtArr() As Variant
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
    
    ' === Step 3: セルデータを配列で一括コピー ===
    firstRow = FromRange.Row
    lastRow = FromRange.Row + FromRange.Rows.Count - 1
    
    ' 列書式を保存
    ReDim fmtArr(1 To FromRange.Columns.Count)
    For c = 1 To FromRange.Columns.Count
        fmtArr(c) = fromsheet.Cells(firstRow, FromRange.Column + c - 1).NumberFormat
    Next c
    
    ' データを配列で取得
    dataArr = FromRange.Value2
    
    ' 配列内で空白処理（必要に応じて）
    For r = 1 To UBound(dataArr, 1)
        For c = 1 To UBound(dataArr, 2)
            If IsEmpty(dataArr(r, c)) Or IsError(dataArr(r, c)) Then
                dataArr(r, c) = ""
            End If
        Next c
    Next r
    DoEvents
    
    ' 一括書き込み（文字列として書き込んで桁落ち防止）
    Dim ToRange As Range
    Set ToRange = ToSheet.Range(ToSheet.Cells(FromRange.Row, FromRange.Column), _
                                ToSheet.Cells(FromRange.Row + FromRange.Rows.Count - 1, _
                                              FromRange.Column + FromRange.Columns.Count - 1))
    
    ToRange.NumberFormat = "@"
    ToRange.Value2 = dataArr
    
    ' 列書式を復元
    For c = 1 To FromRange.Columns.Count
        ToRange.Columns(c).NumberFormat = fmtArr(c)
    Next c
    DoEvents
    
    ' === Step 4: 書式を正規化（条件書式を実書式へ）===
    書式を正規化 ToRange
    DoEvents
    
    ' === Step 5: 図形を画像としてコピー ===
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
    
    ' その他の図形（必要なもののみ）
    For Each sh In fromsheet.Shapes
        If Not Intersect(Range(sh.TopLeftCell, sh.BottomRightCell), FromRange) Is Nothing Then
            ' グラフ以外の図形（テキストボックスなど必要に応じて）
            If sh.Type <> msoChart And sh.Type <> 17 Then
                On Error Resume Next
                imgIndex = imgIndex + 1
                Dim shapeFile As String
                shapeFile = tempPath & "shape_" & imgIndex & ".png"
                
                ' 図形を画像としてエクスポート
                sh.Copy
                ' クリップボードから画像を取得して保存（別途関数が必要）
                ' 簡易版：図形をコピーして貼り付け
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
    
    ' === Step 6: ウィンドウ設定のコピー ===
    ToSheet.Activate
    ToSheet.Cells(1).Select
    ToSheet.Parent.Windows(1).Zoom = fromsheet.Parent.Windows(1).Zoom
    
    If fromsheet.Parent.Windows(1).FreezePanes Then
        ToSheet.Parent.Windows(1).SplitVertical = fromsheet.Parent.Windows(1).SplitVertical
        ToSheet.Parent.Windows(1).SplitHorizontal = fromsheet.Parent.Windows(1).SplitHorizontal
        ToSheet.Application.ActiveWindow.FreezePanes = True
    End If
    DoEvents
    
    ' === Step 7: 名前定義をコピー ===
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
    
    ' === Step 8: 選択範囲外の削除 ===
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

' === 既存のインターフェースを維持 ===
Public Function 切り出し(FromBook As Workbook, fromsheet As Worksheet, FromRange As Range, ToBook As Workbook, ToSheet As Worksheet, Optional selectiononly = False, Optional fitpage = False) As Workbook
    ' 既存の呼び出しを高速版にリダイレクト
    Set 切り出し = 切り出し高速(FromBook, fromsheet, FromRange, ToBook, ToSheet, selectiononly, fitpage)
End Function

' === 既存のサブルーチン（インターフェース維持） ===
Sub 切り出し_全体()
    Dim FromBook As Workbook, ToBook As Workbook
    Dim fromsheet As Worksheet, ToSheet As Worksheet
    STVBA
    Set FromBook = ActiveWorkbook: Set fromsheet = ActiveSheet
    Set ToBook = Workbooks.Add
    Set ToSheet = ToBook.Worksheets.Add
    If ToSheet.Name <> fromsheet.Name Then ToSheet.Name = fromsheet.Name
    ToBook.Worksheets("Sheet1").Delete
    RFVBA
    切り出し FromBook, fromsheet, fromsheet.UsedRange, ToBook, ToSheet
End Sub

Sub 切り出し_選択()
    Dim FromBook As Workbook, ToBook As Workbook
    Dim fromsheet As Worksheet, ToSheet As Worksheet
    Dim SelectRange As Range
    STVBA
    Set SelectRange = Selection
    Set FromBook = ActiveWorkbook: Set fromsheet = ActiveSheet
    Set ToBook = Workbooks.Add
    Set ToSheet = ToBook.Worksheets.Add
    If ToSheet.Name <> fromsheet.Name Then ToSheet.Name = fromsheet.Name
    ToBook.Worksheets("Sheet1").Delete
    RFVBA
    切り出し FromBook, fromsheet, SelectRange, ToBook, ToSheet, True
End Sub

' === 名前定義からの切り出し（既存互換） ===
Sub 切り出し_定義(T As Variant, Optional itemname = "", Optional fitpage = False)
    ' ... 既存のコードを維持 ...
End Sub

' === その他の既存サブルーチン ===
' 切り出し_定義_全体、切り出し_定義_シート内、
' 切り出し_定義とスライサー、切り出し_定義とスライサー_シート内、
' 雛形適応、DirectCall、条件書式の正規化、書式を正規化、
' GetFormatKey、ApplyFormatFromKey、MergeRanges、シートコピー
' は既存のまま維持
