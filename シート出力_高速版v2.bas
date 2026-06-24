' シート出力高速化版v2
' 画像ファイル経由方式：図形を画像化して高速コピー

Global maked As New Collection

' === メイン関数：高速版v2（画像ファイル経由）===
Public Function 切り出し高速v2(FromBook As Workbook, fromsheet As Worksheet, FromRange As Range, ToBook As Workbook, ToSheet As Worksheet, Optional selectiononly = False, Optional fitpage = False) As Workbook
    Dim pt As PivotTable
    Dim lo As ListObject
    Dim sh As Shape
    Dim chartObj As ChartObject
    Dim dataArr As Variant
    Dim r As Long, c As Long
    Dim tempPath As String
    Dim imgIndex As Long: imgIndex = 0
    
    STVBA
    tempPath = Environ("TEMP") & "\SheetExport_"
    
    ' テーマ適用
    ToBook.ApplyTheme "C:\Users\h_ikegami\AppData\Roaming\Microsoft\Templates\Document Themes\default.thmx"
    DoEvents
    
    ' === Step 0: FromRangeをUsedRangeで絞る ===
    Dim usedRng As Range
    Set usedRng = fromsheet.UsedRange
    If Not Intersect(FromRange, usedRng) Is Nothing Then
        Set FromRange = Intersect(FromRange, usedRng)
    End If
    DoEvents
    
    ' === Step 1: ピボットテーブルの範囲を保存 ===
    Dim ptRanges As Collection
    Set ptRanges = New Collection
    
    For Each pt In fromsheet.PivotTables
        If Not Intersect(pt.TableRange2, FromRange) Is Nothing Then
            On Error Resume Next
            Dim ptRange As Range
            If pt.ShowValuesRow Then
                Set ptRange = Intersect(pt.TableRange1.Offset(1, 0).Resize(pt.TableRange1.Rows.Count - 1), FromRange)
            Else
                Set ptRange = Intersect(pt.TableRange1, FromRange)
            End If
            If Not ptRange Is Nothing Then ptRanges.Add ptRange
            
            If pt.PageFields.Count > 0 Then
                Dim ptPageRange As Range
                Set ptPageRange = Intersect(pt.PageRange, FromRange)
                If Not ptPageRange Is Nothing Then ptRanges.Add ptPageRange
            End If
            On Error GoTo 0
        End If
    Next pt
    DoEvents
    
    ' === Step 2: テーブル（リストオブジェクト）を値化 ===
    Dim tblRanges As Collection
    Set tblRanges = New Collection
    
    For Each lo In fromsheet.ListObjects
        If Not Intersect(lo.Range, FromRange) Is Nothing Then
            On Error Resume Next
            Dim loIntersect As Range
            Set loIntersect = Intersect(lo.Range, FromRange)
            If Not loIntersect Is Nothing Then tblRanges.Add loIntersect
            On Error GoTo 0
        End If
    Next lo
    DoEvents
    
    ' === Step 3: データを配列で一括コピー ===
    dataArr = FromRange.Value2
    
    For r = 1 To UBound(dataArr, 1)
        For c = 1 To UBound(dataArr, 2)
            If IsEmpty(dataArr(r, c)) Or IsError(dataArr(r, c)) Then
                dataArr(r, c) = ""
            End If
        Next c
    Next r
    DoEvents
    
    ' === Step 4: 一括書き込み ===
    Dim ToRange As Range
    Set ToRange = ToSheet.Range(ToSheet.Cells(FromRange.Row, FromRange.Column), _
                                ToSheet.Cells(FromRange.Row + FromRange.Rows.Count - 1, _
                                              FromRange.Column + FromRange.Columns.Count - 1))
    
    ToRange.NumberFormat = "@"
    ToRange.Value2 = dataArr
    DoEvents
    
    ' === Step 5: 書式をPasteSpecialで一括適用 ===
    On Error GoTo recopy1
    fromsheet.Activate: FromRange.Copy: ToSheet.Activate
    ToRange.PasteSpecial Paste:=xlPasteFormats
    On Error GoTo 0
    DoEvents
    
    ' === Step 5.5: 列幅をコピー ===
    Dim col As Long
    For col = FromRange.Column To FromRange.Column + FromRange.Columns.Count - 1
        ToSheet.Columns(col).ColumnWidth = fromsheet.Columns(col).ColumnWidth
    Next col
    DoEvents
    
    ' === Step 6: ピボットテーブルを上書きコピー ===
    Dim ptItem As Range
    For Each ptItem In ptRanges
        On Error Resume Next
        Dim ptIntersect As Range
        Set ptIntersect = Intersect(ptItem, FromRange)
        If Not ptIntersect Is Nothing Then
            ptIntersect.Copy
            ToSheet.Cells(ptIntersect.Cells(1, 1).Row, ptIntersect.Cells(1, 1).Column).PasteSpecial Paste:=xlPasteValuesAndNumberFormats
            ToSheet.Cells(ptIntersect.Cells(1, 1).Row, ptIntersect.Cells(1, 1).Column).PasteSpecial Paste:=xlPasteFormats
        End If
        On Error GoTo 0
    Next ptItem
    DoEvents
    
    ' === Step 6.5: 条件書式を正規化 ===
    書式を正規化 ToRange
    DoEvents
    
    ' === Step 7: 図形を画像ファイル経由でコピー ===
    ' 高速化：Activate不要、直接ファイル出力
    
    ' グラフ（ChartObjects）を画像化
    For Each chartObj In fromsheet.ChartObjects
        If Not Intersect(Range(chartObj.TopLeftCell, chartObj.BottomRightCell), FromRange) Is Nothing Then
            On Error Resume Next
            imgIndex = imgIndex + 1
            Dim chartFile As String
            chartFile = tempPath & "chart_" & imgIndex & ".png"
            chartObj.Chart.Export Filename:=chartFile, FilterName:="PNG"
            
            ' 画像を挿入して位置・サイズを設定
            Dim insertedPic As Shape
            Set insertedPic = ToSheet.Pictures.Insert(chartFile)
            With insertedPic
                .Top = ToSheet.Range(chartObj.TopLeftCell.Address).Top
                .Left = ToSheet.Range(chartObj.TopLeftCell.Address).Left
                .Width = chartObj.Width
                .Height = chartObj.Height
            End With
            
            ' 一時ファイル削除
            If Dir(chartFile) <> "" Then Kill chartFile
            On Error GoTo 0
        End If
    Next chartObj
    DoEvents
    
    ' その他の図形（ChartObjects以外）を画像化
    For Each sh In fromsheet.Shapes
        If Not Intersect(Range(sh.TopLeftCell, sh.BottomRightCell), FromRange) Is Nothing Then
            ' ChartObjectsは既に処理済みなのでスキップ
            If sh.Type <> msoChart And sh.Type <> 17 Then
                On Error Resume Next
                imgIndex = imgIndex + 1
                Dim shapeFile As String
                shapeFile = tempPath & "shape_" & imgIndex & ".png"
                
                ' 図形を画像としてエクスポート
                sh.Copy
                ' 一時的に画像として貼り付けてエクスポート
                ToSheet.Paste
                Dim tempShape As Shape
                Set tempShape = ToSheet.Shapes(ToSheet.Shapes.Count)
                tempShape.Copy
                
                ' クリップボードから画像を取得して保存（別途APIが必要な場合あり）
                ' 簡易版：直接コピー方式にフォールバック
                With tempShape
                    .Top = ToSheet.Range(sh.TopLeftCell.Address).Top
                    .Left = ToSheet.Range(sh.TopLeftCell.Address).Left
                End With
                On Error GoTo 0
            End If
        End If
    Next sh
    DoEvents
    
    ' === Step 8: ウィンドウ設定のコピー ===
    ToSheet.Activate
    ToSheet.Cells(1).Select
    ToSheet.Parent.Windows(1).Zoom = fromsheet.Parent.Windows(1).Zoom
    
    If fromsheet.Parent.Windows(1).FreezePanes Then
        ToSheet.Parent.Windows(1).SplitVertical = fromsheet.Parent.Windows(1).SplitVertical
        ToSheet.Parent.Windows(1).SplitHorizontal = fromsheet.Parent.Windows(1).SplitHorizontal
        ToSheet.Application.ActiveWindow.FreezePanes = True
    End If
    DoEvents
    
    ' === Step 9: 名前定義をコピー ===
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
    
    ' === Step 10: 選択範囲外の削除 ===
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
    
    Set 切り出し高速v2 = ToBook
    Exit Function
    
recopy1:
    fromsheet.Activate: FromRange.Copy: ToSheet.Activate: DoEvents: Resume
End Function

' === 既存のインターフェースを維持 ===
Public Function 切り出し(FromBook As Workbook, fromsheet As Worksheet, FromRange As Range, ToBook As Workbook, ToSheet As Worksheet, Optional selectiononly = False, Optional fitpage = False) As Workbook
    Set 切り出し = 切り出し高速v2(FromBook, fromsheet, FromRange, ToBook, ToSheet, selectiononly, fitpage)
End Function
