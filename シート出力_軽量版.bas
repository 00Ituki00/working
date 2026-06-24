' シート出力軽量版
' 整理・軽量化：不要なDoEvents削減、変数スコープ縮小、処理の集約

Global maked As New Collection

' === メイン関数：軽量版 ===
Public Function 切り出し軽量(FromBook As Workbook, fromsheet As Worksheet, FromRange As Range, ToBook As Workbook, ToSheet As Worksheet, Optional selectiononly = False, Optional fitpage = False) As Workbook
    Dim pt As PivotTable, lo As ListObject, sh As Shape
    Dim dataArr As Variant
    Dim r As Long, c As Long
    
    STVBA
    
    ' テーマ適用
    ToBook.ApplyTheme "C:\Users\h_ikegami\AppData\Roaming\Microsoft\Templates\Document Themes\default.thmx"
    
    ' === Step 0: FromRangeをUsedRangeで絞る ===
    Set FromRange = Intersect(FromRange, fromsheet.UsedRange)
    
    ' === Step 1: ピボットテーブル範囲を保存 ===
    Dim ptRanges As New Collection
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
    
    ' === Step 2: テーブル範囲を保存 ===
    Dim tblRanges As New Collection
    For Each lo In fromsheet.ListObjects
        If Not Intersect(lo.Range, FromRange) Is Nothing Then
            On Error Resume Next
            Dim loIntersect As Range
            Set loIntersect = Intersect(lo.Range, FromRange)
            If Not loIntersect Is Nothing Then tblRanges.Add loIntersect
            On Error GoTo 0
        End If
    Next lo
    
    ' === Step 3: データを配列で一括コピー ===
    dataArr = FromRange.Value2
    
    ' === Step 4: 一括書き込み ===
    Dim ToRange As Range
    Set ToRange = ToSheet.Range(ToSheet.Cells(FromRange.Row, FromRange.Column), _
                                ToSheet.Cells(FromRange.Row + FromRange.Rows.Count - 1, _
                                              FromRange.Column + FromRange.Columns.Count - 1))
    ToRange.NumberFormat = "@"
    ToRange.Value2 = dataArr
    
    ' === Step 5: 書式・列幅をコピー ===
    On Error GoTo recopy1
    fromsheet.Activate: FromRange.Copy: ToSheet.Activate
    ToRange.PasteSpecial Paste:=xlPasteFormats
    On Error GoTo 0
    
    ' 列幅コピー
    Dim col As Long
    For col = FromRange.Column To FromRange.Column + FromRange.Columns.Count - 1
        ToSheet.Columns(col).ColumnWidth = fromsheet.Columns(col).ColumnWidth
    Next col
    
    ' === Step 6: ピボットテーブルを上書き ===
    Dim ptItem As Range
    For Each ptItem In ptRanges
        On Error Resume Next
        ptItem.Copy
        ToSheet.Cells(ptItem.Cells(1, 1).Row, ptItem.Cells(1, 1).Column).PasteSpecial Paste:=xlPasteValuesAndNumberFormats
        ToSheet.Cells(ptItem.Cells(1, 1).Row, ptItem.Cells(1, 1).Column).PasteSpecial Paste:=xlPasteFormats
        On Error GoTo 0
    Next ptItem
    
    ' === Step 7: テーブルを上書き ===
    Dim tblItem As Range
    For Each tblItem In tblRanges
        On Error Resume Next
        tblItem.Copy
        ToSheet.Cells(tblItem.Cells(1, 1).Row, tblItem.Cells(1, 1).Column).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
        On Error GoTo 0
    Next tblItem
    
    ' === Step 8: 条件書式を正規化 ===
    書式を正規化 ToRange
    
    ' === Step 9: 図形をコピー ===
    ToSheet.Activate
    Dim newShape As Shape
    For Each sh In fromsheet.Shapes
        If Not Intersect(Range(sh.TopLeftCell, sh.BottomRightCell), FromRange) Is Nothing Then
            If sh.Type = msoChart Or sh.Type = 17 Or sh.Type = 13 Then
                On Error GoTo recopy
                Application.CutCopyMode = False
                sh.Copy
                ToSheet.PasteSpecial Format:=0
                Set newShape = ToSheet.Shapes(ToSheet.Shapes.Count)
                newShape.Top = ToSheet.Range(sh.TopLeftCell.Address).Top
                newShape.Left = ToSheet.Range(sh.TopLeftCell.Address).Left
                newShape.SetShapesDefaultProperties
                On Error GoTo 0
            End If
        End If
        If ToSheet.Shapes.Count > 0 Then 図形スナップ ToSheet.Shapes(ToSheet.Shapes.Count)
    Next sh
    
    ' === Step 10: ウィンドウ設定 ===
    ToSheet.Cells(1).Select
    ToSheet.Parent.Windows(1).Zoom = fromsheet.Parent.Windows(1).Zoom
    If fromsheet.Parent.Windows(1).FreezePanes Then
        ToSheet.Parent.Windows(1).SplitVertical = fromsheet.Parent.Windows(1).SplitVertical
        ToSheet.Parent.Windows(1).SplitHorizontal = fromsheet.Parent.Windows(1).SplitHorizontal
        ToSheet.Application.ActiveWindow.FreezePanes = True
    End If
    
    ' === Step 11: 名前定義をコピー ===
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
    
    ' === Step 12: 選択範囲外の削除 ===
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
    
    Set 切り出し軽量 = ToBook
    Exit Function
    
recopy:
    DoEvents: Resume
recopy1:
    fromsheet.Activate: FromRange.Copy: ToSheet.Activate: DoEvents: Resume
End Function

' === 既存のインターフェースを維持 ===
Public Function 切り出し(FromBook As Workbook, fromsheet As Worksheet, FromRange As Range, ToBook As Workbook, ToSheet As Worksheet, Optional selectiononly = False, Optional fitpage = False) As Workbook
    Set 切り出し = 切り出し軽量(FromBook, fromsheet, FromRange, ToBook, ToSheet, selectiononly, fitpage)
End Function
