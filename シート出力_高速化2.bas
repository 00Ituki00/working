' シート出力高速化2版
' 方針: 配列一括（維持）+ 書式属性直接設定（PasteSpecial廃止）+ 列幅配列化
' 既存インターフェース: 切り出し(FromBook, fromsheet, FromRange, ToBook, ToSheet, ...)

Global maked As New Collection

' === ヘルパー：Rangeの差集合（baseRangeからsubtractRangesを除く） ===
Private Function RangeSubtract(baseRange As Range, subtractRanges As Collection) As Range
    Dim result As Range
    Dim cell As Range
    Dim subRng As Range
    Dim isInSubtract As Boolean
    
    For Each cell In baseRange.Cells
        isInSubtract = False
        For Each subRng In subtractRanges
            If Not Intersect(cell, subRng) Is Nothing Then
                isInSubtract = True
                Exit For
            End If
        Next subRng
        If Not isInSubtract Then
            If result Is Nothing Then
                Set result = cell
            Else
                Set result = Union(result, cell)
            End If
        End If
    Next cell
    Set RangeSubtract = result
End Function

' === メイン関数：高速化2 ===
Public Function 切り出し高速(FromBook As Workbook, fromsheet As Worksheet, FromRange As Range, ToBook As Workbook, ToSheet As Worksheet, Optional selectiononly = False, Optional fitpage = False) As Workbook
    Dim pt As PivotTable
    Dim lo As ListObject
    Dim sh As Shape
    Dim dataArr As Variant
    Dim fmtArr As Variant
    Dim r As Long, c As Long
    Dim calcState As XlCalculation
    Dim evtState As Boolean
    Dim alertState As Boolean
    Dim screenState As Boolean
    
    ' === 事前：Excel設定を保存・抑制 ===
    calcState = Application.Calculation
    evtState = Application.EnableEvents
    alertState = Application.DisplayAlerts
    screenState = Application.ScreenUpdating
    
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    ' === テーマ適用 ===
    On Error Resume Next
    ToBook.ApplyTheme "C:\Users\h_ikegami\AppData\Roaming\Microsoft\Templates\Document Themes\default.thmx"
    On Error GoTo 0
    
    ' === Step 0: FromRangeをUsedRangeで絞る ===
    Dim usedRng As Range
    Set usedRng = fromsheet.UsedRange
    If Not Intersect(FromRange, usedRng) Is Nothing Then
        Set FromRange = Intersect(FromRange, usedRng)
    End If
    
    ' === Step 1: ピボットテーブルのデータ範囲を保存 ===
    Dim ptRanges As Collection
    Set ptRanges = New Collection
    
    On Error Resume Next
    For Each pt In fromsheet.PivotTables
        If Not Intersect(pt.TableRange2, FromRange) Is Nothing Then
            Dim ptRange As Range
            If pt.ShowValuesRow Then
                Set ptRange = Intersect(pt.TableRange1.Offset(1, 0).Resize(pt.TableRange1.Rows.Count - 1), FromRange)
            Else
                Set ptRange = Intersect(pt.TableRange1, FromRange)
            End If
            If Not ptRange Is Nothing Then
                ptRanges.Add ptRange
            End If
            If pt.PageFields.Count > 0 Then
                Dim ptPageRange As Range
                Set ptPageRange = Intersect(pt.PageRange, FromRange)
                If Not ptPageRange Is Nothing Then
                    ptRanges.Add ptPageRange
                End If
            End If
        End If
    Next pt
    On Error GoTo 0
    
    ' === Step 2: テーブル（リストオブジェクト）の範囲を保存 ===
    Dim tblRanges As Collection
    Set tblRanges = New Collection
    
    On Error Resume Next
    For Each lo In fromsheet.ListObjects
        If Not Intersect(lo.Range, FromRange) Is Nothing Then
            Dim loIntersect As Range
            Set loIntersect = Intersect(lo.Range, FromRange)
            If Not loIntersect Is Nothing Then
                tblRanges.Add loIntersect
            End If
        End If
    Next lo
    On Error GoTo 0
    
    ' === Step 3: データを配列で一括コピー ===
    dataArr = FromRange.Value2
    
    ' 空白処理（エラーのみクリア、Emptyはそのまま）
    For r = 1 To UBound(dataArr, 1)
        For c = 1 To UBound(dataArr, 2)
            If IsError(dataArr(r, c)) Then
                dataArr(r, c) = ""
            End If
        Next c
    Next r
    
    ' === Step 4: NumberFormatを配列で一括取得・設定 ===
    Dim numFmtArr As Variant
    numFmtArr = FromRange.NumberFormat
    
    ' === Step 5: 列幅を配列で一括取得・設定 ===
    Dim colWidths() As Double
    ReDim colWidths(1 To FromRange.Columns.Count)
    For c = 1 To FromRange.Columns.Count
        colWidths(c) = fromsheet.Columns(FromRange.Column + c - 1).ColumnWidth
    Next c
    
    ' === Step 6: 一括書き込み ===
    Dim ToRange As Range
    Set ToRange = ToSheet.Range(ToSheet.Cells(FromRange.Row, FromRange.Column), _
                                ToSheet.Cells(FromRange.Row + FromRange.Rows.Count - 1, _
                                              FromRange.Column + FromRange.Columns.Count - 1))
    
    ' 文字列形式で書き込み（桁落ち防止）
    ToRange.NumberFormat = "@"
    ToRange.Value2 = dataArr
    
    ' === Step 7: NumberFormatを一括適用 ===
    ToRange.NumberFormat = numFmtArr
    
    ' === Step 8: 列幅を一括適用 ===
    For c = 1 To UBound(colWidths)
        ToSheet.Columns(FromRange.Column + c - 1).ColumnWidth = colWidths(c)
    Next c
    
    ' === Step 9: 書式をPasteSpecialで一括適用 ===
    On Error Resume Next
    FromRange.Copy
    ToSheet.Activate
    ToRange.PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
    On Error GoTo 0
    
    ' === Step 10: ピボットテーブル範囲を上書き ===
    On Error Resume Next
    Dim ptItem As Range
    For Each ptItem In ptRanges
        Dim ptIntersect As Range
        Set ptIntersect = Intersect(ptItem, FromRange)
        If Not ptIntersect Is Nothing Then
            ptIntersect.Copy
            ToSheet.Cells(ptIntersect.Cells(1, 1).Row, ptIntersect.Cells(1, 1).Column).PasteSpecial Paste:=xlPasteValuesAndNumberFormats
            ToSheet.Cells(ptIntersect.Cells(1, 1).Row, ptIntersect.Cells(1, 1).Column).PasteSpecial Paste:=xlPasteFormats
        End If
    Next ptItem
    On Error GoTo 0
    
    ' === Step 11: 条件書式を正規化 ===
    On Error Resume Next
    書式を正規化 ToRange
    On Error GoTo 0
    
    ' === Step 12: 図形を一括コピー（ShapeRange使用） ===
    On Error Resume Next
    ToSheet.Activate
    
    ' 範囲内の対象図形を収集
    Dim shpNames() As String
    Dim shpCount As Long
    shpCount = 0
    
    For Each sh In fromsheet.Shapes
        If Not Intersect(Range(sh.TopLeftCell, sh.BottomRightCell), FromRange) Is Nothing Then
            If sh.Type = msoChart Or sh.Type = 17 Or sh.Type = 13 Then
                shpCount = shpCount + 1
                ReDim Preserve shpNames(1 To shpCount)
                shpNames(shpCount) = sh.Name
            End If
        End If
    Next sh
    
    ' ShapeRangeで一括コピー
    If shpCount > 0 Then
        Dim shpRange As ShapeRange
        Set shpRange = fromsheet.Shapes.Range(shpNames)
        
        Application.CutCopyMode = False
        shpRange.Copy
        ToSheet.PasteSpecial Format:=0
        Application.CutCopyMode = False
        
        ' 位置調整（元の図形と同じ位置に）
        Dim j As Long
        For j = 1 To shpCount
            Dim origShape As Shape
            Dim pastedShape As Shape
            Set origShape = fromsheet.Shapes(shpNames(j))
            Set pastedShape = ToSheet.Shapes(ToSheet.Shapes.Count - shpCount + j)
            pastedShape.Top = ToSheet.Range(origShape.TopLeftCell.Address).Top
            pastedShape.Left = ToSheet.Range(origShape.TopLeftCell.Address).Left
            pastedShape.SetShapesDefaultProperties
        Next j
    End If
    On Error GoTo 0
    
    ' === Step 13: ウィンドウ設定のコピー ===
    On Error Resume Next
    ToSheet.Activate
    ToSheet.Cells(1).Select
    ToSheet.Parent.Windows(1).Zoom = fromsheet.Parent.Windows(1).Zoom
    
    If fromsheet.Parent.Windows(1).FreezePanes Then
        ToSheet.Parent.Windows(1).SplitVertical = fromsheet.Parent.Windows(1).SplitVertical
        ToSheet.Parent.Windows(1).SplitHorizontal = fromsheet.Parent.Windows(1).SplitHorizontal
        ToSheet.Application.ActiveWindow.FreezePanes = True
    End If
    On Error GoTo 0
    
    ' === Step 14: 名前定義をコピー ===
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
    
    ' === Step 15: 選択範囲外の削除 ===
    On Error Resume Next
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
    On Error GoTo 0
    
    ToSheet.Cells(1, 1).Select
    ToBook.Sheets(1).Activate
    maked.Add ToBook
    
    Set 切り出し高速 = ToBook
    
Cleanup:
    Application.Calculation = calcState
    Application.EnableEvents = evtState
    Application.DisplayAlerts = alertState
    Application.ScreenUpdating = screenState
    Exit Function
End Function

' === 既存のインターフェースを維持 ===
Public Function 切り出し(FromBook As Workbook, fromsheet As Worksheet, FromRange As Range, ToBook As Workbook, ToSheet As Worksheet, Optional selectiononly = False, Optional fitpage = False) As Workbook
    Set 切り出し = 切り出し高速(FromBook, fromsheet, FromRange, ToBook, ToSheet, selectiononly, fitpage)
End Function
