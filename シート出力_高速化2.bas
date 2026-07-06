' シート出力高速化2版
' 方針: 配列一括（維持）+ 書式属性直接設定（PasteSpecial廃止）+ 列幅配列化
' 既存インターフェース: 切り出し(FromBook, fromsheet, FromRange, ToBook, ToSheet, ...)

Global maked As New Collection

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
    
    On Error GoTo Cleanup
    
    ' === Step 0: FromRangeをUsedRangeで絞る ===
    Dim usedRng As Range
    Set usedRng = fromsheet.UsedRange
    If Not Intersect(FromRange, usedRng) Is Nothing Then
        Set FromRange = Intersect(FromRange, usedRng)
    End If
    
    ' === Step 1: ピボットテーブルのデータ範囲を保存 ===
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
            On Error GoTo Cleanup
        End If
    Next pt
    
    ' === Step 2: テーブル（リストオブジェクト）の範囲を保存 ===
    Dim tblRanges As Collection
    Set tblRanges = New Collection
    
    For Each lo In fromsheet.ListObjects
        If Not Intersect(lo.Range, FromRange) Is Nothing Then
            On Error Resume Next
            Dim loIntersect As Range
            Set loIntersect = Intersect(lo.Range, FromRange)
            If Not loIntersect Is Nothing Then
                tblRanges.Add loIntersect
            End If
            On Error GoTo Cleanup
        End If
    Next lo
    
    ' === Step 3: データを配列で一括コピー ===
    dataArr = FromRange.Value2
    
    ' 空白処理
    For r = 1 To UBound(dataArr, 1)
        For c = 1 To UBound(dataArr, 2)
            If IsEmpty(dataArr(r, c)) Or IsError(dataArr(r, c)) Then
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
    
    ' === Step 9: 書式属性を直接設定（PasteSpecial xlPasteFormats 代替） ===
    Call 書式を直接設定(fromsheet, FromRange, ToSheet, ToRange)
    
    ' === Step 10: ピボットテーブル範囲を上書き ===
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
        On Error GoTo Cleanup
    Next ptItem
    
    ' === Step 11: 条件書式を正規化 ===
    書式を正規化 ToRange
    
    ' === Step 12: 図形をコピー（Activateなし） ===
    For Each sh In fromsheet.Shapes
        If Not Intersect(Range(sh.TopLeftCell, sh.BottomRightCell), FromRange) Is Nothing Then
            If sh.Type = msoChart Or sh.Type = 17 Or sh.Type = 13 Then
                On Error Resume Next
                sh.Copy
                ToSheet.PasteSpecial Format:=0
                
                Dim newShape As Shape
                Set newShape = ToSheet.Shapes(ToSheet.Shapes.Count)
                newShape.Top = ToSheet.Range(sh.TopLeftCell.Address).Top
                newShape.Left = ToSheet.Range(sh.TopLeftCell.Address).Left
                newShape.SetShapesDefaultProperties
                On Error GoTo Cleanup
            End If
        End If
        If ToSheet.Shapes.Count > 0 Then 図形スナップ ToSheet.Shapes(ToSheet.Shapes.Count)
    Next sh
    
    ' === Step 13: ウィンドウ設定のコピー ===
    ToSheet.Activate
    ToSheet.Cells(1).Select
    ToSheet.Parent.Windows(1).Zoom = fromsheet.Parent.Windows(1).Zoom
    
    If fromsheet.Parent.Windows(1).FreezePanes Then
        ToSheet.Parent.Windows(1).SplitVertical = fromsheet.Parent.Windows(1).SplitVertical
        ToSheet.Parent.Windows(1).SplitHorizontal = fromsheet.Parent.Windows(1).SplitHorizontal
        ToSheet.Application.ActiveWindow.FreezePanes = True
    End If
    
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
    On Error GoTo Cleanup
    
    ' === Step 15: 選択範囲外の削除 ===
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

' === サブルーチン：書式属性を直接設定（PasteSpecial xlPasteFormats 代替） ===
Private Sub 書式を直接設定(srcSheet As Worksheet, srcRange As Range, dstSheet As Worksheet, dstRange As Range)
    Dim r As Long, c As Long
    Dim srcRow As Long, srcCol As Long
    Dim lastRow As Long, lastCol As Long
    lastRow = srcRange.Rows.Count
    lastCol = srcRange.Columns.Count
    
    ' 行単位で処理（Activateなし）
    For r = 1 To lastRow
        srcRow = srcRange.Row + r - 1
        
        ' この行のセルごとに書式を取得（配列に格納）
        Dim cellBgColors() As Variant
        Dim cellFontColors() As Variant
        Dim cellBold() As Variant
        Dim cellItalic() As Variant
        Dim cellHAlign() As Variant
        Dim cellVAlign() As Variant
        
        ReDim cellBgColors(1 To lastCol)
        ReDim cellFontColors(1 To lastCol)
        ReDim cellBold(1 To lastCol)
        ReDim cellItalic(1 To lastCol)
        ReDim cellHAlign(1 To lastCol)
        ReDim cellVAlign(1 To lastCol)
        
        For c = 1 To lastCol
            srcCol = srcRange.Column + c - 1
            Dim srcCell As Range
            Set srcCell = srcSheet.Cells(srcRow, srcCol)
            
            On Error Resume Next
            cellBgColors(c) = srcCell.Interior.Color
            cellFontColors(c) = srcCell.Font.Color
            cellBold(c) = srcCell.Font.Bold
            cellItalic(c) = srcCell.Font.Italic
            cellHAlign(c) = srcCell.HorizontalAlignment
            cellVAlign(c) = srcCell.VerticalAlignment
            On Error GoTo 0
        Next c
        
        ' 列方向に連続した同じ書式をグループ化
        Dim groupStart As Long
        Dim currentBg As Variant
        Dim currentFontColor As Variant
        Dim currentBold As Variant
        Dim currentItalic As Variant
        Dim currentHAlign As Variant
        Dim currentVAlign As Variant
        
        groupStart = 1
        If lastCol > 0 Then
            currentBg = cellBgColors(1)
            currentFontColor = cellFontColors(1)
            currentBold = cellBold(1)
            currentItalic = cellItalic(1)
            currentHAlign = cellHAlign(1)
            currentVAlign = cellVAlign(1)
        End If
        
        For c = 2 To lastCol
            If cellBgColors(c) <> currentBg Or _
               cellFontColors(c) <> currentFontColor Or _
               cellBold(c) <> currentBold Or _
               cellItalic(c) <> currentItalic Or _
               cellHAlign(c) <> currentHAlign Or _
               cellVAlign(c) <> currentVAlign Then
                
                ' 前のグループを設定
                If groupStart <= c - 1 Then
                    Call 設定書式グループ(dstSheet, dstRange.Row + r - 1, dstRange.Column + groupStart - 1, _
                                          c - groupStart, currentBg, currentFontColor, currentBold, _
                                          currentItalic, currentHAlign, currentVAlign)
                End If
                
                ' 新しいグループ開始
                groupStart = c
                currentBg = cellBgColors(c)
                currentFontColor = cellFontColors(c)
                currentBold = cellBold(c)
                currentItalic = cellItalic(c)
                currentHAlign = cellHAlign(c)
                currentVAlign = cellVAlign(c)
            End If
        Next c
        
        ' 最後のグループを設定
        If groupStart <= lastCol Then
            Call 設定書式グループ(dstSheet, dstRange.Row + r - 1, dstRange.Column + groupStart - 1, _
                                  lastCol - groupStart + 1, currentBg, currentFontColor, currentBold, _
                                  currentItalic, currentHAlign, currentVAlign)
        End If
    Next r
End Sub

' === サブルーチン：書式グループの一括設定 ===
Private Sub 設定書式グループ(ws As Worksheet, rowNum As Long, colNum As Long, colCount As Long, _
                              bgColor As Variant, fontColor As Variant, isBold As Variant, _
                              isItalic As Variant, hAlign As Variant, vAlign As Variant)
    On Error Resume Next
    Dim targetRange As Range
    Set targetRange = ws.Range(ws.Cells(rowNum, colNum), ws.Cells(rowNum, colNum + colCount - 1))
    
    With targetRange.Interior
        If bgColor <> xlNone And bgColor <> 0 Then
            .Color = bgColor
        Else
            .ColorIndex = xlNone
        End If
    End With
    
    With targetRange.Font
        If fontColor <> xlNone And fontColor <> 0 Then
            .Color = fontColor
        Else
            .ColorIndex = xlAutomatic
        End If
        If IsNull(isBold) Or IsEmpty(isBold) Then
            .Bold = False
        Else
            .Bold = isBold
        End If
        If IsNull(isItalic) Or IsEmpty(isItalic) Then
            .Italic = False
        Else
            .Italic = isItalic
        End If
    End With
    
    If Not IsNull(hAlign) And Not IsEmpty(hAlign) Then
        targetRange.HorizontalAlignment = hAlign
    End If
    If Not IsNull(vAlign) And Not IsEmpty(vAlign) Then
        targetRange.VerticalAlignment = vAlign
    End If
    
    On Error GoTo 0
End Sub
