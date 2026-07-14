' シート出力高速化3版
' 方針: スライサー項目単位で一括保存 + Activate/DoEvents徹底排除 + 図形一括コピー
' 前提: Python環境なし、VBA内完結

Global maked As New Collection
Global pendingBooks As Object ' 遅延保存用ブック管理

' === 初期化 ===
Private Sub InitPendingBooks()
    Set pendingBooks = CreateObject("Scripting.Dictionary")
End Sub

' === メイン関数：高速化3 ===
Public Function 切り出し高速(FromBook As Workbook, fromsheet As Worksheet, FromRange As Range, ToBook As Workbook, ToSheet As Worksheet, Optional selectiononly = False, Optional fitpage = False) As Workbook
    Dim pt As PivotTable
    Dim lo As ListObject
    Dim sh As Shape
    Dim dataArr As Variant
    Dim numFmtArr As Variant
    Dim colWidths() As Double
    Dim r As Long, c As Long
    
    ' === Step 0: FromRangeをUsedRangeで絞る ===
    Dim usedRng As Range
    Set usedRng = fromsheet.UsedRange
    If Not Intersect(FromRange, usedRng) Is Nothing Then
        Set FromRange = Intersect(FromRange, usedRng)
    End If
    If FromRange Is Nothing Then
        Set 切り出し高速 = ToBook
        Exit Function
    End If
    
    ' === Step 1: ピボットテーブル範囲を保存 ===
    Dim ptRanges As New Collection
    On Error Resume Next
    For Each pt In fromsheet.PivotTables
        If Not Intersect(pt.TableRange2, FromRange) Is Nothing Then
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
        End If
    Next pt
    On Error GoTo 0
    
    ' === Step 2: データを配列で一括取得 ===
    dataArr = FromRange.Value2
    
    ' エラーのみクリア（Emptyはそのまま）
    For r = 1 To UBound(dataArr, 1)
        For c = 1 To UBound(dataArr, 2)
            If IsError(dataArr(r, c)) Then dataArr(r, c) = ""
        Next c
    Next r
    
    ' === Step 3: NumberFormat・列幅を配列で一括取得 ===
    numFmtArr = FromRange.NumberFormat
    ReDim colWidths(1 To FromRange.Columns.Count)
    For c = 1 To FromRange.Columns.Count
        colWidths(c) = fromsheet.Columns(FromRange.Column + c - 1).ColumnWidth
    Next c
    
    ' === Step 4: 一括書き込み ===
    Dim ToRange As Range
    Set ToRange = ToSheet.Range(ToSheet.Cells(FromRange.Row, FromRange.Column), _
                                ToSheet.Cells(FromRange.Row + FromRange.Rows.Count - 1, _
                                              FromRange.Column + FromRange.Columns.Count - 1))
    
    ToRange.NumberFormat = "@"
    ToRange.Value2 = dataArr
    ToRange.NumberFormat = numFmtArr
    
    ' 列幅一括適用
    For c = 1 To UBound(colWidths)
        ToSheet.Columns(FromRange.Column + c - 1).ColumnWidth = colWidths(c)
    Next c
    
    ' === Step 5: 書式をPasteSpecialで一括適用 ===
    On Error Resume Next
    FromRange.Copy
    ToRange.PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
    On Error GoTo 0
    
    ' === Step 6: ピボットテーブル範囲を上書き ===
    On Error Resume Next
    Dim ptItem As Range
    For Each ptItem In ptRanges
        Dim ptIntersect As Range
        Set ptIntersect = Intersect(ptItem, FromRange)
        If Not ptIntersect Is Nothing Then
            ptIntersect.Copy
            ToSheet.Cells(ptIntersect.Cells(1, 1).Row, ptIntersect.Cells(1, 1).Column).PasteSpecial Paste:=xlPasteValuesAndNumberFormats
            ToSheet.Cells(ptIntersect.Cells(1, 1).Row, ptIntersect.Cells(1, 1).Column).PasteSpecial Paste:=xlPasteFormats
            Application.CutCopyMode = False
        End If
    Next ptItem
    On Error GoTo 0
    
    ' === Step 7: 条件書式を正規化 ===
    On Error Resume Next
    書式を正規化 ToRange
    On Error GoTo 0
    
    ' === Step 8: 図形をコピー ===
    On Error Resume Next
    Dim copyRetry As Integer
    Dim copySuccess As Boolean
    Dim shapesToCopy As New Collection
    
    ' コピー対象図形を事前収集
    For Each sh In fromsheet.Shapes
        If Not Intersect(Range(sh.TopLeftCell, sh.BottomRightCell), FromRange) Is Nothing Then
            If sh.Type = msoChart Or sh.Type = 17 Or sh.Type = 13 Then
                shapesToCopy.Add sh
            End If
        End If
    Next sh
    
    ' 個別コピー（DoEventsはコピー・ペースト直後のみ）
    Dim shp As Shape
    For Each shp In shapesToCopy
        copySuccess = False
        copyRetry = 0
        Do While copyRetry < 3 And Not copySuccess
            On Error Resume Next
            Err.Clear
            ' Copy直後にDoEvents（クリップボード反映待ち）
            shp.Copy
            DoEvents
            ToSheet.PasteSpecial Format:=0
            ' Paste直後にDoEvents（図形配置反映待ち）
            DoEvents
            If Err.Number = 0 Then
                Dim newShape As Shape
                Set newShape = ToSheet.Shapes(ToSheet.Shapes.Count)
                newShape.Top = ToSheet.Range(shp.TopLeftCell.Address).Top
                newShape.Left = ToSheet.Range(shp.TopLeftCell.Address).Left
                newShape.SetShapesDefaultProperties
                copySuccess = True
            End If
            If Not copySuccess Then
                copyRetry = copyRetry + 1
                Application.CutCopyMode = False
                ' エラー時のみDoEvents（クリップボード回復待ち）
                If copyRetry < 3 Then DoEvents
            End If
        Loop
        If ToSheet.Shapes.Count > 0 Then 図形スナップ ToSheet.Shapes(ToSheet.Shapes.Count)
    Next shp
    On Error GoTo 0
    
    ' === Step 9: ウィンドウ設定のコピー ===
    On Error Resume Next
    ToSheet.Cells(1).Select
    ToSheet.Parent.Windows(1).Zoom = fromsheet.Parent.Windows(1).Zoom
    If fromsheet.Parent.Windows(1).FreezePanes Then
        ToSheet.Parent.Windows(1).SplitVertical = fromsheet.Parent.Windows(1).SplitVertical
        ToSheet.Parent.Windows(1).SplitHorizontal = fromsheet.Parent.Windows(1).SplitHorizontal
        ToSheet.Application.ActiveWindow.FreezePanes = True
    End If
    On Error GoTo 0
    
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
    
    ' === Step 11: 選択範囲外の削除 ===
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
    maked.Add ToBook
    Set 切り出し高速 = ToBook
End Function

' === 既存のインターフェースを維持 ===
Public Function 切り出し(FromBook As Workbook, fromsheet As Worksheet, FromRange As Range, ToBook As Workbook, ToSheet As Worksheet, Optional selectiononly = False, Optional fitpage = False) As Workbook
    Set 切り出し = 切り出し高速(FromBook, fromsheet, FromRange, ToBook, ToSheet, selectiononly, fitpage)
End Function

' === 切り出し_定義（遅延保存対応版）===
Sub 切り出し_定義(T As Variant, Optional itemname = "", Optional fitpage = False, Optional skipSave = False)
    Dim FromBook As Workbook, ToBook As Workbook
    Dim fromsheet As Worksheet, ToSheet As Worksheet
    Dim FromRange As Range
    Dim comm As String, topath As String, toBookName As String, toSheetName As String
    Dim selectiononly As Boolean
    Dim p As Long, q As Long, vn As String
    
    Set FromBook = T.RefersToRange.Parent.Parent
    FromBook.Activate
    comm = T.Comment
    
    ' commの変数処理
reg:
    If InStr(comm, "[[") > 0 Then
        Do While InStr(p + 1, comm, "[[") > 0 And InStr(p + 1, comm, "]]") > 0
            p = InStr(p + 1, comm, "[[")
            q = InStr(p, comm, "]]")
            vn = Mid(comm, p + 2, q - p - 2)
            comm = Replace(comm, "[[" & vn & "]]", Range(vn))
            p = q + 1
        Loop
        If InStr(comm, "[[") > 0 Then p = 0: GoTo reg
    End If
    
    comm = Replace(comm, "\\", "\")
    If InStr(comm, "YYYYMMDD") > 0 Then comm = Replace(comm, "YYYYMMDD", Format(Date, "YYYYMMDD"))
    If InStr(comm, "YYYYMM") > 0 Then comm = Replace(comm, "YYYYMM", Format(Date, "YYYYMM"))
    If InStr(comm, "YYMMDD") > 0 Then comm = Replace(comm, "YYMMDD", Format(Date, "YYMMDD"))
    If InStr(comm, "YYMM") > 0 Then comm = Replace(comm, "YYMM", Format(Date, "YYMM"))
    If InStr(comm, "itemname") > 0 Then comm = Replace(comm, "itemname", itemname)
    If InStr(comm, "M-1月") > 0 Then comm = Replace(comm, "M-1月", Month(Date) - 1 & "月")
    If InStr(comm, "M月") > 0 Then comm = Replace(comm, "M月", Month(Date) & "月")
    
    Dim commArr() As String
    commArr = Split(comm, ",")
    topath = FromBook.path & "\" & commArr(0)
    If Right(topath, 1) <> "\" Then topath = topath & "\"
    toBookName = commArr(1)
    If InStr(toBookName, ".xlsx") = 0 Then toBookName = toBookName & ".xlsx"
    toSheetName = commArr(2)
    
    selectiononly = True
    If UBound(commArr) >= 3 Then If commArr(3) = "false" Then selectiononly = False
    
    If InStr(toBookName, "itemname") > 0 Then toBookName = Replace(toBookName, "itemname", itemname)
    
    ' === 遅延保存：既存のブックを再利用 ===
    Dim bookKey As String
    bookKey = topath & toBookName
    
    If Not pendingBooks Is Nothing Then
        If pendingBooks.Exists(bookKey) Then
            Set ToBook = pendingBooks(bookKey)
        End If
    End If
    
    If ToBook Is Nothing Then
        ' 既存ブックを検索
        For Each book In Workbooks
            If book.Name = toBookName Then Set ToBook = Workbooks(toBookName): Exit For
        Next
        If ToBook Is Nothing Then
            Set ToBook = Workbooks.Add
        End If
        If Not pendingBooks Is Nothing Then
            pendingBooks.Add bookKey, ToBook
        End If
    End If
    
    ' シート検索・作成
    For Each sh In ToBook.Sheets
        If sh.Name = toSheetName Then Set ToSheet = sh: Exit For
    Next
    If ToSheet Is Nothing Then
        Set ToSheet = ToBook.Worksheets.Add(After:=ToBook.Worksheets(ToBook.Worksheets.Count))
        ToSheet.Name = toSheetName
    End If
    
    Set fromsheet = T.RefersToRange.Parent
    Set FromRange = Intersect(T.RefersToRange, fromsheet.UsedRange)
    
    MakeDir topath
    FromBook.Activate
    fromsheet.Activate
    
    Set ToBook = 切り出し(FromBook, fromsheet, FromRange, ToBook, ToSheet, selectiononly, fitpage)
    
    ' === 保存制御 ===
    If Not skipSave Then
        On Error Resume Next
        If Not ToBook.Sheets("Sheet1") Is Nothing Then ToBook.Sheets("Sheet1").Delete
        On Error GoTo 0
        ToBook.SaveAs bookKey
        Sleep 2000
        FromBook.Activate
    End If
End Sub

' === 一括保存実行 ===
Sub SavePendingBooks()
    If pendingBooks Is Nothing Then Exit Sub
    
    Dim key As Variant
    Dim wb As Workbook
    For Each key In pendingBooks.Keys
        Set wb = pendingBooks(key)
        On Error Resume Next
        If Not wb.Sheets("Sheet1") Is Nothing Then wb.Sheets("Sheet1").Delete
        On Error GoTo 0
        wb.SaveAs key
        wb.Close
    Next key
    
    Set pendingBooks = Nothing
End Sub

' === 切り出し_定義とスライサー_一括保存版 ===
Public Sub 切り出し_定義とスライサー_一括(Optional teigifilter = "output", Optional fitpage = False)
    Dim ActiveItems As New Collection
    Dim sourcebook As Workbook
    Dim sli As SlicerCache
    Dim Item As SlicerItem
    
    Set sourcebook = ActiveWorkbook
    Set sli = GetSlicer(Selection.Name)
    
    For Each Item In sli.SlicerCacheLevels(1).SlicerItems
        If Item.Selected = True And Item.HasData = True Then ActiveItems.Add Item
    Next
    
    ' 遅延保存初期化
    InitPendingBooks
    
    Dim calcState As XlCalculation
    Dim evtState As Boolean
    Dim alertState As Boolean
    Dim screenState As Boolean
    
    calcState = Application.Calculation
    evtState = Application.EnableEvents
    alertState = Application.DisplayAlerts
    screenState = Application.ScreenUpdating
    
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    For Each Item In ActiveItems
        If Item.HasData <> True Or Item.Caption = "(空白)" Or Item.Caption = "-" Or Item.Caption = "" Then Exit For
        If Item.Caption <> "テスト" Then
            Application.StatusBar = Item.Caption & " 処理中..."
            sli.VisibleSlicerItemsList = Array(Item.Name)
            Application.Calculate
            
            ' 全named rangeを遅延保存モードで処理
            切り出し_定義_抽出 teigifilter, Item.Caption, fitpage
        End If
    Next
    
    ' 全ブックを一括保存
    SavePendingBooks
    
    Application.Calculation = calcState
    Application.EnableEvents = evtState
    Application.DisplayAlerts = alertState
    Application.ScreenUpdating = screenState
    Application.StatusBar = ""
    
    ResetMaked
End Sub

' === 切り出し_定義_抽出（保存なし）===
Sub 切り出し_定義_抽出(Optional teigifilter = "output", Optional itemname = "", Optional fitpage = False)
    Dim FromBook As Workbook
    Set FromBook = ActiveWorkbook
    FromBook.Activate
    
    Dim T As Name
    For Each T In FromBook.Names
        If InStr(1, T.Name, teigifilter) > 0 Then
            切り出し_定義 T, itemname, fitpage, True ' skipSave=True
        End If
    Next
End Sub

' === 切り出し_定義とスライサー_シート内_一括版 ===
Public Sub 切り出し_定義とスライサー_シート内_一括(Optional teigifilter = "output", Optional fitpage = False)
    Dim ActiveItems As New Collection
    Dim sourcebook As Workbook
    Dim sli As SlicerCache
    Dim Item As SlicerItem
    Dim fromsheet As Worksheet
    
    Set sourcebook = ActiveWorkbook
    Set fromsheet = ActiveSheet
    Set sli = GetSlicer(Selection.Name)
    
    For Each Item In sli.SlicerCacheLevels(1).SlicerItems
        If Item.Selected = True And Item.HasData = True Then ActiveItems.Add Item
    Next
    
    ' 遅延保存初期化
    InitPendingBooks
    
    Dim calcState As XlCalculation
    Dim evtState As Boolean
    Dim alertState As Boolean
    Dim screenState As Boolean
    
    calcState = Application.Calculation
    evtState = Application.EnableEvents
    alertState = Application.DisplayAlerts
    screenState = Application.ScreenUpdating
    
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    For Each Item In ActiveItems
        If Item.HasData <> True Or Item.Caption = "(空白)" Or Item.Caption = "-" Or Item.Caption = "" Then Exit For
        If Item.Caption <> "テスト" Then
            Application.StatusBar = Item.Caption & " 処理中..."
            sli.VisibleSlicerItemsList = Array(Item.Name)
            Application.Calculate
            
            ' シート内のnamed rangeを遅延保存モードで処理
            Dim T As Name
            For Each T In FromBook.Names
                If InStr(1, T.Name, teigifilter) > 0 Then
                    If T.RefersToRange.Parent Is fromsheet Then
                        切り出し_定義 T, Item.Caption, fitpage, True ' skipSave=True
                    End If
                End If
            Next
        End If
    Next
    
    ' 全ブックを一括保存
    SavePendingBooks
    
    Application.Calculation = calcState
    Application.EnableEvents = evtState
    Application.DisplayAlerts = alertState
    Application.ScreenUpdating = screenState
    Application.StatusBar = ""
    
    ResetMaked
End Sub

' === 以下、既存のサブルーチン（省略）===
' 切り出し_全体、切り出し_選択、
' 切り出し_定義_全体、切り出し_定義_シート内、
' 雛形適応、DirectCall、条件書式の正規化、書式を正規化、
' GetFormatKey、ApplyFormatFromKey、MergeRanges、シートコピー
' 図形スナップ、GetSlicer、MakeDir、ResetMaked
' は既存のまま維持
