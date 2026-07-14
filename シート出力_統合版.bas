'外部呼出し用
'Sub 切り出し()
'    Application.Run "個人用マクロ.xlam!切り出し_全体"
'End Sub
Global maked As New Collection
Global pendingBooks As Object ' 遅延保存用ブック管理

Public Sub 切り出し_定義とスライサー_シート内_call2(): 切り出し_定義とスライサー_シート内 ("output時間外集計16_残業手当推移グラフ"): End Sub
Public Sub 切り出し_定義_シート内_call2(): 切り出し_定義_シート内 ("output時間外集計16_残業手当推移グラフ"): End Sub

' === 遅延保存用：初期化 ===
Private Sub InitPendingBooks()
    Set pendingBooks = CreateObject("Scripting.Dictionary")
End Sub

' === 遅延保存用：一括保存実行 ===
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

Public Function 切り出し(FromBook As Workbook, fromsheet As Worksheet, FromRange As Range, ToBook As Workbook, ToSheet As Worksheet, Optional selectiononly = False, Optional fitpage = False) As Workbook
'対象の範囲を対象のシートに再現出力
    Dim pt As PivotTable
    Dim ptRange As Range
    Dim ptRanges As Collection
    Dim cell As Range
    Dim dataArr As Variant
    Dim numFmtArr As Variant
    Dim colWidths() As Double
    Dim r As Long, c As Long
    
    STVBA
    ToBook.ApplyTheme ("C:\Users\h_ikegami\AppData\Roaming\Microsoft\Templates\Document Themes\default.thmx")
    DoEvents

    Dim firstRow As Long, lastRow As Long
    firstRow = FromRange.Row
    lastRow = FromRange.Row + FromRange.Rows.Count - 1
    
    ToSheet.Cells.EntireRow.AutoFit
    DoEvents
    
    ' === 高速化：データを配列で一括コピー ===
    dataArr = FromRange.Value2
    ' エラーのみクリア（Emptyはそのまま）
    For r = 1 To UBound(dataArr, 1)
        For c = 1 To UBound(dataArr, 2)
            If IsError(dataArr(r, c)) Then dataArr(r, c) = ""
        Next c
    Next r
    
    ' NumberFormat・列幅を配列で一括取得
    numFmtArr = FromRange.NumberFormat
    ReDim colWidths(1 To FromRange.Columns.Count)
    For c = 1 To FromRange.Columns.Count
        colWidths(c) = fromsheet.Columns(FromRange.Column + c - 1).ColumnWidth
    Next c
    
    ' 一括書き込み
    Dim ToRange As Range
    Set ToRange = ToSheet.Range( _
        ToSheet.Cells(FromRange.Row, FromRange.Column), _
        ToSheet.Cells(FromRange.Row + FromRange.Rows.Count - 1, _
                      FromRange.Column + FromRange.Columns.Count - 1) _
    )
    ToRange.NumberFormat = "@"
    ToRange.Value2 = dataArr
    ToRange.NumberFormat = numFmtArr
    ' 列幅一括適用
    For c = 1 To UBound(colWidths)
        ToSheet.Columns(FromRange.Column + c - 1).ColumnWidth = colWidths(c)
    Next c
    DoEvents
    
    ' 書式をPasteSpecialで一括適用
    On Error GoTo recopy1
    fromsheet.Activate: FromRange.Copy: ToSheet.Activate
    ToRange.PasteSpecial Paste:=xlPasteFormats
    On Error GoTo 0
    DoEvents
    
    '条件書式を実書式へ反映
    書式を正規化 ToRange
    DoEvents
    
    ' ピボットテーブルの再現
    Set ptRanges = New Collection
    For Each pt In fromsheet.PivotTables
    If Not Intersect(pt.TableRange2, FromRange) Is Nothing Then
         'ピボットテーブルの列フィルタ（1行目）を除く
        If pt.ShowValuesRow = True Then
            Set ptRange = pt.TableRange1.Offset(1, 0).Resize(pt.TableRange1.Rows.Count - 1)
        Else
            Set ptRange = pt.TableRange1 '全体を含めると書式コピーできない
        End If
        ' 値と書式をコピー
        ptRange.Copy
        ToSheet.Cells(ptRange.Cells(1, 1).Row, ptRange.Cells(1, 1).Column).PasteSpecial Paste:=xlPasteValuesAndNumberFormats
        ToSheet.Cells(ptRange.Cells(1, 1).Row, ptRange.Cells(1, 1).Column).PasteSpecial Paste:=xlPasteFormats
        If pt.PageFields.Count > 0 Then 'フィルター行がある場合はコピー
            pt.PageRange.Copy
            ToSheet.Cells(pt.PageRange.Cells(1, 1).Row, pt.PageRange.Cells(1, 1).Column).PasteSpecial Paste:=xlPasteValuesAndNumberFormats
            ToSheet.Cells(pt.PageRange.Cells(1, 1).Row, pt.PageRange.Cells(1, 1).Column).PasteSpecial Paste:=xlPasteFormats
        End If
        End If
    Next pt
    DoEvents
    
    If fitpage = True Then 不要な行を非表示
    
    'シートのズームと固定をコピー
    ToSheet.Activate
    ToSheet.Cells(1).Select
    ToSheet.Parent.Windows(1).Zoom = fromsheet.Parent.Windows(1).Zoom
    fromsheet.Activate
    If fromsheet.Parent.Windows(1).FreezePanes = True Then
        ToSheet.Activate
        ToSheet.Parent.Windows(1).SplitVertical = fromsheet.Parent.Windows(1).SplitVertical
        ToSheet.Parent.Windows(1).SplitHorizontal = fromsheet.Parent.Windows(1).SplitHorizontal
        ToSheet.Application.ActiveWindow.FreezePanes = True
    End If
    
    'シート内の画像をコピー
    Dim sh As Shape
    Dim copyRetry As Integer
    Dim copySuccess As Boolean
    Dim shapesToCopy As New Collection
    
    ' コピー対象図形を事前収集（左上セルがFromRange内にあるもののみ）
    For Each sh In fromsheet.Shapes
        If Not Intersect(sh.TopLeftCell, FromRange) Is Nothing Then
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
            Application.CutCopyMode = False
            ' Copy直後にDoEvents（クリップボード反映待ち）
            shp.Copy
            DoEvents
            ToSheet.Activate
            ToSheet.PasteSpecial Format:=0
            ' Paste直後にDoEvents（図形配置反映待ち）
            DoEvents
            If Err.Number = 0 Then
                Dim newShape As Shape
                Set newShape = ToSheet.Shapes(ToSheet.Shapes.Count)
                newShape.Top = ToSheet.Range(shp.TopLeftCell.Address).Top
                newShape.Left = ToSheet.Range(shp.TopLeftCell.Address).Left
                If newShape.TopLeftCell.Address <> shp.TopLeftCell.Address Then Stop
                newShape.SetShapesDefaultProperties
                copySuccess = True
            End If
            If Not copySuccess Then
                copyRetry = copyRetry + 1
                ' エラー時のみDoEvents（クリップボード回復待ち）
                If copyRetry < 3 Then DoEvents
            End If
        Loop
        If ToSheet.Shapes.Count > 0 Then 図形スナップ ToSheet.Shapes(ToSheet.Shapes.Count)
    Next shp
    On Error GoTo 0

    '名前定義をコピー
    On Error Resume Next
    Dim nr As Range
    Dim n As Name, ton As Name
    
    For Each n In FromBook.Names
        If n.Visible Then
            Set nr = Nothing
            On Error Resume Next
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
    Next
    DoEvents
    
    '選択範囲外のセルを削除
    If selectiononly Then
        ToSheet.Range(FromRange.Columns(FromRange.Columns.Count + 1).Address, ToSheet.Columns(99999).Address).Delete '右
        ToSheet.Range(FromRange.Rows(FromRange.Rows.Count + 1).Address, ToSheet.Rows(99999).Address).Delete '下
        If FromRange.Cells(1, 1).Column - 1 > 0 Then ToSheet.Range(ToSheet.Columns(1), ToSheet.Columns(FromRange.Cells(1, 1).Column - 1)).Delete '左
        If FromRange.Cells(1, 1).Row - 1 > 0 Then ToSheet.Range(ToSheet.Rows(1), ToSheet.Rows(FromRange.Cells(1, 1).Row - 1)).Delete '上
        
        ' 範囲外の図形も削除
        On Error Resume Next
        Dim chkShape As Shape
        Dim delShapes As New Collection
        For Each chkShape In ToSheet.Shapes
            If Intersect(chkShape.TopLeftCell, ToRange) Is Nothing Then
                delShapes.Add chkShape
            End If
        Next chkShape
        Dim delShp As Variant
        For Each delShp In delShapes
            delShp.Delete
        Next delShp
        On Error GoTo 0
    End If
    
    ToSheet.Cells(1, 1).Select
    ToBook.Sheets(1).Activate
    maked.Add ToBook
    RFVBA
    Set 切り出し = ToBook
    Exit Function

recopy:
    DoEvents: Resume
recopy1:
    fromsheet.Activate: FromRange.Copy: ToSheet.Activate: DoEvents: Resume
End Function

Sub 切り出し_全体() '現在シートを切り出し
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

Sub 切り出し_選択() '選択範囲を切り出し
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

' === 切り出し_定義（遅延保存対応版）===
Sub 切り出し_定義(T As Variant, Optional itemname = "", Optional fitpage = False, Optional skipSave = False)
    Dim FromBook As Workbook, ToBook As Workbook
    Dim fromsheet As Worksheet, ToSheet As Worksheet
    Dim FromRange As Range
    Dim comm As String, topath As String, toBookName As String, toSheetName As String
    Dim selectiononly As Boolean
    Dim p As Long, q As Long, vn As String
    
    STVBA
    Set FromBook = T.RefersToRange.Parent.Parent
    FromBook.Activate
    comm = T.Comment
    'commの変数処理
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
     If InStr(comm, "YYMMDD") > 0 Then comm = Replace(comm, "YYYYMM", Format(Date, "YYMMDD"))
     If InStr(comm, "YYMM") > 0 Then comm = Replace(comm, "YYMM", Format(Date, "YYMM"))
     If InStr(comm, "itemname") > 0 Then comm = Replace(comm, "itemname", itemname) '可変フォルダ
     If InStr(comm, "M-1月") > 0 Then comm = Replace(comm, "M-1月", Month(Date) - 1 & "月")
     If InStr(comm, "M月") > 0 Then comm = Replace(comm, "M月", Month(Date) & "月")
   
     comm = Split(comm, ",")
     topath = FromBook.path & "\" & comm(0): If Right(topath, 1) <> "\" Then topath = topath & "\"     'コメントから設定 AddPathは連続切り出し用
     toBookName = comm(1): If InStr(toBookName, ".xlsx") = 0 Then toBookName = toBookName & ".xlsx"
     toSheetName = comm(2)
    
     selectiononly = True
     If UBound(comm) >= 3 Then If comm(3) = "false" Then selectiononly = False

     If InStr(toBookName, "itemname") > 0 Then toBookName = Replace(toBookName, "itemname", itemname) '可変ファイル名
     
     ' === 遅延保存：既存のブックを再利用 ===
     Dim bookKey As String
     bookKey = topath & toBookName
     
     If Not pendingBooks Is Nothing Then
         If pendingBooks.Exists(bookKey) Then
             Set ToBook = pendingBooks(bookKey)
         End If
     End If
     
     If ToBook Is Nothing Then
         'すでに同名ブックを開いていれば取得、無ければ新規
         For Each book In Workbooks
             If book.Name = toBookName Then Set ToBook = Workbooks(toBookName): Exit For
         Next
         If ToBook Is Nothing Then Set ToBook = Workbooks.Add
         If Not pendingBooks Is Nothing Then
             pendingBooks.Add bookKey, ToBook
         End If
     End If
     
     For Each sh In ToBook.Sheets
         If sh.Name = toSheetName Then Set ToSheet = sh: Exit For
     Next
     If ToSheet Is Nothing Then Set ToSheet = ToBook.Worksheets.Add(After:=ToBook.Worksheets(ToBook.Worksheets.Count)): ToSheet.Name = toSheetName
     Set fromsheet = T.RefersToRange.Parent
     Set FromRange = Intersect(T.RefersToRange, fromsheet.UsedRange)
     MakeDir topath '出力予定フォルダ作成
     FromBook.Activate
     fromsheet.Activate
    Set ToBook = 切り出し(FromBook, fromsheet, FromRange, ToBook, ToSheet, selectiononly, fitpage)

    ' === 保存制御 ===
    If Not skipSave Then
        STVBA
        On Error Resume Next
        If Not ToBook.Sheets("Sheet1") Is Nothing Then ToBook.Sheets("Sheet1").Delete
        On Error GoTo 0
        ToBook.SaveAs topath & toBookName
        RFVBA
        Sleep (2000)
        FromBook.Activate
    End If
End Sub

Sub 切り出し_定義_全体_call(): 切り出し_定義_全体: End Sub
Sub 切り出し_定義_全体(Optional teigifilter = "output", Optional fitpage = False) 'ブック内の名前定義からすべて切り出し
'path,book,sheet

    STVBA
    Set FromBook = ActiveWorkbook
    FromBook.Activate
    For Each T In FromBook.Names    '名前定義一覧を取得
        If InStr(1, T.Name, teigifilter) > 0 Then    'マクロ用の名前の場合
          切り出し_定義 T, itemname, fitpage
            DoEvents
        End If
    Next
    RFVBA
    ResetMaked
End Sub

Sub 切り出し_定義_シート内_Call(): 切り出し_定義_シート内: End Sub
Sub 切り出し_定義_シート内(Optional teigifilter = "output", Optional fitpage = False) 'シート内の名前定義からすべて切り出し
'path,book,sheet
    STVBA
    Set FromBook = ActiveWorkbook
    Set fromsheet = ActiveSheet
    FromBook.Activate
    If teigifilter = "" Then teigifilter = "output"
    For Each T In FromBook.Names    '名前定義一覧を取得
        If InStr(1, T.Name, teigifilter) > 0 Then    'マクロ用の名前の場合
        If T.RefersToRange.Parent Is fromsheet Then
            切り出し_定義 T, itemname, fitpage
            DoEvents
        End If
        End If
    Next
    RFVBA
    ResetMaked
End Sub

' === 旧版：互換性維持 ===
Public Sub 切り出し_定義とスライサー_call(): 切り出し_定義とスライサー: End Sub
Sub 切り出し_定義とスライサー(Optional teigifilter = "output", Optional fitpage = False)
    Dim ActiveItems As Collection
    Set ActiveItems = New Collection
    Set sourcebook = ActiveWorkbook
    Set sli = GetSlicer(Selection.Name)
    For Each Item In sli.SlicerCacheLevels(1).SlicerItems
        If Item.Selected = True And Item.HasData = True Then ActiveItems.Add Item
    Next
    For Each Item In ActiveItems
        If Item.HasData <> True Or Item.Caption = "(空白)" Or Item.Caption = "-" Or Item.Caption = "" Then: Exit For
        If Item.Caption <> "テスト" Then
            Application.StatusBar = (Item.Caption & " 処理中...")
            sli.VisibleSlicerItemsList = Array(Item.Name)
            Application.Calculate
            切り出し_定義_全体 teigifilter, fitpage
        End If
    Next
    ResetMaked
    Application.StatusBar = ""
End Sub

Public Sub 切り出し_定義とスライサー_シート内_call(): 切り出し_定義とスライサー_シート内: End Sub
Sub 切り出し_定義とスライサー_シート内(Optional teigifilter = "output", Optional fitpage = False)
    Dim ActiveItems As Collection
    Set ActiveItems = New Collection
    Set sourcebook = ActiveWorkbook
    Set sli = GetSlicer(Selection.Name)
    For Each Item In sli.SlicerCacheLevels(1).SlicerItems
        If Item.Selected = True And Item.HasData = True Then ActiveItems.Add Item
    Next
    For Each Item In ActiveItems
        If Item.HasData <> True Or Item.Caption = "(空白)" Or Item.Caption = "-" Or Item.Caption = "" Then: Exit For
        If Item.Caption <> "テスト" Then
            Application.StatusBar = (Item.Caption & " 処理中...")
            sli.VisibleSlicerItemsList = Array(Item.Name)
            Application.Calculate
            切り出し_定義_シート内 teigifilter, fitpage
        End If
    Next
    ResetMaked
    Application.StatusBar = ""
End Sub

' === 高速化版：スライサー項目単位で一括保存 ===
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
            Dim T As Name
            For Each T In sourcebook.Names
                If InStr(1, T.Name, teigifilter) > 0 Then
                    切り出し_定義 T, Item.Caption, fitpage, True ' skipSave=True
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
            For Each T In sourcebook.Names
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

Sub 雛形適応() 'シートの設定を雛形ブックから同名のシートへコピーする
    Set ToBook = ActiveWorkbook
    
    For Each bk In Workbooks
        If InStr(bk.Name, "雛形") > 0 Then Set FromBook = bk: Exit For
    Next
    STVBA
    Application.PrintCommunication = False
    For Each sh In ToBook.Worksheets
    For Each sh2 In FromBook.Worksheets
        If sh.Name = sh2.Name Then
        sh2.Activate
        v = ActiveWindow.View
        sh.Activate
    
            sh.Tab.Color = sh2.Tab.Color
            With sh.PageSetup
                ' 1. 用紙設定関連（最初に設定）
                .PaperSize = sh2.PageSetup.PaperSize
                .Orientation = sh2.PageSetup.Orientation
                .FirstPageNumber = sh2.PageSetup.FirstPageNumber
                .Order = sh2.PageSetup.Order
                '.PrintQuality = sh2.PageSetup.PrintQuality
                ' 2. 拡大縮小設定（Zoom→Fitの順）
                .Zoom = sh2.PageSetup.Zoom
                .FitToPagesWide = sh2.PageSetup.FitToPagesWide
                .FitToPagesTall = sh2.PageSetup.FitToPagesTall
            
                ' 3. 印刷エリアや余白など
                .PrintArea = sh2.PageSetup.PrintArea
                .LeftMargin = sh2.PageSetup.LeftMargin
                .RightMargin = sh2.PageSetup.RightMargin
                .TopMargin = sh2.PageSetup.TopMargin
                .BottomMargin = sh2.PageSetup.BottomMargin
                .HeaderMargin = sh2.PageSetup.HeaderMargin
                .FooterMargin = sh2.PageSetup.FooterMargin
            
                ' 4. 表示位置設定
                .CenterHorizontally = sh2.PageSetup.CenterHorizontally
                .CenterVertically = sh2.PageSetup.CenterVertically
            
                ' 5. 印刷オプション
                .BlackAndWhite = sh2.PageSetup.BlackAndWhite
                .Draft = sh2.PageSetup.Draft
                .PrintComments = sh2.PageSetup.PrintComments
            
            End With
            Cells(1, 1).Select
            Application.GoTo Cells(1, 1), True
            ActiveWindow.View = v
            DoEvents
            Exit For
        End If
    Next
    Next
Application.PrintCommunication = True
    On Error Resume Next
    For i = 1 To FromBook.Sheets.Count
        ToBook.Sheets(FromBook.Sheets(i).Name).Move After:=ToBook.Sheets(ToBook.Sheets.Count)
    Next i
    For Each sh In ToBook.Sheets
    HIT = False
        For Each sh2 In FromBook.Sheets
         If sh.Name = sh2.Name Then HIT = True
        Next
    Next
    Application.ScreenUpdating = True
    On Error GoTo 0
    ToBook.Sheets(1).Activate
    DoEvents
    ToBook.Save
    ToBook.Close
    RFVBA
End Sub

Sub DirectCall()
    DirectCaller.Show False
End Sub

Sub 条件書式の正規化()
   書式を正規化 Selection
End Sub

Sub 書式を正規化(rng As Range)
    Dim cell As Range
    Dim dict As Object
    Dim key As Variant
    Dim displayKey As String, currentKey As String
    Dim styleRange As Range
    Dim fcRange As Range

    Set dict = CreateObject("Scripting.Dictionary")

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    '条件付き書式のあるセルだけを対象範囲として集める
    For Each cell In rng
        If cell.FormatConditions.Count > 0 Then
            If fcRange Is Nothing Then
                Set fcRange = cell
            Else
                Set fcRange = Union(fcRange, cell)
            End If

            displayKey = GetFormatKey(cell.DisplayFormat)
            currentKey = GetFormatKey(cell)

            If displayKey <> currentKey Then
                If Not dict.Exists(displayKey) Then
                    Set dict(displayKey) = cell
                Else
                    Set dict(displayKey) = Union(dict(displayKey), cell)
                End If
            End If
        End If
    Next cell

    DoEvents

    '書式キーごとに一括反映
    For Each key In dict.Keys
        Set styleRange = dict(key)
        ApplyFormatFromKey styleRange, CStr(key)
    Next key

    '条件付き書式があったセルだけ削除
    If Not fcRange Is Nothing Then
        fcRange.FormatConditions.Delete
    End If

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub

Function GetFormatKey(obj As Object) As String
    Dim f As String
    Dim b As Variant
    Dim border As Border
    Dim borderTypes As Variant
    borderTypes = Array(xlEdgeTop, xlEdgeBottom, xlEdgeLeft, xlEdgeRight, xlInsideHorizontal, xlInsideVertical)

    With obj
        f = Join(Array( _
            .Interior.Color, _
            .NumberFormat, _
            .Font.Name, _
            .Font.Size, _
            .Font.Color, _
            .Font.Bold, _
            .Font.Italic, _
            .Font.Underline, _
            .HorizontalAlignment, _
            .VerticalAlignment, _
            .WrapText, _
            .ShrinkToFit _
        ), "|")

        ' 罫線情報（各方向）
        For Each b In borderTypes
            With .Borders(b)
                f = f & "|" & b & ":" & .LineStyle & "," & .Color & "," & .Weight
            End With
        Next b
    End With

    GetFormatKey = f
End Function

Sub ApplyFormatFromKey(rng As Range, key As Variant)
    Dim parts() As String
    Dim i As Long
    Dim borderTypes As Variant
    Dim bIndex As Long

    borderTypes = Array(xlEdgeTop, xlEdgeBottom, xlEdgeLeft, xlEdgeRight, xlInsideHorizontal, xlInsideVertical)

    parts = Split(key, "|")
    On Error Resume Next
    With rng
        If CLng(parts(0)) <> 16777215 Then .Interior.Color = CLng(parts(0))
        .NumberFormat = parts(1)

        With .Font
            .Name = parts(2)
            .Size = Val(parts(3))
            .Color = CLng(parts(4))
            .Bold = CBool(parts(5))
            .Italic = CBool(parts(6))
            .Underline = parts(7)
        End With

        .HorizontalAlignment = parts(8)
        .VerticalAlignment = parts(9)
        .WrapText = CBool(parts(10))
        .ShrinkToFit = CBool(parts(11))
    End With

    ' 罫線の適用
    For i = 0 To 5
        Dim borderPart() As String
        Dim borderVals() As String

        borderPart = Split(parts(12 + i), ":")
        borderVals = Split(borderPart(1), ",")

        With rng.Borders(borderTypes(i))
            If CLng(borderVals(0)) <> -4142 Then
            .LineStyle = CLng(borderVals(0))
            .Color = CLng(borderVals(1))
            .Weight = CLng(borderVals(2))
            End If
        End With
    Next i
    On Error GoTo 0
End Sub

Sub MergeRanges(MergeCells() As Range) 'Range配列を個別にマージ
    Dim muluticells As String
    muluticells = ""
    For i = 0 To UBound(MergeCells) - 1
       muluticells = muluticells & MergeCells(i).Address & ","
       If Len(muluticells) >= 200 Then '256文字制限対応
        muluticells = Left(muluticells, Len(muluticells) - 1)
        Range(muluticells).Merge: Range(muluticells).HorizontalAlignment = xlCenter
        Range(muluticells).VerticalAlignment = xlCenter
        muluticells = ""
    End If
    Next
End Sub

Sub シートコピー(wb1, wb2, targetRange, targetsheet)
    wb1.Activate
    Columns(targetRange).Select
    Application.CutCopyMode = False
    Selection.Copy: DoEvents
    wb2.Activate
    Sheets(targetsheet).Activate
    Range("A:A").Select
    ActiveSheet.Paste: DoEvents: Cells(1, 1).Select
    Cells(1, 1).Select
End Sub

Public Sub AddRange(ByRef ranges() As Range, ByVal newRange As Range)
    Dim count As Long

    On Error GoTo InitArray
    count = UBound(ranges) + 1
    ReDim Preserve ranges(count)
    Set ranges(count) = newRange
    Exit Sub

InitArray:
    ReDim ranges(0)
    Set ranges(0) = newRange
End Sub

Sub 図形位置のセル合わせ()
Dim sh As Object
    For Each sh In ActiveSheet.Shapes
        図形スナップ sh
    Next
End Sub

Private Sub 図形スナップ(shp As Shape)
    Dim ws As Worksheet
    Set ws = shp.Parent

    Dim topLeftCell As Range, bottomRightCell As Range
    Set topLeftCell = shp.TopLeftCell
    Set bottomRightCell = shp.BottomRightCell

    Dim minRow As Long, maxRow As Long
    Dim minCol As Long, maxCol As Long

    minRow = Application.Max(1, topLeftCell.Row - 1)
    maxRow = Application.Min(ws.Rows.Count, bottomRightCell.Row + 1)
    minCol = Application.Max(1, topLeftCell.Column - 1)
    maxCol = Application.Min(ws.Columns.Count, bottomRightCell.Column + 1)

    Dim leftPos As Double, topPos As Double
    Dim rightPos As Double, bottomPos As Double
    leftPos = shp.Left
    topPos = shp.Top
    rightPos = shp.Left + shp.Width
    bottomPos = shp.Top + shp.Height

    Dim nearestLeft As Double, nearestTop As Double
    Dim nearestRight As Double, nearestBottom As Double
    Dim minLeftDiff As Double: minLeftDiff = 10000000000#
    Dim minTopDiff As Double: minTopDiff = 10000000000#
    Dim minRightDiff As Double: minRightDiff = 10000000000#
    Dim minBottomDiff As Double: minBottomDiff = 10000000000#

    Dim r As Long, c As Long
    For r = minRow To maxRow
        For c = minCol To maxCol
            With ws.Cells(r, c)
                ' 左端
                If Abs(.Left - leftPos) < minLeftDiff Then
                    minLeftDiff = Abs(.Left - leftPos)
                    nearestLeft = .Left
                End If
                ' 上端
                If Abs(.Top - topPos) < minTopDiff Then
                    minTopDiff = Abs(.Top - topPos)
                    nearestTop = .Top
                End If
                ' 右端
                If Abs(.Left + .Width - rightPos) < minRightDiff Then
                    minRightDiff = Abs(.Left + .Width - rightPos)
                    nearestRight = .Left + .Width
                End If
                ' 下端
                If Abs(.Top + .Height - bottomPos) < minBottomDiff Then
                    minBottomDiff = Abs(.Top + .Height - bottomPos)
                    nearestBottom = .Top + .Height
                End If
            End With
        Next c
    Next r

    ' 図形の位置とサイズを更新
    With shp
        .LockAspectRatio = False
        .Left = nearestLeft
        .Top = nearestTop
        .Width = nearestRight - nearestLeft
        .Height = nearestBottom - nearestTop
    End With
End Sub

Function GetSlicer(name) As SlicerCache
    Dim slcCache As SlicerCache
    Dim slc As Slicer
    Dim isSlicerFound As Boolean

    isSlicerFound = False
    On Error GoTo re
    For Each slcCache In ActiveWorkbook.SlicerCaches
        For Each slc In slcCache.Slicers
re:
            If slc.Shape.Parent Is ActiveSheet Then
                If Split(slc.Caption, " ")(0) = Split(name, " ")(0) Then
                    Set GetSlicer = slc.SlicerCache
                    Exit Function
                End If
            End If
        Next slc
    Next slcCache

    If Not isSlicerFound Then
        MsgBox "指定されたスライサーは見つかりませんでした。"
    End If
End Function

Sub ExMerge(firstRow, Scanname, Targetname)
'項目のマージ状態を他の項目列へ反映
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    lastcolumn = Cells(firstRow, Columns.Count).End(xlToLeft).Column
    For c = 1 To lastcolumn
        If Cells(firstRow, c) = Scanname Then ScanCol = c: Exit For
    Next c
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    For c = 1 To lastcolumn
        For r = firstRow To lastRow
            If Cells(firstRow, c).Value = Targetname Then
                mergecount = Cells(r, ScanCol).MergeArea.Rows.Count
                If mergecount <> 1 Then
                    Cells(r, c).Resize(mergecount, 1).Merge
                    r = r + mergecount - 1
                End If
                Cells(r, c).HorizontalAlignment = xlCenter
                Cells(r, c).VerticalAlignment = xlCenter
            End If
        Next r
        DoEvents
    Next c
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

Sub MergeBook(wb1 As Workbook, wb2 As Workbook)
    wb2.Worksheets(1).Move After:=wb1.Worksheets(1)
    wb1.Worksheets(1).Activate
    wb1.Save
    wb1.Close
End Sub

Public Sub MakeDir(outdir)
    Dim dirpath As String
    dirpath = outdir

    If Dir(dirpath, vbDirectory) = "" Then
        Dim parentPath As String
        parentPath = Left(dirpath, InStrRev(dirpath, "\") - 1)
        If Dir(parentPath, vbDirectory) = "" Then
            Do
                parentPath = Left(parentPath, InStrRev(parentPath, "\") - 1)
                If Dir(parentPath, vbDirectory) <> "" Then Exit Do
            Loop
            Do
                If Dir(parentPath, vbDirectory) = "" Then MkDir Replace(parentPath, "\\", "\")
                On Error GoTo out
                parentPath = parentPath & Mid(dirpath, Len(parentPath) + 1, InStr(Len(parentPath) + 1, dirpath, "\") - Len(parentPath))
            Loop Until Len(parentPath) >= Len(dirpath)
out:
        End If
        dirpath = Replace(dirpath, "\\", "\")
        If Dir(dirpath, vbDirectory) = "" Then MkDir dirpath
    End If
    DoEvents
End Sub

Sub フィルタ解除(sliname As String)
    Set sli2 = GetSlicer(sliname)
    If sli2.FilterCleared = False Then sli2.ClearManualFilter
End Sub

Sub フィルタ設定(sliname As String, sli_select)
    Set sli = GetSlicer(sliname)
    For Each Items In sli.VisibleSlicerItemsList
        If Items = sli_select Then Exit Sub
    Next
    sli.VisibleSlicerItemsList = Array(sli_select)
End Sub

Sub ResetMaked()
    On Error Resume Next
    STVBA
    For Each Item In maked
        Item.Close
    Next
    Set maked = Nothing
    RFVBA
End Sub
