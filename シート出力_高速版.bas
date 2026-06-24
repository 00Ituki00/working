' シート出力高速化版
' ハイブリッド方式：データは配列一括、書式はPasteSpecial

Global maked As New Collection

' === メイン関数：高速版 ===
Public Function 切り出し高速(FromBook As Workbook, fromsheet As Worksheet, FromRange As Range, ToBook As Workbook, ToSheet As Worksheet, Optional selectiononly = False, Optional fitpage = False) As Workbook
    Dim pt As PivotTable
    Dim lo As ListObject
    Dim sh As Shape
    Dim chartObj As ChartObject
    Dim dataArr As Variant
    Dim r As Long, c As Long
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
    ' テーブルのスタイルを実書式として保持するため、テーブル構造は維持せず値のみをコピー
    Dim tblRanges As Collection
    Set tblRanges = New Collection
    
    For Each lo In fromsheet.ListObjects
        If Not Intersect(lo.Range, FromRange) Is Nothing Then
            On Error Resume Next
            ' テーブル範囲を保存（後で個別にコピー）
            tblRanges.Add lo.Range
            On Error GoTo 0
        End If
    Next lo
    DoEvents
    
    ' === Step 3: データを配列で一括コピー ===
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
    
    ' === Step 4: 一括書き込み ===
    Dim ToRange As Range
    Set ToRange = ToSheet.Range(ToSheet.Cells(FromRange.Row, FromRange.Column), _
                                ToSheet.Cells(FromRange.Row + FromRange.Rows.Count - 1, _
                                              FromRange.Column + FromRange.Columns.Count - 1))
    
    ' 文字列として書き込んで桁落ち防止
    ToRange.NumberFormat = "@"
    ToRange.Value2 = dataArr
    DoEvents
    
    ' === Step 5: 書式をPasteSpecialで一括適用 ===
    On Error GoTo recopy1
    fromsheet.Activate: FromRange.Copy: ToSheet.Activate
    ToRange.PasteSpecial Paste:=xlPasteFormats
    On Error GoTo 0
    DoEvents
    
    ' === Step 5.5: テーブルのスタイルを個別にコピー ===
    ' テーブル範囲は通常セルとしてコピーすることで、テーブルスタイルを実書式として適用
    Dim tblRange As Range
    For Each tblRange In tblRanges
        On Error Resume Next
        ' テーブル範囲をコピー（テーブル構造は維持せず、表示書式のみ）
        tblRange.Copy
        ' 貼り付け先を計算
        Dim tblDestRange As Range
        Set tblDestRange = ToSheet.Range(tblRange.Address)
        tblDestRange.PasteSpecial Paste:=xlPasteAllUsingSourceTheme
        ' 値は上書き（テーブル構造を排除）
        tblDestRange.Value = tblDestRange.Value
        On Error GoTo 0
    Next tblRange
    DoEvents
    
    ' === Step 6: 条件書式を正規化 ===
    書式を正規化 ToRange
    DoEvents
    
    ' === Step 7: 図形を画像としてコピー ===
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
    
    Set 切り出し高速 = ToBook
    Exit Function
    
recopy1:
    fromsheet.Activate: FromRange.Copy: ToSheet.Activate: DoEvents: Resume
End Function

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
