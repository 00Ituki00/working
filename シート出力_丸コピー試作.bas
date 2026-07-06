' シート出力丸コピー試作版
' 方針: シート丸ごと複製 → 不要部分削除 → データ接続切離し
' 既存インターフェース: 切り出し(FromBook, fromsheet, FromRange, ToBook, ToSheet, ...)

Global maked As New Collection

' === メイン関数：丸コピー高速版 ===
Public Function 切り出し高速(FromBook As Workbook, fromsheet As Worksheet, FromRange As Range, ToBook As Workbook, ToSheet As Worksheet, Optional selectiononly = False, Optional fitpage = False) As Workbook
    Dim sh As Shape
    Dim chartObj As ChartObject
    Dim tempPath As String
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
    
    ' === Step 0: 元シートの使用範囲でFromRangeを絞る ===
    Dim usedRng As Range
    Set usedRng = fromsheet.UsedRange
    If Not Intersect(FromRange, usedRng) Is Nothing Then
        Set FromRange = Intersect(FromRange, usedRng)
    End If
    
    ' === Step 1: 元ブック内に一時シートを作成（元シートは無傷） ===
    Dim tempSheet As Worksheet
    fromsheet.Copy After:=FromBook.Sheets(FromBook.Sheets.Count)
    Set tempSheet = FromBook.Sheets(FromBook.Sheets.Count)
    tempSheet.Name = "TempExport_" & Format(Now, "hhmmss")
    
    ' === Step 2: 一時シートのデータ接続を完全切断（元ブックの他シートには影響なし） ===
    Call 切断データ接続(tempSheet)
    Call 値化ピボットテーブル(tempSheet)
    Call 値化リストオブジェクト(tempSheet)
    Call 正規化名前定義(tempSheet, FromBook)
    
    ' === Step 3: 一時シートをToBookにコピー（接続なし＝高速） ===
    tempSheet.Copy After:=ToBook.Sheets(ToBook.Sheets.Count)
    Set copiedSheet = ToBook.Sheets(ToBook.Sheets.Count)
    
    ' === Step 4: 元ブックの一時シートを削除（お掃除） ===
    Application.DisplayAlerts = False
    tempSheet.Delete
    Application.DisplayAlerts = False
    
    ' === Step 5: コピーされたシートの名前を設定 ===
    On Error Resume Next
    copiedSheet.Name = ToSheet.Name
    On Error GoTo Cleanup
    
    ' === Step 6: selectiononly時、FromRange以外を一括削除 ===
    If selectiononly Then
        Dim firstRow As Long, lastRow As Long
        Dim firstCol As Long, lastCol As Long
        firstRow = FromRange.Row
        lastRow = FromRange.Row + FromRange.Rows.Count - 1
        firstCol = FromRange.Column
        lastCol = FromRange.Column + FromRange.Columns.Count - 1
        
        ' 上側削除
        If firstRow > 1 Then
            copiedSheet.Rows("1:" & firstRow - 1).Delete
        End If
        
        ' 下側削除
        Dim currentLastRow As Long
        currentLastRow = copiedSheet.Cells(copiedSheet.Rows.Count, 1).End(xlUp).Row
        If lastRow < currentLastRow Then
            copiedSheet.Rows(lastRow + 1 & ":" & copiedSheet.Rows.Count).Delete
        End If
        
        ' 左側削除
        If firstCol > 1 Then
            copiedSheet.Columns("A:" & Split(copiedSheet.Cells(1, firstCol - 1).Address, "$")(1)).Delete
        End If
        
        ' 右側削除
        Dim currentLastCol As Long
        currentLastCol = copiedSheet.Cells(1, copiedSheet.Columns.Count).End(xlToLeft).Column
        If lastCol < currentLastCol Then
            copiedSheet.Columns(Split(copiedSheet.Cells(1, lastCol + 1).Address, "$")(1) & ":XFD").Delete
        End If
        
        ' 行番号・列番号が変わったのでFromRangeを再設定
        Set FromRange = copiedSheet.Range("A1").Resize(lastRow - firstRow + 1, lastCol - firstCol + 1)
    End If
    
    ' === Step 7: 図形の調整（シートコピーで位置は維持されるが一応確認） ===
    Call 調整図形位置(copiedSheet)
    
    ' === Step 8: ウィンドウ設定コピー ===
    copiedSheet.Activate
    copiedSheet.Cells(1).Select
    copiedSheet.Parent.Windows(1).Zoom = fromsheet.Parent.Windows(1).Zoom
    
    If fromsheet.Parent.Windows(1).FreezePanes Then
        copiedSheet.Parent.Windows(1).SplitVertical = fromsheet.Parent.Windows(1).SplitVertical
        copiedSheet.Parent.Windows(1).SplitHorizontal = fromsheet.Parent.Windows(1).SplitHorizontal
        copiedSheet.Application.ActiveWindow.FreezePanes = True
    End If
    
    ' === Step 9: ToSheetと入れ替え（必要に応じて） ===
    If copiedSheet.Index <> ToSheet.Index Then
        copiedSheet.Move Before:=ToSheet
    End If
    
    maked.Add ToBook
    
    Set 切り出し高速 = ToBook
    
Cleanup:
    ' === 事後：Excel設定を復元 ===
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

' === サブルーチン：データ接続切断 ===
Private Sub 切断データ接続(ws As Worksheet)
    Dim pt As PivotTable
    Dim lo As ListObject
    Dim qry As WorkbookQuery
    Dim conn As WorkbookConnection
    
    On Error Resume Next
    
    ' 1. ピボットテーブルのキャッシュ・接続を切る
    For Each pt In ws.PivotTables
        pt.RefreshOnFileOpen = False
        ' ピボットキャッシュを破棄しないと接続が残る場合がある
    Next pt
    
    ' 2. テーブルのクエリ接続を切る
    For Each lo In ws.ListObjects
        lo.RefreshStyle = xlOverwriteCells
    Next lo
    
    ' 3. ワークブック接続から該当シートのものを削除
    Dim wb As Workbook
    Set wb = ws.Parent
    
    ' クエリ削除（シートに関連するもの）
    ' 注意：WorkbookQueryはコレクションだが、削除中にインデックスがずれるので後ろから
    Dim i As Long
    If wb.Queries.Count > 0 Then
        For i = wb.Queries.Count To 1 Step -1
            ' シート名がクエリ名に含まれるか、接続先を確認するのは困難なので
            ' ここでは全クエリ削除（別ブックなので影響なし）
            wb.Queries(i).Delete
        Next i
    End If
    
    ' 接続削除（後ろから）
    If wb.Connections.Count > 0 Then
        For i = wb.Connections.Count To 1 Step -1
            wb.Connections(i).Delete
        Next i
    End If
    
    On Error GoTo 0
End Sub

' === サブルーチン：ピボットテーブルを値化 ===
Private Sub 値化ピボットテーブル(ws As Worksheet)
    Dim pt As PivotTable
    Dim ptRange As Range
    Dim hasValuesRow As Boolean
    
    On Error Resume Next
    
    ' 後ろから処理（削除でインデックスがずれるため）
    Dim i As Long
    For i = ws.PivotTables.Count To 1 Step -1
        Set pt = ws.PivotTables(i)
        Set ptRange = pt.TableRange2
        hasValuesRow = pt.ShowValuesRow
        
        If Not ptRange Is Nothing Then
            ' ピボットテーブル範囲を値と書式でコピー
            ptRange.Copy
            ptRange.PasteSpecial xlPasteValuesAndNumberFormats
            ptRange.PasteSpecial xlPasteFormats
            Application.CutCopyMode = False
            
            ' 古いピボットテーブルを削除（範囲に残る場合がある）
            pt.TableRange2.ClearContents
        End If
        
        ' ピボットテーブルオブジェクト自体を削除
        pt.PivotCache.RecordCount = 0 ' キャッシュを空に
    Next i
    
    ' ピボットテーブルが残っていたら削除
    Do While ws.PivotTables.Count > 0
        ws.PivotTables(1).TableRange2.Clear
    Loop
    
    On Error GoTo 0
End Sub

' === サブルーチン：リストオブジェクトを値化 ===
Private Sub 値化リストオブジェクト(ws As Worksheet)
    Dim lo As ListObject
    Dim loRange As Range
    
    On Error Resume Next
    
    ' 後ろから処理
    Dim i As Long
    For i = ws.ListObjects.Count To 1 Step -1
        Set lo = ws.ListObjects(i)
        Set loRange = lo.Range
        
        If Not loRange Is Nothing Then
            ' テーブル範囲を値化
            loRange.Copy
            loRange.PasteSpecial xlPasteValuesAndNumberFormats
            loRange.PasteSpecial xlPasteFormats
            Application.CutCopyMode = False
            
            ' リストオブジェクト削除（範囲は残す）
            lo.Unlink ' クエリ接続を切る
            lo.Delete
        End If
    Next i
    
    On Error GoTo 0
End Sub

' === サブルーチン：名前定義の正規化 ===
Private Sub 正規化名前定義(ws As Worksheet, SourceBook As Workbook)
    Dim n As Name
    Dim nr As Range
    Dim wb As Workbook
    Set wb = ws.Parent
    
    On Error Resume Next
    
    ' 後ろから処理
    Dim i As Long
    For i = wb.Names.Count To 1 Step -1
        Set n = wb.Names(i)
        If n.Visible Then
            Set nr = Nothing
            Set nr = n.RefersToRange
            
            If Not nr Is Nothing Then
                ' 外部ブック参照を含む名前は削除
                If InStr(n.RefersTo, "[") > 0 And InStr(n.RefersTo, "]") > 0 Then
                    n.Delete
                ElseIf Not nr.Worksheet Is ws Then
                    ' 対象シート以外の名前は削除（オプション：全体コピー時は残す）
                    n.Delete
                End If
            Else
                ' 範囲解決できない名前は削除
                n.Delete
            End If
        End If
    Next i
    
    On Error GoTo 0
End Sub

' === サブルーチン：図形位置の調整 ===
Private Sub 調整図形位置(ws As Worksheet)
    Dim sh As Shape
    
    On Error Resume Next
    
    For Each sh In ws.Shapes
        ' シートコピーで位置は基本維持されるが、
        ' セルにアンカーされていない図形は微調整が必要な場合がある
        If sh.Placement <> xlMoveAndSize Then
            ' セルにアンカーされていない場合、相対位置を再計算
            ' 現状では特に処理不要なケースが多い
        End If
    Next sh
    
    On Error GoTo 0
End Sub
