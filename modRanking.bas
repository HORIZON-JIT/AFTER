Attribute VB_Name = "modRanking"
Option Explicit

' ==========================================================
' アフター部門 出荷ランキング TOP100  (Excel VBA版)
'
' セットアップ手順:
'   1. 新規 Excel ブックを開き .xlsm で保存
'   2. Alt+F11 → VBA エディタを開く
'   3. [ファイル] → [ファイルのインポート] でこのファイルを選択
'   4. VBA エディタを閉じる
'   5. Alt+F8 → InitRanking を実行
'   6. シート上のボタンで操作開始
'
' Oracle 接続には Oracle Client が必要です。
' 接続できない場合はデモデータで動作します。
' ==========================================================

' ── DB 接続設定 ──
Private Const DB_USER As String = "ECOREAD"
Private Const DB_PASS As String = "ECOread01#"
Private Const DB_DSN  As String = "172.25.3.119:1521/orcl.hrz.local"
Private Const MAX_RANK As Long = 100

' ── レイアウト定数 ──
Private Const SH_NAME   As String = "ランキング"
Private Const HDR_ROW   As Long = 9     ' テーブルヘッダー行
Private Const DAT_ROW   As Long = 10    ' データ開始行
Private Const MODE_CELL As String = "G1" ' モード保存セル (非表示列)
Private Const DATE_CELL As String = "D3" ' 日付入力セル
Private Const HINT_CELL As String = "E3" ' ヒント表示セル
Private Const SORT_CELL As String = "G2" ' ソート保存セル (非表示列)

' ===========================================================
'  初期セットアップ (Alt+F8 → InitRanking)
' ===========================================================
Public Sub InitRanking()
    Dim ws As Worksheet

    ' ---- シート準備 ----
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        If ThisWorkbook.Worksheets.Count = 1 _
           And WorksheetFunction.CountA(ThisWorkbook.Worksheets(1).UsedRange) = 0 Then
            ThisWorkbook.Worksheets(1).Name = SH_NAME
            Set ws = ThisWorkbook.Worksheets(1)
        Else
            Set ws = ThisWorkbook.Worksheets.Add( _
                After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
            ws.Name = SH_NAME
        End If
    End If

    ws.Activate
    Application.ScreenUpdating = False

    ' 全クリア
    ws.Cells.Clear
    Dim shp As Shape
    For Each shp In ws.Shapes
        shp.Delete
    Next shp

    ' ---- 列幅 ----
    ws.Columns("A").ColumnWidth = 8    ' 順位
    ws.Columns("B").ColumnWidth = 18   ' 品番
    ws.Columns("C").ColumnWidth = 42   ' 品名
    ws.Columns("D").ColumnWidth = 14   ' 出荷数合計
    ws.Columns("E").ColumnWidth = 14   ' 出荷回数
    ws.Columns("F").ColumnWidth = 2    ' 余白
    ws.Columns("G").ColumnWidth = 0.5  ' モード格納 (非表示)
    ws.Columns("G").Hidden = True

    ' グリッド線 OFF
    ActiveWindow.DisplayGridlines = False

    ' ---- Row 1: タイトルバー ----
    With ws.Range("A1:E1")
        .Merge
        .Value = "アフター部門 出荷ランキング TOP100"
        .Font.Size = 16
        .Font.Bold = True
        .Font.Color = vbWhite
        .Interior.Color = RGB(26, 86, 219)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 42
    End With

    ' ---- Row 2: 余白 ----
    ws.Rows(2).RowHeight = 6

    ' ---- Row 3: モードボタン + 日付入力 ----
    ws.Rows(3).RowHeight = 30

    Dim r3Top As Double: r3Top = ws.Range("A3").Top + 4
    Dim bh As Double:    bh = 22

    ' モードボタン
    Dim x As Double: x = ws.Range("A3").Left + 2
    CreateBtn ws, "btnYearly", x, r3Top, 48, bh, "SetModeYearly", "年間", RGB(26, 86, 219), vbWhite
    x = x + 52
    CreateBtn ws, "btnMonthly", x, r3Top, 48, bh, "SetModeMonthly", "月間", RGB(229, 231, 235), RGB(31, 41, 55)
    x = x + 52
    CreateBtn ws, "btnDaily", x, r3Top, 48, bh, "SetModeDaily", "日別", RGB(229, 231, 235), RGB(31, 41, 55)

    ' 「対象期間:」ラベル
    With ws.Range("C3")
        .Value = "対象期間:"
        .Font.Bold = True
        .Font.Size = 10
        .HorizontalAlignment = xlRight
    End With

    ' 日付入力セル
    With ws.Range(DATE_CELL)
        .NumberFormat = "@"
        .Value = CStr(Year(Date))
        .Interior.Color = RGB(239, 242, 255)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(26, 86, 219)
        .Borders.Weight = xlThin
        .HorizontalAlignment = xlCenter
        .Font.Size = 11
    End With

    ' ヒント
    With ws.Range(HINT_CELL)
        .Value = "例: 2024"
        .Font.Color = RGB(156, 163, 175)
        .Font.Size = 9
    End With

    ' モード初期値
    ws.Range(MODE_CELL).Value = "年間"

    ' ---- Row 4: アクションボタン ----
    ws.Rows(4).RowHeight = 30
    Dim r4Top As Double: r4Top = ws.Range("A4").Top + 4

    x = ws.Range("A4").Left + 2
    CreateBtn ws, "btnPrev", x, r4Top, 52, bh, "NavigatePrev", "<< 前へ", RGB(229, 231, 235), RGB(31, 41, 55)
    x = x + 56
    CreateBtn ws, "btnSearch", x, r4Top, 52, bh, "FetchRanking", "検索", RGB(26, 86, 219), vbWhite
    x = x + 56
    CreateBtn ws, "btnNext", x, r4Top, 52, bh, "NavigateNext", "次へ >>", RGB(229, 231, 235), RGB(31, 41, 55)

    CreateBtn ws, "btnCSV", ws.Range("D4").Left + 2, r4Top, 75, bh, _
              "ExportCSV", "CSV出力", RGB(16, 185, 129), vbWhite

    ' ---- Row 5: ソート切替 ----
    ws.Rows(5).RowHeight = 28
    ws.Range("A5").Value = "並び順:"
    ws.Range("A5").Font.Bold = True
    ws.Range("A5").Font.Size = 10
    Dim r5Top As Double: r5Top = ws.Range("A5").Top + 3
    x = ws.Range("B5").Left
    CreateBtn ws, "btnSortQty", x, r5Top, 56, bh, "SetSortByQty", "出荷数順", RGB(26, 86, 219), vbWhite
    x = x + 60
    CreateBtn ws, "btnSortFreq", x, r5Top, 56, bh, "SetSortByFreq", "出荷回数順", RGB(229, 231, 235), RGB(31, 41, 55)

    ' ソート初期値
    ws.Range(SORT_CELL).Value = "出荷数"

    ' ---- Row 6: サマリー ----
    ws.Range("A6").Value = "集計期間:"
    ws.Range("A6").Font.Bold = True
    ws.Range("A6").Font.Size = 10
    ws.Range("B6").Font.Size = 11
    ws.Range("B6").Font.Color = RGB(26, 86, 219)
    ws.Range("B6").Font.Bold = True
    ws.Range("C6").Value = "品番数:"
    ws.Range("C6").Font.Bold = True
    ws.Range("C6").Font.Size = 10
    ws.Range("C6").HorizontalAlignment = xlRight
    ws.Range("D6").Font.Size = 10
    ws.Range("D6").Font.Color = RGB(26, 86, 219)
    ws.Range("E6").Font.Size = 10
    ws.Range("E6").Font.Color = RGB(26, 86, 219)
    ws.Range("E6").Font.Bold = True

    ' ---- Row 7: フィルター ----
    ws.Range("A7").Value = "絞り込み:"
    ws.Range("A7").Font.Bold = True
    ws.Range("A7").Font.Size = 10
    With ws.Range("B7:C7")
        .Merge
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(200, 200, 200)
        .Borders.Weight = xlThin
        .Font.Size = 10
    End With

    Dim r7Top As Double: r7Top = ws.Range("A7").Top + 3
    CreateBtn ws, "btnFilter", ws.Range("D7").Left + 2, r7Top, 52, bh, _
              "ApplyFilter", "絞り込み", RGB(107, 114, 128), vbWhite
    CreateBtn ws, "btnClear", ws.Range("E7").Left + 2, r7Top, 52, bh, _
              "ClearFilter", "解除", RGB(229, 231, 235), RGB(31, 41, 55)

    ' ---- Row 8: 余白 ----
    ws.Rows(8).RowHeight = 6

    ' ---- Row 9: テーブルヘッダー ----
    Dim headers As Variant
    headers = Array("順位", "品番", "品名", "出荷数合計", "出荷回数")
    Dim i As Long
    For i = 0 To 4
        With ws.Cells(HDR_ROW, i + 1)
            .Value = headers(i)
            .Font.Bold = True
            .Font.Size = 10
            .Font.Color = RGB(107, 114, 128)
            .Interior.Color = RGB(249, 250, 251)
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Color = RGB(209, 213, 219)
            .Borders(xlEdgeBottom).Weight = xlMedium
        End With
    Next i
    ws.Cells(HDR_ROW, 1).HorizontalAlignment = xlCenter
    ws.Cells(HDR_ROW, 4).HorizontalAlignment = xlRight
    ws.Cells(HDR_ROW, 5).HorizontalAlignment = xlRight

    ' 初期メッセージ
    With ws.Range("A" & DAT_ROW & ":E" & DAT_ROW)
        .Merge
        .Value = "「検索」ボタンを押してデータを取得してください"
        .HorizontalAlignment = xlCenter
        .Font.Color = RGB(156, 163, 175)
        .Font.Size = 11
    End With

    Application.ScreenUpdating = True

    MsgBox "セットアップ完了！" & vbCrLf & vbCrLf & _
           "モードと対象期間を設定して「検索」を押してください。", _
           vbInformation, "出荷ランキング"
End Sub

' ===========================================================
'  ボタン作成ヘルパー
' ===========================================================
Private Sub CreateBtn(ws As Worksheet, sName As String, _
                      x As Double, y As Double, w As Double, h As Double, _
                      macroName As String, caption As String, _
                      bgColor As Long, fgColor As Long)
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, y, w, h)
    With shp
        .Name = sName
        .Fill.ForeColor.RGB = bgColor
        .Line.Visible = msoFalse
        .TextFrame2.TextRange.Text = caption
        .TextFrame2.TextRange.Font.Size = 9
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = fgColor
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.MarginLeft = 2
        .TextFrame2.MarginRight = 2
        .OnAction = macroName
        .Adjustments(1) = 0.25
    End With
End Sub

' ===========================================================
'  モード切替
' ===========================================================
Public Sub SetModeYearly()
    SetMode "年間"
End Sub

Public Sub SetModeMonthly()
    SetMode "月間"
End Sub

Public Sub SetModeDaily()
    SetMode "日別"
End Sub

Private Sub SetMode(newMode As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SH_NAME)

    ' 現在の日付入力からパーツ取得
    Dim cur As String: cur = Trim(CStr(ws.Range(DATE_CELL).Value))
    Dim yr As String:  yr = CStr(Year(Date))
    Dim mo As String:  mo = Format(Month(Date), "00")
    Dim dy As String:  dy = Format(Day(Date), "00")

    If Len(cur) >= 4 And IsNumeric(Left(cur, 4)) Then yr = Left(cur, 4)
    If Len(cur) >= 7 And InStr(cur, "-") > 0 Then
        Dim p() As String: p = Split(cur, "-")
        If UBound(p) >= 1 Then mo = Right("00" & p(1), 2)
        If UBound(p) >= 2 Then dy = Right("00" & p(2), 2)
    End If

    ' モード保存 & 日付更新
    ws.Range(MODE_CELL).Value = newMode
    Select Case newMode
        Case "年間"
            ws.Range(DATE_CELL).Value = yr
            ws.Range(HINT_CELL).Value = "例: 2024"
        Case "月間"
            ws.Range(DATE_CELL).Value = yr & "-" & mo
            ws.Range(HINT_CELL).Value = "例: 2024-03"
        Case "日別"
            ws.Range(DATE_CELL).Value = yr & "-" & mo & "-" & dy
            ws.Range(HINT_CELL).Value = "例: 2024-03-15"
    End Select

    ' ボタン色更新
    UpdateModeButtons ws, newMode

    ' 自動検索
    FetchRanking
End Sub

Private Sub UpdateModeButtons(ws As Worksheet, activeMode As String)
    Dim modes As Variant: modes = Array("年間", "月間", "日別")
    Dim names As Variant: names = Array("btnYearly", "btnMonthly", "btnDaily")
    Dim i As Long

    For i = 0 To 2
        On Error Resume Next
        Dim shp As Shape
        Set shp = ws.Shapes(names(i))
        If Not shp Is Nothing Then
            If modes(i) = activeMode Then
                shp.Fill.ForeColor.RGB = RGB(26, 86, 219)
                shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = vbWhite
            Else
                shp.Fill.ForeColor.RGB = RGB(229, 231, 235)
                shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(31, 41, 55)
            End If
        End If
        Set shp = Nothing
        On Error GoTo 0
    Next i
End Sub

' ===========================================================
'  ソート切替
' ===========================================================
Public Sub SetSortByQty()
    SetSort "出荷数"
End Sub

Public Sub SetSortByFreq()
    SetSort "出荷回数"
End Sub

Private Sub SetSort(newSort As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SH_NAME)
    ws.Range(SORT_CELL).Value = newSort
    UpdateSortButtons ws, newSort
    FetchRanking
End Sub

Private Sub UpdateSortButtons(ws As Worksheet, activeSort As String)
    Dim sorts As Variant: sorts = Array("出荷数", "出荷回数")
    Dim names As Variant: names = Array("btnSortQty", "btnSortFreq")
    Dim i As Long
    For i = 0 To 1
        On Error Resume Next
        Dim shp As Shape
        Set shp = ws.Shapes(names(i))
        If Not shp Is Nothing Then
            If sorts(i) = activeSort Then
                shp.Fill.ForeColor.RGB = RGB(26, 86, 219)
                shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = vbWhite
            Else
                shp.Fill.ForeColor.RGB = RGB(229, 231, 235)
                shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(31, 41, 55)
            End If
        End If
        Set shp = Nothing
        On Error GoTo 0
    Next i
End Sub

' ===========================================================
'  データ取得  (メイン処理)
' ===========================================================
Public Sub FetchRanking()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SH_NAME)

    Dim mode As String:     mode = ws.Range(MODE_CELL).Value
    Dim target As String:   target = Trim(CStr(ws.Range(DATE_CELL).Value))
    Dim sortMode As String: sortMode = ws.Range(SORT_CELL).Value

    If mode = "" Then mode = "年間"
    If sortMode = "" Then sortMode = "出荷数"
    If target = "" Then
        MsgBox "対象期間を入力してください。", vbExclamation
        Exit Sub
    End If

    ' 入力チェック
    If Not IsValidInput(mode, target) Then Exit Sub

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.StatusBar = "データ取得中..."

    ' データエリアクリア
    ClearDataArea ws

    ' 日付条件 & 期間ラベル構築
    Dim dateCondition As String
    Dim periodLabel As String
    Dim parts() As String

    Select Case mode
        Case "年間"
            dateCondition = "to_char(sm.shukka_j_date, 'YYYY') = '" & Left(target, 4) & "'"
            periodLabel = Left(target, 4) & "年"

        Case "月間"
            parts = Split(target, "-")
            dateCondition = "to_char(sm.shukka_j_date, 'YYYY') = '" & parts(0) & "'" & _
                            " AND to_char(sm.shukka_j_date, 'MM') = '" & parts(1) & "'"
            periodLabel = parts(0) & "年" & CInt(parts(1)) & "月"

        Case "日別"
            parts = Split(target, "-")
            dateCondition = "to_char(sm.shukka_j_date, 'YYYY/MM/DD') = '" & _
                            parts(0) & "/" & parts(1) & "/" & parts(2) & "'"
            periodLabel = parts(0) & "年" & CInt(parts(1)) & "月" & CInt(parts(2)) & "日"
    End Select

    ' SQL 構築 (ソートモードで集計・ORDER BYを切替)
    Dim sql As String
    If sortMode = "出荷回数" Then
        ' 出荷回数ランキング: COUNT(*) で出荷レコード件数を集計
        sql = "SELECT ROWNUM AS rank_no, t.hinban, t.hm_nm, " & _
              "t.total_shukka_suu, t.shukka_count " & _
              "FROM (" & _
              "  SELECT jm.hinban, " & _
              "    MAX(jm.juchuu_hm_nm) AS hm_nm, " & _
              "    SUM(sm.shukka_j_suu) AS total_shukka_suu, " & _
              "    COUNT(*) AS shukka_count " & _
              "  FROM ecouser.t_shukka_m sm " & _
              "  INNER JOIN ecouser.t_lot_info li ON li.lot_no = sm.lot_no " & _
              "  INNER JOIN ecouser.t_juchuu_m jm " & _
              "    ON jm.juchuu_no = li.juchuu_no AND jm.juchuu_line_no = li.juchuu_line_no " & _
              "  INNER JOIN ecouser.t_juchuu_h jh " & _
              "    ON jh.juchuu_no = jm.juchuu_no AND jh.juchuu_hansuu = jm.juchuu_hansuu " & _
              "  WHERE jh.juchuu_kyoten_cd = 'A' " & _
              "    AND " & dateCondition & " " & _
              "    AND jm.hinban NOT LIKE 'C%' " & _
              "    AND jm.hinban NOT LIKE 'U%' " & _
              "  GROUP BY jm.hinban " & _
              "  ORDER BY COUNT(*) DESC, SUM(sm.shukka_j_suu) DESC" & _
              ") t " & _
              "WHERE ROWNUM <= " & MAX_RANK
    Else
        ' 出荷数ランキング: SUM で出荷数量を集計
        sql = "SELECT ROWNUM AS rank_no, t.hinban, t.hm_nm, " & _
              "t.total_shukka_suu, t.shukka_count " & _
              "FROM (" & _
              "  SELECT jm.hinban, " & _
              "    MAX(jm.juchuu_hm_nm) AS hm_nm, " & _
              "    SUM(sm.shukka_j_suu) AS total_shukka_suu, " & _
              "    COUNT(*) AS shukka_count " & _
              "  FROM ecouser.t_shukka_m sm " & _
              "  INNER JOIN ecouser.t_lot_info li ON li.lot_no = sm.lot_no " & _
              "  INNER JOIN ecouser.t_juchuu_m jm " & _
              "    ON jm.juchuu_no = li.juchuu_no AND jm.juchuu_line_no = li.juchuu_line_no " & _
              "  INNER JOIN ecouser.t_juchuu_h jh " & _
              "    ON jh.juchuu_no = jm.juchuu_no AND jh.juchuu_hansuu = jm.juchuu_hansuu " & _
              "  WHERE jh.juchuu_kyoten_cd = 'A' " & _
              "    AND " & dateCondition & " " & _
              "    AND jm.hinban NOT LIKE 'C%' " & _
              "    AND jm.hinban NOT LIKE 'U%' " & _
              "  GROUP BY jm.hinban " & _
              "  ORDER BY SUM(sm.shukka_j_suu) DESC, COUNT(*) DESC" & _
              ") t " & _
              "WHERE ROWNUM <= " & MAX_RANK
    End If

    ' ---- Oracle 接続試行 ----
    Dim conn As Object
    Dim rs As Object
    Dim oracleOK As Boolean: oracleOK = False

    On Error Resume Next
    Set conn = CreateObject("ADODB.Connection")
    conn.ConnectionTimeout = 10

    ' OraOLEDB で接続
    conn.Open "Provider=OraOLEDB.Oracle;Data Source=" & DB_DSN & _
              ";User Id=" & DB_USER & ";Password=" & DB_PASS & ";"

    If Err.Number <> 0 Then
        Err.Clear
        ' MSDAORA フォールバック
        conn.Open "Provider=MSDAORA;Data Source=" & DB_DSN & _
                  ";User Id=" & DB_USER & ";Password=" & DB_PASS & ";"
    End If

    If Err.Number = 0 Then
        Set rs = CreateObject("ADODB.Recordset")
        rs.Open sql, conn, 0, 1  ' ForwardOnly, ReadOnly
        If Err.Number = 0 Then
            oracleOK = True
        Else
            Err.Clear
        End If
    Else
        Err.Clear
    End If
    On Error GoTo 0

    ' ---- データ書き込み (配列一括) ----
    Dim totalQty As Long: totalQty = 0
    Dim cnt As Long:      cnt = 0
    Dim buf() As Variant

    If oracleOK Then
        ReDim buf(0 To MAX_RANK - 1, 0 To 4)
        Do While Not rs.EOF
            buf(cnt, 0) = rs.Fields("rank_no").Value
            buf(cnt, 1) = rs.Fields("hinban").Value
            buf(cnt, 2) = rs.Fields("hm_nm").Value
            buf(cnt, 3) = rs.Fields("total_shukka_suu").Value
            buf(cnt, 4) = rs.Fields("shukka_count").Value
            totalQty = totalQty + CLng(rs.Fields("total_shukka_suu").Value)
            cnt = cnt + 1
            rs.MoveNext
        Loop
        rs.Close
        conn.Close
        If cnt > 0 Then
            ReDim Preserve buf(0 To cnt - 1, 0 To 4)
            ws.Range("A" & DAT_ROW).Resize(cnt, 5).Value = buf
        End If
    Else
        ' デモデータ
        If Not conn Is Nothing Then
            On Error Resume Next
            If conn.State = 1 Then conn.Close
            On Error GoTo 0
        End If
        periodLabel = periodLabel & " (デモ)"
        LoadDemoData ws, cnt, totalQty, buf, sortMode
        If cnt > 0 Then
            ws.Range("A" & DAT_ROW).Resize(cnt, 5).Value = buf
        End If
    End If

    Set rs = Nothing
    Set conn = Nothing

    ' ---- 一括書式設定 ----
    If cnt > 0 Then FormatDataBulk ws, cnt

    ' ---- タイトル・サマリー更新 ----
    If sortMode = "出荷回数" Then
        ws.Range("A1").Value = "アフター部門 出荷回数ランキング TOP100"
    Else
        ws.Range("A1").Value = "アフター部門 出荷数ランキング TOP100"
    End If
    ws.Range("B6").Value = periodLabel
    ws.Range("D6").Value = cnt & " 品番"
    ws.Range("E6").Value = "合計: " & Format(totalQty, "#,##0") & " 個"

    ' ---- データバー (条件付き書式) ----
    If cnt > 0 Then ApplyDataBars ws, sortMode

    Application.StatusBar = False
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

' ===========================================================
'  入力チェック
' ===========================================================
Private Function IsValidInput(mode As String, target As String) As Boolean
    IsValidInput = True
    Select Case mode
        Case "年間"
            If Not (Len(target) = 4 And IsNumeric(target)) Then
                MsgBox "年間モード: 年を4桁で入力" & vbCrLf & "例: 2024", vbExclamation
                IsValidInput = False
            End If
        Case "月間"
            If Not target Like "####-##" Then
                MsgBox "月間モード: 年月を入力" & vbCrLf & "例: 2024-03", vbExclamation
                IsValidInput = False
            End If
        Case "日別"
            If Not target Like "####-##-##" Then
                MsgBox "日別モード: 日付を入力" & vbCrLf & "例: 2024-03-15", vbExclamation
                IsValidInput = False
            End If
    End Select
End Function

' ===========================================================
'  データエリアクリア
' ===========================================================
Private Sub ClearDataArea(ws As Worksheet)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    If lastRow >= DAT_ROW Then
        ws.Range("A" & DAT_ROW & ":E" & lastRow).Clear
        ws.Rows(DAT_ROW & ":" & lastRow).Hidden = False
    End If
    ' マージセル解除
    On Error Resume Next
    ws.Range("A" & DAT_ROW & ":E" & DAT_ROW).UnMerge
    On Error GoTo 0
End Sub

' ===========================================================
'  データ全行の書式を一括設定
' ===========================================================
Private Sub FormatDataBulk(ws As Worksheet, cnt As Long)
    Dim lastRow As Long: lastRow = DAT_ROW + cnt - 1
    Dim rng As Range
    Set rng = ws.Range("A" & DAT_ROW & ":E" & lastRow)

    ' フォントサイズ一括
    rng.Font.Size = 10

    ' 下罫線一括
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = RGB(229, 231, 235)
        .Weight = xlHairline
    End With
    With rng.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Color = RGB(229, 231, 235)
        .Weight = xlHairline
    End With

    ' 順位列: 中央揃え
    ws.Range("A" & DAT_ROW & ":A" & lastRow).HorizontalAlignment = xlCenter
    ' 数値列: 右揃え・カンマ区切り
    With ws.Range("D" & DAT_ROW & ":D" & lastRow)
        .HorizontalAlignment = xlRight
        .NumberFormat = "#,##0"
    End With
    With ws.Range("E" & DAT_ROW & ":E" & lastRow)
        .HorizontalAlignment = xlRight
        .NumberFormat = "#,##0"
    End With

    ' Top 3 メダル風
    With ws.Cells(DAT_ROW, 1)
        .Interior.Color = RGB(254, 243, 199)
        .Font.Color = RGB(217, 119, 6)
        .Font.Bold = True
    End With
    If cnt >= 2 Then
        With ws.Cells(DAT_ROW + 1, 1)
            .Interior.Color = RGB(243, 244, 246)
            .Font.Color = RGB(107, 114, 128)
            .Font.Bold = True
        End With
    End If
    If cnt >= 3 Then
        With ws.Cells(DAT_ROW + 2, 1)
            .Interior.Color = RGB(254, 243, 199)
            .Font.Color = RGB(180, 83, 9)
            .Font.Bold = True
        End With
    End If
End Sub

' ===========================================================
'  データバー (条件付き書式)
' ===========================================================
Private Sub ApplyDataBars(ws As Worksheet, sortMode As String)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 4).End(xlUp).row
    If lastRow < DAT_ROW Then Exit Sub

    ' 両列のデータバーをクリア
    ws.Range("D" & DAT_ROW & ":D" & lastRow).FormatConditions.Delete
    ws.Range("E" & DAT_ROW & ":E" & lastRow).FormatConditions.Delete

    ' ソート対象列にデータバー適用
    Dim col As String
    If sortMode = "出荷回数" Then col = "E" Else col = "D"

    Dim rng As Range
    Set rng = ws.Range(col & DAT_ROW & ":" & col & lastRow)

    Dim db As Object
    Set db = rng.FormatConditions.AddDatabar
    db.BarColor.Color = RGB(191, 210, 247)
    db.ShowValue = True
End Sub

' ===========================================================
'  デモデータ (配列で返す)
' ===========================================================
Private Sub LoadDemoData(ws As Worksheet, ByRef cnt As Long, _
                         ByRef totalQty As Long, ByRef buf() As Variant, _
                         sortMode As String)
    Dim items(0 To 9, 0 To 1) As String
    items(0, 0) = "4012273-00": items(0, 1) = "ベアリング A"
    items(1, 0) = "4012274-01": items(1, 1) = "シャフト B"
    items(2, 0) = "M167001-18": items(2, 1) = "モータ C"
    items(3, 0) = "E505712-06": items(3, 1) = "センサー D"
    items(4, 0) = "A910246-01": items(4, 1) = "バルブ E"
    items(5, 0) = "4000064-01": items(5, 1) = "ギア F"
    items(6, 0) = "M206597-02": items(6, 1) = "ポンプ G"
    items(7, 0) = "A970831-00": items(7, 1) = "フィルタ H"
    items(8, 0) = "4015500-03": items(8, 1) = "カップリング I"
    items(9, 0) = "E300100-02": items(9, 1) = "コントローラ J"

    Const DEMO_COUNT As Long = 20
    ReDim buf(0 To DEMO_COUNT - 1, 0 To 4)
    Dim i As Long, idx As Long
    Dim qty As Long, freq As Long, hinban As String

    For i = 0 To DEMO_COUNT - 1
        idx = i Mod 10
        hinban = items(idx, 0)
        If i >= 10 Then hinban = hinban & "-" & Format(i \ 10, "00")

        If sortMode = "出荷回数" Then
            ' 出荷回数順: 回数が多い順 (回数と数量は必ずしも比例しない)
            freq = 120 - i * 5
            If freq < 1 Then freq = 1
            qty = freq * (3 + (idx Mod 5)) + (10 - idx) * 2
        Else
            ' 出荷数順: 数量が多い順
            qty = 500 - i * 20
            If qty < 1 Then qty = 1
            freq = WorksheetFunction.Max(1, 10 + (20 - i) * 3)
        End If

        buf(i, 0) = i + 1
        buf(i, 1) = hinban
        buf(i, 2) = items(idx, 1) & " (" & (i + 1) & ")"
        buf(i, 3) = qty
        buf(i, 4) = freq

        totalQty = totalQty + qty
    Next i
    cnt = DEMO_COUNT
End Sub

' ===========================================================
'  日付ナビゲーション
' ===========================================================
Public Sub NavigatePrev()
    StepDate -1
End Sub

Public Sub NavigateNext()
    StepDate 1
End Sub

Private Sub StepDate(delta As Long)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SH_NAME)

    Dim mode As String:   mode = ws.Range(MODE_CELL).Value
    Dim target As String: target = Trim(CStr(ws.Range(DATE_CELL).Value))

    Select Case mode
        Case "年間"
            If IsNumeric(target) Then
                ws.Range(DATE_CELL).Value = CStr(CLng(target) + delta)
            End If

        Case "月間"
            If InStr(target, "-") > 0 Then
                Dim mp() As String: mp = Split(target, "-")
                Dim md As Date: md = DateSerial(CInt(mp(0)), CInt(mp(1)) + delta, 1)
                ws.Range(DATE_CELL).Value = Format(md, "yyyy-mm")
            End If

        Case "日別"
            If target Like "####-##-##" Then
                Dim dp() As String: dp = Split(target, "-")
                Dim dd As Date: dd = DateSerial(CInt(dp(0)), CInt(dp(1)), CInt(dp(2)) + delta)
                ws.Range(DATE_CELL).Value = Format(dd, "yyyy-mm-dd")
            End If
    End Select

    FetchRanking
End Sub

' ===========================================================
'  絞り込みフィルター
' ===========================================================
Public Sub ApplyFilter()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SH_NAME)

    Dim filterText As String
    filterText = LCase(Trim(CStr(ws.Range("B7").Value)))

    If filterText = "" Then
        ClearFilter
        Exit Sub
    End If

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    If lastRow < DAT_ROW Then Exit Sub

    Application.ScreenUpdating = False

    Dim row As Long
    For row = DAT_ROW To lastRow
        Dim hinban As String: hinban = LCase(CStr(ws.Cells(row, 2).Value))
        Dim hm_nm As String:  hm_nm = LCase(CStr(ws.Cells(row, 3).Value))

        ws.Rows(row).Hidden = (InStr(hinban, filterText) = 0 And InStr(hm_nm, filterText) = 0)
    Next row

    Application.ScreenUpdating = True
End Sub

Public Sub ClearFilter()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SH_NAME)

    ws.Range("B7").Value = ""

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    If lastRow >= DAT_ROW Then
        ws.Rows(DAT_ROW & ":" & lastRow).Hidden = False
    End If
End Sub

' ===========================================================
'  CSV 出力
' ===========================================================
Public Sub ExportCSV()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SH_NAME)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    If lastRow < DAT_ROW Then
        MsgBox "データがありません。", vbExclamation
        Exit Sub
    End If

    ' ファイル保存ダイアログ
    Dim period As String: period = CStr(ws.Range("B6").Value)
    If period = "" Then period = "ranking"

    Dim fileName As Variant
    fileName = Application.GetSaveAsFilename( _
        InitialFileName:="出荷ランキング_" & period & ".csv", _
        FileFilter:="CSV ファイル (*.csv),*.csv")

    If fileName = False Then Exit Sub

    ' BOM 付き UTF-8 で出力 (ADODB.Stream)
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2         ' adTypeText
    stm.Charset = "UTF-8"
    stm.Open

    ' ヘッダー
    stm.WriteText "順位,品番,品名,出荷数合計,出荷回数", 1  ' 1 = adWriteLine

    ' データ (表示行のみ)
    Dim row As Long
    For row = DAT_ROW To lastRow
        If Not ws.Rows(row).Hidden Then
            Dim csvLine As String
            csvLine = ws.Cells(row, 1).Value & "," & _
                      """" & Replace(CStr(ws.Cells(row, 2).Value), """", """""") & """," & _
                      """" & Replace(CStr(ws.Cells(row, 3).Value), """", """""") & """," & _
                      ws.Cells(row, 4).Value & "," & _
                      ws.Cells(row, 5).Value
            stm.WriteText csvLine, 1
        End If
    Next row

    stm.SaveToFile CStr(fileName), 2  ' adSaveCreateOverWrite
    stm.Close
    Set stm = Nothing

    MsgBox "CSV出力完了:" & vbCrLf & CStr(fileName), vbInformation
End Sub
