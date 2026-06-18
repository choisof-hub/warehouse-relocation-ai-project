Option Explicit

'==========================================================
' 同一ブック版（変換表は同じExcelブック内）
' 仕様:
'  - データ側: シート「データ」、ヘッダーに「棚番」
'  - 変換表側: シート「新旧棚番変換表」、ヘッダーに「旧棚番」「新棚番」
'  - 空白棚番は完全スキップ（未定義に数えない／出力もしない／ハイライトもしない）
'  - 進捗%をステータスバー表示、DoEventsで応答性確保
' 作者: M365 Copilot
'==========================================================
Public Sub Create_Update_File()

    '=========== 設定（必要に応じて変更） ===========
    Const DATA_SHEET_NAME      As String = "データ"          ' 対象データのシート名
    Const MAP_SHEET_NAME       As String = "新旧棚番変換表"  ' 変換表のシート名
    Const HDR_OLD              As String = "旧棚番"
    Const HDR_NEW              As String = "新棚番"
    Const HDR_SHELF            As String = "棚番"
    Const OUTPUT_HEADER        As String = "新棚番(変換後)"

    ' 出力モード: True=元の棚番を上書き / False=新しい列に出力
    Const OVERWRITE_ORIGINAL   As Boolean = False

    ' "XXXX" の扱い: True="XXXX" をそのまま書く / False=変換せず元の棚番を残す
    Const WRITE_XXXX_AS_IS     As Boolean = True

    ' 未定義セルをハイライト（空白棚番はそもそもスキップ対象）
    Const HIGHLIGHT_NOT_FOUND  As Boolean = True

    ' 進捗表示の更新間隔（秒）
    Const PROGRESS_UPDATE_SEC  As Double = 0.2
    '===============================================

    Dim wb As Workbook
    Dim wsData As Worksheet, wsMap As Worksheet
    Dim colShelf As Long, colOld As Long, colNew As Long
    Dim lastRowData As Long, lastRowMap As Long, outCol As Long

    Dim dict As Object ' Scripting.Dictionary
    Dim dataArr As Variant, outArr As Variant
    Dim oldArr As Variant, newArr As Variant

    Dim total As Long, i As Long
    Dim v As String, nv As String
    Dim exists As Boolean

    Dim changed As Long, notFound As Long, sameReplaced As Long, mappedXXXX As Long
    Dim rngNotFound As Range

    Dim nextTick As Double, t0 As Double

    On Error GoTo EH

    Set wb = ThisWorkbook
    Set wsData = GetSheetByName(wb, DATA_SHEET_NAME, True)
    Set wsMap = GetSheetByName(wb, MAP_SHEET_NAME, True)

    ' 列検出
    colShelf = FindHeaderColumn(wsData, HDR_SHELF)
    If colShelf = 0 Then Err.Raise vbObjectError + 101, , "データ側にヘッダー「" & HDR_SHELF & "」が見つかりません。"

    colOld = FindHeaderColumn(wsMap, HDR_OLD)
    colNew = FindHeaderColumn(wsMap, HDR_NEW)
    If colOld = 0 Or colNew = 0 Then Err.Raise vbObjectError + 102, , "変換表にヘッダー「" & HDR_OLD & "」「" & HDR_NEW & "」が見つかりません。"

    ' 最終行（堅牢）
    lastRowData = wsData.Cells(wsData.Rows.Count, colShelf).End(xlUp).Row
    If lastRowData < 2 Then
        MsgBox "データが見つかりません。", vbExclamation
        Exit Sub
    End If

    lastRowMap = wsMap.Cells(wsMap.Rows.Count, colOld).End(xlUp).Row
    If lastRowMap < 2 Then Err.Raise vbObjectError + 103, , "変換表にデータがありません。"

    ' Excel最適化
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayStatusBar = True

    ' 変換表を辞書にロード（大/小文字無視）
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1 ' TextCompare

    oldArr = wsMap.Range(wsMap.Cells(2, colOld), wsMap.Cells(lastRowMap, colOld)).Value
    newArr = wsMap.Range(wsMap.Cells(2, colNew), wsMap.Cells(lastRowMap, colNew)).Value

    For i = 1 To UBound(oldArr, 1)
        v = TrimSafe(oldArr(i, 1))
        nv = TrimSafe(newArr(i, 1))
        If LenB(v) > 0 Then
            dict(v) = nv    ' 同一旧棚番が複数あれば最後を優先
        End If
    Next i

    ' データ配列取得
    dataArr = wsData.Range(wsData.Cells(2, colShelf), wsData.Cells(lastRowData, colShelf)).Value
    total = UBound(dataArr, 1)

    ' 出力列の決定
    If OVERWRITE_ORIGINAL Then
        outCol = colShelf
        ReDim outArr(1 To total, 1 To 1)
    Else
        outCol = wsData.Cells(1, wsData.Columns.Count).End(xlToLeft).Column + 1
        wsData.Cells(1, outCol).Value = OUTPUT_HEADER
        ReDim outArr(1 To total, 1 To 1)
    End If

    ' ===== 変換（空白は完全スキップ） =====
    t0 = Timer
    nextTick = Timer + PROGRESS_UPDATE_SEC

    For i = 1 To total
        v = TrimSafe(dataArr(i, 1))

        ' --- 空白棚番は完全スキップ ---
        If LenB(v) = 0 Then
            ' 新規列の場合は空のまま、上書きの場合も元が空なので変更なし
            GoTo ProgressOnly
        End If

        exists = dict.exists(v)
        If exists Then
            nv = CStr(dict(v))
            If UCase$(nv) = "XXXX" And Not WRITE_XXXX_AS_IS Then
                ' "XXXX" は採用せず元の棚番を維持
                nv = v
            Else
                changed = changed + 1
                If UCase$(nv) = "XXXX" Then mappedXXXX = mappedXXXX + 1
                If nv = v Then sameReplaced = sameReplaced + 1
            End If
        Else
            ' 未定義（空白でないのにマップ無し）
            notFound = notFound + 1
            nv = v
            If HIGHLIGHT_NOT_FOUND Then
                If rngNotFound Is Nothing Then
                    Set rngNotFound = wsData.Cells(i + 1, outCol) ' +1はヘッダー行
                Else
                    Set rngNotFound = Union(rngNotFound, wsData.Cells(i + 1, outCol))
                End If
            End If
        End If

        outArr(i, 1) = nv

ProgressOnly:
        ' ---- 進捗（0.2秒ごと）----
        If Timer >= nextTick Then
            Application.StatusBar = "変換中 … " & Format$(i / total, "0%") & _
                                    "  (" & i & "/" & total & ")"
            DoEvents
            nextTick = Timer + PROGRESS_UPDATE_SEC
        End If
    Next i

    ' 一括書き戻し（新規列出力 or 上書き）
    wsData.Range(wsData.Cells(2, outCol), wsData.Cells(lastRowData, outCol)).Value = outArr

    ' 未定義セルのみ一括ハイライト（空白は対象外）
    If HIGHLIGHT_NOT_FOUND And Not rngNotFound Is Nothing Then
        rngNotFound.Interior.Color = RGB(255, 255, 153) ' 薄黄
    End If

    ' 完了表示
    Dim sec As Double, msg As String
    sec = Timer - t0
    Application.StatusBar = "完了: " & Format$(sec, "0.0") & " 秒 / " & total & " 行"

    msg = "棚番 → 新棚番 変換が完了しました。" & vbCrLf & vbCrLf & _
          "処理時間: " & Format$(sec, "0.0") & " 秒" & vbCrLf & _
          "変換件数: " & changed & vbCrLf & _
          "（うち ""XXXX"" 変換: " & mappedXXXX & "）" & vbCrLf & _
          "旧=新（同一値）: " & sameReplaced & vbCrLf & _
          "未定義（変換表に無し）: " & notFound & vbCrLf & vbCrLf & _
          "出力列: " & IIf(OVERWRITE_ORIGINAL, "棚番（上書き）", wsData.Cells(1, outCol).Address(False, False))
    MsgBox msg, vbInformation

CleanExit:
    On Error Resume Next
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Exit Sub

EH:
    MsgBox "エラー: " & Err.Description, vbCritical
    Resume CleanExit
End Sub

'=== ヘルパー ===
Public Function GetSheetByName(ByVal wb As Workbook, ByVal sheetName As String, _
                               Optional ByVal mustExist As Boolean = True) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing And mustExist Then
        Err.Raise vbObjectError + 201, , "シート「" & sheetName & "」が見つかりません。"
    End If
    Set GetSheetByName = ws
End Function

Public Function FindHeaderColumn(ByVal ws As Worksheet, ByVal headerName As String) As Long
    Dim lastCol As Long, c As Range
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For Each c In ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))
        If TrimSafe(c.Value) = headerName Then
            FindHeaderColumn = c.Column
            Exit Function
        End If
    Next c
    FindHeaderColumn = 0
End Function

Public Function TrimSafe(ByVal v As Variant) As String
    If IsError(v) Or IsNull(v) Then
        TrimSafe = ""
    Else
        TrimSafe = Trim$(CStr(v))
    End If
End Function

