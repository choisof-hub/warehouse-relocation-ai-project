Option Explicit

Sub create_conversion_table()

    Dim wsOld As Worksheet
    Dim wsNew As Worksheet
    Dim wsOut As Worksheet
    
    Dim r As Long, c As Long
    Dim outRow As Long
    Dim oldVal As String
    Dim newVal As String

    '--- シート設定
    Set wsOld = ThisWorkbook.Worksheets("旧棚番配置図")
    Set wsNew = ThisWorkbook.Worksheets("新棚番配置図")

    '--- 既存の出力シート削除
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets("新旧棚番変換表").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    '--- 出力シート作成
    Set wsOut = ThisWorkbook.Worksheets.Add
    wsOut.Name = "新旧棚番変換表"

    '--- 見出し（文字列指定）
    wsOut.Columns("A:B").NumberFormat = "@"
    wsOut.Cells(1, 1).Value = "旧棚番"
    wsOut.Cells(1, 2).Value = "新棚番"
    wsOut.Rows(1).Font.Bold = True

    outRow = 2

    '--- A1:Y24（600セル）走査
    For r = 1 To 24
        For c = 1 To 25   ' A～Y

            oldVal = Trim(CStr(wsOld.Cells(r, c).Value))

            ' 旧棚番が空白ならスキップ
            If oldVal <> "" Then

                newVal = Trim(CStr(wsNew.Cells(r, c).Value))

                '--- 出力セルを文字列に固定
                wsOut.Cells(outRow, 1).NumberFormat = "@"
                wsOut.Cells(outRow, 2).NumberFormat = "@"

                wsOut.Cells(outRow, 1).Value = oldVal

                If newVal <> "" Then
                    wsOut.Cells(outRow, 2).Value = newVal
                Else
                    wsOut.Cells(outRow, 2).Value = "■■■"
                End If

                outRow = outRow + 1
            End If

        Next c
    Next r

    '--- 見た目調整
    wsOut.Columns("A:B").AutoFit

    MsgBox "新旧棚番変換表を作成しました。", vbInformation

End Sub
