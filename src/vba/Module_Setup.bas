Attribute VB_Name = "Module_Setup"
Option Explicit

'=============================================================
' Module_Setup: 初回セットアップ用マクロ
' SetupSheets を一度だけ実行してシート構造を作成する
'=============================================================

Private Const SHEET_TEMPLATES As String = "テンプレート一覧"

'-------------------------------------------------------------
' SetupSheets: テンプレート一覧と本文シートを一括作成する
'-------------------------------------------------------------
Public Sub SetupSheets()
    If SheetExists(SHEET_TEMPLATES) Then
        If MsgBox("すでにセットアップ済みです。シートを再作成しますか？" & vbCrLf & _
                  "（既存の内容は削除されます）", _
                  vbYesNo + vbQuestion, "セットアップ") = vbNo Then
            Exit Sub
        End If
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    CreateTemplateListSheet
    CreateBodySheets

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    ThisWorkbook.Sheets(SHEET_TEMPLATES).Activate

    MsgBox "セットアップが完了しました。" & vbCrLf & vbCrLf & _
           "次の手順で設定してください：" & vbCrLf & _
           "  1. B2〜B4 に案件情報を入力（毎回変える）" & vbCrLf & _
           "  2. 7〜11行目の宛先・CC・件名を編集" & vbCrLf & _
           "  3. 本文_1〜本文_5 シートのA2に本文を入力", _
           vbInformation, "セットアップ完了"
End Sub

'-------------------------------------------------------------
' CreateTemplateListSheet: テンプレート一覧シートを作成する
'-------------------------------------------------------------
Private Sub CreateTemplateListSheet()
    If SheetExists(SHEET_TEMPLATES) Then
        ThisWorkbook.Sheets(SHEET_TEMPLATES).Delete
    End If

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
    ws.Name = SHEET_TEMPLATES
    ws.Tab.Color = RGB(68, 114, 196)

    ' タイトル
    With ws.Range("A1:H1")
        .Merge
        .Value = "Mail Template Launcher"
        .Font.Bold = True
        .Font.Size = 14
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
        .RowHeight = 28
    End With

    ' 案件情報入力エリア（毎回ここに入力する）
    ws.Range("A2").Value = "案件名:"
    ws.Range("A3").Value = "案件番号:"
    ws.Range("A4").Value = "顧客名:"
    ws.Range("A2:A4").Font.Bold = True
    ws.Range("A2:A4").Font.Color = RGB(68, 114, 196)
    ws.Range("B2:B4").Interior.Color = RGB(255, 255, 200)
    ws.Range("A2:H4").RowHeight = 22

    ' 区切り線
    ws.Range("A5:H5").Interior.Color = RGB(200, 200, 200)
    ws.Rows(5).RowHeight = 4

    ' ヘッダー行（6行目）
    Dim headers As Variant
    headers = Array("ID", "テンプレート名", "形式", "宛先 (To)", "CC", "件名", "本文シート", "起動")
    Dim c As Integer
    For c = 0 To 7
        ws.Cells(6, c + 1).Value = headers(c)
    Next c
    With ws.Range("A6:H6")
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .RowHeight = 22
    End With

    ' テンプレート5件（7〜11行目）
    Dim i As Integer
    For i = 1 To 5
        Dim r As Long
        r = 6 + i
        ws.Cells(r, 1).Value = i
        ws.Cells(r, 2).Value = "テンプレート" & i
        ws.Cells(r, 3).Value = "TEXT"
        ws.Cells(r, 4).Value = ""
        ws.Cells(r, 5).Value = ""
        ws.Cells(r, 6).Value = "件名" & i
        ws.Cells(r, 7).Value = "本文_" & i
        ws.Rows(r).RowHeight = 22

        ' 起動ボタン
        Dim cell As Range
        Set cell = ws.Cells(r, 8)
        Dim btn As Object
        Set btn = ws.Buttons.Add(cell.Left, cell.Top, cell.Width, cell.Height)
        btn.Caption = "起動"
        btn.OnAction = "Launch_" & i
    Next i

    ' 列幅
    ws.Columns("A").ColumnWidth = 5
    ws.Columns("B").ColumnWidth = 22
    ws.Columns("C").ColumnWidth = 8
    ws.Columns("D").ColumnWidth = 28
    ws.Columns("E").ColumnWidth = 22
    ws.Columns("F").ColumnWidth = 30
    ws.Columns("G").ColumnWidth = 12
    ws.Columns("H").ColumnWidth = 10
End Sub

'-------------------------------------------------------------
' CreateBodySheets: 本文_1〜本文_5 シートを作成する
'-------------------------------------------------------------
Private Sub CreateBodySheets()
    Dim i As Integer
    For i = 1 To 5
        Dim sheetName As String
        sheetName = "本文_" & i

        If SheetExists(sheetName) Then
            ThisWorkbook.Sheets(sheetName).Delete
        End If

        Dim ws As Worksheet
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = sheetName
        ws.Tab.Color = RGB(180, 198, 231)

        ' プレースホルダー説明（A1）
        With ws.Range("A1")
            .Value = "【利用可能なプレースホルダー】  {案件名}  {案件番号}  {顧客名}"
            .Interior.Color = RGB(255, 255, 200)
            .Font.Color = RGB(128, 100, 0)
            .Font.Size = 9
            .Font.Italic = True
        End With

        ' 本文入力エリア（A2）
        With ws.Range("A2")
            .Value = ""
            .WrapText = True
            .VerticalAlignment = xlTop
            .Font.Size = 11
            .RowHeight = 300
        End With

        ws.Columns("A").ColumnWidth = 80
    Next i
End Sub

'-------------------------------------------------------------
' SheetExists: シートが存在するか確認する
'-------------------------------------------------------------
Private Function SheetExists(name As String) As Boolean
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(name)
    SheetExists = Not (ws Is Nothing)
    On Error GoTo 0
End Function
