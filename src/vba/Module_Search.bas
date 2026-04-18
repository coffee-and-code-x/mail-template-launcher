Attribute VB_Name = "Module_Search"
Option Explicit

'=============================================================
' Module_Search: 案件検索モジュール
' 登録された外部Excelファイルからキーワードで案件を検索する
'=============================================================

' 検索結果の開始行
Private Const RESULT_START_ROW As Long = 6

'-------------------------------------------------------------
' SearchProjects: 全アクティブファイルでキーワード検索を実行する
' 「検索実行」ボタンから呼び出される
'-------------------------------------------------------------
Public Sub SearchProjects()
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_SEARCH)

    ' キーワード取得
    Dim keyword As String
    keyword = Trim(CStr(ws.Range("rng_SearchKeyword").Value))

    ' 前回の結果をクリア
    ClearResultRows ws

    ' アクティブな設定を全取得
    Dim settings As Collection
    Set settings = GetAllActiveFileSettings()

    If settings.Count = 0 Then
        SetStatusMessage ws, "A4", _
            "⚠ 有効なファイル設定がありません。「ファイル設定」シートで設定を行ってください。", True
        Exit Sub
    End If

    ' 検索実行
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    SetStatusMessage ws, "A4", "検索中... (0/" & settings.Count & " ファイル処理済み)", False

    Dim allResults As New Collection
    Dim i As Integer
    For i = 1 To settings.Count
        Dim setting As CFileSetting
        Set setting = settings(i)
        Application.StatusBar = "検索中: " & setting.DisplayName & " (" & i & "/" & settings.Count & ")"

        Dim fileResults As Collection
        Set fileResults = SearchInFile(setting, keyword)

        Dim j As Integer
        For j = 1 To fileResults.Count
            allResults.Add fileResults(j)
        Next j

        SetStatusMessage ws, "A4", _
            "検索中... (" & i & "/" & settings.Count & " ファイル処理済み / " & allResults.Count & " 件見つかりました)", False
    Next i

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = False

    ' 結果をシートに書き込む
    If allResults.Count = 0 Then
        SetStatusMessage ws, "A4", "検索結果: 0件  （キーワード: 「" & keyword & "」）", False
    Else
        WriteResultsToSheet ws, allResults
        SetStatusMessage ws, "A4", _
            "✔ " & allResults.Count & " 件見つかりました。行をクリックして「この案件を選択」ボタンで選択してください。", False
    End If

    Exit Sub
ErrHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = False
    LogError "SearchProjects", Err.Number, Err.Description
    SetStatusMessage ws, "A4", "エラーが発生しました: " & Err.Description, True
End Sub

'-------------------------------------------------------------
' WriteResultsToSheet: 検索結果をシートの指定行から書き込む
'-------------------------------------------------------------
Private Sub WriteResultsToSheet(ws As Worksheet, results As Collection)
    Dim row As Long
    row = RESULT_START_ROW

    Dim i As Integer
    For i = 1 To results.Count
        Dim pd As CProjectData
        Set pd = results(i)

        ws.Cells(row, 1).Value = pd.案件名
        ws.Cells(row, 2).Value = pd.案件番号
        ws.Cells(row, 3).Value = pd.顧客名
        ws.Cells(row, 4).Value = pd.担当者名
        ws.Cells(row, 5).Value = pd.期日
        ws.Cells(row, 6).Value = pd.SourceFile

        ' 警告行（⚠で始まる案件名）はオレンジ背景
        If Left(pd.案件名, 1) = Chr(9888) Or Left(pd.案件名, 2) = Chr(226) Then
            ws.Rows(row).Interior.Color = RGB(255, 235, 200)
            ws.Rows(row).Font.Color = RGB(180, 80, 0)
        Else
            ' 交互に背景色を設定（見やすさ向上）
            If (row - RESULT_START_ROW) Mod 2 = 0 Then
                ws.Rows(row).Interior.Color = RGB(255, 255, 255)
            Else
                ws.Rows(row).Interior.Color = RGB(242, 242, 242)
            End If
            ws.Rows(row).Font.Color = RGB(0, 0, 0)
        End If

        ws.Rows(row).RowHeight = 20
        row = row + 1
    Next i
End Sub

'-------------------------------------------------------------
' ClearResultRows: 検索結果行をクリアする
'-------------------------------------------------------------
Public Sub ClearResultRows(Optional ws As Worksheet = Nothing)
    If ws Is Nothing Then
        If SheetExists(SHEET_SEARCH) Then
            Set ws = ThisWorkbook.Sheets(SHEET_SEARCH)
        Else
            Exit Sub
        End If
    End If

    ' 結果行（6行目以降28行目まで）をクリア
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow >= RESULT_START_ROW Then
        ws.Range(ws.Cells(RESULT_START_ROW, 1), ws.Cells(lastRow, 6)).ClearContents
        ws.Range(ws.Cells(RESULT_START_ROW, 1), ws.Cells(lastRow, 6)).Interior.ColorIndex = xlNone
        ws.Range(ws.Cells(RESULT_START_ROW, 1), ws.Cells(lastRow, 6)).Font.ColorIndex = xlAutomatic
    End If
End Sub

'-------------------------------------------------------------
' ClearSearchResults: 検索結果とステータスをクリアする（ボタン用）
'-------------------------------------------------------------
Public Sub ClearSearchResults()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_SEARCH)
    ClearResultRows ws
    ClearStatusMessage ws, "A4"
    ws.Range("rng_SearchKeyword").Value = ""
End Sub

'-------------------------------------------------------------
' SelectProject: アクティブセル行の案件を選択状態にする
' 「この案件を選択」ボタンから呼び出される
'-------------------------------------------------------------
Public Sub SelectProject()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_SEARCH)

    ' アクティブなシートが案件検索シートかチェック
    If ActiveSheet.Name <> SHEET_SEARCH Then
        MsgBox "案件検索シートで行を選択してから、このボタンを押してください。", _
               vbInformation, "案件選択"
        Exit Sub
    End If

    Dim selectedRow As Long
    selectedRow = ActiveCell.Row

    ' 選択行の範囲チェック
    If selectedRow < RESULT_START_ROW Then
        MsgBox "案件の行（" & RESULT_START_ROW & "行目以降）をクリックしてから「この案件を選択」を押してください。", _
               vbInformation, "案件選択"
        Exit Sub
    End If

    ' 選択行に案件名がない場合
    If Trim(CStr(ws.Cells(selectedRow, 1).Value)) = "" Then
        MsgBox "選択した行に案件データがありません。" & vbCrLf & _
               "案件が表示されている行をクリックしてから「この案件を選択」を押してください。", _
               vbInformation, "案件選択"
        Exit Sub
    End If

    ' ⚠ 警告行は選択不可
    Dim cellVal As String
    cellVal = CStr(ws.Cells(selectedRow, 1).Value)
    If Left(cellVal, 1) = "?" Or InStr(cellVal, "ファイルが見つかりません") > 0 Then
        MsgBox "この行は警告メッセージです。案件データの行を選択してください。", _
               vbInformation, "案件選択"
        Exit Sub
    End If

    ' 名前付き範囲に選択値を書き込む
    ws.Range("rng_Sel_案件名").Value    = ws.Cells(selectedRow, 1).Value
    ws.Range("rng_Sel_案件番号").Value  = ws.Cells(selectedRow, 2).Value
    ws.Range("rng_Sel_顧客名").Value    = ws.Cells(selectedRow, 3).Value
    ws.Range("rng_Sel_担当者名").Value  = ws.Cells(selectedRow, 4).Value
    ws.Range("rng_Sel_期日").Value      = ws.Cells(selectedRow, 5).Value

    ' 選択行をハイライト
    HighlightSelectedRow ws, selectedRow

    ' ステータス表示
    Dim selectedName As String
    selectedName = CStr(ws.Cells(selectedRow, 1).Value)
    SetStatusMessage ws, "A4", "✔ 案件「" & selectedName & "」を選択しました。テンプレート一覧へ戻ってメールを起動してください。", False

    MsgBox "案件「" & selectedName & "」を選択しました。" & vbCrLf & vbCrLf & _
           "「テンプレート一覧へ」ボタンで戻り、テンプレートの「起動」ボタンを押してください。", _
           vbInformation, "案件選択完了"
End Sub

'-------------------------------------------------------------
' HighlightSelectedRow: 選択した結果行をハイライトする
'-------------------------------------------------------------
Private Sub HighlightSelectedRow(ws As Worksheet, selectedRow As Long)
    ' 全結果行のハイライトを一旦解除（偶数/奇数の背景に戻す）
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim r As Long
    For r = RESULT_START_ROW To lastRow
        If Trim(CStr(ws.Cells(r, 1).Value)) <> "" Then
            If (r - RESULT_START_ROW) Mod 2 = 0 Then
                ws.Rows(r).Interior.Color = RGB(255, 255, 255)
            Else
                ws.Rows(r).Interior.Color = RGB(242, 242, 242)
            End If
            ws.Rows(r).Font.Color = RGB(0, 0, 0)
        End If
    Next r

    ' 選択行を青ハイライト
    ws.Range(ws.Cells(selectedRow, 1), ws.Cells(selectedRow, 6)).Interior.Color = RGB(173, 216, 230)
    ws.Range(ws.Cells(selectedRow, 1), ws.Cells(selectedRow, 6)).Font.Color = RGB(0, 0, 128)
    ws.Range(ws.Cells(selectedRow, 1), ws.Cells(selectedRow, 6)).Font.Bold = True
End Sub

'-------------------------------------------------------------
' GetSelectedProject: 現在選択中の案件情報を返す
' ProjectData が空（案件名=""）の場合は未選択
'-------------------------------------------------------------
Public Function GetSelectedProject() As CProjectData
    Dim pd As New CProjectData
    On Error GoTo ErrHandler

    pd.案件名   = CStr(ThisWorkbook.Names("rng_Sel_案件名").RefersToRange.Value)
    pd.案件番号 = CStr(ThisWorkbook.Names("rng_Sel_案件番号").RefersToRange.Value)
    pd.顧客名   = CStr(ThisWorkbook.Names("rng_Sel_顧客名").RefersToRange.Value)
    pd.担当者名 = CStr(ThisWorkbook.Names("rng_Sel_担当者名").RefersToRange.Value)
    pd.期日     = CStr(ThisWorkbook.Names("rng_Sel_期日").RefersToRange.Value)

    Set GetSelectedProject = pd
    Exit Function
ErrHandler:
    Set GetSelectedProject = pd
End Function

'-------------------------------------------------------------
' ClearSelectedProject: 選択中案件をクリアする
'-------------------------------------------------------------
Public Sub ClearSelectedProject()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_SEARCH)
    ws.Range("rng_Sel_案件名").Value    = ""
    ws.Range("rng_Sel_案件番号").Value  = ""
    ws.Range("rng_Sel_顧客名").Value    = ""
    ws.Range("rng_Sel_担当者名").Value  = ""
    ws.Range("rng_Sel_期日").Value      = ""
End Sub

'-------------------------------------------------------------
' NavigateToSearch: 案件検索シートへ移動する
'-------------------------------------------------------------
Public Sub NavigateToSearch()
    If SheetExists(SHEET_SEARCH) Then
        ThisWorkbook.Sheets(SHEET_SEARCH).Activate
        ThisWorkbook.Sheets(SHEET_SEARCH).Range("B2").Select
    End If
End Sub

'-------------------------------------------------------------
' NavigateToTemplateList: テンプレート一覧シートへ移動する
'-------------------------------------------------------------
Public Sub NavigateToTemplateList()
    If SheetExists(SHEET_TEMPLATE_LIST) Then
        ThisWorkbook.Sheets(SHEET_TEMPLATE_LIST).Activate
    End If
End Sub
