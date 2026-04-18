Attribute VB_Name = "Module_Init"
Option Explicit

'=============================================================
' Module_Init: ワークブック初期化モジュール
' シート構造・書式・サンプルデータ・名前付き範囲を一括セットアップする
' InitializeWorkbook() を初回に一度だけ実行すればよい
'=============================================================

Private Const INIT_FLAG As String = "INITIALIZED_V1"

'-------------------------------------------------------------
' InitializeWorkbook: ワークブック全体を初期化する（初回のみ）
' Workbook_Open から呼ばれる。すでに初期化済みの場合はスキップ。
'-------------------------------------------------------------
Public Sub InitializeWorkbook()
    ' 既に初期化済みか確認
    If IsAlreadyInitialized() Then
        ' 名前付き範囲だけ再作成（削除された場合の復旧）
        EnsureNamedRanges
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    On Error GoTo InitError

    ' --- シート作成 ---
    CreateInternalSheet
    CreateSettingsSheet
    CreateErrorLogSheet
    CreateTemplateListSheet
    CreateSearchSheet
    CreateFileConfigSheet
    CreateSampleBodySheet

    ' --- 名前付き範囲の設定 ---
    EnsureNamedRanges

    ' --- 初期化完了フラグを設定 ---
    ThisWorkbook.Sheets(SHEET_INTERNAL).Range("A1").Value = INIT_FLAG

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    ' テンプレート一覧を最初に表示
    ThisWorkbook.Sheets(SHEET_TEMPLATE_LIST).Activate
    MsgBox "Mail Template Launcher の初期化が完了しました。" & vbCrLf & vbCrLf & _
           "■ 使い方" & vbCrLf & _
           "1. 「ファイル設定」シートで案件データファイルを登録してください" & vbCrLf & _
           "2. 「新規テンプレート追加」ボタンでテンプレートを作成してください" & vbCrLf & _
           "3. 「案件を検索」ボタンで案件を選択してからメールを起動してください", _
           vbInformation, "初期化完了"
    Exit Sub

InitError:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    LogError "InitializeWorkbook", Err.Number, Err.Description
    MsgBox "初期化中にエラーが発生しました。" & vbCrLf & _
           "エラー " & Err.Number & ": " & Err.Description, vbCritical, "初期化エラー"
End Sub

'-------------------------------------------------------------
' IsAlreadyInitialized: 初期化済みかチェックする
'-------------------------------------------------------------
Private Function IsAlreadyInitialized() As Boolean
    On Error Resume Next
    If Not SheetExists(SHEET_INTERNAL) Then
        IsAlreadyInitialized = False
        Exit Function
    End If
    IsAlreadyInitialized = (ThisWorkbook.Sheets(SHEET_INTERNAL).Range("A1").Value = INIT_FLAG)
End Function

'-------------------------------------------------------------
' CreateInternalSheet: 内部データシートを作成する（非表示）
'-------------------------------------------------------------
Private Sub CreateInternalSheet()
    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(SHEET_INTERNAL)
    ws.Visible = xlSheetVeryHidden

    ws.Range("A1").Value = ""           ' 初期化フラグ（後で設定）
    ws.Range("A2").Value = "次テンプレートID"
    ws.Range("B2").Value = 0            ' AddNewTemplate が +1 して使う
    ws.Range("A1").Font.Size = 9
End Sub

'-------------------------------------------------------------
' CreateSettingsSheet: 設定シートを作成する
'-------------------------------------------------------------
Private Sub CreateSettingsSheet()
    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(SHEET_SETTINGS)

    ws.Cells.Clear
    ws.Name = SHEET_SETTINGS
    ws.Tab.Color = RGB(255, 192, 0)     ' オレンジタブ

    ' タイトル
    ws.Range("A1").Value = "設定"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14
    ws.Range("A1:D1").Merge
    ws.Range("A1").Interior.Color = RGB(255, 192, 0)

    ' ヘッダー行
    ws.Range("A2").Value = "設定キー"
    ws.Range("B2").Value = "値"
    ws.Range("C2").Value = "説明"
    FormatHeaderRow ws, 2, 3, RGB(255, 192, 0)
    ws.Range("A2:C2").Font.Color = RGB(0, 0, 0)

    ' 設定値
    Dim data As Variant
    data = Array( _
        Array(CFG_DATE_FORMAT, "yyyy/mm/dd", "日付の表示形式（例: yyyy/mm/dd または yyyy年m月d日）"), _
        Array(CFG_MAX_RESULTS, "100", "案件検索の最大表示件数"), _
        Array("検索後に案件検索シートへ移動", "TRUE", "検索実行後に案件検索シートに自動移動する（TRUE/FALSE）"), _
        Array("Outlookパス", "", "Office 365版など特定のOutlookを使う場合にパスを指定（例: C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE）"), _
        Array("Outlook起動待機秒数", "5", "Outlookパス指定時、起動完了まで待機する最大秒数") _
    )

    Dim i As Integer
    For i = 0 To UBound(data)
        ws.Cells(i + 3, 1).Value = data(i)(0)
        ws.Cells(i + 3, 2).Value = data(i)(1)
        ws.Cells(i + 3, 3).Value = data(i)(2)
        ws.Cells(i + 3, 3).Font.Color = RGB(128, 128, 128)
    Next i

    ' 列幅調整
    ws.Columns("A").ColumnWidth = 30
    ws.Columns("B").ColumnWidth = 20
    ws.Columns("C").ColumnWidth = 50

    ' 設定初期化ボタン
    AddButtonToCell ws, "E3", "設定を初期化", "Module_Init.ResetSettings"
    ' エラーログ表示ボタン
    AddButtonToCell ws, "E5", "エラーログを表示", "Module_Init.ShowErrorLog"
    ' 再初期化ボタン（危険なので注意書き付き）
    ws.Range("E8").Value = "※ワークブックを再初期化する場合:"
    ws.Range("E8").Font.Color = RGB(180, 0, 0)
    ws.Range("E8").Font.Size = 9
    AddButtonToCell ws, "E9", "ワークブックを再初期化", "Module_Init.ForceReinitialize"
End Sub

'-------------------------------------------------------------
' CreateErrorLogSheet: エラーログシートを作成する（非表示）
'-------------------------------------------------------------
Private Sub CreateErrorLogSheet()
    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(SHEET_ERROR_LOG)
    ws.Visible = xlSheetVeryHidden

    ws.Range("A1").Value = "タイムスタンプ"
    ws.Range("B1").Value = "処理名"
    ws.Range("C1").Value = "エラー番号"
    ws.Range("D1").Value = "エラーメッセージ"
    FormatHeaderRow ws, 1, 4, RGB(220, 80, 80)

    ws.Columns("A").ColumnWidth = 22
    ws.Columns("B").ColumnWidth = 30
    ws.Columns("C").ColumnWidth = 12
    ws.Columns("D").ColumnWidth = 60
End Sub

'-------------------------------------------------------------
' CreateTemplateListSheet: テンプレート一覧シートを作成する
'-------------------------------------------------------------
Private Sub CreateTemplateListSheet()
    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(SHEET_TEMPLATE_LIST)

    ws.Cells.Clear
    ws.Tab.Color = RGB(68, 114, 196)    ' 青タブ

    ' タイトル行
    ws.Range("A1").Value = "Mail Template Launcher"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 16
    ws.Range("A1:I1").Merge
    ws.Range("A1").Interior.Color = RGB(68, 114, 196)
    ws.Range("A1").Font.Color = RGB(255, 255, 255)
    ws.Range("A1").RowHeight = 32

    ' 操作ボタン行
    ws.Range("A2").RowHeight = 30
    AddButtonToCell ws, "A2", "案件を検索", "Module_Search.NavigateToSearch"
    AddButtonToCell ws, "C2", "新規テンプレート追加", "Module_Launcher.AddNewTemplate"
    AddButtonToCell ws, "E2", "ファイル設定を開く", "Module_Init.NavigateToFileConfig"

    ' マクロ有効化の案内（マクロ無効時には見えるが、有効時はVBAで非表示化）
    ws.Range("G2").Value = "✔ マクロ有効"
    ws.Range("G2").Font.Color = RGB(0, 130, 0)
    ws.Range("G2").Font.Bold = True

    ' ヘッダー行（3行目）
    ws.Cells(3, 1).Value = "ID"
    ws.Cells(3, 2).Value = "テンプレート名"
    ws.Cells(3, 3).Value = "形式"
    ws.Cells(3, 4).Value = "宛先 (To)"
    ws.Cells(3, 5).Value = "CC"
    ws.Cells(3, 6).Value = "件名"
    ws.Cells(3, 7).Value = "本文シート"
    ws.Cells(3, 8).Value = "最終更新"
    ws.Cells(3, 9).Value = "起動"
    FormatHeaderRow ws, 3, 9

    ' 列幅設定
    ws.Columns("A").ColumnWidth = 5
    ws.Columns("B").ColumnWidth = 22
    ws.Columns("C").ColumnWidth = 7
    ws.Columns("D").ColumnWidth = 25
    ws.Columns("E").ColumnWidth = 20
    ws.Columns("F").ColumnWidth = 30
    ws.Columns("G").ColumnWidth = 12
    ws.Columns("H").ColumnWidth = 18
    ws.Columns("I").ColumnWidth = 10

    ' 行高さ
    ws.Rows("3").RowHeight = 22

    ' ウィンドウ枠の固定（ヘッダー行）
    ws.Activate
    ws.Range("A4").Select
    ActiveWindow.FreezePanes = True

    ' サンプルテンプレートを追加
    AddSampleTemplate ws
End Sub

'-------------------------------------------------------------
' AddSampleTemplate: サンプルテンプレートを1件追加する
'-------------------------------------------------------------
Private Sub AddSampleTemplate(ws As Worksheet)
    Dim templateID As Long
    templateID = 1
    ThisWorkbook.Sheets(SHEET_INTERNAL).Range("B2").Value = 1

    Dim row As Long
    row = 4

    ws.Cells(row, 1).Value = templateID
    ws.Cells(row, 2).Value = "見積送付メール（サンプル）"
    ws.Cells(row, 3).Value = "HTML"
    ws.Cells(row, 4).Value = "{担当者メール}"
    ws.Cells(row, 5).Value = ""
    ws.Cells(row, 6).Value = "【{案件名}】お見積書のご送付"
    ws.Cells(row, 7).Value = "本文_1"
    ws.Cells(row, 8).Value = Now()
    ws.Cells(row, 8).NumberFormat = "yyyy/mm/dd hh:mm"

    ' 起動ボタン
    AddButtonToCell ws, "I" & row, "起動", "Launch_1"

    ' 行の高さ
    ws.Rows(row).RowHeight = 25

    ' サンプル本文シートを作成
    CreateBodySheet 1, "見積送付メール（サンプル）", True
End Sub

'-------------------------------------------------------------
' CreateSearchSheet: 案件検索シートを作成する
'-------------------------------------------------------------
Private Sub CreateSearchSheet()
    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(SHEET_SEARCH)

    ws.Cells.Clear
    ws.Tab.Color = RGB(0, 176, 80)      ' 緑タブ

    ' タイトル
    ws.Range("A1").Value = "案件検索"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14
    ws.Range("A1:G1").Merge
    ws.Range("A1").Interior.Color = RGB(0, 176, 80)
    ws.Range("A1").Font.Color = RGB(255, 255, 255)
    ws.Range("A1").RowHeight = 28

    ' 検索入力エリア
    ws.Range("A2").Value = "検索キーワード:"
    ws.Range("A2").Font.Bold = True
    ws.Range("B2").Value = ""           ' 入力セル
    ws.Range("B2").Interior.Color = RGB(255, 255, 200)  ' 薄い黄色
    ws.Range("B2:D2").Merge
    ws.Range("B2").Font.Size = 11
    ws.Range("A2:D2").RowHeight = 26

    ' 操作ボタン行
    ws.Range("A3").RowHeight = 30
    AddButtonToCell ws, "A3", "検索実行", "Module_Search.SearchProjects"
    AddButtonToCell ws, "C3", "この案件を選択", "Module_Search.SelectProject"
    AddButtonToCell ws, "E3", "テンプレート一覧へ", "Module_Search.NavigateToTemplateList"
    AddButtonToCell ws, "G3", "クリア", "Module_Search.ClearSearchResults"

    ' ステータス表示エリア
    ws.Range("A4").Value = ""           ' ステータスメッセージ表示用
    ws.Range("A4:G4").Merge
    ws.Range("A4").RowHeight = 22

    ' 検索結果ヘッダー行（5行目）
    ws.Cells(5, 1).Value = "案件名"
    ws.Cells(5, 2).Value = "案件番号"
    ws.Cells(5, 3).Value = "顧客名"
    ws.Cells(5, 4).Value = "担当者名"
    ws.Cells(5, 5).Value = "期日"
    ws.Cells(5, 6).Value = "ソースファイル"
    FormatHeaderRow ws, 5, 6, RGB(0, 176, 80)
    ws.Cells(5, 1).Font.Color = RGB(255, 255, 255)
    ws.Cells(5, 2).Font.Color = RGB(255, 255, 255)
    ws.Cells(5, 3).Font.Color = RGB(255, 255, 255)
    ws.Cells(5, 4).Font.Color = RGB(255, 255, 255)
    ws.Cells(5, 5).Font.Color = RGB(255, 255, 255)
    ws.Cells(5, 6).Font.Color = RGB(255, 255, 255)

    ' 列幅設定
    ws.Columns("A").ColumnWidth = 25
    ws.Columns("B").ColumnWidth = 15
    ws.Columns("C").ColumnWidth = 20
    ws.Columns("D").ColumnWidth = 15
    ws.Columns("E").ColumnWidth = 14
    ws.Columns("F").ColumnWidth = 40

    ' 選択中案件 表示エリア
    SetupSearchSelectionArea ws

    ' ウィンドウ枠の固定
    ws.Activate
    ws.Range("A6").Select
    ActiveWindow.FreezePanes = True
End Sub

'-------------------------------------------------------------
' SetupSearchSelectionArea: 案件検索シートの選択中案件表示エリアを設定する
'-------------------------------------------------------------
Private Sub SetupSearchSelectionArea(ws As Worksheet)
    ws.Range("A29").Value = "─────────────────────────────"
    ws.Range("A30").Value = "■ 選択中の案件"
    ws.Range("A30").Font.Bold = True
    ws.Range("A30").Font.Size = 11

    ws.Range("A31").Value = "案件名:"
    ws.Range("A32").Value = "案件番号:"
    ws.Range("A33").Value = "顧客名:"
    ws.Range("A34").Value = "担当者名:"
    ws.Range("A35").Value = "期日:"

    Dim labels As Variant
    labels = Array("A31", "A32", "A33", "A34", "A35")
    Dim r As Integer
    For r = 0 To UBound(labels)
        ws.Range(labels(r)).Font.Bold = True
        ws.Range(labels(r)).Font.Color = RGB(68, 114, 196)
    Next r

    ' B31:B35 は名前付き範囲で設定される（EnsureNamedRanges）
    ws.Range("B31:B35").Interior.Color = RGB(240, 248, 255)
End Sub

'-------------------------------------------------------------
' CreateFileConfigSheet: ファイル設定シートを作成する
'-------------------------------------------------------------
Private Sub CreateFileConfigSheet()
    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(SHEET_FILE_CONFIG)

    ws.Cells.Clear
    ws.Tab.Color = RGB(255, 102, 0)     ' オレンジタブ

    ' タイトル
    ws.Range("A1").Value = "外部ファイル設定"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14
    ws.Range("A1:N1").Merge
    ws.Range("A1").Interior.Color = RGB(255, 102, 0)
    ws.Range("A1").Font.Color = RGB(255, 255, 255)
    ws.Range("A1").RowHeight = 28

    ' 操作ボタン
    ws.Range("A2").RowHeight = 30
    AddButtonToCell ws, "A2", "設定行を追加", "Module_FileIO.AddFileConfigRow"
    AddButtonToCell ws, "C2", "テンプレート一覧へ", "Module_Init.NavigateToTemplateList"

    ' ヘッダー行（3行目）
    Dim headers As Variant
    headers = Array("ID", "表示名", "ファイルパス", "シート名", "ヘッダー行", _
                    "案件名列", "案件番号列", "顧客名列", "担当者名列", "期日列", _
                    "検索対象列(カンマ区切り)", "有効(○/×)", "参照", "接続テスト")
    Dim c As Integer
    For c = 0 To UBound(headers)
        ws.Cells(3, c + 1).Value = headers(c)
    Next c
    FormatHeaderRow ws, 3, 14, RGB(255, 102, 0)
    ws.Rows(3).Font.Color = RGB(255, 255, 255)

    ' 注意書き
    ws.Range("A4").Value = "※ 列番号は数字(例:3)または列記号(例:C)どちらでも入力できます。0または空白は未設定。"
    ws.Range("A4:N4").Merge
    ws.Range("A4").Font.Color = RGB(128, 128, 128)
    ws.Range("A4").Font.Size = 9
    ws.Range("A4").Font.Italic = True

    ' 列幅設定
    ws.Columns("A").ColumnWidth = 5
    ws.Columns("B").ColumnWidth = 18
    ws.Columns("C").ColumnWidth = 45
    ws.Columns("D").ColumnWidth = 15
    ws.Columns("E").ColumnWidth = 10
    ws.Columns("F").ColumnWidth = 10
    ws.Columns("G").ColumnWidth = 10
    ws.Columns("H").ColumnWidth = 10
    ws.Columns("I").ColumnWidth = 12
    ws.Columns("J").ColumnWidth = 10
    ws.Columns("K").ColumnWidth = 20
    ws.Columns("L").ColumnWidth = 10
    ws.Columns("M").ColumnWidth = 10
    ws.Columns("N").ColumnWidth = 12

    ' サンプル設定行を追加
    AddSampleFileConfigRow ws
End Sub

'-------------------------------------------------------------
' AddSampleFileConfigRow: ファイル設定のサンプル行を追加する
'-------------------------------------------------------------
Private Sub AddSampleFileConfigRow(ws As Worksheet)
    Dim row As Long
    row = 5

    ws.Cells(row, 1).Value = 1
    ws.Cells(row, 2).Value = "営業案件管理表（サンプル）"
    ws.Cells(row, 3).Value = "C:\案件データ\営業案件管理表.xlsx"
    ws.Cells(row, 4).Value = "案件一覧"
    ws.Cells(row, 5).Value = 1
    ws.Cells(row, 6).Value = 1      ' A列: 案件名
    ws.Cells(row, 7).Value = 2      ' B列: 案件番号
    ws.Cells(row, 8).Value = 3      ' C列: 顧客名
    ws.Cells(row, 9).Value = 4      ' D列: 担当者名
    ws.Cells(row, 10).Value = 5     ' E列: 期日
    ws.Cells(row, 11).Value = "1,2,3"
    ws.Cells(row, 12).Value = "×"
    ws.Cells(row, 3).Font.Color = RGB(128, 128, 128)
    ws.Cells(row, 3).Font.Italic = True

    ' 有効/無効のデータ入力規則
    With ws.Cells(row, 12).Validation
        .Delete
        .Add Type:=xlValidateList, Formula1:="○,×"
        .ShowError = False
    End With

    ' 参照・テストボタン
    AddButtonToCell ws, "M" & row, "参照...", "BrowseFile_1"
    AddButtonToCell ws, "N" & row, "テスト", "TestConn_1"
End Sub

'-------------------------------------------------------------
' CreateSampleBodySheet: サンプル本文シートを作成する
'-------------------------------------------------------------
Private Sub CreateSampleBodySheet()
    ' CreateBodySheet は AddSampleTemplate から呼ばれるので不要な場合は何もしない
End Sub

'-------------------------------------------------------------
' CreateBodySheet: 指定テンプレートの本文シートを作成する
'-------------------------------------------------------------
Public Sub CreateBodySheet(templateID As Long, templateName As String, _
                            Optional isHTML As Boolean = True)
    Dim sheetName As String
    sheetName = "本文_" & templateID

    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(sheetName)

    ws.Cells.Clear
    ws.Tab.Color = RGB(180, 198, 231)   ' 薄い青タブ

    ' テンプレート情報
    ws.Range("A1").Value = "テンプレートID:"
    ws.Range("B1").Value = templateID
    ws.Range("A2").Value = "テンプレート名:"
    ws.Range("B2").Value = templateName
    ws.Range("A1:B2").Font.Color = RGB(128, 128, 128)
    ws.Range("A1:B2").Font.Size = 9

    ' プレースホルダー説明
    ws.Range("A3").Value = "【利用可能なプレースホルダー】 " & _
        "{案件名}  {案件番号}  {顧客名}  {担当者名}  {期日}  {今日の日付}"
    ws.Range("A3:F3").Merge
    ws.Range("A3").Interior.Color = RGB(255, 255, 200)
    ws.Range("A3").Font.Color = RGB(128, 100, 0)
    ws.Range("A3").Font.Size = 9
    ws.Range("A3").Font.Italic = True
    ws.Range("A3").RowHeight = 20

    ' 本文入力エリア（A4から）
    ws.Range("A4").Value = GetSampleBodyContent(isHTML)
    ws.Range("A4").WrapText = True
    ws.Range("A4").VerticalAlignment = xlTop
    ws.Range("A4").Font.Size = 11
    ws.Range("A4:F30").Merge
    ws.Range("A4").RowHeight = 400

    ' 列幅
    ws.Columns("A").ColumnWidth = 80
End Sub

'-------------------------------------------------------------
' GetSampleBodyContent: サンプル本文テキストを返す
'-------------------------------------------------------------
Private Function GetSampleBodyContent(isHTML As Boolean) As String
    If isHTML Then
        GetSampleBodyContent = _
            "<p>{顧客名} {担当者名} 様</p>" & vbCrLf & _
            "<p>お世話になっております。<br>株式会社○○ 営業部の△△でございます。</p>" & vbCrLf & _
            "<p>下記の件につきまして、お見積書をご送付いたします。<br>ご確認のほど、よろしくお願いいたします。</p>" & vbCrLf & _
            "<ul>" & vbCrLf & _
            "  <li><strong>案件名:</strong> {案件名}</li>" & vbCrLf & _
            "  <li><strong>案件番号:</strong> {案件番号}</li>" & vbCrLf & _
            "  <li><strong>納期:</strong> {期日}</li>" & vbCrLf & _
            "</ul>" & vbCrLf & _
            "<p>ご不明な点がございましたら、お気軽にお申し付けください。</p>" & vbCrLf & _
            "<p>どうぞよろしくお願いいたします。</p>"
    Else
        GetSampleBodyContent = _
            "{顧客名} {担当者名} 様" & vbCrLf & vbCrLf & _
            "お世話になっております。" & vbCrLf & _
            "株式会社○○ 営業部の△△でございます。" & vbCrLf & vbCrLf & _
            "下記の件につきまして、お見積書をご送付いたします。" & vbCrLf & _
            "ご確認のほど、よろしくお願いいたします。" & vbCrLf & vbCrLf & _
            "■ 案件名:   {案件名}" & vbCrLf & _
            "■ 案件番号: {案件番号}" & vbCrLf & _
            "■ 納期:     {期日}" & vbCrLf & vbCrLf & _
            "ご不明な点がございましたら、お気軽にお申し付けください。" & vbCrLf & vbCrLf & _
            "どうぞよろしくお願いいたします。"
    End If
End Function

'-------------------------------------------------------------
' EnsureNamedRanges: 名前付き範囲を冪等に再作成する
' Workbook_Open から毎回呼ばれて名前が削除された場合も復旧する
'-------------------------------------------------------------
Public Sub EnsureNamedRanges()
    If Not SheetExists(SHEET_SEARCH) Then Exit Sub
    If Not SheetExists(SHEET_SETTINGS) Then Exit Sub

    SetNamedRange "rng_SearchKeyword",  SHEET_SEARCH,   "$B$2"
    SetNamedRange "rng_Sel_案件名",     SHEET_SEARCH,   "$B$31"
    SetNamedRange "rng_Sel_案件番号",   SHEET_SEARCH,   "$B$32"
    SetNamedRange "rng_Sel_顧客名",     SHEET_SEARCH,   "$B$33"
    SetNamedRange "rng_Sel_担当者名",   SHEET_SEARCH,   "$B$34"
    SetNamedRange "rng_Sel_期日",       SHEET_SEARCH,   "$B$35"
    SetNamedRange "cfg_DateFormat",     SHEET_SETTINGS, "$B$3"
    SetNamedRange "cfg_MaxResults",     SHEET_SETTINGS, "$B$4"
End Sub

'-------------------------------------------------------------
' SetNamedRange: 名前付き範囲を設定する（既存は削除して再作成）
'-------------------------------------------------------------
Private Sub SetNamedRange(rangeName As String, sheetName As String, cellAddr As String)
    On Error Resume Next
    Dim nm As Name
    Set nm = ThisWorkbook.Names(rangeName)
    If Not nm Is Nothing Then nm.Delete
    On Error GoTo 0

    If Not SheetExists(sheetName) Then Exit Sub
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)
    ThisWorkbook.Names.Add Name:=rangeName, RefersTo:=ws.Range(cellAddr)
End Sub

'-------------------------------------------------------------
' ResetSettings: 設定値をデフォルトに戻す
'-------------------------------------------------------------
Public Sub ResetSettings()
    If MsgBox("設定をデフォルト値にリセットしますか？", vbYesNo + vbQuestion, "設定初期化") = vbNo Then
        Exit Sub
    End If
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_SETTINGS)
    ws.Range("B3").Value = "yyyy/mm/dd"
    ws.Range("B4").Value = "100"
    ws.Range("B5").Value = "TRUE"
    MsgBox "設定をデフォルト値にリセットしました。", vbInformation, "設定初期化"
End Sub

'-------------------------------------------------------------
' ShowErrorLog: エラーログシートを表示する
'-------------------------------------------------------------
Public Sub ShowErrorLog()
    If Not SheetExists(SHEET_ERROR_LOG) Then
        MsgBox "エラーログが見つかりません。", vbInformation, "エラーログ"
        Exit Sub
    End If
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_ERROR_LOG)
    ws.Visible = xlSheetVisible
    ws.Activate
End Sub

'-------------------------------------------------------------
' ForceReinitialize: ワークブックを強制的に再初期化する
'-------------------------------------------------------------
Public Sub ForceReinitialize()
    Dim ans As Integer
    ans = MsgBox("ワークブックを再初期化します。" & vbCrLf & _
                 "テンプレートデータは保持されますが、シートが再作成されます。" & vbCrLf & vbCrLf & _
                 "本当に実行しますか？", vbYesNo + vbCritical, "再初期化の確認")
    If ans = vbNo Then Exit Sub

    ' 初期化フラグをクリア
    If SheetExists(SHEET_INTERNAL) Then
        ThisWorkbook.Sheets(SHEET_INTERNAL).Range("A1").Value = ""
    End If
    InitializeWorkbook
End Sub

'-------------------------------------------------------------
' NavigateToFileConfig: ファイル設定シートへ移動する
'-------------------------------------------------------------
Public Sub NavigateToFileConfig()
    If SheetExists(SHEET_FILE_CONFIG) Then
        ThisWorkbook.Sheets(SHEET_FILE_CONFIG).Activate
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
