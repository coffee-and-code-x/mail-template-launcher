Attribute VB_Name = "Module_FileIO"
Option Explicit

'=============================================================
' Module_FileIO: 外部Excelファイル読み込みモジュール
' ファイル設定シートの情報をもとに外部案件データを読み込む
' 外部ファイルは ReadOnly で開き、読み込み後に必ず閉じる
'=============================================================

'-------------------------------------------------------------
' GetAllActiveFileSettings: 有効な全ファイル設定を取得する
'-------------------------------------------------------------
Public Function GetAllActiveFileSettings() As Collection
    Dim col As New Collection
    On Error GoTo ErrHandler

    If Not SheetExists(SHEET_FILE_CONFIG) Then
        Set GetAllActiveFileSettings = col
        Exit Function
    End If

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_FILE_CONFIG)

    ' ヘッダー行(3)、注意書き行(4) を除いて5行目以降を処理
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim i As Long
    For i = 5 To lastRow
        ' A列にIDが入っていない行はスキップ
        If IsEmpty(ws.Cells(i, 1).Value) Or ws.Cells(i, 1).Value = "" Then GoTo NextRow
        ' L列（有効フラグ）が "○" のものだけ取得
        If Trim(CStr(ws.Cells(i, 12).Value)) <> "○" Then GoTo NextRow

        Dim setting As FileSetting
        setting = ReadFileSetting(ws, i)
        col.Add setting

NextRow:
    Next i

    Set GetAllActiveFileSettings = col
    Exit Function
ErrHandler:
    LogError "GetAllActiveFileSettings", Err.Number, Err.Description
    Set GetAllActiveFileSettings = col
End Function

'-------------------------------------------------------------
' ReadFileSetting: ファイル設定シートの1行を FileSetting 構造体に読み込む
'-------------------------------------------------------------
Private Function ReadFileSetting(ws As Worksheet, rowNum As Long) As FileSetting
    Dim s As FileSetting
    s.SettingID     = CLng(ws.Cells(rowNum, 1).Value)
    s.DisplayName   = CStr(ws.Cells(rowNum, 2).Value)
    s.FilePath      = CStr(ws.Cells(rowNum, 3).Value)
    s.SheetName     = CStr(ws.Cells(rowNum, 4).Value)
    s.HeaderRow     = IIf(IsNumeric(ws.Cells(rowNum, 5).Value), _
                          CLng(ws.Cells(rowNum, 5).Value), 1)
    s.Col_案件名    = ColNumOrLetter(CStr(ws.Cells(rowNum, 6).Value))
    s.Col_案件番号  = ColNumOrLetter(CStr(ws.Cells(rowNum, 7).Value))
    s.Col_顧客名    = ColNumOrLetter(CStr(ws.Cells(rowNum, 8).Value))
    s.Col_担当者名  = ColNumOrLetter(CStr(ws.Cells(rowNum, 9).Value))
    s.Col_期日      = ColNumOrLetter(CStr(ws.Cells(rowNum, 10).Value))
    s.SearchColStr  = CStr(ws.Cells(rowNum, 11).Value)
    s.IsActive      = (Trim(CStr(ws.Cells(rowNum, 12).Value)) = "○")
    ReadFileSetting = s
End Function

'-------------------------------------------------------------
' GetFileSettingByID: 指定IDのファイル設定を取得する
'-------------------------------------------------------------
Public Function GetFileSettingByID(settingID As Long) As FileSetting
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_FILE_CONFIG)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim i As Long
    For i = 5 To lastRow
        If IsNumeric(ws.Cells(i, 1).Value) Then
            If CLng(ws.Cells(i, 1).Value) = settingID Then
                GetFileSettingByID = ReadFileSetting(ws, i)
                Exit Function
            End If
        End If
    Next i

    ' 見つからない場合は空の構造体を返す
    Dim empty As FileSetting
    GetFileSettingByID = empty
End Function

'-------------------------------------------------------------
' SearchInFile: 指定のファイル設定でキーワード検索を実行する
' 戻り値: ProjectData の Collection
'-------------------------------------------------------------
Public Function SearchInFile(setting As FileSetting, keyword As String) As Collection
    Dim results As New Collection
    On Error GoTo ErrHandler

    ' ファイルの存在確認
    If setting.FilePath = "" Then
        Set SearchInFile = results
        Exit Function
    End If
    If Dir(setting.FilePath) = "" Then
        LogError "SearchInFile", 0, "ファイルが見つかりません: " & setting.FilePath
        ' 警告用のダミーレコードを追加
        Dim warnPd As ProjectData
        warnPd.案件名 = "⚠ ファイルが見つかりません: " & setting.DisplayName
        warnPd.SourceFile = setting.FilePath
        results.Add warnPd
        Set SearchInFile = results
        Exit Function
    End If

    ' 外部ファイルを読み取り専用で開く
    Application.StatusBar = "検索中: " & setting.DisplayName & " ..."
    Dim wb As Workbook
    Dim bNewlyOpened As Boolean
    bNewlyOpened = False

    ' すでに開いている場合はそれを使用
    On Error Resume Next
    Set wb = Workbooks(Dir(setting.FilePath))
    On Error GoTo ErrHandler

    If wb Is Nothing Then
        Set wb = Workbooks.Open(Filename:=setting.FilePath, _
                                ReadOnly:=True, _
                                UpdateLinks:=False, _
                                AddToMRU:=False)
        bNewlyOpened = True
    End If

    ' 対象シートを取得
    Dim srcWs As Worksheet
    On Error Resume Next
    Set srcWs = wb.Sheets(setting.SheetName)
    On Error GoTo CloseAndError

    If srcWs Is Nothing Then
        LogError "SearchInFile", 0, "シートが見つかりません: " & setting.SheetName & _
                 " (" & setting.DisplayName & ")"
        Dim sheetWarnPd As ProjectData
        sheetWarnPd.案件名 = "⚠ シート「" & setting.SheetName & "」が見つかりません: " & setting.DisplayName
        sheetWarnPd.SourceFile = setting.FilePath
        results.Add sheetWarnPd
        GoTo CloseFile
    End If

    ' 検索対象列の解析
    Dim searchCols() As Long
    searchCols = ParseSearchColumns(setting)

    ' データ行を検索
    Dim lastRow As Long
    lastRow = srcWs.Cells(srcWs.Rows.Count, 1).End(xlUp).Row

    Dim maxResults As Long
    maxResults = CLng(GetConfig(CFG_MAX_RESULTS))
    If maxResults <= 0 Then maxResults = 100

    Dim dataRow As Long
    For dataRow = setting.HeaderRow + 1 To lastRow
        ' 最大件数チェック
        If results.Count >= maxResults Then Exit For

        ' 空行スキップ
        Dim allEmpty As Boolean
        allEmpty = True
        Dim sc As Integer
        For sc = 0 To UBound(searchCols)
            If searchCols(sc) > 0 Then
                If Trim(SafeStr(srcWs.Cells(dataRow, searchCols(sc)).Value)) <> "" Then
                    allEmpty = False
                    Exit For
                End If
            End If
        Next sc
        If allEmpty Then GoTo NextDataRow

        ' キーワード検索（空キーワードは全件返す）
        Dim isMatch As Boolean
        isMatch = (Trim(keyword) = "")

        If Not isMatch Then
            For sc = 0 To UBound(searchCols)
                If searchCols(sc) > 0 Then
                    Dim cellVal As String
                    cellVal = SafeStr(srcWs.Cells(dataRow, searchCols(sc)).Value)
                    If InStr(1, cellVal, keyword, vbTextCompare) > 0 Then
                        isMatch = True
                        Exit For
                    End If
                End If
            Next sc
        End If

        ' 一致した場合は ProjectData に格納
        If isMatch Then
            Dim pd As ProjectData
            pd.案件名    = IIf(setting.Col_案件名 > 0, SafeStr(srcWs.Cells(dataRow, setting.Col_案件名).Value), "")
            pd.案件番号  = IIf(setting.Col_案件番号 > 0, SafeStr(srcWs.Cells(dataRow, setting.Col_案件番号).Value), "")
            pd.顧客名    = IIf(setting.Col_顧客名 > 0, SafeStr(srcWs.Cells(dataRow, setting.Col_顧客名).Value), "")
            pd.担当者名  = IIf(setting.Col_担当者名 > 0, SafeStr(srcWs.Cells(dataRow, setting.Col_担当者名).Value), "")
            pd.期日      = IIf(setting.Col_期日 > 0, SafeStr(srcWs.Cells(dataRow, setting.Col_期日).Value), "")
            pd.SourceFile = setting.DisplayName
            pd.SourceRow  = dataRow
            results.Add pd
        End If
NextDataRow:
    Next dataRow

CloseFile:
    ' 新たに開いたファイルのみ閉じる
    If bNewlyOpened And Not wb Is Nothing Then
        wb.Close SaveChanges:=False
    End If

    Application.StatusBar = False
    Set SearchInFile = results
    Exit Function

CloseAndError:
    If bNewlyOpened And Not wb Is Nothing Then
        On Error Resume Next
        wb.Close SaveChanges:=False
    End If
    LogError "SearchInFile", Err.Number, Err.Description
    Application.StatusBar = False
    Set SearchInFile = results
    Exit Function

ErrHandler:
    LogError "SearchInFile", Err.Number, Err.Description
    Application.StatusBar = False
    Set SearchInFile = results
End Function

'-------------------------------------------------------------
' ParseSearchColumns: SearchColStr から検索列番号配列を生成する
'-------------------------------------------------------------
Private Function ParseSearchColumns(setting As FileSetting) As Long()
    Dim cols() As Long

    If Trim(setting.SearchColStr) = "" Then
        ' 未設定の場合は定義済み全列を対象にする
        ReDim cols(4)
        cols(0) = setting.Col_案件名
        cols(1) = setting.Col_案件番号
        cols(2) = setting.Col_顧客名
        cols(3) = setting.Col_担当者名
        cols(4) = setting.Col_期日
    Else
        ' カンマ区切りの文字列をパース
        Dim parts() As String
        parts = Split(setting.SearchColStr, ",")
        ReDim cols(UBound(parts))
        Dim i As Integer
        For i = 0 To UBound(parts)
            cols(i) = ColNumOrLetter(Trim(parts(i)))
        Next i
    End If

    ParseSearchColumns = cols
End Function

'-------------------------------------------------------------
' TestConnection: ファイル設定の接続テストを実行する
'-------------------------------------------------------------
Public Sub TestConnection(settingID As Long)
    Dim setting As FileSetting
    setting = GetFileSettingByID(settingID)

    If setting.FilePath = "" Then
        MsgBox "設定ID " & settingID & " が見つかりません。", vbExclamation, "接続テスト"
        Exit Sub
    End If

    If Dir(setting.FilePath) = "" Then
        MsgBox "ファイルが見つかりません:" & vbCrLf & setting.FilePath, _
               vbCritical, "接続テスト - 失敗"
        Exit Sub
    End If

    On Error GoTo TestError
    Application.StatusBar = "接続テスト中: " & setting.DisplayName

    Dim wb As Workbook
    Dim bNewlyOpened As Boolean
    bNewlyOpened = False

    On Error Resume Next
    Set wb = Workbooks(Dir(setting.FilePath))
    On Error GoTo TestError

    If wb Is Nothing Then
        Set wb = Workbooks.Open(Filename:=setting.FilePath, ReadOnly:=True, _
                                UpdateLinks:=False, AddToMRU:=False)
        bNewlyOpened = True
    End If

    ' シートの確認
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets(setting.SheetName)
    On Error GoTo TestError

    Dim msg As String
    If ws Is Nothing Then
        msg = "⚠ シート「" & setting.SheetName & "」が見つかりません。" & vbCrLf & vbCrLf & _
              "【利用可能なシート一覧】" & vbCrLf
        Dim sh As Object
        For Each sh In wb.Sheets
            msg = msg & "  ・" & sh.Name & vbCrLf
        Next sh
        If bNewlyOpened Then wb.Close SaveChanges:=False
        Application.StatusBar = False
        MsgBox msg, vbExclamation, "接続テスト - シートエラー"
        Exit Sub
    End If

    ' ヘッダー行の読み取り
    Dim headerInfo As String
    headerInfo = ""
    Dim fieldDefs As Variant
    fieldDefs = Array( _
        Array("案件名", setting.Col_案件名), _
        Array("案件番号", setting.Col_案件番号), _
        Array("顧客名", setting.Col_顧客名), _
        Array("担当者名", setting.Col_担当者名), _
        Array("期日", setting.Col_期日) _
    )

    Dim j As Integer
    For j = 0 To UBound(fieldDefs)
        Dim colNum As Long
        colNum = fieldDefs(j)(1)
        If colNum > 0 Then
            Dim headerVal As String
            headerVal = SafeStr(ws.Cells(setting.HeaderRow, colNum).Value)
            headerInfo = headerInfo & "  " & fieldDefs(j)(0) & ": 列" & colNum & _
                         " = 「" & headerVal & "」" & vbCrLf
        Else
            headerInfo = headerInfo & "  " & fieldDefs(j)(0) & ": （未設定）" & vbCrLf
        End If
    Next j

    ' データ行数の確認
    Dim dataRowCount As Long
    dataRowCount = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row - setting.HeaderRow

    If bNewlyOpened Then wb.Close SaveChanges:=False
    Application.StatusBar = False

    MsgBox "✔ 接続テスト成功: " & setting.DisplayName & vbCrLf & vbCrLf & _
           "【列マッピング確認】" & vbCrLf & headerInfo & vbCrLf & _
           "データ行数: 約 " & dataRowCount & " 行", _
           vbInformation, "接続テスト - 成功"
    Exit Sub

TestError:
    If bNewlyOpened Then
        On Error Resume Next
        wb.Close SaveChanges:=False
    End If
    Application.StatusBar = False
    LogError "TestConnection", Err.Number, Err.Description
    MsgBox "接続テスト中にエラーが発生しました。" & vbCrLf & _
           "エラー " & Err.Number & ": " & Err.Description, vbCritical, "接続テスト - エラー"
End Sub

'-------------------------------------------------------------
' BrowseFilePath: ファイルパス参照ダイアログを表示し設定行に書き込む
'-------------------------------------------------------------
Public Sub BrowseFilePath(settingID As Long)
    Dim filePath As Variant
    filePath = Application.GetOpenFilename( _
        FileFilter:="Excelファイル (*.xlsx;*.xlsm;*.xls),*.xlsx;*.xlsm;*.xls", _
        Title:="案件データファイルを選択してください", _
        MultiSelect:=False)

    If filePath = False Then Exit Sub  ' キャンセル

    ' ファイル設定シートの対応行を更新
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_FILE_CONFIG)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim i As Long
    For i = 5 To lastRow
        If IsNumeric(ws.Cells(i, 1).Value) Then
            If CLng(ws.Cells(i, 1).Value) = settingID Then
                ws.Cells(i, 3).Value = CStr(filePath)
                ws.Cells(i, 3).Font.Color = RGB(0, 0, 0)
                ws.Cells(i, 3).Font.Italic = False
                MsgBox "ファイルパスを設定しました:" & vbCrLf & CStr(filePath), _
                       vbInformation, "ファイルパス設定"
                Exit Sub
            End If
        End If
    Next i

    MsgBox "設定ID " & settingID & " が見つかりませんでした。", vbExclamation, "ファイルパス設定"
End Sub

'-------------------------------------------------------------
' AddFileConfigRow: ファイル設定シートに新しい設定行を追加する
'-------------------------------------------------------------
Public Sub AddFileConfigRow()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_FILE_CONFIG)

    ' 最終行の次に追加
    Dim newRow As Long
    newRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If newRow < 5 Then newRow = 5

    ' 新しい設定IDを決定
    Dim newID As Long
    newID = 1
    Dim i As Long
    For i = 5 To newRow - 1
        If IsNumeric(ws.Cells(i, 1).Value) Then
            If CLng(ws.Cells(i, 1).Value) >= newID Then
                newID = CLng(ws.Cells(i, 1).Value) + 1
            End If
        End If
    Next i

    ' 行データの初期値を設定
    ws.Cells(newRow, 1).Value = newID
    ws.Cells(newRow, 2).Value = "新しい設定 " & newID
    ws.Cells(newRow, 3).Value = ""
    ws.Cells(newRow, 4).Value = "Sheet1"
    ws.Cells(newRow, 5).Value = 1
    ws.Cells(newRow, 6).Value = 1
    ws.Cells(newRow, 7).Value = 2
    ws.Cells(newRow, 8).Value = 3
    ws.Cells(newRow, 9).Value = 4
    ws.Cells(newRow, 10).Value = 5
    ws.Cells(newRow, 11).Value = "1,2,3"
    ws.Cells(newRow, 12).Value = "×"
    ws.Rows(newRow).RowHeight = 25

    ' データ入力規則
    With ws.Cells(newRow, 12).Validation
        .Delete
        .Add Type:=xlValidateList, Formula1:="○,×"
        .ShowError = False
    End With

    ' ボタンを追加（ID が 20 以下の場合のみ Module_ButtonHandlers に対応するボタンがある）
    If newID <= 20 Then
        AddButtonToCell ws, "M" & newRow, "参照...", "BrowseFile_" & newID
        AddButtonToCell ws, "N" & newRow, "テスト", "TestConn_" & newID
    Else
        ws.Cells(newRow, 13).Value = "（参照は設定シートで手動入力）"
        ws.Cells(newRow, 13).Font.Color = RGB(128, 128, 128)
        ws.Cells(newRow, 13).Font.Size = 9
    End If

    ' 追加した行にスクロール
    ws.Cells(newRow, 1).Select
    MsgBox "設定行 (ID=" & newID & ") を追加しました。" & vbCrLf & _
           "ファイルパスの「参照...」ボタンでファイルを選択してください。", _
           vbInformation, "設定行追加"
End Sub
