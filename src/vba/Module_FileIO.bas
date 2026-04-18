Attribute VB_Name = "Module_FileIO"
Option Explicit

'=============================================================
' Module_FileIO: 外部Excelファイル読み込みモジュール
' ファイル設定シートの情報をもとに外部案件データを読み込む
' 外部ファイルは ReadOnly で開き、読み込み後に必ず閉じる
'=============================================================

'-------------------------------------------------------------
' GetAllActiveFileSettings: 有効な全ファイル設定を Collection で返す
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
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim i As Long
    For i = 5 To lastRow
        If IsValidConfigRow(ws, i) Then
            Dim s As CFileSetting
            Set s = ReadFileSetting(ws, i)
            col.Add s
        End If
    Next i

    Set GetAllActiveFileSettings = col
    Exit Function
ErrHandler:
    LogError "GetAllActiveFileSettings", Err.Number, Err.Description
    Set GetAllActiveFileSettings = col
End Function

'-------------------------------------------------------------
' IsValidConfigRow: 設定行が有効かどうかチェックする（内部用）
'-------------------------------------------------------------
Private Function IsValidConfigRow(ws As Worksheet, rowNum As Long) As Boolean
    If IsEmpty(ws.Cells(rowNum, 1).Value) Or ws.Cells(rowNum, 1).Value = "" Then
        IsValidConfigRow = False
        Exit Function
    End If
    IsValidConfigRow = (Trim(CStr(ws.Cells(rowNum, 12).Value)) = "○")
End Function

'-------------------------------------------------------------
' ReadFileSetting: ファイル設定シートの1行を CFileSetting に読み込む
'-------------------------------------------------------------
Private Function ReadFileSetting(ws As Worksheet, rowNum As Long) As CFileSetting
    Dim s As New CFileSetting
    s.SettingID    = CLng(ws.Cells(rowNum, 1).Value)
    s.DisplayName  = CStr(ws.Cells(rowNum, 2).Value)
    s.FilePath     = CStr(ws.Cells(rowNum, 3).Value)
    s.SheetName    = CStr(ws.Cells(rowNum, 4).Value)
    s.HeaderRow    = IIf(IsNumeric(ws.Cells(rowNum, 5).Value), CLng(ws.Cells(rowNum, 5).Value), 1)
    s.Col_案件名   = ColNumOrLetter(CStr(ws.Cells(rowNum, 6).Value))
    s.Col_案件番号 = ColNumOrLetter(CStr(ws.Cells(rowNum, 7).Value))
    s.Col_顧客名   = ColNumOrLetter(CStr(ws.Cells(rowNum, 8).Value))
    s.Col_担当者名 = ColNumOrLetter(CStr(ws.Cells(rowNum, 9).Value))
    s.Col_期日     = ColNumOrLetter(CStr(ws.Cells(rowNum, 10).Value))
    s.SearchColStr = CStr(ws.Cells(rowNum, 11).Value)
    s.IsActive     = (Trim(CStr(ws.Cells(rowNum, 12).Value)) = "○")
    Set ReadFileSetting = s
End Function

'-------------------------------------------------------------
' GetFileSettingByID: 指定IDのファイル設定を返す
'-------------------------------------------------------------
Public Function GetFileSettingByID(settingID As Long) As CFileSetting
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_FILE_CONFIG)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim i As Long
    For i = 5 To lastRow
        If IsNumeric(ws.Cells(i, 1).Value) Then
            If CLng(ws.Cells(i, 1).Value) = settingID Then
                Set GetFileSettingByID = ReadFileSetting(ws, i)
                Exit Function
            End If
        End If
    Next i
    Set GetFileSettingByID = Nothing
End Function

'-------------------------------------------------------------
' SearchInFile: 指定設定でキーワード検索を実行する
' 戻り値: CProjectData の Collection
'-------------------------------------------------------------
Public Function SearchInFile(setting As CFileSetting, keyword As String) As Collection
    Dim results As New Collection
    On Error GoTo ErrHandler

    If Not setting.IsValid() Then
        Set SearchInFile = results
        Exit Function
    End If

    If Not FileExists(setting.FilePath) Then
        results.Add MakeWarningProject("ファイルが見つかりません: " & setting.DisplayName, setting.FilePath)
        Set SearchInFile = results
        Exit Function
    End If

    Application.StatusBar = "検索中: " & setting.DisplayName & " ..."
    Dim wb As Workbook
    Dim bNewlyOpened As Boolean
    Set wb = OpenWorkbookReadOnly(setting.FilePath, bNewlyOpened)
    If wb Is Nothing Then
        results.Add MakeWarningProject("ファイルを開けません: " & setting.DisplayName, setting.FilePath)
        Set SearchInFile = results
        Exit Function
    End If

    Dim srcWs As Worksheet
    Set srcWs = GetWorksheet(wb, setting.SheetName)
    If srcWs Is Nothing Then
        LogError "SearchInFile", 0, "シートが見つかりません: " & setting.SheetName
        results.Add MakeWarningProject("シート「" & setting.SheetName & "」が見つかりません: " & setting.DisplayName, setting.FilePath)
        GoTo CloseAndReturn
    End If

    SearchRowsInSheet srcWs, setting, keyword, results

CloseAndReturn:
    If bNewlyOpened Then wb.Close SaveChanges:=False
    Application.StatusBar = False
    Set SearchInFile = results
    Exit Function

ErrHandler:
    LogError "SearchInFile", Err.Number, Err.Description
    If bNewlyOpened Then
        On Error Resume Next
        wb.Close SaveChanges:=False
    End If
    Application.StatusBar = False
    Set SearchInFile = results
End Function

'-------------------------------------------------------------
' FileExists: ファイルの存在確認（内部用）
'-------------------------------------------------------------
Private Function FileExists(path As String) As Boolean
    FileExists = (path <> "" And Dir(path) <> "")
End Function

'-------------------------------------------------------------
' OpenWorkbookReadOnly: 外部ファイルを読み取り専用で開く（内部用）
' 既に開いている場合はそれを返す
'-------------------------------------------------------------
Private Function OpenWorkbookReadOnly(filePath As String, ByRef bNewlyOpened As Boolean) As Workbook
    bNewlyOpened = False
    On Error Resume Next
    Dim wb As Workbook
    Set wb = Workbooks(Dir(filePath))
    On Error GoTo 0
    If wb Is Nothing Then
        On Error Resume Next
        Set wb = Workbooks.Open(Filename:=filePath, ReadOnly:=True, _
                                UpdateLinks:=False, AddToMRU:=False)
        On Error GoTo 0
        bNewlyOpened = True
    End If
    Set OpenWorkbookReadOnly = wb
End Function

'-------------------------------------------------------------
' GetWorksheet: ブック内のシートを名前で取得（内部用）
'-------------------------------------------------------------
Private Function GetWorksheet(wb As Workbook, sheetName As String) As Worksheet
    On Error Resume Next
    Set GetWorksheet = wb.Sheets(sheetName)
    On Error GoTo 0
End Function

'-------------------------------------------------------------
' SearchRowsInSheet: シートの全行をキーワード検索して結果を追加する（内部用）
'-------------------------------------------------------------
Private Sub SearchRowsInSheet(srcWs As Worksheet, setting As CFileSetting, _
                               keyword As String, results As Collection)
    Dim maxResults As Long
    maxResults = CLng(GetConfig(CFG_MAX_RESULTS))
    If maxResults <= 0 Then maxResults = 100

    Dim searchCols() As Long
    searchCols = setting.GetSearchColumns()

    Dim lastRow As Long
    lastRow = srcWs.Cells(srcWs.Rows.Count, 1).End(xlUp).Row

    Dim dataRow As Long
    For dataRow = setting.HeaderRow + 1 To lastRow
        If results.Count >= maxResults Then Exit For
        If IsRowEmpty(srcWs, dataRow, searchCols) Then GoTo NextRow
        If IsRowMatch(srcWs, dataRow, searchCols, keyword) Then
            results.Add BuildProjectData(srcWs, dataRow, setting)
        End If
NextRow:
    Next dataRow
End Sub

'-------------------------------------------------------------
' IsRowEmpty: 検索対象列が全て空かどうか確認する（内部用）
'-------------------------------------------------------------
Private Function IsRowEmpty(ws As Worksheet, rowNum As Long, searchCols() As Long) As Boolean
    Dim sc As Integer
    For sc = 0 To UBound(searchCols)
        If searchCols(sc) > 0 Then
            If Trim(SafeStr(ws.Cells(rowNum, searchCols(sc)).Value)) <> "" Then
                IsRowEmpty = False
                Exit Function
            End If
        End If
    Next sc
    IsRowEmpty = True
End Function

'-------------------------------------------------------------
' IsRowMatch: 行のいずれかの検索列にキーワードが含まれるか確認する（内部用）
' 空キーワードは全件一致とする
'-------------------------------------------------------------
Private Function IsRowMatch(ws As Worksheet, rowNum As Long, _
                             searchCols() As Long, keyword As String) As Boolean
    If Trim(keyword) = "" Then
        IsRowMatch = True
        Exit Function
    End If
    Dim sc As Integer
    For sc = 0 To UBound(searchCols)
        If searchCols(sc) > 0 Then
            If InStr(1, SafeStr(ws.Cells(rowNum, searchCols(sc)).Value), keyword, vbTextCompare) > 0 Then
                IsRowMatch = True
                Exit Function
            End If
        End If
    Next sc
    IsRowMatch = False
End Function

'-------------------------------------------------------------
' BuildProjectData: シートの行から CProjectData を生成する（内部用）
'-------------------------------------------------------------
Private Function BuildProjectData(ws As Worksheet, rowNum As Long, _
                                   setting As CFileSetting) As CProjectData
    Dim pd As New CProjectData
    pd.案件名   = IIf(setting.Col_案件名 > 0, SafeStr(ws.Cells(rowNum, setting.Col_案件名).Value), "")
    pd.案件番号 = IIf(setting.Col_案件番号 > 0, SafeStr(ws.Cells(rowNum, setting.Col_案件番号).Value), "")
    pd.顧客名   = IIf(setting.Col_顧客名 > 0, SafeStr(ws.Cells(rowNum, setting.Col_顧客名).Value), "")
    pd.担当者名 = IIf(setting.Col_担当者名 > 0, SafeStr(ws.Cells(rowNum, setting.Col_担当者名).Value), "")
    pd.期日     = IIf(setting.Col_期日 > 0, SafeStr(ws.Cells(rowNum, setting.Col_期日).Value), "")
    pd.SourceFile = setting.DisplayName
    pd.SourceRow  = rowNum
    Set BuildProjectData = pd
End Function

'-------------------------------------------------------------
' MakeWarningProject: 警告用のダミー CProjectData を生成する（内部用）
'-------------------------------------------------------------
Private Function MakeWarningProject(msg As String, filePath As String) As CProjectData
    Dim pd As New CProjectData
    pd.案件名    = "⚠ " & msg
    pd.SourceFile = filePath
    Set MakeWarningProject = pd
End Function

'-------------------------------------------------------------
' TestConnection: ファイル設定の接続テストを実行する
'-------------------------------------------------------------
Public Sub TestConnection(settingID As Long)
    Dim setting As CFileSetting
    Set setting = GetFileSettingByID(settingID)

    If setting Is Nothing Then
        MsgBox "設定ID " & settingID & " が見つかりません。", vbExclamation, "接続テスト"
        Exit Sub
    End If

    If Not FileExists(setting.FilePath) Then
        MsgBox "ファイルが見つかりません:" & vbCrLf & setting.FilePath, vbCritical, "接続テスト - 失敗"
        Exit Sub
    End If

    Application.StatusBar = "接続テスト中: " & setting.DisplayName
    Dim bNewlyOpened As Boolean
    Dim wb As Workbook
    Set wb = OpenWorkbookReadOnly(setting.FilePath, bNewlyOpened)

    If wb Is Nothing Then
        Application.StatusBar = False
        MsgBox "ファイルを開けませんでした: " & setting.FilePath, vbCritical, "接続テスト - 失敗"
        Exit Sub
    End If

    Dim srcWs As Worksheet
    Set srcWs = GetWorksheet(wb, setting.SheetName)

    If srcWs Is Nothing Then
        Dim sheetList As String
        sheetList = BuildSheetList(wb)
        If bNewlyOpened Then wb.Close SaveChanges:=False
        Application.StatusBar = False
        MsgBox "シート「" & setting.SheetName & "」が見つかりません。" & vbCrLf & vbCrLf & _
               "【利用可能なシート】" & vbCrLf & sheetList, vbExclamation, "接続テスト - シートエラー"
        Exit Sub
    End If

    Dim headerInfo As String
    headerInfo = BuildHeaderInfo(srcWs, setting)
    Dim dataRowCount As Long
    dataRowCount = srcWs.Cells(srcWs.Rows.Count, 1).End(xlUp).Row - setting.HeaderRow

    If bNewlyOpened Then wb.Close SaveChanges:=False
    Application.StatusBar = False

    MsgBox "✔ 接続テスト成功: " & setting.DisplayName & vbCrLf & vbCrLf & _
           "【列マッピング確認】" & vbCrLf & headerInfo & vbCrLf & _
           "データ行数: 約 " & dataRowCount & " 行", vbInformation, "接続テスト - 成功"
End Sub

'-------------------------------------------------------------
' BuildSheetList: ブック内の全シート名を箇条書き文字列で返す（内部用）
'-------------------------------------------------------------
Private Function BuildSheetList(wb As Workbook) As String
    Dim result As String
    Dim sh As Object
    For Each sh In wb.Sheets
        result = result & "  ・" & sh.Name & vbCrLf
    Next sh
    BuildSheetList = result
End Function

'-------------------------------------------------------------
' BuildHeaderInfo: 列マッピングの確認文字列を生成する（内部用）
'-------------------------------------------------------------
Private Function BuildHeaderInfo(ws As Worksheet, setting As CFileSetting) As String
    Dim result As String
    Dim fieldCols As Variant
    fieldCols = setting.GetFieldColumns()
    Dim j As Integer
    For j = 0 To UBound(fieldCols)
        Dim colNum As Long
        colNum = fieldCols(j)(1)
        If colNum > 0 Then
            Dim headerVal As String
            headerVal = SafeStr(ws.Cells(setting.HeaderRow, colNum).Value)
            result = result & "  " & fieldCols(j)(0) & ": 列" & colNum & _
                     " = 「" & headerVal & "」" & vbCrLf
        Else
            result = result & "  " & fieldCols(j)(0) & ": （未設定）" & vbCrLf
        End If
    Next j
    BuildHeaderInfo = result
End Function

'-------------------------------------------------------------
' BrowseFilePath: ファイルパス参照ダイアログを表示し設定行に書き込む
'-------------------------------------------------------------
Public Sub BrowseFilePath(settingID As Long)
    Dim filePath As Variant
    filePath = Application.GetOpenFilename( _
        FileFilter:="Excelファイル (*.xlsx;*.xlsm;*.xls),*.xlsx;*.xlsm;*.xls", _
        Title:="案件データファイルを選択してください", _
        MultiSelect:=False)

    If filePath = False Then Exit Sub

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

    Dim newRow As Long
    newRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If newRow < 5 Then newRow = 5

    Dim newID As Long
    newID = GetNextFileSettingID(ws, newRow)

    WriteDefaultConfigRow ws, newRow, newID
    AddConfigRowButtons ws, newRow, newID

    ws.Cells(newRow, 1).Select
    MsgBox "設定行 (ID=" & newID & ") を追加しました。" & vbCrLf & _
           "ファイルパスの「参照...」ボタンでファイルを選択してください。", _
           vbInformation, "設定行追加"
End Sub

'-------------------------------------------------------------
' GetNextFileSettingID: 次の設定IDを計算する（内部用）
'-------------------------------------------------------------
Private Function GetNextFileSettingID(ws As Worksheet, lastRow As Long) As Long
    Dim newID As Long
    newID = 1
    Dim i As Long
    For i = 5 To lastRow - 1
        If IsNumeric(ws.Cells(i, 1).Value) Then
            If CLng(ws.Cells(i, 1).Value) >= newID Then
                newID = CLng(ws.Cells(i, 1).Value) + 1
            End If
        End If
    Next i
    GetNextFileSettingID = newID
End Function

'-------------------------------------------------------------
' WriteDefaultConfigRow: デフォルト値で設定行を書き込む（内部用）
'-------------------------------------------------------------
Private Sub WriteDefaultConfigRow(ws As Worksheet, rowNum As Long, newID As Long)
    ws.Cells(rowNum, 1).Value = newID
    ws.Cells(rowNum, 2).Value = "新しい設定 " & newID
    ws.Cells(rowNum, 3).Value = ""
    ws.Cells(rowNum, 4).Value = "Sheet1"
    ws.Cells(rowNum, 5).Value = 1
    ws.Cells(rowNum, 6).Value = 1
    ws.Cells(rowNum, 7).Value = 2
    ws.Cells(rowNum, 8).Value = 3
    ws.Cells(rowNum, 9).Value = 4
    ws.Cells(rowNum, 10).Value = 5
    ws.Cells(rowNum, 11).Value = "1,2,3"
    ws.Cells(rowNum, 12).Value = "×"
    ws.Rows(rowNum).RowHeight = 25

    With ws.Cells(rowNum, 12).Validation
        .Delete
        .Add Type:=xlValidateList, Formula1:="○,×"
        .ShowError = False
    End With
End Sub

'-------------------------------------------------------------
' AddConfigRowButtons: 設定行にボタンを追加する（内部用）
'-------------------------------------------------------------
Private Sub AddConfigRowButtons(ws As Worksheet, rowNum As Long, newID As Long)
    If newID <= 20 Then
        AddButtonToCell ws, "M" & rowNum, "参照...", "BrowseFile_" & newID
        AddButtonToCell ws, "N" & rowNum, "テスト", "TestConn_" & newID
    Else
        ws.Cells(rowNum, 13).Value = "（参照は手動入力）"
        ws.Cells(rowNum, 13).Font.Color = RGB(128, 128, 128)
        ws.Cells(rowNum, 13).Font.Size = 9
    End If
End Sub
