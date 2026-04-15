Attribute VB_Name = "Module_Utils"
Option Explicit

'=============================================================
' Module_Utils: ユーティリティ関数モジュール
' 他モジュールから共通利用する汎用関数群
'=============================================================

' シート名定数
Public Const SHEET_TEMPLATE_LIST    As String = "テンプレート一覧"
Public Const SHEET_SEARCH           As String = "案件検索"
Public Const SHEET_FILE_CONFIG      As String = "ファイル設定"
Public Const SHEET_SETTINGS         As String = "設定"
Public Const SHEET_ERROR_LOG        As String = "エラーログ"
Public Const SHEET_INTERNAL         As String = "_内部データ"

' 設定キー定数
Public Const CFG_DATE_FORMAT        As String = "今日の日付フォーマット"
Public Const CFG_MAX_RESULTS        As String = "検索結果最大件数"

'-------------------------------------------------------------
' ColNumOrLetter: 列の指定を列番号(Long)に変換する
' 入力: "3" → 3, "C" → 3, "AA" → 27
'-------------------------------------------------------------
Public Function ColNumOrLetter(val As String) As Long
    Dim trimmed As String
    trimmed = Trim(val)
    If trimmed = "" Then
        ColNumOrLetter = 0
        Exit Function
    End If
    ' 数値文字列の場合はそのまま変換
    If IsNumeric(trimmed) Then
        ColNumOrLetter = CLng(trimmed)
        Exit Function
    End If
    ' 列記号の場合は番号に変換（A=1, B=2, ... Z=26, AA=27 ...）
    ColNumOrLetter = ColumnLetterToNumber(UCase(trimmed))
End Function

'-------------------------------------------------------------
' ColumnLetterToNumber: 列記号を列番号に変換（内部用）
'-------------------------------------------------------------
Private Function ColumnLetterToNumber(col As String) As Long
    Dim result As Long
    Dim i As Integer
    result = 0
    For i = 1 To Len(col)
        result = result * 26 + (Asc(Mid(col, i, 1)) - Asc("A") + 1)
    Next i
    ColumnLetterToNumber = result
End Function

'-------------------------------------------------------------
' SafeStr: セル値を安全に文字列変換する
' Null/Empty/エラー値 → "" を返す
' 日付型 → 設定された日付フォーマットで文字列化
'-------------------------------------------------------------
Public Function SafeStr(val As Variant) As String
    On Error GoTo ErrHandler
    If IsEmpty(val) Or IsNull(val) Then
        SafeStr = ""
        Exit Function
    End If
    If IsError(val) Then
        SafeStr = ""
        Exit Function
    End If
    ' 日付型（vbDate）の場合はフォーマット適用
    If VarType(val) = vbDate Then
        SafeStr = Format(val, GetConfig(CFG_DATE_FORMAT))
        Exit Function
    End If
    ' 文字列に変換（日付シリアル値も自動変換される）
    Dim s As String
    s = CStr(val)
    ' Excelの数値日付シリアルを日付文字列に変換（セル値がDate型に見える場合）
    If IsDate(s) And Not IsNumeric(val) Then
        SafeStr = Format(CDate(s), GetConfig(CFG_DATE_FORMAT))
    Else
        SafeStr = s
    End If
    Exit Function
ErrHandler:
    SafeStr = ""
End Function

'-------------------------------------------------------------
' GetConfig: 設定シートからキーに対応する値を取得する
'-------------------------------------------------------------
Public Function GetConfig(key As String) As String
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_SETTINGS)
    Dim i As Long
    For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If Trim(ws.Cells(i, 1).Value) = key Then
            GetConfig = CStr(ws.Cells(i, 2).Value)
            Exit Function
        End If
    Next i
    ' キーが見つからない場合のデフォルト値
    Select Case key
        Case CFG_DATE_FORMAT:   GetConfig = "yyyy/mm/dd"
        Case CFG_MAX_RESULTS:   GetConfig = "100"
        Case Else:              GetConfig = ""
    End Select
    Exit Function
ErrHandler:
    ' 設定シートが存在しない場合のデフォルト値
    Select Case key
        Case CFG_DATE_FORMAT:   GetConfig = "yyyy/mm/dd"
        Case CFG_MAX_RESULTS:   GetConfig = "100"
        Case Else:              GetConfig = ""
    End Select
End Function

'-------------------------------------------------------------
' SetStatusMessage: シートの指定セルにステータスメッセージを表示する
' isError=True → 赤背景、isError=False → 緑背景
'-------------------------------------------------------------
Public Sub SetStatusMessage(ws As Worksheet, cellAddr As String, _
                             msg As String, isError As Boolean)
    On Error Resume Next
    Dim cell As Range
    Set cell = ws.Range(cellAddr)
    cell.Value = msg
    If isError Then
        cell.Interior.Color = RGB(255, 200, 200)    ' 薄い赤
        cell.Font.Color = RGB(180, 0, 0)             ' 赤文字
    Else
        cell.Interior.Color = RGB(200, 255, 200)    ' 薄い緑
        cell.Font.Color = RGB(0, 130, 0)             ' 緑文字
    End If
End Sub

'-------------------------------------------------------------
' ClearStatusMessage: ステータスメッセージをクリアする
'-------------------------------------------------------------
Public Sub ClearStatusMessage(ws As Worksheet, cellAddr As String)
    On Error Resume Next
    Dim cell As Range
    Set cell = ws.Range(cellAddr)
    cell.Value = ""
    cell.Interior.ColorIndex = xlNone
    cell.Font.ColorIndex = xlAutomatic
End Sub

'-------------------------------------------------------------
' LogError: エラーログシートにエラー情報を記録する
'-------------------------------------------------------------
Public Sub LogError(context As String, errNum As Long, errMsg As String)
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_ERROR_LOG)
    If ws Is Nothing Then Exit Sub

    ' 最終行の次に追記
    Dim nextRow As Long
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If nextRow < 2 Then nextRow = 2

    ws.Cells(nextRow, 1).Value = Now()
    ws.Cells(nextRow, 1).NumberFormat = "yyyy/mm/dd hh:mm:ss"
    ws.Cells(nextRow, 2).Value = context
    ws.Cells(nextRow, 3).Value = errNum
    ws.Cells(nextRow, 4).Value = errMsg
End Sub

'-------------------------------------------------------------
' SheetExists: 指定名のシートが存在するか確認する
'-------------------------------------------------------------
Public Function SheetExists(sheetName As String) As Boolean
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)
    SheetExists = Not (ws Is Nothing)
    On Error GoTo 0
End Function

'-------------------------------------------------------------
' GetOrCreateSheet: シートを取得または新規作成する
'-------------------------------------------------------------
Public Function GetOrCreateSheet(sheetName As String) As Worksheet
    If SheetExists(sheetName) Then
        Set GetOrCreateSheet = ThisWorkbook.Sheets(sheetName)
    Else
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = sheetName
        Set GetOrCreateSheet = ws
    End If
End Function

'-------------------------------------------------------------
' GetNextTemplateID: 次のテンプレートIDを取得してインクリメントする
'-------------------------------------------------------------
Public Function GetNextTemplateID() As Long
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_INTERNAL)
    Dim currentID As Long
    currentID = CLng(ws.Range("B1").Value)
    ws.Range("B1").Value = currentID + 1
    GetNextTemplateID = currentID + 1
    Exit Function
ErrHandler:
    GetNextTemplateID = 1
End Function

'-------------------------------------------------------------
' AddButtonToCell: 指定セル範囲にフォームコントロールボタンを追加する
'-------------------------------------------------------------
Public Function AddButtonToCell(ws As Worksheet, cellAddr As String, _
                                 caption As String, macroName As String) As Object
    On Error GoTo ErrHandler
    Dim rng As Range
    Set rng = ws.Range(cellAddr)

    ' 既存ボタンの削除（同じセル位置のもの）
    Dim btn As Object
    For Each btn In ws.Buttons
        If btn.TopLeftCell.Address = rng.Address Then
            btn.Delete
            Exit For
        End If
    Next btn

    ' 新しいボタンを追加
    Dim newBtn As Object
    Set newBtn = ws.Buttons.Add(rng.Left + 2, rng.Top + 2, rng.Width - 4, rng.Height - 4)
    newBtn.Caption = caption
    newBtn.OnAction = macroName
    newBtn.Font.Size = 10
    Set AddButtonToCell = newBtn
    Exit Function
ErrHandler:
    Set AddButtonToCell = Nothing
    LogError "AddButtonToCell", Err.Number, Err.Description
End Function

'-------------------------------------------------------------
' FormatHeaderRow: ヘッダー行を太字・背景色付きで書式設定する
'-------------------------------------------------------------
Public Sub FormatHeaderRow(ws As Worksheet, rowNum As Long, _
                            lastCol As Long, Optional bgColor As Long = -1)
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(rowNum, 1), ws.Cells(rowNum, lastCol))
    rng.Font.Bold = True
    rng.Font.Size = 10
    If bgColor = -1 Then
        rng.Interior.Color = RGB(68, 114, 196)  ' 濃い青
        rng.Font.Color = RGB(255, 255, 255)      ' 白文字
    Else
        rng.Interior.Color = bgColor
    End If
    rng.HorizontalAlignment = xlCenter
    rng.VerticalAlignment = xlCenter
    rng.RowHeight = 22
End Sub
