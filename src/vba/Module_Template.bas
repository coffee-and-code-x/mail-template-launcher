Attribute VB_Name = "Module_Template"
Option Explicit

'=============================================================
' Module_Template: テンプレートデータの読み取りとプレースホルダー置換
'=============================================================

' テンプレート一覧シートのレイアウト
Private Const SHEET_TEMPLATES As String = "テンプレート一覧"
Private Const DATA_START_ROW  As Long = 8   ' テンプレートデータ開始行

' テンプレート一覧の列番号
Private Const COL_ID         As Long = 1   ' A: テンプレートID
Private Const COL_FORMAT     As Long = 3   ' C: 形式 (HTML/TEXT)
Private Const COL_TO         As Long = 4   ' D: 宛先
Private Const COL_CC         As Long = 5   ' E: CC
Private Const COL_SUBJECT    As Long = 6   ' F: 件名
Private Const COL_BODY_SHEET As Long = 7   ' G: 本文シート名

' 案件情報入力セル（B列）
Private Const ROW_案件名   As Long = 2
Private Const ROW_案件番号 As Long = 3
Private Const ROW_顧客名   As Long = 4

'-------------------------------------------------------------
' FindTemplateRow: IDに対応する行番号を返す（未発見は 0）
'-------------------------------------------------------------
Private Function FindTemplateRow(templateID As Long) As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_TEMPLATES)
    Dim i As Long
    For i = DATA_START_ROW To ws.Cells(ws.Rows.Count, COL_ID).End(xlUp).Row
        If IsNumeric(ws.Cells(i, COL_ID).Value) Then
            If CLng(ws.Cells(i, COL_ID).Value) = templateID Then
                FindTemplateRow = i
                Exit Function
            End If
        End If
    Next i
    FindTemplateRow = 0
End Function

'-------------------------------------------------------------
' SubstitutePlaceholders: {案件名} {案件番号} {顧客名} を置換する
'-------------------------------------------------------------
Public Function SubstitutePlaceholders(text As String) As String
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_TEMPLATES)
    Dim result As String
    result = text
    result = Replace(result, "{案件名}",   Trim(CStr(ws.Cells(ROW_案件名,   2).Value)))
    result = Replace(result, "{案件番号}", Trim(CStr(ws.Cells(ROW_案件番号, 2).Value)))
    result = Replace(result, "{顧客名}",   Trim(CStr(ws.Cells(ROW_顧客名,   2).Value)))
    SubstitutePlaceholders = result
End Function

'-------------------------------------------------------------
' TemplateExists: テンプレートIDが存在するか確認する
'-------------------------------------------------------------
Public Function TemplateExists(templateID As Long) As Boolean
    TemplateExists = (FindTemplateRow(templateID) > 0)
End Function

'-------------------------------------------------------------
' IsHTMLFormat: HTML形式か返す
'-------------------------------------------------------------
Public Function IsHTMLFormat(templateID As Long) As Boolean
    Dim row As Long
    row = FindTemplateRow(templateID)
    If row = 0 Then IsHTMLFormat = False : Exit Function
    IsHTMLFormat = (UCase(Trim(CStr(ThisWorkbook.Sheets(SHEET_TEMPLATES).Cells(row, COL_FORMAT).Value))) = "HTML")
End Function

'-------------------------------------------------------------
' GetToAddress / GetCCAddress / GetSubject: 置換済みフィールドを返す
'-------------------------------------------------------------
Public Function GetToAddress(templateID As Long) As String
    Dim row As Long : row = FindTemplateRow(templateID)
    If row = 0 Then Exit Function
    GetToAddress = SubstitutePlaceholders(CStr(ThisWorkbook.Sheets(SHEET_TEMPLATES).Cells(row, COL_TO).Value))
End Function

Public Function GetCCAddress(templateID As Long) As String
    Dim row As Long : row = FindTemplateRow(templateID)
    If row = 0 Then Exit Function
    GetCCAddress = SubstitutePlaceholders(CStr(ThisWorkbook.Sheets(SHEET_TEMPLATES).Cells(row, COL_CC).Value))
End Function

Public Function GetSubject(templateID As Long) As String
    Dim row As Long : row = FindTemplateRow(templateID)
    If row = 0 Then Exit Function
    GetSubject = SubstitutePlaceholders(CStr(ThisWorkbook.Sheets(SHEET_TEMPLATES).Cells(row, COL_SUBJECT).Value))
End Function

'-------------------------------------------------------------
' GetBody: 本文シートの A2 から置換済み本文を返す
'-------------------------------------------------------------
Public Function GetBody(templateID As Long) As String
    Dim row As Long : row = FindTemplateRow(templateID)
    If row = 0 Then GetBody = "" : Exit Function

    Dim bodySheetName As String
    bodySheetName = Trim(CStr(ThisWorkbook.Sheets(SHEET_TEMPLATES).Cells(row, COL_BODY_SHEET).Value))
    If bodySheetName = "" Then GetBody = "" : Exit Function

    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(bodySheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        GetBody = "（本文シート「" & bodySheetName & "」が見つかりません）"
        Exit Function
    End If

    GetBody = SubstitutePlaceholders(CStr(ws.Range("A2").Value))
End Function
