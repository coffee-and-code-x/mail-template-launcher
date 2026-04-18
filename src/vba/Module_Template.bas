Attribute VB_Name = "Module_Template"
Option Explicit

'=============================================================
' Module_Template: テンプレート処理モジュール
' テンプレートデータの読み取りと {プレースホルダー} 置換を行う
'=============================================================

' テンプレート一覧シートの列番号
Private Const COL_ID            As Long = 1
Private Const COL_NAME          As Long = 2
Private Const COL_FORMAT        As Long = 3
Private Const COL_TO            As Long = 4
Private Const COL_CC            As Long = 5
Private Const COL_SUBJECT       As Long = 6
Private Const COL_BODY_SHEET    As Long = 7
Private Const COL_UPDATED       As Long = 8

' テンプレート一覧のデータ開始行
Private Const TEMPLATE_DATA_ROW As Long = 4

'-------------------------------------------------------------
' FindTemplateRow: テンプレートIDに対応する行番号を返す
' 見つからない場合は 0 を返す
'-------------------------------------------------------------
Public Function FindTemplateRow(templateID As Long) As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_TEMPLATE_LIST)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_ID).End(xlUp).Row

    Dim i As Long
    For i = TEMPLATE_DATA_ROW To lastRow
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
' GetTemplateFormat: テンプレートの形式（"HTML" or "TEXT"）を返す
'-------------------------------------------------------------
Public Function GetTemplateFormat(templateID As Long) As String
    Dim row As Long
    row = FindTemplateRow(templateID)
    If row = 0 Then
        GetTemplateFormat = "TEXT"
        Exit Function
    End If
    Dim fmt As String
    fmt = UCase(Trim(CStr(ThisWorkbook.Sheets(SHEET_TEMPLATE_LIST).Cells(row, COL_FORMAT).Value)))
    If fmt = "HTML" Then
        GetTemplateFormat = "HTML"
    Else
        GetTemplateFormat = "TEXT"
    End If
End Function

'-------------------------------------------------------------
' GetBodySheetName: テンプレートに対応する本文シート名を返す
'-------------------------------------------------------------
Public Function GetBodySheetName(templateID As Long) As String
    Dim row As Long
    row = FindTemplateRow(templateID)
    If row = 0 Then
        GetBodySheetName = ""
        Exit Function
    End If
    GetBodySheetName = CStr(ThisWorkbook.Sheets(SHEET_TEMPLATE_LIST).Cells(row, COL_BODY_SHEET).Value)
End Function

'-------------------------------------------------------------
' BuildEmailBody: テンプレート本文を読み取り、プレースホルダーを置換して返す
'-------------------------------------------------------------
Public Function BuildEmailBody(templateID As Long, project As CProjectData) As String
    Dim bodySheetName As String
    bodySheetName = GetBodySheetName(templateID)

    If bodySheetName = "" Or Not SheetExists(bodySheetName) Then
        BuildEmailBody = "（本文シートが見つかりません: " & bodySheetName & "）"
        Exit Function
    End If

    ' 本文シートの A4 セルから本文を読み取る
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(bodySheetName)
    Dim rawBody As String
    rawBody = CStr(ws.Range("A4").Value)

    ' プレースホルダー置換
    Dim result As String
    result = SubstitutePlaceholders(rawBody, project)

    ' TEXT 形式の場合は改行コードを正規化
    If GetTemplateFormat(templateID) = "TEXT" Then
        result = Replace(result, Chr(10), vbCrLf)
        result = Replace(result, Chr(13) & Chr(13), Chr(13))
    End If

    BuildEmailBody = result
End Function

'-------------------------------------------------------------
' BuildSubjectLine: テンプレート件名にプレースホルダー置換を適用して返す
'-------------------------------------------------------------
Public Function BuildSubjectLine(templateID As Long, project As CProjectData) As String
    Dim row As Long
    row = FindTemplateRow(templateID)
    If row = 0 Then
        BuildSubjectLine = ""
        Exit Function
    End If

    Dim rawSubject As String
    rawSubject = CStr(ThisWorkbook.Sheets(SHEET_TEMPLATE_LIST).Cells(row, COL_SUBJECT).Value)
    BuildSubjectLine = SubstitutePlaceholders(rawSubject, project)
End Function

'-------------------------------------------------------------
' BuildToAddress: 宛先 (To) にプレースホルダー置換を適用して返す
'-------------------------------------------------------------
Public Function BuildToAddress(templateID As Long, project As CProjectData) As String
    Dim row As Long
    row = FindTemplateRow(templateID)
    If row = 0 Then
        BuildToAddress = ""
        Exit Function
    End If

    Dim rawTo As String
    rawTo = CStr(ThisWorkbook.Sheets(SHEET_TEMPLATE_LIST).Cells(row, COL_TO).Value)
    BuildToAddress = SubstitutePlaceholders(rawTo, project)
End Function

'-------------------------------------------------------------
' BuildCCAddress: CC にプレースホルダー置換を適用して返す
'-------------------------------------------------------------
Public Function BuildCCAddress(templateID As Long, project As CProjectData) As String
    Dim row As Long
    row = FindTemplateRow(templateID)
    If row = 0 Then
        BuildCCAddress = ""
        Exit Function
    End If

    Dim rawCC As String
    rawCC = CStr(ThisWorkbook.Sheets(SHEET_TEMPLATE_LIST).Cells(row, COL_CC).Value)
    BuildCCAddress = SubstitutePlaceholders(rawCC, project)
End Function

'-------------------------------------------------------------
' SubstitutePlaceholders: テキスト内のプレースホルダーを案件データで置換する
' 全てのフィールド（宛先・件名・本文）で使用可能
'-------------------------------------------------------------
Public Function SubstitutePlaceholders(text As String, project As CProjectData) As String
    Dim result As String
    result = text

    ' 日付フォーマット取得
    Dim dateFormat As String
    dateFormat = GetConfig(CFG_DATE_FORMAT)
    If dateFormat = "" Then dateFormat = "yyyy/mm/dd"

    ' 標準プレースホルダーを置換
    result = Replace(result, "{案件名}",     project.案件名)
    result = Replace(result, "{案件番号}",   project.案件番号)
    result = Replace(result, "{顧客名}",     project.顧客名)
    result = Replace(result, "{担当者名}",   project.担当者名)
    result = Replace(result, "{期日}",       project.期日)
    result = Replace(result, "{今日の日付}", Format(Now(), dateFormat))

    ' 未解決プレースホルダーの検出と警告
    CheckUnresolvedPlaceholders result

    SubstitutePlaceholders = result
End Function

'-------------------------------------------------------------
' CheckUnresolvedPlaceholders: 未解決の {プレースホルダー} を検出して警告する
'-------------------------------------------------------------
Private Sub CheckUnresolvedPlaceholders(text As String)
    Dim pos As Long
    pos = InStr(1, text, "{")
    If pos = 0 Then Exit Sub

    Dim endPos As Long
    endPos = InStr(pos, text, "}")
    If endPos = 0 Then Exit Sub

    ' 未解決プレースホルダーを収集
    Dim unresolved As String
    unresolved = ""
    Dim searchPos As Long
    searchPos = 1
    Do
        pos = InStr(searchPos, text, "{")
        If pos = 0 Then Exit Do
        endPos = InStr(pos, text, "}")
        If endPos = 0 Then Exit Do

        Dim placeholder As String
        placeholder = Mid(text, pos, endPos - pos + 1)
        unresolved = unresolved & placeholder & "  "
        searchPos = endPos + 1
    Loop

    If Trim(unresolved) <> "" Then
        MsgBox "以下のプレースホルダーが置換されませんでした:" & vbCrLf & vbCrLf & _
               unresolved & vbCrLf & vbCrLf & _
               "メール内容を確認してから送信してください。", _
               vbExclamation, "未解決プレースホルダー"
    End If
End Sub

'-------------------------------------------------------------
' TouchTemplateUpdated: テンプレートの最終更新日時を現在時刻に更新する
'-------------------------------------------------------------
Public Sub TouchTemplateUpdated(templateID As Long)
    Dim row As Long
    row = FindTemplateRow(templateID)
    If row > 0 Then
        With ThisWorkbook.Sheets(SHEET_TEMPLATE_LIST).Cells(row, COL_UPDATED)
            .Value = Now()
            .NumberFormat = "yyyy/mm/dd hh:mm"
        End With
    End If
End Sub

'-------------------------------------------------------------
' GetTemplateName: テンプレート名を返す
'-------------------------------------------------------------
Public Function GetTemplateName(templateID As Long) As String
    Dim row As Long
    row = FindTemplateRow(templateID)
    If row = 0 Then
        GetTemplateName = "（不明）"
        Exit Function
    End If
    GetTemplateName = CStr(ThisWorkbook.Sheets(SHEET_TEMPLATE_LIST).Cells(row, COL_NAME).Value)
End Function
