Attribute VB_Name = "Module_Launcher"
Option Explicit

'=============================================================
' Module_Launcher: 「起動」ボタンのエントリポイント
' テンプレート一覧シートの各起動ボタンから呼ばれる
'=============================================================

Public Sub Launch_1() : LaunchTemplate 1 : End Sub
Public Sub Launch_2() : LaunchTemplate 2 : End Sub
Public Sub Launch_3() : LaunchTemplate 3 : End Sub
Public Sub Launch_4() : LaunchTemplate 4 : End Sub
Public Sub Launch_5() : LaunchTemplate 5 : End Sub

'-------------------------------------------------------------
' InsertPlaceholder: 本文シートのボタンから呼ばれる
' テンプレート一覧の現在のラベル名を読み、{ラベル名} をクリップボードにコピーする
' ボタン名は "PlaceholderBtn_N"（N=1〜3）で、末尾の番号でフィールドを特定する
'-------------------------------------------------------------
Public Sub InsertPlaceholder()
    On Error GoTo ErrHandler

    Dim callerName As String
    callerName = CStr(Application.Caller)

    Dim parts() As String
    parts = Split(callerName, "_")
    Dim fieldNum As Long
    fieldNum = CLng(parts(UBound(parts)))

    Dim label As String
    label = GetFieldLabel(fieldNum)

    If label = "" Then
        MsgBox "フィールド " & fieldNum & " のラベルが設定されていません。" & vbCrLf & _
               "テンプレート一覧の A" & (1 + fieldNum) & " にラベルを入力してください。", _
               vbExclamation, "ラベル未設定"
        Exit Sub
    End If

    Dim placeholder As String
    placeholder = "{" & label & "}"

    ' ボタンのキャプションを現在のラベルに同期
    ActiveSheet.Buttons(callerName).Caption = placeholder

    ' クリップボードにコピー
    CreateObject("htmlfile").ParentWindow.ClipboardData.SetData "text", placeholder

    Application.StatusBar = placeholder & "  をコピーしました - A2 セルで Ctrl+V で貼り付けてください"
    Exit Sub

ErrHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbExclamation, "エラー"
End Sub

'-------------------------------------------------------------
' LaunchTemplate: テンプレートIDを指定してメール作成画面を開く
'-------------------------------------------------------------
Private Sub LaunchTemplate(templateID As Long)
    If Not TemplateExists(templateID) Then
        MsgBox "テンプレートID " & templateID & " が見つかりません。", _
               vbExclamation, "テンプレートエラー"
        Exit Sub
    End If

    CreateEmail _
        GetToAddress(templateID), _
        GetCCAddress(templateID), _
        GetSubject(templateID), _
        GetBody(templateID), _
        IsHTMLFormat(templateID)
End Sub
