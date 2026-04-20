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
