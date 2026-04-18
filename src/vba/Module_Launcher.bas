Attribute VB_Name = "Module_Launcher"
Option Explicit

'=============================================================
' Module_Launcher: メール起動トップレベルモジュール
' 「起動」ボタンから呼ばれるエントリポイントを提供する
' テンプレート選択・案件取得・Outlook起動を統括する
'=============================================================

'-------------------------------------------------------------
' LaunchTemplate: テンプレートIDを指定してメールを起動する
' Module_ButtonHandlers の Launch_N から呼び出される
'-------------------------------------------------------------
Public Sub LaunchTemplate(templateID As Long)
    On Error GoTo LaunchError

    ' テンプレートの存在確認
    If FindTemplateRow(templateID) = 0 Then
        MsgBox "テンプレートID " & templateID & " が見つかりません。" & vbCrLf & _
               "テンプレート一覧を確認してください。", vbExclamation, "テンプレートエラー"
        Exit Sub
    End If

    ' 選択中の案件を取得
    Dim project As CProjectData
    Set project = GetSelectedProject()

    ' 案件未選択の場合の確認
    If project.IsEmpty() Then
        Dim ans As Integer
        ans = MsgBox("案件が選択されていません。" & vbCrLf & vbCrLf & _
                     "案件情報なし（プレースホルダーは空のまま）でメールを作成しますか？" & vbCrLf & vbCrLf & _
                     "[はい] → そのままメールを作成" & vbCrLf & _
                     "[いいえ] → 案件検索シートへ移動", _
                     vbYesNo + vbQuestion, "案件未選択")
        If ans = vbNo Then
            NavigateToSearch
            Exit Sub
        End If
        ' 空の ProjectData でそのまま続行
    End If

    ' テンプレート名を取得して確認メッセージ用に使用
    Dim templateName As String
    templateName = GetTemplateName(templateID)

    ' Outlook でメールを作成
    CreateEmailFromTemplate templateID, project

    Exit Sub

LaunchError:
    LogError "LaunchTemplate(" & templateID & ")", Err.Number, Err.Description
    MsgBox "メール起動中にエラーが発生しました。" & vbCrLf & _
           "エラー " & Err.Number & ": " & Err.Description, vbCritical, "起動エラー"
End Sub

'-------------------------------------------------------------
' AddNewTemplate: 新規テンプレートをテンプレート一覧に追加する
' 「新規テンプレート追加」ボタンから呼び出される
'-------------------------------------------------------------
Public Sub AddNewTemplate()
    ' 新しいテンプレートIDを取得
    Dim newID As Long
    newID = GetNextTemplateID()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_TEMPLATE_LIST)

    ' 最終データ行の次に新規行を追加
    Dim newRow As Long
    newRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If newRow < 4 Then newRow = 4

    ' テンプレート名の入力
    Dim newName As String
    newName = InputBox("新しいテンプレートの名前を入力してください:", "テンプレート名", "新しいテンプレート " & newID)
    If Trim(newName) = "" Then
        ' ID カウンターを元に戻す
        ThisWorkbook.Sheets(SHEET_INTERNAL).Range("B2").Value = newID - 1
        Exit Sub
    End If

    ' メール形式の選択
    Dim fmtAns As Integer
    fmtAns = MsgBox("メール本文の形式を選択してください。" & vbCrLf & vbCrLf & _
                    "[はい] → HTML形式（太字・色・表などの書式が使えます）" & vbCrLf & _
                    "[いいえ] → テキスト形式（シンプルなプレーンテキスト）", _
                    vbYesNo + vbQuestion, "メール形式の選択")
    Dim newFormat As String
    Dim isHTML As Boolean
    If fmtAns = vbYes Then
        newFormat = "HTML"
        isHTML = True
    Else
        newFormat = "TEXT"
        isHTML = False
    End If

    ' テンプレート一覧に行を追加
    ws.Cells(newRow, 1).Value = newID
    ws.Cells(newRow, 2).Value = newName
    ws.Cells(newRow, 3).Value = newFormat
    ws.Cells(newRow, 4).Value = ""           ' 宛先（後で入力）
    ws.Cells(newRow, 5).Value = ""           ' CC（後で入力）
    ws.Cells(newRow, 6).Value = newName      ' 件名の初期値をテンプレート名に設定
    ws.Cells(newRow, 7).Value = "本文_" & newID
    ws.Cells(newRow, 8).Value = Now()
    ws.Cells(newRow, 8).NumberFormat = "yyyy/mm/dd hh:mm"
    ws.Rows(newRow).RowHeight = 25

    ' 起動ボタンの追加（ID が 30 以下の場合）
    If newID <= 30 Then
        AddButtonToCell ws, "I" & newRow, "起動", "Launch_" & newID
    Else
        ws.Cells(newRow, 9).Value = "（ID=" & newID & "）起動にはModule_ButtonHandlersに追加が必要"
        ws.Cells(newRow, 9).Font.Color = RGB(180, 0, 0)
        ws.Cells(newRow, 9).Font.Size = 8
    End If

    ' 本文シートを作成
    CreateBodySheet newID, newName, isHTML

    ' 案内メッセージと本文シートへの移動
    MsgBox "テンプレート「" & newName & "」（ID=" & newID & "）を追加しました。" & vbCrLf & vbCrLf & _
           "本文シート「本文_" & newID & "」に移動します。" & vbCrLf & _
           "・A4 セルに本文を入力してください" & vbCrLf & _
           "・{プレースホルダー} を使って案件情報を差し込めます" & vbCrLf & _
           "・宛先・CC・件名はテンプレート一覧シートで編集してください", _
           vbInformation, "テンプレート追加完了"

    ' 本文シートへ移動して編集を促す
    If SheetExists("本文_" & newID) Then
        ThisWorkbook.Sheets("本文_" & newID).Activate
        ThisWorkbook.Sheets("本文_" & newID).Range("A4").Select
    End If
End Sub

'-------------------------------------------------------------
' OpenBodySheet: テンプレートIDの本文シートを開く（ボタン用）
'-------------------------------------------------------------
Public Sub OpenBodySheet(templateID As Long)
    Dim bodySheetName As String
    bodySheetName = GetBodySheetName(templateID)

    If bodySheetName = "" Then
        MsgBox "テンプレートID " & templateID & " の本文シート名が設定されていません。", _
               vbExclamation, "本文シートを開く"
        Exit Sub
    End If

    If Not SheetExists(bodySheetName) Then
        Dim ans As Integer
        ans = MsgBox("本文シート「" & bodySheetName & "」が見つかりません。" & vbCrLf & _
                     "新しく作成しますか？", vbYesNo + vbQuestion, "本文シートを開く")
        If ans = vbYes Then
            CreateBodySheet templateID, GetTemplateName(templateID), _
                            (GetTemplateFormat(templateID) = "HTML")
        End If
        Exit Sub
    End If

    ThisWorkbook.Sheets(bodySheetName).Activate
    ThisWorkbook.Sheets(bodySheetName).Range("A4").Select
End Sub
