Attribute VB_Name = "Module_Outlook"
Option Explicit

'=============================================================
' Module_Outlook: Outlook COM連携モジュール
' レイトバインディングで Outlook を操作し、メール作成画面を開く
' 注意: .Display を使用するため、ユーザーが内容を確認してから送信できる
'=============================================================

' Outlook MailItem の定数（レイトバインディングのため数値で定義）
Private Const olMailItem    As Long = 0
Private Const olFormatHTML  As Long = 2
Private Const olFormatPlain As Long = 1

'-------------------------------------------------------------
' IsOutlookAvailable: Outlook が利用可能か確認する
'-------------------------------------------------------------
Public Function IsOutlookAvailable() As Boolean
    On Error Resume Next
    Dim olApp As Object
    Set olApp = CreateObject("Outlook.Application")
    IsOutlookAvailable = Not (olApp Is Nothing)
    If Not olApp Is Nothing Then Set olApp = Nothing
    On Error GoTo 0
End Function

'-------------------------------------------------------------
' CreateEmail: Outlook のメール作成画面を開く
' toAddr   : 宛先（セミコロン区切り）
' ccAddr   : CC（セミコロン区切り、省略可）
' subject  : 件名
' body     : 本文
' isHTML   : True=HTML形式, False=テキスト形式
'-------------------------------------------------------------
Public Sub CreateEmail(toAddr As String, ccAddr As String, _
                        subject As String, body As String, _
                        isHTML As Boolean)
    On Error GoTo OutlookError

    Dim olApp As Object
    Dim olMail As Object

    ' まず既存の Outlook プロセスに接続を試みる（新しいプロセスを増やさない）
    On Error Resume Next
    Set olApp = GetObject(, "Outlook.Application")
    On Error GoTo OutlookError

    ' 起動していない場合は新規に起動する
    If olApp Is Nothing Then
        Set olApp = CreateObject("Outlook.Application")
        If olApp Is Nothing Then
            MsgBox "Outlook を起動できませんでした。" & vbCrLf & _
                   "Outlook がインストールされているか確認してください。", _
                   vbCritical, "Outlook 起動エラー"
            Exit Sub
        End If
    End If

    ' 新規メールアイテムを作成
    Set olMail = olApp.CreateItem(olMailItem)

    If olMail Is Nothing Then
        MsgBox "メールアイテムの作成に失敗しました。", vbCritical, "メール作成エラー"
        Exit Sub
    End If

    ' メールフィールドの設定
    With olMail
        ' 宛先とCC
        If Trim(toAddr) <> "" Then
            .To = toAddr
        End If
        If Trim(ccAddr) <> "" Then
            .CC = ccAddr
        End If

        ' 件名
        .Subject = subject

        ' 本文（HTML または テキスト）
        ' 重要: .HTMLBody と .Body は相互上書きするため、最後に設定した方が有効
        ' HTML形式の場合は .HTMLBody のみ設定
        ' テキスト形式の場合は .Body のみ設定
        If isHTML Then
            .HTMLBody = body
        Else
            .Body = body
        End If

        ' メール作成画面を表示（.Send ではなく .Display を使用）
        ' ユーザーが内容を確認してから手動で送信する
        .Display
    End With

    ' 参照をクリア（Outlook を終了させない）
    Set olMail = Nothing
    Set olApp = Nothing

    Exit Sub

OutlookError:
    On Error Resume Next
    If Not olMail Is Nothing Then Set olMail = Nothing
    If Not olApp Is Nothing Then Set olApp = Nothing

    Dim errNum As Long
    Dim errDesc As String
    errNum = Err.Number
    errDesc = Err.Description

    LogError "CreateEmail", errNum, errDesc

    MsgBox "Outlook でメールを作成できませんでした。" & vbCrLf & vbCrLf & _
           "確認事項:" & vbCrLf & _
           "  ・Outlook がインストールされているか確認してください" & vbCrLf & _
           "  ・Outlook を起動してから再度お試しください" & vbCrLf & vbCrLf & _
           "エラー情報: " & errNum & " - " & errDesc, _
           vbCritical, "Outlook エラー"
End Sub

'-------------------------------------------------------------
' CreateEmailFromTemplate: テンプレートIDと案件データからメールを作成する
' Module_Launcher.LaunchTemplate から呼び出される統合関数
'-------------------------------------------------------------
Public Sub CreateEmailFromTemplate(templateID As Long, project As ProjectData)
    ' Outlook 利用可能チェック
    If Not IsOutlookAvailable() Then
        MsgBox "Outlook が起動していないか、インストールされていません。" & vbCrLf & _
               "Outlook を起動してから再度お試しください。", _
               vbExclamation, "Outlook が見つかりません"
        Exit Sub
    End If

    ' テンプレートが存在するか確認
    If FindTemplateRow(templateID) = 0 Then
        MsgBox "テンプレートID " & templateID & " が見つかりません。", _
               vbExclamation, "テンプレートエラー"
        Exit Sub
    End If

    ' メール各フィールドを生成
    Dim toAddr As String
    Dim ccAddr As String
    Dim subject As String
    Dim body As String
    Dim isHTML As Boolean

    toAddr  = BuildToAddress(templateID, project)
    ccAddr  = BuildCCAddress(templateID, project)
    subject = BuildSubjectLine(templateID, project)
    body    = BuildEmailBody(templateID, project)
    isHTML  = (GetTemplateFormat(templateID) = "HTML")

    ' テンプレート最終更新日時を記録
    TouchTemplateUpdated templateID

    ' Outlook でメール作成画面を開く
    CreateEmail toAddr, ccAddr, subject, body, isHTML
End Sub
