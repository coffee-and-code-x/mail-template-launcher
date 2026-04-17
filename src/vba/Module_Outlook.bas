Attribute VB_Name = "Module_Outlook"
Option Explicit

'=============================================================
' Module_Outlook: Outlook COM連携モジュール
' レイトバインディングで Outlook を操作し、メール作成画面を開く
' 注意: .Display を使用するため、ユーザーが内容を確認してから送信できる
'
' 【複数バージョンの Outlook が共存している場合】
' 設定シートの「Outlookパス」に Office 365 版の実行ファイルパスを設定すると
' そのバージョンを優先して起動します。
' 例: C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE
'=============================================================

' Outlook MailItem の定数（レイトバインディングのため数値で定義）
Private Const olMailItem    As Long = 0
Private Const olFormatHTML  As Long = 2
Private Const olFormatPlain As Long = 1

' 設定キー
Private Const CFG_OUTLOOK_PATH  As String = "Outlookパス"
Private Const CFG_OUTLOOK_WAIT  As String = "Outlook起動待機秒数"

'-------------------------------------------------------------
' GetOutlookExePath: 設定シートから Outlook 実行ファイルパスを取得する
' 未設定または空の場合は "" を返す
'-------------------------------------------------------------
Private Function GetOutlookExePath() As String
    Dim path As String
    path = Trim(GetConfig(CFG_OUTLOOK_PATH))
    ' パスが設定されているが存在しない場合は警告
    If path <> "" And Dir(path) = "" Then
        MsgBox "設定シートの「Outlookパス」に指定されたファイルが見つかりません。" & vbCrLf & _
               path & vbCrLf & vbCrLf & _
               "設定シートのパスを確認するか、空欄にしてください。", _
               vbExclamation, "Outlookパス設定エラー"
        GetOutlookExePath = ""
        Exit Function
    End If
    GetOutlookExePath = path
End Function

'-------------------------------------------------------------
' LaunchSpecificOutlook: 指定パスの Outlook を Shell で起動し、
'                        COM オブジェクトが取得できるまで待機する
'-------------------------------------------------------------
Private Function LaunchSpecificOutlook(exePath As String) As Object
    Dim waitSec As Long
    waitSec = 5  ' デフォルト待機秒数
    Dim cfgWait As String
    cfgWait = Trim(GetConfig(CFG_OUTLOOK_WAIT))
    If IsNumeric(cfgWait) And CLng(cfgWait) > 0 Then
        waitSec = CLng(cfgWait)
    End If

    ' すでに起動中か確認（パスが異なっても COM は共有）
    Dim olApp As Object
    On Error Resume Next
    Set olApp = GetObject(, "Outlook.Application")
    On Error GoTo 0
    If Not olApp Is Nothing Then
        Set LaunchSpecificOutlook = olApp
        Exit Function
    End If

    ' 指定パスで Outlook を起動
    Shell Chr(34) & exePath & Chr(34), vbNormalFocus

    ' COM オブジェクトが取得できるまでポーリング（最大 waitSec 秒）
    Dim startTime As Single
    startTime = Timer
    Do
        On Error Resume Next
        Set olApp = GetObject(, "Outlook.Application")
        On Error GoTo 0
        If Not olApp Is Nothing Then Exit Do
        Application.Wait Now + TimeValue("00:00:01")
    Loop While Timer - startTime < waitSec

    Set LaunchSpecificOutlook = olApp  ' 取得できなければ Nothing
End Function

'-------------------------------------------------------------
' IsOutlookAvailable: Outlook が利用可能か確認する
'-------------------------------------------------------------
Public Function IsOutlookAvailable() As Boolean
    On Error Resume Next
    Dim olApp As Object
    Dim exePath As String
    exePath = GetOutlookExePath()

    If exePath <> "" Then
        ' 指定パスの Outlook を起動して確認
        Set olApp = LaunchSpecificOutlook(exePath)
    Else
        Set olApp = GetObject(, "Outlook.Application")
        If olApp Is Nothing Then
            Set olApp = CreateObject("Outlook.Application")
        End If
    End If

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
    Dim exePath As String
    exePath = GetOutlookExePath()

    If exePath <> "" Then
        ' 設定に Outlook パスが指定されている → そのバージョンを優先起動
        Set olApp = LaunchSpecificOutlook(exePath)
        If olApp Is Nothing Then
            MsgBox "指定された Outlook を起動できませんでした。" & vbCrLf & _
                   "パス: " & exePath & vbCrLf & vbCrLf & _
                   "設定シートの「Outlookパス」を確認してください。", _
                   vbCritical, "Outlook 起動エラー"
            Exit Sub
        End If
    Else
        ' 設定なし → 起動中の Outlook に接続、なければ COM で起動
        On Error Resume Next
        Set olApp = GetObject(, "Outlook.Application")
        On Error GoTo OutlookError
        If olApp Is Nothing Then
            Set olApp = CreateObject("Outlook.Application")
        End If
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
        If Trim(toAddr) <> "" Then .To = toAddr
        If Trim(ccAddr) <> "" Then .CC = ccAddr
        .Subject = subject
        ' 重要: .HTMLBody と .Body は相互上書きするため、一方のみ設定する
        If isHTML Then
            .HTMLBody = body
        Else
            .Body = body
        End If
        ' .Send ではなく .Display → ユーザーが確認してから手動送信
        .Display
    End With

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
           "  ・Office 365 版を使う場合は設定シートの「Outlookパス」を設定してください" & vbCrLf & vbCrLf & _
           "エラー情報: " & errNum & " - " & errDesc, _
           vbCritical, "Outlook エラー"
End Sub

'-------------------------------------------------------------
' CreateEmailFromTemplate: テンプレートIDと案件データからメールを作成する
' Module_Launcher.LaunchTemplate から呼び出される統合関数
'-------------------------------------------------------------
Public Sub CreateEmailFromTemplate(templateID As Long, project As ProjectData)
    If FindTemplateRow(templateID) = 0 Then
        MsgBox "テンプレートID " & templateID & " が見つかりません。", _
               vbExclamation, "テンプレートエラー"
        Exit Sub
    End If

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

    TouchTemplateUpdated templateID
    CreateEmail toAddr, ccAddr, subject, body, isHTML
End Sub
