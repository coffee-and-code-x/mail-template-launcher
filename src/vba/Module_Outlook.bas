Attribute VB_Name = "Module_Outlook"
Option Explicit

'=============================================================
' Module_Outlook: Outlook COM連携モジュール
' レイトバインディングで Outlook を操作し、メール作成画面を開く
' B5 に Outlook の実行ファイルパスを設定すると、そのバージョンを優先起動する
' 例: C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE
'=============================================================

Private Const olMailItem      As Long = 0
Private Const SHEET_TEMPLATES As String = "テンプレート一覧"
Private Const ROW_OUTLOOK_PATH As Long = 5  ' B5: Outlookパス（任意）

'-------------------------------------------------------------
' GetOutlookExePath: シートから Outlook 実行ファイルパスを取得する
'-------------------------------------------------------------
Private Function GetOutlookExePath() As String
    On Error Resume Next
    Dim path As String
    path = Trim(CStr(ThisWorkbook.Sheets(SHEET_TEMPLATES).Cells(ROW_OUTLOOK_PATH, 2).Value))
    On Error GoTo 0
    If path = "" Then Exit Function

    If Dir(path) = "" Then
        MsgBox "Outlookパスに指定されたファイルが見つかりません。" & vbCrLf & path & vbCrLf & vbCrLf & _
               "B5 のパスを確認するか、空欄にしてください。", vbExclamation, "Outlookパス設定エラー"
        GetOutlookExePath = ""
        Exit Function
    End If
    GetOutlookExePath = path
End Function

'-------------------------------------------------------------
' LaunchSpecificOutlook: 指定パスの Outlook を Shell で起動し COM を返す
'-------------------------------------------------------------
Private Function LaunchSpecificOutlook(exePath As String) As Object
    Dim olApp As Object

    ' すでに起動中ならそのまま使う
    On Error Resume Next
    Set olApp = GetObject(, "Outlook.Application")
    On Error GoTo 0
    If Not olApp Is Nothing Then Set LaunchSpecificOutlook = olApp : Exit Function

    ' 指定パスで新規起動
    Shell Chr(34) & exePath & Chr(34), vbNormalFocus

    ' COM が取得できるまで最大10秒ポーリング
    Dim startTime As Single
    startTime = Timer
    Do
        On Error Resume Next
        Set olApp = GetObject(, "Outlook.Application")
        On Error GoTo 0
        If Not olApp Is Nothing Then Exit Do
        Application.Wait Now + TimeValue("00:00:01")
    Loop While Timer - startTime < 10

    Set LaunchSpecificOutlook = olApp
End Function

'-------------------------------------------------------------
' CreateEmail: Outlook のメール作成画面を開く
'-------------------------------------------------------------
Public Sub CreateEmail(toAddr As String, ccAddr As String, _
                       subject As String, body As String, _
                       isHTML As Boolean)
    On Error GoTo ErrHandler

    Dim olApp As Object
    Dim exePath As String
    exePath = GetOutlookExePath()

    If exePath <> "" Then
        Set olApp = LaunchSpecificOutlook(exePath)
        If olApp Is Nothing Then
            MsgBox "指定した Outlook を起動できませんでした。" & vbCrLf & "パス: " & exePath, _
                   vbCritical, "Outlook 起動エラー"
            Exit Sub
        End If
    Else
        On Error Resume Next
        Set olApp = GetObject(, "Outlook.Application")
        On Error GoTo ErrHandler
        If olApp Is Nothing Then Set olApp = CreateObject("Outlook.Application")
    End If

    If olApp Is Nothing Then
        MsgBox "Outlook を起動できませんでした。" & vbCrLf & _
               "Outlook がインストールされているか確認してください。", _
               vbCritical, "Outlook 起動エラー"
        Exit Sub
    End If

    Dim olMail As Object
    Set olMail = olApp.CreateItem(olMailItem)

    With olMail
        If Trim(toAddr) <> "" Then .To = toAddr
        If Trim(ccAddr) <> "" Then .CC = ccAddr
        .Subject = subject
        If isHTML Then
            .HTMLBody = body
        Else
            .Body = body
        End If
        .Display
    End With

    Set olMail = Nothing
    Set olApp = Nothing
    Exit Sub

ErrHandler:
    On Error Resume Next
    If Not olMail Is Nothing Then Set olMail = Nothing
    If Not olApp Is Nothing Then Set olApp = Nothing
    MsgBox "Outlook でメールを作成できませんでした。" & vbCrLf & _
           "エラー " & Err.Number & ": " & Err.Description, _
           vbCritical, "Outlook エラー"
End Sub
