Attribute VB_Name = "Module_Outlook"
Option Explicit

'=============================================================
' Module_Outlook: Outlook COM連携モジュール
' レイトバインディングで Outlook を操作し、メール作成画面を開く
' .Display を使用するため、ユーザーが確認してから手動送信できる
'=============================================================

Private Const olMailItem As Long = 0

'-------------------------------------------------------------
' CreateEmail: Outlook のメール作成画面を開く
'-------------------------------------------------------------
Public Sub CreateEmail(toAddr As String, ccAddr As String, _
                       subject As String, body As String, _
                       isHTML As Boolean)
    On Error GoTo ErrHandler

    Dim olApp As Object
    On Error Resume Next
    Set olApp = GetObject(, "Outlook.Application")
    On Error GoTo ErrHandler

    If olApp Is Nothing Then
        Set olApp = CreateObject("Outlook.Application")
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
