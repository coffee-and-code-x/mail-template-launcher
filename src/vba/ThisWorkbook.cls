VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'=============================================================
' ThisWorkbook: ワークブックイベントモジュール
' ブックを開いた際に初期化チェックと名前付き範囲の設定を行う
'=============================================================

'-------------------------------------------------------------
' Workbook_Open: ブックを開いた時に実行される
'-------------------------------------------------------------
Private Sub Workbook_Open()
    On Error GoTo OpenError

    ' Outlook の利用可能チェック（警告のみ、処理は止めない）
    If Not IsOutlookAvailable() Then
        ' 起動時には警告しない（ユーザーの邪魔をしない）
        ' MsgBox は出さずエラーログにのみ記録
        LogError "Workbook_Open", 0, "Outlook が起動していないか利用できません。"
    End If

    ' 初期化チェック（未初期化なら初期化を実行）
    Module_Init.InitializeWorkbook

    Exit Sub

OpenError:
    ' 開く際のエラーは無視（ユーザーがキャンセルした場合など）
    On Error Resume Next
    LogError "Workbook_Open", Err.Number, Err.Description
End Sub

'-------------------------------------------------------------
' Workbook_BeforeClose: ブックを閉じる前に実行される
'-------------------------------------------------------------
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' ステータスバーをリセット
    Application.StatusBar = False
End Sub
