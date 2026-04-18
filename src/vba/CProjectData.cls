VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CProjectData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'=============================================================
' CProjectData: 案件データクラス
' 外部Excelファイルから取得した案件1件分の情報を保持する
'=============================================================

Public 案件名      As String
Public 案件番号    As String
Public 顧客名      As String
Public 担当者名    As String
Public 期日        As String
Public SourceFile  As String   ' 取得元ファイル表示名
Public SourceRow   As Long     ' 取得元行番号

'-------------------------------------------------------------
' IsEmpty: 主要フィールドが全て空かどうか確認する
'-------------------------------------------------------------
Public Function IsEmpty() As Boolean
    IsEmpty = (Len(Trim(案件名)) = 0 And Len(Trim(顧客名)) = 0)
End Function

'-------------------------------------------------------------
' ToString: デバッグ用の文字列表現を返す
'-------------------------------------------------------------
Public Function ToString() As String
    ToString = "[" & 案件番号 & "] " & 案件名 & " / " & 顧客名 & " (" & SourceFile & ")"
End Function
