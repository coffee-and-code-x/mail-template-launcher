VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CFileSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'=============================================================
' CFileSetting: 外部ファイル設定クラス
' ファイル設定シートの1行分の設定情報を保持する
'=============================================================

Public SettingID       As Long
Public DisplayName     As String
Public FilePath        As String
Public SheetName       As String
Public HeaderRow       As Long
Public Col_案件名      As Long
Public Col_案件番号    As Long
Public Col_顧客名      As Long
Public Col_担当者名    As Long
Public Col_期日        As Long
Public SearchColStr    As String   ' 検索対象列（カンマ区切り）
Public IsActive        As Boolean

'-------------------------------------------------------------
' GetSearchColumns: SearchColStr をパースして列番号の配列を返す
' SearchColStr が空の場合は定義済み全フィールド列を返す
'-------------------------------------------------------------
Public Function GetSearchColumns() As Long()
    Dim cols() As Long

    If Trim(SearchColStr) = "" Then
        ' 未設定時は定義されている全列を検索対象にする
        ReDim cols(4)
        cols(0) = Col_案件名
        cols(1) = Col_案件番号
        cols(2) = Col_顧客名
        cols(3) = Col_担当者名
        cols(4) = Col_期日
    Else
        Dim parts() As String
        parts = Split(SearchColStr, ",")
        ReDim cols(UBound(parts))
        Dim i As Integer
        For i = 0 To UBound(parts)
            cols(i) = ColNumOrLetter(Trim(parts(i)))
        Next i
    End If

    GetSearchColumns = cols
End Function

'-------------------------------------------------------------
' IsValid: 最低限必要な設定が揃っているか確認する
'-------------------------------------------------------------
Public Function IsValid() As Boolean
    IsValid = (Trim(FilePath) <> "" And Trim(SheetName) <> "" And HeaderRow > 0)
End Function

'-------------------------------------------------------------
' GetFieldColumns: フィールド名と列番号のマッピングを配列で返す（テスト表示用）
'-------------------------------------------------------------
Public Function GetFieldColumns() As Variant
    GetFieldColumns = Array( _
        Array("案件名", Col_案件名), _
        Array("案件番号", Col_案件番号), _
        Array("顧客名", Col_顧客名), _
        Array("担当者名", Col_担当者名), _
        Array("期日", Col_期日) _
    )
End Function
