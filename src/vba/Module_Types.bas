Attribute VB_Name = "Module_Types"
Option Explicit

'=============================================================
' Module_Types: 共通データ型定義モジュール
' 全モジュールで使用するユーザー定義型(UDT)を定義する
' 注意: このモジュールは他の全モジュールより先にインポートすること
'=============================================================

'-------------------------------------------------------------
' ProjectData: 案件（プロジェクト）情報
'-------------------------------------------------------------
Public Type ProjectData
    案件名      As String   ' プロジェクト名称
    案件番号    As String   ' プロジェクトID/番号
    顧客名      As String   ' 顧客・取引先名
    担当者名    As String   ' 担当者氏名
    期日        As String   ' 納期・締切日（文字列として保持）
    SourceFile  As String   ' 取得元ファイルパス（表示用）
    SourceRow   As Long     ' 取得元行番号（デバッグ用）
End Type

'-------------------------------------------------------------
' FileSetting: 外部Excelファイルの設定情報
'-------------------------------------------------------------
Public Type FileSetting
    SettingID       As Long     ' 設定行の識別番号
    DisplayName     As String   ' 表示名（例: "営業案件管理表"）
    FilePath        As String   ' 外部ファイルの絶対パス
    SheetName       As String   ' 対象シート名
    HeaderRow       As Long     ' ヘッダー行番号（通常1）
    Col_案件名      As Long     ' 案件名の列番号（0=未設定）
    Col_案件番号    As Long     ' 案件番号の列番号（0=未設定）
    Col_顧客名      As Long     ' 顧客名の列番号（0=未設定）
    Col_担当者名    As Long     ' 担当者名の列番号（0=未設定）
    Col_期日        As Long     ' 期日の列番号（0=未設定）
    SearchColStr    As String   ' 検索対象列番号（カンマ区切り、例: "1,2,3"）
    IsActive        As Boolean  ' 有効/無効フラグ
End Type
