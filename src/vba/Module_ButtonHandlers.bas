Attribute VB_Name = "Module_ButtonHandlers"
Option Explicit

'=============================================================
' Module_ButtonHandlers: ボタンハンドラーモジュール
' フォームコントロールボタンの OnAction はパラメータを渡せないため、
' テンプレートIDを固定した薄いラッパーサブルーチンを事前定義する
'
' テンプレート起動:  Launch_1 ～ Launch_30  （ID 1～30 に対応）
' ファイルパス参照:  BrowseFile_1 ～ BrowseFile_20
' 接続テスト:       TestConn_1 ～ TestConn_20
'
' ID が上限を超えた場合は、このモジュールに手動で追加するか、
' テンプレート一覧で直接マクロ名を入力してください。
'=============================================================

' ============================================================
' テンプレート起動ボタン (Launch_1 ～ Launch_30)
' ============================================================

Sub Launch_1():  Module_Launcher.LaunchTemplate 1:  End Sub
Sub Launch_2():  Module_Launcher.LaunchTemplate 2:  End Sub
Sub Launch_3():  Module_Launcher.LaunchTemplate 3:  End Sub
Sub Launch_4():  Module_Launcher.LaunchTemplate 4:  End Sub
Sub Launch_5():  Module_Launcher.LaunchTemplate 5:  End Sub
Sub Launch_6():  Module_Launcher.LaunchTemplate 6:  End Sub
Sub Launch_7():  Module_Launcher.LaunchTemplate 7:  End Sub
Sub Launch_8():  Module_Launcher.LaunchTemplate 8:  End Sub
Sub Launch_9():  Module_Launcher.LaunchTemplate 9:  End Sub
Sub Launch_10(): Module_Launcher.LaunchTemplate 10: End Sub
Sub Launch_11(): Module_Launcher.LaunchTemplate 11: End Sub
Sub Launch_12(): Module_Launcher.LaunchTemplate 12: End Sub
Sub Launch_13(): Module_Launcher.LaunchTemplate 13: End Sub
Sub Launch_14(): Module_Launcher.LaunchTemplate 14: End Sub
Sub Launch_15(): Module_Launcher.LaunchTemplate 15: End Sub
Sub Launch_16(): Module_Launcher.LaunchTemplate 16: End Sub
Sub Launch_17(): Module_Launcher.LaunchTemplate 17: End Sub
Sub Launch_18(): Module_Launcher.LaunchTemplate 18: End Sub
Sub Launch_19(): Module_Launcher.LaunchTemplate 19: End Sub
Sub Launch_20(): Module_Launcher.LaunchTemplate 20: End Sub
Sub Launch_21(): Module_Launcher.LaunchTemplate 21: End Sub
Sub Launch_22(): Module_Launcher.LaunchTemplate 22: End Sub
Sub Launch_23(): Module_Launcher.LaunchTemplate 23: End Sub
Sub Launch_24(): Module_Launcher.LaunchTemplate 24: End Sub
Sub Launch_25(): Module_Launcher.LaunchTemplate 25: End Sub
Sub Launch_26(): Module_Launcher.LaunchTemplate 26: End Sub
Sub Launch_27(): Module_Launcher.LaunchTemplate 27: End Sub
Sub Launch_28(): Module_Launcher.LaunchTemplate 28: End Sub
Sub Launch_29(): Module_Launcher.LaunchTemplate 29: End Sub
Sub Launch_30(): Module_Launcher.LaunchTemplate 30: End Sub

' ============================================================
' ファイルパス参照ボタン (BrowseFile_1 ～ BrowseFile_20)
' ============================================================

Sub BrowseFile_1():  Module_FileIO.BrowseFilePath 1:  End Sub
Sub BrowseFile_2():  Module_FileIO.BrowseFilePath 2:  End Sub
Sub BrowseFile_3():  Module_FileIO.BrowseFilePath 3:  End Sub
Sub BrowseFile_4():  Module_FileIO.BrowseFilePath 4:  End Sub
Sub BrowseFile_5():  Module_FileIO.BrowseFilePath 5:  End Sub
Sub BrowseFile_6():  Module_FileIO.BrowseFilePath 6:  End Sub
Sub BrowseFile_7():  Module_FileIO.BrowseFilePath 7:  End Sub
Sub BrowseFile_8():  Module_FileIO.BrowseFilePath 8:  End Sub
Sub BrowseFile_9():  Module_FileIO.BrowseFilePath 9:  End Sub
Sub BrowseFile_10(): Module_FileIO.BrowseFilePath 10: End Sub
Sub BrowseFile_11(): Module_FileIO.BrowseFilePath 11: End Sub
Sub BrowseFile_12(): Module_FileIO.BrowseFilePath 12: End Sub
Sub BrowseFile_13(): Module_FileIO.BrowseFilePath 13: End Sub
Sub BrowseFile_14(): Module_FileIO.BrowseFilePath 14: End Sub
Sub BrowseFile_15(): Module_FileIO.BrowseFilePath 15: End Sub
Sub BrowseFile_16(): Module_FileIO.BrowseFilePath 16: End Sub
Sub BrowseFile_17(): Module_FileIO.BrowseFilePath 17: End Sub
Sub BrowseFile_18(): Module_FileIO.BrowseFilePath 18: End Sub
Sub BrowseFile_19(): Module_FileIO.BrowseFilePath 19: End Sub
Sub BrowseFile_20(): Module_FileIO.BrowseFilePath 20: End Sub

' ============================================================
' 接続テストボタン (TestConn_1 ～ TestConn_20)
' ============================================================

Sub TestConn_1():  Module_FileIO.TestConnection 1:  End Sub
Sub TestConn_2():  Module_FileIO.TestConnection 2:  End Sub
Sub TestConn_3():  Module_FileIO.TestConnection 3:  End Sub
Sub TestConn_4():  Module_FileIO.TestConnection 4:  End Sub
Sub TestConn_5():  Module_FileIO.TestConnection 5:  End Sub
Sub TestConn_6():  Module_FileIO.TestConnection 6:  End Sub
Sub TestConn_7():  Module_FileIO.TestConnection 7:  End Sub
Sub TestConn_8():  Module_FileIO.TestConnection 8:  End Sub
Sub TestConn_9():  Module_FileIO.TestConnection 9:  End Sub
Sub TestConn_10(): Module_FileIO.TestConnection 10: End Sub
Sub TestConn_11(): Module_FileIO.TestConnection 11: End Sub
Sub TestConn_12(): Module_FileIO.TestConnection 12: End Sub
Sub TestConn_13(): Module_FileIO.TestConnection 13: End Sub
Sub TestConn_14(): Module_FileIO.TestConnection 14: End Sub
Sub TestConn_15(): Module_FileIO.TestConnection 15: End Sub
Sub TestConn_16(): Module_FileIO.TestConnection 16: End Sub
Sub TestConn_17(): Module_FileIO.TestConnection 17: End Sub
Sub TestConn_18(): Module_FileIO.TestConnection 18: End Sub
Sub TestConn_19(): Module_FileIO.TestConnection 19: End Sub
Sub TestConn_20(): Module_FileIO.TestConnection 20: End Sub
