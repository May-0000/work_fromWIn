Attribute VB_Name = "Module1"
Option Explicit
'*********************************
'       ポップアップの処理
'・年間集計
'・バックアップ
'・初期化
'*********************************

'年間集計用関数（年間集計シート有無確認）
Function ExistsWorkSheet(sheetName) As Boolean
    Dim b As Boolean
    b = False
    
    Dim sh
    For Each sh In Worksheets
        If sh.Name = sheetName Then
            b = True
            Exit For
        End If
    Next sh
    
    ExistsWorkSheet = b
End Function

'年間集計（既存シート削除、実行確認、バックアップ対象確認）
Sub yearTotal()
    If ExistsWorkSheet("年間集計") = True Then
        Application.DisplayAlerts = False
        Worksheets("年間集計").Delete
        Application.DisplayAlerts = True
    End If

    Dim res As Integer
    res = MsgBox("年間集計しますか？", vbOKCancel)
    
    If res = 1 Then
    
        If ThisWorkbook.Sheets.Count < 2 Then
            MsgBox "年間集計するシートがありません。"
            Exit Sub
        End If
        
        Call yearCreate
        
    End If
End Sub

'バックアップ処理（実行確認、バックアップ対象確認）
Sub BkUp()
    Dim res As Integer
    res = MsgBox("バックアップを作成しますか。", vbOKCancel)
    
    If res = 1 Then
    
        If ThisWorkbook.Sheets.Count < 2 Then
            MsgBox "バックアップするシートがありません。"
            Exit Sub
        End If
        
        Call bk
    End If
    
End Sub

'初期化（実行確認）
Sub Initialization()
    Dim res As Integer
    res = MsgBox("初期化しますか。" & vbNewLine & "あらかじめバックアップを取得することを推奨します。", vbOKCancel)
    If res = 1 Then
        Call initial
    End If
End Sub
