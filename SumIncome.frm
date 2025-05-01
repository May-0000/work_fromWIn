VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SumIncome 
   Caption         =   "収入"
   ClientHeight    =   3640
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   5470
   OleObjectBlob   =   "SumIncome.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "Sumincome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**************************************
'   収入入力のユーザーフォーム処理
'**************************************

'文字数制限設定、入力スタート位置設定
Private Sub UserForm_Initialize()
    TextBox2.MaxLength = 2
    TextBox2.SetFocus
End Sub

'月_入力制限（数値のみ受け付け）
Private Sub TextBox2_Change()
    If Len(TextBox2.Text) = 0 Then Exit Sub
    
    Dim a
    Dim i
    
    a = TextBox2.Text
    
    For i = Len(a) To 1 Step -1
        If IsNumeric(Mid(a, i, 1)) = False Then
            a = Replace(a, Mid(a, i, 1), "")
        End If
    Next i
    
    TextBox2.Text = a
            
End Sub

'月_入力チェックエラーメッセージ（1-12のみ受け付け）
Private Sub TextBox2_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If TextBox2.Text = "" Then Exit Sub

    If TextBox2.Text > 12 Or TextBox2.Text = 0 Then
        MsgBox "適切な数字を入力してください。"
        TextBox2.BackColor = vbRed
        Cancel = False
        Exit Sub
    End If
    
    Dim inmonth As String
    inmonth = Format(TextBox2.Text, "00")
    TextBox2.Text = inmonth
    
    TextBox2.BackColor = vbWhite

End Sub

'金額_入力制限（数値のみ受け付け）
Private Sub TextBox1_Change()
    If Len(TextBox1.Text) = 0 Then Exit Sub
    
    Dim a
    Dim i
    
    a = TextBox1.Text
    
    For i = Len(TextBox1) To 1 Step -1
        If IsNumeric(Mid(a, i, 1)) = False Then
            a = Replace(a, Mid(a, i, 1), "")
        End If
    Next i
    
    TextBox1.Text = a
        
End Sub

'登録ボタン押下後処理（未入力チェック、入力対象シート存在チェック、完了メッセージ）
Private Sub CommandButton1_Click()
    '未入力チェック
    Dim w As Integer
    w = 0
    If TextBox2.Text = "" Then
        w = 2
        GoTo Nullerror
    ElseIf TextBox1.Text = "" Then
        w = 1
        GoTo Nullerror
    End If

    '入力対象シート存在チェック、完了メッセージ
    If Sheets.Count < 2 Then
        MsgBox "月間シートを作成してください。"
        Unload Sumincome
        Exit Sub
    Else
        Dim i As Integer
        For i = 1 To Worksheets.Count
            Dim month As String
            month = Replace(Mid(Sheets(i).Cells(1, 1).Text, 6), "月", "")
            Debug.Print "月" & month

            If Sumincome.TextBox2.Text = month Then
                MsgBox "登録しました。"
                Call SumIncomeDisplay
                Unload Sumincome
                Exit Sub
            End If
        Next i
            MsgBox "入力したい月の月間シートを作成してください。"
            Unload Sumincome
            Exit Sub
    End If
    
    Exit Sub
    
Nullerror:
MsgBox "未入力の項目があります。"
If w = 2 Then
    TextBox2.BackColor = vbRed
ElseIf w = 1 Then
    TextBox1.BackColor = vbRed
End If

End Sub

'キャンセルボタン押下後処理
Private Sub CommandButton2_Click()
    Unload Sumincome
End Sub
