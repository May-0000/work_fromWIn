VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NewMonth 
   Caption         =   "新規作成"
   ClientHeight    =   3360
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   6000
   OleObjectBlob   =   "NewMonth.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "NewMonth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'******************************************
'   月間シート作成のユーザーフォーム処理
'******************************************

'文字数制限設定、入力スタート位置設定
Private Sub UserForm_Initialize()
    TextBox1.MaxLength = 2
    TextBox3.MaxLength = 4
    TextBox3.SetFocus
End Sub

'年_入力制御（数値のみ受け付け）
Private Sub TextBox3_Change()
    If Len(TextBox3.Text) = 0 Then Exit Sub
    
    Dim a
    Dim i
    
    a = TextBox3.Text
    
    For i = Len(a) To 1 Step -1
        If IsNumeric(Mid(a, i, 1)) = False Then
            a = Replace(a, Mid(a, i, 1), "")
        End If
    Next i
    
    TextBox3.Text = a
            
End Sub

'年_入力チェックエラーメッセージ（1000-2100年の間の年のみ有効）
Private Sub TextBox3_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If TextBox3.Text = "" Then Exit Sub

    If TextBox3.Text > 2100 Or TextBox3.Text < 1000 Then
        MsgBox "適切な数字を入力してください。"
        TextBox3.BackColor = vbRed
        Cancel = False
        Exit Sub
    End If
    
    TextBox3.BackColor = vbWhite

End Sub

'月_入力制限（数値のみ受け付け）
Private Sub TextBox1_Change()
    If Len(TextBox1.Text) = 0 Then Exit Sub
    
    Dim a
    Dim i
    
    a = TextBox1.Text
    
    For i = Len(a) To 1 Step -1
        If IsNumeric(Mid(a, i, 1)) = False Then
            a = Replace(a, Mid(a, i, 1), "")
        End If
    Next i
    
    TextBox1.Text = a
            
End Sub

'月_入力チェックエラーメッセージ（1-12まで受け付け）
Private Sub TextBox1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If TextBox1.Text = "" Then Exit Sub

    If TextBox1.Text > 12 Or TextBox1.Text = 0 Then
        MsgBox "適切な数字を入力してください。"
        TextBox1.BackColor = vbRed
        Cancel = False
        Exit Sub
    End If
    
    Dim mtext As String
    mtext = Format(TextBox1.Text, "00")
    TextBox1.Text = mtext
    
    TextBox1.BackColor = vbWhite

End Sub

'予算の入力制限（数値のみ受け付け）
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

'登録ボタン押下後処理（未入力チェック、重複チェック、完了メッセージ）
Private Sub CommandButton1_Click()
    '未入力チェック
    Dim w As Integer
    If TextBox3.Text = "" Then
        w = 3
        GoTo Nullerror
    ElseIf TextBox1.Text = "" Then
        w = 1
        GoTo Nullerror
    ElseIf TextBox2.Text = "" Then
        w = 2
        GoTo Nullerror
    End If
        
    '重複チェック
    Dim i As Integer
    For i = 2 To Worksheets.Count
        Dim inp As String
        inp = TextBox3.Text & "年" & TextBox1.Text & "月"
        
        If inp = Sheets(i).Name Then
            MsgBox "既に作成済みの月です。" & vbCrLf & _
                    "シートを確認してください。"
            TextBox1.BackColor = vbRed
            TextBox3.BackColor = vbRed
            Exit Sub
        End If
    Next i
    
    '完了メッセージ
    MsgBox "新しい月を作成しました。"
    Call NewMonthCreate
    Unload NewMonth
    
    Exit Sub
    
Nullerror:
MsgBox "未入力の項目があります。"
If w = 3 Then
TextBox3.BackColor = vbRed
ElseIf w = 1 Then
TextBox1.BackColor = vbRed
ElseIf w = 2 Then
TextBox2.BackColor = vbRed
End If

End Sub

'キャンセルボタン押下後処理
Private Sub CommandButton2_Click()
    Unload NewMonth
End Sub


