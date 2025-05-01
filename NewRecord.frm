VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NewRecord 
   Caption         =   "支出入力"
   ClientHeight    =   6820
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   7760
   OleObjectBlob   =   "NewRecord.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "newRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**************************************
'   支出入力のユーザーフォーム処理
'**************************************

'使用ジャンル、満足度リスト数値設定、内容の入力設定
Private Sub UserForm_Initialize()
    TextBox2.MaxLength = 10
    TextBox3.WordWrap = True
    TextBox3.EnterKeyBehavior = True
    
    With ListBox1
        .AddItem "食費"
        .AddItem "外食費"
        .AddItem "光熱費"
        .AddItem "水道代"
        .AddItem "通信費"
        .AddItem "日用品"
        .AddItem "家賃"
        .AddItem "衣服"
        .AddItem "美容代"
        .AddItem "趣味"
        .AddItem "交通費"
        .AddItem "交際費"
        .AddItem "特別費"
        .AddItem "経費"
    End With
    
    Dim i As Integer
    For i = 1 To 10
        ListBox2.AddItem i
    Next i
    
    TextBox2.SetFocus
    
End Sub

'日付スペース処理(数字とスラッシュのみ受付）
Private Sub TextBox2_Change()
    Dim strDate As String
    strDate = Trim(TextBox2.Text)
    
    If strDate = "" Then
        Exit Sub
    End If
    
    TextBox2.Text = strDate
End Sub

'日付フォーマットチェック（YYYY/MM/DD）
Private Sub TextBox2_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    Dim strDate As String
    strDate = TextBox2.Text
    Dim i
    
    Debug.Print "strDate: " & strDate
    Debug.Print "Length: " & Len(strDate)
    
    For i = Len(strDate) To 1 Step -1
    
        If i = 8 Or i = 5 Then
            Debug.Print "Error at position1: " & i
            Debug.Print "Error at position_contens: " & Mid(strDate, i, 1)
            If StrComp(Mid(strDate, i, 1), "/") <> 0 Then
                Debug.Print "Error at position2: " & i
                GoTo Formmaterror
                Cancel = False
            End If
        Else
            If IsNumeric(Mid(strDate, i, 1)) = False Then
                Debug.Print "Error at position3: " & i
                GoTo Formmaterror
                Cancel = False
            End If
        End If
            
    Next i
    
    If Len(strDate) < 10 And Len(strDate) > 0 Then
        GoTo Formmaterror
    End If
    
    TextBox2.BackColor = vbWhite
    
Exit Sub

Formmaterror:
    MsgBox "日付は正しい入力形式(YYYY/MM/DD)で入力してください。"
    TextBox2.BackColor = vbRed
    
    
End Sub

'金額入力フォーマット（数値のみ受け付け）
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
    
    TextBox1.BackColor = vbWhite

End Sub

'未入力エラー解除
Private Sub ListBox1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Cancel = False
    ListBox1.BackColor = vbWhite
End Sub

'未入力エラー解除
Private Sub ListBox2_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Cancel = False
    ListBox2.BackColor = vbWhite
End Sub

'内容入力文字数制限（25文字以内を受け付け）
Private Sub TextBox3_Change()
    If Len(TextBox3.Text) = 0 Then Exit Sub
    If Len(TextBox3.Text) > 25 Then
        TextBox3 = Left(TextBox3.Text, 25)
    End If
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
    ElseIf ListBox1.Text = "" Then
        w = 3
        GoTo Nullerror
    ElseIf ListBox2.Text = "" Then
        w = 4
        GoTo Nullerror
    End If
    
    '入力対象シート存在チェック、完了メッセージ
    If Sheets.Count < 2 Then
        MsgBox "月間シートを作成してください。"
        Unload newRecord
        Exit Sub
    Else
        Dim i As Integer
        For i = 1 To Worksheets.Count
        
            Dim inYm As String
            inYm = Left(newRecord.TextBox2.Text, 4) & "年" & Mid(newRecord.TextBox2.Text, 6, 2) & "月"
            Debug.Print "入力月：" & inYm
            Debug.Print "セルA1：" & Sheets(i).Cells(1, 1).Text
            
            If inYm = Sheets(i).Cells(1, 1).Text Then
                MsgBox "登録が完了しました。"
                Call newRecordAdd
                Unload newRecord
                Exit Sub
            End If
        Next i
            MsgBox "入力したい月の月間シートを作成してください。"
            Unload newRecord
            Exit Sub
    End If
    
    Exit Sub
    
Nullerror:
MsgBox "未入力の項目があります。"
If w = 2 Then
    TextBox2.BackColor = vbRed
ElseIf w = 1 Then
    TextBox1.BackColor = vbRed
ElseIf w = 3 Then
    ListBox1.BackColor = vbRed
ElseIf w = 4 Then
    ListBox2.BackColor = vbRed
End If

End Sub

'キャンセルボタン押下後処理
Private Sub CommandButton2_Click()
    Unload newRecord
End Sub


