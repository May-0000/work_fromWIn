VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NewMonth 
   Caption         =   "�V�K�쐬"
   ClientHeight    =   3360
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   6000
   OleObjectBlob   =   "NewMonth.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "NewMonth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'******************************************
'   ���ԃV�[�g�쐬�̃��[�U�[�t�H�[������
'******************************************

'�����������ݒ�A���̓X�^�[�g�ʒu�ݒ�
Private Sub UserForm_Initialize()
    TextBox1.MaxLength = 2
    TextBox3.MaxLength = 4
    TextBox3.SetFocus
End Sub

'�N_���͐���i���l�̂ݎ󂯕t���j
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

'�N_���̓`�F�b�N�G���[���b�Z�[�W�i1000-2100�N�̊Ԃ̔N�̂ݗL���j
Private Sub TextBox3_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If TextBox3.Text = "" Then Exit Sub

    If TextBox3.Text > 2100 Or TextBox3.Text < 1000 Then
        MsgBox "�K�؂Ȑ�������͂��Ă��������B"
        TextBox3.BackColor = vbRed
        Cancel = False
        Exit Sub
    End If
    
    TextBox3.BackColor = vbWhite

End Sub

'��_���͐����i���l�̂ݎ󂯕t���j
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

'��_���̓`�F�b�N�G���[���b�Z�[�W�i1-12�܂Ŏ󂯕t���j
Private Sub TextBox1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If TextBox1.Text = "" Then Exit Sub

    If TextBox1.Text > 12 Or TextBox1.Text = 0 Then
        MsgBox "�K�؂Ȑ�������͂��Ă��������B"
        TextBox1.BackColor = vbRed
        Cancel = False
        Exit Sub
    End If
    
    Dim mtext As String
    mtext = Format(TextBox1.Text, "00")
    TextBox1.Text = mtext
    
    TextBox1.BackColor = vbWhite

End Sub

'�\�Z�̓��͐����i���l�̂ݎ󂯕t���j
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

'�o�^�{�^�������㏈���i�����̓`�F�b�N�A�d���`�F�b�N�A�������b�Z�[�W�j
Private Sub CommandButton1_Click()
    '�����̓`�F�b�N
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
        
    '�d���`�F�b�N
    Dim i As Integer
    For i = 2 To Worksheets.Count
        Dim inp As String
        inp = TextBox3.Text & "�N" & TextBox1.Text & "��"
        
        If inp = Sheets(i).Name Then
            MsgBox "���ɍ쐬�ς݂̌��ł��B" & vbCrLf & _
                    "�V�[�g���m�F���Ă��������B"
            TextBox1.BackColor = vbRed
            TextBox3.BackColor = vbRed
            Exit Sub
        End If
    Next i
    
    '�������b�Z�[�W
    MsgBox "�V���������쐬���܂����B"
    Call NewMonthCreate
    Unload NewMonth
    
    Exit Sub
    
Nullerror:
MsgBox "�����͂̍��ڂ�����܂��B"
If w = 3 Then
TextBox3.BackColor = vbRed
ElseIf w = 1 Then
TextBox1.BackColor = vbRed
ElseIf w = 2 Then
TextBox2.BackColor = vbRed
End If

End Sub

'�L�����Z���{�^�������㏈��
Private Sub CommandButton2_Click()
    Unload NewMonth
End Sub


