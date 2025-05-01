VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SumIncome 
   Caption         =   "����"
   ClientHeight    =   3640
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   5470
   OleObjectBlob   =   "SumIncome.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "Sumincome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**************************************
'   �������͂̃��[�U�[�t�H�[������
'**************************************

'�����������ݒ�A���̓X�^�[�g�ʒu�ݒ�
Private Sub UserForm_Initialize()
    TextBox2.MaxLength = 2
    TextBox2.SetFocus
End Sub

'��_���͐����i���l�̂ݎ󂯕t���j
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

'��_���̓`�F�b�N�G���[���b�Z�[�W�i1-12�̂ݎ󂯕t���j
Private Sub TextBox2_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If TextBox2.Text = "" Then Exit Sub

    If TextBox2.Text > 12 Or TextBox2.Text = 0 Then
        MsgBox "�K�؂Ȑ�������͂��Ă��������B"
        TextBox2.BackColor = vbRed
        Cancel = False
        Exit Sub
    End If
    
    Dim inmonth As String
    inmonth = Format(TextBox2.Text, "00")
    TextBox2.Text = inmonth
    
    TextBox2.BackColor = vbWhite

End Sub

'���z_���͐����i���l�̂ݎ󂯕t���j
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

'�o�^�{�^�������㏈���i�����̓`�F�b�N�A���͑ΏۃV�[�g���݃`�F�b�N�A�������b�Z�[�W�j
Private Sub CommandButton1_Click()
    '�����̓`�F�b�N
    Dim w As Integer
    w = 0
    If TextBox2.Text = "" Then
        w = 2
        GoTo Nullerror
    ElseIf TextBox1.Text = "" Then
        w = 1
        GoTo Nullerror
    End If

    '���͑ΏۃV�[�g���݃`�F�b�N�A�������b�Z�[�W
    If Sheets.Count < 2 Then
        MsgBox "���ԃV�[�g���쐬���Ă��������B"
        Unload Sumincome
        Exit Sub
    Else
        Dim i As Integer
        For i = 1 To Worksheets.Count
            Dim month As String
            month = Replace(Mid(Sheets(i).Cells(1, 1).Text, 6), "��", "")
            Debug.Print "��" & month

            If Sumincome.TextBox2.Text = month Then
                MsgBox "�o�^���܂����B"
                Call SumIncomeDisplay
                Unload Sumincome
                Exit Sub
            End If
        Next i
            MsgBox "���͂��������̌��ԃV�[�g���쐬���Ă��������B"
            Unload Sumincome
            Exit Sub
    End If
    
    Exit Sub
    
Nullerror:
MsgBox "�����͂̍��ڂ�����܂��B"
If w = 2 Then
    TextBox2.BackColor = vbRed
ElseIf w = 1 Then
    TextBox1.BackColor = vbRed
End If

End Sub

'�L�����Z���{�^�������㏈��
Private Sub CommandButton2_Click()
    Unload Sumincome
End Sub
