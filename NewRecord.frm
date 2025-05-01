VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NewRecord 
   Caption         =   "�x�o����"
   ClientHeight    =   6820
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   7760
   OleObjectBlob   =   "NewRecord.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "newRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**************************************
'   �x�o���͂̃��[�U�[�t�H�[������
'**************************************

'�g�p�W�������A�����x���X�g���l�ݒ�A���e�̓��͐ݒ�
Private Sub UserForm_Initialize()
    TextBox2.MaxLength = 10
    TextBox3.WordWrap = True
    TextBox3.EnterKeyBehavior = True
    
    With ListBox1
        .AddItem "�H��"
        .AddItem "�O�H��"
        .AddItem "���M��"
        .AddItem "������"
        .AddItem "�ʐM��"
        .AddItem "���p�i"
        .AddItem "�ƒ�"
        .AddItem "�ߕ�"
        .AddItem "���e��"
        .AddItem "�"
        .AddItem "��ʔ�"
        .AddItem "���۔�"
        .AddItem "���ʔ�"
        .AddItem "�o��"
    End With
    
    Dim i As Integer
    For i = 1 To 10
        ListBox2.AddItem i
    Next i
    
    TextBox2.SetFocus
    
End Sub

'���t�X�y�[�X����(�����ƃX���b�V���̂ݎ�t�j
Private Sub TextBox2_Change()
    Dim strDate As String
    strDate = Trim(TextBox2.Text)
    
    If strDate = "" Then
        Exit Sub
    End If
    
    TextBox2.Text = strDate
End Sub

'���t�t�H�[�}�b�g�`�F�b�N�iYYYY/MM/DD�j
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
    MsgBox "���t�͐��������͌`��(YYYY/MM/DD)�œ��͂��Ă��������B"
    TextBox2.BackColor = vbRed
    
    
End Sub

'���z���̓t�H�[�}�b�g�i���l�̂ݎ󂯕t���j
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

'�����̓G���[����
Private Sub ListBox1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Cancel = False
    ListBox1.BackColor = vbWhite
End Sub

'�����̓G���[����
Private Sub ListBox2_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Cancel = False
    ListBox2.BackColor = vbWhite
End Sub

'���e���͕����������i25�����ȓ����󂯕t���j
Private Sub TextBox3_Change()
    If Len(TextBox3.Text) = 0 Then Exit Sub
    If Len(TextBox3.Text) > 25 Then
        TextBox3 = Left(TextBox3.Text, 25)
    End If
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
    ElseIf ListBox1.Text = "" Then
        w = 3
        GoTo Nullerror
    ElseIf ListBox2.Text = "" Then
        w = 4
        GoTo Nullerror
    End If
    
    '���͑ΏۃV�[�g���݃`�F�b�N�A�������b�Z�[�W
    If Sheets.Count < 2 Then
        MsgBox "���ԃV�[�g���쐬���Ă��������B"
        Unload newRecord
        Exit Sub
    Else
        Dim i As Integer
        For i = 1 To Worksheets.Count
        
            Dim inYm As String
            inYm = Left(newRecord.TextBox2.Text, 4) & "�N" & Mid(newRecord.TextBox2.Text, 6, 2) & "��"
            Debug.Print "���͌��F" & inYm
            Debug.Print "�Z��A1�F" & Sheets(i).Cells(1, 1).Text
            
            If inYm = Sheets(i).Cells(1, 1).Text Then
                MsgBox "�o�^���������܂����B"
                Call newRecordAdd
                Unload newRecord
                Exit Sub
            End If
        Next i
            MsgBox "���͂��������̌��ԃV�[�g���쐬���Ă��������B"
            Unload newRecord
            Exit Sub
    End If
    
    Exit Sub
    
Nullerror:
MsgBox "�����͂̍��ڂ�����܂��B"
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

'�L�����Z���{�^�������㏈��
Private Sub CommandButton2_Click()
    Unload newRecord
End Sub


