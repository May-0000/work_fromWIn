Attribute VB_Name = "Module1"
Option Explicit
'*********************************
'       �|�b�v�A�b�v�̏���
'�E�N�ԏW�v
'�E�o�b�N�A�b�v
'�E������
'*********************************

'�N�ԏW�v�p�֐��i�N�ԏW�v�V�[�g�L���m�F�j
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

'�N�ԏW�v�i�����V�[�g�폜�A���s�m�F�A�o�b�N�A�b�v�Ώۊm�F�j
Sub yearTotal()
    If ExistsWorkSheet("�N�ԏW�v") = True Then
        Application.DisplayAlerts = False
        Worksheets("�N�ԏW�v").Delete
        Application.DisplayAlerts = True
    End If

    Dim res As Integer
    res = MsgBox("�N�ԏW�v���܂����H", vbOKCancel)
    
    If res = 1 Then
    
        If ThisWorkbook.Sheets.Count < 2 Then
            MsgBox "�N�ԏW�v����V�[�g������܂���B"
            Exit Sub
        End If
        
        Call yearCreate
        
    End If
End Sub

'�o�b�N�A�b�v�����i���s�m�F�A�o�b�N�A�b�v�Ώۊm�F�j
Sub BkUp()
    Dim res As Integer
    res = MsgBox("�o�b�N�A�b�v���쐬���܂����B", vbOKCancel)
    
    If res = 1 Then
    
        If ThisWorkbook.Sheets.Count < 2 Then
            MsgBox "�o�b�N�A�b�v����V�[�g������܂���B"
            Exit Sub
        End If
        
        Call bk
    End If
    
End Sub

'�������i���s�m�F�j
Sub Initialization()
    Dim res As Integer
    res = MsgBox("���������܂����B" & vbNewLine & "���炩���߃o�b�N�A�b�v���擾���邱�Ƃ𐄏����܂��B", vbOKCancel)
    If res = 1 Then
        Call initial
    End If
End Sub
