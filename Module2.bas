Attribute VB_Name = "Module2"
Option Explicit
'*********************************
'            ���C������
'*********************************

'-------------------
'���ԃV�[�g�쐬����
'-------------------
Sub NewMonthCreate()
    '**�V�[�g���ݒ�**
    Dim year As String
    Dim month As String
    
    year = NewMonth.TextBox3.Text
    month = NewMonth.TextBox1.Text

    Worksheets.add After:=Worksheets(Worksheets.Count)
    ActiveSheet.Name = year & "�N" & month & "��"
    
    With Cells(1, 1)
        .NumberFormat = "@"
        .Value = year & "�N" & month & "��"
    End With

    '**�x�o�\�쐬**
    Cells(2, 2) = "���t"
    Cells(2, 3) = "�j��"
    Cells(2, 4) = "�W������"
    Cells(2, 5) = "���e"
    Cells(2, 6) = "�����x"
    Cells(2, 7) = "���z"

    Range("B2:G2").BorderAround LineStyle:=xlContinuous
    Range("B2:G2").Borders(xlInsideVertical).LineStyle = xlContinuous
    Range("B2:G2").Borders(xlInsideHorizontal).LineStyle = xlContinuous

    Range("B2:G2").Interior.Color = RGB(255, 220, 140)

    '**���v�x�o�\�쐬**
    Cells(2, 9) = "TOTAL"
    Cells(3, 9) = "�\�Z"
    Cells(4, 9) = "�\�Z�B��"

    Cells(3, 10) = NewMonth.TextBox2.Text & "�~"

    Range("I2:J4").BorderAround LineStyle:=xlContinuous
    Range("I2:J4").Borders(xlInsideVertical).LineStyle = xlContinuous
    Range("I2:J4").Borders(xlInsideHorizontal).LineStyle = xlContinuous

    Range("I2:I4").Interior.Color = RGB(255, 220, 140)
    
    '**���x�\�쐬**
    Cells(6, 9) = "�������v"
    Cells(7, 9) = "���x"

    Cells(2, 10) = 0 & "�~"
    Cells(6, 10) = 0 & "�~"
    Cells(7, 10) = 0 & "�~"

    Range("I6:J7").BorderAround LineStyle:=xlContinuous
    Range("I6:J7").Borders(xlInsideVertical).LineStyle = xlContinuous
    Range("I6:J7").Borders(xlInsideHorizontal).LineStyle = xlContinuous

    Range("I6:I7").Interior.Color = RGB(255, 220, 140)

    '**�W�������\�쐬**
    Cells(2, 12) = "�W������"
    Cells(2, 13) = "���v���z"
    Cells(3, 12) = "�H��"
    Cells(4, 12) = "�O�H��"
    Cells(5, 12) = "���M��"
    Cells(6, 12) = "������"
    Cells(7, 12) = "�ʐM��"
    Cells(8, 12) = "���p�i"
    Cells(9, 12) = "�ƒ�"
    Cells(10, 12) = "�ߕ�"
    Cells(11, 12) = "���e��"
    Cells(12, 12) = "�"
    Cells(13, 12) = "��ʔ�"
    Cells(14, 12) = "���۔�"
    Cells(15, 12) = "���ʔ�"
    Cells(16, 12) = "�o��"
    
    Dim i As Integer
    For i = 3 To 16
        Cells(i, 13) = 0
    Next i
    
    Range("L2:M16").BorderAround LineStyle:=xlContinuous
    Range("L2:M16").Borders(xlInsideVertical).LineStyle = xlContinuous
    Range("L2:M16").Borders(xlInsideHorizontal).LineStyle = xlContinuous

    Range("L2:M2").Interior.Color = RGB(255, 220, 140)
    
    '**�~�O���t�쐬**
    Dim chartObj As ChartObject
    Set chartObj = ActiveSheet.ChartObjects.add(Left:=100, Top:=100, Width:=400, Height:=300)

    With chartObj.chart
        .ChartType = xlPie
        .SetSourceData Source:=ActiveSheet.Range(ActiveSheet.Cells(2, 12), ActiveSheet.Cells(16, 13))
        .HasTitle = True
        .ChartTitle.Text = "�W�������ʎg�p����"
        .ApplyDataLabels
    End With

    chartObj.Activate
    ActiveChart.FullSeriesCollection(1).DataLabels.ShowPercentage = True
    ActiveChart.FullSeriesCollection(1).DataLabels.ShowValue = False


    With chartObj
        .Top = ActiveSheet.Cells(3, 15).Top
        .Left = ActiveSheet.Cells(3, 15).Left
        .Height = 300
        .Width = 400
    End With
   
    '**�W�v�{�^���쐬**
    With ActiveSheet.Buttons.add(Range("I10").Left, _
                                Range("I10").Top, _
                                Range("I10:J11").Width, _
                                Range("I10:J11").Height)
                            .OnAction = "MonthTotal"
                            .Characters.Text = "�W�v"
    End With

End Sub

'-------------------
'   �������͏���
'-------------------
Sub SumIncomeDisplay()
    '**�����ǉ������i�V�[�g�̒T���A��������**
    Dim i As Integer
    For i = 2 To Worksheets.Count
            Debug.Print i
            
            Dim month As String
            month = Replace(Mid(Sheets(i).Cells(1, 1).Text, 6), "��", "")
            
        '�V�[�g�̒T���A��������
        If Sumincome.TextBox2.Text = month Then
            Debug.Print i
            Dim sum As Long
            sum = Val(Replace(Sheets(i).Cells(6, 10).Value, "�~", ""))
            sum = sum + Val(Sumincome.TextBox1.Text)
            Debug.Print sum
            Sheets(i).Cells(6, 10) = sum & "�~"
            
        '**���W�v���b�Z�[�W�\��**
            Dim j As Long
            j = Val(Replace(Sheets(i).Cells(2, 10).Value, "�~", "")) + Val(Replace(Sheets(i).Cells(6, 10).Value, "�~", ""))
            If Sheets(i).Cells(7, 10).Value = "" Or j <> Val(Replace(Sheets(i).Cells(7, 10).Value, "�~", "")) Then
                With Sheets(i).Cells(1, 2)
                        .Value = "�W�v����Ă��܂���B"
                        .Font.Bold = True
                        .Font.Color = RGB(255, 0, 0)
                End With
            End If
                
        End If
    Next i
    

End Sub

'-------------------
'   �x�o���͏���
'-------------------
Sub newRecordAdd()
    '**�x�o�\�ւ̍s�ǉ��i�V�[�g�T���A���R�[�h���́A�����x���f�j**
    Dim i As Integer
    For i = 1 To Worksheets.Count
        Dim inYm As String
        inYm = Left(newRecord.TextBox2.Text, 4) & "�N" & Mid(newRecord.TextBox2.Text, 6, 2) & "��"
        Dim lastRow As Long
        '�V�[�g�T��
        If inYm = Sheets(i).Cells(1, 1).Text Then
            '���R�[�h����
            lastRow = Sheets(i).Cells(Rows.Count, 2).End(xlUp).Offset(1, 0).Row
            With Sheets(i)
                .Cells(lastRow, 2).Value = newRecord.TextBox2.Text
                .Cells(lastRow, 7).Value = newRecord.TextBox1.Text
                .Cells(lastRow, 4).Value = newRecord.ListBox1.Text
                .Cells(lastRow, 6).Value = newRecord.ListBox2.Text
                .Cells(lastRow, 5).Value = newRecord.TextBox3.Text
                .Cells(lastRow, 3).Value = WeekdayName(Weekday(newRecord.TextBox2.Text))
                .Columns("B:H").EntireColumn.AutoFit
                .Range(.Cells(lastRow, 2), .Cells(lastRow, 7)).Borders.LineStyle = xlContinuous
            End With
                       
            '�����x���f
            If Val(newRecord.ListBox2.Text) < 3 Then
                Dim j As Integer
                For j = 2 To 7
                    Sheets(i).Cells(lastRow, j).Interior.Color = RGB(255, 128, 128)
                Next j
            ElseIf Val(newRecord.ListBox2.Text) > 7 Then
                Dim k As Integer
                For k = 2 To 7
                    Sheets(i).Cells(lastRow, k).Interior.Color = RGB(255, 255, 153)
                Next k
            End If
            
        '**�W�������\�ւ̓���**
            Dim m As Integer
            Dim sum As Long
            For m = 3 To 16
                If newRecord.ListBox1.Text = Sheets(i).Cells(m, 12) Then
                    sum = Sheets(i).Cells(m, 13)
                    sum = sum + newRecord.TextBox1.Text
                    Sheets(i).Cells(m, 13).Value = sum
                End If
            Next m
        '**���W�v���b�Z�[�W�\��**
            Dim l As Long
            l = Val(Replace(Sheets(i).Cells(2, 10).Value, "�~", "")) + Val(Replace(Sheets(i).Cells(6, 10).Value, "�~", ""))
            If Sheets(i).Cells(4, 10).Value = "" Or l <> Val(Replace(Sheets(i).Cells(7, 10).Value, "�~", "")) Then
                With Sheets(i).Cells(1, 2)
                        .Value = "�W�v����Ă��܂���B"
                        .Font.Bold = True
                        .Font.Color = RGB(255, 0, 0)
                End With
            End If
            
        End If
    Next i

End Sub


'-------------------
'   ���ԏW�v����
'-------------------
Sub MonthTotal()
    Dim t As Long
    Dim lastRow As Long
    lastRow = ActiveSheet.Cells(Rows.Count, 7).End(xlUp).Row
    t = WorksheetFunction.sum(ActiveSheet.Range(Cells(3, 7), Cells(lastRow, 7)))
    ActiveSheet.Cells(2, 10).Value = t & "�~"
    
    Dim u As Long
    u = Left(ActiveSheet.Cells(3, 10).Value, Len(ActiveSheet.Cells(3, 10).Value) - 1)
    
    If t < u Then
    ActiveSheet.Cells(4, 10).Value = "�B��"
    Else
        With ActiveSheet.Cells(4, 10)
            .Value = "���B��"
            .Interior.Color = RGB(255, 128, 128)
        End With
    End If
    
    Dim b As Long
    b = Val(ActiveSheet.Cells(6, 10).Value) - Val(ActiveSheet.Cells(2, 10).Value)
    ActiveSheet.Cells(7, 10).Value = b & "�~"
    
    ActiveSheet.Cells(1, 2).Value = ""
    
End Sub


'-------------------
'   �N�ԏW�v����
'-------------------
Sub yearCreate()
    '**���ԃV�[�g���W�v�`�F�b�N**
    Dim s1 As Integer
    Dim ckTotal() As String
    ReDim ckTotal(0)
    ckTotal(0) = "�@���W�v�̌��F"
    
    For s1 = 2 To Worksheets.Count
        If Sheets(s1).Cells(1, 2).Value <> "" Then
            ReDim Preserve ckTotal(UBound(ckTotal) + 1)
            ckTotal(UBound(ckTotal)) = Sheets(s1).Cells(1, 1).Value
        End If
    Next s1
        
    If UBound(ckTotal) <> 0 Then
        Dim s2 As Integer
        Dim msg As String
        msg = ""
        For s2 = 0 To UBound(ckTotal)
            msg = msg & ckTotal(s2) & " "
        Next s2
        
        MsgBox "�W�v����Ă��Ȃ���������܂��B���Ԃ̏W�v���s���Ă��������B" & _
                vbCrLf & msg
        Exit Sub
    End If
    
'**�N�ԏW�v����**
    Worksheets.add After:=Worksheets(Worksheets.Count)
    ActiveSheet.Name = "�N�ԏW�v"
    Dim targetSheet As Worksheet
    Set targetSheet = ActiveSheet
    
'**�\�̍쐬**
    '���x���ڂ̕\�쐬
    Dim i As Integer
    targetSheet.Cells(2, 2).Value = "���x���ځi�~�j"
    For i = 2 To Sheets.Count - 1
        With targetSheet
            .Cells(3, i).Value = Worksheets(i).Cells(1, 1).Value
            .Cells(4, i).Value = Left(Worksheets(i).Cells(7, 10).Value, Len(Worksheets(i).Cells(7, 10).Value) - 1)
            .Range(targetSheet.Cells(3, i), targetSheet.Cells(4, i)).Borders.LineStyle = xlContinuous
            .Cells(3, i).Interior.Color = RGB(255, 220, 140)
            .Range("B3:M3").Columns.AutoFit
        End With
    Next i
    Debug.Print i
    
    '�N�Ԓ��~���ڂ̕\�쐬
    Dim j As Integer
    Dim yearSum As Long
    yearSum = 0
    targetSheet.Cells(6, 2).Value = "�N�Ԓ��~���ځi�~�j"
    For j = 2 To Sheets.Count - 1
        With targetSheet
            .Cells(7, j).Value = Worksheets(j).Cells(1, 1).Value
            yearSum = yearSum + Left(Worksheets(j).Cells(7, 10).Value, Len(Worksheets(j).Cells(7, 10).Value) - 1)
            .Cells(8, j).Value = yearSum
            .Range(targetSheet.Cells(7, j), targetSheet.Cells(8, j)).Borders.LineStyle = xlContinuous
            .Cells(7, j).Interior.Color = RGB(255, 220, 140)
        End With
    Next j
    Debug.Print j
    
    '�W�������ʏW�v�̕\�쐬
    targetSheet.Cells(24, 2).Value = "�W�������ʏW�v�i�~�j"
    With targetSheet
        .Cells(25, 2).Value = "�W������"
        .Cells(25, 3).Value = "���v���z"
        .Cells(26, 2).Value = "�H��"
        .Cells(27, 2).Value = "�O�H��"
        .Cells(28, 2).Value = "���M��"
        .Cells(29, 2).Value = "������"
        .Cells(30, 2).Value = "�ʐM��"
        .Cells(31, 2).Value = "���p�i"
        .Cells(32, 2).Value = "�ƒ�"
        .Cells(33, 2).Value = "�ߕ�"
        .Cells(34, 2).Value = "���e��"
        .Cells(35, 2).Value = "�"
        .Cells(36, 2).Value = "��ʔ�"
        .Cells(37, 2).Value = "���۔�"
        .Cells(38, 2).Value = "���ʔ�"
        .Cells(39, 2).Value = "�o��"
        .Range("B25:C39").BorderAround LineStyle:=xlContinuous
        .Range("B25:C39").Borders(xlInsideVertical).LineStyle = xlContinuous
        .Range("B25:C39").Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Range("B25:C25").Interior.Color = RGB(255, 220, 140)
    End With
    
    Dim k As Integer
    Dim l As Integer
    Dim sumAll(13) As Long
    Dim add As Long
    
    For k = 0 To 13
        sumAll(k) = 0
    Next k

    For k = 0 To 13
        For l = 2 To Sheets.Count - 1
            add = Replace(Sheets(l).Cells(k + 3, 13).Value, "�~", "")
            sumAll(k) = sumAll(k) + add
        Next l
        targetSheet.Cells(k + 26, 3).Value = sumAll(k)
    Next k
    
'**�O���t�̍쐬**
    '���x���ڂ̖_�O���t�쐬
    Dim chartObj1 As ChartObject
    Set chartObj1 = targetSheet.ChartObjects.add(Left:=100, Top:=100, Width:=400, Height:=200)

    With chartObj1.chart
        .ChartType = xlColumnClustered
        .SetSourceData Source:=targetSheet.Range(targetSheet.Cells(3, 2), targetSheet.Cells(4, i - 1))
        .HasTitle = True
        .ChartTitle.Text = "���x����"
        .Legend.Delete
        .Axes(xlValue, 1).HasTitle = True
        .Axes(xlValue, 1).AxisTitle.Text = "���z(�~)"
        .Axes(xlValue, 1).AxisTitle.Font.Size = 9
        .Axes(xlValue, 1).AxisTitle.Font.Bold = False
    End With

    With chartObj1
        .Top = targetSheet.Cells(11, 2).Top
        .Left = targetSheet.Cells(11, 2).Left
        .Height = 200
        .Width = 400
    End With
    
    '�N�Ԓ��~���ڂ̖_�O���t�쐬
    Dim chartObj2 As ChartObject
    Set chartObj2 = targetSheet.ChartObjects.add(Left:=100, Top:=100, Width:=400, Height:=200)

    With chartObj2.chart
        .ChartType = xlColumnClustered
        .SetSourceData Source:=targetSheet.Range(targetSheet.Cells(7, 2), targetSheet.Cells(8, j - 1))
        .HasTitle = True
        .ChartTitle.Text = "�N�Ԓ��~����"
        .Legend.Delete
        .Axes(xlValue, 1).HasTitle = True
        .Axes(xlValue, 1).AxisTitle.Text = "���z(�~)"
        .Axes(xlValue, 1).AxisTitle.Font.Size = 9
        .Axes(xlValue, 1).AxisTitle.Font.Bold = False
    End With

    With chartObj2
        .Top = targetSheet.Cells(11, 11).Top
        .Left = targetSheet.Cells(11, 11).Left
        .Height = 200
        .Width = 400
    End With
    
    '�W�������ʎg�p�����̉~�O���t�쐬
    Dim chartObj3 As ChartObject
    Set chartObj3 = targetSheet.ChartObjects.add(Left:=100, Top:=100, Width:=400, Height:=300)
       
    With chartObj3.chart
        .ChartType = xlPie
        .SetSourceData Source:=targetSheet.Range(targetSheet.Cells(26, 2), targetSheet.Cells(39, 3))
        .HasTitle = True
        .ChartTitle.Text = "�W�������ʎg�p����"
        .ApplyDataLabels
    End With
    
    chartObj3.Activate
    ActiveChart.FullSeriesCollection(1).DataLabels.ShowPercentage = True
    ActiveChart.FullSeriesCollection(1).DataLabels.ShowValue = False
    
    With chartObj3
        .Top = targetSheet.Cells(24, 6).Top
        .Left = targetSheet.Cells(24, 6).Left
        .Height = 300
        .Width = 400
    End With

End Sub

'-------------------
'       bk����
'-------------------
Sub bk()
    '�V�K�u�b�N�쐬
    Dim newBook As Workbook
    Set newBook = Workbooks.add
    Dim i As Integer
    Dim j As Integer
    j = 1
    
    '2���ڈȍ~�̃R�s�[����
    For i = 2 To ThisWorkbook.Sheets.Count
        ThisWorkbook.Worksheets(i).Copy before:=newBook.Sheets(j)
        j = j + 1
    Next i
    
    '�ۑ��u�b�N���ݒ�
    Dim pathName As String
    pathName = ThisWorkbook.Path & "\�ƌv��bk" & Format(Now, "Long Date") & ".xls"
    
    Dim res As Integer
    If Dir(pathName) <> "" Then
        res = MsgBox("�������O�̃t�@�C�������݂��܂��B" & vbCrLf & "�㏑�����܂����H", vbYesNo)
        If res = 7 Then
            MsgBox "�ۑ�����܂���ł����B" & vbCrLf & "���O�����ĕۑ����Ă��������B"
        Else
            Application.DisplayAlerts = False
            newBook.SaveAs Filename:=pathName
            Application.DisplayAlerts = True
            MsgBox pathName & "�ŕۑ�����܂����B"
        End If
    End If
        
End Sub


'-------------------
'   ����������
'-------------------
Sub initial()
    '�������i1���ڈȊO�̍폜�j
    Application.DisplayAlerts = False
    
    While Sheets.Count > 1
        Worksheets(2).Delete
    Wend
    
    Application.DisplayAlerts = True
    
    MsgBox "���������܂����B"
    
End Sub
