Attribute VB_Name = "Module2"
Option Explicit
'*********************************
'            メイン処理
'*********************************

'-------------------
'月間シート作成処理
'-------------------
Sub NewMonthCreate()
    '**シート名設定**
    Dim year As String
    Dim month As String
    
    year = NewMonth.TextBox3.Text
    month = NewMonth.TextBox1.Text

    Worksheets.add After:=Worksheets(Worksheets.Count)
    ActiveSheet.Name = year & "年" & month & "月"
    
    With Cells(1, 1)
        .NumberFormat = "@"
        .Value = year & "年" & month & "月"
    End With

    '**支出表作成**
    Cells(2, 2) = "日付"
    Cells(2, 3) = "曜日"
    Cells(2, 4) = "ジャンル"
    Cells(2, 5) = "内容"
    Cells(2, 6) = "満足度"
    Cells(2, 7) = "金額"

    Range("B2:G2").BorderAround LineStyle:=xlContinuous
    Range("B2:G2").Borders(xlInsideVertical).LineStyle = xlContinuous
    Range("B2:G2").Borders(xlInsideHorizontal).LineStyle = xlContinuous

    Range("B2:G2").Interior.Color = RGB(255, 220, 140)

    '**合計支出表作成**
    Cells(2, 9) = "TOTAL"
    Cells(3, 9) = "予算"
    Cells(4, 9) = "予算達成"

    Cells(3, 10) = NewMonth.TextBox2.Text & "円"

    Range("I2:J4").BorderAround LineStyle:=xlContinuous
    Range("I2:J4").Borders(xlInsideVertical).LineStyle = xlContinuous
    Range("I2:J4").Borders(xlInsideHorizontal).LineStyle = xlContinuous

    Range("I2:I4").Interior.Color = RGB(255, 220, 140)
    
    '**収支表作成**
    Cells(6, 9) = "収入合計"
    Cells(7, 9) = "収支"

    Cells(2, 10) = 0 & "円"
    Cells(6, 10) = 0 & "円"
    Cells(7, 10) = 0 & "円"

    Range("I6:J7").BorderAround LineStyle:=xlContinuous
    Range("I6:J7").Borders(xlInsideVertical).LineStyle = xlContinuous
    Range("I6:J7").Borders(xlInsideHorizontal).LineStyle = xlContinuous

    Range("I6:I7").Interior.Color = RGB(255, 220, 140)

    '**ジャンル表作成**
    Cells(2, 12) = "ジャンル"
    Cells(2, 13) = "合計金額"
    Cells(3, 12) = "食費"
    Cells(4, 12) = "外食費"
    Cells(5, 12) = "光熱費"
    Cells(6, 12) = "水道代"
    Cells(7, 12) = "通信費"
    Cells(8, 12) = "日用品"
    Cells(9, 12) = "家賃"
    Cells(10, 12) = "衣服"
    Cells(11, 12) = "美容代"
    Cells(12, 12) = "趣味"
    Cells(13, 12) = "交通費"
    Cells(14, 12) = "交際費"
    Cells(15, 12) = "特別費"
    Cells(16, 12) = "経費"
    
    Dim i As Integer
    For i = 3 To 16
        Cells(i, 13) = 0
    Next i
    
    Range("L2:M16").BorderAround LineStyle:=xlContinuous
    Range("L2:M16").Borders(xlInsideVertical).LineStyle = xlContinuous
    Range("L2:M16").Borders(xlInsideHorizontal).LineStyle = xlContinuous

    Range("L2:M2").Interior.Color = RGB(255, 220, 140)
    
    '**円グラフ作成**
    Dim chartObj As ChartObject
    Set chartObj = ActiveSheet.ChartObjects.add(Left:=100, Top:=100, Width:=400, Height:=300)

    With chartObj.chart
        .ChartType = xlPie
        .SetSourceData Source:=ActiveSheet.Range(ActiveSheet.Cells(2, 12), ActiveSheet.Cells(16, 13))
        .HasTitle = True
        .ChartTitle.Text = "ジャンル別使用割合"
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
   
    '**集計ボタン作成**
    With ActiveSheet.Buttons.add(Range("I10").Left, _
                                Range("I10").Top, _
                                Range("I10:J11").Width, _
                                Range("I10:J11").Height)
                            .OnAction = "MonthTotal"
                            .Characters.Text = "集計"
    End With

End Sub

'-------------------
'   収入入力処理
'-------------------
Sub SumIncomeDisplay()
    '**収入追加処理（シートの探索、収入入力**
    Dim i As Integer
    For i = 2 To Worksheets.Count
            Debug.Print i
            
            Dim month As String
            month = Replace(Mid(Sheets(i).Cells(1, 1).Text, 6), "月", "")
            
        'シートの探索、収入入力
        If Sumincome.TextBox2.Text = month Then
            Debug.Print i
            Dim sum As Long
            sum = Val(Replace(Sheets(i).Cells(6, 10).Value, "円", ""))
            sum = sum + Val(Sumincome.TextBox1.Text)
            Debug.Print sum
            Sheets(i).Cells(6, 10) = sum & "円"
            
        '**未集計メッセージ表示**
            Dim j As Long
            j = Val(Replace(Sheets(i).Cells(2, 10).Value, "円", "")) + Val(Replace(Sheets(i).Cells(6, 10).Value, "円", ""))
            If Sheets(i).Cells(7, 10).Value = "" Or j <> Val(Replace(Sheets(i).Cells(7, 10).Value, "円", "")) Then
                With Sheets(i).Cells(1, 2)
                        .Value = "集計されていません。"
                        .Font.Bold = True
                        .Font.Color = RGB(255, 0, 0)
                End With
            End If
                
        End If
    Next i
    

End Sub

'-------------------
'   支出入力処理
'-------------------
Sub newRecordAdd()
    '**支出表への行追加（シート探索、レコード入力、満足度反映）**
    Dim i As Integer
    For i = 1 To Worksheets.Count
        Dim inYm As String
        inYm = Left(newRecord.TextBox2.Text, 4) & "年" & Mid(newRecord.TextBox2.Text, 6, 2) & "月"
        Dim lastRow As Long
        'シート探索
        If inYm = Sheets(i).Cells(1, 1).Text Then
            'レコード入力
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
                       
            '満足度反映
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
            
        '**ジャンル表への入力**
            Dim m As Integer
            Dim sum As Long
            For m = 3 To 16
                If newRecord.ListBox1.Text = Sheets(i).Cells(m, 12) Then
                    sum = Sheets(i).Cells(m, 13)
                    sum = sum + newRecord.TextBox1.Text
                    Sheets(i).Cells(m, 13).Value = sum
                End If
            Next m
        '**未集計メッセージ表示**
            Dim l As Long
            l = Val(Replace(Sheets(i).Cells(2, 10).Value, "円", "")) + Val(Replace(Sheets(i).Cells(6, 10).Value, "円", ""))
            If Sheets(i).Cells(4, 10).Value = "" Or l <> Val(Replace(Sheets(i).Cells(7, 10).Value, "円", "")) Then
                With Sheets(i).Cells(1, 2)
                        .Value = "集計されていません。"
                        .Font.Bold = True
                        .Font.Color = RGB(255, 0, 0)
                End With
            End If
            
        End If
    Next i

End Sub


'-------------------
'   月間集計処理
'-------------------
Sub MonthTotal()
    Dim t As Long
    Dim lastRow As Long
    lastRow = ActiveSheet.Cells(Rows.Count, 7).End(xlUp).Row
    t = WorksheetFunction.sum(ActiveSheet.Range(Cells(3, 7), Cells(lastRow, 7)))
    ActiveSheet.Cells(2, 10).Value = t & "円"
    
    Dim u As Long
    u = Left(ActiveSheet.Cells(3, 10).Value, Len(ActiveSheet.Cells(3, 10).Value) - 1)
    
    If t < u Then
    ActiveSheet.Cells(4, 10).Value = "達成"
    Else
        With ActiveSheet.Cells(4, 10)
            .Value = "未達成"
            .Interior.Color = RGB(255, 128, 128)
        End With
    End If
    
    Dim b As Long
    b = Val(ActiveSheet.Cells(6, 10).Value) - Val(ActiveSheet.Cells(2, 10).Value)
    ActiveSheet.Cells(7, 10).Value = b & "円"
    
    ActiveSheet.Cells(1, 2).Value = ""
    
End Sub


'-------------------
'   年間集計処理
'-------------------
Sub yearCreate()
    '**月間シート未集計チェック**
    Dim s1 As Integer
    Dim ckTotal() As String
    ReDim ckTotal(0)
    ckTotal(0) = "　未集計の月："
    
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
        
        MsgBox "集計されていない月があります。月間の集計を行ってください。" & _
                vbCrLf & msg
        Exit Sub
    End If
    
'**年間集計処理**
    Worksheets.add After:=Worksheets(Worksheets.Count)
    ActiveSheet.Name = "年間集計"
    Dim targetSheet As Worksheet
    Set targetSheet = ActiveSheet
    
'**表の作成**
    '収支推移の表作成
    Dim i As Integer
    targetSheet.Cells(2, 2).Value = "収支推移（円）"
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
    
    '年間貯蓄推移の表作成
    Dim j As Integer
    Dim yearSum As Long
    yearSum = 0
    targetSheet.Cells(6, 2).Value = "年間貯蓄推移（円）"
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
    
    'ジャンル別集計の表作成
    targetSheet.Cells(24, 2).Value = "ジャンル別集計（円）"
    With targetSheet
        .Cells(25, 2).Value = "ジャンル"
        .Cells(25, 3).Value = "合計金額"
        .Cells(26, 2).Value = "食費"
        .Cells(27, 2).Value = "外食費"
        .Cells(28, 2).Value = "光熱費"
        .Cells(29, 2).Value = "水道代"
        .Cells(30, 2).Value = "通信費"
        .Cells(31, 2).Value = "日用品"
        .Cells(32, 2).Value = "家賃"
        .Cells(33, 2).Value = "衣服"
        .Cells(34, 2).Value = "美容代"
        .Cells(35, 2).Value = "趣味"
        .Cells(36, 2).Value = "交通費"
        .Cells(37, 2).Value = "交際費"
        .Cells(38, 2).Value = "特別費"
        .Cells(39, 2).Value = "経費"
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
            add = Replace(Sheets(l).Cells(k + 3, 13).Value, "円", "")
            sumAll(k) = sumAll(k) + add
        Next l
        targetSheet.Cells(k + 26, 3).Value = sumAll(k)
    Next k
    
'**グラフの作成**
    '収支推移の棒グラフ作成
    Dim chartObj1 As ChartObject
    Set chartObj1 = targetSheet.ChartObjects.add(Left:=100, Top:=100, Width:=400, Height:=200)

    With chartObj1.chart
        .ChartType = xlColumnClustered
        .SetSourceData Source:=targetSheet.Range(targetSheet.Cells(3, 2), targetSheet.Cells(4, i - 1))
        .HasTitle = True
        .ChartTitle.Text = "収支推移"
        .Legend.Delete
        .Axes(xlValue, 1).HasTitle = True
        .Axes(xlValue, 1).AxisTitle.Text = "金額(円)"
        .Axes(xlValue, 1).AxisTitle.Font.Size = 9
        .Axes(xlValue, 1).AxisTitle.Font.Bold = False
    End With

    With chartObj1
        .Top = targetSheet.Cells(11, 2).Top
        .Left = targetSheet.Cells(11, 2).Left
        .Height = 200
        .Width = 400
    End With
    
    '年間貯蓄推移の棒グラフ作成
    Dim chartObj2 As ChartObject
    Set chartObj2 = targetSheet.ChartObjects.add(Left:=100, Top:=100, Width:=400, Height:=200)

    With chartObj2.chart
        .ChartType = xlColumnClustered
        .SetSourceData Source:=targetSheet.Range(targetSheet.Cells(7, 2), targetSheet.Cells(8, j - 1))
        .HasTitle = True
        .ChartTitle.Text = "年間貯蓄推移"
        .Legend.Delete
        .Axes(xlValue, 1).HasTitle = True
        .Axes(xlValue, 1).AxisTitle.Text = "金額(円)"
        .Axes(xlValue, 1).AxisTitle.Font.Size = 9
        .Axes(xlValue, 1).AxisTitle.Font.Bold = False
    End With

    With chartObj2
        .Top = targetSheet.Cells(11, 11).Top
        .Left = targetSheet.Cells(11, 11).Left
        .Height = 200
        .Width = 400
    End With
    
    'ジャンル別使用割合の円グラフ作成
    Dim chartObj3 As ChartObject
    Set chartObj3 = targetSheet.ChartObjects.add(Left:=100, Top:=100, Width:=400, Height:=300)
       
    With chartObj3.chart
        .ChartType = xlPie
        .SetSourceData Source:=targetSheet.Range(targetSheet.Cells(26, 2), targetSheet.Cells(39, 3))
        .HasTitle = True
        .ChartTitle.Text = "ジャンル別使用割合"
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
'       bk処理
'-------------------
Sub bk()
    '新規ブック作成
    Dim newBook As Workbook
    Set newBook = Workbooks.add
    Dim i As Integer
    Dim j As Integer
    j = 1
    
    '2枚目以降のコピー処理
    For i = 2 To ThisWorkbook.Sheets.Count
        ThisWorkbook.Worksheets(i).Copy before:=newBook.Sheets(j)
        j = j + 1
    Next i
    
    '保存ブック名設定
    Dim pathName As String
    pathName = ThisWorkbook.Path & "\家計簿bk" & Format(Now, "Long Date") & ".xls"
    
    Dim res As Integer
    If Dir(pathName) <> "" Then
        res = MsgBox("同じ名前のファイルが存在します。" & vbCrLf & "上書きしますか？", vbYesNo)
        If res = 7 Then
            MsgBox "保存されませんでした。" & vbCrLf & "名前をつけて保存してください。"
        Else
            Application.DisplayAlerts = False
            newBook.SaveAs Filename:=pathName
            Application.DisplayAlerts = True
            MsgBox pathName & "で保存されました。"
        End If
    End If
        
End Sub


'-------------------
'   初期化処理
'-------------------
Sub initial()
    '初期化（1枚目以外の削除）
    Application.DisplayAlerts = False
    
    While Sheets.Count > 1
        Worksheets(2).Delete
    Wend
    
    Application.DisplayAlerts = True
    
    MsgBox "初期化しました。"
    
End Sub
