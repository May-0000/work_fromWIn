Attribute VB_Name = "Module0"
Option Explicit
'*****************************
'   各ボタン押下直後の処理
'*****************************

'『支出の入力』ボタン押下
Sub button1()
    newRecord.Show
End Sub

'『収入の入力』ボタン押下
Sub button2()
    Sumincome.Show
End Sub

'『月間シート作成』ボタン押下
Sub button3()
    NewMonth.Show
End Sub

'『年間集計』ボタン押下
Sub button4()
    Call yearTotal
End Sub

'『バックアップ』ボタン押下
Sub buttonbk()
    Call BkUp
End Sub

'『初期化』ボタン押下
Sub buttonIni()
    Call Initialization
End Sub
