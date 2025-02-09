Attribute VB_Name = "家庭学習カード_月設定"
Option Explicit


Sub Sheet2の月変更()
  
  Dim i As Long
  Dim youbi As Date
  Dim eo As Long
  
  
  Application.ScreenUpdating = False
  
  Cells(1, 20).Value = 家庭学習カード_月設定Form.TextBox2.Value
    ' 月末日を取得
    eo = Day(WorksheetFunction.EoMonth(DateSerial(家庭学習カード_月設定Form.TextBox1.Value, 家庭学習カード_月設定Form.TextBox2.Value, 1), 0))
    Cells(7, 1).Value = 1
    Cells(7, 1).Resize(eo).DataSeries step:=1 ' A7セルから1～月末日の値をを入力。
    
    ' 日付ごとに曜日をB列に設定
    For i = 7 To Cells(Rows.Count, 1).End(xlUp).Row
        youbi = DateSerial(家庭学習カード_月設定Form.TextBox1.Value, 家庭学習カード_月設定Form.TextBox2.Value, Cells(i, 1).Value)
  
            Cells(i, 2).Value = Format(Weekday(youbi), "aaa")
    Next i
     
End Sub
