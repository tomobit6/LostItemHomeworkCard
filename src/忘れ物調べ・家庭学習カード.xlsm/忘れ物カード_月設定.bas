Attribute VB_Name = "忘れ物カード_月設定"
Option Explicit

Sub Sheet1の月変更()
  
  Dim i As Long
  Dim youbi As Date
  Dim eo As Long
  
  Application.ScreenUpdating = False
  
  Cells(1, 1).Value = 忘れ物カード_月設定Form.TextBox2.Value
  
  ' 月末日を取得
  eo = Day(WorksheetFunction.EoMonth(DateSerial(忘れ物カード_月設定Form.TextBox1.Value, 忘れ物カード_月設定Form.TextBox2.Value, 1), 0))
  Cells(4, 1).Value = 1
  Cells(4, 1).Resize(eo).DataSeries step:=1 ' A4セルから1～月末日の値をを入力。
  
  ' 日付ごとに曜日を設定し、土日を非表示
  For i = 4 To Cells(Rows.Count, 1).End(xlUp).Row
    youbi = DateSerial(忘れ物カード_月設定Form.TextBox1.Value, 忘れ物カード_月設定Form.TextBox2.Value, Cells(i, 1).Value)
  
        Cells(i, 2).Value = Format(Weekday(youbi), "aaa")
    
    With Cells(i, 2)
        If (.Value = "土" Or .Value = "日") Then
            Rows(i).Hidden = True
        End If
    End With

  Next i
    
End Sub
