Attribute VB_Name = "�Y�ꕨ�J�[�h_���ݒ�"
Option Explicit

Sub Sheet1�̌��ύX()
  
  Dim i As Long
  Dim youbi As Date
  Dim eo As Long
  
  Application.ScreenUpdating = False
  
  Cells(1, 1).Value = �Y�ꕨ�J�[�h_���ݒ�Form.TextBox2.Value
  
  ' ���������擾
  eo = Day(WorksheetFunction.EoMonth(DateSerial(�Y�ꕨ�J�[�h_���ݒ�Form.TextBox1.Value, �Y�ꕨ�J�[�h_���ݒ�Form.TextBox2.Value, 1), 0))
  Cells(4, 1).Value = 1
  Cells(4, 1).Resize(eo).DataSeries step:=1 ' A4�Z������1�`�������̒l������́B
  
  ' ���t���Ƃɗj����ݒ肵�A�y�����\��
  For i = 4 To Cells(Rows.Count, 1).End(xlUp).Row
    youbi = DateSerial(�Y�ꕨ�J�[�h_���ݒ�Form.TextBox1.Value, �Y�ꕨ�J�[�h_���ݒ�Form.TextBox2.Value, Cells(i, 1).Value)
  
        Cells(i, 2).Value = Format(Weekday(youbi), "aaa")
    
    With Cells(i, 2)
        If (.Value = "�y" Or .Value = "��") Then
            Rows(i).Hidden = True
        End If
    End With

  Next i
    
End Sub
