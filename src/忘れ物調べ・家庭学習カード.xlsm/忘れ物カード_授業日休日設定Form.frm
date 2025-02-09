VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 忘れ物カード_授業日休日設定Form 
   Caption         =   "授業日設定　休日・代休日設定"
   ClientHeight    =   9504
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   8496
   OleObjectBlob   =   "忘れ物カード_授業日休日設定Form.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "忘れ物カード_授業日休日設定Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CommandButton1_Click() 'OKボタン
    Dim i As Long
    
    ' CheckBoxが31より大きい番号→右の休日及び代休設定。
    ' CheckBoxが31以下の番号の場合→左の週休日の授業日設定。
    For i = 1 To 62
        If Me.Controls("CheckBox" & i).Value = True Then
            If i > 31 Then
                Rows(i - 28).Hidden = True
            Else
                Rows(i + 3).Hidden = False
            End If
        End If
    Next
    
    Unload 忘れ物カード_授業日休日設定Form
End Sub

Private Sub CommandButton2_Click() 'キャンセルボタン
    Unload 忘れ物カード_授業日休日設定Form
End Sub

Private Sub UserForm_Click()
    ' 空のクリックイベントハンドラ
End Sub
