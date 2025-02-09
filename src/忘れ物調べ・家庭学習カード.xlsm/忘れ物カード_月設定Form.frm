VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 忘れ物カード_月設定Form 
   Caption         =   "年度及び月の切替"
   ClientHeight    =   3168
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   3672
   OleObjectBlob   =   "忘れ物カード_月設定Form.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "忘れ物カード_月設定Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CommandButton1_Click()
    Rows("4:35").Hidden = False
    Range(Cells(4, 1), Cells(35, 2)).ClearContents
    
    Call Sheet1の月変更
    
    Unload 忘れ物カード_月設定Form
End Sub

Private Sub UserForm_Click()
    ' 空のクリックイベントハンドラ
End Sub
