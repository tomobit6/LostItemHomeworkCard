VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 家庭学習カード_月設定Form 
   Caption         =   "年度月切替"
   ClientHeight    =   3012
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   3696
   OleObjectBlob   =   "家庭学習カード_月設定Form.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "家庭学習カード_月設定Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
    Range(Cells(7, 1), Cells(37, 2)).ClearContents
    
    Call Sheet2の月変更
    
    Unload 家庭学習カード_月設定Form
End Sub

Private Sub UserForm_Click()
    ' 空のクリックイベントハンドラ
End Sub
