VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} �ƒ�w�K�J�[�h_���ݒ�Form 
   Caption         =   "�N�x���ؑ�"
   ClientHeight    =   3012
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   3696
   OleObjectBlob   =   "�ƒ�w�K�J�[�h_���ݒ�Form.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "�ƒ�w�K�J�[�h_���ݒ�Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
    Range(Cells(7, 1), Cells(37, 2)).ClearContents
    
    Call Sheet2�̌��ύX
    
    Unload �ƒ�w�K�J�[�h_���ݒ�Form
End Sub

Private Sub UserForm_Click()
    ' ��̃N���b�N�C�x���g�n���h��
End Sub
