VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} �Y�ꕨ�J�[�h_���ݒ�Form 
   Caption         =   "�N�x�y�ь��̐ؑ�"
   ClientHeight    =   3168
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   3672
   OleObjectBlob   =   "�Y�ꕨ�J�[�h_���ݒ�Form.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "�Y�ꕨ�J�[�h_���ݒ�Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CommandButton1_Click()
    Rows("4:35").Hidden = False
    Range(Cells(4, 1), Cells(35, 2)).ClearContents
    
    Call Sheet1�̌��ύX
    
    Unload �Y�ꕨ�J�[�h_���ݒ�Form
End Sub

Private Sub UserForm_Click()
    ' ��̃N���b�N�C�x���g�n���h��
End Sub
