VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} �Y�ꕨ�J�[�h_���Ɠ��x���ݒ�Form 
   Caption         =   "���Ɠ��ݒ�@�x���E��x���ݒ�"
   ClientHeight    =   9504
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   8496
   OleObjectBlob   =   "�Y�ꕨ�J�[�h_���Ɠ��x���ݒ�Form.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "�Y�ꕨ�J�[�h_���Ɠ��x���ݒ�Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CommandButton1_Click() 'OK�{�^��
    Dim i As Long
    
    ' CheckBox��31���傫���ԍ����E�̋x���y�ё�x�ݒ�B
    ' CheckBox��31�ȉ��̔ԍ��̏ꍇ�����̏T�x���̎��Ɠ��ݒ�B
    For i = 1 To 62
        If Me.Controls("CheckBox" & i).Value = True Then
            If i > 31 Then
                Rows(i - 28).Hidden = True
            Else
                Rows(i + 3).Hidden = False
            End If
        End If
    Next
    
    Unload �Y�ꕨ�J�[�h_���Ɠ��x���ݒ�Form
End Sub

Private Sub CommandButton2_Click() '�L�����Z���{�^��
    Unload �Y�ꕨ�J�[�h_���Ɠ��x���ݒ�Form
End Sub

Private Sub UserForm_Click()
    ' ��̃N���b�N�C�x���g�n���h��
End Sub
