VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "������"
   ClientHeight    =   1572
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   1650
   OleObjectBlob   =   "UserForm1.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
    If MsgBox("���~���܂����H", 292) = vbYes Then flag = True
End Sub

Private Sub UserForm_Click()

End Sub
