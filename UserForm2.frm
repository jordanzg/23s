VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "���ʏ���"
   ClientHeight    =   2304
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3375
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton6_Click()
    'PW�Ȃ��ł��@csv(sjis)�ł�
    fmt = "csvsjis"
    Unload Me
End Sub
Private Sub UserForm_Initialize()
'    MsgBox "�₟"
    TextBox1.Value = passwordGet(10)
End Sub
Private Sub CommandButton1_Click()
    'PW����ł� xlsx�ł�
    'MsgBox TextBox1.Value
    dw = TextBox1.Value
    fmt = "xlsx"
    Unload Me
End Sub
Private Sub CommandButton2_Click()
    'PW�Ȃ��ł��@xlsx�ł�
    'MsgBox TextBox1.Value
    fmt = "xlsx"
    Unload Me
End Sub
Private Sub CommandButton3_Click()
    'PW����ł� xlsb�ł�
    'MsgBox TextBox1.Value
    dw = TextBox1.Value
    fmt = "xlsb"
    Unload Me
End Sub
Private Sub CommandButton4_Click()
    'PW�Ȃ��ł��@xlsb�ł�
    'MsgBox TextBox1.Value
    fmt = "xlsb"
    Unload Me
End Sub
Private Sub CommandButton5_Click()
    'PW�Ȃ��ł��@csv(utf-8)�ł�
    fmt = "csv"
    Unload Me
End Sub
