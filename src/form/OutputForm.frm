VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} OutputForm
   Caption         =   "�G���["
   ClientHeight    =   6960
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9750
   OleObjectBlob   =   "OutputForm.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "OutputForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
    Unload Me
End Sub

' �t�H�[�����J�����Ƃ�
Private Sub UserForm_Initialize()
    ' �t�H�[���̕\���ʒu�𒆉��ɂ���
    Me.StartUpPosition = 0
    Me.Top = Application.Top + ((Application.Height - Me.Height) / 2)
    Me.Left = Application.Left + ((Application.Width - Me.Width) / 2)
End Sub

