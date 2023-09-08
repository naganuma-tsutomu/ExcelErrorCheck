VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} OutputForm
   Caption         =   "エラー"
   ClientHeight    =   6960
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9750
   OleObjectBlob   =   "OutputForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
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

' フォームを開いたとき
Private Sub UserForm_Initialize()
    ' フォームの表示位置を中央にする
    Me.StartUpPosition = 0
    Me.Top = Application.Top + ((Application.Height - Me.Height) / 2)
    Me.Left = Application.Left + ((Application.Width - Me.Width) / 2)
End Sub

