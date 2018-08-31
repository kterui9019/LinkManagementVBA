VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TopForm 
   Caption         =   "UserForm1"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   4710
   OleObjectBlob   =   "TopForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "TopForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnClose_Click()
'    Application.Quit
'    Windows("リンク集").Close True
    Unload Me
    Application.Quit
    Windows("リンク集.xlsm").Close True
End Sub

Private Sub btnRegisterForm_Click()
    RegisterLink.Show
End Sub

Private Sub btnSelectForm_Click()
    SelectForm.Show
End Sub
