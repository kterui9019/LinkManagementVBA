VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RegisterLink 
   Caption         =   "�����N�V�K�o�^"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   8910
   OleObjectBlob   =   "RegisterLink.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "RegisterLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnEnd_Click()
    Unload Me
End Sub

Private Sub btnRegister_Click()
    
    Dim Msg As String
    Dim title As String
    Dim res As Integer
    Dim TargetRow As Integer
    
    If txtTitle = "" Then
        MsgBox "�^�C�g���͕K�{���͍��ڂł�", vbExclamation, "�G���["
        txtTitle.SetFocus
        Exit Sub
    End If
    
    If txtUrl = "" Then
        MsgBox "URL�͕K�{���͍��ڂł�", vbExclamation, "�G���["
        txtUrl.SetFocus
        Exit Sub
    End If
    
    Msg = "�����N��o�^���܂��B��낵���ł����H"
    title = "�����N�o�^�̊m�F"
    res = MsgBox(Msg, vbYesNo + vbQuestion + vbDefaultButton2, title)
    If res = vbNo Then Exit Sub
    
    With Worksheets("sheet2")
        TargetRow = .Range("B65536").End(xlUp).Offset(1).Row
        .Range("B" & TargetRow).Value = txtTitle.Text
        .Range("C" & TargetRow).Value = txtUrl.Text
        .Range("D" & TargetRow).Value = txtOption.Text
    End With
    
    MsgBox "�o�^���������܂���", vbInfomation, "�o�^����"

    Msg = "�����ēo�^���܂����H"
    title = "�m�F"
    res = MsgBox(Msg, vbYesNo + vbQuestion + vbDefaultButton2, title)
    If res = vbNo Then Unload Me
    
    txtTitle.Value = ""
    txtUrl.Value = ""
    txtOption.Value = ""
    
End Sub
