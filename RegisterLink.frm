VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RegisterLink 
   Caption         =   "リンク新規登録"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   8910
   OleObjectBlob   =   "RegisterLink.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
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
        MsgBox "タイトルは必須入力項目です", vbExclamation, "エラー"
        txtTitle.SetFocus
        Exit Sub
    End If
    
    If txtUrl = "" Then
        MsgBox "URLは必須入力項目です", vbExclamation, "エラー"
        txtUrl.SetFocus
        Exit Sub
    End If
    
    Msg = "リンクを登録します。よろしいですか？"
    title = "リンク登録の確認"
    res = MsgBox(Msg, vbYesNo + vbQuestion + vbDefaultButton2, title)
    If res = vbNo Then Exit Sub
    
    With Worksheets("sheet2")
        TargetRow = .Range("B65536").End(xlUp).Offset(1).Row
        .Range("B" & TargetRow).Value = txtTitle.Text
        .Range("C" & TargetRow).Value = txtUrl.Text
        .Range("D" & TargetRow).Value = txtOption.Text
    End With
    
    MsgBox "登録が完了しました", vbInfomation, "登録完了"

    Msg = "続けて登録しますか？"
    title = "確認"
    res = MsgBox(Msg, vbYesNo + vbQuestion + vbDefaultButton2, title)
    If res = vbNo Then Unload Me
    
    txtTitle.Value = ""
    txtUrl.Value = ""
    txtOption.Value = ""
    
End Sub
