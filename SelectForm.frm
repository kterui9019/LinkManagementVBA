VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SelectForm 
   Caption         =   "UserForm2"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   8910
   OleObjectBlob   =   "SelectForm.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "SelectForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const titleColumn As Integer = 2
Const urlColumn As Integer = 3
Dim lastRow As Integer

'�}�N���N�����̏���
Private Sub UserForm_Initialize()
    
    lastRow = Worksheets("Sheet2").Range("B65536").End(xlUp).Row '�V�[�g2��B65536�Z�������֌������Ēl�����͂���Ă���Z����T����lastRow�ɓ����
    
    '�R���{�{�b�N�X�ɒl�����鏈��
    '�N���A
    cboSelect.Clear
    
    '2��ڂ�2�s�ڂ���ŏI�s�܂Œl���R���{�{�b�N�X�ɓ����
    For startRow = 2 To lastRow
       cboSelect.AddItem Worksheets("Sheet2").Cells(startRow, 2).Value
    Next
    '1�Ԗڂ̃��X�g�̒l��I����Ԃɂ���
    cboSelect.Value = cboSelect.List(0)
    
End Sub

Private Sub btnJump_Click()
    'Dim selectedCombobox As String  '�I�����ꂽ�R���{�{�b�N�X�̒l
    Dim linkurl As String           '�����N���URL/PATH
    Dim linkRow As Integer          '�����N��̂���s
    
    ''�I�����ꂽ�R���{�{�b�N�X�̒l��ϐ��Ɋi�[
    'selectedCombobox = cboSelect.Text
    
    '2�s�ڂ���ŏI�s�܂ł̊ԂőI�����ꂽ�R���{�{�b�N�X�ƈ�v����^�C�g���̂���s��T��
    For startRow = 2 To lastRow
        If cboSelect.Text = Worksheets("Sheet2").Cells(startRow, titleColumn).Value Then
            linkRow = startRow
            Exit For
        End If
    Next
    
    'Do While cboSelect.Text <> Worksheets("Sheet2").Cells(startRow, titleColumn).Value
    
    
    '�R���{�{�b�N�X�̒l�ƈ�v�����^�C�g���̃����NURL��ϐ��Ɋi�[
    linkurl = Worksheets("Sheet2").Cells(linkRow, urlColumn).Value
    
    '�����N��http:~�̏ꍇ
    If Worksheets("Sheet2").Cells(startRow, urlColumn).Value Like "http*" Then
        Call OpenIE(linkurl)
        
    '����ȊO�i�t�@�C���p�X�j�̏ꍇ
    Else
        OpenFE linkurl
        
    End If
    
End Sub

Sub OpenIE(linkurl As String)
    Dim objIE As Object
    
    Set objIE = CreateObject("InternetExplorer.Application")
    objIE.Visible = True
    
    objIE.Navigate linkurl
    
    Call ApplicationClose
    
End Sub

Sub OpenFE(ByVal filePath As String)
    Shell "EXPLORER.EXE /select,""" & filePath & """", vbNormalFocus

    Call ApplicationClose
End Sub

Private Sub btnDelete_Click()
    Dim Msg As String
    Dim res As Integer
    Dim deleteRow As Integer
    
    Msg = "�{���ɍ폜���Ă�낵���ł����H"
    res = MsgBox(Msg, vbYesNo + vbExclamation + vbDefaultButton2, "�m�F")
    '���������������ꂽ�ꍇ�I������
    If res = vbNo Then Exit Sub
    
    '�폜�Ώۍs�̐ݒ�
    For startRow = 2 To lastRow
        If cboSelect.Text = Worksheets("Sheet2").Cells(startRow, titleColumn).Value Then
            deleteRow = startRow
            Exit For
        End If
    Next
    
    '�폜�̎��s
    Rows(deleteRow).Delete

    MsgBox "�폜���܂����B", vbInfomation, "�폜����"

    Unload Me
    
End Sub

Private Sub btnEnd_Click()
    Unload Me
End Sub

Sub ApplicationClose()
    Unload Me
    Application.Quit
    Windows("�����N�W.xlsm").Close True
End Sub
