VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SelectForm 
   Caption         =   "UserForm2"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   8910
   OleObjectBlob   =   "SelectForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "SelectForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const titleColumn As Integer = 2
Const urlColumn As Integer = 3
Dim lastRow As Integer

'マクロ起動時の処理
Private Sub UserForm_Initialize()
    
    lastRow = Worksheets("Sheet2").Range("B65536").End(xlUp).Row 'シート2のB65536セルから上へ向かって値が入力されているセルを探してlastRowに入れる
    
    'コンボボックスに値を入れる処理
    'クリア
    cboSelect.Clear
    
    '2列目の2行目から最終行まで値をコンボボックスに入れる
    For startRow = 2 To lastRow
       cboSelect.AddItem Worksheets("Sheet2").Cells(startRow, 2).Value
    Next
    '1番目のリストの値を選択状態にする
    cboSelect.Value = cboSelect.List(0)
    
End Sub

Private Sub btnJump_Click()
    'Dim selectedCombobox As String  '選択されたコンボボックスの値
    Dim linkurl As String           'リンク先のURL/PATH
    Dim linkRow As Integer          'リンク先のある行
    
    ''選択されたコンボボックスの値を変数に格納
    'selectedCombobox = cboSelect.Text
    
    '2行目から最終行までの間で選択されたコンボボックスと一致するタイトルのある行を探す
    For startRow = 2 To lastRow
        If cboSelect.Text = Worksheets("Sheet2").Cells(startRow, titleColumn).Value Then
            linkRow = startRow
            Exit For
        End If
    Next
    
    'Do While cboSelect.Text <> Worksheets("Sheet2").Cells(startRow, titleColumn).Value
    
    
    'コンボボックスの値と一致したタイトルのリンクURLを変数に格納
    linkurl = Worksheets("Sheet2").Cells(linkRow, urlColumn).Value
    
    'リンクがhttp:~の場合
    If Worksheets("Sheet2").Cells(startRow, urlColumn).Value Like "http*" Then
        Call OpenIE(linkurl)
        
    'それ以外（ファイルパス）の場合
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
    
    Msg = "本当に削除してよろしいですか？"
    res = MsgBox(Msg, vbYesNo + vbExclamation + vbDefaultButton2, "確認")
    'いいえが押下された場合終了する
    If res = vbNo Then Exit Sub
    
    '削除対象行の設定
    For startRow = 2 To lastRow
        If cboSelect.Text = Worksheets("Sheet2").Cells(startRow, titleColumn).Value Then
            deleteRow = startRow
            Exit For
        End If
    Next
    
    '削除の実行
    Rows(deleteRow).Delete

    MsgBox "削除しました。", vbInfomation, "削除成功"

    Unload Me
    
End Sub

Private Sub btnEnd_Click()
    Unload Me
End Sub

Sub ApplicationClose()
    Unload Me
    Application.Quit
    Windows("リンク集.xlsm").Close True
End Sub
