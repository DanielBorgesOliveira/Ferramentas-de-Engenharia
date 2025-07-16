VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "UserForm2"
   ClientHeight    =   5955
   ClientLeft      =   60
   ClientTop       =   270
   ClientWidth     =   7905
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
Option Base 0

Private Sub Label10_Click()

End Sub

Private Sub Label2_Click()

End Sub

Private Sub Label12_Click()

End Sub

Private Sub Label13_Click()

End Sub

Private Sub Label3_Click()

End Sub

Private Sub UserForm_Initialize()
    TextBox3.AddItem "Anglo American"
    TextBox3.AddItem "CBMM"
    TextBox3.AddItem "Itaminas"
    TextBox3.AddItem "Samarco"
    TextBox3.AddItem "Vale"
    TextBox3.AddItem "Vallourec"
End Sub

Private Sub CommandButton1_Click()
    Me.Hide
End Sub

Private Sub CommandButton5_Click()
    'Variaveis.userformDirname = libUtils.UseFolderDialog
    UserForm2.TextBox5.text = libUtils.UseFileDialog(DialogType:=msoFileDialogFilePicker)
End Sub

Private Sub CommandButton7_Click()
    'Variaveis.userformDirname = libUtils.UseFolderDialog
    UserForm2.TextBox7.text = libUtils.UseFileDialog(DialogType:=msoFileDialogSaveAs)
End Sub

Private Sub CommandButton6_Click()
    'Variaveis.userformDirname = libUtils.UseFolderDialog
    UserForm2.TextBox6.text = libUtils.UseFileDialog(DialogType:=msoFileDialogSaveAs)
End Sub


