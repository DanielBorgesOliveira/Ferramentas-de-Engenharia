VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   10155
   ClientLeft      =   60
   ClientTop       =   270
   ClientWidth     =   6480
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
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

Private Sub Label3_Click()

End Sub

Private Sub UserForm_Initialize()
    Cliente.AddItem "Anglo American"
    Cliente.AddItem "CBMM"
    Cliente.AddItem "Itaminas"
    Cliente.AddItem "Samarco"
    Cliente.AddItem "Vale"
    Cliente.AddItem "Vallourec"
End Sub

Private Sub CommandButton1_Click()
    Me.Hide
End Sub
