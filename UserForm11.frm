VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm11 
   Caption         =   "Express Login"
   ClientHeight    =   3750
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   5700
   OleObjectBlob   =   "UserForm11.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGetin_Click()
Application.Run "Module11.easy_login"
End Sub

Private Sub CommandButton2_Click()
Unload UserForm11
End Sub

