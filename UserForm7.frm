VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm7 
   Caption         =   "Change Password"
   ClientHeight    =   3045
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   4380
   OleObjectBlob   =   "UserForm7.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdChangePW_Click()
Application.Run "Module11.change_password"
End Sub

Private Sub cmdQuit_Click()
Unload UserForm7

End Sub

Private Sub txtNewP1_Change()
UserForm7.txtNewP1.PasswordChar = "*"
End Sub

Private Sub txtNewP2_Change()
UserForm7.txtNewP2.PasswordChar = "*"
End Sub
