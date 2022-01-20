VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm13 
   Caption         =   "Send Email"
   ClientHeight    =   1965
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   3960
   OleObjectBlob   =   "UserForm13.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
Unload UserForm13
End Sub

Private Sub cmdSendEmail_Click()
Application.Run "Module11.SAP_clerk_email_update"
End Sub
