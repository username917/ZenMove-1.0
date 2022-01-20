VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm10 
   Caption         =   "Submit PO to Google Sheets"
   ClientHeight    =   3825
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   2784
   OleObjectBlob   =   "UserForm10.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdQuit_Click()
Unload UserForm10
End Sub

Private Sub cmdSubmit_Click()
Application.Run "Module11.SAP_login_gmail"
End Sub

