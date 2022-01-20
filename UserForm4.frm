VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm4 
   Caption         =   "Order Print Menu"
   ClientHeight    =   1920
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   4404
   OleObjectBlob   =   "UserForm4.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdPrint_Click()
Call Module11.print_stuff
End Sub

Private Sub cmdQuit_Click()
UserForm4.cboPOs.Clear
Unload UserForm4
End Sub

