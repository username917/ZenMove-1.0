VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "Email Update"
   ClientHeight    =   4500
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   5220
   OleObjectBlob   =   "UserForm3.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ComboBox1_Change()
End Sub

Private Sub cboYesNo_Change()
Application.Run "Module11.enable_email"

    

    
End Sub

Private Sub chkPO_Click()
With UserForm3

    If .chkStaged = True And .chkReturn = True And .chkBlank = True Then

    .lblSure.Visible = True
    .cboYesNo.Visible = True
    
    End If
    
End With
End Sub

Private Sub cmdSend_Click()
Application.Run "Module11.sendmail"
End Sub

