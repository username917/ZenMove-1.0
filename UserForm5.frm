VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm5 
   Caption         =   "Create New Profile"
   ClientHeight    =   4695
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   7644
   OleObjectBlob   =   "UserForm5.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmdCreate_Click()
Application.Run "Module11.new_profile"
End Sub

Private Sub cmdExit_Click()
Unload UserForm5
UserForm2.Show vbModeless
End Sub

Private Sub CommandButton1_Click()
Application.Run "Module11.confirm_profile_deets"
End Sub

Private Sub lblSummary_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
With UserForm5.lblSummary

    .Caption = ("This uitlity allows you to create up to one user profile at a time. Enter your credentials and confirm them before proceeding.")
    
End With
End Sub

Private Sub txtPword1_Change()
UserForm5.txtPword1.PasswordChar = "*"
End Sub

Private Sub txtPword2_Change()
UserForm5.txtPword2.PasswordChar = "*"
End Sub

