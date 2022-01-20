VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "Login Credentials"
   ClientHeight    =   4710
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   5688
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdChgPW_Click()

With UserForm2

    If IsEmpty(.txtEmail) Then
    
        MsgBox "Please enter your username before trying to change your password", vbInformation
        
    End If
    
End With
UserForm7.lblUname = UserForm2.txtEmail
UserForm7.Show



End Sub

Private Sub cmdLogin_Click()
UserForm2.Hide
Application.Run "Module11.login"
End Sub

Private Sub cmdNewProfile_Click()
UserForm5.Show vbModeless
UserForm2.Hide

End Sub

Private Sub cmdTest_Click()
Application.Run "Module1.testmail"
End Sub

Private Sub lblExplain_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
UserForm2.lblExplain = "This utility prompts you to enter your password twice, in order to ensure that it is entered correctly. It will compare the two inputs and alert you if the " _
& "passwords do not match. You can also send a test email to yourself to see that you have entered your details correctly."

End Sub

Private Sub txtPW1_Change()
UserForm2.txtPW1.PasswordChar = "*"
End Sub

Private Sub txtPW2_Change()
UserForm2.txtPW2.PasswordChar = "*"
End Sub


