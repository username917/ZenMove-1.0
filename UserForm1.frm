VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Active Clerk Utility"
   ClientHeight    =   2820
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   4704
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cboClerk_Change()
With ThisWorkbook.Sheets("TopSecret")

    lstcl = .Range("C10000").End(xlUp).row
    i = UserForm1.cboClerk.ListIndex
    
    For Each cell In .Range("C2:C" & lstcl)
    
        If cell & " " & cell.Offset(0, 1) = UserForm1.cboClerk.List(i) Then
        
            UserForm1.lblEmail = cell.Offset(0, 2) & "@diversey.com"
            
        End If
        
    Next
    
End With


Application.Run "Module11.set_clerk"
    
End Sub

Private Sub cmdHelp_Click()

UserForm1.lblHelp = "This utility helps you set the clerk on shift to whom you will send your movement and returns updates"


End Sub

Private Sub cmdOK_Click()
Application.Run "Module11.set_clerk"
End Sub

Private Sub cmdQuit_Click()
MsgBox "You cannot quit. You must choose a clerk on duty before proceeding", vbInformation
End Sub

