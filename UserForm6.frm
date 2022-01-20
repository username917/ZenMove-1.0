VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm6 
   Caption         =   "Admin Tools"
   ClientHeight    =   4200
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   7752
   OleObjectBlob   =   "UserForm6.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmdAdd_Click()
With UserForm6

.frmPeople.Visible = True
.frmPeople.Caption = "Add Person"

    
End With

End Sub

Private Sub cmdAddNext_Click()
With UserForm6

    .txtFirst = ""
    .txtLast = ""
    
End With
End Sub

Private Sub cmdChange_Click()
With UserForm6

    .frmPeople.Visible = True
    .frmPeople.Caption = "Change Name"
    .cmdConfAdd.Caption = "Confirm and Change"
    
    
End With
End Sub

Private Sub cmdClear_Click()

With UserForm6

    .txtFirst = ""
    .txtLast = ""
    
End With

End Sub

Private Sub cmdConfAdd_Click()
With UserForm6

    'ensuring there are no empty textboxes

    If IsEmpty(.txtFirst) Or IsEmpty(.txtLast) Then
    
        MsgBox "Please fill all names before proceeding!", vbExclamation
        Exit Sub
        
    End If

    If .optClerks.Value = True And .frmPeople.Caption = "Add Person" Then
    
        Application.Run "Module11.admin_add_clerks"
        Exit Sub
        
    ElseIf .optDrivers.Value = True And .frmPeople.Caption = "Add Person" Then
    
        Application.Run "Module11.admin_add_drivers"
        Exit Sub
        
    ElseIf .optClerks.Value = True And .frmPeople.Caption = "Change Name" Then
    
        Application.Run "Module11.change_person"
        .cmdConfAdd.Caption = "Confirm and Change"
        
    ElseIf .optDrivers.Value = True And .frmPeople.Caption = "Change Name" Then
    
        Application.Run "Module11.change_person"
        .cmdConfAdd.Caption = "Confirm and Change"
    
    Else
    
        MsgBox "Please select an option from the Personnel Management menu.", vbExclamation
        Exit Sub
        
    End If
    
    
    If .frmPeople.Caption = "Add Person" Then
    
        .cmdConfAdd.Caption = "Confirm and Add"
        
    End If
    
    
.frmPeople.Visible = False

    
End With
End Sub

Private Sub cmdGo_Click()
'Application.Run "Module1.make_roster"
End Sub

Private Sub cmdGTFO_Click()
Unload UserForm6
End Sub

Private Sub cmdHide_Click()
UserForm6.frmPeople.Visible = False
End Sub

Private Sub cmdOpenTS_Click()

End Sub

Private Sub cmdRemove_Click()
Application.Run "Module11.remove_person"
End Sub


Private Sub optClerks_Click()
Application.Run "Module11.make_roster"
UserForm6.frmPeople.Visible = False
End Sub

Private Sub optDrivers_Click()
Application.Run "Module11.make_roster"
UserForm6.frmPeople.Visible = False
End Sub


