VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm9 
   Caption         =   "Movementology 101"
   ClientHeight    =   7965
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   7980
   OleObjectBlob   =   "UserForm9.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cboOOs_Change()
With UserForm9

If .optReg.Value = True And .optReturn.Value = True Then
Application.Run "Module11.reg_return"
End If

End With

End Sub

Private Sub cboOPO_Change()

Dim i As Integer
i = cboOPO.ListIndex
If i > -1 Then
UserForm9.lblActivePO = cboOPO.List(i)
End If
UserForm9.txtPO = ""
Application.Run "Module11.opt_adjjj"

End Sub

Private Sub chkSame_Click()
If chkSame.Value = False Then

    'UserForm9.txtSKU = ""
    
End If

    
End Sub

Private Sub cmdClear_Click()

With UserForm9

    .txtPO = ""
    .txtSKU = ""
    .txtFrom = ""
    .txtTo = ""
    .txtQty = ""
    
End With

End Sub

Private Sub cmdHelpMe_Click()

End Sub

Private Sub cmdMoveMe_Click()

With UserForm9

    If .optAdjust = True And .optStaging = True Then
    
        Application.Run "Module11.check_movt_inputs"
        'Application.Run "Module11.repeating_orders"
        Exit Sub
        
    End If
    
        

    If .optAdjust.Value = True And .optReturn.Value = True Then
    
        Application.Run "Module11.STOCK_rtn_movts"
        
    'Else
    
        'Application.Run "Module11.check_movt_inputs"
        
        
    End If
    
    If .optReg = True And .optStaging = True Then
    
        Application.Run "Module11.check_movt_inputs"
        
        'check_movt_inputs feeds directly into repeating_orders, so it does not need to be referenced from the button
        'Application.Run "Module11.repeating_orders"
    
    End If
    
    
End With

'disable after each movement until all inputs are correct as evaluated in userform9 code
cmdMoveMe.Enabled = False

End Sub

Private Sub cmdNext_Click()
With UserForm9
If .chkSame.Value = False Then

    .txtSKU = ""
    
End If

.txtFrom = ""
.txtTo = ""
.txtQty = ""

End With

End Sub

Private Sub cmdQuit_Click()
Unload UserForm9
End Sub


Private Sub cmdRefresh_Click()

End Sub

Private Sub cmdRtn_Click()
Application.Run "Module11.do_rtn"

'disable command after move until next set of data is evaluated as good in userform9
cmdRtn.Enabled = False
End Sub

Private Sub Label22_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
Label22.Caption = "This application shows the current PO status, " _
& "tracks movement accuracy in real time and allows for staging, returning, stock to stock and adjusted movements."

End Sub



Private Sub lblHelp_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
With UserForm9.lblHelp

.Caption = "This utility allows you to do regular movements and secondary movements after an order has been submitted for posting."

End With
End Sub

Private Sub lstUnfinishedBusiness_Click()
Dim cell2 As Range, lstcl2 As Variant, i As Integer, j As Integer, mat As Double, PO As Variant

'set the active selection in the listbox and assign the contents of the position to a variable

With UserForm9

    If .optReg.Value = True And .optReturn.Value = True Then
    
    Application.Run "Module11.do_rtn"
    
    End If
    
End With

With UserForm9

    If .optAdjust.Value = True And .optReturn.Value = True Then

i = .lstUnfinishedBusiness.ListIndex
mat = UserForm9.lstUnfinishedBusiness.List(i)
j = .cboOOs.ListIndex
If j > -1 Then
PO = .cboOOs.List(j)
End If

    With ThisWorkbook.Sheets("Staging")
    
        lstcl2 = .Range("M10000").End(xlUp).row
        
        For Each cell2 In .Range("M4:M" & lstcl2)
        
            If cell2 = mat And cell2.Offset(0, -1) = PO Then
            
                With UserForm9
                
                    'insert from location into user form
                    .txtFrom2 = cell2.Offset(0, 1)
                    'populate material label
                    .labMat = mat
                
                End With
        
            End If
    
        Next
    
    End With
End If
End With
End Sub

Private Sub optAdjust_Click()

If Not switch = False Then

    If optAdjust.Value = True Then
    Application.Run "Module11.opt_adjjj"
    End If

End If

End Sub

Private Sub optReg_Click()

If Not swtich = False Then

    If optReg.Value = True Then
    Application.Run "Module11.opt_adjjj"
    End If

End If

End Sub

Private Sub optReturn_Click()

    'does not need switch variable, clicking return does not clean out any listboxes...

    Application.Run "Module11.reg_open_orders"

End Sub

Private Sub optStaging_Click()



    If UserForm9.optStaging.Value = True Then
    
        'MsgBox "Staged movements must be returned, unless permanently on line.", vbInformation
    End If
    
    UserForm9.txtPO.Enabled = True
    
    With UserForm9
    
        .lstUnfinishedBusiness.Clear
        .cboOOs.Clear
        
    End With
    


End Sub

Private Sub txtPO_Change()
If Not txtPO = "" Then
UserForm9.lblActivePO.Caption = txtPO
Else
Dim i As Integer
i = UserForm9.cboOPO.ListIndex
If i > -1 Then
UserForm9.lblActivePO.Caption = UserForm9.cboOPO.List(i)
End If
End If
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)


'the logic of this routine is that combo box need to be refreshed based on the changes in the made movements, and these _
'are connected to the choices made in the type of movement

'public variable that is intended to prevent radio buttons running opt_adjjj when mouse moves in form
switch = False

'running opt_adjjj is usually done when the user switches options for the type of movement
Application.Run "Module11.opt_adjjj"

'make lstactivepo bold

lblActivePO.Font.Bold = True

With UserForm9

    If .optReg.Value = True Then
    
        .optReg.Value = False
        .optReg.Value = True
        
    End If
    
    If .optAdjust.Value = True Then
    
        .optAdjust.Value = False
        .optAdjust.Value = True
        
    End If
    
    'take out staging routine altogether from the autorefresh
    'If .optStaging.Value = True Then
    
        '.optReturn.Value = True
        '.optStaging.Value = True
        
    'End If
    
    If .optReturn.Value = True Then
    
        .optReturn.Value = False
        .optReturn.Value = True
        
    End If
    
    If .optAdjust.Value = True And .optReturn.Value = True Then
    
        .optReturn.Caption = "STOCK TO STOCK Movement"
        
    Else
    
        .optReturn.Caption = "Return Movement"
        
    End If
    
    'checking staging for accuracy before enabling the Move Me button
    
    'ensure that PO numbers are present
    If Not IsEmpty(.txtPO.Value) Or Not IsEmpty(.lblActivePO) Then
    
        'ensure that no fields are left empty
        If Not IsEmpty(.txtSKU.Value) And Not IsEmpty(.txtQty.Value) And Not IsEmpty(.txtFrom.Value) And Not IsEmpty(.txtTo.Value) Then
        
            'checking criteria for SKU: that it's 9 digits long, begins with 300 and is a number
            If Len(.txtSKU.Value) = 9 And Left(.txtSKU.Value, 3) = 300 And IsNumeric(.txtSKU.Value) Then
            
                'checking crtieria for quantity
                If IsNumeric(.txtQty.Value) Then
                
                    'checking from location: that it begins with a letter, it's no less than 5 and up to 6 digits long and
                    If Not IsNumeric(Left(.txtFrom.Value, 1)) Then
                    
                        'checking length for from location to be 5 or 6 digits
                        If Len(.txtFrom.Value) = 2 Or Len(.txtFrom.Value) = 5 Or Len(.txtFrom.Value) = 6 Then
                        
                            'checking To location criteria: identical to from loocation criteria
                            If Not IsNumeric(Left(.txtTo.Value, 1)) Then
                            
                                'checking to see if SAP quantity is numeric
                                If IsNumeric(.txtSAPQ) Then
                            
                                    If Len(.txtTo.Value) = 2 Or Len(.txtTo.Value) = 5 Or Len(.txtTo.Value) = 6 Then
                                    
                                        .cmdMoveMe.Enabled = True
                                        .lblWarn.Caption = ""
                                    
                                    Else
                                    
                                        .lblWarn.Caption = "Check the length of TO location address"
                                        
                                    End If
                                    
                                Else
                                .lblWarn = "SAP Quantity must be a number"
                                
                                End If
                                
                            Else
                                .lblWarn.Caption = "TO location address must begin with a letter"
                                
                            End If
                        
                        Else
                        
                            .lblWarn.Caption = "Check the length of FROM Location"
                            
                        End If
                        
                    Else
                    
                        .lblWarn.Caption = "FROM location needs to begin with a letter"
                        
                    End If
                    
                Else
                
                    .lblWarn.Caption = "Please check your quantity"
                    
                End If
            
            
            Else
            
                .lblWarn.Caption = "Please check accuracy and length of SKU number"
                
            End If
            
        End If
        
        End If
        
        'checking return locations
        If Not IsEmpty(.txtFrom2.Value) And Not IsEmpty(.txtTo2.Value) And Not IsEmpty(.txtQty2.Value) Then
        
            'From location populates automatically based on what the user clicks in the unfinished business listbox
            
            'checking TO box
            If Not IsNumeric(Left(.txtTo2.Value, 1)) Then
                            
                If Len(.txtTo2.Value) = 5 Or Len(.txtTo2.Value) = 6 Then
                
                    'checking Qty2 box
                    If IsNumeric(.txtQty2.Value) Then
                    
                        .cmdRtn.Enabled = True
                        .lblWarn2.Caption = ""
                    
                    
                    Else
                    
                        .lblWarn2.Caption = "Please check your quantity"
                    
                    End If
            
                Else
                                
                    .lblWarn2.Caption = "Check the length of TO location address"
                                    
                End If
                                
            Else
            
                .lblWarn2.Caption = "TO location address must begin with a letter"
                                
        End If
            
    End If
      
    'enable return button if there is indeed no return to be made for the material
    If .chkNoRtn.Value = True Then
      
        .cmdRtn.Enabled = True
        
    End If
    
        
            
        
    

End With
    
    
End Sub
