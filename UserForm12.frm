VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm12 
   Caption         =   "PO Information System/SAP Movement Confirmation"
   ClientHeight    =   4845
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   8160
   OleObjectBlob   =   "UserForm12.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer, j As Integer, mat As Double, mat_r As Double
Public eval_rng As Range

Private Sub cmdExit_Click()
Unload UserForm12
End Sub

Sub cmdGetMovts_Click()
ThisWorkbook.Sheets("Staging").Range("AA3").Interior.Color = vbYellow
Application.Run "Module11.analyze_SAP_PO"
End Sub

Private Sub cmdI1_Click()

'rationale for turning the cell red is to go to the correct subroutine when a PO is picked and found
'to avoid writing the same routines over and over again

ThisWorkbook.Sheets("Staging").Range("AA2").Interior.Color = vbBlue
Application.Run "Module11.analyze_SAP_PO"
'Application.Run "Module11.get_SKUs_to_sign"

End Sub

Private Sub cmdI2_Click()

'rationale for turning the cell red is to go to the correct subroutine when a PO is picked and found
'to avoid writing the same routines over and over again

ThisWorkbook.Sheets("Staging").Range("AA2").Interior.Color = vbBlack
Application.Run "Module11.analyze_SAP_PO"
'Application.Run "Module11.get_SKUs_to_sign"

End Sub

Private Sub cmdRefresh_Click()

With UserForm12

    .lstSTG.Clear
    .lstRTN.Clear
    .cboPOSAP.Clear
    .lblF1.Caption = ""
    .lblT1.Caption = ""
    .lblQ1.Caption = ""
    .lblF2.Caption = ""
    .lblT2.Caption = ""
    .lblQ2.Caption = ""
    
End With

Application.Run "Module11.POss_for_SAP"
End Sub





Private Sub lstRTN_Click()
ThisWorkbook.Sheets("Staging").Range("AA1").Interior.Color = vbCyan
Application.Run "Module11.analyze_SAP_PO"

End Sub



Private Sub lstSTG_Click()
ThisWorkbook.Sheets("Staging").Range("AA1").Interior.Color = vbGreen
Application.Run "Module11.analyze_SAP_PO"

End Sub



Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
'this routine will keep checking if the SAP textbook is not empty or numeric

With UserForm12

'test if staging listbox is empty and return movements are ready to be started (re-enable the returns listbox and the button)
If .lstSTG.ListCount = 0 Then

    'ensure that SAPQ for returns is not empty
    If Not IsEmpty(.txtSAPQ2.Value) Then
    
        'ensure that SAPQ is a number
        If Not IsNumeric(.txtSAPQ2) Then
        
            .lstRTN.Enabled = True
            .cmdI2.Enabled = True
            
            
        Else
        
            .lblSAPWarn.Caption = "SAP Quantity must be numeric"
            
        End If
    
    .lblSAPWarn.Caption = ""
    
    End If
    
End If
    
End With


End Sub
