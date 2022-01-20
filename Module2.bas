'Callback for login onAction
Sub Callback(control As IRibbonControl)

'easy login for drivers

UserForm11.Show vbModeless

End Sub

'Callback for logout onAction
Sub Callback2(control As IRibbonControl)

'log out option for drivers

With ThisWorkbook.Sheets("Staging")

    .Range("E1:I1").Clear
    .txtSPW = ""
    
End With

End Sub

'Callback for movtutil onAction
Sub Callback3(control As IRibbonControl)

'movement utility for drivers

UserForm9.Show vbModeless


End Sub

'Callback for GSheetsUpdate onAction
Sub Callback4(control As IRibbonControl)

'submit driver update to gmail sheet

Application.Run "Module1.login_gmail"
End Sub

'Callback for Email onAction
Sub Callback5(control As IRibbonControl)

'back up driver email update to clerk on duty

Application.Run "Module1.are_you_sure"
End Sub

'Callback for loginn onAction
Sub Callback6(control As IRibbonControl)

'easy login for SAP clerks

UserForm11.Show vbModeless

End Sub

'Callback for logoutt onAction
Sub Callback7(control As IRibbonControl)

'log out option for SAP clerks

With ThisWorkbook.Sheets("Staging")

    .Range("E1:I1").Clear
    .txtSPW = ""
    
End With
End Sub

'Callback for dwldDrv onAction
Sub Callback8(control As IRibbonControl)

'SAP clerk download of driver-submitted update

Application.Run "Module1.SAP_Monkey_download"

End Sub

'Callback for openSysInfo onAction
Sub Callback9(control As IRibbonControl)

'this control opens the SAP clerk's PO information system

Application.Run "Module1.POss_for_SAP"

End Sub

'Callback for GSheets onAction
Sub Callback10(control As IRibbonControl)

'this control opens the menu from which clerk submit orders to a final google sheet

Application.Run "Module1.evaluate_orders_for_Googlee"

End Sub

'Callback for backupEmail onAction
Sub Callback11(control As IRibbonControl)

'this control allows the SAP clerk to open the file they've downloaded from the driver and update the sheet in Excel

Application.Run "Module1.SAP_Clerk_update"


End Sub

'Callback for print onAction
Sub Callback12(control As IRibbonControl)

'this control allows the SAP clerk to print a selected order

Application.Run "Module1.show_print_utility"
    
End Sub

'Callback for archive onAction
Sub Callback13(control As IRibbonControl)

'this control activates the archiving option

Application.Run "Module1.archiving"

End Sub

'Callback for loginutil onAction
Sub Callback14(control As IRibbonControl)

'this control activates userform2, the more sophisticated login utility

UserForm2.Show vbModeless

End Sub

'Callback for profmgt onAction
Sub Callback15(control As IRibbonControl)

'this control open the master menu for managing SAP clerk and driver profiles from an admin

UserForm6.Show vbModeless

End Sub

'Callback for profcrt onAction
Sub Callback16(control As IRibbonControl)

'this control allow for the creation of a new profile, which can also be accessed from the login menu

UserForm5.Show vbModeless

End Sub

'Callback for chgpwd onAction
Sub Callback17(control As IRibbonControl)

'this control allows for changing the password on a profile

UserForm2.Show vbModeless

End Sub

'Callback for clearDriverSheet onAction
Sub Callback19(control As IRibbonControl)

'this routine is going to clear the driver sheet

Application.Run "Module1.clear_DRV_sheet"

End Sub

'Callback for PickClerk onAction
Sub Callback20(control As IRibbonControl)

Module1.clerk_list
'Application.Run "Module1.clerk_list"

End Sub

'Callback for backupEmailSend onAction
Sub Callback21(control As IRibbonControl)

'this control is going to extract the Staging and Archive worksheets into a blank file and send them to a recipient

UserForm13.Show vbModeless

End Sub


