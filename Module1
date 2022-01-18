Public username As String, pw1 As Variant, pw2 As Variant
Public switch As Boolean 'this variable is intended to prevent listboxes from triggering opt_adjjj when _
they are refreshed in userform9 on the mousemove event

Public eval_rng As Range   'this variable is going to prevent an eternal loop with get_SKUs routine

Dim myHTML_Element As IHTMLElement
Dim Driver As New WebDriver

'Dim eval_rng As Range
Dim cell_po As Range 'need this variable for going between the SAP PO routine and the analysis of each order


Sub clerk_list()
Application.ScreenUpdating = False

'this routine builds and rebuilds the list of clerks every time it is called from userform1

Dim arr_clerk() As Variant, i As Integer, lstcl As Variant, cell As Range



With ThisWorkbook.Sheets("TopSecret")

    .Visible = True
    .Activate

    lstcl = .Range("C10000").End(xlUp).row
       
    If Not IsEmpty(.Range("C2")) Then
        
        For i = 2 To lstcl
            
                'remove second loop, make i as the cell row and go through the motions
                            
                ReDim Preserve arr_clerk(i)
                arr_clerk(i - 2) = Range("C" & i) & " " & Range("C" & i).Offset(0, 1)
            
        Next
        UserForm1.cboClerk.List = arr_clerk()
        UserForm1.Show vbModeless
    End If
    
    .Visible = xlVeryHidden
    
End With
    
End Sub

Private Sub set_clerk()
Application.ScreenUpdating = False


'this routine sets the active clerk

Dim name As String, clerk As String, email As Variant
Dim rng As Range


With UserForm1.cboClerk

    'populate the combobox


    'set the name equal to the active choice in the combo box, set the selection equal to position in email array _
    'to recall the right email
    
    'in case name is nothing, exit sub
    If name <> "" Then
        Exit Sub
    Else
        name = .List(.ListIndex)
    End If

  
    
    With ThisWorkbook.Sheets("TopSecret")
    
        .Visible = True
        .Activate
    
        lstcl = .Range("C10000").End(xlUp).row
        
        
        'make a range from the
        Set rng = .Range("C2:C" & lstcl)
        
            For Each cell In rng
            
                If cell & " " & cell.Offset(0, 1) = name Then
                
                    email = cell.Offset(0, 2) & "@diversey.com"
                
                    With ThisWorkbook.Sheets("Staging")
                    
                        .Activate
                        .Range("A1") = "Clerk on duty: " & name
                        .Range("C1") = email
                        Columns("A:C").AutoFit
                        
                        MsgBox "Clerk on duty set as: " & cell & " " & cell.Offset(0, 1), vbInformation
                        Unload UserForm1
                        
                    End With
            
                End If
            
            Next
            
        .Visible = xlVeryHidden
    
    End With
    
End With



End Sub



Private Sub sendmail()

Application.ScreenUpdating = False

'saveas the file that will be sent to the clerk on duty

'method is to copy the first sheet in a new workbook, save it on the desktop and send it to the clerk

Dim xNewWb As Workbook, lstcl As Variant, data_to_send As Variant, fname As Variant, uname As String, myname As Variant

Set xNewWb = Workbooks.Add
uname = ThisWorkbook.Sheets(1).Range("E1")
fname = "Driver Update - " & uname & " " & Format(Date, "yyyymmdd") & ".xlsx"

ThisWorkbook.Activate

With ThisWorkbook

lstcl = .Sheets("Staging").Range("C10000").End(xlUp).row

data_to_send = .Sheets("Staging").Range("A2:Z" & lstcl).Copy

'    Range(data_to_send).Copy
xNewWb.Sheets(1).Range("A1").PasteSpecial xlPasteValuesAndNumberFormats

With xNewWb
    myname = Environ("username")
    .SaveAs Filename:="C:\Users\" & myname & "\Desktop\" & fname
    
    'sort everything according to the timestamps
    
    'determine the used range
    
    Dim lstcl2 As Variant, xSort As Range
    
    With .Sheets(1)
        .Activate
        lstcl2 = .Range("U10000").End(xlUp).row
        
        Set xSort = .Range(Range("A4").Address, Range("X" & lstcl2).Address)
        
        Set x = Range(.Range("V4").Address, .Range("V" & lstcl2).Address)
        'X.Select
        
        xSort.Select
        .Sort.SortFields.Clear
        .Sort.SortFields.Add Key:=x, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        Columns("A:AA").AutoFit

       With .Sort
        .SetRange xSort
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
        End With

    End With
    
    .Save
    .Close
    
End With

With ThisWorkbook.Sheets("Staging")

    'set username and password
    username = .Range("E1")
    pw1 = .txtSPW
    

End With

End With
    
            Dim mail As CDO.Message
            Dim config As CDO.Configuration
            
            Set mail = CreateObject("CDO.message")
            Set config = CreateObject("Cdo.configuration")
            
            config.Fields(cdoSendUsingMethod).Value = cdoSendUsingPort
            config.Fields(cdoSMTPServer).Value = "aspmx.l.google.com"
            'config.Fields(cdoSMTPServer).Value = "smtp.gmail.com"
            config.Fields(cdoSMTPServerPort).Value = 25 '25
            config.Fields(cdoSMTPAuthenticate).Value = cdoNTLM ' cdoBasic  'both cdoNTLM and cdoBasic work in communicating with gmail
            config.Fields(cdoSendUserName).Value = username
            config.Fields(cdoSendPassword).Value = pw1
            config.Fields.Update
            
            Set mail.Configuration = config
            
            With mail
            
                'get clerk from name set on sheet
            
                Dim rcpt As Variant
                rcpt = ActiveSheet.Range("C1")
            
                .To = rcpt
                .From = username & "@diversey.com"
                .Subject = "Updated Staging Information - " & username
                .TextBody = "Please update these accordingly."
                
                .AddAttachment "C:\Users\" & myname & "\Desktop\" & fname
                
                
                .Send
                
            End With
            
            Set config = Nothing
            Set mail = Nothing

UserForm3.Hide


End Sub

Private Sub testmail()

Application.ScreenUpdating = False
            
With UserForm2

username = .txtEmail
pw1 = .txtPW1
pw2 = .txtPW2
            
  If Not IsEmpty(.txtEmail) And Not IsEmpty(.txtPW1) And Not IsEmpty(.txtPW2) Then

    If pw1 = pw2 Then
    
        MsgBox "You are about to send a test mail to yourself. Check your email to ensure credentials are accurate."
                    
            Dim mail As CDO.Message
            Dim config As CDO.Configuration
            
            Set mail = CreateObject("CDO.message")
            Set config = CreateObject("Cdo.configuration")
            
            config.Fields(cdoSendUsingMethod).Value = cdoSendUsingPort
            config.Fields(cdoSMTPServer).Value = "aspmx.l.google.com"
            config.Fields(cdoSMTPServerPort).Value = 25 '25
            config.Fields(cdoSMTPAuthenticate).Value = cdoNTLM ' cdoBasic  'both cdoNTLM and cdoBasic work in communicating with gmail
            config.Fields(cdoSendUserName).Value = username
            config.Fields(cdoSendPassword).Value = pw1
            config.Fields.Update
            
            Set mail.Configuration = config
            
            With mail
            
                'get clerk from name set on sheet
            
                Dim rcpt As Variant
                rcpt = ActiveSheet.Range("C1")
            
                .To = rcpt
                .From = username & "@diversey.com"
                .Subject = "Self-test email - " & username
                .TextBody = "Your login credentials are accurate."
                
                '.addattachment "Path"
                
                .Send
                
            End With
            
            Set config = Nothing
            Set mail = Nothing
            

    Else
    
        MsgBox "Check your login information!"


    End If
    
End If

End With


End Sub

Private Sub SAP_clerk_email_update()

'this sub is going to be the backup transfer protocol for SAP clerks

'logic: extract Staging and Archive to a blank file, save and email to someone on a list or as typed in a textbox

Dim xNewWb As Workbook, lstA As Variant, bldSTG As Range, bldARCH As Range, fname As String
Dim username As String, pw1 As Variant

'start by exporting Staging sheet

Set xNewWb = Workbooks.Add

With ThisWorkbook.Sheets("Staging")

    .Activate

    'make a range of the entire range of data and populate in the new spreadsheet
    lstA = .Range("A100000").End(xlUp).row
    Set bldSTG = Range("A2:AA" & lstA)
    
    xNewWb.Sheets(1).Range("A2") = bldSTG

End With

'export Archive spreadsheet

With ThisWorkbook.Sheets("Archive")

.Activate

    lstA = .Range("A100000").End(xlUp).row
    Set bldARCH = Range("A2:A" & lstA)
    
    'set fname as the name of the person sending the email and use it in the filename when saving
    fname = ThisWorkbook.Sheets("Staging").Range("E1")
    
    With xNewWb
        'add sheet2 and paste the archive into it
        .Sheets.Add after:=Sheets(.Sheets.Count)
        
        With .Sheets(2)
        
        .Range("A2") = bldARCH
        
        End With
        
    End With
    
End With

'save as the new workbook that will be sent to the manager or whoever
With xNewWb

.Activate

    myname = Environ("username")
    .SaveAs Filename:="C:\Users\" & myname & "\Desktop\" & fname
    .Close
    
End With

'set username and password in preparation  for emailing

With ThisWorkbook.Sheets("Staging")

    username = .Range("E1")
    pw1 = .txtSPW
    
End With


            Dim mail As CDO.Message
            Dim config As CDO.Configuration
            
            Set mail = CreateObject("CDO.message")
            Set config = CreateObject("Cdo.configuration")
            
            config.Fields(cdoSendUsingMethod).Value = cdoSendUsingPort
            config.Fields(cdoSMTPServer).Value = "aspmx.l.google.com"
            config.Fields(cdoSMTPServerPort).Value = 25 '25
            config.Fields(cdoSMTPAuthenticate).Value = cdoNTLM ' cdoBasic  'both cdoNTLM and cdoBasic work in communicating with gmail
            config.Fields(cdoSendUserName).Value = username
            config.Fields(cdoSendPassword).Value = pw1
            config.Fields.Update
            
            Set mail.Configuration = config
            
            With mail
            
                'set the manager's email on the info in the textbox
                'ensure there is something in the textbox
            
                If Not IsEmpty(UserForm13.txtMGR) Then
                
                    Dim rcpt As Variant
                    rcpt = UserForm13.txtMGR.Value & "@diversey.com"
                
                Else
                
                    MsgBox "You must enter a recipient in the textbox", vbExclamation
                    Exit Sub
                
                End If
            
                .To = rcpt
                .From = username & "@diversey.com"
                .Subject = "Self-test email - " & username
                .TextBody = "Your login credentials are accurate."
                
                .AddAttachment "C:\Users\" & myname & "\Desktop\" & fname
                
                .Send
                
            End With
            
            Set config = Nothing
            Set mail = Nothing
End Sub
Private Sub are_you_sure()

Application.ScreenUpdating = False

'this sub puts in the options for the yes and no combo box before drivers send email to clerk

Dim arr_yesno(0 To 1) As String

With UserForm3.cboYesNo

arr_yesno(0) = "Yes!"
arr_yesno(1) = "No!"

.List = arr_yesno()

End With

UserForm3.Show vbModeless

End Sub

Private Sub enable_email()

Application.ScreenUpdating = False

'this sub checks whether the checks have been performed before sending email

With UserForm3

If .chkStaged = True And .chkReturn = True And .chkPO = True And .chkBlank = True And .cboYesNo.ListIndex = 0 Then

    .cmdSend.Enabled = True
    
End If

If .cboYesNo.ListIndex = 1 Then

    .cmdSend.Enabled = False
    
    MsgBox "Please double check your inputs!", vbExclamation
    UserForm3.Hide
    
End If

End With

End Sub

Private Sub end_shift()

Application.ScreenUpdating = False

'this routine is going to allow the drivers to clear the field at the end of the shift and start over the next day

'may be renamed as a clear sheet option

Dim answer As Integer, lstcl As Variant, rng_del As Range

answer = MsgBox("You are about to end your shift erase all data. Are you sure?", vbYesNo)

lstcl = ActiveSheet.Range("C10000").End(xlUp).row

With ActiveSheet

Set rng_del = Range(.Range("A4").Address, .Range("X" & lstcl).Address)

End With

If answer = vbYes Then

rng_del.Delete

With ThisWorkbook

    .Save
    .Close
    
End With

ElseIf answer = vbNo Then

Exit Sub

End If



End Sub

Private Sub SAP_Clerk_update()

Application.ScreenUpdating = False

'this routine will be used by the SAP Clerk to open the emailed file and automatically update its contents into the master list


'open the driver's update file

Dim myfile As Variant, drv_wbk As Workbook

myfile = Application.GetOpenFilename(Title:="Please choose driver update file to open:", FileFilter:="Excel Files *.xls* (*.xls*),")

If myfile <> False Then

    Workbooks.Open Filename:=myfile
    Set drv_wbk = Workbooks.Open(myfile)

End If

If myfile = False Then

    MsgBox "File not chosen!"
    Exit Sub
End If




'this is the part that will import information and establish the universal order of things with a time stamp

'build the range of timestamps in the Staging sheet

With ThisWorkbook.Sheets("Staging")

    .Activate

    lstcl = .Range("U100000").End(xlUp).row
    
    Set r1 = .Range("U5")
    Set r2 = .Range("U" & lstcl)
    Set rng_to_scan = .Range(r1, r2)
    
    
End With

'build the range which will be compared in the DriverUpdate sheet

With drv_wbk.Sheets(1)

    .Activate

    Columns("A:B").Delete
    lstcl2 = .Range("U100000").End(xlUp).row
    
    Set c1 = .Range("U4")
    Set c2 = .Range("U" & lstcl2)
    Set rng_to_eval = .Range(c1, c2)
    

End With

drv_wbk.Sheets(1).Activate
    
    'range in DriverUpdate which will be downloaded and compared against the Staging timestamps
    For Each cell In rng_to_eval
    
        cell.Activate
        
        'with each pass, new info is put into Staging, so the range should grow to include that data
        Set lstcl = ThisWorkbook.Sheets("Staging").Range("U100000").End(xlUp).row
        Set r2 = ThisWorkbook.Sheets("Staging").Range("U" & lstcl)
        Set rng_to_scan = Range(r1, r2)
        
        'range in Staging that will be searched for every timestamp in DriverUpdate
        With rng_to_scan
        
            '.Activate
        
            Set to_find = .Find(cell)
            
            'this section is to handle pasting return movement for orders that are already staged
            If Not to_find Is Nothing Then
            
                'when it finds the timestamp, the routine is going to make a range of the return movements and put them in
                'it will paste existing movements, as well as new ones
                Set t1 = cell.Offset(0, -6)
                Set t2 = cell.Offset(0, -3)
                'Set t3 = cell.Offset(0, -2)
                Set rng_to_tfr = Range(t1, t2)
                rng_to_tfr.Select
                rng_to_tfr.Copy
                
                With ThisWorkbook.Sheets("Staging")
                
                    .Activate
                    
                    'make the range where the return movement information is going to fit in
                    Set s1 = to_find.Offset(0, -6)
                    Set s2 = to_find.Offset(0, -3)
                    Set rng_paste = Range(s1, s2)
                    rng_paste.Select
                    
                    'populate the range in Staging with rtn movement information for the appopriate time stamp
                    rng_paste.Value = rng_to_tfr.Value
                
                End With
            
            End If
            
            If to_find Is Nothing Then
            
                'copy the row in the next available line in the main Staging sheet
                'build a range out of the active row
                
                
                
                Set t1 = cell.Offset(0, -20)
                Set t2 = cell
                Set rng_to_tfr = Range(t1, t2)
                rng_to_tfr.Select
                rng_to_tfr.Copy
                
                With ThisWorkbook.Sheets("Staging")
                
                    .Activate
                
                    lstcl = .Range("U100000").End(xlUp).row
                    
                    'copy the range to the Staging sheet
                    
                    'if pasting a movmement, it will have the same timestamp as above it, so offset paste range to next row
                    If cell.Offset(-1, 0) = cell Then
                        Set s1 = .Range("A" & lstcl + 1)
                    End If
                    
                    'if starting a new order, it will have a blank cell in the above row, so offset paste range by 2 to create a break btw orders
                    If cell.Offset(-1, 0) = "" Then
                        Set s1 = .Range("A" & lstcl + 2)
                    End If
                    Set s2 = s1.Offset(0, 20)
                    Set rng_paste = Range(s1, s2)
                    'rng_paste.Select
                    
                    '.Range("A" & lstcl).Offset(1, 0).Activate
                    rng_paste.Value = rng_to_tfr.Value
                    
                End With
                
            End If
            
        End With
        
        drv_wbk.Sheets(1).Activate
        
    Next
    
'sort orders by timestamp - critical part about establishing the continuity of production runs for the entire factory

'make a range of the contents in the spreadsheet

Dim lstC As Variant, rng_sort As Range

With ThisWorkbook.Sheets("Staging")

    .Activate

    lstC = .Range("C100000").End(xlUp).row
    Set rng_sort = .Range("A4:Z" & lstC)
    rng_sort.Select
    
    .Sort.SortFields.Add Key:=Range("U5:U" & lstC), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    
    With .Sort
    
        .SetRange rng_sort
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
End With


    
'insert blank rows as breaks between newly inserted orders

Dim cell3 As Range

With ThisWorkbook.Sheets("Staging")

    .Activate

    lstcl = .Range("U100000").End(xlUp).row
    
    For Each cell3 In .Range("B4:B" & lstcl)
    
        If Not IsEmpty(cell3) Then
        
            If Not IsEmpty(cell3.Offset(-1, -1)) Then
            
                Range(cell3.Address).EntireRow.Insert
            
                
            End If
    
        End If
    
    Next


End With


End Sub

Private Sub show_print_utility()

Application.ScreenUpdating = False

'this routine will allow the SAP clerk to print off the relevant movements for a production or mix order

'add REFRESH button to the list app?

Dim lstcl As Variant, cell As Range, arr_po() As Variant, n As Integer

With ThisWorkbook.Sheets("Staging")

lstcl = .Range("C10000").End(xlUp).row


n = 1
For Each cell In .Range("B4:B" & lstcl)

    If Not IsEmpty(cell) Then
    
        ReDim Preserve arr_po(n)
        arr_po(n) = cell
        UserForm4.cboPOs.List = arr_po()
        n = n + 1
        
    End If
    
Next

Erase arr_po()
n = 0

End With

UserForm4.Show vbModeless

End Sub

Sub print_stuff()

Application.DisplayAlerts = False

'this routine will work with the above routine in finding and printing the correct range, based on the choice in the listbox

Dim i As Integer, lstC As Variant, lstB As Variant, PO As Variant, fndPO As Range
Dim r1 As Range, r2 As Range, rng_PO As Range, cell As Range

With ThisWorkbook.Sheets("Staging")


'set the extent of the range with data for columns B and C and find the value of the combo box when chosen
lstC = .Range("C10000").End(xlUp).row
lstB = .Range("B10000").End(xlUp).row

i = UserForm4.cboPOs.ListIndex
If i > -1 Then
PO = UserForm4.cboPOs.List(i)
End If


    With .Range("B4:B" & lstB)
    
        Set fndPO = .Find(PO)
        
        'set the top of the range as the PO to be found
        Set r1 = fndPO
        
        If Not fndPO Is Nothing Then
        
            For Each cell In .Range(fndPO.Offset(1, 0), "B" & lstB)
            
                'ensure that the PO is not the last order in the list
                If fndPO = Range("B" & lstB) Then
                    
                    'set the bottom of th range as the bottom of the data
                    Set r2 = Range("B" & lstC)
                    Set rng_PO = Range(r1, r2.Offset(0, 19))
                    Exit For
                    
                End If
                
                'if not the last PO in the list, scan unitl finding the next order in the list
                
                If Not IsEmpty(cell) Then
                
                    'set the bottom range just above the next order in the list
                    Set r2 = cell.Offset(-1, 0)
                    Set rng_PO = Range(r1, r2.Offset(0, 19))
                    Exit For
                    
                End If
                
            Next
            
        End If
        
    End With
    
End With
    
        
        rng_PO.Select
        rng_PO.Copy
              
      
        Set xNewWs = Worksheets.Add
                                
        With xNewWs
                                
            .Activate
            .Range("A2").PasteSpecial xlPasteAllUsingSourceTheme
            With .Range("C1")
                                    
                .Value = "PULLS"
                .Font.Bold = True
                .Font.Size = 14
                .Font.Underline = True
                                    
            End With
                                    
            With .Range("N1")
                                    
                Value = "PUTBACKS"
                
                With .Font
                    .Bold = True
                    .Size = 14
                    .Underline = True
                End With
                                        
            End With
                                    
                                  
            .Columns("A:AA").AutoFit
            .PageSetup.Orientation = xlLandscape
            .PageSetup.PaperSize = xlPaperLetter
            .PageSetup.FitToPagesTall = 1
            .PageSetup.FitToPagesWide = 1
            .PageSetup.Zoom = False
            .PrintPreview
                                    
                                    
        End With
                                
    ThisWorkbook.Activate
    Selection.Clear
    xNewWs.Delete
    Range("A1").Activate

End Sub

Private Sub make_topsecret()

'this routine is going to differentiate between the drivers and SAP clerk logins and store their information in a separate sheet.

Dim xTopSecret As Worksheet


Set xTopSecret = Worksheets.Add

xTopSecret.name = "TopSecret"

ThisWorkbook.Sheets("TopSecret").Visible = xlSheetVeryHidden

End Sub

Private Sub admin_topsecret()

Dim ws As Worksheet

For Each ws In ThisWorkbook.Sheets

If ws.name = "TopSecret" Then

    ws.Visible = xlSheetVisible
    
End If

Next

End Sub

Private Sub admin_add_clerks()

Application.ScreenUpdating = False


'this rotuine enables adding people to the list of username

'create a string out of the first and last name

Dim fname As String, lname As String, usname As Variant, cell As Range, lstcl As Variant, cname As Variant


With UserForm6

fname = UCase(Left(.txtFirst, 1)) & Mid(.txtFirst, 2)
lname = UCase(Left(.txtLast, 1)) & Mid(.txtLast, 2)
usname = fname & " " & lname

'add an if condition to ensure that the first and last name does not exist

    With ThisWorkbook.Sheets("TopSecret")
    
    .Visible = True
    .Activate
    
    lstcl = .Range("C10000").End(xlUp).row
    
        With .Range("D2:D" & lstcl)
        
            Set cname = .Find(lname)
            
        End With
        
            If cname Is Nothing Then
                        
                
                    .Range("D" & lstcl).Offset(1, 0) = lname
                    .Range("D" & lstcl).Offset(1, -1) = fname
                    
                    
                    'stitch together username and add it to the roster
                    
                    'put it in lower case for potential email purposes
    
                    usname = LCase(fname) & "." & LCase(lname)
                    
                    .Range("D" & lstcl).Offset(1, 1) = usname
                    
            Else
            
                MsgBox "SAP Clerk arleady exists", vbInformation
                
            End If
            
    .Visible = xlVeryHidden
        
    End With
 
End With


With ThisWorkbook.Sheets("TopSecret")

    .Visible = True
    .Activate
    
    .Range("C1") = "SAP Clerks"
    .Range("G1") = "Drivers"
    
    .Visible = xlVeryHidden

End With

Application.Run "Module11.make_roster"

End Sub
Private Sub admin_add_drivers()

Application.ScreenUpdating = False

'this rotuine enables adding people to the list of username

'create a string out of the first and last name

Dim fname As String, lname As String, usname As Variant, cell As Range, lstcl As Variant, cname As Variant


With UserForm6

fname = UCase(Left(.txtFirst, 1)) & Mid(.txtFirst, 2)
lname = UCase(Left(.txtLast, 1)) & Mid(.txtLast, 2)
usname = fname & " " & lname

'add an if condition to ensure that the first and last name does not exist

    With ThisWorkbook.Sheets("TopSecret")
    
    .Visible = True
    .Activate
    
    lstcl = .Range("H10000").End(xlUp).row
    
        With .Range("H2:H" & lstcl)
        
            Set cname = .Find(lname)
            
        End With
        
            If cname Is Nothing Then
                        
                
                    .Range("H" & lstcl).Offset(1, 0) = lname
                    .Range("H" & lstcl).Offset(1, -1) = fname
                    
                    
                    'stitch together username and add it to the roster
    
                    
                    'put it in lower case for potential email purposes
                    
                    usname = LCase(fname) & "." & LCase(lname)
                    
                    .Range("H" & lstcl).Offset(1, 1) = usname
                    
            Else
            
                MsgBox "Driver already exists!", vbInformation
                
                
            End If
        
    .Visible = xlVeryHidden
    End With
    

 
End With


With ThisWorkbook.Sheets("TopSecret")

    .Visible = True
    .Activate
    
    .Range("C1") = "SAP Clerks"
    .Range("G1") = "Drivers"
    
    .Visible = xlVeryHidden

End With

Application.Run "Module11.make_roster"

End Sub
Private Sub make_roster()

Application.ScreenUpdating = False

Dim arr_clerks() As Variant, arr_drivers() As Variant, n As Integer, lstcl As Variant


With UserForm6
    
    
    
    'if clerks are checked
    

    If .optClerks.Value = True Then
    
    .cboRoster.Clear
        
        With ThisWorkbook.Sheets("TopSecret")
        
        .Visible = True
        .Activate
        
            lstcl = .Range("C10000").End(xlUp).row
            
            If Not IsEmpty(.Range("C2")) Then
            
                For n = 2 To lstcl
                
                    ReDim Preserve arr_clerks(n)
                    arr_clerks(n - 2) = .Range("C" & n) & " " & .Range("C" & n).Offset(0, 1)
                                        
                Next
                
                
            UserForm6.cboRoster.List = arr_clerks()
            
            End If
            
            .Visible = xlVeryHidden
            
        End With
            

    End If
    
    
    'if drivers are checked
    
    
    If .optDrivers.Value = True Then
    
    .cboRoster.Clear
        
        With ThisWorkbook.Sheets("TopSecret")
        
            .Visible = True
            .Activate
        
            lstcl = .Range("G10000").End(xlUp).row
            
            For n = 2 To lstcl
            
                ReDim Preserve arr_drivers(n)
                arr_drivers(n - 2) = .Range("G" & n) & " " & .Range("G" & n).Offset(0, 1)
                UserForm6.cboRoster.List = arr_drivers()

            Next
            
            .Visible = xlVeryHidden
                      
        End With
    
    End If
   
    
End With
End Sub

Private Sub remove_person()

Application.ScreenUpdating = False

'this routine will allow for the removal of people from the list of employees

Dim cell As Range, lstcl As Variant, i As Integer

'deleting a clerk

With UserForm6

If .optClerks.Value = True Then


    With ThisWorkbook.Sheets("TopSecret")
    
    .Visible = True
    .Activate
    
        lstcl = .Range("C1000").End(xlUp).row
        i = UserForm6.cboRoster.ListIndex
        
            For Each cell In .Range("C2:C" & lstcl)
            
                    'prevent the remove function from engaging if no selecitotn is made in the combo box
                    If i <> -1 Then
                
                    If cell & " " & cell.Offset(0, 1) = UserForm6.cboRoster.List(i) Then
                    
                        Dim rng As Range
                        Set rng = Range(Range(cell.Address), Range(cell.Offset(0, 2).Address))

                        rng.Select
                        Selection.Delete Shift:=xlUp
                        
                    End If
                
                End If
                
            Next
            
    .Visible = xlVeryHidden
            
    End With
    
End If
        
If .optDrivers.Value = True Then
 
    With ThisWorkbook.Sheets("TopSecret")
    
    .Visible = Hidden
    
        lstcl = .Range("G10000").End(xlUp).row
        i = UserForm6.cboRoster.ListIndex
        
            For Each cell In .Range("G2:G" & lstcl)
            
                'prevent the remove function from engaging if no selecitotn is made in the combo box

                If i <> -1 Then
            
                    If cell & " " & cell.Offset(0, 1) = UserForm6.cboRoster.List(i) Then
                    
                        Set rng = Range(Range(cell.Address), Range(cell.Offset(0, 2).Address))

                        rng.Select
                        Selection.Delete Shift:=xlUp
                        
                        
                    End If
                    
                End If
    
            Next
            
    .Visible = xlVeryHidden
    
    End With

End If

End With

Application.Run "Module1.make_roster"

End Sub

Private Sub change_person()

Application.ScreenUpdating = False

'this routine will allow edits in the name of the person

Dim cell As Range, lstcl As Variant, i As Integer, old As Variant, uname As Variant, new_name As Variant, new_uname As Variant



With UserForm6

If .optClerks.Value = True Then

    With ThisWorkbook.Sheets("TopSecret")
    
    .Visible = True
    .Activate
    
        lstcl = .Range("C10000").End(xlUp).row
        i = UserForm6.cboRoster.ListIndex
        
        For Each cell In .Range("C2:C" & lstcl)
        
            'store the name to be changed if it is not erased, and delete the contents from the relevant cells
        
            If cell & " " & cell.Offset(0, 1) = UserForm6.cboRoster.List(i) Then
            
                Dim old_f As String, old_l As String
            
                old = cell & " " & cell.Offset(0, 1)
                old_f = cell
                old_l = cell.Offset(0, 1)
                uname = cell.Offset(0, 2)
                
                Dim rng As Range
                Set rng = Range(Range(cell.Address), Range(cell.Offset(0, 2).Address))
                rng.Select
                rng.Clear
                
                With UserForm6
                
                    'assign the changed name to a new set of variables
                
                    If Not IsEmpty(.txtFirst) And Not IsEmpty(.txtLast) Then

                        new_name = UCase(Left(.txtFirst, 1)) & Mid(.txtFirst, 2) & " " & UCase(Left(.txtLast, 1)) & Mid(.txtLast, 2)
               
                        new_uname = LCase(.txtFirst) & "." & LCase(.txtLast)
                        
                        'compare the old and new names
                        
                        If Not new_name = old Then
                        
                            cell = UCase(Left(.txtFirst, 1)) & Mid(.txtFirst, 2)
                            cell.Offset(0, 1) = UCase(Left(.txtLast, 1)) & Mid(.txtLast, 2)
                            cell.Offset(0, 2) = new_uname
                            
                        Else
                            
                        'put in the old information back
                        
                            cell = old_f
                            cell.Offset(0, 1) = old_l
                            cell.Offset(0, 2) = uname
                            
                        End If
                        
                    End If
                    
                End With
                
                
            End If
            
        Next
        
    .Visible = xlVeryHidden
        
    End With
    
End If

If .optDrivers.Value = True Then

    With ThisWorkbook.Sheets("TopSecret")
    
    .Visible = True
    .Activate
    
        lstcl = .Range("G10000").End(xlUp).row
        i = UserForm6.cboRoster.ListIndex
        
        For Each cell In .Range("G2:G" & lstcl)
        
            'store the name to be changed if it is not erased, and delete the contents from the relevant cells
        
            If cell & " " & cell.Offset(0, 1) = UserForm6.cboRoster.List(i) Then
            
           
                old = cell & " " & cell.Offset(0, 1)
                old_f = cell
                old_l = cell.Offset(0, 1)
                uname = cell.Offset(0, 2)
                
                Set rng = Range(Range(cell.Address), Range(cell.Offset(0, 2).Address))
                rng.Select
                rng.Clear
                
                With UserForm6
                
                    'assign the changed name to a new set of variables
                
                    If Not IsEmpty(.txtFirst) And Not IsEmpty(.txtLast) Then
                    
                        new_name = UCase(Left(.txtFirst, 1)) & Mid(.txtFirst, 2) & " " & UCase(Left(.txtLast, 1)) & Mid(.txtLast, 2)
                        
                        new_uname = LCase(.txtFirst) & "." & LCase(.txtLast)
                        
                        'compare the old and new names
                        
                        If Not new_name = old Then
                        
                            cell = UCase(Left(.txtFirst, 1)) & Mid(.txtFirst, 2)
                            cell.Offset(0, 1) = UCase(Left(.txtLast, 1)) & Mid(.txtLast, 2)
                            cell.Offset(0, 2) = new_uname
                            
                        Else
                            
                        'put in the old information back
                        
                            cell = old_f
                            cell.Offset(0, 1) = old_l
                            cell.Offset(0, 2) = uname
                            
                        End If
                        
                    End If
                    
                End With
                
                
            End If
            
        Next
        
    .Visible = xlVeryHidden
        
    End With
    
End If

End With
    
                            
Application.Run "Module1.make_roster"
        
End Sub
Private Sub new_profile()

Application.ScreenUpdating = False

'this routine is going to guide the person through forming a new profile, either a driver or SAP clerk


Dim uname As Variant, pword1 As Variant, pword2 As Variant

With UserForm5

uname = .txtUName
pword1 = .txtPword1
pword2 = .txtPword2

End With

'ensure that passwords are the same, before populating the proper cells with the proper info

With ThisWorkbook.Sheets("TopSecret")

.Visible = True
.Activate

    If pword1 = pword2 Then
    
        .Range("A1") = uname
        .Range("A100") = pword1
        
        
        
            If UserForm5.optDriver.Value = True Then
            
                .Range("A3") = UserForm5.optDriver.Caption
                
            End If
            
            If UserForm5.optSAPClerk.Value = True Then
            
                .Range("A3") = UserForm5.optSAPClerk.Caption
                
            End If
    
    End If
    
        .Rows("100").EntireRow.Hidden = True
    
.Visible = xlVeryHidden

End With

MsgBox "Profile created!"

'disable create button once profile is generated

UserForm5.cmdCreate.Enabled = False



End Sub

Private Sub confirm_profile_deets()

Application.ScreenUpdating = False

With UserForm5

If Not IsEmpty(.txtUName) And Not IsEmpty(.txtPword1) And Not IsEmpty(.txtPword2) Then

.lblUname = .txtUName

    If .optDriver.Value = True Then
    
        .lblProfile = .optDriver.Caption
        
    Else
    
        MsgBox "Please select your profile"
        Exit Sub
        
    End If
    
    If .optSAPClerk.Value = True Then
    
        .lblProfile = .optSAPClerk.Caption
        
    Else
    
        MsgBox "Please select your profile"
        Exit Sub
        
    End If
    
End If

End With
End Sub

Public Sub login()

'this program is going to manage driver's login and email updated file to the clerk on duty
Application.ScreenUpdating = False


With UserForm2

    username = LCase(.txtEmail)
    ThisWorkbook.Sheets(1).Range("E1") = username
    pw1 = .txtPW1
    pw2 = .txtPW2
    
    
    'ensure there is no empty fields
    
    If IsEmpty(.txtEmail) Or IsEmpty(.txtPW1) Or IsEmpty(.txtPW2) Then
    
        MsgBox "Please enter missing details!"
        .txtPW1 = ""
        .txtPW2 = ""
    
        Exit Sub
    
    End If
                
    Dim cname As Variant, lstcl As Variant, rng As Range
    
    If .optSAPClerk.Value = True Then
    
        With ThisWorkbook.Sheets("TopSecret")
        
            .Visible = True
            .Activate
            
            lstcl = .Range("E10000").End(xlUp).row
            Set rng = ThisWorkbook.Sheets("TopSecret").Range("E2:E" & lstcl)
            
            .Visible = xlveyrhidden
        
        End With
    
    End If
    
    If .optDriver.Value = True Then
    
        With ThisWorkbook.Sheets("TopSecret")
        
            .Visible = True
            .Activate
        
            lstcl = .Range("I10000").End(xlUp).row
            Set rng = .Range("I2:I" & lstcl)
            
            .Visible = xlVeryHidden
            
        End With
    
    End If
    
    If .optSAPClerk.Value = False And .optDriver.Value = False Then
    
        MsgBox "You must select a profile option to login!", vbInformation
        UserForm2.Show
        
    End If
    
    
    
    'find the username in the database and see if it exists, quit the program if it does not
    
    With ThisWorkbook.Sheets("TopSecret")
    
        .Visible = True
        .Activate
    
        'lstcl = .Range("C10000").End(xlUp).Row
        
            'preventing an error with negative ranges from being created
        If lstcl >= 2 Then
            rng.Select
        
            With rng 'Range("C2:C" & lstcl)
            
                Set cname = .Find(username)
                
            End With
                
        End If
        
        'created by way of the cname variable, which throws an error either with Is Nothing or = "" options
        
        Dim i As Integer
        i = UserForm1.cboClerk.ListIndex
        
        'prevent an error with a negative option in the combo box
        If i > -1 Then
            
            If UserForm1.cboClerk.List(i) = "" Or IsEmpty(cname) Then
            
                MsgBox "This profile does not exist!", vbExclamation
                
                Exit Sub
                
            End If
        
        End If
        
    .Visible = xlveryhidedn
        
    End With
    
    'ensure there are no empty fields and that passwords match
    
    'if passwords match, driver can proceed to send emaisl to the clerk on duty
    
    If Not IsEmpty(.txtEmail) And Not IsEmpty(.txtPW1) And Not IsEmpty(.txtPW2) Then
    
        If pw1 = pw2 Then
                                          
            With ThisWorkbook.Sheets("TopSecret")
            
                If Not cname Is Nothing And username = cname Then
            
                    If username = .Range("A1") Then
                    
                        If pw1 = .Range("A100") Then
                        
                        
                            MsgBox "You can proceed to choose the clerk on duty"
                            
                            With ThisWorkbook
                            
                                .Sheets("Staging").Visible = True
                                .Sheets("STOCK to STOCK").Visible = True
                                .Sheets("Print").Visible = True
                                .Sheets("Archive").Visible = True
                                '.Sheets(5).Visible = True
                                .Sheets("STOCK to STOCK").Activate
                                

                            
                            End With
                            
                        Else
                        
                            MsgBox "Please re-enter passwords, they do not match."
                            
                            With UserForm2
                                .Show
                                .txtPW1 = ""
                                .txtPW2 = ""
                            End With
                            
                        End If
                    
                    End If
                    
                End If
            
            End With
            
        Else
        
            MsgBox "Password not correct! Please re-enter password.", vbExclamation
            .txtPW1 = ""
            .txtPW2 = ""
    
    
        End If
        
    End If

End With

Application.Run "Module1.clerk_list"
End Sub

Private Sub logout()

Application.ScreenUpdating = False

'this routine is going to log the user off the system and close access to the sheets

With ThisWorkbook

.Save
.Sheets("Sheet1").Visible = True
.Sheets("STOCK to STOCK").Visible = xlVeryHidden
.Sheets("Print").Visible = xlVeryHidden
.Sheets("Archive").Visible = xlVeryHidden
.Sheets("TopSecret").Visible = xlVeryHidden
.Sheets("Sheet1").Activate


MsgBox "Your work has been saved and you can safely exit MS Excel.", vbInformation

End With

End Sub

Private Sub change_password()

Application.ScreenUpdating = False

'this sub will allow the user to change their password, in case they forgot it

Dim newp1 As Variant, newp2 As Variant, uname As Variant

'ensure that the entered username matches one of the ones on the list - do admin tools first

With UserForm7

    'ensuring that both textboxes have stuff them and their values match
    
    .lblUname = UserForm2.txtEmail
    uname = .lblUname
    
    If Not IsEmpty(.txtNewP1) And Not IsEmpty(.txtNewP2) Then
    
        newp1 = .txtNewP1
        newp2 = .txtNewP2
        
        'see if passwords match
        If newp1 = newp2 Then
        
            'see if username exists in database
            Dim lstcl As Variant, cell As Range
            
            With ThisWorkbook.Sheets("TopSecret")
            
                .Visible = True
                
                
                lstcl = .Range("E10000").End(xlUp).row
                
                    For Each cell In .Range("E2:E" & lstcl)
                    
                        'if found in the list
                        If cell = uname Then
                        
                            'if profile matches
                            If .Range("A1") = uname Then
                            
                                'change password
                                .Range("A100") = newp1
                                Rows("100").EntireRow.Hidden = True
                            End If
                            
                        Else
                        
                            'in case user name is not found
                            MsgBox "User name is misspelled or does not exist. Please check again.", vbExclamation
                            UserForm7.Hide
                            Exit Sub
                            
                        End If
                    Next
                    
                .Visible = xlVeryHidden
            End With
        End If
    End If
    .Hide
End With


End Sub

Private Sub archiving()

Application.ScreenUpdating = False

Dim lstcl As Variant, cell As Range, dat As Date, dat2 As Long, diff As Double
Dim r1 As Range, r2 As Range, rng_copy As Range
Dim lstU As Range

'this sub is going to enable archiving in the following ways:

'archiving is supposed to archive details that are 90 days or older in the Staging sheet; _
it should also be able to keep archives for 180 days and have an option to wipe anything older than 90 days _
another thought: should the option to wipe it be available to the clerks, or happen automatically? _
use the developed protoytpe to evaluat the differences

'technical approach: _
1) make a data variable, get the first ten digits of each timestamp, convert to long, subtract 90 and if equal or greater _
2) take the range, copy it to the next available row in the archive sheet _
    and delete the original row to prevent duplicate transfers in the future _
    and then insert blank rows to space out orders
    
With ThisWorkbook.Sheets("Staging")

lstcl = .Range("U100000").End(xlUp).row
 
    For Each cell In .Range("U4:U" & lstcl)
    
        If Not cell = "" Then
        
            'get the date out of the left side of the cell, convert it to a number and subtract from the current date
            dat = Left(cell, 10)
            diff = CLng(dat) - CLng(dat)
            
            'If the difference is greater than 90, copy to archive and erase row
            
            If diff > 90 Then
            
            
                'build a range to copy
                Set r1 = cell.Offset(0, -20)
                Set r2 = cell.Offset(0, 2)
                Set rng_copy = Range(r1, r2)
                
                'find a place in Archive to copy to
                With ThisWorkbook.Sheets("Archive")
                
                    lstU = .Range("U1000000").End(xlUp).row
                    .Range("A" & lstU + 1) = rng_copy
                    
                End With
                
            End If
            
        End If
        
    Next
    
End With

'insert blank rows between the copied rows based on the orders

With ThisWorkbook.Sheets("Archive")

    lstB = .Range("B100000").End(xlUp).row
    
    For Each cell2 In .Range("B4:B" & lstB)
    
        If Not IsEmpty(cell2) Then
        
            If Not IsEmpty(cell2.Offset(-1, -1)) Then
            
                cell2.EntireRow.Insert
                
            End If
        
        End If
        
    Next
    
End With

End Sub



Private Sub check_movt_inputs()


With UserForm9

    'ensure FROM location begins with a letter

    If Not IsNumeric(Left(.txtFrom, 1)) Then
        'check location input for a length of 5 or 6 characters, depending if locations are ZA, ZB or ZC + number
        If Len(.txtFrom.Value) = 5 Or Len(.txtFrom.Value) = 6 Or Len(.txtFrom.Value) = 2 Or Left(.txtFrom, 1) = "L" Then
    
        .txtTo.Enabled = True
        .txtTo.SetFocus
        
    Else
    
        MsgBox "FROM Location must begin with a letter, not a number, and be 5 characters long.", vbExclamation
        .txtFrom = ""
        
        End If
    End If
    
    
    


    'ensure that the PO textbox is only numeric and has 10 digits that start with 5
    
    If IsNumeric(.txtPO) And Len(.txtPO.Value) = 10 And Left(.txtPO, 1) = 5 And Not IsEmpty(.txtPO) Then
    
        .txtSKU.Enabled = True
        .txtSKU.SetFocus
        
    ElseIf IsEmpty(.lblActivePO) Then
    
        MsgBox "Production order must be numeric, 10 digits and start with 5.", vbExclamation
        .txtSKU = ""
    
    End If
    
    
    'ensure quantity is numeric

    If IsNumeric(.txtQty) And Not IsEmpty(.txtQty) Then
    
        .txtFrom.Enabled = True
        .txtFrom.SetFocus
        
    Else
    
        MsgBox "Quantity must be a numeric value!", vbExclamation
        .txtQty = ""
        
    End If
    
    
    'ensure the SAP# is begins with 300, is 9 digits long and is numeric


    If IsNumeric(.txtSKU) And Len(.txtSKU.Value) = 9 And Left(.txtSKU, 3) = 300 And Not IsEmpty(.txtSKU) Then
    
        .txtQty.Enabled = True
        .txtQty.SetFocus
        
    Else
    
        MsgBox "Material SAP SKU number must be numeric, 9 digits and beginning with 300", vbExclamation
        .txtSKU = ""
        
    End If
    
    
    
    'ensure that the TO location begins with a letter

    If IsNumeric(Left(.txtTo, 1)) Then
        If Not Len(.txtTo.Value) = 5 Or Not Len(.txtFrom.Value) = 6 Then
    
        MsgBox "TO Location must begin with a letter, not a number and be 5 characters long.", vbExclamation
        .txtTo = ""
        
        End If
    End If
    
End With

If UserForm9.optAdjust.Value = True And UserForm9.optStaging.Value = True Then

repeating_orders
secondary_movts_staging

Else

repeating_orders

End If

End Sub
Private Sub secondary_movts_staging()

'this sub is going to compensate for movements added on to or returned from the line well afrer the order is started or finished

'return movements will be treated like stock to stock movements

'staged movements will be stored in an array and posted in a listbox to remind drivers that they need to be finished before posted to the clerks

'handling the Move Me section, unfihished movements may well be a second sub

Dim cell As Range, cell2 As Range, lstcl As Variant, lstcl2 As Variant, rng As Range, dat As Date, fnd As Double
Dim PO As Double, fndPO As Range

dat = Format(Date, "MMMM-dd-yyyy")

'ensure that order does exist and is closed

With ThisWorkbook.Sheets("Staging")

        'create driver's initials from the username in E1

        With .Range("E1")
        
        Dim DI As String, name As String, i1 As String, i2 As String, counter As Integer
        
        name = .Value
        
        i1 = UCase(Left(name, 1))
        
        For counter = 1 To Len(name)
        
            If Mid(name, counter, 1) = "." Then
            
                i2 = UCase(Mid(name, counter + 1, 1))
                Exit For
                
            End If
            
        Next
        
        DI = i1 & i2
        
        End With

    lstcl = .Range("B10000").End(xlUp).row
    Set rng = .Range(Range("B4").Address, Range("B" & lstcl).Address)
    rng.Select
    
    With UserForm9
        If Not .txtPO = "" Then
            PO = UserForm9.txtPO.Value
        End If
    End With
    
    With rng
    
        'if the order exists and the cell is green, meaning order is fully returned
        Set fndPO = .Find(PO)
        
        If Not fndPO Is Nothing Then 'And cell.Interior.Color = vbGreen Then
        
            MsgBox "Order found!", vbInformation
            
            'find the last cell with a material number in it
            
            lstcl2 = .Range("C10000").End(xlUp).row
            
            'set the entry point for the adjusted order
            Set cell2 = Cells(lstcl2 + 2, 2)
            cell2.Activate
            
            'populate row with relevant data and populate the returns accordingly

            With UserForm9
            
                'populate staging info
            
                'insert date
                cell2.Offset(0, -1) = Format(Date, "MMMM-dd-yyyy")
                'insert PO
                cell2 = "ADJ " & .lblActivePO
                'insert material
                cell2.Offset(0, 1) = UCase(.txtSKU)
                'insert FROM loc
                cell2.Offset(0, 2) = UCase(.txtFrom)
                'insert TO Loc
                cell2.Offset(0, 3) = UCase(.txtTo)
                'insert quantity
                cell2.Offset(0, 4) = .txtQty
                'populate stager
                cell2.Offset(0, 6) = DI
                
                'populate return info
                'extra staged movements will not be returned, returns will be handled with the standard PO movts
                
                'return date
                cell2.Offset(0, 9) = "NO RTN"
                'return PO
                cell2.Offset(0, 10) = "NO RTN"
                'return SKU
                cell2.Offset(0, 11) = "NO RTN"
                'return from location
                cell2.Offset(0, 12) = "NO RTN"
                'return to location
                cell2.Offset(0, 13) = "NO RTN"
                'stager initials
                cell2.Offset(0, 16) = DI
                'insert time stamp
                cell2.Offset(0, 19) = Format(Now, "mm-dd-yyyy, hh:mm:ss")
                'insert staged code
                cell2.Offset(0, 20) = "STG"
                'fill the range with yellow to signify that it has been staged
                Range(cell2.Offset(0, -1), cell2.Offset(0, 20)).Interior.Color = vbYellow
                
                
                'populate listbox with materials
                
                're-build an array of unfinished movements after a new movement is entered
                
                Dim come_undone() As Variant, i As Integer, cell3 As Range
                
                    With ThisWorkbook.Sheets("Staging")
                    
                        lstcl = .Range("B10000").End(xlUp).row
                        i = 1
                        For Each cell3 In .Range("B4:B" & lstcl)
                        
                            If Left(cell3, 3) = "ADJ" Then
                            
                                ReDim Preserve come_undone(i)
                                come_undone(i) = "ADJ " & cell3.Offset(0, 1)
                                i = i + 1
                        
                            End If
                            
                        Next
                    
                    End With
                
                If i > 0 Then
                'Set .lstUnfinishedBusiness.List = come_undone()
               End If
                
            End With
            
        End If

    End With
            
          'in case the order was not found in the list
        If fndPO Is Nothing Then
        
            MsgBox "Order does not exist, is not posted or finalized yet. You cannot stage this movement.", vbInformation
        
        End If
    
End With

'autofit columns

Columns("A:AA").AutoFit

'clear input textboxes

With UserForm9

    If .chkSame.Value = False Then
    .txtSKU = ""
    End If
    
    If .chkLockTAG.Value = False Then
    .txtQty = ""
    End If

    If .chkLockSAP.Value = False Then
    .txtSAPQ = ""
    End If
    
    .txtFrom = ""
    
    'account for the Lock feature to prevent erasing the destination for the movement
    
    If .chkLock.Value = False Then
        .txtTo = ""
    End If
    
End With

End Sub

Private Sub STOCK_rtn_movts()

'this routine is going to handle the posting of return movements and attach a time stamp for them
'the time stamp approach developed here will be used to update the main timestamping application to _
'reflect the return, not the staging movements

'REVISE: the point is to post orders in the sequence they are staged; return movements don't need a timestamp

'two types of returns

'one type is a stock to stock movement without a staged movement - handled within the return movement optoin

'movement will be populated in the return movement section without anything in the staged section

'plan is: planned through the staging section, populated in the returns section!!

Dim cell As Range, lstcl As Variant

With ThisWorkbook.Sheets("Staging")

        'create driver's initials from the username in E1

        With .Range("E1")
        
        Dim DI As String, name As String, i1 As String, i2 As String, counter As Integer
        
        name = .Value
        
        i1 = UCase(Left(name, 1))
        
        For counter = 1 To Len(name)
        
            If Mid(name, counter, 1) = "." Then
            
                i2 = UCase(Mid(name, counter + 1, 1))
                Exit For
                
            End If
            
        Next
        
        DI = i1 & i2
        
        End With

    lstcl = .Range("C10000").End(xlUp).row
        
    Set cell = .Range("C" & lstcl + 2)
    cell.Activate
    
       
        With UserForm9
        
            'put today's date in the return column
            cell.Offset(0, -2) = Format(Date, "MMMM-dd-yyyy")
            'everything from the PO to the SAP clerk field remains blank
            Range(cell.Offset(0, -1), cell.Offset(0, 5)) = "STK MVNT"
            
            'STK movement date
            cell.Offset(0, 8) = Format(Date, "MMMM-dd-yyyy")
            'PO in the return section remains blank
            cell.Offset(0, 9) = "STK MOVNT"
            'put in the material number
            cell.Offset(0, 10) = .txtSKU
            'put the location FROM
            cell.Offset(0, 11) = UCase(.txtFrom)
            'put the location TO
            cell.Offset(0, 12) = UCase(.txtTo)
            'put in the quantity
            cell.Offset(0, 13) = UCase(.txtQty)
            'put in the stager initials
            cell.Offset(0, 15) = DI
            'put in time stamp
            cell.Offset(0, 18) = Format(Now, "mm-dd-yyyy, hh:mm:ss")
            'insert RTN code
            cell.Offset(0, 19) = "RTN"
            
            'paint the range green to signify a completed movement
            Range(cell.Offset(0, -2), cell.Offset(0, 19)).Interior.Color = vbGreen

        End With
        

End With

Columns("A:AA").AutoFit

'clean up textboxes after doing the movement
With UserForm9

    If .chkSame.Value = False Then
    
        .txtSKU = ""
        
    End If
    
    .txtFrom = ""
    If .chkLock = False Then
    .txtTo = ""
    End If
    
    If .chkLockTAG = False Then
    .txtQty = ""
    End If
    
    If .chkLockSAP.Value = False Then
    .txtSAPQ = ""
    End If
    
End With

End Sub

Private Sub opt_adjjj()

'this routine is going to make a quick list of the option adjusting staged movements and fill the combo box
'it also works for regular staging

'must make it possible to enter the associated movements depending on the order picked in the combo box

Dim cell As Range, lstcl As Variant, arr_adj() As Variant, i As Integer, n As Integer, k As Integer

Erase arr_adj()
i = 0

If UserForm9.optAdjust.Value = True Then

UserForm9.cboOOs.Clear
UserForm9.lstUnfinishedBusiness.Clear

With ThisWorkbook.Sheets("Staging")
    
    lstcl = .Range("B10000").End(xlUp).row
    

    For Each cell In .Range("B4:B" & lstcl)
    
        If Not cell.Interior.Color = vbGreen Then
    
            If Left(cell, 3) = "ADJ" Then
        
            'check that any one of the requisite return quantities are filled in, before making the array
  
                If IsEmpty(cell.Offset(0, 13)) Or IsEmpty(cell.Offset(0, 14)) Or IsEmpty(cell.Offset(0, 15)) Then
                
                
                        ReDim Preserve arr_adj(i)
                        arr_adj(i) = cell
                        i = i + 1
                        
                End If
                
            End If
            
        End If
        
    Next
    
End With

If i > 0 Then
UserForm9.cboOOs.List = arr_adj()
End If

End If

'make a list of open orders and populate the regular combo box

Dim cell2 As Range, arr_reg() As Variant, j As Integer

j = 0
Erase arr_reg()

If UserForm9.optReg.Value = True Or UserForm9.optAdjust.Value = True Then

    With ThisWorkbook.Sheets("Staging")
    
    lstcl = .Range("B10000").End(xlUp).row
    
        For Each cell In .Range("B4:B" & lstcl)
        
            If Not cell.Interior.Color = vbGreen Then
        
                If Not Left(cell, 3) = "ADJ" And Not IsEmpty(cell) Then
                
                    ReDim Preserve arr_reg(j)
                    arr_reg(j) = cell
                    j = j + 1
                    
                End If
            
            End If
            
        Next
        
    End With
    
    If j > 0 Then
    UserForm9.cboOPO.List = arr_reg()
    End If
End If

End Sub



Private Sub cboOO_selections()

'this routine is going to retrieve the materials for the adjusted orders

Dim cell As Range, lstcl As Variant, i As Integer, sel As Variant
Dim arr_mats() As Variant, j As Integer

With UserForm9

    '.cboOOs.Clear

    i = .cboOOs.ListIndex
    If i > -1 Then
    sel = .cboOOs.List(i)
    End If
    
End With

With ThisWorkbook.Sheets("Staging")

    lstcl = .Range("B10000").End(xlUp).row
    
    For Each cell In .Range("B4:B" & lstcl)
    
        If Left(cell, 3) = "ADJ" Then
            
            ReDim Preserve arr_mats(j)
            arr_mats(j) = cell.Offset(0, 1)
            j = j + 1
            
        End If
        
    Next
UserForm9.lstUnfinishedBusiness.List = arr_mats()
    
End With


End Sub

Private Sub unfinished_rtns()

'the other type is an unfinished movement from the list in the listbox - handled from the second panel with the listbox

'this sub is going to take care of outstanding irregular staged movements

Dim cell2 As Range, lstcl2 As Variant, i As Integer, j As Integer, mat As Double, PO As Variant

'set the active selection in the listbox and assign the contents of the position to a variable

i = UserForm9.lstUnfinishedBusiness.ListIndex
mat = UserForm9.lstUnfinishedBusiness.List(i)
j = UserForm9.cboOOs.ListIndex
PO = UserForm9.cboOOs.List(j)

With ThisWorkbook.Sheets("Staging")

    lstcl2 = .Range("M10000").End(xlUp).row
    
    For Each cell2 In .Range("M4:M" & lstcl2)
    
        If cell2 = mat And cell2.Offset(0, -1) = PO Then
        
            With UserForm9
            
                'insert from location into user form
                .txtFrom2 = cell2.Offset(0, 1)
                'populate destination in sheet
                cell2.Offset(0, 2) = .txtTo2
                'populate return quantity in sheet
                cell2.Offset(0, 3) = .txtQty2
            
            End With
    
        End If

    Next

End With

'put regular staging and return movements in the same utility?

End Sub



Private Sub repeating_orders()

'this short routine checks that the same order number is not entered twice

Dim cell As Range, cell2 As Range, lstcl As Variant, PO As Double, fndPO As Range, rng As Range

With ThisWorkbook.Sheets("Staging")

    lstcl = .Range("B10000").End(xlUp).row

    Set rng = .Range("B4", Range("B" & lstcl).Address)
    n = 0
    
    'for newly entered movements
    
    With rng
        
        If UserForm9.optReg.Value = True Then
    
            PO = UserForm9.lblActivePO
            
            Set fndPO = .Find(PO)
    
                If Not fndPO Is Nothing Then
                
                    'exit only if the duplicating order is being newly entered
                    If Not UserForm9.txtPO = "" Then
                
                        MsgBox "You cannot enter duplicate order numbers!"
                        fndPO.Activate
                        'Rows(cell.row).EntireRow.Delete
                        Exit Sub
                        
                    End If
                    
                End If
            
        End If
    
    'For adjusted movements
    
        If UserForm9.optAdjust.Value = True Then
        
            rng.Select
        
            Dim POadj As Variant
            
            POadj = "ADJ " & UserForm9.lblActivePO
            
            Set fndPO = .Find(POadj)
        
                If Not fndPO Is Nothing Then
                
                    'only exit if the order being adjusted is attempted to be entered as a new order
                    If Not UserForm9.txtPO = "" Then
            
                        MsgBox "You cannot enter duplicate order numbers!"
                        fndPO.Activate
                        'Rows(cell.row).EntireRow.Delete
                        Exit Sub
                    
                    End If
                
                End If
            
        End If


End With

End With

'if the duplicate is found, it would exit the sub and not affect this section, so going to the staging should not be a
'problem


If Not fndPO Is Nothing And UserForm9.optAdjust.Value = False Then

    staging

End If

If fndPO Is Nothing And UserForm9.optReg.Value = True Then

    staging
    
End If


End Sub

Private Sub staging_prep()

'this app is no longer used

Application.ScreenUpdating = False

With UserForm9
    
    'run the same routine if it is a new routine or if the order exists
        If Not IsEmpty(.txtPO) And Len(.txtPO) = 10 And .lblActivePO = .txtPO Then
        
            staging
        
        
        End If
        
        If .txtPO = "" And .lblActivePO = .cboOPO Then
        
            staging
            
        
        End If
 End With
        
End Sub

        
Private Sub staging()
'this routine will allow staging and returning of regular staging movements

Dim cell As Range, dat As Variant, lstcl As Variant, tag As Double, sap As Double, rng As Range, fnd2 As Range
Dim bld_rng, r1 As Range, r2 As Range, PO2 As Double, lstcl2 As Variant, cell2 As Range
Dim arr_stgd() As Double, a As Integer, total As Double, fnd As Range ', lstcl2 As Range
Dim lstA As Variant, clA As Range


dat = Format(Date, "MMMM-dd-yyyy")

'staging a new order and an existing order

'staging a new order

With ThisWorkbook.Sheets("Staging")

        'create driver's initials from the username in E1

        With .Range("E1")
        
        Dim DI As String, name As String, i1 As String, i2 As String, counter As Integer
        
        name = .Value
        
        i1 = UCase(Left(name, 1))
        
        For counter = 1 To Len(name)
        
            If Mid(name, counter, 1) = "." Then
            
                i2 = UCase(Mid(name, counter + 1, 1))
                Exit For
                
            End If
            
        Next
        
        DI = i1 & i2
        
        End With

    lstcl = .Range("C10000").End(xlUp).row
    lstcl2 = .Range("B10000").End(xlUp).row
    
    
    Set cell = Range("B" & lstcl).Offset(1, 0)
    Set rng = Range("B4", Range("B" & lstcl).Address)
    
    With UserForm9
    
            

            'use a .find function to prevent entering the order beyond the first mat movement
            Dim PO As Double
                       
            With rng
            
            'need to build a range to insert the material in the correct spot
            
                If Not UserForm9.txtPO = "" Then
            
                    PO = UserForm9.txtPO
                    'Set cell = cell.Offset(0, -1)
                    cell = PO
                    
                Else: PO = UserForm9.lblActivePO
                
                    Set fnd2 = .Find(PO)
                
                    fnd2.Activate
                    
                    'If Not fnd2 Is Nothing Then
                    
                    'build the relevant range for the selected order
                    Dim bld_rng2 As Range, fnd3 As Range
                    
                    For Each cell3 In Range(fnd2.Offset(1, 0), "B" & lstcl2)
                    
                        'if it's the last order on the list, select the end of hte range with data
                        
                        'check that the order is not adjusted, if it matches, and that it is not the last one on the list
      
                        
                            If Not cell3 = "ADJ" & Range("B" & lstcl2) And Not Left(Range("B" & lstcl2), 3) = "ADJ" _
                            And fnd2 = Range("B" & lstcl2) Then
                            
                                'MsgBox Range("B" & lstcl2)
                            
                                Set bld_rng2 = Range(fnd2.Offset(0, 1), "B" & lstcl)
                                bld_rng2.Select
                                Set clA = Range("B" & lstcl).Offset(1, 0)
                                Set cell = clA
                                cell.Activate
                                Exit For
                                
                            End If
                        
          
                        cell3.Activate
                        Set cell = cell3
                        
                        If Not IsEmpty(cell3) And Not Left(cell3, 3) = "ADJ " & cell3 Then
                        
                            Set bld_rng2 = Range(fnd2.Offset(0, 1), cell3.Offset(-1, 1))
                            bld_rng2.Select
                            'Exit For
                            
                            'insert the new movement for the material at the end of the relevant range
                            With cell3
                                .Activate
                                
                                'if posting on the first line of the spreadsheet under the headings
                                
                                If .Offset(-1, 1) = "SKU Number" Then
                            
                                    lstA = Sheets("Staging").Range("A100000").End(xlUp).row + 1
                                    Set clA = Sheets("Staging").Range("A" & lstA).Offset(1, 0)
                                    
                                    clA.Activate
                                    clA.EntireRow.Insert , CopyOrigin:=xlFormatFromLeftOrAbove

                                    cell = clA.Offset(-2, 1)
                                        
                                End If
                                
                                'if posting somewhere among the POs already listed
                                If Not IsEmpty(.Value) Then
                                
                                    Set cell = cell3.Offset(-1, 0)
                                    cell.Activate
                                    cell.EntireRow.Insert , CopyOrigin:=xlFormatFromLeftOrAbove
                                    Set cell = cell3.Offset(-2, 0)
                                    
                                End If
                            
                            End With
                            
                            Exit For
                            
                        End If
                        
                    Next
                
                End If
                
                
            End With
            
            'must enter things in the appropriate place, depending on if the order exist or if it is new
            
            'if order is new
            
            lstcl = ThisWorkbook.Sheets("Staging").Range("C10000").End(xlUp).row

            If Not IsEmpty(.txtPO) And Len(.txtPO) = 10 Then
        
                Set cell = Range("B" & lstcl + 1)
                Set r2 = Range("B" & lstcl + 2) 'if there is a new order, leave a blank row btw it and the prev PO
                cell.Activate
        
            End If

        
            'insert the date
            cell.Offset(0, -1) = dat
            'insert the material
            cell.Offset(0, 1) = .txtSKU
            'insert the from location
            cell.Offset(0, 2) = UCase(.txtFrom)
            'insert to location
            cell.Offset(0, 3) = UCase(.txtTo)
            'insert quantity
            cell.Offset(0, 4) = .txtQty
            'insert SAP LS24 quantity
            cell.Offset(0, 5) = .txtSAPQ
            'insert driver's initials
            cell.Offset(0, 6) = DI
            'insert the all important timestamp
            cell.Offset(0, 19) = Format(Now, "mm-dd-yyyy, hh:mm:ss")
            'insert staging code
            cell.Offset(0, 20) = "STG"
            'colour the range yellow to signify that the movement is staged
            Range(cell.Offset(0, -1), cell.Offset(0, 20)).Interior.Color = vbYellow
            
            'remove green filling from previous orders that may have been finalized
            Rows(cell.row).Interior.Color = xlNone
            
            
            'DO RETURN Transfers and put materials there only if they are not already there
            'second movements will be treated like a stock to stock movement
            
            
            Dim r_mat As Range, xRng As Range
            
            If fnd2 Is Nothing Then
            
            cell.Offset(0, 11).Activate
            
                'if order does not exist, set the range of materials as its own row at the end of the list
            
                Set xRng = Range(cell.Offset(0, 10), cell.Offset(0, 10))
                
            Else
                
                With Sheets("Staging")
                
                    lstcl = .Range("C10000").End(xlUp).row
                    'Set cell3 = .Range("B" & lstcl)
                    Set xRng = Range(fnd2.Offset(0, 11), cell.Offset(0, 11))
                    
                End With
                
                xRng.Select
                
            End If
            
            'revise how return materials are entered, maybe do  a find function
            'DONE
            
            With xRng
            
                Set r_mat = .Find(UserForm9.txtSKU)
            
                If r_mat Is Nothing Then
                
                    'insert date
                    cell.Offset(0, 9) = dat
                    'insert PO - no need, but OK
                    cell.Offset(0, 10).Activate
                    cell.Offset(0, 10) = UCase(UserForm9.txtPO.Value)
                    'insert material
                    cell.Offset(0, 11).Activate
                    cell.Offset(0, 11) = cell.Offset(0, 1).Value
                    'insert from location as old TO location
                    cell.Offset(0, 12) = UCase(UserForm9.txtTo)
                    
                    'remove green filling from previous orders that may have been finalized
                    Rows(cell.row).Interior.Color = xlNone
                    
                End If
            
            End With
            
            'create a running total of the material that is being staged - DO NOT FORGET TO TRANSFER SUM TO USERFORM9
            
            'here there will be a need to build a range of the active order, as per the print_stuff routine
            
                
                With ThisWorkbook.Sheets("Staging")
                
                lstcl = .Range("C10000").End(xlUp).row
                lstcl2 = .Range("B10000").End(xlUp).row
                
                Set rng = Range("B4:B" & lstcl)
                rng.Select
                
                    With rng
                        
                        'check that the order is not newly entered or already existing by checking the txtbx or label
                         If Not UserForm9.txtPO = "" Then
            
                            PO2 = UserForm9.txtPO
                            Set fnd = .Find(PO2)
                            
                            
                    
                         Else: PO2 = UserForm9.lblActivePO
                         
                            Set fnd = .Find(PO2)
                            fnd.Select
                
                         End If
                        
                        
                    End With
                        
                        'if the order is found, make a range of the materails
                        If Not fnd Is Nothing Then
                                                
                                Set r1 = fnd
                                r1.Select
                                
                                'set r2 and build the PO range
                                
                                'in case the last PO is chosen
                                For Each cell2 In Range(r1.Offset(1, 0), "B" & lstcl2)
                                    
                                    If cell2 = Range("B" & lstcl2) Then
                                
                                        Set r2 = Range("B" & lstcl)
                                        Exit For
                                        
                                    End If
                                    
                                    'if it's a PO somewhere along the list
                                    If Not IsEmpty(cell2) Then
                                    
                                        Set r2 = cell2.Offset(-1, 0)
                                        Exit For
                                        
                                    End If
                                    
                                Next
                                
                                'create a range out of the materials for the particular order
                                Set bld_rng = .Range(r1.Offset(0, 1), r2.Offset(0, 1))
                                bld_rng.Select
                            
                                'build an array and count all the instances of the material quantity staged so far
                                
                                        
                                    For Each cell In bld_rng
                                            
                                        If cell = UserForm9.txtSKU Then
                                                
                                            ReDim Preserve arr_stgd(a)
                                            arr_stgd(a) = cell.Offset(0, 3)
                                            a = a + 1
                                                    
                                        End If
                                                
                                            
                                    Next
                                    
                                    'if there is nothing to add
                                    
                                    If a = 0 Then
                                    
                                        UserForm9.lblQinfo = ActiveCell.Offset(0, 3)
                                    
                                    End If
                                    
                                    'if there is something to add
                                    
                                    If a > 0 Then
                                    
                                        total = WorksheetFunction.Sum(arr_stgd())
                                    
                                    End If
                                    
                                    'transfer relevant data to be displayed
                                    
                                    With UserForm9
                                    
                                        .lblCurrentPO = .lblActivePO
                                        .lblSKUinfo = .txtSKU
                                        
                                        If a > 0 Then
                                        .lblQinfo = total
                                        End If
                                        
                                    End With
                                        
                            
                        End If
                End With
            
        
             
    End With
End With



'xRng changes its makeup, depending on if a new or existing order is being updated
'bld_rng2 also relies on fnd2 being something, but it will be passed anyway

'must transfer cell if the same timestamp is going to be used across all staged movements
'cell3 is important, because the ranges must include newly inserted movements

If Not cell Is Nothing Then

timestamping xRng, bld_rng2, fnd2, cell, cell3

End If

If Not fnd2 Is Nothing Then

timestamping xRng, bld_rng2, fnd2, cell, cell3

End If

'things that need to happen: 1) clear inputs for next movement, 2) put breaks btw orders, 3) watch for NM, _
4) refresh cboOPO and clean out txtPO

'refresh the list of open orders with the new inputs, find and activate the last entered order in that list
opt_adjjj

With UserForm9

    'see if the textbox is not empty, then refresh opt_adjjj
    If Not IsEmpty(.txtPO) Then
    
    'recreate the list of open POs
    opt_adjjj
    
    'put the PO in a variable and erase the textbox
    
    Dim fndPO As Double
    
    If Not .txtPO = "" Then
    fndPO = .txtPO.Value
    .txtPO = ""
    
    'search the cboOPO for the PO
    Dim arr_OPO() As Double, k As Integer
    
    For i = 0 To .cboOPO.ListCount
    
        If .cboOPO.Value = fndPO Then
        
            .cboOPO.Text = fndPO
            .lblActivePO = fndPO
            
        End If
        
    Next
    
    End If
    
    'clear inputs for next movement
    
    If .chkSame.Value = False Then
    
        .txtSKU = ""
        
    End If
    
    If .chkLockTAG.Value = False Then
    .txtQty = ""
    End If
    
    If .chkLockSAP.Value = False Then
    .txtSAPQ = ""
    End If
    
    .txtFrom = ""
    
    If .chkLock = False Then
        .txtTo = ""
    End If
    
    .txtSKU.SetFocus
    
    'put breaks between orders
    
    With Sheets("Staging")
    
        lstcl2 = .Range("B10000").End(xlUp).row
        
        Dim brk As Range
        
        For Each brk In .Range("B4:B" & lstcl2)
        
            'if there is an order
            If Not IsEmpty(brk) Then
        
                'if there is not break between dates, put a break in
                
                If Not IsEmpty(brk.Offset(-1, -1)) Then
                
                    brk.EntireRow.Insert
                    Rows(brk.Offset(-1, 0).row).ClearFormats
                    
                    
                End If
                
            End If
            
        Next
        
    End With
    
    End If
    
End With
                
End Sub

Private Sub timestamping(xRng, bld_rng2, fnd2, cell, cell3)

'can just select bld_rng and timestamp it, but an insurance policy would be to rebuild the ranges and correctly timestamp them

Application.ScreenUpdating = False

Dim cell2 As Range, bld_rng As Range
Dim lstB As Variant, lstC As Variant



xRng.Select

'hedge an error, where the

If Not fnd2 Is Nothing Then

With ThisWorkbook.Sheets("Staging")

    lstB = .Range("B100000").End(xlUp).row
    lstC = .Range("C100000").End(xlUp).row
    
    Range(fnd2.Offset(1, 0), "B" & lstB).Select
    
    For Each cell2 In Range(fnd2.Offset(1, 0), "B" & lstB)
    
    cell2.Select
    
        'if it is the last PO in the range
        If fnd2 = Range("B" & lstB) Then
        
            Set cell2 = Range("B" & lstC)
            Set bld_rng = Range(fnd2, "B" & lstC)
            Exit For
            
        End If
        
        'if the PO is somewhere in the list
        
        If Not IsEmpty(cell2) Then
        
            'set the bottom extent of the range to the last material, so no timestamp is put in the blank row
            Set cell2 = cell2.Offset(-2, 0)
            'build a range of the materials
            Set bld_rng = Range(fnd2.Offset(0, 1), cell2.Offset(0, 1))
            Exit For
            
        End If
        
    Next
    
    'select the range
    bld_rng.Select
    
    'set the timestamp valie
    timestamp = fnd2.Offset(0, 19)
    
    For Each cell3 In bld_rng
    
        cell3.Offset(0, 19) = timestamp

    Next


End With

End If


End Sub



Private Sub reg_open_orders()

'this routine and the ones after it are going to facilitate the returns for regularly staged production orders

'make a list of open production orders and populate the combo box cboOOs

Dim cell As Range, lstcl As Variant, arr_OO() As Double, i As Integer

With ThisWorkbook.Sheets("Staging")

    'criteria for open orders: not adjusted movements, not empty and not coloured in green to ensure that there are _
    'open orders still to be done.

    lstcl = .Range("B10000").End(xlUp).row
    
    For Each cell In .Range("B4:B" & lstcl)
    
        If Not IsEmpty(cell) And Not cell.Interior.Color = vbGreen Then
    
        If Not Left(cell, 3) = "ADJ" And IsNumeric(cell) Then
        
            ReDim Preserve arr_OO(i)
            arr_OO(i) = cell
            i = i + 1

        End If
        
        End If

    Next
    
    If i > 0 Then
    UserForm9.cboOOs.List = arr_OO()
    End If
    

End With

End Sub

Private Sub reg_return()

'this routine is going to display the returns that need to be made and finalize them on driver input

'find the order and build a range around it

Dim PO As Variant, r1 As Range, r2 As Range, lstcl As Variant, rng As Range, fnd_PO As Range
Dim cell As Range, bld_rng As Range
Dim i As Integer, j As Integer 'for the choices in the combo box and the listbox

With ThisWorkbook.Sheets("Staging")

    lstcl = .Range("C10000").End(xlUp).row
    
    With UserForm9
    
        'set the choice of open order
        i = .cboOOs.ListIndex
        If i > -1 Then
        PO = .cboOOs.Value
        End If
        
        'build a range out of the order list and find the order
        Set rng = Range("B4:B" & lstcl + 1)
        
        With rng
        
            Set fnd_PO = .Find(PO)
            
        End With
        
        'fnd_PO will be the top of the range, find the bottom as the first empty cell from the fnd_PO to the last cell with data
        
        For Each cell In Range(fnd_PO.Offset(0, 1), "C" & lstcl + 1)
        
            If cell = 0 Then
            
                'build the range that encompasses the chosen order
                Set r1 = fnd_PO.Offset(0, 1)
                Set r2 = cell.Offset(-1, 0)
                Set bld_rng = Range(r1, r2)
                bld_rng.Select
                Exit For
            
        
            End If

        Next
        
        'an array is going to collect all the outstanding materials based on the range built above
        
        Dim arr_rtns() As Variant, a As Integer, cell2 As Integer
        
        For Each cell In Range(r1.Offset(0, 10), r2.Offset(0, 10))
        
            If Not IsEmpty(cell) And IsEmpty(cell.Offset(0, 2)) Then
            
                'only count cells that have outstanding movements
            
                    ReDim Preserve arr_rtns(a)
                    arr_rtns(a) = cell
                    a = a + 1
                    
        
            End If
        
        Next
        
            
        'if any remaining movements remain open, populate the listbox, and if not, clear the listbox when the order is activated
        If a > 0 Then
        UserForm9.lstUnfinishedBusiness.List = arr_rtns()
        Else
        UserForm9.lstUnfinishedBusiness.Clear
        End If
        

    
    End With

End With



End Sub

Private Sub do_rtn()

'this routine is going to rebuild the range, locate the cell and populate the requisite boxes and cells
Dim PO As Variant, r1 As Range, r2 As Range, lstcl As Variant, rng As Range, fnd_PO As Range
Dim cell As Range, bld_rng As Range
Dim mat As Double, cell_mat




With ThisWorkbook.Sheets("Staging")

        'create driver's initials from the username in E1

        With .Range("E1")
        
        Dim DI As String, name As String, i1 As String, i2 As String, counter As Integer
        
        name = .Value
        
        i1 = UCase(Left(name, 1))
        
        For counter = 1 To Len(name)
        
            If Mid(name, counter, 1) = "." Then
            
                i2 = UCase(Mid(name, counter + 1, 1))
                Exit For
                
            End If
            
        Next
        
        DI = i1 & i2
        
        End With

        lstcl = .Range("C10000").End(xlUp).row
    
        With UserForm9
        
                'i = .cboOOs.ListIndex
                PO = .cboOOs.Value
                j = .lstUnfinishedBusiness.ListIndex
                mat = .lstUnfinishedBusiness.List(j)
        
            'set the choice of open order

            
            'build a range out of the order list and find the order
            Set rng = Range("B4:B" & lstcl + 1)
            
            With rng
            
                Set fnd_PO = .Find(PO)
                
            End With
            
                'fnd_PO will be the top of the range, find the bottom as the first empty cell from the fnd_PO to the last cell with data
            
                For Each cell In Range(fnd_PO.Offset(0, 1), "C" & lstcl + 1)
                
                    If cell = 0 Then
                    
                        'build the range that encompasses the chosen order
                        Set r1 = fnd_PO.Offset(0, 1)
                        Set r2 = cell.Offset(-1, 0)
                        Set bld_rng = Range(r1, r2)
                        bld_rng.Select
                        Exit For
                    
                
                    End If
    
                Next
            
                'populate the relevant info from the range into the proper boxes on the user form
            
                For Each cell_mat In Range(r1.Offset(0, 10), r2.Offset(0, 10))
                
                    If cell_mat = mat Then
                    
                        .txtFrom2 = cell_mat.Offset(0, 1)
                        .labMat = cell_mat
                    
                    End If
                
                Next
                
                'populate the relevant cells if all info is filled out
                
                Dim mat_fnd As Variant
                
                s = 1
                With Range(r1.Offset(0, 10), r2.Offset(0, 10))
                
                    Set mat_fnd = .Find(mat)
                    
                    If Not mat_fnd Is Nothing Then
                    
                        'create a sum of the materials in the staged section to compare against the return
                        
                        With Range(r1, r2)
                        
                            .Select
                            
                            Dim mat_stg As Variant, arr_stg() As Double, s As Integer, mat_sum As Double
                            
                            For Each mat_stg In Range(r1, r2)
                            
                            'annoying bit to remove quotes from mat_rng so comparison can catch
                            'set a double variable so the code can execute
                            Dim mat_stg2 As Double
                            
                            mat_stg2 = mat_stg
                            
                            'mat_stg2 = Mid(mat_stg, 1, Len(mat_stg))
                            
                                If mat_stg2 = mat_fnd.Value Then
                                
                                    ReDim Preserve arr_stg(s)
                                    arr_stg(s) = mat_stg.Offset(0, 3)
                                    s = s + 1
                                    
                                End If
                            
                            Next
                            
                            'create a sum of the staged quantities if the material is found _
                            'and it will be if it listed in the returns
                            If s > 0 Then
                            
                                mat_sum = Application.WorksheetFunction.Sum(arr_stg())
                            
                            End If
                            
                        
                        
                        End With
                    
                        With UserForm9
                        
                            'insert the return location and make the row green to signify a completed movement
                    
                            If Not .txtTo2 = "" And Not .txtQty2 = "" Then
                                    
                                    If .txtQty2.Value < mat_sum Then
    
                                        mat_fnd.Offset(0, 2) = UCase(.txtTo2)
                                        mat_fnd.Offset(0, 3) = UCase(.txtQty2)
                                        
                                        'calculate scrap value
                                        mat_fnd.Offset(0, 5) = mat_fnd.Offset(0, 4) - mat_fnd.Offset(0, 3)
                                        'insert driver's initials
                                        mat_fnd.Offset(0, 6) = DI
                                        'put the RTN code in
                                        mat_fnd.Offset(0, 9) = "RTN"
                                        'fill the range with green when the return movement is done
                                        Range(mat_fnd.Offset(0, -12), mat_fnd.Offset(0, 9)).Interior.Color = vbGreen
                                    
                                    Else
                                    
                                        MsgBox "You cannot return a quantity greater than what was staged!", vbExclamation
                                        Exit Sub
                                    
                                    End If
                                
                            End If
                                
                            If .txtTo2 = "" And .txtQty2 = "" And .chkNoRtn.Value = True Then
                                    
                                    mat_fnd.Offset(0, 2) = "NO RTN"
                                    mat_fnd.Offset(0, 3) = "NO RTN"
                                    mat_fnd.Offset(0, 5) = DI
                                    Range(mat_fnd.Offset(0, -12), mat_fnd.Offset(0, 8)).Interior.Color = vbGreen
                            
                            End If
                        End With
                        
                    End If
                    
                End With
                

           
    End With

End With

'clear up no rtn checkbox and the textboxes

With UserForm9

    .chkNoRtn.Value = False
    '.txtFrom2 = ""
    .txtTo2 = ""
    .txtQty2 = ""
    .cmdRtn.Enabled = False
    
End With

'check if order is completed and color it in green

                'scan the active range for completed movements, and if all are done, color the entire range green
                
                Dim mat_chk As Range
                
                'create a counter to see how many returns there are
                
                Dim RT_ctr As Integer, RT_chk As Integer
                
                RT_ctr = 0
                RT_chk = 0
                
                'count the number of materials to potentially return
                
                Range(r1.Offset(0, 10), r2.Offset(0, 10)).Select
                For Each mat_chk In Range(r1.Offset(0, 10), r2.Offset(0, 10))
                
                
                
                    If Not IsEmpty(mat_chk) Then
                    
                        RT_ctr = RT_ctr + 1
                        
                    End If
                    
                Next
                    
                'count how many are returned
                For Each mat_chk In Range(r1.Offset(0, 10), r2.Offset(0, 10))

                    
                    If Not IsEmpty(mat_chk.Offset(0, 2)) And Not IsEmpty(mat_chk.Offset(0, 3)) Then
                    
                        RT_chk = RT_chk + 1
                        
                    End If
                    
                Next
                
                'if the counters match in value, paint the whole order green as finalized
                
                If RT_ctr = RT_chk Then
                        
                            Range(r1.Offset(0, -2), r2.Offset(0, 16)).Interior.Color = vbGreen
                        
                End If

reg_open_orders
        
End Sub
        
Private Sub easy_login()



'this routine is going to provide an easy login that does not rely on the profile management tools above

Dim uname As Variant, SAPuname As Variant, pw1 As Variant, pw2 As Variant

'get the variable inputs from the userform11
With UserForm11

    uname = .txtEmail
    SAPuname = .txtSAP
    pw1 = .txtPW1
    pw2 = .txtPW2
    
End With


'ensure there are not empty fields, the passwords match, and populate fields



    With UserForm11
    
        If Not IsEmpty(.txtEmail) And Not IsEmpty(.txtSAP) And Not IsEmpty(.txtPW1) And Not IsEmpty(.txtPW2) Then
    
            If pw1 = pw2 Then
            
                With ThisWorkbook.Sheets("Staging")
                
                    .Range("E1") = uname
                    .Range("F1") = SAPuname
                    .txtSPW = UserForm11.txtPW1
                    
                End With
                
            Else
            
                MsgBox "Your password does not match!", vbExclamation
                .txtPW1 = ""
                .txtPW2 = ""
                
            End If
            
        End If
        
    End With
    
    'create a set of initials
    
    With ThisWorkbook.Sheets("Staging")
    
        With .Range("E1")
        
            Dim INIT As String, name As String, i1 As String, i2 As String, counter As Integer
            
            name = .Value
            
            i1 = UCase(Left(name, 1))
            
            For counter = 1 To Len(name)
            
                If Mid(name, counter, 1) = "." Then
                
                    i2 = UCase(Mid(name, counter + 1, 1))
                    Exit For
                    
                End If
                
            Next
            
            INIT = i1 & i2
        
        End With
        .Range("H1") = INIT
        
        'establish the type of profile - driver or SAP clerk
    
        If UserForm11.optDriver.Value = True Then
        
            .Range("I1") = "DRV"
            
        End If
        
        If UserForm11.optSAPClerk.Value = True Then
        
            .Range("I1") = "SAP"
            
        End If

        
    End With
    


    'login into gmail and don't worry about doing it every time an order has to be submitted
    With UserForm11
    
        If .optDriver.Value = True Then
        
            login_gmail
        
        End If
    
        If .optSAPClerk.Value = True Then
        
            SAP_login_gmail
        
        End If
    
    End With
    
Unload UserForm11

End Sub

Private Sub login_gmail()

'login to gmail and then pass to a routine that will pass the data to google from the driver's side


Driver.Start "chrome", "https://gmail.com"
Driver.Get ("https://gmail.com")
Driver.FindElementById("identifierId").SendKeys ThisWorkbook.Sheets("Staging").Range("E1") & "@diversey.com"
Driver.FindElementById("identifierNext").Click
Driver.FindElementById("uname").SendKeys ThisWorkbook.Sheets("Staging").Range("F1")
Driver.FindElementById("pass").SendKeys ThisWorkbook.Sheets("Staging").txtSPW
Driver.FindElementByName("loginButton2").Click

'login once, redirect to a second routine that creates an array of the items to be fed to google and then loops through the next few lines to refresh the existing page

pass_info

End Sub


Private Sub pass_info()

'this routine is going to build and pass the information to google


Dim arr_elem(0 To 20) As Variant '21 element array to hold the info for each row to be passed
Dim rng As Range, r1 As Range, r2 As Range, cell As Range, lstcl As Variant 'find the range to be passed

'build the range to be passed

With ThisWorkbook.Sheets("Staging")

.Activate

lstcl = .Range("C10000").End(xlUp).row

    For Each cell In .Range("V4:V" & lstcl)
    
        If Not cell = "TFR" Then
        
            Set r1 = cell.Offset(0, -21) 'cell A with the first cell to be transferred
            Set r2 = .Range("A" & lstcl) 'bottom of the sheet with a value in it in column A
            Set rng = Range(r1, r2) 'make a range
            rng.Select
            Exit For

    
        End If

    Next
    
    'exit sub in case there are no movements to upload
    
    If r1 Is Nothing And r2 Is Nothing Then
                            
        MsgBox "No movements to upload", vbExclamation
        Exit Sub
            
    End If
    
    
    'build ranges of each row that needs to be populated in the URL
    
    Dim rng2 As Range, c1 As Range, c2 As Range 'variables to make each row to be copied
    Dim cell2 As Range 'variable to cycle through the rng made above
    Dim copy_cell As Range 'variable to cycle through each cell in rng2 when it's made
    
    For Each cell2 In rng
    
        
    
            Set c1 = cell2 'first cell in the row to be copied into google (date the movement is made)
            Set c2 = cell2.Offset(0, 21) 'last cell in the row to be copied into google (timestamp)
            Set rng2 = Range(c1, c2) 'make a range of the row to be fed into google
            rng2.Select
            
            Dim link As String 'the link to google where things will be submitted
            Dim link2 As String 'the link to the google sheet that will be activated last
            
            'cycle through the cells in rng2
            
            For Each copy_cell In rng2
            
                'feed each element into an array location
                Set arr_elem(0) = copy_cell
                Set arr_elem(1) = copy_cell.Offset(0, 1)
                Set arr_elem(2) = copy_cell.Offset(0, 2)
                Set arr_elem(3) = copy_cell.Offset(0, 3)
                Set arr_elem(4) = copy_cell.Offset(0, 4)
                Set arr_elem(5) = copy_cell.Offset(0, 5)
                Set arr_elem(6) = copy_cell.Offset(0, 6)
                Set arr_elem(7) = copy_cell.Offset(0, 7)
                Set arr_elem(8) = copy_cell.Offset(0, 8)
                Set arr_elem(9) = copy_cell.Offset(0, 9)
                Set arr_elem(10) = copy_cell.Offset(0, 10)
                Set arr_elem(11) = copy_cell.Offset(0, 11)
                Set arr_elem(12) = copy_cell.Offset(0, 12)
                Set arr_elem(13) = copy_cell.Offset(0, 13)
                Set arr_elem(14) = copy_cell.Offset(0, 14)
                Set arr_elem(15) = copy_cell.Offset(0, 15)
                Set arr_elem(16) = copy_cell.Offset(0, 16)
                Set arr_elem(17) = copy_cell.Offset(0, 17)
                Set arr_elem(18) = copy_cell.Offset(0, 18)
                Set arr_elem(19) = copy_cell.Offset(0, 19)
                Set arr_elem(20) = copy_cell.Offset(0, 20)
            Exit For
            
            
            
        Next
        
        
            
            'send each element to google
            link = "https://docs.google.com/forms/d/e/1FAIpQLScDo57NRUxeEQ6EoasJj0WiizCinzZt8G-3_iOclNktT_34-w/formResponse?" _
            & "entry.1662196013=" & arr_elem(0) _
            & "&entry.1184587076=" & arr_elem(1) _
            & "&entry.870170607=" & arr_elem(2) _
            & "&entry.1504473992=" & arr_elem(3) _
            & "&entry.1080983622=" & arr_elem(4) _
            & "&entry.1822817214=" & arr_elem(5) _
            & "&entry.851488661=" & arr_elem(6) _
            & "&entry.1822034012=" & arr_elem(7) _
            & "&entry.1011704712=" & arr_elem(8) _
            & "&entry.104233611=" & arr_elem(9) _
            & "&entry.1194179537=" & arr_elem(10) _
            & "&entry.1737519045=" & arr_elem(11) _
            & "&entry.1399879278=" & arr_elem(12) _
            & "&entry.2066595584=" & arr_elem(13) _
            & "&entry.1654147617=" & arr_elem(14) _
            & "&entry.1480800053=" & arr_elem(15) _
            & "&entry.1093959551=" & arr_elem(16) _
            & "&entry.1457032635=" & arr_elem(17) _
            & "&entry.729463285=" & arr_elem(18) _
            & "&entry.576636500=" & arr_elem(19) _
            & "&entry.650295314=" & arr_elem(20) _
            & "&fvv=1&draftResponse=%5Bnull%2Cnull%2C%22-2106807567670054409%22%5D%0D%0A&pageHistory=0&fbzx=-2106807567670054409"

            'insert TFR in the row being copied to ascertain that it has been transferred and will not be transferered again
            If Not IsEmpty(cell2.Offset(0, 21)) Then
            
                If cell2.Offset(0, 21) = "RTN" Or cell2.Offset(0, 21) = "STGU" Then
                
                    cell2.Offset(0, 21) = "TFR"
                    cell2.Offset(0, 21).Interior.Color = vbGreen
                
                End If
                
                'STGU code will signify submitted staged orders that have not been returned yet
                If cell2.Offset(0, 21) = "STG" Then
                
                    cell2.Offset(0, 21) = "STGU"
                    
                End If
                
            End If
            
            'refresh the page
            'MsgBox link
            Driver.Get (link)
            
        
        
    Next
    
End With

'set link2 as the spreadsheet to be opened after update is done
            link2 = "https://docs.google.com/spreadsheets/d/1LjpmfMi9UAjyoRfoQESS-ZtuCYHxcOZQ2G9m8AXwtcA/edit#gid=1093102243"
            Driver.Get (link2)


End Sub

Private Sub SAP_Monkey_download()

'this routine will allow to download the driver's update from Google sheets, instead of emails and integrate it with the existing database


Dim cell As Range, cell2 As Range
Dim link As String 'URL from the table where infomration will be downloaded
Dim to_find As Range, rng_to_eval As Range, c1 As Range, c2 As Range 'variables for DriverUpdate
Dim rng_to_scan As Range, lstcl As Variant, r1 As Range, r2 As Range 'variables for Staging
Dim rng_to_tfr As Range, t1 As Range, t2 As Range 'variables to pass data from Driver Sheet to staging
Dim s1 As Range, s2 As Range, rng_paste As Range 'variables to pass data in the proper location in Staging

'download data to DriverUpdate sheet

Sheets("DriverUpdate").Activate


link = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSQ3h9BfQLkJtGMI3d9a3xPuxie7L-RLh2GhtMBMf9ZwW4OF0zCWAw4es0eFgMIUnTYGSDvo4gxXwRt/pubhtml"

Sheet4.QueryTables(1).Connection = "URL;" & link
Sheet4.QueryTables(1).Refresh False
Sheet4.Columns(1).ColumnWidth = 10

'build the range of timestamps in the Staging sheet

With ThisWorkbook.Sheets("Staging")

    .Activate

    lstcl = .Range("U100000").End(xlUp).row
    
    Set r1 = .Range("U5")
    Set r2 = .Range("U" & lstcl)
    Set rng_to_scan = .Range(r1, r2)
    
    
End With

'build the range which will be compared in the DriverUpdate sheet

With ThisWorkbook.Sheets("DriverUpdate")

    .Activate

    Columns("A:B").Delete
    lstcl2 = .Range("U100000").End(xlUp).row
    
    Set c1 = .Range("U8")
    Set c2 = .Range("U" & lstcl2)
    Set rng_to_eval = .Range(c1, c2)
    

End With

Sheets("DriverUpdate").Activate
    
    'range in DriverUpdate which will be downloaded and compared against the Staging timestamps
    For Each cell In rng_to_eval
    
        cell.Activate
        
        'with each pass, new info is put into Staging, so the range should grow to include that data
        lstcl = ThisWorkbook.Sheets("Staging").Range("U100000").End(xlUp).row
        Set r2 = ThisWorkbook.Sheets("Staging").Range("U" & lstcl)
        Set rng_to_scan = Range(r1, r2)
        
        'range in Staging that will be searched for every timestamp in DriverUpdate
        With rng_to_scan
        
            '.Activate
            
            
        
            Set to_find = .Find(cell)
            
            'this section is to handle pasting return movement for orders that are already staged
            If Not to_find Is Nothing Then
            
                'when it finds the timestamp, the routine is going to make a range of the return movements and put them in
                'it will paste existing movements, as well as new ones
                Set t1 = cell.Offset(0, -6)
                Set t2 = cell.Offset(0, -3)
                'Set t3 = cell.Offset(0, -2)
                Set rng_to_tfr = Range(t1, t2)
                rng_to_tfr.Select
                rng_to_tfr.Copy
                
                With ThisWorkbook.Sheets("Staging")
                
                    .Activate
                    
                    'make the range where the return movement information is going to fit in
                    Set s1 = to_find.Offset(0, -6)
                    Set s2 = to_find.Offset(0, -3)
                    Set rng_paste = Range(s1, s2)
                    rng_paste.Select
                    
                    'populate the range in Staging with rtn movement information for the appopriate time stamp
                    rng_paste.Value = rng_to_tfr.Value
                
                End With
            
            End If
            
            If to_find Is Nothing Then
            
                'copy the row in the next available line in the main Staging sheet
                'build a range out of the active row
                
                
                
                Set t1 = cell.Offset(0, -20)
                Set t2 = cell
                Set rng_to_tfr = Range(t1, t2)
                rng_to_tfr.Select
                rng_to_tfr.Copy
                
                With ThisWorkbook.Sheets("Staging")
                
                    .Activate
                
                    lstcl = .Range("U100000").End(xlUp).row
                    
                    'copy the range to the Staging sheet
                    
                    'if pasting a movmement, it will have the same timestamp as above it, so offset paste range to next row
                    If cell.Offset(-1, 0) = cell Then
                        Set s1 = .Range("A" & lstcl + 1)
                    End If
                    
                    'if starting a new order, it will have a blank cell in the above row, so offset paste range by 2 to create a break btw orders
                    If cell.Offset(-1, 0) = "" Then
                        Set s1 = .Range("A" & lstcl + 2)
                    End If
                    Set s2 = s1.Offset(0, 20)
                    Set rng_paste = Range(s1, s2)
                    'rng_paste.Select
                    
                    '.Range("A" & lstcl).Offset(1, 0).Activate
                    rng_paste.Value = rng_to_tfr.Value
                    
                End With
                
            End If
            
        End With
        
        Sheets("DriverUpdate").Activate
        
    Next
    
'sort orders by timestamp - critical part about establishing the continuity of production runs for the entire factory

'make a range of the contents in the spreadsheet

Dim lstC As Variant, rng_sort As Range

With ThisWorkbook.Sheets("Staging")

    .Activate

    lstC = .Range("C100000").End(xlUp).row
    Set rng_sort = .Range("A4:Z" & lstC)
    rng_sort.Select
    
    .Sort.SortFields.Add Key:=Range("U5:U" & lstC), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    
    With .Sort
    
        .SetRange rng_sort
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
End With


    
'insert blank rows as breaks between newly inserted orders

Dim cell3 As Range

With ThisWorkbook.Sheets("Staging")

    .Activate

    lstcl = .Range("U100000").End(xlUp).row
    
    For Each cell3 In .Range("B4:B" & lstcl)
    
        If Not IsEmpty(cell3) Then
        
            If Not IsEmpty(cell3.Offset(-1, -1)) Then
            
                Range(cell3.Address).EntireRow.Insert
            
                
            End If
    
        End If
    
    Next


End With

End Sub



Private Sub POss_for_SAP()

'this routine is going to create the list of POs and populate the combo box with them

Dim lstcl As Variant, arr_po() As Variant, x As Integer
Dim r1 As Range, r2 As Range, rng As Range, cell As Range, cell_m As Range
Dim ctr As Integer

'cell_po As Range declared as a public variable at the top of the module for use by this routine and analyze_SAP_PO

'UserForm12.cboPOSAP.Clear

With ThisWorkbook.Sheets("Staging")

lstclB = .Range("B10000").End(xlUp).row
lstcl = .Range("C10000").End(xlUp).row

'UserForm12.cboPOSAP.Clear

    For Each cell_po In .Range("B4:B" & lstclB)
    
        'find each order number in the cell that isn't empty
        If Not IsEmpty(cell_po) Then
        
        'set the top of the range as the PO that is located
        Set r1 = cell_po
        
            'loop through each cell from that PO to the bottom of the data
            For Each cell In .Range(r1.Offset(1, 0), "B" & lstcl)
            
                cell.Select
            
                'if it's the last PO in the spreadsheet, set r2 at the bottom of the data
                If r1 = Range("B" & lstclB) Then
                
                        Set r2 = Range("B" & lstcl)
                        Set rng = Range(r1, r2)
                        rng.Select
                    
                End If
                    
                'otherwise, set r2 just before the next order
                If Not IsEmpty(cell) Then
                
                    'set a range of the materials in each order
                
                    Set r2 = cell.Offset(-1, 0)
                    Set rng = Range(r1.Offset(0, 1), r2.Offset(0, 1))
                    rng.Select
                    
                End If
                
                    'check if there are unsigned materials in each range
                    ctr = 0
                    If Not rng Is Nothing Then
                        
                        With rng
                        
                        .Select
                    
                            For Each cell_m In rng
                            
                                If IsEmpty(cell_m.Offset(0, 7)) Or IsEmpty(cell_m.Offset(0, 19)) Then
                                
                                    cell_m.Offset(0, 7).Select
                                    cell_m.Offset(0, 19).Select
                                    cell_po.Select
                                
                                    ctr = ctr + 1
                                    
                                End If
                                
                            Next
                        
                        End With
                        
                        'exit the inner cell cycling through each line of hte main order range _
                        'so that it's not repeated for the blank cells between PO entries
                        Exit For
                        
                    End If
           Next
           
        'add the matching order to the combo box and reset the counter for the next PO that matches
        
        If ctr > 0 Then
            UserForm12.cboPOSAP.AddItem cell_po
        End If
        ctr = 0
        
        End If

    Next

End With


    
If Not UserForm12.Visible = True Then

UserForm12.Show vbModeless

End If

End Sub

Sub analyze_SAP_PO()

Dim k As Integer, PO As Variant, lstcl As Variant, lstcl2 As Variant
Dim fnd_PO As Variant, cell As Range, cell2 As Range, r1 As Range, r2 As Range
Dim rng_to_fnd As Range
Dim b As Integer, c As Integer, arr_stg() As Range, arr_rtn() As Range


'prevent listboxes from being erased if they are activated after an order is changed
If ThisWorkbook.Sheets("Staging").Range("AA5").Interior.Color = vbRed Or _
ThisWorkbook.Sheets("Staging").Range("AA2").Interior.Color = vbRed Then

UserForm12.lstSTG.Clear
UserForm12.lstRTN.Clear
ThisWorkbook.Sheets("Staging").Range("AA5").Interior.Color = vbGreen
ThisWorkbook.Sheets("Staging").Range("AA2").Interior.Color = vbGreen

End If

'this routine is going to analyze the list of orders after downloading and determine which ones need to have _
'movements in SAP done based on the lack of SAP Clerk initials

k = UserForm12.cboPOSAP.ListIndex
If k > -1 Then
PO = UserForm12.cboPOSAP.List(k)
End If


'no need to make an entire collection of POs, only need to find the one specified by the use r
'the only array that needs to be made is to populate te comb obox initially
With ThisWorkbook.Sheets("Staging")

lstcl = .Range("B10000").End(xlUp).row
lstcl2 = .Range("C10000").End(xlUp).row

'make a range out of all the orders
Set rng_to_fnd = .Range("B4:B" & lstcl)

    'find the PO specified in the combo box
    With rng_to_fnd
    
        .Select
    
        'locate PO
        Set fnd_PO = .Find(PO)
        
        'if found, make it the top of the range to be evaluated for SAP Clerk initials
        If Not fnd_PO Is Nothing Then
    
            Set r1 = fnd_PO
            r1.Select
        
            'if the PO in question is the last on the list, set the bottom extent of the range with the last cell w. data
            Dim c1 As Range, c2 As Range, PO_rng As Range
            
            For Each cell In Range(r1.Offset(1, 0), "B" & lstcl)
            
                cell.Activate
            
                If r1 = Range("B" & lstcl) Then
                
                    Set c1 = r1.Offset(1, 0) 'set as offset a row so the loop doesn't count the top of the range and exit prematurely
                    Set c2 = Range("B" & lstcl)
                    Set PO_rng = Range(c1, c2)
                    PO_rng.Select

                    Set r2 = .Range("B" & lstcl2)
                    r2.Select
                    Exit For
                 
                End If
                 
                 'else make the bottom end of the range just shy of the next order in the list
                
                 
                If Not IsEmpty(cell) Then
                     
                    Set r2 = cell.Offset(-1, 0)
                    r2.Select
                    Exit For
                         
                End If
            
            Next
            
            'make a range out of the materials in the relevant order to be evaluated to be evaluated
            Set eval_rng = Range(r1.Offset(0, 1), r2.Offset(0, 1))
            eval_rng.Select
            
            'in case the last order in the list is selected
            If r1 = Range("B" & lstcl) Then
                Set r1 = r1.Offset(0, 1)
                Set r2 = Range("B" & lstcl2).Offset(0, 1)
                Set eval_rng = Range(r1, r2)
            End If
            
            eval_rng.Select
            
        End If
        
    End With
            
End With

'logic of the following:
'initializing requires to find the particular movements in particular orders
'POs and movements need to be found once to display them and evaluate them and found again to initial them_
'to avoid having to find a PO all over again, a system of colour codes is implemented to reuse the same code _
'when different triggers are applied
'at the end of each routine, the color of the cell changes to prevent it from being called again by another _
'routine

'triggers being:
'Get Movements - reestablish the PO and get the materials that need to be initialed
'Clicking on a listbox - to get the data for each material displayed in the labels inside UF12
'INIT - initializing the correct material on the correct line in the correct spot


With ThisWorkbook.Sheets("Staging")

    'trigger when Get Movements button is clicked
    'send to analzye what movements need to be autographed
    If .Range("AA3").Interior.Color = vbYellow Then
        UserForm12.lstSTG.Clear
        UserForm12.lstRTN.Clear
        get_SKUs_to_sign eval_rng
        Exit Sub
    End If
    
    'feeds from Get Movement button click event to populate the relevant staged and returned mats to listboxes
    If .Range("AA1").Interior.Color = vbGreen Then
        get_mats_STG r1, r2, eval_rng
    End If
    
    'feeds from Get Movement button click event to populate the relevant staged and returned mats to listboxes
    If .Range("AA1").Interior.Color = vbCyan Then
        get_mats_RTN r1, r2, eval_rng
    End If
    
    'trigger when pressing the INIT button on the staging or returning to initials SAP moves
    If .Range("AA2").Interior.Color = vbBlue Then
        'send for SAP clerk signature
        SAP_Clerk_Init_STG r1, r2, eval_rng
    End If
    
    'trigger when pressing the INIT button on the staging or returning to initials SAP moves
    If .Range("AA2").Interior.Color = vbBlack Then
        'send for SAP clerk signature
        SAP_Clerk_Init_RTN r1, r2, eval_rng
    End If
    
    'this triggers from the get_SKUs_to_sign app to remove materials that have been signed off
    'If .Range("AA4").Interior.Color = vbBlack Then
        
        'ctr2 = ctr2 + 1
        'get_SKUs_to_sign eval_rng, ctr2
        'Exit Sub 'to prevent an infinite loop from the app refreshing the list of movements
    'End If
            
End With



End Sub

Private Sub get_SKUs_to_sign(eval_rng)

'this routine analyzes the movements that need to be signed by the SAP clerk
'Dim b As Integer, arr_stg() As Range, arr_rtn() As Range, c  As Integer

'use of addiitem is neccesitated because using an array populates the listbox, but makes the content invisible
'could not figure out why that happens

With UserForm12
    .lstRTN.Clear
    .lstSTG.Clear
End With

    'go through the materials to check for initials for the staged materials
    For Each cell2 In eval_rng
            
        If IsEmpty(cell2.Offset(0, 6)) Then
            
            UserForm12.lstSTG.AddItem cell2
            'ReDim Preserve arr_stg(b)
            'Set arr_stg(b) = cell2
            'b = b + 1
            
        End If
        
        'go through the materials in the returns section to check for movements that need to be returned
        
        'testing code
        w = cell2.Offset(0, 10)
        q = cell2.Offset(0, 17)
        
        'check if there is a material in the return section and if there is, if it is initialed
        
        If Not IsEmpty(cell2.Offset(0, 10)) And IsEmpty(cell2.Offset(0, 17)) Then
        
            'ReDim Preserve arr_rtn(c)
            'Set arr_rtn(c) = cell2.Offset(0, 9)
            'c = c + 1
            UserForm12.lstRTN.AddItem cell2.Offset(0, 10)
        End If
        
    Next
    
'turn cell AA3 red to prevent triggering the Get Movements command again
ThisWorkbook.Sheets("Staging").Range("AA3").Interior.Color = vbRed
'trigger Get Movements again to get rid of material that has already been initialed (as analyze PO is w/ each trigger)

'turn AA1 green to trigger this routine again after
'ThisWorkbook.Sheets("Staging").Range("AA4").Interior.Color = vbBlack

'purpose of counter variable is to prevent an eternal loop with analyze_PO
'If ctr2 >= 2 Then
'    Exit Sub
'End If

'analyze_SAP_PO


End Sub

Private Sub get_mats_STG(r1, r2, eval_rng)

'this sub is going to populate the labels with the data for each material that needs to be SAP processed

Dim i As Integer, j As Integer, mat As Double, fmat As Range

With UserForm12

    i = .lstSTG.ListIndex
    
    If i > -1 Then
        mat = .lstSTG.List(i)
        a = i
    End If
    


End With


'populate the labels with the info that has to be input into SAP

If Not mat = 0 Then

    Set eval_rng = Range(r1.Offset(-1, 0), r2)

    With eval_rng
    
    .Select
        
        Set fmat = .Find(mat)
        
        If Not fmat Is Nothing Then
        
                Dim c1 As Range, c2 As Range, c As Range
                Set c1 = fmat
                Set c2 = fmat.Offset(0, 6)
                Set c = Range(c1, c2)
                c.Select
                
                With UserForm12
                    
                    .lblF1 = fmat.Offset(0, 1)
                    .lblT1 = fmat.Offset(0, 2)
                    .lblQ1 = fmat.Offset(0, 3)
                        
                End With
            
                
        End If
                
    End With

End If

'turn AA1 Blue to avoid triggering this routine again until a material is selected from a listbox
'(as analyze PO is w/ each trigger)
ThisWorkbook.Sheets("Staging").Range("AA1").Interior.Color = vbBlue

End Sub


Private Sub get_mats_RTN(r1, r2, eval_rng)

Dim j As Integer, b As Integer, fmat_r As Range, mat_r As Double

With UserForm12

    j = .lstRTN.ListIndex
    
    If j > -1 Then
        mat_r = .lstRTN.List(j)
        b = j + 1
    End If
    
End With

'populate the labels with the info that has to be input into SAP
If Not mat_r = 0 Then

    Set eval_rng = Range(r1.Offset(-1, 10), r2.Offset(0, 10))
    
    With eval_rng
    
        .Select
        
        Set fmat_r = .Find(mat_r)
        
        If Not fmat_r Is Nothing Then
        
            
                fmat_r.Activate
    
                With UserForm12
                
                    .lblF2 = fmat_r.Offset(0, 1)
                    .lblT2 = fmat_r.Offset(0, 2)
                    .lblQ2 = fmat_r.Offset(0, 3)
                    
                End With
                
            
        End If
                
    End With

End If

'turn AA1 Blue to avoid triggering this routine again until a material is selected from a listbox
'(as analyze PO is w/ each trigger)
ThisWorkbook.Sheets("Staging").Range("AA1").Interior.Color = vbBlue

End Sub

Private Sub SAP_Clerk_Init_STG(r1, r2, eval_rng)

'this routine is going to allow the SAP clerk to inititalize each SAP movement

Dim i As Integer, j As Integer, mat As Double, fmat As Range

With UserForm12

    i = .lstSTG.ListIndex
    
    If i > -1 Then
        mat = .lstSTG.List(i)
        a = i
    End If
    


End With


'populate the labels with the info that has to be input into SAP for STAGING

If Not mat = 0 Then

    Set eval_rng = Range(r1.Offset(a - 1, 0), r2)

    With eval_rng
    
        .Select
        
        Set fmat = .Find(mat)
        
        If Not fmat Is Nothing Then
        
            
            fmat.Activate

            'initiialize the material with the SAP clerk inititals
            fmat.Offset(0, 6) = ThisWorkbook.Sheets("Staging").Range("H1")
       
        End If
                
    End With


End If

'avoid triggering this routine again until the Init button is clicked again (as analyze PO is w/ each trigger)
ThisWorkbook.Sheets("Staging").Range("AA2").Interior.Color = vbGreen

End Sub

Private Sub SAP_Clerk_Init_RTN(r1, r2, eval_rng)

Dim fmat_r As Range, mat_r As Double, b As Integer, j As Integer

With UserForm12

    j = .lstRTN.ListIndex
    
    If j > -1 Then
        mat_r = .lstRTN.List(j)
        b = j + 1
    End If
    
End With

'populate the labels with the info that has to be input into SAP for RETURNS
If Not mat_r = 0 Then

    Set eval_rng = Range(r1.Offset(-1, 11), r2.Offset(0, 11))
    
    With eval_rng
    
        .Select
        
        Set fmat_r = .Find(mat_r)
        
        If Not fmat_r Is Nothing Then
            
            fmat_r.Activate
            
            'enter the SAP return quantity, as input by the SAP clerk from LX02
            If Not IsEmpty(UserForm12.txtSAPQ2) Then
                fmat.Offset(0, 4) = UserForm12.txtSAPQ2
            End If
            'initialize the fields with SAP clerk initials
            fmat_r.Offset(0, 7) = ThisWorkbook.Sheets("Staging").Range("H1")

        End If
                
    End With

End If

'avoid triggering this routine again until the Init button is clicked again (as analyze PO is w/ each trigger)
ThisWorkbook.Sheets("Staging").Range("AA2").Interior.Color = vbGreen

'get_SKUs_to_sign eval_rng

End Sub

Private Sub evaluate_orders_for_GoogleE()

'this sub is going to feed the SAP clerk update into a finalized sheet
'need a google form like the one used for the driver uploads
'need a routine identical to the one used to send stuff back to google

Dim arr_upld(0 To 20) As Variant 'array submitting each row of an order to the URL
Dim cell As Range, lstcl As Variant, rng As Range, r1 As Range, r2 As Range 'building order ranges
Dim link As String 'link to google sheet that will be updated
Dim arr_orders() As Range, x As Integer

With ThisWorkbook.Sheets("Staging")

    lstcl = .Range("B10000").End(xlUp).row
    
    For Each cell In .Range("B4:B" & lstcl)
    
        'build a range of orders and then evaluate them
    
        If Not IsEmpty(cell) Then
        
            ReDim Preserve arr_orders(x)
            Set arr_orders(x) = cell
            Debug.Print arr_orders(x)
            'UserForm10.lstPOs.AddItem cell
            x = x + 1

        End If
    
    Next

    lstcl = .Range("C10000").End(xlUp).row
    
    
    'cycle through each element of the array, build a range from each order
    
    Dim a As Long, cell2 As Range 'variable to go through the elements in array
    Dim lstrowB As Variant
    
    lstrowB = .Range("B10000").End(xlUp).row
    
    For a = LBound(arr_orders) To UBound(arr_orders)
    
        Set r1 = arr_orders(a)
        
        For Each cell2 In Range(r1.Offset(1, 0), "B" & lstcl)
        
            'to capture the last dataset, make the end of the range be the end of data in the active sheet
            If r1 = .Range("B" & lstrowB) Then
                
                Set r2 = .Range("B" & lstcl)
                Exit For
                
            End If
        
            'else just define the extent of each order in the list
            If Not IsEmpty(cell2) Then
            
                Set r2 = cell2.Offset(-1, 0)
                Exit For
                
            End If
            
        Next
        
        Set rng = Range(r1, r2)
        rng.Select
        
        'this is where the program checks for SAP clerk initials and for relevant details in the returns section
        
        Dim cell_chk As Range
        
        For Each cell_chk In rng
        
            'check that SAP clerk initials are filled in for the staging
            If Not IsEmpty(cell_chk.Offset(0, 7)) Then
                
                'make the staged movt yellow
                Range(cell_chk.Offset(0, -1), cell_chk.Offset(0, 7)).Interior.Color = vbYellow
                
                'make the timestamp yellow
                cell_chk.Offset(0, 19).Interior.Color = vbYellow
                
                'make the TFR stamp yellow
                cell_chk.Offset(0, 20).Interior.Color = vbYellow
                
                'put the SAP STGD mark in column W
                cell_chk.Offset(0, 20) = "SAP STGD"
                cell_chk.Offset(0, 20).Interior.Color = vbYellow
                
            End If
            
            'check the return movements and fill them in with green
            
            Dim n As Integer
            
            'if there is a returned material, check that all of its cells afterward are filled correctly
            If Not IsEmpty(cell_chk.Offset(0, 11)) Then
            
                For n = 12 To 17
                    
                    'check that material numbers exist before filling in colours
                    If Not IsEmpty(cell_chk.Offset(0, n)) Then
                    
                        'recalculate scrap
                        cell_chk.Offset(0, 16) = cell_chk.Offset(0, 15) - cell_chk.Offset(0, 14)
                        
                        'make the return movement field green if all conditions are met
                        Range(cell_chk.Offset(0, 9), cell_chk.Offset(0, 17)).Interior.Color = vbGreen
                        
                        'create a SAP RTND stamp at the end of the sheet and make it green
                        cell_chk.Offset(0, 21) = "SAP RTND"
                        cell_chk.Offset(0, 21).Interior.Color = vbGreen
                    
                    Else
                    
                        MsgBox "You are missing values in row " & cell_chk.row & " for PO " & r1, vbExclamation
                        Exit Sub
                    
                    End If
                    
                Next
                
     
            End If
            
            'check if material numbers don't exist, but there is info for an oprhaned movement
            If IsEmpty(cell_chk.Offset(0, 11)) Then
                
                For n = 12 To 17
                
                    If Not IsEmpty(cell_chk.Offset(0, n)) Then
                    
                        MsgBox "There is a movement without an attached material on row " & cell_chk.row & ". " _
                        & "Please check if those details are connected to a material.", vbExclamation
                        cell_chk.Offset(0, n).Activate
                        Exit Sub
                    
                    End If
                    
                Next
                
            End If
        
        Next
        
     Next
            
    'get the orders that are completed
    
    'recreate the same ranges as above, and include the criteria (e.g. colors), created in the foregoing section
    
    Dim arr_donePOs() As Range, Y As Integer 'the array that will hold all the completed POs once all conditions are tested
    Dim btm_rng As Range, cel_eval As Range
    
    'reaffirm extent of data for materials and orders
    lstcl = .Range("C10000").End(xlUp).row
    lstrowB = .Range("B10000").End(xlUp).row

    For a = LBound(arr_orders) To UBound(arr_orders)
    
        Set r1 = arr_orders(a)
        
        For Each btm_rng In Range(r1, "B" & lstcl)
            
            'set the bottom most range
            If r1 = .Range("B" & lstrowB) Then
                
                Set r2 = .Range("B" & lstcl)
                Exit For
                
            End If
            
            'set each order range
            If Not IsEmpty(btm_rng) Then
            
                Set r2 = btm_rng.Offset(-1, 0)
                
                
            End If
        Next
            Set rng = Range(r1, r2)
            rng.Select
    
            'evaluate each order according to all the criteria and add it to the array
            For Each cel_eval In rng
            
                'ensure that the cell is yellow in colour
                If cel_eval.Interior.Color = vbYellow Then
                
                    'ensure that timestamps is there and yellow
                    If cel_eval.Offset(0, 19).Interior.Color = vbYellow And Not IsEmpty(cel_eval.Offset(0, 19)) Then
                    
                        'ensure that SAP STGD confirmation is there and in yellow
                        If cel_eval.Offset(0, 20).Interior.Color = vbYellow And cel_eval.Offset(0, 20) = "SAP STGD" Then
                        
                            'see if there is a returned material
                            If cel_eval.Offset(0, 11).Interior.Color = vbGreen And Not IsEmpty(cel_eval.Offset(0, 11)) Then
                            
                                'ensure that orders with and without returns are captured
                                If cel_eval.Offset(0, 11).Interior.Color = vbGreen And Not IsEmpty(cel_eval.Offset(0, 11)) Or _
                                    cel_eval.Offset(0, 11) = "NO RTNS" Then
                                
                                    'ensure that the order has not already been transferred to Google
                                    If Not cel_eval.Offset(0, 22) = "GOOG TFR" And Not cel_eval.Offset(0, 22).Interior.Color = vbRed Then
                                    
                                        'see that all return material data is there
                                        For n = 12 To 17
                                            
                                            'check that there is content and the cells are green
                                            If cel_eval.Offset(0, n).Interior.Color = vbGreen And Not IsEmpty(cel_eval.Offset(0, n)) Or cel_eval.Offset(0, n) = "NO RTNS" Then
                                            
                                                'check that SAP RTND confirmation is in place
                                                If cel_eval.Offset(0, 21).Interior.Color = vbGreen And cel_eval.Offset(0, 21) = "SAP RTND" Or cel_eval.Offset(0, 21) = "" Then
                                                    
                                                    'put the value of r1 100 columns away, and use it after to build an array of unique values
                                                    cel_eval.Offset(0, 100) = r1
                                                    
                                                End If
                                            
                                            End If
                                            
                                        Next
                                    
                                    End If
                                    
                               End If
                               
                        End If
                        
                    End If
                    
                End If
                
            End If
            
        Next
        
    
      
Next

    'go through the posted orders, remove the duplicates and make an array of the items that make the criteria
    
    Dim lstCX As Variant, celCX As Range
    UserForm10.lstPOs.Clear
    
    lstCX = .Range("CX10000").End(xlUp).row
    
    With Range("CX4:CX" & lstCX)
        
        
        .Select
        .RemoveDuplicates Columns:=1, Header:=xlNo
        
        
    End With
        
        For Each celCX In .Range("CX4:CX" & lstCX)
        
            If Not IsEmpty(celCX) Then
            
                UserForm10.lstPOs.AddItem celCX
                
            End If
            
        Next
    .Range("CX4:CX" & lstCX).Clear
    
    UserForm10.Show vbModeless
    

End With

End Sub

Private Sub SAP_login_gmail()
'login to gmail and then pass to a routine that will pass the data to google

Driver.Start "chrome", "https://gmail.com"
Driver.Get ("https://gmail.com")
Driver.FindElementById("identifierId").SendKeys ThisWorkbook.Sheets("Staging").Range("E1") & "@diversey.com"
Driver.FindElementById("identifierNext").Click
Driver.FindElementById("uname").SendKeys ThisWorkbook.Sheets("Staging").Range("F1")
Driver.FindElementById("pass").SendKeys ThisWorkbook.Sheets("Staging").txtSPW
Driver.FindElementByName("loginButton2").Click

'login once, redirect to a second routine that creates an array of the items to be fed to google and then loops through the next few lines to refresh the existing page

publish_to_Google

End Sub


Private Sub publish_to_Google()

Dim arr_to_publ(0 To 23) As Variant, i As Integer, PO As Double


With UserForm10

    i = .lstPOs.ListIndex
    PO = .lstPOs.List(i)
    
End With

'find the active selection in the sheet

Dim rng As Range, lstcl As Variant, lstcl2 As Variant, fnd_PO As Variant, r1 As Range, r2 As Range, rng2 As Range

With ThisWorkbook.Sheets("Staging")

    lstcl = .Range("B10000").End(xlUp).row
    lstcl2 = .Range("C10000").End(xlUp).row
    
    'create range to search for selected PO
    
    Set r1 = .Range("B4")
    Set r2 = .Range("B" & lstcl2)
    Set rng = Range(r1, r2)
    rng.Select
    'find the PO that is chosen by the user and make a range out of it, with a provision if they chose the _
    'last PO in the sheet
    
    With rng
    
        .Select
    
        Set fnd_PO = .Find(PO)
        
        If Not fnd_PO Is Nothing Then
        
            Set r1 = fnd_PO
            r1.Offset(1, 0).Select
            r2.Select
            
            For Each cell In Range(r1.Offset(1, 0), r2)
            
                If r1 = .Range("B" & lstcl) Then
                
                    Set r2 = .Range("B" & lstcl2)
                    Exit For
                    
                End If
            
                If Not cell = "" Then
                
                    Set r2 = cell.Offset(-1, 0)
                    Exit For
                        
                End If
                
            Next
    
            'make a range out of each selected order
            Set rng2 = Range(r1, r2)
            rng2.Select
            
        End If
        
    End With
    
    'with the established range of the PO, take each cell in the row, apply it to an array and pass it to google
    
    Dim cell2 As Range, goog_rng As Range, c1 As Range, c2 As Range, link As String, link2 As String, goog As Range
    
    For Each cell2 In rng2
    
        'need to make a horizontal range out of each row that will be passed to google
        Set c1 = cell2.Offset(0, -1)
        Set c2 = cell2.Offset(0, 22)
        Set goog_rng = Range(c1, c2)
        goog_rng.Select
        
        For Each goog In goog_rng
        
            'assign items to array
            
            Set arr_to_publ(0) = goog
            Set arr_to_publ(1) = goog.Offset(0, 1)
            Set arr_to_publ(2) = goog.Offset(0, 2)
            Set arr_to_publ(3) = goog.Offset(0, 3)
            Set arr_to_publ(4) = goog.Offset(0, 4)
            Set arr_to_publ(5) = goog.Offset(0, 5)
            Set arr_to_publ(6) = goog.Offset(0, 6)
            Set arr_to_publ(7) = goog.Offset(0, 7)
            Set arr_to_publ(8) = goog.Offset(0, 8)
            Set arr_to_publ(9) = goog.Offset(0, 9)
            Set arr_to_publ(10) = goog.Offset(0, 10)
            Set arr_to_publ(11) = goog.Offset(0, 11)
            Set arr_to_publ(12) = goog.Offset(0, 12)
            Set arr_to_publ(13) = goog.Offset(0, 13)
            Set arr_to_publ(14) = goog.Offset(0, 14)
            Set arr_to_publ(15) = goog.Offset(0, 15)
            Set arr_to_publ(16) = goog.Offset(0, 16)
            Set arr_to_publ(17) = goog.Offset(0, 17)
            Set arr_to_publ(18) = goog.Offset(0, 18)
            Set arr_to_publ(19) = goog.Offset(0, 19)
            Set arr_to_publ(20) = goog.Offset(0, 20)
            Set arr_to_publ(21) = goog.Offset(0, 21)
            Set arr_to_publ(22) = goog.Offset(0, 22)
            Set arr_to_publ(23) = goog.Offset(0, 23)
            
            Exit For
            
        Next
        
            link = "https://docs.google.com/forms/d/e/1FAIpQLSfaj8DSCRFG5JeGVqcdLcVoFSVlCvtBoefRZeAELulgM3XElg/formResponse?" _
            & "entry.1662196013=" & arr_to_publ(0) _
            & "&entry.1184587076=" & arr_to_publ(1) _
            & "&entry.870170607=" & arr_to_publ(2) _
            & "&entry.1504473992=" & arr_to_publ(3) _
            & "&entry.1080983622=" & arr_to_publ(4) _
            & "&entry.1822817214=" & arr_to_publ(5) _
            & "&entry.851488661=" & arr_to_publ(6) _
            & "&entry.1822034012=" & arr_to_publ(7) _
            & "&entry.1011704712=" & arr_to_publ(8) _
            & "&entry.104233611=" & arr_to_publ(9) _
            & "&entry.1194179537=" & arr_to_publ(10) _
            & "&entry.1737519045=" & arr_to_publ(11) _
            & "&entry.1399879278=" & arr_to_publ(12) _
            & "&entry.2066595584=" & arr_to_publ(13) _
            & "&entry.1654147617=" & arr_to_publ(14) _
            & "&entry.1480800053=" & arr_to_publ(15) _
            & "&entry.1093959551=" & arr_to_publ(16) _
            & "&entry.1457032635=" & arr_to_publ(17) _
            & "&entry.729463285=" & arr_to_publ(18) _
            & "&entry.576636500=" & arr_to_publ(19) _
            & "&entry.650295314=" & arr_to_publ(22) _
            & "&%2C+19%3A51%3A32&fvv=1&draftResponse=%5Bnull%2Cnull%2C%223416597135105934321%22%5D%0D%0A&pageHistory=0&fbzx=3416597135105934321"


            'insert the GOOGLE STAMP at the end of the row
            goog.Offset(0, 23) = "GOOG TFR"
            goog.Offset(0, 23).Interior.Color = vbRed
            
    Driver.Get (link)
    Next
            
End With

link2 = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSQ0vf3l54bMToDiL0u4U436u3EjfiCxQGOZhzDlMo82ypP1oUzGPcUU8o0mpcam_Vrd36eb8ky0rX8/pubhtml"
Driver.Get (link2)

'hide the menu and re-evaluate the orders that still need to be sent to google
UserForm10.Hide
evaluate_orders_for_Google

End Sub

Private Sub clear_DRV_sheet()

'this routine is intended to clear the driver sheet to facilitate easier data transfer and reduce processing time

Dim lstcl As Variant, answer As Integer

With ThisWorkbook.Sheets("DriverUpdate")

lstcl = .Range("A10000").End(xlUp).row

'to prevent making a range of items that are not there

    If lstcl > 8 Then

        answer = MsgBox("You are about to erase all data in the DriverUpdate sheet. Do you want to continue?", vbYesNo + vbQuestion)
        
        If answer = vbYes Then
        
            Rows("8:" & lstcl).Select
            Selection.Delete Shift:=xlUp
            
        End If
    
    End If


End With

End Sub

'this sub and the ones after it are simply for trying different details and have no bearing on the program's function

Sub errors()

For Index = 1 To 500
    Debug.Print Error$(Index)
    Next
End Sub

Sub stagingg()
Dim ws As Worksheet

For Each ws In ThisWorkbook.Sheets
    If ws.name = "Staging" Then
    
    MsgBox "Found!"
    
    End If
    Next
    
End Sub

Sub sfsfsdf()

ActiveCell.Interior.Color = vbGreen
Rows(ActiveCell.row).EntireRow.Interior.Color = vbGreen



End Sub

Sub timestamppp()

ActiveCell = Format(Now, "mm-dd-yyyy, hh:mm:ss")
End Sub

Sub dateasinteger()

Dim x As Long
a = Date
b = Date - 7

x = CLng(a)
Y = CLng(b)

MsgBox x
End Sub

Sub datesinsheet()

'this is the prototype for calculating the dates in the archive

'simply convert the date portion of the timestamp to a long variable, subtract 90 from the current date and if it is that _
or greater, transfer the enture row to the next row in the archive and erase it from the staging sheet to prevent duplicate _
entrues in teh archice
Dim x As Date

For Each cell In Range("U4:U20")
    
    
    If Not cell = "" Then
    x = Left(cell, 10)
    Y = CLng(x)
    End If
    
Next
End Sub

