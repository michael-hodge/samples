Option Explicit
Dim vDBHost, vDBName, vDBUser, vDBPass, vSQL, vCN, vRS, vOutlook, vMessage, vMailItem, vExUser, vProgramNumber, vStartDate, vProgramType, vVenueCity, vVenueState, vSpeakerFirstName, vSpeakerLastName, vSpeakerEmail, vSpeakerPrefix, vArrivalTime, vVenueName, vVenueAddr1, vVenueAddr2, vVenueAddr, vVenueZip, vVenuePhone, vRepFirstName, vRepLastName, vRepPhone, vRepEmail, vProgramInput, vMessageTime, vTab, vUserRole, vUserExt, vUserName, vUserFirstName, vRoomName, vEstAttendees, vHeadcountDueDate, vCancellationNotes, vAVScreen, vAVFlatScreen, vAVCords, vAVProjector, vAVLaptop, vAVMicrophone, vMenu, vMealPrice, vProgramTitle, vPickQuery, vRepFullName, vLongStartDate, vAsstEmail, vHonorarium, vInvitationCount, vCompanyName, vVenueAddrNoTab, vVendorEmail, vVendorContactName, vUserAddress
Const vAttachRepIpadLogin = "S:\Customers - Active\Entera 2016\ENTSB16\Master Templates\Program Correspondence\Rep\Meeting Materials\Virtual Materials\iPad Login Steps.pdf"
Const vAttachRepSystemCheck = "S:\Customers - Active\Entera 2016\ENTSB16\Master Templates\Program Correspondence\Rep\Meeting Materials\Virtual Materials\System Check - Web Conference.pdf"
Const vAttachSpeakerTips = "S:\Customers - Active\Entera 2016\ENTSB16\Master Templates\Program Correspondence\Speaker\Meeting Materials\Virtual Materials\Webinar Speaker Tips and System Check.pdf"
' test programs:  1653AU2015, 1604MR1815

Sub EnteraLSpeakerConfEmail()

vPickQuery = "Standard"
Call EnteraGetData
    If vProgramInput = "" Then
        Exit Sub
    End If
    
Do Until vRS.EOF
    Call EnteraDefineVariables
   
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' build message
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Set vOutlook = CreateObject("outlook.application")
    Set vMessage = vOutlook.CreateItem(vMailItem)
    
    With vMessage
        .Subject = "CONFIRMATION:  Speaking Engagement for " & vCompanyName & " - " & vLongStartDate & " - " & vVenueCity & ", " & vVenueState & " - (" & vProgramNumber & ")"
        .HTMLBody = "<span LANG=EN>" _
        & "<p class=style2><span LANG=EN><font face=Calibri size=3>" _
        & "Good " & vMessageTime & " " & vSpeakerPrefix & " " & vSpeakerLastName & "," _
        & "<br><br> We have you confirmed for an <b>" & vCompanyName & " Speakers Bureau " & vProgramType & " Program </b> at <b>" & vArrivalTime & "</b> on <b>" & vLongStartDate & "</b> in <b>" & vVenueCity & ", " & vVenueState & "</b>. The title of this program is <b><i>" & vProgramTitle & "</i></b>" _
        & "<br><br> We will send a reminder prior to this program, which will include venue details and an expense report. We will be reimbursing you for any out of pocket expenses, as well as mailing you your honorarium." _
        & "<br><br><b> TRAVEL <font color = 'red'>(DELETE IF TRAVEL ISN'T NEEDED) </font> </b> " _
        & "<br> We will have Maupin Travel contact you to set up your travel arrangements for this program. Someone will be in contact with you soon." _
        & "<br><br> <b> ACTION REQUESTED<font color = 'red'> (DELETE SECTION IF TITLE ISN'T NEEDED) </font> </b> " _
        & "<br> Please provide the following information as you would like it to be <u>published</u> on invitations for this and future programs: " _
        & "<br> " & vTab & " Name, Credentials " _
        & "<br> " & vTab & " Professional Title " _
        & "<br> " & vTab & " Affiliation " _
        & "<br><br> You will need a laptop for this presentation.  If this is not available to you, please let us know." _
        & "<br><br> Best Regards, " _
        & "<br><br> " & vUserFirstName

        .To = vSpeakerEmail
        .CC = vAsstEmail & ";" & vRepEmail & "; jlyzen@plan365inc.com"
        .Display

    End With
    
    vRS.MoveNext
Loop

vRS.Close
Set vRS = Nothing
Set vCN = Nothing

End Sub

Sub EnteraVSpeakerConfEmail()

vPickQuery = "Standard"
Call EnteraGetData
    If vProgramInput = "" Then
        Exit Sub
    End If
    
Do Until vRS.EOF
    Call EnteraDefineVariables
   
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' build message
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Set vOutlook = CreateObject("outlook.application")
    Set vMessage = vOutlook.CreateItem(vMailItem)
    
    With vMessage
        .Subject = "CONFIRMATION:  Virtual Speaking Engagement for " & vCompanyName & " - " & vLongStartDate & " - " & vVenueCity & ", " & vVenueState & " - (" & vProgramNumber & ")"
        .HTMLBody = "<span LANG=EN>" _
        & "<p class=style2><span LANG=EN><font face=Calibri size=3>" _
        & "Good " & vMessageTime & " " & vSpeakerPrefix & " " & vSpeakerLastName & "," _
        & "<br><br> We have you confirmed for an <b>" & vCompanyName & " Virtual Speakers Bureau Program </b> on <b>" & vLongStartDate & "</b>.  The title of this program is <b><i>" & vProgramTitle & "</i></b>.  We will send you a calendar invitation with the webinar login instructions.  After the program, we will mail you your honorarium." _
        & "<table> <tr> <td style='width:150px'> <font face=Calibri size=3> <b><u>VENUE</u></b> </font> </td> <td style='width:600px'> </td> </tr> <tr>  <td style='width:150px'> <font face=Calibri size=3> Office Location:</td> </font> <td style='width:600px'> <font face=Calibri size=3> " & vVenueName & " </font> </td> </tr> <tr> <td style='width:150px'></td> <td style='width:600px'> <font face=Calibri size=3> " & vVenueCity & ", " & vVenueState & " " & vVenueZip & " </font> </td> </tr> <tr> <td style='width:150px'> <font face=Calibri size=3> Presentation Time: </font> </td> <td style='width:600px'> <font face=Calibri size=3> " & vArrivalTime & "<font color='red'> TIME ZONE</font>" & " </font> </td> </tr> <tr> <td style='width:150px'></td> <td style='width:600px'> <font face=Calibri size=3> <b> **Please plan to login 20 minutes prior to the program start time for a dry-run of the webinar platform.  Please perform the system check (attached) prior to the program.** </b> </font> </td> </tr>  </table> " _
        & "<br><br> <b><u>ACTION REQUESTED </u> <font color = 'red'>(DELETE SECTION IF TITLE ISN'T NEEDED)</font></u></b> " _
        & "<br> Please provide the following information as you would like it to be <u>published</u> on invitations for this and future programs: " _
        & "<br> " & vTab & " Name, Credentials " _
        & "<br> " & vTab & " Professional Title " _
        & "<br> " & vTab & " Affiliation " _
        & "<br><br> You will need a laptop and a landline phone for this presentation.  If this is not available to you, please let us know." _
        & "<br><br> Best Regards, " _
        & "<br>" & vUserFirstName

        .To = vSpeakerEmail
        .CC = vAsstEmail & ";" & vRepEmail & "; jlyzen@plan365inc.com"
        If Dir(vAttachSpeakerTips) = "" Then
            MsgBox "speaker tips file not found/attached"
        Else
            .Attachments.Add vAttachSpeakerTips
        End If
        .Display

    End With
    
    vRS.MoveNext
Loop

vRS.Close
Set vRS = Nothing
Set vCN = Nothing

End Sub

Sub EnteraLSpeakerReminderEmail()

vPickQuery = "Standard"
Call EnteraGetData
    If vProgramInput = "" Then
        Exit Sub
    End If

Do Until vRS.EOF
    Call EnteraDefineVariables
   
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' build message
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Set vOutlook = CreateObject("outlook.application")
    Set vMessage = vOutlook.CreateItem(vMailItem)
    
    With vMessage
        .Subject = "REMINDER: Speaking Engagement for " & vCompanyName & " - " & vLongStartDate & " - " & vVenueCity & ", " & vVenueState & " (" & vProgramNumber & ")"
        .HTMLBody = "<span LANG=EN>" _
        & "<p class=style2><span LANG=EN><font face=Calibri size=3>" _
        & "<b>Attachments:</b> <font color = 'red'> Expense Report, Travel Itinerary </font> " _
        & "<br><br> Good " & vMessageTime & " " & vSpeakerPrefix & " " & vSpeakerLastName & ", " _
        & "<br><br> Thank you for agreeing to speak at this <b>" & vCompanyName & " Speakers Bureau Program</b>, titled <b>" & vProgramTitle & "</b>. We will be handling your honorarium and expense reimbursement." _
        & "<br><br> <b><u>PROGRAM DETAILS</u></b> " _
        & "<br> " & vRepFullName & " is hosting this program and can be reached at " & vRepPhone & " or via email (copied)." _
        & "<br><br>" & vTab & "<b>Date: &nbsp;&nbsp; </b> " & vLongStartDate _
        & "<br><b>" & vTab & "Time: &nbsp;&nbsp;</b>  " & Format(vArrivalTime, "medium time") _
        & "<br>" & vTab & "<b>Venue:</b> " & vVenueName _
        & "<br>" & vTab & vTab & vTab & "&nbsp;&nbsp;" & vVenueAddr & "<br>" & vTab & vTab & vTab & "&nbsp;&nbsp;" & vVenueCity & ", " & vVenueState & " " & vVenueZip & "<br>" & vTab & vTab & vTab & "&nbsp;&nbsp;" & vVenuePhone _
        & "<br><br> <b><u>ATTACHED YOU WILL FIND:</u></b> " _
        & "<br>" & vTab & "&#9702; <font color='red'> Itinerary - (ENTER SPECIFIC TRAVEL LOGISTICS OR DELETE) </font> <br>" & vTab & "&#9702; Expense Report " _
        & "<br><br> Please bring a laptop for the presentation." _
        & "<br><br>Best Regards, <br> " & vUserFirstName

        .To = vSpeakerEmail
        .CC = vAsstEmail & ";" & vRepEmail & "; jlyzen@plan365inc.com"
        .Display

    End With
    
    vRS.MoveNext
Loop

vRS.Close
Set vRS = Nothing
Set vCN = Nothing

End Sub

Sub EnteraVSpeakerReminderEmail()

vPickQuery = "Standard"
Call EnteraGetData
    If vProgramInput = "" Then
        Exit Sub
    End If

Do Until vRS.EOF
    Call EnteraDefineVariables
   
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' build message
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Set vOutlook = CreateObject("outlook.application")
    Set vMessage = vOutlook.CreateItem(vMailItem)
    
    With vMessage
        .Subject = "REMINDER: Virtual Speaking Engagement for " & vCompanyName & " - " & vLongStartDate & " - " & vArrivalTime & " (" & vProgramNumber & ")"
        .HTMLBody = "<span LANG=EN>" _
        & "<p class=style2><span LANG=EN><font face=Calibri size=3>" _
        & "<b>Attachments:</b> System Check & Speaker Tips, Rep-selected Slide Deck" _
        & "<br><br> Good " & vMessageTime & " " & vSpeakerPrefix & " " & vSpeakerLastName & ", " _
        & "<br><br> Thank you for agreeing to speak at this <b>" & vCompanyName & " Virtual Speakers Bureau Program</b>, titled <b>" & vProgramTitle & "</b>. We will be handling your honorarium reimbursement." _
        & "<br><br> <b><u>PROGRAM DETAILS</u></b> " _
        & "<br> " & vRepFullName & " is hosting this program and can be reached at " & vRepPhone & " or via email (copied)." _
        & "<br><br>" & vTab & "<b>Date:</b> " & vLongStartDate _
        & "<br><b>" & vTab & "Time:</b>  " & Format(vArrivalTime, "medium time") & "<font color = 'red'> TIME ZONE</font>" _
        & "<br>" & vTab & "<b>Venue:</b> " & vVenueName _
        & "<br>" & vTab & vTab & vTab & "&nbsp;&nbsp;" & vVenueCity & ", " & vVenueState _
        & "<br>" & vTab & "<b>Web Conference Room and Dial-in:</b>  Please refer to your calendar invite for login and dial-in information for this program." _
        & "<br><br> <b><u>ATTACHED YOU WILL FIND:</u></b> " _
        & "<br>" & vTab & "&bull; System Check & Speaker Tips <br>" & vTab & "&bull; <font color = 'red'>IBD Lunch Presentation/IBS-D Lunch Presentation </font> Slide Deck " _
        & "<br><br> <b><u>AUDIO/VISUAL</u></b> " _
        & "<br>" & vTab & "&bull; A landline phone and hardwired internet is recommended for this presentation." _
        & "<br><br>Best Regards, <br> " & vUserFirstName

        .To = vSpeakerEmail
        .CC = vAsstEmail & ";" & vRepEmail & "; jlyzen@plan365inc.com"
        If Dir(vAttachSpeakerTips) = "" Then
            MsgBox "speaker tips file not found/attached"
        Else
            .Attachments.Add vAttachSpeakerTips
        End If
        .Display

    End With
    
    vRS.MoveNext
Loop

vRS.Close
Set vRS = Nothing
Set vCN = Nothing

End Sub

Sub EnteraCancellationEmail()

vPickQuery = "Standard"
Call EnteraGetData
    If vProgramInput = "" Then
        Exit Sub
    End If
    
Do Until vRS.EOF
    Call EnteraDefineVariables
   
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' build message
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Set vOutlook = CreateObject("outlook.application")
    Set vMessage = vOutlook.CreateItem(vMailItem)
    
    With vMessage
        .Subject = "Entera Health Speaker Program - " & Format(vStartDate, "Long Date") & " - " & vVenueCity & ", " & vVenueState & " - CANCELLATION NOTICE - " & vProgramNumber
        .HTMLBody = "<span LANG=EN>" _
        & "<p class=style2><span LANG=EN><font face=Calibri size=3>" _
        & "Good " & vMessageTime & " " & vSpeakerPrefix & " " & vSpeakerLastName & ", " _
        & "<br><br> We are writing to inform you that the Entera Health " & vProgramType & " program on " & Format(vStartDate, "Long Date") & " in " & vVenueCity & ", " & vVenueState & " has been cancelled." _
        & "<br><br> <font color = 'red'> Your flight / hotel / ground have been cancelled and require no action from you. </font> " _
        & "<br><br> We greatly apologize for any inconvenience this has caused you.  We look forward to working with you on future programs.  Please feel free to contact us with any questions or concerns." _
        & "<br><br> Sincerely, " _
        & "<br> " & vUserFirstName

        .To = vSpeakerEmail
        .CC = vAsstEmail & ";" & vRepEmail & "; jlyzen@plan365inc.com"
        .Display

    End With
    
    vRS.MoveNext
Loop

vRS.Close
Set vRS = Nothing
Set vCN = Nothing

End Sub

Sub EnteraLRepVenueConfEmail()

vPickQuery = "Standard"
Call EnteraGetData
    If vProgramInput = "" Then
        Exit Sub
    End If
    
Do Until vRS.EOF
    Call EnteraDefineVariables
   
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' build message
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Set vOutlook = CreateObject("outlook.application")
    Set vMessage = vOutlook.CreateItem(vMailItem)
    
    With vMessage
        .Subject = "CONFIRMATION:  Speakers Bureau Program - " & vLongStartDate & " - " & vVenueCity & ", " & vVenueState & " - (" & vProgramNumber & ")"
        .HTMLBody = "<span LANG=EN>" _
        & "<p class=style2><span LANG=EN><font face=Calibri size=3>" _
        & "<b>Attachments:</b> Electronic Invitation " _
        & "<br><br> Good " & vMessageTime & " " & vRepFirstName & ", " _
        & "<br><br> We have confirmed your <b> Speakers Bureau " & vProgramType & " Program </b>with<b> " & vSpeakerPrefix & " " & vSpeakerLastName & "</b> on <b>" & vLongStartDate & "</b> at <b>" & vVenueName & "</b> in <b>" & vVenueCity & ", " & vVenueState & "</b>. The title of this program is <b>" & vProgramTitle & "</b>." _
        & "<br><br> <b> <u> INVITATION  </b> </u> " _
        & "<br>An electronic copy of your invitation is attached. You may also download a copy of the invitation on Plan 365's <a href='http://entera.plan365stat.com/'>Stat Sales</a> Portal. " _
        & "<br><br> <b> <u> VENUE DETAILS </b> </u> " _
        & "<br> " & vTab & "Arrival Time: " & vArrivalTime _
        & "<br> " & vTab & "Room Name: " & vRoomName _
        & "<br> " & vTab & "Set For: " & vEstAttendees & " (including you and the speaker) " _
        & "<br> " & vTab & "Guarantee needed by: " & vHeadcountDueDate & " (Plan 365 will provide this to the venue) <br> " & vTab & "Cancellation policy: " & vCancellationNotes _
        & "<br><br> <b> <u> MENU </b> </u> " _
        & "<br>Please let me know if the following menu is acceptable:" _
        & "<br> " & vMenu _
        & "<br><br> <b>**If we have not received changes within 2 days, we will proceed with the above menu.**</b> " _
        & "<br><br> <b> <u> AUDIO VISUAL </b> </u> " _
        & "<br> Per your request, the following equipment has been ordered: " _
        & "<br>Laptop: " & vAVLaptop & "<br> Projector: " & vAVProjector & "<br> Screen: " & vAVScreen & "<br> <font color = 'red'>Other:</font> " _
        & "<br><br> <b> <u> SLIDE DECK</b> </u> " _
        & "<br> You may access the slide decks on the <a href='http://entera.plan365stat.com/'>Stat Sales</a> Portal." _
        & "<br><br> If you have any questions, please contact me at (919) 534-" & vUserExt & ". <br><br> Thank you! <br> " & vUserFirstName
        
        
        .To = vRepEmail
        .CC = "jlyzen@plan365inc.com"
        .Display

    End With
    
    vRS.MoveNext
Loop

vRS.Close
Set vRS = Nothing
Set vCN = Nothing

End Sub

Sub EnteraVRepVenueConfEmail()

vPickQuery = "Standard"
Call EnteraGetData
    If vProgramInput = "" Then
        Exit Sub
    End If
    
Do Until vRS.EOF
    Call EnteraDefineVariables
   
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' build message
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Set vOutlook = CreateObject("outlook.application")
    Set vMessage = vOutlook.CreateItem(vMailItem)
    
    With vMessage
        .Subject = "CONFIRMATION:  Virtual Speakers Bureau Program - " & vLongStartDate & " - " & vVenueCity & ", " & vVenueState & " - (" & vProgramNumber & ")"
        .HTMLBody = "<span LANG=EN>" _
        & "<p class=style2><span LANG=EN><font face=Calibri size=3>" _
        & "<b>Attachments:</b>  Electronic Invitation, System Check, Ipad Login Instructions, Meeting Materials - Full Merge " _
        & "<br><br> Good " & vMessageTime & " " & vRepFirstName & ", " _
        & "<br><br> We have confirmed your <b> Virtual Speakers Bureau Program </b>with<b> " & vSpeakerPrefix & " " & vSpeakerLastName & "</b> on <b>" & vLongStartDate & "</b> at <b>" & vVenueName & "</b> in <b>" & vVenueCity & ", " & vVenueState & "</b>. The title of this program is <b>" & vProgramTitle & "</b>.  We will send you a calendar invitation with the webinar login instructions.  <b>Please perform the system check (attached) prior to the program. </b>" _
        & "<br><br> <b> <u> INVITATION  </b> </u> " _
        & "<br>An electronic copy of your invitation is attached. You may also download a copy of the invitation on Plan 365's <a href='http://entera.plan365stat.com/'>Stat Sales</a> Portal. " _
        & "<table> <tr> <td style='width:150px'> <font face=Calibri size=3> <b><u>VENUE DETAILS</u></b> </font> </td> <td style='width:600px'> </td> </tr> <tr>  <td style='width:150px'> <font face=Calibri size=3> Office Location:</td> </font> <td style='width:600px'> <font face=Calibri size=3> " & vVenueName & " </font> </td> </tr>    <tr> <td style='width:150px'></td> <td style='width:600px'> <font face=Calibri size=3> " & vVenueAddrNoTab & " </font> </td> </tr> <tr> <td style='width:150px'></td> <td style='width:600px'> <font face=Calibri size=3> " & vVenueCity & ", " & vVenueState & " " & vVenueZip & " </font> </td> </tr> " _
        & "<tr> <td style='width:150px'> <font face=Calibri size=3> Presentation Time: </font> </td> <td style='width:600px'> <font face=Calibri size=3> " & vArrivalTime & " </font> </td> </tr> <tr> <td style='width:150px'></td> <td style='width:600px'> <font face=Calibri size=3> <b> **Please plan to login in 20 minutes prior to the program start time** </b> </font> </td> </tr>  </table> " _
        & "<br><br> <b> <u> AUDIO VISUAL </b> </u> " _
        & "<br> Per your request, the following equipment has been ordered: " _
        & "<br>Laptop: " & vAVLaptop & "<br> Projector: " & vAVProjector & "<br> Screen: " & vAVScreen & "<br> <font color = 'red'>Other:</font> " _
        & "<br><br> <b> <u> SLIDE DECK</b> </u> " _
        & "<br> You may access the slide decks on the <a href='http://entera.plan365stat.com/'>Stat Sales</a> Portal." _
        & "<br><br> If you have any questions, please contact me at (919) 534-" & vUserExt & ". <br><br> Thank you! <br> " & vUserFirstName

        .To = vRepEmail
        .CC = "jlyzen@plan365inc.com"
        If Dir(vAttachRepIpadLogin) = "" Then
            MsgBox "ipad steps file not found/attached"
        Else
            .Attachments.Add vAttachRepIpadLogin
        End If
        If Dir(vAttachRepSystemCheck) = "" Then
            MsgBox "system check file not found/attached"
        Else
            .Attachments.Add vAttachRepSystemCheck
        End If
        .Display

    End With
    
    vRS.MoveNext
Loop

vRS.Close
Set vRS = Nothing
Set vCN = Nothing

End Sub

Sub EnteraLRepProgramReminderEmail()

vPickQuery = "Standard"
Call EnteraGetData
    If vProgramInput = "" Then
        Exit Sub
    End If
    
Do Until vRS.EOF
    Call EnteraDefineVariables
   
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' build message
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Set vOutlook = CreateObject("outlook.application")
    Set vMessage = vOutlook.CreateItem(vMailItem)
    
    With vMessage
        .Subject = "REMINDER: Speakers Bureau Program - " & vLongStartDate & " - " & vVenueCity & ", " & vVenueState & " - (" & vProgramNumber & ")"
        .HTMLBody = "<span LANG=EN>" _
        & "<p class=style2><span LANG=EN><font face=Calibri size=3>" _
        & "<b>Attachments:</b> Program Information Sheet, Checklist, Evaluation, Sign-In Sheets, Itinerary (if applicable), Onsite Venue Compliance Document (if applicable) OR Full Meeting Merge " _
        & "<br><br> Hi " & vRepFirstName & ", " _
        & "<br><br> Attached are the following materials for your Speakers Bureau program on <b>" & vLongStartDate & "</b> at <b>" & vVenueName & "</b> in <b>" & vVenueCity & ", " & vVenueState & "</b> with <b>" & vSpeakerPrefix & " " & vSpeakerLastName & "</b>. Please print these documents and bring them with you onsite." _
        & "<br><br> <b> <u>PROGRAM MATERIALS</u></b> " _
        & "<br> You may also download copies of the program materials on Plan 365's <a href='http://entera.plan365stat.com/'>Stat Sales</a> Portal." _
        & "<br> " & vTab & "&bull; Program Information Sheet, Checklist & Attendee Evaluation" _
        & "<br> " & vTab & "&bull; Sign-In Sheets" _
        & "<br> " & vTab & vTab & "&#9702; <i><b>IMPORTANT:</b> All signatures are required." _
        & "<br> " & vTab & vTab & "&#9702; The fields with *asterisks are the fields required in Stat Sales, so please be sure your attendees provide this information." _
        & "<br> " & vTab & vTab & "&#9702; State License # and NPI # are not required for non-HCPs or Residents.</i> " _
        & "<br> " & vTab & "&bull; Program Evaluation: <a href='http://https://www.surveymonkey.com/r/enterahealth2016/'>https://www.surveymonkey.com/r/enterahealth2016</a>." _
        & "<br> " & vTab & "&bull; Speaker's Itinerary <font color = 'red'>(delete if not applicable)</font> " _
        & "<br> " & vTab & "&bull; Onsite Venue Compliance Guidelines  <font color = 'red'>(delete if not applicable)</font> " _
        & "<br><br> <b> <u>SLIDE DECK</u></b>" _
        & "<br>You can access the slide deck by logging in to <a href='http://entera.plan365stat.com/'>Stat Sales</a> and clicking on the ""Resources"" tab." _
        & "<br><br> <b> <u>POST PROGRAM</u></b>" _
        & "<br>Log into <a href='http://entera.plan365stat.com/'>Stat Sales</a> and in the ""Post Program Wrap Up"" section of your completed program, click the ""Launch Program Wrap Up"" button." _
        & "<br><br> If you have any questions regarding your program please contact me at (919) 534-" & vUserExt & "." _
        & "<br><br> Thank you! " _
        & "<br> " & vUserFirstName
    
        .To = vRepEmail
        .CC = "jlyzen@plan365inc.com"
        .Display

    End With
    
    vRS.MoveNext
Loop

vRS.Close
Set vRS = Nothing
Set vCN = Nothing

End Sub

Sub EnteraVRepProgramReminderEmail()

vPickQuery = "Standard"
Call EnteraGetData
    If vProgramInput = "" Then
        Exit Sub
    End If
    
Do Until vRS.EOF
    Call EnteraDefineVariables
   
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' build message
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Set vOutlook = CreateObject("outlook.application")
    Set vMessage = vOutlook.CreateItem(vMailItem)
    
    With vMessage
        .Subject = "REMINDER: Speakers Bureau Program - " & vLongStartDate & " - " & vVenueCity & ", " & vVenueState & " - (" & vProgramNumber & ")"
        .HTMLBody = "<span LANG=EN>" _
        & "<p class=style2><span LANG=EN><font face=Calibri size=3>" _
        & "<b>Attachments:</b> Virtual Meeting Materials - Full Merge, System Check, Ipad Login Instructions, Sign-in Sheets (blank & pre-pop) " _
        & "<br><br> Hi " & vRepFirstName & ", " _
        & "<br><br> Attached are the following materials for your <b>Virtual Speakers Bureau Program </b> on <b>" & vLongStartDate & "</b> at <b>" & vVenueName & "</b> in <b>" & vVenueCity & ", " & vVenueState & "</b> with <b>" & vSpeakerPrefix & " " & vSpeakerLastName & "</b>. The start time for this program is <b>" & vArrivalTime & " <font color = 'red'>TIME ZONE</font></b>.  Please print these documents and bring them with you onsite.</b>" _
        & "<br><br> <b> <u>PROGRAM MATERIALS</u></b> " _
        & "<br> You may also download copies of the program materials on Plan 365's <a href='http://entera.plan365stat.com/'>Stat Sales</a> Portal" _
        & "<br> " & vTab & "&bull; Program Information Sheet, Checklist & Attendee Evaluation" _
        & "<br> " & vTab & "&bull; Sign-In Sheets" _
        & "<br> " & vTab & vTab & "&#9702; <i><b>IMPORTANT:</b> All signatures are required." _
        & "<br> " & vTab & vTab & "&#9702; The fields with *asterisks are the fields required in Stat Sales, so please be sure your attendees provide this information." _
        & "<br> " & vTab & vTab & "&#9702; State License # and NPI # are not required for non-HCPs or Residents.</i> " _
        & "<br> " & vTab & "&bull; Program Evaluation: <a href='http://https://www.surveymonkey.com/r/enterahealth2016/'>https://www.surveymonkey.com/r/enterahealth2016</a>." _
        & "<br><br> <b> <u>WEBINAR ACCESS</u></b>" _
        & "<br> " & vTab & "&bull; A copy of the System Check/Ipad Login Instructions are attached" _
        & "<br> " & vTab & vTab & "&#9702; <i><b>IMPORTANT:</b> You must complete the system check prior to the program</i>" _
        & "<br> " & vTab & "&bull; The webinar login information is located in your meeting calendar request as well as on the Program Information Sheet (attached)." _
        & "<br><br> <b> <u>SLIDE DECK</u></b> " _
        & "<br>You can access the slide deck by logging in to <a href='http://entera.plan365stat.com/'>Stat Sales</a> and clicking on the ""Resources"" tab." _
        & "<br><br> <b> <u>POST PROGRAM</u></b>" _
        & "<br>Log into <a href='http://entera.plan365stat.com/'>Stat Sales</a> and in the ""Post Program Wrap Up"" section of your completed program, click the ""Launch Program Wrap Up"" button." _
        & "<br><br> <b>**If this program was intended for virtual attendees only, Plan 365 will email you a list of participants after the program for attendee portal entry (in lieu of a sign-in sheet).** </b>" _
        & "<br><br> If you have any questions regarding your program please contact me at (919) 534-" & vUserExt & "." _
        & "<br><br> Thank you! <br> " & vUserFirstName
    
        .To = vRepEmail
        .CC = "jlyzen@plan365inc.com"
        If Dir(vAttachRepIpadLogin) = "" Then
            MsgBox "ipad steps file not found/attached"
        Else
            .Attachments.Add vAttachRepIpadLogin
        End If
        If Dir(vAttachRepSystemCheck) = "" Then
            MsgBox "system check file not found/attached"
        Else
            .Attachments.Add vAttachRepSystemCheck
        End If
        .Display

    End With
    
    vRS.MoveNext
Loop

vRS.Close
Set vRS = Nothing
Set vCN = Nothing

End Sub

Sub EnteraFBContract()

vPickQuery = "Standard"
Call EnteraGetData
    If vProgramInput = "" Then
        Exit Sub
    End If
    
Do Until vRS.EOF
    Call EnteraDefineVariables
   
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' build message
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Set vOutlook = CreateObject("outlook.application")
    Set vMessage = vOutlook.CreateItem(vMailItem)
    
    With vMessage
        .Subject = "Entera Health Speakers Bureau Program - " & vLongStartDate & " - Contract - " & vProgramNumber
        .HTMLBody = "<span LANG=EN>" _
        & "<p class=style2><span LANG=EN><font face=Calibri size=3>" _
        & "<b>Attachments:</b> Entera Health Dinner Contract " _
        & "<br><br> Dear " & vVendorContactName & ", " _
        & "<br><br> Attached is the Credit Card Authorization Form for the Entera Health / " & vRepFullName & " " & vProgramType & " program on " & vLongStartDate & " at " & vArrivalTime & " for approximately " & vEstAttendees & " guests. <br> Please note the following: " _
        & "<br><br>" & vTab & " &#9642; Menu Selection: $" & vMealPrice _
        & "<br>" & vTab & " &#9642; <b>2 glasses</b> of beer or wine only--not to exceed <b>$12/drink</b> or $50/bottle. <b>NO CORDIALS/LIQUOR</b> unless cash transaction." _
        & "<br>" & vTab & " &#9642; F&B cost at no more than <b>$100/per person</b> exclusive of tax/gratuity" _
        & "<p align='center'><font size=3>Please let me know if you have any questions.</font></p>" _
        & "<p align='center'><font size=3><b><i>***At the conclusion of the dinner, please fax a deposit receipt and <br> <u>ITEMIZED</u> F&B receipt to (919) 573-9364.***</i></b></font></p> " _
        & "<p align='center'><font size=3><i>**PLEASE CONFIRM RECEIPT OF THIS EMAIL AS SOON IT'S RECEIVED. <br> MY DIRECT LINE IS (919) 534-" & vUserExt & ". </i></font></p> " _
        & "<br><br> Best Regards, " _
        & "<br><br>" & vUserName _
        & "<br>" & vUserRole _
        & "<br> Plan 365, Inc. " _
        & "<br> (919) 534-" & vUserExt _
        & "<br> <a href='mailto:" & vUserAddress & "'>" & vUserAddress & "</a>"

    
        .To = vVendorEmail
        '.CC = "jlyzen@plan365inc.com"
        .Display

    End With
    
    vRS.MoveNext
Loop

vRS.Close
Set vRS = Nothing
Set vCN = Nothing

End Sub

Sub EnteraGetData()

vProgramInput = InputBox("Enter Program Number(s)")
vProgramInput = Replace(vProgramInput, " ", "")
vProgramInput = Replace(vProgramInput, ",", "','")

If vProgramInput = "" Then
    Exit Sub
End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' db connection info (production warehouse)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
vDBHost = "********************"
vDBName = "********************"
vDBUser = "******"
vDBPass = "******"

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' connect to db
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Set vCN = CreateObject("ADODB.Connection")
Set vRS = CreateObject("ADODB.Recordset")
vCN.ConnectionTimeout = 60
vCN.CommandTimeout = 60
vCN.Open "Driver={MySQL ODBC 5.3 Unicode Driver};Server=" & vDBHost & ";Database=" & vDBName & _
";Uid=" & vDBUser & ";Pwd=" & vDBPass & ";"


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' retrieve data
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If vPickQuery = "Standard" Then

    vSQL = "select p.number program_number, p.start_date start_date, p.arrival_time, pt.description program_type, p.title  program_title, p.invitation_count, " _
            & "v.name venue_name, v.straddr1 venue_addr1, v.straddr2 venue_addr2, v.city venue_city, v.state venue_state, v.zip venue_zip, v.phone_primary venue_phone, " _
            & "p.venue_room_name private_room_name, p.estimated_attendees, p.foodbev_final_headcount_due_date, p.cancellation_notes, replace(ifnull(p.foodbev_menu, ''), '\n', '*****') foodbev_menu, p.foodbev_meal_price_per_person, " _
            & "p.av_screen, p.av_flatscreen_tv, p.av_tv_cords, p.av_lcd_projector, p.av_laptop, p.av_microphone, cl.company_name, " _
            & "c1.firstname speaker_firstname, c1.LastName speaker_lastname, c1.email speaker_email, c1.assistant_email, " _
            & "c1.prefix speaker_prefix, c2.firstname rep_firstname, c2.lastname rep_lastname, " _
            & "c2.email rep_email, c2.phone_mobile rep_phone, " _
            & "(select sum(ifnull(amount_actual, 0)) honorarium_amt from ledger where fk_registrant_id = r.id and fk_ledger_category_id = 5) as honorarium_amt, " _
            & "(select fv.contact_email from vendor fv where p.foodbev_fk_vendor_id = fv.id) as vendor_email, " _
            & "(select fv.contact_fullname from vendor fv where p.foodbev_fk_vendor_id = fv.id) as vendor_contact_name " _
            & "from program p, project pj, program_type pt, vendor v, registrant r, speaker_assignment sa, speaker_profile sp, speaker s, contact c1, client_staff_project_assignment cspa, client_staff cs, contact c2, client cl " _
            & "where p.fk_program_type_id = pt.id " _
            & "and p.venue_fk_vendor_id = v.id " _
            & "and r.fk_program_id = p.id " _
            & "and r.fk_speaker_assignment_id = sa.id " _
            & "and sa.fk_speaker_profile_id = sp.id " _
            & "and sp.fk_speaker_id = s.id " _
            & "and s.fk_contact_id = c1.id " _
            & "and p.fk_client_staff_project_assignment_id = cspa.id " _
            & "and cspa.fk_client_staff_id = cs.id " _
            & "and cs.fk_contact_id = c2.id " _
            & "and p.fk_project_id = pj.id " _
            & "and pj.fk_client_id = cl.id " _
            & "  and ifnull(r.fk_attendance_status_id, 0) <> 2 " _
            & "  and p.number in ('" & vProgramInput & "') "
        
End If

If vPickQuery = "Engage" Then

    vSQL = "select p.number program_number, p.start_date start_date, p.arrival_time, pt.description program_type, replace(p.title, ""Â  Weâ€™ve"",  "" We've"") program_title, " _
            & "v.name venue_name, v.straddr1 venue_addr1, v.straddr2 venue_addr2, v.city venue_city, v.state venue_state, v.zip venue_zip, v.phone_primary venue_phone, " _
            & "p.venue_room_name private_room_name, p.estimated_attendees, p.foodbev_final_headcount_due_date, p.cancellation_notes, replace(ifnull(p.foodbev_menu, ''), '\n', '*****') foodbev_menu, p.foodbev_meal_price_per_person, " _
            & "p.av_screen, p.av_flatscreen_tv, p.av_tv_cords, p.av_lcd_projector, p.av_laptop, p.av_microphone, " _
            & "'' as speaker_firstname, " _
            & "'' as speaker_lastname, " _
            & "'' as speaker_email, " _
            & "'' as speaker_prefix, " _
            & "'' as assistant_email, " _
            & "'' as honorarium_amt, " _
            & "'' as invitation_count, " _
            & "c2.firstname rep_firstname, " _
            & "c2.lastname rep_lastname, " _
            & "c2.email rep_email, " _
            & "c2.phone_mobile rep_phone " _
            & "from program p, program_type pt, vendor v, client_staff_project_assignment cspa, client_staff cs, contact c2 " _
            & "where p.fk_program_type_id = pt.id " _
            & "and p.venue_fk_vendor_id = v.id " _
            & "and p.fk_client_staff_project_assignment_id = cspa.id " _
            & "and cspa.fk_client_staff_id = cs.id " _
            & "and cs.fk_contact_id = c2.id " _
            & "  and p.number in ('" & vProgramInput & "') "
        
End If

vRS.CursorLocation = adUseClient
vRS.Open vSQL, vCN

If vRS.recordcount = 0 Then
    MsgBox "Program Number '" & vProgramInput & "' not found." & _
    Chr(13) & "Program doesn't exist or is missing information." & _
    Chr(13) & Chr(13) & "Check that Stat has been updated with" & _
    Chr(13) & "necessary speaker, rep, and venue information."
    Exit Sub
End If

End Sub

Sub EnteraDefineVariables()

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' define variables
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Set vExUser = Application.Session.CurrentUser.AddressEntry.GetExchangeUser

vProgramNumber = vRS.Fields("program_number").Value
vStartDate = vRS.Fields("start_date").Value
vProgramType = vRS.Fields("program_type").Value
vProgramTitle = "Advances In The Management of Enteropathy and other Gastrointestinal Conditions"
vVenueCity = vRS.Fields("venue_city").Value
vVenueState = vRS.Fields("venue_state").Value
vSpeakerFirstName = vRS.Fields("speaker_firstname").Value
vSpeakerLastName = vRS.Fields("speaker_lastname").Value
vSpeakerEmail = vRS.Fields("speaker_email").Value
vSpeakerPrefix = vRS.Fields("speaker_prefix").Value
vArrivalTime = vRS.Fields("arrival_time").Value
vVenueName = vRS.Fields("venue_name").Value
vVenueAddr1 = vRS.Fields("venue_addr1").Value
vVenueAddr2 = vRS.Fields("venue_addr2").Value
vVenueZip = vRS.Fields("venue_zip").Value
vVenuePhone = vRS.Fields("venue_phone").Value
vRepFirstName = vRS.Fields("rep_firstname").Value
vRepLastName = vRS.Fields("rep_lastname").Value
vRepPhone = vRS.Fields("rep_phone").Value
vRepEmail = vRS.Fields("rep_email").Value
vRoomName = vRS.Fields("private_room_name").Value
vEstAttendees = vRS.Fields("estimated_attendees").Value
vHeadcountDueDate = vRS.Fields("foodbev_final_headcount_due_date").Value
vCancellationNotes = vRS.Fields("cancellation_notes").Value
vAVScreen = vRS.Fields("av_screen").Value
vAVFlatScreen = vRS.Fields("av_flatscreen_tv").Value
vAVCords = vRS.Fields("av_tv_cords").Value
vAVProjector = vRS.Fields("av_lcd_projector").Value
vAVLaptop = vRS.Fields("av_laptop").Value
vAVMicrophone = vRS.Fields("av_microphone").Value
vMenu = Replace(vRS.Fields("foodbev_menu").Value, "**********", "<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;")
vMenu = Replace(vMenu, "*****", "<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;")
vMealPrice = vRS.Fields("foodbev_meal_price_per_person").Value
vTab = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
vUserName = vExUser.Name
vUserFirstName = vExUser.FirstName
vUserAddress = vExUser.PrimarySmtpAddress
vUserExt = "2200"
vRepFullName = vRepFirstName & " " & vRepLastName
vLongStartDate = Format(vStartDate, "Long Date")
vAsstEmail = vRS.Fields("assistant_email").Value
vHonorarium = vRS.Fields("honorarium_amt").Value
vInvitationCount = vRS.Fields("invitation_count").Value
vCompanyName = vRS.Fields("company_name").Value
vArrivalTime = Format(vArrivalTime, "medium time")

If Left(Format(vArrivalTime, "medium time"), 1) = 0 Then
    vArrivalTime = Mid(Format(vArrivalTime, "medium time"), 2, 1) & ":" & Mid(Format(vArrivalTime, "medium time"), 4, 2)
Else
    vArrivalTime = Mid(Format(vArrivalTime, "medium time"), 1, 2) & ":" & Mid(Format(vArrivalTime, "medium time"), 4, 2)
End If

vVenueAddrNoTab = vVenueAddr1 & "<br>" & vVenueAddr2
vVendorEmail = vRS.Fields("vendor_email").Value
vVendorContactName = vRS.Fields("vendor_contact_name").Value

If vUserName = "Michael Hodge" Then
    vUserRole = "Data Analyst"
    vUserExt = "2214"
End If

If vUserName = "Jennifer Lyzen" Then
    vUserRole = "Project Manager"
    vUserExt = "2228"
End If

If vUserName = "Stephanie Brown" Then
    vUserRole = "Project Coordinator"
    vUserExt = "2231"
End If

If vUserName = "Denise Schuck" Then
    vUserRole = "Project Manager"
    vUserExt = "2221"
End If

If Time() < #12:00:00 PM# Then
    vMessageTime = "morning"
ElseIf Time() > #12:00:00 PM# Then
    vMessageTime = "afternoon"
End If

vVenueAddr = vVenueAddr1

If Len(vVenueAddr2) > 0 Then
    vVenueAddr = vVenueAddr1 & "<br>" & vTab & vTab & vTab & "&nbsp;&nbsp;" & vVenueAddr2
End If
    
End Sub

Sub test()
 
Dim TimeTest

TimeTest = DateAdd("h", 12, Now())

'MsgBox Now()
MsgBox TimeTest

If Left(Format(TimeTest, "medium time"), 1) = 0 Then
    TimeTest = Mid(Format(TimeTest, "medium time"), 2, 1) & ":" & Mid(Format(TimeTest, "medium time"), 4, 2)
Else
    TimeTest = Mid(Format(TimeTest, "medium time"), 1, 2) & ":" & Mid(Format(TimeTest, "medium time"), 4, 2)
End If
    
MsgBox TimeTest

End Sub

Function IsFile(ByVal fName As String) As Boolean
'Returns TRUE if the provided name points to an existing file.
'Returns FALSE if not existing, or if it's a folder
    On Error Resume Next
    IsFile = ((GetAttr(fName) And vbDirectory) <> vbDirectory)
End Function




