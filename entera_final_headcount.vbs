option explicit

dim fso, vLogFile, vDBHost, vDBName, vDBUser, vDBPass, vSQL, vCN, vRS, vOutlook, vMessage, vMailItem, vProgramNumber, vLongProgramDate, vRepName, vRepEmail,vVenueName, vSpeakerName, vArrivalTime, vRecordCount, vErrorStep, vErrorDesc

'on error resume next 

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' open logfile
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
set fso = CreateObject("Scripting.FileSystemObject")
set vLogFile = fso.OpenTextFile("\\PLAN365FILE\Shared\Plan 365\IT\Email\Entera\entera_final_headcount_logfile.txt", 8)
		vLogFile.WriteLine("------------------------------------------")
		vLogFile.WriteLine("task started:  " & now())

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' db connection info (production warehouse)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
vDBHost = "*********************"
vDBName = "*********************"
vDBUser = "************"
vDBPass = "************"

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' connect to db
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
set vCN = CreateObject("ADODB.Connection")
set vRS = CreateObject("ADODB.Recordset")
vCN.ConnectionTimeout = 60
vCN.CommandTimeout = 60
vCN.Open "Driver={MySQL ODBC 5.3 Unicode Driver};Server=" & vDBHost & ";Database=" & vDBName & _
";Uid=" & vDBUser & ";Pwd=" & vDBPass & ";"

		vLogFile.WriteLine("connect to db:  " & now())
 
if Err.Number <> 0 then
  vErrorStep = "Connect to Database"
  vErrorDesc = Err.Description
		vLogFile.WriteLine("ERROR: " & vErrorStep & "  " & vErrorDesc & "  " & now())
  Err.Clear
  call ErrEmail
  wscript.quit
end if

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' retrieve data
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
vSQL = "select count(*) from (select " _
        & " pj.code project_code, " _
        & " p.number program_number, " _
        & " date_format(p.start_date, '%M %d, %Y') long_program_date, " _
        & " p.start_date program_date, " _
		& " time_format(p.arrival_time, '%l:%i') arrival_time, " _
        & " date_add(p.start_date, interval -7 day) one_week_prior, " _
        & " pt.description program_type, " _
        & " c1.email rep_email, " _
        & " concat(c1.firstname, ' ', c1.lastname) rep_name, " _
        & " v.name venue_name, " _
        & " concat(c2.prefix, ' ', c2.lastname) speaker_name " _
        & " from " _
        & " project pj, program p, program_type pt, client_staff_project_assignment cspa, client_staff cs, contact c1, vendor v,  " _
		& " registrant r, speaker_assignment sa, speaker_profile sp, speaker s, contact c2 " _
        & " where " _
        & " p.fk_project_id = pj.id " _
        & " and p.fk_program_type_id = pt.id " _
        & " and p.fk_client_staff_project_assignment_id = cspa.id " _
        & " and cspa.fk_client_staff_id = cs.id " _
        & " and cs.fk_contact_id = c1.id " _
        & " and p.venue_fk_vendor_id = v.id " _
        & " and r.fk_program_id = p.id " _
        & " and r.fk_speaker_assignment_id = sa.id " _
        & " and sa.fk_speaker_profile_id = sp.id " _
        & " and sp.fk_speaker_id = s.id " _
        & " and s.fk_contact_id = c2.id " _
        & "   and pj.code in ('ENTSB15', 'ENTSB16') " _
        & "   and pt.description = 'Dinner' " _
		& "  and curdate() = date_add(p.start_date, interval -7 day)) x " 
        
		'& "   and p.number in ('1620NV1215P', '1640NV1915')) x " 
        
vRS.CursorLocation = 3        
vRS.CursorType = 3		
vRS.Open vSQL, vCN
vRecordCount = vRS.fields(0).value
vRS.close

if vRecordCount = "0" then
		vLogFile.WriteLine("checked for records:  " & now())
		vLogFile.WriteLine("no records found.  calling NoResultsEmail:  " & now())
	call NoResultsEmail
    'wscript.quit
end if

vSQL = "select " _
        & " pj.code project_code, " _
        & " p.number program_number, " _
        & " date_format(p.start_date, '%M %d, %Y') long_program_date, " _
        & " p.start_date program_date, " _
		& " time_format(p.arrival_time, '%l:%i') arrival_time, " _
        & " date_add(p.start_date, interval -7 day) one_week_prior, " _
        & " pt.description program_type, " _
        & " c1.email rep_email, " _
        & " concat(c1.firstname, ' ', c1.lastname) rep_name, " _
        & " v.name venue_name, " _
        & " concat(c2.prefix, ' ', c2.lastname) speaker_name " _
        & " from " _
        & " project pj, program p, program_type pt, client_staff_project_assignment cspa, client_staff cs, contact c1, vendor v,  " _
		& " registrant r, speaker_assignment sa, speaker_profile sp, speaker s, contact c2 " _
        & " where " _
        & " p.fk_project_id = pj.id " _
        & " and p.fk_program_type_id = pt.id " _
        & " and p.fk_client_staff_project_assignment_id = cspa.id " _
        & " and cspa.fk_client_staff_id = cs.id " _
        & " and cs.fk_contact_id = c1.id " _
        & " and p.venue_fk_vendor_id = v.id " _
        & " and r.fk_program_id = p.id " _
        & " and r.fk_speaker_assignment_id = sa.id " _
        & " and sa.fk_speaker_profile_id = sp.id " _
        & " and sp.fk_speaker_id = s.id " _
        & " and s.fk_contact_id = c2.id " _
        & "   and pj.code in ('ENTSB15', 'ENTSB16') " _
        & "   and pt.description = 'Dinner' " _
		& "   and curdate() = date_add(p.start_date, interval -7 day) " 
       
	   '& "   and p.number in ('1620NV1215P', '1640NV1915') " 
        
vRS.CursorLocation = 3        
vRS.CursorType = 3
vRS.Open vSQL, vCN

		vLogFile.WriteLine("retrieved records:  " & now())

if Err.Number <> 0 then
  vErrorStep = "Retrieve Data"
  vErrorDesc = Err.Description
		vLogFile.WriteLine("ERROR: " & vErrorStep & "  " & vErrorDesc & "  " & now())
  Err.Clear
  call ErrEmail
  wscript.quit
end if

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' build message
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
vRS.movefirst
do until vRS.eof

	' define variables
	vLongProgramDate = vRS.Fields("long_program_date").Value
	vRepEmail = vRS.Fields("rep_email").Value
	vRepName = vRS.Fields("rep_name").Value
	vVenueName = vRS.Fields("venue_name").Value
	vSpeakerName = vRS.Fields("speaker_name").Value
	vArrivalTime = vRS.Fields("arrival_time").Value
	vProgramNumber = vRS.Fields("program_number").Value

	'set vOutlook = GetObject(, "outlook.application")
	set vOutlook = CreateObject("outlook.application")
	set vMessage = vOutlook.CreateItem(vMailItem)
	set vMessage.SendUsingAccount = vOutlook.Session.Accounts.item(3)
		vLogFile.WriteLine("opening outlook:  " & now())
		
	with vMessage
		'.To = "mhodge@plan365inc.com"
		.To = vRepEmail
		.CC = "enterahealth@plan365inc.com"
		.BCC = "mhodge@plan365inc.com"
		.Subject = "ACTION REQUIRED:  Final Head Count Due - " & vLongProgramDate & " - " & vProgramNumber
		.HTMLBody = "<span LANG=EN>" _
					& "<p class=style2><span LANG=EN><font face=Calibri size=3> " _
					& "Dear " & vRepName & "," _
					& "<br><br> " _
					& "Please provide a final attendee headcount within 24 hours of receiving this email for your upcoming dinner program on " _
					& vLongProgramDate & " at " & vVenueName & " with " & vSpeakerName &  " at " & vArrivalTime & ".  The total number should " _ 
					& "include yourself, other Entera Health staff and the speaker.  Headcounts can be called in to (919) 534-2231." _
					& "<br><br> " _
					& "Reminder:  If you do not have 3 HCPs confirmed to attend and in the portal 4 days prior to the program, your program " _
					& "will cancel. "  _
					& "Please feel free to contact us with any questions." _
					& "<br><br> " _
					& "Best Regards," _
					& "<br> " _
					& "Plan 365, Inc." _
					& "<br> " _
					& "(919) 534-2200"
		.Send
	end with
		
	if Err.Number <> 0 then
	  vErrorStep = "Build Message"
	  vErrorDesc = Err.Description
			vLogFile.WriteLine("ERROR: " & vErrorStep & "  " & vErrorDesc & "  " & now())
	  Err.Clear
	  call ErrEmail
	  wscript.quit
	end if

vRS.MoveNext
loop

vRS.Close
set vRS = Nothing
set vCN = Nothing

vOutlook.quit

		vLogFile.WriteLine("message sent:  " & now())
		vLogFile.Close

wscript.quit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' error email
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sub ErrEmail()

'set vOutlook = GetObject(, "outlook.application")
set vOutlook = CreateObject("outlook.application")
set vMessage = vOutlook.CreateItem(vMailItem)
set vMessage.SendUsingAccount = vOutlook.Session.Accounts.item(3) 
    
with vMessage
	'.To = "mhodge@plan365inc.com"
    .To = "mhodge@plan365inc.com; enterahealth@plan365inc.com"
    .Subject = "ERROR: Entera Headcount Email"
	.Body = "Error on Step: " & vErrorStep & Chr(13) & "Description: " & vErrorDesc
	.Send
end with
    
'vRS.Close
set vRS = Nothing
set vCN = Nothing

vOutlook.quit

		vLogFile.WriteLine("error message sent:  " & now())
		vLogFile.Close

wscript.quit

end sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' no results email
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sub NoResultsEmail()

'set vOutlook = GetObject(, "outlook.application")
set vOutlook = CreateObject("outlook.application")
set vMessage = vOutlook.CreateItem(vMailItem)
set vMessage.SendUsingAccount = vOutlook.Session.Accounts.item(3) 
    
with vMessage
	'.To = "mhodge@plan365inc.com"
    .To = "mhodge@plan365inc.com; enterahealth@plan365inc.com"
    .Subject = "NO RESULTS: Entera Final Headcount Email"
	.Body = "no programs scheduled 7 days out from today."
	.Send
end with
    
'vRS.Close
set vRS = Nothing
set vCN = Nothing

vOutlook.quit

		vLogFile.WriteLine("no results message sent:  " & now())
		vLogFile.Close

wscript.quit

end sub