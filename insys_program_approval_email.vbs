option explicit
dim fso, vLogFile, vNamespace, shell, vDBHost, vDBName, vDBUser, vDBPass, vSQL, vCN, vRS, vOutlook, vMessage, vMailItem, vProgramNumber, vProgramDate, vRepName, vDistrict, vRecordCount, vErrorStep, vErrorDesc

'on error resume next 

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' open logfile
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
set fso = CreateObject("Scripting.FileSystemObject")
set vLogFile = fso.OpenTextFile("\\PLAN365FILE\Shared\Plan 365\IT\Email\Insys\insys_program_approval_logfile.txt", 8)	
		vLogFile.WriteLine("------------------------------------------")
		vLogFile.WriteLine("task started:  " & now())

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' open outlook
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Set vOutlook = createobject("Outlook.Application")
		vLogFile.WriteLine("outlook opened:  " & now())

 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' db connection info (production warehouse)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
vDBHost = "**********************"
vDBName = "**********************"
vDBUser = "********"
vDBPass = "********"

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
        & "concat('Rep Name:  ', c.firstname, ' ', c.lastname, '          ') rep_name, " _
        & "concat('Program Date:  ', date_format(p.start_date, '%M %d, %Y'), '          ') program_date, " _
        & "concat('District:  ', cspa.region, '          ') district " _
        & "from " _
        & "program p, project pj, client_staff_project_assignment cspa, client_staff cs, contact c, program_status ps, program_status_explanation pse " _
        & "where " _
        & "p.fk_project_id = pj.id " _
        & "and p.fk_client_staff_project_assignment_id = cspa.id " _
        & "and cspa.fk_client_staff_id = cs.id " _
        & "and cs.fk_contact_id = c.id " _
        & "and p.fk_program_status_id = ps.id " _
        & "and p.fk_program_status_explanation_id = pse.id " _
        & "  and pj.code in ('INSSB15', 'INSSB16') " _
		& "  and p.fk_program_status_id = 1 " _
		& "  and p.fk_program_status_explanation_id = 1) x " 
		
	'& "  and p.number in ('8035JA2115E', '8223JA0715W', '8035JA1915E')) x "				
				
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
        & "concat('Rep Name:  ', c.firstname, ' ', c.lastname) rep_name, " _
        & "concat('Program Date:  ', date_format(p.start_date, '%M %d, %Y')) program_date, " _
        & "concat('District:  ', cspa.region) district " _
        & "from " _
        & "program p, project pj, client_staff_project_assignment cspa, client_staff cs, contact c, program_status ps, program_status_explanation pse " _
        & "where " _
        & "p.fk_project_id = pj.id " _
        & "and p.fk_client_staff_project_assignment_id = cspa.id " _
        & "and cspa.fk_client_staff_id = cs.id " _
        & "and cs.fk_contact_id = c.id " _
        & "and p.fk_program_status_id = ps.id " _
        & "and p.fk_program_status_explanation_id = pse.id " _
        & "  and pj.code = 'INSSB15' " _
		& "  and p.fk_program_status_id = 1 " _
		& "  and p.fk_program_status_explanation_id = 1 " 
		
		'& "  and p.number in ('8035JA2115E', '8223JA0715W', '8035JA1915E') "
					
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
set vMessage = vOutlook.CreateItem(vMailItem)
set vMessage.SendUsingAccount = vOutlook.Session.Accounts.item(1) 
		vLogFile.WriteLine("opening outlook:  " & now())
    
with vMessage
	'.To = "mhodge@plan365inc.com"
    .To = "ISP@insysrx.com"
    .CC = "insys@plan365inc.com"
	.BCC = "mhodge@plan365inc.com"
    .Subject = "Programs Awaiting Approval in Plan365 Portal"
    .HTMLBody = "<span LANG=EN>" _
                & "<p class=style2><span LANG=EN><font face=Calibri size=3>" _
                & "The following programs have been submitted by the field and are pending final review and approval online before the Plan365 team can begin planning. " _
                & "<br><br> " _
                & "<table width=""80%""><font face=Calibri size=3><tr><td>" & vRS.GetString(2, , "</td><td>", "</td></tr><tr><td>") & "</td></tr></font></table> " _
                & "</font></p> "
	.Send
end with
    
vRS.Close
set vRS = Nothing
set vCN = Nothing

if Err.Number <> 0 then
  vErrorStep = "Build Message"
  vErrorDesc = Err.Description
		vLogFile.WriteLine("ERROR: " & vErrorStep & "  " & vErrorDesc & "  " & now())
  Err.Clear
  call ErrEmail
  wscript.quit
end if


vOutlook.quit

		vLogFile.WriteLine("message sent:  " & now())
		vLogFile.Close

wscript.quit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' error email
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sub ErrEmail()

set vMessage = vOutlook.CreateItem(vMailItem)
set vMessage.SendUsingAccount = vOutlook.Session.Accounts.item(1) 
    
with vMessage
	'.To = "mhodge@plan365inc.com"
    .To = "insys@plan365inc.com; mhodge@plan365inc.com"
    .Subject = "ERROR: Insys Program Approval Email"
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

set vMessage = vOutlook.CreateItem(vMailItem)
set vMessage.SendUsingAccount = vOutlook.Session.Accounts.item(1) 
    
with vMessage
	'.To = "mhodge@plan365inc.com"
    .To = "insys@plan365inc.com; mhodge@plan365inc.com"
    .Subject = "NO RESULTS: Insys Program Approval Email"
	.Body = "no pending programs that need client attention."
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