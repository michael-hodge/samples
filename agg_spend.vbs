option explicit
dim dbHost, dbName, dbUser, dbPass, fileSQL, sqlQuery, sqlQueryEnd, objExcel, colIndex, vStartDate, vEndDate, vProgramNumLabel, vSortField, vClient, vClientName, vVerify, vFile, objShell, CurDir, cn, rs, fso, i, x

set objShell = WScript.CreateObject ("WScript.Shell")
CurDir = (objShell.CurrentDirectory)

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' db connection info (production)                          
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
dbHost = "******************"                
dbName = "******************" 
dbUser = "******"               
dbPass = "******"

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' connect to db
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
set cn = CreateObject("ADODB.Connection")
set rs = CreateObject("ADODB.Recordset")
cn.ConnectionTimeout = 60
cn.CommandTimeout = 60
cn.Open "Driver={MySQL ODBC 5.3 Unicode Driver};Server=" & dbHost & ";Database=" & dbName & _
";Uid=" & dbUser & ";Pwd=" & dbPass & ";"

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' collect client and run date variables from input boxes
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Do Until vVerify = vbYes
	' client input box
	vClient = InputBox("Client:"  & (Chr(13)) & (Chr(13)) & "1: Impax"  & (Chr(13)) & "2: Insys"  & (Chr(13)) & "3: Kaleo" & (Chr(13)) & "4: Arbor" & (Chr(13)) & "5: UT" & (Chr(13)) & "6: Pernix", "Input" ,, 100, 100)

	select case vClient
	case "1"
		vClientName = "Impax"
		vFile = "impax_agg_spend_query_view.sql"
		vProgramNumLabel = "Event ID"
		vSortField = "program_number"
	case "2"
		vClientName = "Insys"
		vFile = "insys_agg_spend_query_view.sql"
		vProgramNumLabel = "Meeting ID Number"
		vSortField = "program_number"
	case "3"
		vClientName = "Kaleo"
		vFile = "kaleo_agg_spend_query_view.sql"
		vProgramNumLabel = "Event ID/Number"
		vSortField = "program_number"
	case "4"
		vClientName = "Arbor"
		vFile = "arbor_agg_spend_query_view.sql"
		vProgramNumLabel = "Program Number (x)"
		vSortField = "date_of_payment, program_number"
	case "5"
		vClientName = "UT"
		vFile = "ut_agg_spend_query_view.sql"
		vProgramNumLabel = "Program Number"
		vSortField = "date_of_payment, program_number"
	case "6"
		vClientName = "Pernix"
		vFile = "pernix_agg_spend_query_view.sql"
		vProgramNumLabel = "Program Number"
		vSortField = "date_of_payment, program_number"
	case ""
		wscript.quit
	case else 
		msgbox "invalid input"
		wscript.quit
	end select
	
	' start date input box
	vStartDate = "'" & InputBox("Start Date:", "Input", "2015-01-01", 100, 100) & "'"
		if vStartDate = "''" then wscript.quit
	
	' end date input box
	vEndDate = "'" & InputBox("End Date:", "Input", "2015-12-31", 100, 100) & "'"
			if vEndDate = "''" then wscript.quit

	' verify parameters message box			
	vVerify = MsgBox ("client:  " & vClientName & (Chr(13)) & "start date:  " & vStartDate & (Chr(13)) & "end date:  " & vEndDate, vbYesNo, "verify parameters:")

Loop

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' loop thru once for completed programs and again for cancelled/rescheduled.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
For i = 1 to 2
	
	if i = 1 then
		sqlQueryEnd = " and program_date between " & vStartDate & " and " & vEndDate & " and program_status_id not in (3,4,5,8,7,20,21,22) order by " & vSortField
	end if
	
	if i = 2 then
		sqlQueryEnd = " and program_date between " & vStartDate & " and " & vEndDate & " and program_status_id in (3,4,5,8,7,20,21,22) order by " & vSortField
	end if
	
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' read query file.  append additional parameters.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set fileSQL = fso.OpenTextFile(CurDir & "\sql\" & vFile, 1)
	sqlQuery = fileSQL.ReadAll()
	sqlQuery = sqlQuery & sqlQueryEnd
	
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' retrieve dataset
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	rs.Open sqlQuery, cn
	
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' copy dataset to excel
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	if i = 1 then
		Set objExcel = CreateObject("Excel.Application")
		objExcel.Visible = False
		objExcel.Workbooks.Add
		objExcel.Sheets("Sheet1").Name = "Complete"
		objExcel.Sheets("Complete").Select
	end if
	
	if i = 2 then
		objExcel.Sheets("Sheet2").Name = "Cancel_Reschedule"
		objExcel.Sheets("Cancel_Reschedule").Select
	end if

	For colIndex = 0 To rs.Fields.Count - 1
		objExcel.cells(1,1).Offset(0, colIndex).Value = rs.Fields(colIndex).Name
	Next

	objExcel.Cells(2,1).CopyFromRecordset rs
	

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' format excel
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''	
	objExcel.Cells.Font.Name = "Calibri"
	objExcel.Cells.Font.Size = 10
	objExcel.Range("1:1").Font.Bold = True
	objExcel.Cells.Select
	objExcel.Selection.ColumnWidth = 100
	objExcel.Cells.EntireRow.AutoFit
	objExcel.Cells.EntireColumn.AutoFit
	objExcel.Range("A1").Select
	objExcel.ActiveWindow.SplitColumn = 0
	objExcel.ActiveWindow.SplitRow = 1
	objExcel.ActiveWindow.FreezePanes = True
	
	colIndex = 1
	Do While Not isempty(objExcel.Cells(colIndex, 1).Value)
		For x = 1 To 100
			if instr(objExcel.Cells(colIndex, x).Value, "(x)") > 0 then
				objExcel.Cells(colIndex, x).EntireColumn.Font.Color = vbRed
			end if
		Next
		colIndex = colIndex + 1
	Loop
	objExcel.Range("A1").Select
	
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' tag program.report_date with current date
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'	if rs.recordcount > 0 then
		rs.movefirst
		do until rs.eof
			cn.execute("update program set report_date = now() where number = " & "'" & rs.Fields(vProgramNumLabel).value) & "'"
			rs.movenext()
		loop
''	end if
'
	rs.close

Next
	
objExcel.Sheets("Sheet3").Delete
objExcel.Sheets("Complete").Select
objExcel.Visible = True

' cleanup
cn.close
set rs = nothing
set cn = nothing