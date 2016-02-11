
option explicit

dim dbHost, dbName, dbUser, dbPass, fileInput, sqlDelete, sqlLoad, sqlRetrieve, objExcel, colIndex, cn, rs

' db connection info                              
dbHost = "*****************"                
dbName = "*****************" 
dbUser = "*******"               
dbPass = "*******"

' path info
fileInput = "S:/Plan 365/IT/Client Support/2015 Clients/Intel/IUP/Assign Points/iup_input.csv"

' query to delete current data from stat_scratchpad.iup_sales table
sqlDelete = "delete from iup_sales"
 
' query to load iup_input.csv to stat_scratchpad.iup_sales table
sqlLoad = "load data local infile '" & fileInput & "' " & _ 
"into table iup_sales " & _ 
"fields terminated by ',' " & _ 
"optionally enclosed by '""' " & _
"lines terminated by '\r\n' " & _
"ignore 1 rows; "
 
' query to retrieve updated data from stat_scratchpad.iup_sales table 
sqlRetrieve = "select " & _
"sales_rep_company 'Sales Rep Company', " & _
"sales_rep 'Sales Rep', " & _
"intel_contact 'Intel Contact', " & _
"po_date 'PO Date', " & _
"po_number 'PO#', " & _
"qty_of_systems_sold 'Quantity of Systems Sold', " & _
"num_processors_in_order '# of Processors in Order', " & _
"vpro 'vPro (Y/N)', " & _
"qty_of_ssd 'Quantity of SSDs', " & _
"product_description 'Description of System or Product', " & _
"intel_processor_number 'Intel Processor number/SKU', " & _
"oem_system_part_number 'OEM System Part Number', " & _
"system_manufacturer_name 'Name of System Manufacturer/OEM', " & _
"ssd_pnt 'Points: General SSD', " & _
"sata_pnt 'Points: SATA SSD', " & _
"pcie_pnt 'Points: PCIE SSD', " & _
"system_pnt 'Points: System', " & _
"hero_pnt 'Points: Hero SKU', " & _
"wireless_docking_pnt 'Points: Wireless Docking Add-On', " & _
"realsense_pnt 'Points: RealSense Add-On', " & _
"vpro_pnt 'Points: vPro Add-On', " & _
"wireless_display_pnt 'Points: Pro Wireless Display Add-On'," & _
"case when reason = '' then 'DNQ' else" & _
"(ssd_pnt + sata_pnt + pcie_pnt + system_pnt + hero_pnt + wireless_docking_pnt + realsense_pnt + vpro_pnt + wireless_display_pnt)*(qty_of_systems_sold) " & _
"end 'Total Points'," & _
"reason 'Reason' " & _
"from " & _
"iup_sales " & _
"order by sales_rep_company, sales_rep"

' connect to db
set cn = CreateObject("ADODB.Connection")
set rs = CreateObject("ADODB.Recordset")
cn.ConnectionTimeout = 60
cn.CommandTimeout = 60
cn.Open "Driver={MySQL ODBC 5.3 Unicode Driver};Server=" & dbHost & ";Database=" & dbName & _
";Uid=" & dbUser & ";Pwd=" & dbPass & ";"

' execute queries
cn.execute (sqlDelete)
cn.execute (sqlLoad)
  cn.execute("call iup_update") ' stored proc
rs.Open sqlRetrieve, cn

' copy results to excel
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = False
objExcel.Workbooks.Add

For colIndex = 0 To rs.Fields.Count - 1
    objExcel.cells(1,1).Offset(0, colIndex).Value = rs.Fields(colIndex).Name
Next

objExcel.Cells(2,1).CopyFromRecordset rs
objExcel.Cells.Font.Name = "Calibri"
objExcel.Cells.Font.Size = 10
objExcel.Range("A1:Y1").Font.Bold = True
objExcel.Cells.Select
objExcel.Selection.ColumnWidth = 100
objExcel.Cells.EntireRow.AutoFit
objExcel.Cells.EntireColumn.AutoFit
objExcel.Range("A1").Select
objExcel.ActiveWindow.SplitColumn = 0
objExcel.ActiveWindow.SplitRow = 1
objExcel.ActiveWindow.FreezePanes = True

objExcel.Visible = True

' cleanup
cn.close
set rs = nothing
set cn = nothing