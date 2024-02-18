Dim vConn As ADODB.Connection
Dim vRS As ADODB.Recordset
Dim vSQL As String
Dim vFile As String
Dim vSheetname As String
Dim vEvent As String
Dim vCellsCount
Dim i

Sub CreateChecklist()

' check if worksheet already exists
For Each sht In ActiveWorkbook.Worksheets

    If sht.Name = "EventCheckIn" Then
        vReponse = MsgBox("Worksheet for Event Check-In already created." & Chr(13) & Chr(13) & "Would you like to delete it and create a new one?", vbYesNo)
       
        If vReponse = vbNo Then
            Exit Sub
        Else
            Worksheets("EventCheckIn").Delete
        End If
       
    End If

Next sht


' catch misc errors
On Error GoTo CatchError

Application.ScreenUpdating = False

vFile = ActiveWorkbook.FullName
vEvent = ActiveCell.Value

' create CheckList tab
Sheets.Add(Before:=Sheets(1)).Name = "EventCheckIn"
Sheets("EventCheckIn").Select
ActiveWindow.DisplayGridlines = False


' pull data from export into CheckList tab
Set vConn = CreateObject("ADODB.Connection")
Set vRS = CreateObject("ADODB.Recordset")

vConn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & vFile _
    & ";Extended Properties=""Excel 12.0;HDR=Yes;IMEX=1"";"

vSQL = "select ucase([First name]), ucase([Last name]), Order, Quantity " & _
       "from [Sheet2$A1:V10000] " & _
       "order by [Last name], [First name] "

vRS.Open vSQL, vConn

Sheets("EventCheckIn").Select
Range("A5").CopyFromRecordset vRS
   
vRS.Close
vConn.Close


' create header row on CheckList tab
Set vConn = CreateObject("ADODB.Connection")
Set vRS = CreateObject("ADODB.Recordset")

vConn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & vFile _
    & ";Extended Properties=""Excel 12.0;HDR=Yes;IMEX=1"";"

vSQL = "select 'First Name', 'Last Name', 'Order ID', '# Tickets', ' ' "

vRS.Open vSQL, vConn

Sheets("EventCheckIn").Select
Range("A4").CopyFromRecordset vRS

vRS.Close
vConn.Close

Set vRS = Nothing
Set vConn = Nothing


' create title row on CheckList tab
Sheets("EventCheckIn").Range("A1").Value = "Event Check-In"
Sheets("EventCheckIn").Range("A2").Formula = "=Sheet2!e2"

Sheets("EventCheckIn").Range("A1:E2").Borders(xlEdgeLeft).Weight = xlMedium
Sheets("EventCheckIn").Range("A1:E2").Borders(xlEdgeRight).Weight = xlMedium
Sheets("EventCheckIn").Range("A1:E2").Borders(xlEdgeTop).Weight = xlMedium
Sheets("EventCheckIn").Range("A1:E2").Borders(xlEdgeBottom).Weight = xlMedium
Sheets("EventCheckIn").Range("A1:E2").Interior.ThemeColor = xlThemeColorDark1
Sheets("EventCheckIn").Range("A1:E2").Interior.TintAndShade = -0.149998474074526


' add checkboxes
With Worksheets("EventCheckIn")
    vCellsCount = .Range("A5", .Range("A5").End(xlDown)).Cells.Count
End With

For i = 5 To vCellsCount + 4

Worksheets("EventCheckIn").CheckBoxes.Add(Cells(i, "E").Left, Cells(i, "E").Top, 72, 17.25).Select
With Selection
    .Caption = ""
    .Value = xlOff '
    .Display3DShading = False
End With

Next

' add links
Dim vCurCell As Range
Dim vCurCellVal As String

With Worksheets("EventCheckIn")
    vCellsCount = .Range("A5", .Range("A5").End(xlDown)).Cells.Count
End With

For i = 5 To vCellsCount + 4
    Set vCurCell = Worksheets("EventCheckIn").Cells(i, 3)
    vCurCellVal = vCurCell.Text
    vCurCell.Hyperlinks.Add Anchor:=vCurCell, Address:="https://app.amilia.com/Clients/en/theartscenter/Invoice/" & vCurCellVal, TextToDisplay:=vCurCellVal
Next i


' format CheckList tab

With Sheets("EventCheckIn")
  .Cells.Font.Name = "Calibri"
  .Cells.Font.Size = 10
End With


Rows("1:4").Select
Selection.Font.Bold = True
With ActiveWindow
    .SplitColumn = 0
    .SplitRow = 1
End With

Cells.Select
Cells.EntireColumn.AutoFit
Columns("A:E").Select
Selection.ColumnWidth = 20
Cells.Select
With Selection
    .HorizontalAlignment = xlLeft
    .VerticalAlignment = xlBottom
    .WrapText = False
    .Orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .ReadingOrder = xlContext
    .MergeCells = False
End With

Columns("E:E").ColumnWidth = 10

Range("A4").Select
Range(Selection, Selection.End(xlToRight)).Select
Range(Selection, Selection.End(xlDown)).Select

Selection.Borders(xlEdgeBottom).Weight = xlThin
Selection.Borders(xlInsideHorizontal).Weight = xlThin


Worksheets("EventCheckIn").Cells.Select
Selection.RowHeight = 18

Worksheets("EventCheckIn").Range("F1").Select
ActiveWindow.FreezePanes = True


Application.ScreenUpdating = True

Exit Sub


CatchError:
Application.ScreenUpdating = True

MsgBox ("shoot." & Chr(13) & Chr(13) & "sorry.. something went wrong :( ")

End Sub

Sub CreateSummary()

' check if worksheet already exists
For Each sht In ActiveWorkbook.Worksheets

    If sht.Name = "EventSummary" Then
        vReponse = MsgBox("Worksheet for Event Summary already created." & Chr(13) & Chr(13) & "Would you like to delete it and create a new one?", vbYesNo)
       
        If vReponse = vbNo Then
            Exit Sub
        Else
            Worksheets("EventSummary").Delete
        End If
       
    End If

Next sht

' catch misc errors
On Error GoTo CatchError

Application.ScreenUpdating = False


vFile = ActiveWorkbook.FullName
vEvent = ActiveCell.Value

' create Summary tab
Sheets.Add(Before:=Sheets(1)).Name = "EventSummary"
Sheets("EventSummary").Select
ActiveWindow.DisplayGridlines = False

With Sheets("EventSummary")
  .Cells.Font.Name = "Calibri"
  .Cells.Font.Size = 12
End With


' format and insert formulas to total amounts from export data
Sheets("EventSummary").Range("A1").Value = "Event Sales Summary Report"
Sheets("EventSummary").Range("A2").Formula = "=Sheet2!e2"

Sheets("EventSummary").Range("A6").Formula = "Gross Sales:"
Sheets("EventSummary").Range("A7").Formula = "Discounts:"
Sheets("EventSummary").Range("A8").Formula = "Cancellations:"
Sheets("EventSummary").Range("A9").Formula = "Net Sales:"
Sheets("EventSummary").Range("A10").Formula = "Sales Tax:"
Sheets("EventSummary").Range("A14").Formula = "Total:"

Sheets("EventSummary").Range("D5").Formula = "(#)"
Sheets("EventSummary").Range("E5").Formula = "($)"

Sheets("EventSummary").Range("D6").Formula = "=Sum(Sheet2!o:o)"                                     'Gross Sales
Sheets("EventSummary").Range("D7").Formula = "=Sum(Sheet2!r:r)"                                     'Discounts
Sheets("EventSummary").Range("D8").Formula = "=SUMIF(Sheet2!J:J,""Cancelled"",Sheet2!O:O)*-1"       'Cancellations
Sheets("EventSummary").Range("D9").Formula = "=SUMIF(Sheet2!J:J,""<>Cancelled"",Sheet2!O:O)"        'Net Sales


Sheets("EventSummary").Range("E6").Formula = "=Sum(Sheet2!q:q)"                                     'Gross Sales
Sheets("EventSummary").Range("E7").Formula = "=SUM(Sheet2!s:s)"                                     'Discounts
Sheets("EventSummary").Range("E8").Formula = "=SUMIF(Sheet2!J:J,""Cancelled"",Sheet2!q:q)*-1"       'Cancellations
Sheets("EventSummary").Range("E9").Formula = "=SUMIF(Sheet2!j:j,""<>Cancelled"",Sheet2!t:t)"        'Net Sales
Sheets("EventSummary").Range("E10").Formula = "=SUMIF(Sheet2!J:J,""<>Cancelled"",Sheet2!u:u)"       'Sales Tax

Sheets("EventSummary").Range("A14").Formula = "Total:"
Sheets("EventSummary").Range("E14").Formula = "=E9+E10"                                             'Total

Sheets("EventSummary").Range("A17").Formula = "Performer Split:"
Sheets("EventSummary").Range("E17").Formula = "=E9*D17"                                             'Performer Split

Sheets("EventSummary").Cells.Font.Name = "Calibri"
Sheets("EventSummary").Cells.Font.Size = 14

Sheets("EventSummary").Columns("A:A").ColumnWidth = 55
Sheets("EventSummary").Columns("B:Z").ColumnWidth = 15

Sheets("EventSummary").Columns("D:E").HorizontalAlignment = xlRight
Sheets("EventSummary").Columns("E:E").NumberFormat = "#,##0.00;[Red]#,##0.00"

Sheets("EventSummary").Range("A2:E2").MergeCells = True
Sheets("EventSummary").Range("A2:E2").WrapText = True

Sheets("EventSummary").Range("A2").Font.Bold = True
Sheets("EventSummary").Range("D5:E5").Font.Bold = True
Sheets("EventSummary").Range("A14:E17").Font.Bold = True

Sheets("EventSummary").Range("A1:E2").Borders(xlEdgeLeft).Weight = xlMedium
Sheets("EventSummary").Range("A1:E2").Borders(xlEdgeRight).Weight = xlMedium
Sheets("EventSummary").Range("A1:E2").Borders(xlEdgeTop).Weight = xlMedium
Sheets("EventSummary").Range("A1:E2").Borders(xlEdgeBottom).Weight = xlMedium
Sheets("EventSummary").Range("A1:E2").Interior.ThemeColor = xlThemeColorDark1
Sheets("EventSummary").Range("A1:E2").Interior.TintAndShade = -0.149998474074526

Sheets("EventSummary").Range("A14:E14").Borders(xlEdgeTop).Weight = xlMedium

Sheets("EventSummary").Range("A17:E17").Borders(xlEdgeTop).Weight = xlMedium
Sheets("EventSummary").Range("A17:E17").Borders(xlEdgeBottom).Weight = xlMedium

Sheets("EventSummary").Range("D17").NumberFormat = "0%"
Sheets("EventSummary").Range("D17").Interior.Color = 4716535

Sheets("EventSummary").Range("A50").Formula = "Reported Generated: " & Date & " @ " & Time

Sheets("EventSummary").PageSetup.FitToPagesWide = 1
Sheets("EventSummary").PageSetup.FitToPagesTall = 1
Sheets("EventSummary").PageSetup.Zoom = False

Worksheets("EventSummary").Range("F1").Select

Application.ScreenUpdating = True

Exit Sub

CatchError:
Application.ScreenUpdating = True

MsgBox ("shoot." & Chr(13) & Chr(13) & "sorry.. something went wrong :( ")

End Sub