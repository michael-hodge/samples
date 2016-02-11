Sub Grifols_Status_Format()

Dim c As Range
Dim x As Integer
Dim y As Integer
Dim RowCnt As Integer
Dim TotalCost As Double
Dim CombRowCnt As Integer
Dim CombTotalCost As Double
Dim DropLoc As String
Dim ProjFileName As String
Dim ExportFileName As String
Dim ProjArray(1 To 9) As String
    ProjArray(1) = "GRIABU15"
    ProjArray(2) = "GRIALP15"
    ProjArray(3) = "GRIGAM15"
    ProjArray(4) = "GRIGPH15"
    ProjArray(5) = "GRIPRO15"
    ProjArray(6) = "GRITHR15"
    ProjArray(7) = "GRIGNE15"
    ProjArray(8) = "GRIALP15_NE"
    ProjArray(9) = "GRIGAM15_NE"
    
Application.ScreenUpdating = False
Application.DisplayAlerts = False

ExportFileName = ActiveWorkbook.Name
DropLoc = "C:\grifols_status_test\"
'DropLoc = "S:\Customers - Active\Grifols 2015\ALL BRANDS - General Info\Reports - Status Reports\"
    
Cells.Select
Cells.Font.Name = "Calibri"
Cells.Font.Size = 10
Cells.ColumnWidth = 50
Cells.EntireColumn.AutoFit
Cells.RowHeight = 12.75

Range("A1:AC5000").Sort Key1:=Range("A1"), Order1:=xlAscending, Header:=xlYes

' delete program id column
Columns("B:B").Select
Selection.Delete Shift:=xlToLeft

' format Total Estimated Program Costs column as currency
Range("AB:AB").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.NumberFormat = "$#,##0.00"

' add sheets
Cells.Select
Selection.Copy

For i = 1 To UBound(ProjArray)

    Worksheets.Add(After:=Worksheets(1)).Name = ProjArray(i)
    Sheets(ProjArray(i)).Select
    ActiveSheet.Paste

Next i


' delete rows.  highlight upcoming rows.
For i = 1 To UBound(ProjArray)

    Sheets(ProjArray(i)).Select
    
    For Each c In Worksheets(ProjArray(i)).Range("A2:A5000").Cells
        Select Case True
        Case ProjArray(i) Like "*_NE"
            If c.Value <> Left(ProjArray(i), 8) Or c.Offset(0, 4).Value <> "Nurse Education Program" Then c.EntireRow.ClearContents
        Case ProjArray(i) = "GRIALP15"
            If c.Value <> ProjArray(i) Or c.Offset(0, 4).Value = "Nurse Education Program" Then c.EntireRow.ClearContents
        Case ProjArray(i) = "GRIGAM15"
            If c.Value <> ProjArray(i) Or c.Offset(0, 4).Value = "Nurse Education Program" Then c.EntireRow.ClearContents
        Case Else
            If c.Value <> ProjArray(i) Then c.EntireRow.ClearContents
        End Select
        
        If c.Offset(0, 5).Value = "Confirmed" Or c.Offset(0, 5).Value = "Pending" Or c.Offset(0, 5).Value = "Work-in-Progress" Then
            Range(c, c.Offset(0, 27)).Interior.Color = RGB(220, 230, 241)
        End If
    Next
   
Next i

' sort and lock headers
For i = 1 To UBound(ProjArray)

Worksheets(ProjArray(i)).Select

Range("A1:AC5000").Sort Key1:=Range("A1"), Order1:=xlAscending, Header:=xlYes

    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True

' draw borders
Range("A1").Select
Range(Selection, Selection.End(xlDown)).Select
Range(Selection, Selection.End(xlToRight)).Select

Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Selection.Borders(xlInsideVertical).LineStyle = xlContinuous
Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous
    
Next i

'validation
RowCnt = Application.CountA(Sheets("Program_Status_Report").Range("AB:AB"))
CombRowCnt = Application.CountA(Sheets("GRIABU15").Range("AB:AB")) + Application.CountA(Sheets("GRIALP15").Range("AB:AB")) + Application.CountA(Sheets("GRIGAM15").Range("AB:AB")) + Application.CountA(Sheets("GRIGPH15").Range("AB:AB")) + Application.CountA(Sheets("GRIPRO15").Range("AB:AB")) + Application.CountA(Sheets("GRITHR15").Range("AB:AB")) + Application.CountA(Sheets("GRIALP15_NE").Range("AB:AB")) + Application.CountA(Sheets("GRIGAM15_NE").Range("AB:AB")) + Application.CountA(Sheets("GRIGNE15").Range("AB:AB"))

TotalCost = Application.Sum(Sheets("Program_Status_Report").Range("AB:AB"))
CombTotalCost = Application.Sum(Sheets("GRIABU15").Range("AB:AB")) + Application.Sum(Sheets("GRIALP15").Range("AB:AB")) + Application.Sum(Sheets("GRIGAM15").Range("AB:AB")) + Application.Sum(Sheets("GRIGPH15").Range("AB:AB")) + Application.Sum(Sheets("GRIPRO15").Range("AB:AB")) + Application.Sum(Sheets("GRITHR15").Range("AB:AB")) + Application.Sum(Sheets("GRIALP15_NE").Range("AB:AB")) + Application.Sum(Sheets("GRIGAM15_NE").Range("AB:AB")) + Application.Sum(Sheets("GRIGNE15").Range("AB:AB"))

Beep
If (MsgBox("Original Row Count:          " & Format((RowCnt - 1), "Standard") & vbCrLf & "Combined Row Count:     " & Format((CombRowCnt - 9), "Standard") & vbCrLf & vbCrLf & "Original Total:        " & Format(TotalCost, "Currency") & vbCrLf & "Combined Total:   " & Format(CombTotalCost, "Currency"), vbOKCancel, "Validation..")) = vbCancel Then
    Exit Sub
End If


' split tabs out to new workbooks
For i = 1 To 7

    Select Case True
        Case ProjArray(i) = "GRIABU15"
            ProjFileName = "Albumin Speakers Bureau"
        Case ProjArray(i) = "GRIALP15"
            ProjFileName = "Alphanate Speakers Bureau"
        Case ProjArray(i) = "GRIGAM15"
            ProjFileName = "Gamunex-C Speakers Bureau"
        Case ProjArray(i) = "GRIGPH15"
            ProjFileName = "Pharmacy Speakers Bureau"
        Case ProjArray(i) = "GRIPRO15"
            ProjFileName = "Prolastin"
        Case ProjArray(i) = "GRITHR15"
            ProjFileName = "Thrombate III Speakers Bureau"
        Case ProjArray(i) = "GRIGNE15"
            ProjFileName = "Grifols Neuropathy"
        Case Else
            ProjFileName = "Undefined"
    End Select

    Workbooks.Add
    ActiveWorkbook.SaveAs Filename:=DropLoc & Format(Now(), "mm-dd-yyyy") & " Status Report - " & ProjFileName & ".xlsx", FileFormat:=51

    
    Windows(ExportFileName).Activate
    Sheets(ProjArray(i)).Copy Before:=Workbooks(Format(Now(), "mm-dd-yyyy") & " Status Report - " & ProjFileName & ".xlsx").Sheets(1)
    
    Workbooks(Format(Now(), "mm-dd-yyyy") & " Status Report - " & ProjFileName & ".xlsx").Activate
    Worksheets("Sheet1").Name = "QuickStats"
    Worksheets("Sheet2").Delete
    Worksheets("Sheet3").Delete
    
    Worksheets("QuickStats").Select
    Range("A1").Value = "Completed:"
    Range("A2").Value = "Upcoming:"
    Range("A3").Value = "Pending RSD Approval:"
    Range("A4").Value = "Cancelled:"
    Range("A5").Value = "Rescheduled:"
    Range("A6").Value = "Total:"
    
    Range("B1").Formula = "=COUNTIF(" & ProjArray(i) & "!F:F,""Completed"")"
    Range("B2").Formula = "=COUNTIF(" & ProjArray(i) & "!F:F,""Confirmed"") + COUNTIF(" & ProjArray(i) & "!F:F,""Work-in-Progress"") "
    Range("B3").Formula = "=COUNTIF(" & ProjArray(i) & "!F:F,""Pending"")"
    Range("B4").Formula = "=COUNTIF(" & ProjArray(i) & "!F:F,""Cancelled"")"
    Range("B5").Formula = "=COUNTIF(" & ProjArray(i) & "!F:F,""Rescheduled"")"
    Range("B6").Formula = "=SUM(B1:B5)"

    Range("A:A").ColumnWidth = 22
    
    'sort sheets
    For x = 1 To Sheets.Count
        For y = 1 To Sheets.Count - 1
            If UCase$(Sheets(y).Name) > UCase$(Sheets(y + 1).Name) Then
                Sheets(y).Move After:=Sheets(y + 1)
            End If
        Next y
    Next x
    
    Sheets(1).Select
    Range("A1").Select
    ActiveWorkbook.Save
    
Next i

' append nurse educator tabs
For i = 8 To 9
    
    Select Case True
        Case ProjArray(i) = "GRIALP15_NE"
            ProjFileName = "Alphanate Speakers Bureau"
        Case ProjArray(i) = "GRIGAM15_NE"
            ProjFileName = "Gamunex-C Speakers Bureau"
    End Select

    
    Windows(ExportFileName).Activate
    Sheets(ProjArray(i)).Copy Before:=Workbooks(Format(Now(), "mm-dd-yyyy") & " Status Report - " & ProjFileName & ".xlsx").Sheets(1)
        Range("C:C,G:G,H:H,I:I,J:J,S:S,T:T,U:U,V:V").Delete Shift:=xlToLeft
    
    Workbooks(Format(Now(), "mm-dd-yyyy") & " Status Report - " & ProjFileName & ".xlsx").Activate
    Sheets.Add.Name = "NurseEducator QuickStats"
    Worksheets("NurseEducator QuickStats").Select
    Range("A1").Value = "Completed:"
    Range("A2").Value = "Upcoming:"
    Range("A3").Value = "Pending NE Approval:"
    Range("A4").Value = "Cancelled:"
    Range("A5").Value = "Rescheduled:"
    Range("A6").Value = "Total:"
    Range("A7").Value = "Program Cost does not equal 50:"
    
    Range("B1").Formula = "=COUNTIF(" & ProjArray(i) & "!E:E,""Completed"")"
    Range("B2").Formula = "=COUNTIF(" & ProjArray(i) & "!E:E,""Confirmed"") + COUNTIF(" & ProjArray(i) & "!E:E,""Work-in-Progress"") "
    Range("B3").Formula = "=COUNTIF(" & ProjArray(i) & "!E:E,""Pending"")"
    Range("B4").Formula = "=COUNTIF(" & ProjArray(i) & "!E:E,""Cancelled"")"
    Range("B5").Formula = "=COUNTIF(" & ProjArray(i) & "!E:E,""Rescheduled"")"
    Range("B6").Formula = "=SUM(B1:B5)"
    Range("B7").Formula = "=B6-COUNTIF(" & ProjArray(i) & "!W:W,50)"
        If Range("B7").Value <> 0 Then
            Range("B7").Font.Color = vbRed
            Range("B7").Font.Bold = True
        End If

    
    Range("A:A").ColumnWidth = 30
    
    Worksheets(ProjArray(i)).Name = "Nurse Educator Programs"
    
    'sort sheets
    For x = 1 To Sheets.Count
        For y = 1 To Sheets.Count - 1
            If UCase$(Sheets(y).Name) > UCase$(Sheets(y + 1).Name) Then
                Sheets(y).Move After:=Sheets(y + 1)
            End If
        Next y
    Next x
    
    Sheets(1).Select
    Range("A1").Select
    ActiveWorkbook.Save
    
Next i

Application.ScreenUpdating = True
Application.DisplayAlerts = True

End Sub