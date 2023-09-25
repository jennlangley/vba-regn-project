Attribute VB_Name = "Module3"
Sub CollectMetrics()
'
' Stores the metrics entered by the data reviewer and copies to table
'
Dim ttl As String 'investigation number
Dim typ As String 'investigation type
Dim tim As String 'hours spent
Dim nm As String 'reviewer initials
Dim nme As String 'reviewer fName
Dim dln As String 'deadline change
Dim wb As Workbook 'workbook
Dim dte As Date 'submit date
Dim tbl As ListObject 'metrics table
Dim nwrw As ListRow 'new row to be added to table

nm = ActiveSheet.Name 'stores name of active excel sheet (which is reviewer initials)
dte = Date 'stores current date = submit date
ttl = ActiveWorkbook.Worksheets(nm).Range("Y3").Value 'stores values where reviewer
typ = ActiveWorkbook.Worksheets(nm).Range("Y4").Value 'inputs metrics on their sheet
tim = ActiveWorkbook.Worksheets(nm).Range("Y5").Value 'in these variables
nme = ActiveWorkbook.Worksheets(nm).Range("B2").Value
dln = ActiveWorkbook.Worksheets(nm).Range("Y6").Value

Set tbl = ActiveWorkbook.Worksheets("Metrics Collection").ListObjects("Metrics") 'sets the metrics table as this variable
Set nwrw = tbl.ListRows.Add(1) 'adds a new row to the table

With nwrw 'adding values to the new table in the row
    .Range(1) = ttl 'adds the values the reviewer had entered on their sheet
    .Range(2) = typ 'which are stored in these variables
    .Range(3) = tim 'to columns 1-3 of the metrics table
    .Range(4) = nme 'automatically adds their initials in column 4
    .Range(5) = dte 'and the submit date in column 5
    .Range(6) = dln
    
End With

ActiveWorkbook.Worksheets(nm).Range("Y3:Y7").ClearContents 'clears the cells where
                                                           'the reviewer entered the values on their sheet
End Sub

