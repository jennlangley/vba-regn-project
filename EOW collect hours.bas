Attribute VB_Name = "Module4"
Sub EOWCollectHours()
'
' Sums the weekly allocation sheet table
' Stores values in the Hours table
' for use once weekly end of day friday!!
'
Dim s As Worksheet
Dim tmsz As Double 'stores number of data reviewers
Dim esst As Variant 'stores number of hours on essential meetings and tasks (from input box)
Dim tbl As ListObject 'hours collection table
Dim nwrw As ListRow 'new row added to the hours collection table
Dim dt As Date 'current date
Dim dtMonday As Date 'date of monday from the current/collection date

dt = Date 'stores current date
dtMonday = dt - Weekday(dt, vbMonday) + 1 'calculates date of monday of current week

'calculate team size from workbook
For Each s In ActiveWorkbook.Sheets 'loops through sheets in workbook
    If Len(s.Name) = 2 Then 'tallys up the number of sheets which are 2 in length, aka team members initials
        tmsz = tmsz + 1 'adds one to the count of members if the length of the sheet name is 2
    End If
Next s

tmsz = tmsz - (0.5) * 2 'subtracts .5 from team size for 2 team leadS
esst = InputBox("Enter hours spent on Essential Meetings & tasks: ") 'stores num. essential meetings plus core work

Set tbl = ActiveWorkbook.Worksheets("Hours Table").ListObjects("HoursCollection") 'sets tbl equal to the hours table
Set nwrw = tbl.ListRows.Add(1) 'adds row to top of hours table

'sums and adds total hours from the team on the first sheet "Weekly Allocation" and stores in columns 1-11 of new row in hours table
With nwrw
    .Range(1) = dtMonday
    .Range(2) = Application.WorksheetFunction.Sum(ActiveWorkbook.Worksheets("Weekly Allocation").Range("E3:M3")) 'email&training
    .Range(3) = Application.WorksheetFunction.Sum(ActiveWorkbook.Worksheets("Weekly Allocation").Range("E4:M4")) 'meetings
    .Range(4) = Application.WorksheetFunction.Sum(ActiveWorkbook.Worksheets("Weekly Allocation").Range("E5:M5")) 'ci&automation
    .Range(5) = Application.WorksheetFunction.Sum(ActiveWorkbook.Worksheets("Weekly Allocation").Range("E6:M6")) 'admin
    .Range(6) = Application.WorksheetFunction.Sum(ActiveWorkbook.Worksheets("Weekly Allocation").Range("E7:M7")) 'pto
    .Range(7) = Application.WorksheetFunction.Sum(ActiveWorkbook.Worksheets("Weekly Allocation").Range("E8:M8")) 'investigation
    .Range(8) = Application.WorksheetFunction.Sum(ActiveWorkbook.Worksheets("Weekly Allocation").Range("E9:M9")) 'ad hoc
    .Range(9) = Application.WorksheetFunction.Sum(ActiveWorkbook.Worksheets("Weekly Allocation").Range("E10:M10")) 'data acquisition
    .Range(10) = Application.WorksheetFunction.Sum(ActiveWorkbook.Worksheets("Weekly Allocation").Range("E11:M11")) 'trend reports
    .Range(11) = Application.WorksheetFunction.Sum(ActiveWorkbook.Worksheets("Weekly Allocation").Range("E12:M12")) 'ppm plans
    .Range(12) = tmsz 'stores number of team members in column 12
    .Range(13) = esst 'stores number of essential meetings plus core work in column 13
End With
End Sub


