Attribute VB_Name = "Module3"
Sub NewWeekHours()
'
' New Week Macro - Hours chart
' Sets the start date of the pivot table to the current monday
' groups by 7 days and per week
' filters starting with the current week and 4 future weeks
' excludes dates with < or > signs
'

Dim p As PivotTable
Dim pf As pivotfield
Dim pit As pivotitem
Dim pr As Range
Dim dt As Date
Dim dtMonday As Date
Dim i As Integer
dt = Date
dtMonday = dt - Weekday(dt, vbMonday) + 1
'Count per week pivot chart
Set p = Sheet10.PivotTables("PivotTable2")
Set pr = p.PivotFields("Date Expected").DataRange
p.RowAxisLayout xlTabularRow
Set pf = p.PivotFields("Date Expected")
pf.LabelRange.Group Start:=dtMonday, End:=True, By:=7, Periods:=Array(False, False, False, True, False, False, False)
'pr.Cells().Ungroup
With pf
    For i = 1 To .PivotItems.Count - 1
        If i <= 6 Then
            .PivotItems(.PivotItems(i).Name).Visible = True
        Else
            .PivotItems(.PivotItems(i).Name).Visible = False
        End If
        
        If InStr(1, .PivotItems(i), "<") Then
            .PivotItems(i).Visible = False
        End If
    Next i
End With
pf.PivotItems(i).Visible = False
End Sub

