Attribute VB_Name = "Transpose"
Sub prog()
Dim i, j, data_i, C, utility As Integer
Dim l As Long
Dim sheetname As String
Dim idd As Long
Dim emp, projId, proj, billStat, isActive As String
Dim sDate, eDate As Date
Dim empData(6) As Variant

sheetname = "Data"
Sheets("TR Information").Activate
'Sheets(sheetname).Activate
Length = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row

ActiveWorkbook.Worksheets("TR Information").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("TR Information").Sort.SortFields.Add Key:=Range( _
        "A2:A4121"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets("TR Information").Sort.SortFields.Add Key:=Range( _
        "G2:G4121"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("TR Information").Sort
        .SetRange Range("A1:L4121")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

j = 2

For i = 2 To Length - 1
    data_i = i
    idd = Sheets("TR Information").Cells(i, 1).Value
    emp = Sheets("TR Information").Cells(i, 2).Value
    empData(0) = Sheets("TR Information").Cells(i, 3).Value 'proj name
    empData(1) = Sheets("TR Information").Cells(i, 4).Value 'proj code
    empData(2) = Sheets("TR Information").Cells(i, 6).Value 'Utilization
    empData(3) = Sheets("TR Information").Cells(i, 7).Value 'start Date
    empData(4) = Sheets("TR Information").Cells(i, 8).Value 'End Date
    empData(5) = Sheets("TR Information").Cells(i, 10).Value 'Billing status
    empData(6) = Sheets("TR Information").Cells(i, 12).Value 'IsActive
    'printing first occurence
    Sheets(sheetname).Cells(j, 1).Value = idd
    Sheets(sheetname).Cells(j, 2).Value = emp
    C = 3
    Sheets(sheetname).Select
    ActiveWorkbook.Sheets(sheetname).Range(Cells(j, C), Cells(j, C + 6)).Value = empData
    While idd = Sheets("TR Information").Cells(data_i + 1, 1).Value
        C = C + 7
        data_i = data_i + 1
        empData(0) = Sheets("TR Information").Cells(data_i, 3).Value 'proj name
        empData(1) = Sheets("TR Information").Cells(data_i, 4).Value 'proj code
        empData(2) = Sheets("TR Information").Cells(data_i, 6).Value 'Utilization
        empData(3) = Sheets("TR Information").Cells(data_i, 7).Value 'start Date
        empData(4) = Sheets("TR Information").Cells(data_i, 8).Value 'End Date
        empData(5) = Sheets("TR Information").Cells(data_i, 10).Value 'Billing status
        empData(6) = Sheets("TR Information").Cells(data_i, 12).Value 'IsActive
        Range(Cells(j, C), Cells(j, C + 6)).Value = empData
    Wend
    i = data_i + 1
    j = j + 1
    i = i - 1
Next
End Sub


