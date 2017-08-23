Attribute VB_Name = "DailyReport"
Public Rng As Range
Public Function fSheetExists(sheetToFind As String) As Boolean
    Dim sheet As Worksheet
    fSheetExists = False
    For Each sheet In Worksheets
        If sheetToFind = sheet.name Then
            fSheetExists = True
            Exit Function
        End If
    Next sheet
    MsgBox (sheetToFind & " is missing or renamed!!")
End Function

Sub UpdateTeam()
' It will update the list of present employee working currently
' on the Summary Sheet from Parameters Sheet
Application.ScreenUpdating = False
Application.DisplayAlerts = False

Dim startTime As Double
Dim totalTime As Double
startTime = Timer

    Dim WB As Workbook
    Dim WS_summ As Worksheet
    Dim WS_param As Worksheet
    Set WB = ActiveWorkbook
    
    If fSheetExists("Summary") Then
        Set WS_summ = WB.Sheets("Summary")
    End If
    If fSheetExists("Parameters") Then
        Set WS_param = WB.Sheets("Parameters")
    End If
    
    WS_summ.Select
    lroSumm = WS_summ.Cells(WS_summ.Rows.Count, "A").End(xlUp).Row
    Range("A3:M" & lroSumm).ClearContents
    'cleaning all the present data in Summary sheet
    WS_param.Select
    lroParam = WS_param.Cells(WS_param.Rows.Count, "B").End(xlUp).Row
    Range("B2:B" & lroParam).Copy
    WS_summ.Select
    WS_summ.Range("A3").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    'copy-pasting the support team names to Summary sheet from Parameter

Application.ScreenUpdating = True
Application.DisplayAlerts = True
totalTime = Round(Timer - startTime, 2)
MsgBox "Update Complete " & totalTime, vbInformation

End Sub

Sub ticketCount()
' Purpose   :   To get no. of Tickets depending on conditions from Data File to Dashboard.
'
' Author    :   Subhankar Paul, 19th August, 2017
' Notes     :   Different Ticket Types: 'INC', 'SRQ', 'ACT', 'PRB' are string constant
Dim startTime As Double
Dim totalTime As Double
startTime = Timer
Application.ScreenUpdating = False
Application.DisplayAlerts = False


Dim BI As Variant
Dim R, C As Long
Dim i As Long
Dim WB As Workbook
Dim WS_summ As Worksheet
Dim WS_data As Worksheet
Dim name As String

Set WB = ActiveWorkbook

'------------ Checking for the Data & Summary Sheets -----------
If fSheetExists("Summary") Then
    Set WS_summ = WB.Sheets("Summary")
End If
If fSheetExists("Source Data") Then
    Set WS_data = WB.Sheets("Source Data")
End If

WS_summ.Select

'------------ Cleaning Previous Data from the cells -----------
Dim lroSumm As Long
lroSumm = WS_summ.Cells(Rows.Count, "A").End(xlUp).Row

Range("B3:M" & lroSumm).ClearContents

'------------------ Range Selection ----------

For i = 3 To lroSumm
    name = WS_summ.Cells(i, 1).Value
    WS_data.Select
    If ActiveSheet.AutoFilterMode Then Cells.AutoFilter
    Call ticket(name, i)
Next i
    WS_summ.Select
    WS_summ.Range(Cells(R + 3, C + 5), Cells(R + lroSumm, C + 5)).Formula = "=Sum(B3:D3)"
    WS_summ.Range(Cells(3, 12), Cells(lroSumm, 12)).Formula = "=Sum(I3:K3)"
Application.ScreenUpdating = True
Application.DisplayAlerts = True
totalTime = Round(Timer - startTime, 2)
MsgBox "Update Complete " & totalTime, vbInformation
End Sub
Sub ticket(name As String, i As Long)
Sheets("Source Data").Select
lro = Cells(Rows.Count, "E").End(xlUp).Row

    Sheets("Source Data").Range("A1:P1").Select
    Selection.AutoFilter
    '------------------ Filtering Data for Assigned Resource ----------------------
    With Selection
        ' For a particular person
        .AutoFilter field:=8, Criteria1:=name
    
    ' QUEUED ticket Count
        .AutoFilter field:=14, Criteria1:="QUEUED"
    Count = ActiveSheet.AutoFilter.Range.Columns(5). _
            SpecialCells(xlCellTypeVisible).Count - 1
    Sheets("Summary").Cells(i, 6).Value = Count
    
    ' INPROG ticket Count
        .AutoFilter field:=14, Criteria1:="INPROG", Operator:=xlOr, Criteria2:="INPRG"
    Count = ActiveSheet.AutoFilter.Range.Columns(5). _
            SpecialCells(xlCellTypeVisible).Count - 1
    Sheets("Summary").Cells(i, 7).Value = Count
    
    ' PENDING ticket Count
        .AutoFilter field:=14, Criteria1:="PENDING"
    Count = ActiveSheet.AutoFilter.Range.Columns(5). _
            SpecialCells(xlCellTypeVisible).Count - 1
    Sheets("Summary").Cells(i, 8).Value = Count
    
    .AutoFilter field:=14
    
    ' Active SRQ Count
        .AutoFilter field:=2, Criteria1:="SRQ"
        .AutoFilter field:=11, Criteria1:=""
    Count = ActiveSheet.AutoFilter.Range.Columns(5). _
            SpecialCells(xlCellTypeVisible).Count - 1
    Sheets("Summary").Cells(i, 2).Value = Count
    
    ' SRQ Resolved count
        .AutoFilter field:=11
        .AutoFilter field:=14, Criteria1:="RESOLVED"
    Count = ActiveSheet.AutoFilter.Range.Columns(5). _
            SpecialCells(xlCellTypeVisible).Count - 1
    Sheets("Summary").Cells(i, 10).Value = Count
    
    .AutoFilter field:=14
    
    ' Active INC Count
        .AutoFilter field:=2, Criteria1:="INC"
        .AutoFilter field:=11, Criteria1:=""
    Count = ActiveSheet.AutoFilter.Range.Columns(5). _
            SpecialCells(xlCellTypeVisible).Count - 1
    Sheets("Summary").Cells(i, 3).Value = Count
    
    ' INC Resolved count
        .AutoFilter field:=11
        .AutoFilter field:=14, Criteria1:="RESOLVED"
    Count = ActiveSheet.AutoFilter.Range.Columns(5). _
            SpecialCells(xlCellTypeVisible).Count - 1
    Sheets("Summary").Cells(i, 9).Value = Count
    
    .AutoFilter field:=14
    
    ' Active ACT Count
        .AutoFilter field:=2, Criteria1:="ACT"
        .AutoFilter field:=11, Criteria1:=""
    Count = ActiveSheet.AutoFilter.Range.Columns(5). _
            SpecialCells(xlCellTypeVisible).Count - 1
    Sheets("Summary").Cells(i, 4).Value = Count
    
    ' ACT Resolved count
        .AutoFilter field:=11
        .AutoFilter field:=14, Criteria1:="COMP"
    Count = ActiveSheet.AutoFilter.Range.Columns(5). _
            SpecialCells(xlCellTypeVisible).Count - 1
    Sheets("Summary").Cells(i, 11).Value = Count
    
    .AutoFilter field:=14
    
    'Actual Efforts
        .AutoFilter field:=2
        .AutoFilter field:=14, Criteria1:="COMP", Operator:=xlOr, Criteria2:="RESOLVED"
    Count = ActiveSheet.AutoFilter.Range.Columns(5). _
            SpecialCells(xlCellTypeVisible).Count - 1
    Sheets("Summary").Cells(i, 13).Value = Application.Sum(ActiveSheet.AutoFilter.Range.Columns(13).SpecialCells(xlCellTypeVisible))
'    Sheets("Summary").Cells(i, 14).Value = Application.WorksheetFunction.Subtotal(9, Range("M3:M" & lro))
    .AutoFilter field:=14
    .AutoFilter field:=8
    End With
Cells(1, 1).Select
End Sub


