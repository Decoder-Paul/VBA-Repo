Attribute VB_Name = "Counting"
Public Rng As Range
Sub ticketCount()

'========================================================================================================
' TicketCount
' -------------------------------------------------------------------------------------------------------
' Purpose   :   To get no. of Tickets depending on conditions from Data File to Dashboard.
'
' Author    :   Subhankar Paul, 9th February, 2017
' Notes     :   Different Ticket Types: 'INC', 'SRQ', 'ACT', 'PRB' are string constant
'
' Parameter :   N/A
' Returns   :   N/A
' ---------------------------------------------------------------
' Revision History
'========================================================================================================
Application.ScreenUpdating = False
Application.DisplayAlerts = False


Dim sheetData As String
Dim sheetDbd As String
Dim BI As Variant
Dim R, lro As Long
Dim C As Long

sheetData = "Consolidated Report"
sheetDbd = "Summary"

'------------ Checking for the Data & Dashboard Sheets -----------
If fSheetExists(sheetData) = True Then
    Sheets(sheetData).Activate
    If fSheetExists(sheetDbd) = True Then
        Sheets(sheetDbd).Activate
    Else
        MsgBox "Dashboard Sheet doesn't Exist"
    End If
Else
    MsgBox "Data Sheet doesn't Exist"
End If

Sheets(sheetDbd).Select

'------------ Cleaning Previous Data from the cells -----------
Dim clean As Range

Set clean = Range("B4:K12")
clean.Select
Selection.Cells.ClearContents

Set clean = Range("B14:K22")
clean.Select
Selection.Cells.ClearContents

Set clean = Range("B24:K32")
clean.Select
Selection.Cells.ClearContents

Set clean = Range("B34:K42")
clean.Select
Selection.Cells.ClearContents

Set clean = Range("B44:K51")
clean.Select
Selection.Cells.ClearContents
'------------------ Range Selection ----------
Sheets(sheetData).Select
lro = Sheets(sheetData).Cells(Rows.Count, "A").End(xlUp).Row
Set Rng = Sheets(sheetData).Range("O2:O" & lro)
For i = 2 To lro
    Cells(i, 18).Value = CLng(Cells(i, 10).Value) 'Creation Date Converted to Integer
    If Cells(i, 12).Value <> "" Then
        Cells(i, 19).Value = CLng(Cells(i, 12).Value) 'Finish Date Converted to Integer
    End If
Next i
Call Trans_SRQ_Ver1
Call Trans_INC_Ver1
Call Trans_PRB_Ver1
Call Trans_ACT_Ver1
Call Atlas_SRQ_Ver1
Call Atlas_INC_Ver1
Call Atlas_PRB_Ver1
Call Atlas_ACT_Ver1


End Sub

Sub Trans_SRQ_Ver1()
    
    Sheets("Consolidated Report").Range("A1:S1").Select
    Selection.AutoFilter
    '------------------ Filtering Data for Version 1 ----------------------
    With Selection
    ' For All Transformer/
    .AutoFilter Field:=9, Criteria1:="Transformers"
    ' For All SRQ Tickets
    .AutoFilter Field:=2, Criteria1:="SRQ"
        ' For Resolved Ticket
        ' end_Date >= 'Finish Date' >= start_Date
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver1_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver1_enDt)
        
        ' SRQ-P1_Resolved
        .AutoFilter Field:=13, Criteria1:="1"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(4, 2).Value = Count
        Sheets("Summary").Cells(4, 14).Value = WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        
        ' SRQ-P2_Resolved
        .AutoFilter Field:=13, Criteria1:="2"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(5, 2).Value = Count
        Sheets("Summary").Cells(5, 14).Value = WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        
        ' SRQ-P3_Resolved
        .AutoFilter Field:=13, Criteria1:="3"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(6, 2).Value = Count
        Sheets("Summary").Cells(6, 14).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))

        ' SRQ-P4 & P5 Resolved
        .AutoFilter Field:=13, Criteria1:="4", Operator:=xlOr, Criteria2:="5"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(7, 2).Value = Count
        Sheets("Summary").Cells(7, 14).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' SRQ-Total_Resolved
        .AutoFilter Field:=13
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(10, 2).Value = Count
        Sheets("Summary").Cells(8, 14).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        .AutoFilter Field:=19
        
        ' SRQ-Total_Opening Balance
        .AutoFilter Field:=18, Criteria1:="<" & CLng(ver1_stDt)
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver1_stDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(8, 2).Value = Count
        
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
        ' SRQ-Total_Received
        .AutoFilter Field:=18, Criteria1:=">=" & CLng(ver1_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver1_enDt)
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(9, 2).Value = Count
            
        .AutoFilter Field:=18
        
        ' SRQ-Total_Carry Forward
        .AutoFilter Field:=18, Criteria1:="<=" & CLng(ver1_enDt)
        .AutoFilter Field:=19, Criteria1:=">" & CLng(ver1_enDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(11, 2).Value = Count
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
    End With
    
End Sub
Sub Trans_INC_Ver1()
    
    Sheets("Consolidated Report").Range("A1:S1").Select
    Selection.AutoFilter
    '------------------ Filtering Data for Version 1 ----------------------
    With Selection
    ' For All Transformer
    .AutoFilter Field:=9, Criteria1:="Transformers"
    ' For All INC Tickets
    .AutoFilter Field:=2, Criteria1:="INC"
        ' For Resolved Ticket
        ' end_Date >= 'Finish Date' >= start_Date
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver1_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver1_enDt)
        
        ' INC-P1_Resolved
        .AutoFilter Field:=13, Criteria1:="1"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(14, 2).Value = Count
        Sheets("Summary").Cells(10, 14).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' INC-P2_Resolved
        .AutoFilter Field:=13, Criteria1:="2"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(15, 2).Value = Count
        Sheets("Summary").Cells(11, 14).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' INC-P3_Resolved
        .AutoFilter Field:=13, Criteria1:="3"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(16, 2).Value = Count
        Sheets("Summary").Cells(12, 14).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' INC-P4 & P5 Resolved
        .AutoFilter Field:=13, Criteria1:="4", Operator:=xlOr, Criteria2:="5"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(17, 2).Value = Count
        Sheets("Summary").Cells(13, 14).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' INC-Total_Resolved
        .AutoFilter Field:=13
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(20, 2).Value = Count
        Sheets("Summary").Cells(14, 14).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        .AutoFilter Field:=19
        
        ' INC-Total_Opening Balance
        .AutoFilter Field:=18, Criteria1:="<" & CLng(ver1_stDt)
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver1_stDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(18, 2).Value = Count
        
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
        ' INC-Total_Received
        .AutoFilter Field:=18, Criteria1:=">=" & CLng(ver1_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver1_enDt)
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(19, 2).Value = Count
            
        .AutoFilter Field:=18
        
        ' INC-Total_Carry Forward
        .AutoFilter Field:=18, Criteria1:="<=" & CLng(ver1_enDt)
        .AutoFilter Field:=19, Criteria1:=">" & CLng(ver1_enDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(21, 2).Value = Count
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
    End With
    
End Sub
Sub Trans_PRB_Ver1()
    
    Sheets("Consolidated Report").Range("A1:S1").Select
    Selection.AutoFilter
    '------------------ Filtering Data for Version 1 ----------------------
    With Selection
    ' For All Transformer
    .AutoFilter Field:=9, Criteria1:="Transformers"
    ' For All INC Tickets
    .AutoFilter Field:=2, Criteria1:="PRB"
        ' For Resolved Ticket
        ' end_Date >= 'Finish Date' >= start_Date
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver1_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver1_enDt)
        
        ' PRB-P1_Resolved
        .AutoFilter Field:=13, Criteria1:="1"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(24, 2).Value = Count
        Sheets("Summary").Cells(16, 14).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' PRB-P2_Resolved
        .AutoFilter Field:=13, Criteria1:="2"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(25, 2).Value = Count
        Sheets("Summary").Cells(17, 14).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' PRB-P3_Resolved
        .AutoFilter Field:=13, Criteria1:="3"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(26, 2).Value = Count
        Sheets("Summary").Cells(18, 14).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' PRB-P4 & P5 Resolved
        .AutoFilter Field:=13, Criteria1:="4", Operator:=xlOr, Criteria2:="5"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(27, 2).Value = Count
        Sheets("Summary").Cells(19, 14).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' PRB-Total_Resolved
        .AutoFilter Field:=13
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(30, 2).Value = Count
        Sheets("Summary").Cells(20, 14).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        .AutoFilter Field:=19
        
        ' PRB-Total_Opening Balance
        .AutoFilter Field:=18, Criteria1:="<" & CLng(ver1_stDt)
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver1_stDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(28, 2).Value = Count
        
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
        ' PRB-Total_Received
        .AutoFilter Field:=18, Criteria1:=">=" & CLng(ver1_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver1_enDt)
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(29, 2).Value = Count
            
        .AutoFilter Field:=18
        
        ' PRB-Total_Carry Forward
        .AutoFilter Field:=18, Criteria1:="<=" & CLng(ver1_enDt)
        .AutoFilter Field:=19, Criteria1:=">" & CLng(ver1_enDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(31, 2).Value = Count
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
    End With
    
End Sub

Sub Trans_ACT_Ver1()
    
    Sheets("Consolidated Report").Range("A1:S1").Select
    Selection.AutoFilter
    '------------------ Filtering Data for Version 1 ----------------------
    With Selection
    ' For All Transformer
    .AutoFilter Field:=9, Criteria1:="Transformers"
    ' For All INC Tickets
    .AutoFilter Field:=2, Criteria1:="ACT"
        ' For Resolved Ticket
        ' end_Date >= 'Finish Date' >= start_Date
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver1_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver1_enDt)
        
        ' ACT-P1_Resolved
        .AutoFilter Field:=13, Criteria1:="1"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(34, 2).Value = Count
        Sheets("Summary").Cells(22, 14).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' ACT-P2_Resolved
        .AutoFilter Field:=13, Criteria1:="2"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(35, 2).Value = Count
        Sheets("Summary").Cells(23, 14).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' ACT-P3_Resolved
        .AutoFilter Field:=13, Criteria1:="3"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(36, 2).Value = Count
        Sheets("Summary").Cells(24, 14).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' ACT-P4 & P5 Resolved
        .AutoFilter Field:=13, Criteria1:="4", Operator:=xlOr, Criteria2:="5"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(37, 2).Value = Count
        Sheets("Summary").Cells(25, 14).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' ACT-Total_Resolved
        .AutoFilter Field:=13
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(40, 2).Value = Count
        Sheets("Summary").Cells(26, 14).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        .AutoFilter Field:=19
        
        ' ACT-Total_Opening Balance
        .AutoFilter Field:=18, Criteria1:="<" & CLng(ver1_stDt)
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver1_stDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(38, 2).Value = Count
        
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
        ' ACT-Total_Received
        .AutoFilter Field:=18, Criteria1:=">=" & CLng(ver1_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver1_enDt)
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(39, 2).Value = Count
            
        .AutoFilter Field:=18
        
        ' ACT-Total_Carry Forward
        .AutoFilter Field:=18, Criteria1:="<=" & CLng(ver1_enDt)
        .AutoFilter Field:=19, Criteria1:=">" & CLng(ver1_enDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(41, 2).Value = Count
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
    End With
    
End Sub

Sub Atlas_SRQ_Ver1()
    
    Sheets("Consolidated Report").Range("A1:S1").Select
    Selection.AutoFilter
    '------------------ Filtering Data for Version 1 ----------------------
    With Selection
    ' For All Atlas
    .AutoFilter Field:=9, Criteria1:="Atlas"
    ' For All SRQ Tickets
    .AutoFilter Field:=2, Criteria1:="SRQ"
        ' For Resolved Ticket
        ' end_Date >= 'Finish Date' >= start_Date
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver1_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver1_enDt)
        
        ' SRQ-P1_Resolved
        .AutoFilter Field:=13, Criteria1:="1"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(4, 7).Value = Count
        Sheets("Summary").Cells(4, 19).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        
        ' SRQ-P2_Resolved
        .AutoFilter Field:=13, Criteria1:="2"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(5, 7).Value = Count
        Sheets("Summary").Cells(5, 19).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        
        ' SRQ-P3_Resolved
        .AutoFilter Field:=13, Criteria1:="3"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(6, 7).Value = Count
        Sheets("Summary").Cells(6, 19).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))

        ' SRQ-P4 & P5 Resolved
        .AutoFilter Field:=13, Criteria1:="4", Operator:=xlOr, Criteria2:="5"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(7, 7).Value = Count
        Sheets("Summary").Cells(7, 19).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' SRQ-Total_Resolved
        .AutoFilter Field:=13
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(10, 7).Value = Count
        Sheets("Summary").Cells(8, 19).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        .AutoFilter Field:=19
        
        ' SRQ-Total_Opening Balance
        .AutoFilter Field:=18, Criteria1:="<" & CLng(ver1_stDt)
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver1_stDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(8, 7).Value = Count
        
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
        ' SRQ-Total_Received
        .AutoFilter Field:=18, Criteria1:=">=" & CLng(ver1_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver1_enDt)
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(9, 7).Value = Count
            
        .AutoFilter Field:=18
        
        ' SRQ-Total_Carry Forward
        .AutoFilter Field:=18, Criteria1:="<=" & CLng(ver1_enDt)
        .AutoFilter Field:=19, Criteria1:=">" & CLng(ver1_enDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(11, 7).Value = Count
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
    End With
    
End Sub
Sub Atlas_INC_Ver1()
    
    Sheets("Consolidated Report").Range("A1:S1").Select
    Selection.AutoFilter
    '------------------ Filtering Data for Version 1 ----------------------
    With Selection
    ' For All Atlas
    .AutoFilter Field:=9, Criteria1:="Atlas"
    ' For All INC Tickets
    .AutoFilter Field:=2, Criteria1:="INC"
        ' For Resolved Ticket
        ' end_Date >= 'Finish Date' >= start_Date
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver1_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver1_enDt)
        
        ' INC-P1_Resolved
        .AutoFilter Field:=13, Criteria1:="1"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(14, 7).Value = Count
        Sheets("Summary").Cells(10, 19).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' INC-P2_Resolved
        .AutoFilter Field:=13, Criteria1:="2"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(15, 7).Value = Count
        Sheets("Summary").Cells(11, 19).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' INC-P3_Resolved
        .AutoFilter Field:=13, Criteria1:="3"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(16, 7).Value = Count
        Sheets("Summary").Cells(12, 19).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' INC-P4 & P5 Resolved
        .AutoFilter Field:=13, Criteria1:="4", Operator:=xlOr, Criteria2:="5"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(17, 7).Value = Count
        Sheets("Summary").Cells(13, 19).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' INC-Total_Resolved
        .AutoFilter Field:=13
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(20, 7).Value = Count
        Sheets("Summary").Cells(14, 19).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        .AutoFilter Field:=19
        
        ' INC-Total_Opening Balance
        .AutoFilter Field:=18, Criteria1:="<" & CLng(ver1_stDt)
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver1_stDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(18, 7).Value = Count
        
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
        ' INC-Total_Received
        .AutoFilter Field:=18, Criteria1:=">=" & CLng(ver1_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver1_enDt)
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(19, 7).Value = Count
            
        .AutoFilter Field:=18
        
        ' INC-Total_Carry Forward
        .AutoFilter Field:=18, Criteria1:="<=" & CLng(ver1_enDt)
        .AutoFilter Field:=19, Criteria1:=">" & CLng(ver1_enDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(21, 7).Value = Count
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
    End With
    
End Sub
Sub Atlas_PRB_Ver1()
    
    Sheets("Consolidated Report").Range("A1:S1").Select
    Selection.AutoFilter
    '------------------ Filtering Data for Version 1 ----------------------
    With Selection
    ' For All Atlas
    .AutoFilter Field:=9, Criteria1:="Atlas"
    ' For All INC Tickets
    .AutoFilter Field:=2, Criteria1:="PRB"
        ' For Resolved Ticket
        ' end_Date >= 'Finish Date' >= start_Date
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver1_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver1_enDt)
        
        ' PRB-P1_Resolved
        .AutoFilter Field:=13, Criteria1:="1"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(24, 7).Value = Count
        Sheets("Summary").Cells(16, 19).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' PRB-P2_Resolved
        .AutoFilter Field:=13, Criteria1:="2"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(25, 7).Value = Count
        Sheets("Summary").Cells(17, 19).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' PRB-P3_Resolved
        .AutoFilter Field:=13, Criteria1:="3"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(26, 7).Value = Count
        Sheets("Summary").Cells(18, 19).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' PRB-P4 & P5 Resolved
        .AutoFilter Field:=13, Criteria1:="4", Operator:=xlOr, Criteria2:="5"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(27, 7).Value = Count
        Sheets("Summary").Cells(19, 19).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' PRB-Total_Resolved
        .AutoFilter Field:=13
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(30, 7).Value = Count
        Sheets("Summary").Cells(20, 19).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        .AutoFilter Field:=19
        
        ' PRB-Total_Opening Balance
        .AutoFilter Field:=18, Criteria1:="<" & CLng(ver1_stDt)
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver1_stDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(28, 7).Value = Count
        
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
        ' PRB-Total_Received
        .AutoFilter Field:=18, Criteria1:=">=" & CLng(ver1_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver1_enDt)
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(29, 7).Value = Count
            
        .AutoFilter Field:=18
        
        ' PRB-Total_Carry Forward
        .AutoFilter Field:=18, Criteria1:="<=" & CLng(ver1_enDt)
        .AutoFilter Field:=19, Criteria1:=">" & CLng(ver1_enDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(31, 7).Value = Count
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
    End With
    
End Sub

Sub Atlas_ACT_Ver1()
    
    Sheets("Consolidated Report").Range("A1:S1").Select
    Selection.AutoFilter
    '------------------ Filtering Data for Version 1 ----------------------
    With Selection
    ' For All Atlas
    .AutoFilter Field:=9, Criteria1:="Atlas"
    ' For All INC Tickets
    .AutoFilter Field:=2, Criteria1:="ACT"
        ' For Resolved Ticket
        ' end_Date >= 'Finish Date' >= start_Date
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver1_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver1_enDt)
        
        ' ACT-P1_Resolved
        .AutoFilter Field:=13, Criteria1:="1"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(34, 7).Value = Count
        Sheets("Summary").Cells(22, 19).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' ACT-P2_Resolved
        .AutoFilter Field:=13, Criteria1:="2"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(35, 7).Value = Count
        Sheets("Summary").Cells(23, 19).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' ACT-P3_Resolved
        .AutoFilter Field:=13, Criteria1:="3"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(36, 7).Value = Count
        Sheets("Summary").Cells(24, 19).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' ACT-P4 & P5 Resolved
        .AutoFilter Field:=13, Criteria1:="4", Operator:=xlOr, Criteria2:="5"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(37, 7).Value = Count
        Sheets("Summary").Cells(25, 19).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' ACT-Total_Resolved
        .AutoFilter Field:=13
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(40, 7).Value = Count
        Sheets("Summary").Cells(26, 19).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        .AutoFilter Field:=19
        
        ' ACT-Total_Opening Balance
        .AutoFilter Field:=18, Criteria1:="<" & CLng(ver1_stDt)
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver1_stDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(38, 7).Value = Count
        
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
        ' ACT-Total_Received
        .AutoFilter Field:=18, Criteria1:=">=" & CLng(ver1_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver1_enDt)
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(39, 7).Value = Count
            
        .AutoFilter Field:=18
        
        ' ACT-Total_Carry Forward
        .AutoFilter Field:=18, Criteria1:="<=" & CLng(ver1_enDt)
        .AutoFilter Field:=19, Criteria1:=">" & CLng(ver1_enDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(41, 7).Value = Count
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
    End With
    
End Sub
