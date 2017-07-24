Attribute VB_Name = "DataExtraction"
 Public CalcState As Long
Public EventState As Boolean
Public PageBreakState As Boolean
Public Const sDQ As String = """"
Public ST As Variant
Public ED As Variant
Option Explicit

Sub pDataTab()
'========================================================================================================
' pDataTab
' -------------------------------------------------------------------------------------------------------
' Purpose of this programm : ' To Extract details from different files the Assignment data details.
'
' Author : Prasanna Kumar 3rd October, 2016
' Notes  : N/A
' Parameters :N/A
' Returns : N/A
' ---------------------------------------------------------------
' Revision History
'========================================================================================================

'On Error GoTo ErrorHandler
Dim StartTime As Double
Dim SecondsElapsed As Double

'Remember time when macro starts
  StartTime = Timer

Dim sData As String
Dim BI As Variant
Dim R As Long
Dim C As Long
Dim sbg As String

sData = "Data"

If fSheetExists(sData) = True Then
  Sheets(sData).Activate
  Sheets(sData).Select
  ST = Cells(2, 2).Value
  ED = Cells(3, 2).Value
    Cells.Select
    Selection.Cells.Clear
  Else
    Call pWorksheetcheck(sData)
End If

  Sheets(sData).Select
  ActiveWindow.DisplayGridlines = False

' As per the pre defined Template/Form at the information is formated and populated below
  If Cells(R + 1, C + 1).Value = " Calculation" Then
    GoTo MM1
  Else
    Cells(R + 1, C + 1).Select
    With Selection
      .Value = "Data calculation"
      .Font.Size = 15
      .Font.Bold = True
      .Font.Name = "Arial"
      .Font.Color = RGB(32, 32, 32)
    End With
    Range(Cells(R + 1, C + 1), Cells(R + 1, C + 83)).Select
    With Selection
      .Borders(xlEdgeBottom).Weight = xlThick
      .Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Borders(xlEdgeBottom).Color = RGB(0, 75, 175)
    End With
    End If

MM1:

  With Sheets(sData)
    .Cells(R + 6, C + 1).Value = "Month"
    .Cells(R + 6, C + 2).Value = "Quarter"
    .Cells(R + 6, C + 3).Value = "Unique"
    .Cells(R + 6, C + 4).Value = "Primary key"
    .Cells(R + 6, C + 5).Value = "Annuity"
    .Cells(R + 6, C + 6).Value = "Project/Staffing"
    .Cells(R + 6, C + 7).Value = "Mgmt type"
    .Cells(R + 6, C + 8).Value = "Location check"
    .Cells(R + 6, C + 9).Value = "Unique for HC"
    .Cells(R + 6, C + 10).Value = "Head count"

    .Cells(R + 6, C + 11).Value = "Utilisation %"
    .Cells(R + 6, C + 12).Value = "Location 1"
    .Cells(R + 6, C + 13).Value = "Emp ID"
    .Cells(R + 6, C + 14).Value = "FULL_NAME"
    .Cells(R + 6, C + 15).Value = "Location"
    .Cells(R + 6, C + 16).Value = "HRMS Location"
    .Cells(R + 6, C + 17).Value = "Emp Type"
    .Cells(R + 6, C + 18).Value = "Emp Classif"
    .Cells(R + 6, C + 19).Value = "DOJ"
    .Cells(R + 6, C + 20).Value = "DOL"

    .Cells(R + 6, C + 21).Value = "GRADE"
    .Cells(R + 6, C + 22).Value = "Status"
    .Cells(R + 6, C + 23).Value = "Classification"
    .Cells(R + 6, C + 24).Value = "Project No"
    .Cells(R + 6, C + 25).Value = "Project Name"
    .Cells(R + 6, C + 26).Value = "Client"
    .Cells(R + 6, C + 27).Value = "Project Type"
    .Cells(R + 6, C + 28).Value = "Horizontal"
    .Cells(R + 6, C + 29).Value = "Practice"
    .Cells(R + 6, C + 30).Value = "Tower"

    .Cells(R + 6, C + 31).Value = "SL"
    .Cells(R + 6, C + 32).Value = "SL1"
    .Cells(R + 6, C + 33).Value = "Sales Channel"
    .Cells(R + 6, C + 34).Value = "Client Partner"
    .Cells(R + 6, C + 35).Value = "Shadow Details"
    .Cells(R + 6, C + 36).Value = "Billable"
    .Cells(R + 6, C + 37).Value = "Non Billable"
    .Cells(R + 6, C + 38).Value = "Internal"

    .Cells(R + 6, C + 39).Value = "Leave/Company Holiday"
    .Cells(R + 6, C + 40).Value = "Unassigned"
    .Cells(R + 6, C + 41).Value = "Total"
    .Cells(R + 6, C + 42).Value = "Billed Hours"
    .Cells(R + 6, C + 43).Value = "Rate"
    .Cells(R + 6, C + 44).Value = "Revenue"
    .Cells(R + 6, C + 45).Value = "Billable Hours 1"
    .Cells(R + 6, C + 46).Value = "Non Billable Hours 1"
    .Cells(R + 6, C + 47).Value = "Assigned Hours 1"
    .Cells(R + 6, C + 48).Value = "Internal 1"

    .Cells(R + 6, C + 49).Value = "Leave/Company Holiday 1"
    .Cells(R + 6, C + 50).Value = "Unassigned 1"
    .Cells(R + 6, C + 51).Value = "Available 1"
    .Cells(R + 6, C + 52).Value = "Total"
    .Cells(R + 6, C + 53).Value = "Register"
    .Cells(R + 6, C + 54).Value = "App Salary"
    .Cells(R + 6, C + 55).Value = "Other cost"
    .Cells(R + 6, C + 56).Value = "Monthly Salary"


    .Cells(R + 6, C + 57).Value = "Perdiem in lieu of sal"
    .Cells(R + 6, C + 58).Value = "Total Salary incl Perdiem"
    .Cells(R + 6, C + 59).Value = "Billed Sal"
    .Cells(R + 6, C + 60).Value = "Unbilled Sal"
    .Cells(R + 6, C + 61).Value = "Internal Sal"
    .Cells(R + 6, C + 62).Value = "Leave Sal"
    .Cells(R + 6, C + 63).Value = "Bech Sal"
    .Cells(R + 6, C + 64).Value = "Total Sal"


    .Cells(R + 6, C + 65).Value = "Perdiem"
    .Cells(R + 6, C + 66).Value = "Travel"
    .Cells(R + 6, C + 67).Value = "Hotel & Lodging"
    .Cells(R + 6, C + 68).Value = "Others"
    .Cells(R + 6, C + 69).Value = "Expenses"
    .Cells(R + 6, C + 70).Value = "Project cost"
    .Cells(R + 6, C + 71).Value = "Contribution"
    .Cells(R + 6, C + 72).Value = "Rev for bill rate"
    .Cells(R + 6, C + 73).Value = "Hours for bill rate"
    .Cells(R + 6, C + 74).Value = "Hours for Realiation rate"
    .Cells(R + 6, C + 75).Value = "Avg Bill rate"
    .Cells(R + 6, C + 76).Value = "Avg Realisation Rate"
    .Cells(R + 6, C + 77).Value = "Utilization"
    
    .Cells(R + 6, C + 80).Value = "From Date"
    .Cells(R + 6, C + 81).Value = "To Date"
    .Cells(R + 6, C + 82).Value = "Duplicates of UniqueKey"
    '.Cells(R + 6, C + 83).Value = "Input for App Salary 1"
    .Cells(R + 6, C + 84).Value = "Input for Revenue2"
    .Cells(R + 6, C + 85).Value = "Input for Revenue1"
    .Cells(R + 6, C + 86).Value = "Input for Classification"

 Range(Cells(R + 6, C + 1), Cells(R + 6, C + 86)).Interior.Color = RGB(0, 102, 204)
 Range(Cells(R + 6, C + 1), Cells(R + 6, C + 86)).Font.Color = RGB(255, 255, 255)
 Range(Cells(R + 6, C + 1), Cells(R + 6, C + 86)).Select
 
End With

Sheets(sData).Select
Cells(2, 2).Value = ST
Cells(3, 2).Value = ED

  With Selection
    .Columns.AutoFit
    .Rows.RowHeight = 45
    .Font.Bold = True
    .VerticalAlignment = xlCenter
  End With

 For Each BI In Array(xlEdgeTop, xlEdgeLeft, xlEdgeBottom, xlEdgeRight, xlInsideHorizontal, xlInsideVertical)
    With Range(Cells(R + 6, C + 1), Cells(R + 6, C + 86)).Borders(BI)
         .Weight = xlThin
         .Color = RGB(148, 138, 84)
    End With
 Next BI


Call Dataextract

    Range("A6").Select
    If Not ActiveSheet.FilterMode Then
         Selection.AutoFilter
    End If
    
    Range("A6:CE6").Select
    Selection.Columns.AutoFit
    Selection.Columns.HorizontalAlignment = xlCenter

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

' Determine how many seconds this code will take to run
  SecondsElapsed = Round(Timer - StartTime, 2)

'Notify user in seconds
  MsgBox "This code ran successfully in " & SecondsElapsed & " seconds", vbInformation


 End Sub

Sub Dataextract()

'========================================================================================================
' Dataextract
' -------------------------------------------------------------------------------------------------------
' Purpose of this programm : To Extract the details of the Data sheet
'
' Author : Prasanna kumar
' Coding Start and End : 5th October,2016
' Notes  : N/A
' Parameters : N/A
' Returns : N/A
' ---------------------------------------------------------------
' Revision History
'========================================================================================================

'On Error GoTo ErrorHandler

Application.ScreenUpdating = False
Application.DisplayAlerts = False

EventState = Application.EnableEvents
Application.EnableEvents = False

CalcState = Application.Calculation
Application.Calculation = xlCalculationManual

PageBreakState = ActiveSheet.DisplayPageBreaks
ActiveSheet.DisplayPageBreaks = False

Dim Asntr As String
Dim sDa As String
Dim sSDH As String
Dim sBt As String
Dim sTs As String
Dim sPc1 As String
Dim sTM As String

sTM = "T" & Chr(38) & "M"

Dim WB_ASnTr As Workbook
Dim WS_SDH As Worksheet
Dim WS_Pc1 As Worksheet
Dim WS_Bt As Worksheet
Dim WS_Lds As Worksheet
Dim WS_PRL As Worksheet
Dim WS_NB As Worksheet
Dim WS_Pls As Worksheet
Dim WS_Ts As Worksheet
Dim WS_AsnTr As Worksheet
Dim WS_Dat As Worksheet
Dim WS_LB As Worksheet

Dim aColumn(20) As Integer
Dim rCellsda As Range
Dim BI As Variant
Dim stmp As String
Dim lTmp1 As Long
Dim lTmp2 As Long
Dim sColName As String
Dim iCol As String
Dim iresult As String
Dim sTmpsheet As String
Dim i As Long
Dim j As Long
Dim R As Long
Dim C As Long


Dim sAsnTr As String
Dim iCounter As Long
Dim sDat As String
Dim sLB As String
Dim sNB As String
Dim sPRl As String

Dim lro As Long
Dim lro1 As Long
Dim lRoCheck As Long
Dim sDbQ As String
Dim Tslro As Long
Dim Btlro As Long
Dim Pc1lro As Long
Dim LBlro As Long
Dim NBlro As Long
Dim PRlro As Long


Dim sdate As Variant

sDbQ = Chr(34)
sDa = "Column Header Data"
sAsnTr = "Master_List"
sDat = "Data"
sBt = "BT"
sTs = "Timesheet"
sPc1 = "PC"
sLB = "Leadership & BEL"
sNB = "Non-Billable"
sPRl = "Perdiem in lieu of salary"

Set WB_ASnTr = ActiveWorkbook
Set WS_Dat = WB_ASnTr.Sheets(sDat)
Set WS_Ts = WB_ASnTr.Sheets(sTs)
Set WS_Bt = WB_ASnTr.Sheets(sBt)
Set WS_Pc1 = WB_ASnTr.Sheets(sPc1)
Set WS_LB = WB_ASnTr.Sheets(sLB)
Set WS_NB = WB_ASnTr.Sheets(sNB)
Set WS_PRL = WB_ASnTr.Sheets(sPRl)

  WB_ASnTr.Sheets(sTs).Activate
  Sheets(sTs).Select
  'Timesheet
  If ActiveSheet.FilterMode Then
     Selection.AutoFilter = False
  End If
   
  Tslro = Cells(Rows.Count, "A").End(xlUp).Row

  'BT
  WB_ASnTr.Sheets(sBt).Activate
  Sheets(sBt).Select

  If ActiveSheet.FilterMode Then
     Selection.AutoFilter = False
  End If
   
   Btlro = Cells(Rows.Count, "A").End(xlUp).Row - 3
   
  'PC
  WB_ASnTr.Sheets(sPc1).Activate
  Sheets(sPc1).Select

    If ActiveSheet.FilterMode Then
         Selection.AutoFilter = False
    End If
   
  Pc1lro = Cells(Rows.Count, "A").End(xlUp).Row
  
  'Non Billable
  WB_ASnTr.Sheets(sNB).Activate
  Sheets(sNB).Select

  NBlro = Cells(Rows.Count, "B").End(xlUp).Row
  
   'Perdium in lieu of salary
  WB_ASnTr.Sheets(sPRl).Activate
  Sheets(sPRl).Select

  PRlro = Cells(Rows.Count, "B").End(xlUp).Row

  'LB
  WB_ASnTr.Sheets(sLB).Activate
  Sheets(sLB).Select

    If ActiveSheet.FilterMode Then
         Selection.AutoFilter = False
    End If
   
  LBlro = Cells(Rows.Count, "A").End(xlUp).Row

  WB_ASnTr.Sheets(sTs).Activate
  Sheets(sTs).Select
  Range("1:1").Select

  lro = Cells(Rows.Count, "B").End(xlUp).Row - 2
  
  WB_ASnTr.Sheets(sDat).Activate
  Sheets(sDat).Select
  
  'Month
  WS_Dat.Range("A5").Value = WS_Bt.Range("$A$4").Value
  WS_Dat.Range("a7").Formula = "=Text(A5," & sDQ & "MMMM" & sDQ & ")"
  WS_Dat.Range("a7").Copy
  WS_Dat.Range("a7").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
  Application.CutCopyMode = False
  
  WS_Dat.Range("a5").Value = ""
  
  'QTR
  WS_Dat.Range("c7").Value = Sheets("std hrs").Range("v12").Value
  WS_Dat.Range("b7").Formula = "=CHOOSE(ROUNDUP(MONTH(c7)/3,0)," & sDQ & "Q4" & sDQ & "," & sDQ & "Q1" & sDQ & "," & sDQ & "Q2" & sDQ & "," & sDQ & "Q3" & sDQ & ")"
  WS_Dat.Range("b7").Copy
  
  WS_Dat.Range("b7").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Application.CutCopyMode = False
  'Populating the data
  
  'Month
  WS_Dat.Range(Cells(R + 8, C + 1), Cells((R + 8 + lro) - 1, C + 1)).Formula = "=$A$7"
  
  'QTR
  WS_Dat.Range(Cells(R + 8, C + 2), Cells(R + 8 + lro - 1, C + 2)).Formula = "=$B$7"
  
  'Emp id
  WS_Dat.Range(Cells(R + 7, C + 13), Cells(R + 7 + lro, C + 13)).Formula = "='Timesheet'!A2"
  
  'Project No.
  WS_Dat.Range(Cells(R + 7, C + 24), Cells(R + 7 + lro, C + 24)).Formula = "='Timesheet'!j2"
  
  'Unique id 1
  WS_Dat.Range(Cells(R + 7, C + 3), Cells(R + 7 + lro, C + 3)).Formula = "=$M7&$X7"
  
   'Location
  WS_Dat.Range(Cells(R + 7, C + 15), Cells(R + 7 + lro, C + 15)).Formula = "='Timesheet'!C2"
  
  'Unique id 2
  WS_Dat.Range(Cells(R + 7, C + 4), Cells(R + 7 + lro, C + 4)).Formula = "=M7&X7&O7"
  
  'Annuity
  WS_Dat.Range(Cells(R + 7, C + 5), Cells(R + 7 + lro, C + 5)).Formula = "=iferror(VLOOKUP(X7,'PC'!$A$2:$Z$" & Pc1lro & ",8,0)," & sDQ & "" & sDQ & ")"
  
  'Project Staffing
  'WS_Dat.Range(Cells(R + 7, C + 6), Cells(R + 7 + lro, C + 6)).Formula = "=VLOOKUP(X7,PC!$A$2:$Z$" & Pc1lro & ",9,0)"

  'Management Type
  WS_Dat.Range(Cells(R + 7, C + 7), Cells(R + 7 + lro, C + 7)).Formula = "=iferror(VLOOKUP(X7,'PC'!$A$2:$Z$" & Pc1lro & ",9,0)," & sDQ & "" & sDQ & ")"

  'Location check
  WS_Dat.Range(Cells(R + 7, C + 8), Cells(R + 7 + lro, C + 8)).Formula = "=IF(O7=""Offshore"",""Offshore"",IF(AND(O7=""Onsite"",OR(L7=""RAK"",L7=""Dubai"")),""ME"",""Onsite""))"
  
  'Unique for HC
  WS_Dat.Range(Cells(R + 7, C + 9), Cells(R + 7 + lro, C + 9)).Formula = "=IF(M7="""",0,CONCATENATE(M7&A7))"
  
  'Head Count
  WS_Dat.Range(Cells(R + 7, C + 10), Cells(R + 7 + lro, C + 10)).Formula = "=IFERROR(MIN(1,SUMIF(I:I,I7,AZ:AZ)/168)*AZ7/SUMIF(I:I,I7,AZ:AZ),0)"
  WS_Dat.Range(Cells(R + 7, C + 10), Cells(R + 7 + lro, C + 10)).NumberFormat = "#,##0.00"
  
  'Utilisation %
  WS_Dat.Range(Cells(R + 7, C + 11), Cells(R + 7 + lro, C + 11)).Formula = "=iferror(VLOOKUP(M7,'BT'!$B$4:$I$" & Btlro & ",8,0)," & sDQ & "" & sDQ & ")"
  WS_Dat.Range(Cells(R + 7, C + 11), Cells(R + 7 + lro, C + 11)).NumberFormat = "0%"

  'Location 1
  WS_Dat.Range(Cells(R + 7, C + 12), Cells(R + 7 + lro, C + 12)).Formula = "=vlookup(m7,'bt'!$M:$P,4,0)"
    
  'FULL_NAME
  WS_Dat.Range(Cells(R + 7, C + 14), Cells(R + 7 + lro, C + 14)).Formula = "='Timesheet'!B2"
  
  'HRMS Location
  WS_Dat.Range(Cells(R + 7, C + 16), Cells(R + 7 + lro, C + 16)).Formula = "='Timesheet'!D2"
    
  'Emp Type
  WS_Dat.Range(Cells(R + 7, C + 17), Cells(R + 7 + lro, C + 17)).Formula = "='Timesheet'!E2"
  
  'Emp Classification
  WS_Dat.Range(Cells(R + 7, C + 18), Cells(R + 7 + lro, C + 18)).Formula = "=iferror(VLOOKUP(M7,'BT'!$B$4:$AF$" & Btlro & ",31,0)," & sDQ & "" & sDQ & ")"
  
  'DOJ
  WS_Dat.Range(Cells(R + 7, C + 19), Cells(R + 7 + lro, C + 19)).Formula = "=iferror(VLOOKUP(M7,'Timesheet'!$A$2:$F$" & Tslro & ",6,0)," & sDQ & "" & sDQ & ")"
  
  'DOL
  WS_Dat.Range(Cells(R + 7, C + 20), Cells(R + 7 + lro, C + 20)).Formula = "=IF(VLOOKUP(M7,'Timesheet'!$A$2:$G$" & Tslro & ",7,0)=" & sDQ & sDQ & "," & sDQ & sDQ & ",VLOOKUP($M7,Timesheet!$A2:$G$" & Tslro & ",7,0))"
  
  WS_Dat.Range(Cells(R + 7, C + 19), Cells(R + lro, C + 19)).NumberFormat = "[$-14009]dd/mm/yyyy;@"
  WS_Dat.Range(Cells(R + 7, C + 20), Cells(R + lro, C + 20)).NumberFormat = "[$-14009]dd/mm/yyyy;@"

  'GRADE
  WS_Dat.Range(Cells(R + 7, C + 21), Cells(R + 7 + lro, C + 21)).Formula = "=iferror(VLOOKUP(M7,'Timesheet'!$A$2:$H$" & Tslro & ",8,0)," & sDQ & "" & sDQ & ")"
  
  'Status
  WS_Dat.Range(Cells(R + 7, C + 22), Cells(R + 7 + lro, C + 22)).Formula = "='Timesheet'!i2"
  
  'Classification
  'WS_Dat.Range(Cells(R + 7, C + 23), Cells(R + 7 + lro, C + 23)).Formula = "=IF(CH7<>0,CH7,IF(or(V7=" & sDQ & "Bench" & sDQ & ",V7=" & sDQ & "Assigned - Internal" & sDQ & "),V7,IF(V7 =" & sDQ & "Assigned - Client" & sDQ & "," & sDQ & "Client Project - Unbilled" & sDQ & ",0)))"
  
  WS_Dat.Range(Cells(R + 7, C + 86), Cells(R + 7 + lro, C + 86)).Formula = "=IF(AR7>=1," & sDQ & "Client Project - Billed" & sDQ & ",IF(OR(r7=" & sDQ & "Leadership" & sDQ & ",R7=" & sDQ & "BEL" & sDQ & "),R7,0))"

  'Project Name
  WS_Dat.Range(Cells(R + 7, C + 25), Cells(R + 7 + lro, C + 25)).Formula = "='Timesheet'!K2"
  
  'Client
  WS_Dat.Range(Cells(R + 7, C + 26), Cells(R + 7 + lro, C + 26)).Formula = "='Timesheet'!L2"
  
  'Project Type
  WS_Dat.Range(Cells(R + 7, C + 27), Cells(R + 7 + lro, C + 27)).Formula = "='Timesheet'!M2"
  
  'Horizontal
  WS_Dat.Range(Cells(R + 7, C + 28), Cells(R + 7 + lro, C + 28)).Formula = "=iferror(VLOOKUP(M7,'BT'!$B$4:BT!$AG$" & Btlro & ",32,0)," & sDQ & "" & sDQ & ")"
  
  'Practice
  WS_Dat.Range(Cells(R + 7, C + 29), Cells(R + 7 + lro, C + 29)).Formula = "=iferror(VLOOKUP(M7,'BT'!$B$4:BT!$AH$" & Btlro & ",33,0)," & sDQ & "" & sDQ & ")"
  
  'Tower
  WS_Dat.Range(Cells(R + 7, C + 30), Cells(R + 7 + lro, C + 30)).Formula = "=AC7"
  
  'SL
  WS_Dat.Range(Cells(R + 7, C + 31), Cells(R + 7 + lro, C + 31)).Formula = "=iferror(VLOOKUP(M7,'BT'!$B$4:BT!$AI$" & Btlro & ",34,0)," & sDQ & "" & sDQ & ")"
  '=iferror(VLOOKUP(M7,Timesheet!$A$2:Timesheet!$Q$" & Tslro & ",17,0)," & sDQ & "" & sDQ & ")"
  
  'SL1
  WS_Dat.Range(Cells(R + 7, C + 32), Cells(R + 7 + lro, C + 32)).Formula = "=IF(OR(AE7=" & sDQ & "ITS Apps" & sDQ & ",AE7=" & sDQ & "Infra" & sDQ & ",AE7=" & sDQ & "Cloud" & sDQ & ",AE7=" & sDQ & "ITS Services" & sDQ & ")," & sDQ & "ITS Services" & sDQ & ",IF(AE7=" & sDQ & "Business Engagement" & sDQ & "," & sDQ & "Business Engagement" & sDQ & "," & sDQ & "Solutions" & sDQ & "))"
  
  'Sales Channel
  WS_Dat.Range(Cells(R + 7, C + 33), Cells(R + 7 + lro, C + 33)).Formula = "=iferror(VLOOKUP(X7,'PC'!$A$2:$AJ$" & Pc1lro & ",36,0)," & sDQ & "" & sDQ & ")"
  
  'Client Partner
  WS_Dat.Range(Cells(R + 7, C + 34), Cells(R + 7 + lro, C + 34)).Formula = "=iferror(VLOOKUP(X7,'PC'!$A$2:PC!$Y$" & Pc1lro & ",25,0)," & sDQ & "" & sDQ & ")"
  
  'Shadow Details
  WS_Dat.Range(Cells(R + 7, C + 35), Cells(R + 7 + lro, C + 35)).Formula = "=iferror(VLOOKUP(C7,'BT'!$D$4:BT!$J$" & Btlro & ",7,0)," & sDQ & "" & sDQ & ")"
  
  'Billable
  WS_Dat.Range(Cells(R + 7, C + 36), Cells(R + 7 + lro, C + 36)).Formula = "=iferror(VLOOKUP(M7,'Timesheet'!$A$2:$R$" & Tslro & ",18,0)," & sDQ & "" & sDQ & ")"
  
  'Non Billable
  WS_Dat.Range(Cells(R + 7, C + 37), Cells(R + 7 + lro, C + 37)).Formula = "=iferror(VLOOKUP(M7,'Timesheet'!$A$2:$S$" & Tslro & ",19,0)," & sDQ & "" & sDQ & ")"
  
  'Internal
  WS_Dat.Range(Cells(R + 7, C + 38), Cells(R + 7 + lro, C + 38)).Formula = "=iferror(VLOOKUP(M7,'Timesheet'!$A$2:$T$" & Tslro & ",20,0)," & sDQ & "" & sDQ & ")"
    
  'Leave/Company Holiday
   WS_Dat.Range(Cells(R + 7, C + 39), Cells(R + 7 + lro, C + 39)).Formula = "=iferror(VLOOKUP(M7,'Timesheet'!$A$2:$U$" & Tslro & ",21,0)," & sDQ & "" & sDQ & ")"
  
  'Unassigned
   WS_Dat.Range(Cells(R + 7, C + 40), Cells(R + 7 + lro, C + 40)).Formula = "=iferror(VLOOKUP(M7,'Timesheet'!$A$2:$V$" & Tslro & ",22,0)," & sDQ & "" & sDQ & ")"

  'Total
   WS_Dat.Range(Cells(R + 7, C + 41), Cells(R + 7 + lro, C + 41)).Formula = "=SUM(AJ7:AN7)"
   
   Call sumbyprimarykey
    
   Sheets(sDat).Select
   
   WS_Dat.Range(Cells(R + 7, C + 23), Cells(R + 7 + lro, C + 23)).Formula = "=IF(V7=""Bench"",""bench"",IF(V7=""Assigned Internal"",IF(ISERROR(INDEX('Leadership & BEL'!G:G,MATCH(M7,'Leadership & BEL'!A:A,0))),""Assigned Internal"",INDEX('Leadership & BEL'!G:G,MATCH(M7,'Leadership & BEL'!A:A,0))),IF(AR7>0,""Assigned-client-billed"",IF(ISERROR(INDEX('Leadership & BEL'!G:G,MATCH(M7,'Leadership & BEL'!A:A,0))),""client-unbilled"",INDEX('Leadership & BEL'!G:G,MATCH(M7,'Leadership & BEL'!A:A,0))))))"
  
  'Billed Hours
  'WS_Dat.Range(Cells(R + 7, C + 42), Cells(R + 7 + lro, C + 42)).Formula = "=iferror(VLOOKUP(C7,'BT'!$D$4:$Z$" & Btlro & ",17,0),0)"
  
  'Rate
  'WS_Dat.Range(Cells(R + 7, C + 43), Cells(R + 7 + lro, C + 43)).Formula = "=iferror(VLOOKUP(C7,'BT'!$D$4:$Z$" & Btlro & ",14,0)," & sDQ & "" & sDQ & ")"
  
  'Revenue
  'WS_Dat.Range(Cells(R + 7, C + 44), Cells(R + 7 + lro, C + 44)).Formula = "=IF(COUNTIFS($D$7:$D7,D7)>1,"""",SUMIFS(CF:CF,D:D,D7))"
  
  'Billable Hours 1
   WS_Dat.Range("Aa5").Value = sTM
  
   WS_Dat.Range(Cells(R + 7, C + 45), Cells(R + 7 + lro, C + 45)).Formula = "=if(AA7=AA$5,AP7,AJ7)"
  
   'Non Billable Hours 1
   WS_Dat.Range(Cells(R + 7, C + 46), Cells(R + 7 + lro, C + 46)).Formula = "=if(AJ7+AK7-AS7>0,(AJ7+AK7-AS7),0)"
       
   'Assigned hours
   WS_Dat.Range(Cells(R + 7, C + 47), Cells(R + 7 + lro, C + 47)).Formula = "=AS7+AT7"
     
   'Internal
   WS_Dat.Range(Cells(R + 7, C + 48), Cells(R + 7 + lro, C + 48)).Formula = "='Timesheet'!T2"

   'Leave/Company Holiday 1
   WS_Dat.Range(Cells(R + 7, C + 49), Cells(R + 7 + lro, C + 49)).Formula = "=AM7"
     
   'Unassigned 1
   WS_Dat.Range(Cells(R + 7, C + 50), Cells(R + 7 + lro, C + 50)).Formula = "=+IF((AO7-AU7-AV7-AW7)>0,(AO7-Au7-AV7-AW7),0)"
   
   'Available 1
   WS_Dat.Range(Cells(R + 7, C + 51), Cells(R + 7 + lro, C + 51)).Formula = "=AU7+AV7+AX7"
   
   'Total
   WS_Dat.Range(Cells(R + 7, C + 52), Cells(R + 7 + lro, C + 52)).Formula = "=AY7+AW7"
   
   'Register
   WS_Dat.Range(Cells(R + 7, C + 53), Cells(R + 7 + lro, C + 53)).Formula = "=VLOOKUP(M7,'Salary'!A:B,2,0)"
   
    Call findDuplicates
   
    Call fromToDate
    
    
   'App Salary
'   WS_Dat.Range(Cells(R + 7, C + 54), Cells(R + 7 + lro, C + 54)).Formula = "=IF(COUNTIFS($C$7:$C7,C7,$BA$7:$BA7,BA7)>1,"""",CD7)"
    
    Call AppSalary
    
    Sheets(sDat).Select
   'Other cost
   WS_Dat.Range(Cells(R + 7, C + 55), Cells(R + 7 + lro, C + 55)).Formula = "=IF(O7=" & sDQ & "Offshore" & sDQ & ",BB7*VLOOKUP(AC7,'Others'!$A$7:$B$29,2,0)/SUMIFS('Data'!BB:BB,'Data'!O:O," & sDQ & "Offshore" & sDQ & ",'Data'!AC:AC,'Data'!AC7),IF(O7=" & sDQ & "Onsite" & sDQ & ",BB7*VLOOKUP(AC7,'Others'!$A$7:$C$29,3,0)/SUMIFS('Data'!BB:BB,'Data'!O:O," & sDQ & "Onsite" & sDQ & ",'Data'!AC:AC,'Data'!AC7),0))"
   
   'Monthly Salary
   WS_Dat.Range(Cells(R + 7, C + 56), Cells(R + 7 + lro, C + 56)).Formula = "=BB7+BC7"
   
   'Perdiem in lieu of sal
   WS_Dat.Range(Cells(R + 7, C + 57), Cells(R + 7 + lro, C + 57)).Formula = "=iferror(VLOOKUP(C7,'Perdiem in lieu of salary'!$D$8:$E$" & PRlro & ",2,0)," & sDQ & "" & sDQ & ")"

   'Total Salary incl Perdiem
   WS_Dat.Range(Cells(R + 7, C + 58), Cells(R + 7 + lro, C + 58)).Formula = "=BD7+BE7"
      
   'Billed Sal
    WS_Dat.Range(Cells(R + 7, C + 59), Cells(R + 7 + lro, C + 59)).Formula = "=IFERROR(BF7*AS7/AT7,0)"
    
    'Unbilled Sal
    WS_Dat.Range(Cells(R + 7, C + 60), Cells(R + 7 + lro, C + 60)).Formula = "=IFERROR(BF7*AV7/AT7,0)"
    
    'Internal Sal
    WS_Dat.Range(Cells(R + 7, C + 61), Cells(R + 7 + lro, C + 61)).Formula = "=IFERROR(BF7*AV7/AZ7,0)"
    
    'Leave Sal
    WS_Dat.Range(Cells(R + 7, C + 62), Cells(R + 7 + lro, C + 62)).Formula = "=IFERROR(BF7*AW7/AZ7,0)"
    
    'Bech Sal
    WS_Dat.Range(Cells(R + 7, C + 63), Cells(R + 7 + lro, C + 63)).Formula = "=IFERROR(BF7*AX7/AZ7,0)"
    
    'Total Sal
    WS_Dat.Range(Cells(R + 7, C + 64), Cells(R + 7 + lro, C + 64)).Formula = "=SUM(Bg7:BK7)"
    
    'Perdiem
    WS_Dat.Range(Cells(R + 7, C + 65), Cells(R + 7 + lro, C + 65)).Formula = "=iferror(Vlookup(m7,'Non-Billable'!$A$8:$H$" & NBlro & ",5,0)," & sDQ & "" & sDQ & ")"
    
    'Travel
    WS_Dat.Range(Cells(R + 7, C + 66), Cells(R + 7 + lro, C + 66)).Formula = "=iferror(Vlookup(m7,'Non-Billable'!$A$8:$H$" & NBlro & ",6,0)," & sDQ & "" & sDQ & ")"
       
    'Hotel & Lodging
    WS_Dat.Range(Cells(R + 7, C + 67), Cells(R + 7 + lro, C + 67)).Formula = "=iferror(Vlookup(m7,'Non-Billable'!$A$8:$H$" & NBlro & ",7,0)," & sDQ & "" & sDQ & ")"
    
    'Others
    WS_Dat.Range(Cells(R + 7, C + 68), Cells(R + 7 + lro, C + 68)).Formula = "=iferror(Vlookup(m7,'Non-Billable'!$A$8:$H$" & NBlro & ",8,0)," & sDQ & "" & sDQ & ")"
    
    'Expenses
    WS_Dat.Range(Cells(R + 7, C + 69), Cells(R + 7 + lro, C + 69)).Formula = "=SUM(BM7:BP7)"
    
    'Project cost
    WS_Dat.Range(Cells(R + 7, C + 70), Cells(R + 7 + lro, C + 70)).Formula = "=BL7+BQ7"

    'Contribution
    WS_Dat.Range(Cells(R + 7, C + 71), Cells(R + 7 + lro, C + 71)).Formula = "=AR7-BR7"
    
    'Rev for bill rate
    WS_Dat.Range(Cells(R + 7, C + 72), Cells(R + 7 + lro, C + 72)).Formula = "=AR7"
    
    'Hours for bill rate
    WS_Dat.Range(Cells(R + 7, C + 73), Cells(R + 7 + lro, C + 73)).Formula = "=AS7"
    
    'Hours for Realiation rate
    WS_Dat.Range(Cells(R + 7, C + 74), Cells(R + 7 + lro, C + 74)).Formula = "=AU7"
    
    'Avg Bill rate
    WS_Dat.Range(Cells(R + 7, C + 75), Cells(R + 7 + lro, C + 75)).Formula = "=AR7/BU7"

    'Avg Realisation Rate
    WS_Dat.Range(Cells(R + 7, C + 76), Cells(R + 7 + lro, C + 76)).Formula = "=AR7/BV7"

    'Utilization
    WS_Dat.Range(Cells(R + 7, C + 77), Cells(R + 7 + lro, C + 77)).Formula = "=AS7/AZ7"

'    WS_Dat.Range(Cells(R + 7, C + 82), Cells(R + 7 + lro, C + 82)).Formula = "=IF(CE7=""TRUE"",BA7*AZ7,BA7*AZ7/SUMIF(M:M,M7,AZ:AZ))"
    WS_Dat.Range(Cells(R + 7, C + 84), Cells(R + 7 + lro, C + 84)).Formula = "=IF(COUNTIFS($D$7:$D7,D7,$CG$7:$CG7,CG7)>1,"""",CG7)"
    WS_Dat.Range(Cells(R + 7, C + 85), Cells(R + 7 + lro, C + 85)).Formula = "=if(AA7=AA$5,AP7*AQ7,IF(OR(AA7 =" & sDQ & "FBP" & sDQ & ",AA7 =" & sDQ & "Fixed Billing" & sDQ & "),VLOOKUP(C7,'BT'!$D$4:$S$" & Btlro & ",16,0),0))"
    

'
'  For i = 7 To lro
'    For j = 4 To Btlro
'      If WS_Dat.Cells(i, 3).Value = WS_Bt.Cells(j, 4).Value Then
'        If WS_Bt.Cells(j, 11).Value = WS_Dat.Cells(2, 2).Value And WS_Bt.Cells(j, 12).Value = WS_Dat.Cells(3, 2).Value Then
'          If WS_Dat.Cells(i, 22).Value = "Assigned - Client" Then
'            WS_Dat.Cells(i, 83).Value = "TRUE"
'            GoTo MM
'          End If
'        End If
'      End If
'      Next j
'MM:
'  Next i

   WB_ASnTr.Sheets(sDat).Activate
   Sheets(sDat).Range("a1").Select


   For Each BI In Array(xlEdgeTop, xlEdgeLeft, xlEdgeBottom, xlEdgeRight, xlInsideHorizontal, xlInsideVertical)
      With Range(Cells(6, 1), Cells(lro + 7, 86)).Borders(BI)
         .Weight = xlThin
         .Color = RGB(148, 138, 84)
    End With
   Next BI
  WS_Dat.Range(Cells(6, 1), Cells(lro + 7, 83)).Font.Bold = True



  Application.EnableEvents = True
  Application.Calculation = xlCalculationAutomatic
  Application.ScreenUpdating = True
  Application.DisplayAlerts = True
  
    WS_Dat.Range("C6:ci30").Select
  Selection.Cells.EntireColumn.AutoFit
  WS_Dat.Range("aw6,am6").ColumnWidth = 10
  WS_Dat.Range("L6").ColumnWidth = 24
  WS_Dat.Range("N6").ColumnWidth = 24
  WS_Dat.Range("Y6").ColumnWidth = 24
  WS_Dat.Range("AG6").ColumnWidth = 24

  Range("A6").Select
    If Not ActiveSheet.FilterMode Then
         Selection.AutoFilter
    End If

End Sub
Sub sumbyprimarykey()
'    Dim StartTime As Double
'    Dim SecondsElapsed As Double
'
'    'Remember time when macro starts
'    'StartTime = Timer
'    Application.ScreenUpdating = False
'    Application.DisplayAlerts = False
'    Application.EnableEvents = False
    
    Dim Wb As Workbook
    Dim WS_MasterList, WS_Bt As Worksheet
    Dim Master_rowCount, BT_rowCount As Integer
    
    Set Wb = ActiveWorkbook
    Set WS_MasterList = Wb.Sheets("Data")
    Set WS_Bt = Wb.Sheets("BT")
    
    WS_MasterList.Activate
    WS_MasterList.Select

    Master_rowCount = Cells(Rows.Count, "AO").End(xlUp).Row
    
    WS_Bt.Activate
    WS_Bt.Select
    
    BT_rowCount = Sheets("BT").Cells(Rows.Count, "C").End(xlUp).Row
    
    Dim BT_i, Master_i As Integer
    Dim lookup_value As String
    Dim bill_rate, revenue, billed_hours As Double

    For Master_i = 7 To Master_rowCount
        
        If Not IsError(Sheets("Data").Cells(Master_i, 4).Value) Then
            lookup_value = Sheets("Data").Cells(Master_i, 4).Value
             BT_i = Application.Match(lookup_value, ActiveSheet.Columns(3), 0) ' returning the row index from BT sheet by matching the Primarkey
            If Not IsError(BT_i) Then
                bill_rate = ActiveSheet.Cells(BT_i, 17).Value ' returning Billed_rate from BT sheet which is 17th(Q) Column
                revenue = ActiveSheet.Cells(BT_i, 19).Value ' returning revenue from BT sheet which is 19th(S) Column
                billed_hours = ActiveSheet.Cells(BT_i, 20).Value ' returning Billable_hours from BT sheet which is 20th(T) Column
                BT_i = BT_i + 1
                
                Do While lookup_value = ActiveSheet.Cells(BT_i, 3).Value 'Checking for the duplicates values on the sorted list of primarykey in BT sheet
                    bill_rate = bill_rate + ActiveSheet.Cells(BT_i, 17).Value
                    revenue = revenue + ActiveSheet.Cells(BT_i, 19).Value
                    billed_hours = billed_hours + ActiveSheet.Cells(BT_i, 20).Value
                    BT_i = BT_i + 1
                Loop
                Sheets("Data").Cells(Master_i, 44).Value = revenue
                Sheets("Data").Cells(Master_i, 43).Value = bill_rate
                Sheets("Data").Cells(Master_i, 42).Value = billed_hours
                                
            End If
        End If
    Next Master_i
'    Application.EnableEvents = True
'    Application.ScreenUpdating = True
'    Application.DisplayAlerts = True
'    ' Determine how many seconds this code will take to run
'    'SecondsElapsed = Round(Timer - StartTime, 2)
'    'Notify user in seconds
'   ' WS_MasterList.Select
'    'MsgBox "This code ran successfully in " & SecondsElapsed & " seconds", vbInformation
End Sub
'========================================================================================================
' findDuplicates
' -------------------------------------------------------------------------------------------------------
' Purpose of this programm : ' To put 'True' in the Duplicates column(82).
'                            And on that basis making 'App salary' and 'Total' column's value 0
'
' Author : Subhankar Paul, 11th January, 2017
' Notes  : N/A
' Parameters :N/A
' Returns : N/A
' ---------------------------------------------------------------
' Revision History
'========================================================================================================
Sub findDuplicates()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Dim Data_rowCount, Data_i, j As Integer
    Data_rowCount = ActiveSheet.Cells(Rows.Count, "D").End(xlUp).Row
    
    For Data_i = 7 To Data_rowCount - 1
        If Cells(Data_i, 82).Value = "" Then
            For j = 8 To Data_rowCount
                If j <> Data_i And ActiveSheet.Cells(Data_i, 3).Value = ActiveSheet.Cells(j, 3).Value Then
                    Cells(j, 82).Value = True
                    Cells(j, 52).Value = 0  'Total & App Salary value for duplicates are 0
                    Cells(j, 54).Value = 0
                End If
            Next j
        End If
    Next Data_i
End Sub
'========================================================================================================
' fromToDate
' -------------------------------------------------------------------------------------------------------
' Purpose of this programm : getting actual From & To dates from BT Sheet for each uniquekeys (empID + Proj Code)
'
' Author : Subhankar Paul, 11th January, 2017
' Notes  : The BT sheet should be sorted first by Uniquekey then run the macro
' Parameters :N/A
' Returns : N/A
' ---------------------------------------------------------------
' Revision History
'========================================================================================================
Sub fromToDate()

    Dim Wb As Workbook
    Dim WS_Dat, WS_Bt As Worksheet
    Dim Data_rowCount, BT_rowCount, Data_i, BT_i As Integer
    Dim from_date, to_date, start_of_month, second_of_month As Date
    Dim date_flag As Boolean
    
    Dim lookup_value As String
    Set Wb = ActiveWorkbook
    Set WS_Dat = Wb.Sheets("Data")
    Set WS_Bt = Wb.Sheets("BT")
    
    WS_Bt.Activate
    WS_Bt.Select
    
    start_of_month = DateSerial(Year(Cells(4, 11).Value), Month(Cells(4, 11).Value), 1)
    second_of_month = DateSerial(Year(Cells(4, 11).Value), Month(Cells(4, 11).Value), 2)
 
    
    BT_rowCount = Sheets("BT").Cells(Rows.Count, "C").End(xlUp).Row
    Data_rowCount = Sheets("Data").Cells(Rows.Count, "D").End(xlUp).Row
    date_flag = True
    
    For Data_i = 7 To Data_rowCount
        If Not IsError(WS_Dat.Cells(Data_i, 3).Value) Then 'Checking UniqueID
            lookup_value = WS_Dat.Cells(Data_i, 3).Value
            date_flag = True
            For BT_i = 4 To BT_rowCount
                If lookup_value = WS_Bt.Cells(BT_i, 4).Value Then
                    date_flag = False

                    from_date = WS_Bt.Cells(BT_i, 11).Value
                    to_date = WS_Bt.Cells(BT_i, 12).Value
                    
                    BT_i = BT_i + 1
                    
                    Do While lookup_value = WS_Bt.Cells(BT_i, 4).Value
                        If WS_Bt.Cells(BT_i, 11).Value < from_date Then
                            from_date = WS_Bt.Cells(BT_i, 11).Value
                        End If
                        If WS_Bt.Cells(BT_i, 12).Value > to_date Then
                            to_date = WS_Bt.Cells(BT_i, 12).Value
                        End If
                        BT_i = BT_i + 1
                    Loop
                    
                    WS_Dat.Cells(Data_i, 80).Value = from_date
                    WS_Dat.Cells(Data_i, 81).Value = to_date
                    Exit For
                End If
            Next BT_i
            If date_flag = True Then 'If this unique is not in BT Sheet putting start date of the month
                WS_Dat.Cells(Data_i, 80).Value = start_of_month
                WS_Dat.Cells(Data_i, 81).Value = second_of_month
            End If
        End If
    Next Data_i
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
End Sub


Function Appotion(empId As Variant, i As Variant) As Long
    Dim sum_of_total, total, register, result As Double
    Dim Data_i As Long
    Data_i = i
    sum_of_total = Application.SumIf(Range("M:M"), empId, Range("AZ:AZ"))
    Do While Cells(Data_i, 13).Value = empId
        total = Cells(Data_i, 52).Value
        register = Cells(Data_i, 53).Value
        If Not IsError(register) Then
            result = (total / sum_of_total) * register
            Cells(Data_i, 54).Value = result
        End If
        Data_i = Data_i + 1
    Loop
    Appotion = Data_i
End Function

Function SuperChecking(ByVal empId As Variant, i As Variant, start_of_month As Variant, end_of_month As Variant) As Boolean

Dim from_date, to_date As Variant
Dim Data_i As Long
Data_i = i
Do While Cells(Data_i, 13).Value = empId And Cells(Data_i, 22).Value = "Assigned - Client"
    
    from_date = DateValue(Cells(Data_i, 80).Value)
    to_date = DateValue(Cells(Data_i, 81).Value)
    
    If from_date = start_of_month And to_date = end_of_month Then
        SuperChecking = True
        Exit Function
    End If
    Data_i = Data_i + 1
Loop

SuperChecking = False

End Function

Function SuperAppotion(ByVal empId As Variant, i As Variant) As Long

    Dim sum_of_total, total, register, result As Double
    Dim Data_i As Long
    Data_i = i
        
    sum_of_total = Application.SumIfs(Range("AZ:AZ"), Range("M:M"), empId, Range("V:V"), "Assigned - Client")
    
    Do While Cells(Data_i, 13).Value = empId And Cells(Data_i, 22).Value = "Assigned - Client"
        total = Cells(Data_i, 52).Value
        register = Cells(Data_i, 53).Value
        If Not IsError(register) Then
            result = (total / sum_of_total) * register
            Cells(Data_i, 54).Value = result
        End If
        Data_i = Data_i + 1
    Loop
    
    Do While Cells(Data_i, 13).Value = empId And Cells(Data_i, 22).Value <> "Assigned - Client"
        Cells(Data_i, 54).Value = 0
        Data_i = Data_i + 1
    Loop
    SuperAppotion = Data_i
End Function

'========================================================================================================
' AppSalary
' -------------------------------------------------------------------------------------------------------
' Purpose of this programm :
'
' Author : Subhankar Paul, 12th January, 2017
' Notes  : TimeSheet Should be sorted by EmpID
' Parameters : N/A
' Returns : N/A
' ---------------------------------------------------------------
' Revision History
'========================================================================================================

Sub AppSalary()
    
    Dim Wb As Workbook
    Dim WS_Dat, WS_Bt As Worksheet
    
    Set Wb = ActiveWorkbook
    Set WS_Dat = Wb.Sheets("Data")
    
    WS_Dat.Activate
    WS_Dat.Select

    Dim LastRow As Long
    
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Range("C7:CH" & LastRow).Sort _
        Key1:=Range("M7:M" & LastRow), Order1:=xlAscending, _
        Key2:=Range("V7:V" & LastRow), Order2:=xlAscending, _
        Header:=xlNo

    
    Dim Data_rowCount, empId, Data_i, j As Long
    Dim UniqueKey, status As String
    Dim start_of_month, end_of_month As Date
    Dim from_date, to_date As Date
    Data_rowCount = ActiveSheet.Cells(Rows.Count, "D").End(xlUp).Row
    
    end_of_month = DateSerial(Year(Cells(7, 80).Value), Month(Cells(7, 80).Value) + 1, 0)
    start_of_month = DateSerial(Year(Cells(7, 80).Value), Month(Cells(7, 80).Value), 1)
    Data_i = 7
    Do While Data_i <= Data_rowCount
        empId = Cells(Data_i, 13).Value
        status = Cells(Data_i, 22).Value
        If status = "Assigned - Client" Then
            If SuperChecking(empId, Data_i, start_of_month, end_of_month) Then
                Data_i = SuperAppotion(empId, Data_i)
                GoTo NextIteration
            Else
                '******** Not full Month Section********
                If Cells(Data_i, 80).Value <> "" Or Cells(Data_i, 81).Value <> "" Then
                    from_date = DateValue(Cells(Data_i, 80).Value)
                    to_date = DateValue(Cells(Data_i, 81).Value)
    '                MsgBox from_date & " " & to_date & " " & start_of_month & " " & end_of_month, vbInformation
                
                    If from_date <> start_of_month Or to_date <> end_of_month Then
                    '   MsgBox "Not Full Month call Appotion function ", vbInformation
                        Data_i = Appotion(empId, Data_i)
                        GoTo NextIteration
                    End If
                End If
                '******** Full Month Section***********
                
                ' Checking whether another Assigned Client of same empID available
                
                If Cells(Data_i + 1, 13).Value = empId Then ' multiple projects
                    If Cells(Data_i + 1, 22).Value = "Assigned - Client" Then
                        Data_i = Appotion(empId, Data_i)
                    Else
                        Cells(Data_i, 54).Value = Cells(Data_i, 53).Value
                        Do While Cells(Data_i + 1, 13).Value = empId And Cells(Data_i + 1, 22).Value <> "Assigned - Client"
                            Data_i = Data_i + 1
                            Cells(Data_i, 54).Value = 0
                        Loop
                        Data_i = Data_i + 1
                    End If
                ' Transfer full register to App salary
                Else
                    If Not IsError(Cells(Data_i, 53).Value) Then
                        Cells(Data_i, 54).Value = Cells(Data_i, 53).Value
                    End If
                    Data_i = Data_i + 1
                End If
            End If
        Else    ' for Status is Bench or Internal
            
            If Cells(Data_i + 1, 13).Value = empId Then 'Checking for multiple projects
                ' Appotioning w.r.t. Total
                Data_i = Appotion(empId, Data_i)
            Else
                ' Transfer full register to App salary
                If Not IsError(Cells(Data_i, 53).Value) Then
                    Cells(Data_i, 54).Value = Cells(Data_i, 53).Value
                End If
                Data_i = Data_i + 1
            End If
        End If
NextIteration:
    Loop
    
End Sub


'========================================================================================================
' SortByEmpId_Status
' -------------------------------------------------------------------------------------------------------
' Purpose of this programm : Multi Level Sorting with EmpID then Status
'                            to make AppSalary procedure work correctly
'
' Author : Subhankar Paul, 12th January, 2017
' Notes  : TimeSheet Should be sorted by EmpID
' Parameters : N/A
' Returns : N/A
' ---------------------------------------------------------------
' Revision History
'========================================================================================================

Sub SortByEmpId_Status()
    
    Dim LastRow As Long
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Range("A7:CH" & LastRow).Sort _
        Key1:=Range("M7:M" & LastRow), Order1:=xlAscending, _
        Key2:=Range("V7:V" & LastRow), Order2:=xlAscending, _
        Header:=xlNo

End Sub

