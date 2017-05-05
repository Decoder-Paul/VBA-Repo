Attribute VB_Name = "DeleteRowM"
Option Explicit

Sub pNYLDD()

Dim lro As Long
Dim lro1 As Long
Dim lroPr As Long

Dim WB As Workbook
Dim WS_In As Worksheet
Dim WS_DA As Worksheet

Dim sDa As String
Dim sIn As String
Dim BI As Variant

Dim i As Long
Dim j As Long
Dim k As Long
Dim R As Long
Dim x As Long
Dim C As Integer
Dim Cellda As Range

Dim iStart As Integer
Dim iClose As Integer

Dim sValue As String
Dim sRes As String

Dim sResp As String
Dim sRespo As String

sIn = "Incident"
sDa = "MainData"

Set WB = ActiveWorkbook
Set WS_In = WB.Sheets(sIn)
Set WS_DA = WB.Sheets(sDa)

WB.Sheets(sIn).Activate
Sheets(sIn).Select
lro = Cells(Rows.Count, "A").End(xlUp).Row

'Sorting to take out the Problem to the next sheet
WS_In.Range("A2:BZ" & lro).Sort Key1:=Range("D2"), Order1:=xlDescending

'identifying Problem Request
WB.Sheets(sIn).Activate
WS_In.Range("d2").Select

For i = 2 To lro
    If WS_In.Cells(i, 4).Value = "" Then
        Exit For
    End If
Next i
lroPr = i - 1
WS_In.Range(WS_In.Cells(2, 4), WS_In.Cells(lroPr, 4)).Value = "PRB"

Sheets(sIn).Activate
Sheets(sIn).Select

lro = Cells(Rows.Count, "A").End(xlUp).Row
'if Problem needs to be removed the below line needs to be executed.
'WS_In.Range(Cells(2, 1), Cells(lro, 16)).Delete Shift:=xlUp

'Incident Sheet for Priority and 64 is for Effort
WS_In.Range(Cells(R + lroPr + 1, C + 4), Cells(R + lro, C + 4)).Formula = "=IF(OR(LOWER(LEFT(TRIM(H" & lroPr + 1 & "),7))=""request"",LOWER(LEFT(H" & lroPr & ",4))=""task""),""SRQ"",""INC"")"
WS_In.Range(Cells(R + 2, C + 4), Cells(R + lro, C + 4)).Copy
WS_In.Range(Cells(R + 2, C + 4), Cells(R + lro, C + 4)).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Application.CutCopyMode = False

WS_In.Range(Cells(R + 2, C + 63), Cells(R + lro, C + 63)).Formula = "=NUMBERVALUE(LEFT(G2,1))" 'Priority to number
WS_In.Range(Cells(R + 2, C + 64), Cells(R + lro, C + 64)).Formula = "0" ' Assign 0 to effort column

WS_In.Range(Cells(R + 2, C + 18), Cells(R + lro, C + 18)).Formula = "=IF(LOWER(RIGHT(B2,5))=""ution"",""Resolution"",""Response"")"
WS_In.Range(Cells(R + 2, C + 19), Cells(R + lro, C + 19)).Formula = "=IF(AND(VALUE(LEFT(G2,1))=VALUE(MID(B2,10,1)),R2=""Resolution""),""Yes"",""No"")"
WS_In.Range(Cells(R + 2, C + 20), Cells(R + lro, C + 20)).Formula = "=IF(S2=""Yes"",""Yes"","""")"


WS_In.Range("A2:BZ" & lro).Sort Key1:=Range("a2"), Order1:=xlAscending


''For i = 2 To lro
''    sValue = WS_In.Cells(i, 1).Value
''    sRes = WS_In.Cells(i, 2).Value
''    sRes = Right(sRes, Len(sRes) - InStrRev(sRes, " "))
''
''    If Not IsError(sValue) Then
''        If sRes = "response" Then
''            If WS_In.Cells(i, 15).Value = "False" Then
''                WS_In.Cells(i, 65).Value = "Y" 'Response
''                WS_In.Cells(i, 66).Value = "NA" ' Resolution
''            ElseIf WS_In.Cells(i, 15).Value = "" Then
''                WS_In.Cells(i, 65).Value = "NA" 'Response
''                WS_In.Cells(i, 66).Value = "NA" ' Resolution
''            ElseIf WS_In.Cells(i, 15).Value = "True" Then
''                WS_In.Cells(i, 65).Value = "N"
''                WS_In.Cells(i, 66).Value = "NA"
''            End If
''        ElseIf sRes = "resolution" Then
''            If WS_In.Cells(i, 15).Value = "False" Then
''                WS_In.Cells(i, 66).Value = "Y"
''                WS_In.Cells(i, 65).Value = "NA"
''
''            ElseIf WS_In.Cells(i, 15).Value = "" Then
''                WS_In.Cells(i, 65).Value = "NA" 'Response
''                WS_In.Cells(i, 66).Value = "NA" ' Resolution
''
''            ElseIf WS_In.Cells(i, 15).Value = "True" Then
''                WS_In.Cells(i, 66).Value = "N"
''                WS_In.Cells(i, 65).Value = "NA"
''            End If
''        End If
''    End If
''
''Next i
''
''Dim Data_i As String
''Dim first_Occurence_Flag As Boolean
''Dim start_index As Long
''Dim end_index As Long
''Dim ultimate_resp_value As String
''ultimate_resp_value = "NA"
''first_Occurence_Flag = True
''
''For x = 2 To lro
''    If WS_In.Cells(x, 65).Value <> "NA" Then
''        ultimate_resp_value = WS_In.Cells(x, 65).Value
''    End If
''    If WS_In.Cells(x, 1).Value = WS_In.Cells(x + 1, 1).Value Then
''          If first_Occurence_Flag = True Then
''               start_index = x
''               first_Occurence_Flag = False
''          End If
''
''    Else
''        If first_Occurence_Flag = True Then
''            start_index = x
''        End If
''
''        end_index = x
''        first_Occurence_Flag = True
''        Range(Cells(start_index, 67), Cells(end_index, 67)).Value = ultimate_resp_value
''        ultimate_resp_value = "NA"
''    End If
''Next x
''
''first_Occurence_Flag = True
''
''For x = 2 To lro
''    If WS_In.Cells(x, 66).Value <> "NA" Then
''        ultimate_resp_value = WS_In.Cells(x, 66).Value
''    End If
''    If WS_In.Cells(x, 1).Value = WS_In.Cells(x + 1, 1).Value Then
''          If first_Occurence_Flag = True Then
''               start_index = x
''               first_Occurence_Flag = False
''          End If
''
''    Else
''        If first_Occurence_Flag = True Then
''            start_index = x
''        End If
''        end_index = x
''        first_Occurence_Flag = True
''        Range(Cells(start_index, 68), Cells(end_index, 68)).Value = ultimate_resp_value
''        ultimate_resp_value = "NA"
''    End If
''Next x


For i = 2 To lro
    sValue = WS_In.Cells(i, 1).Value
    sRes = WS_In.Cells(i, 2).Value
    sRes = Right(sRes, Len(sRes) - InStrRev(sRes, " "))

        If Not IsError(sValue) Then
        If sRes = "response" Then
            If WS_In.Cells(i, 15).Value = "False" Then
                WS_In.Cells(i, 67).Value = "Y" 'Response
                WS_In.Cells(i, 68).Value = "N" ' Resolution
            ElseIf WS_In.Cells(i, 15).Value = "" Then
                WS_In.Cells(i, 67).Value = "NA" 'Response
                WS_In.Cells(i, 68).Value = "NA" ' Resolution
            ElseIf WS_In.Cells(i, 15).Value = "True" Then
                WS_In.Cells(i, 67).Value = "N"
                WS_In.Cells(i, 68).Value = "N"
            End If
        ElseIf sRes = "resolution" Then
            If WS_In.Cells(i, 15).Value = "False" Then
                WS_In.Cells(i, 67).Value = "Y"
                WS_In.Cells(i, 68).Value = "Y"
                
            ElseIf WS_In.Cells(i, 15).Value = "" Then
                WS_In.Cells(i, 67).Value = "NA" 'Response
                WS_In.Cells(i, 68).Value = "NA" ' Resolution
            
            ElseIf WS_In.Cells(i, 15).Value = "True" Then
                WS_In.Cells(i, 67).Value = "Y"
                WS_In.Cells(i, 68).Value = "N"
            End If
        End If
    End If
Next i

'Even though the resolution is met some tikets are not closed , so we need to make the resolution not met
'Actually the condition should be resolution met when the ticket is closed
For x = 2 To lro
    If WS_In.Cells(x, 12).Value = "" Then
        WS_In.Cells(x, 68).Value = "N"
    End If
Next x


WS_In.Range("A2:BZ" & lro).Sort Key1:=Range("T2"), Order1:=xlAscending
WS_In.Range("A2:BZ" & lro).Sort Key1:=Range("T2"), Order1:=xlDescending

'Duplicate deletion logic
lro = Cells(Rows.Count, "A").End(xlUp).Row

 WS_In.Range("S2:S" & lro).Select
  Set Cellda = Selection.Find(What:="No", LookIn:=xlValues, lookat:=xlPart, SearchDirection:=xlNext)
    
      If Cellda Is Nothing Then
            
      Else
            j = Cellda.Row
            WS_In.Range(Cells(j + 1, 1), Cells(lro, 90)).Delete Shift:=xlUp
      End If

lro = WS_In.Cells(Rows.Count, "B").End(xlUp).Row
WS_In.Range("A2:BZ" & lro).Sort Key1:=Range("A2"), Order1:=xlAscending

For i = 2 To lro
    If WS_In.Cells(i, 1).Value = WS_In.Cells(i + 1, 1).Value Then
        WS_In.Cells(i, 1).Value = ""
    End If
Next i

WS_In.Range("A2:BZ" & lro).Sort Key1:=Range("a2"), Order1:=xlAscending

For i = 2 To lro + 10
    If WS_In.Cells(i, 1).Value = "" Then
        WS_In.Range(Cells(i, 1), Cells(i, 80)).Delete Shift:=xlUp
    End If
Next i

lro = Cells(Rows.Count, "A").End(xlUp).Row

WB.Sheets(sDa).Activate
Sheets(sDa).Range("A1").Select

'Clearing the contents in the MainData Sheet Cells
lro1 = WS_DA.Cells(WS_DA.Rows.Count, "A").End(xlUp).Row
If lro1 > 4 Then
    WS_DA.Range("A4:Z" & (lro1)).Clear
End If

For i = 4 To (lro + 2)
    Sheets(sDa).Cells(i, 1).Value = i - 3
Next i

'copying column D(Problem) from Incident sheet to MainData sheet column B(Ticket Type)
WS_In.Range("D2:D" & lro).Copy
WS_DA.Range("B4").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

'copying column BM(Calculated Yes or No based on Made SLA ) to MainData sheet column C(Response SLA)
WS_In.Range("BO2:BO" & lro).Copy
WS_DA.Range("C4").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
'
WS_In.Range("BP2:BP" & lro).Copy
WS_DA.Range("D4").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

WS_In.Range("A2:A" & lro).Copy
WS_DA.Range("E4").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

WS_In.Range("N2:N" & lro).Copy
WS_DA.Range("F4").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

WS_In.Range("E2:E" & lro).Copy
WS_DA.Range("g4").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

WS_In.Range("F2:F" & lro).Copy
WS_DA.Range("H4").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

WS_In.Range("I2:I" & lro).Copy
WS_DA.Range("I4").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
WS_DA.Range("P4").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

WS_DA.Range(Cells(R + 4, C + 9), Cells(R + lro + 2, C + 9)).NumberFormat = "[$-14009]dd-mm-yyyy;@"
WS_DA.Range(Cells(R + 4, C + 16), Cells(R + lro + 2, C + 16)).NumberFormat = "[$-14009]dd-mm-yyyy;@"

WS_In.Range("L2:L" & lro).Copy
WS_DA.Range("J4").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
WS_DA.Range(Cells(R + 4, C + 10), Cells(R + lro + 2, C + 10)).NumberFormat = "[$-14009]dd-mm-yyyy;@"

WS_In.Range("Bk2:Bk" & lro).Copy
WS_DA.Range("K4").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

WS_In.Range("BL2:BL" & lro).Copy
WS_DA.Range("L4").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

WS_In.Range("J2:J" & lro).Copy
WS_DA.Range("M4").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

WS_DA.Range("N4:N" & lro + 2).Value = "NYL"

Range(Cells(R + 4, C + 1), Cells(R + lro + 3, C + 16)).Select
With Selection
    .Columns.AutoFit
End With

For Each BI In Array(xlEdgeTop, xlEdgeLeft, xlEdgeBottom, xlEdgeRight, xlInsideHorizontal, xlInsideVertical)
    With Range(Cells(R + 4, C + 1), Cells(R + 3 + lro, C + 16)).Borders(BI)
         .Weight = xlThin
         .Color = RGB(148, 138, 84)
    End With
 Next BI

End Sub
Sub pMASDD()

Application.DisplayAlerts = False

Dim lro As Long
Dim lro1 As Long

Dim WB As Workbook
Dim WS_In As Worksheet
Dim WS_DA As Worksheet

Dim sDa As String
Dim sIn As String
Dim BI As Variant

Dim i As Long
Dim j As Long

Dim Cellda As Range


sIn = "Incident"
sDa = "MainData"

Set WB = ActiveWorkbook
Set WS_In = WB.Sheets(sIn)
Set WS_DA = WB.Sheets(sDa)

WB.Sheets(sIn).Activate
Sheets(sIn).Select
lro = Cells(Rows.Count, "A").End(xlUp).Row

'Sorting to take out the Problem to the next sheet
WS_In.Range("A2:P" & lro).Sort Key1:=Range("B2"), Order1:=xlAscending

Sheets(sIn).Activate
Sheets(sIn).Select

lro = WS_In.Cells(Rows.Count, "A").End(xlUp).Row

'Sorting to take out the Problem to the next sheet
WS_In.Range("A2:P" & lro).Sort Key1:=Range("a2"), Order1:=xlAscending
For i = 2 To lro
    If WS_In.Cells(i, 2).Value = "ACT" Then
        WS_In.Cells(i, 2).Value = "CHG"
    End If
'If ticket type is CHG and Priority is empty , considering priority 3 @Shambhavi
    If WS_In.Cells(i, 2).Value = "CHG" And WS_In.Cells(i, 12).Value = "" Then
        WS_In.Cells(i, 12).Value = 3
    End If
    
'Making all P4 and P5 tickets Resolution and Response "Y" because there is no SLA for the tickets @mathews
    If WS_In.Cells(i, 12).Value = 4 Or WS_In.Cells(i, 12).Value = 5 And WS_In(i, 11).Value <> "" Then
       WS_In.Cells(i, 3).Value = "Y"
       WS_In.Cells(i, 4).Value = "Y"
    End If
Next i



'Converting String Data to Number and Date
WS_In.Range(Cells(2, 20), Cells(lro, 20)).Formula = "=Numbervalue(M2)"
WS_In.Range(Cells(2, 19), Cells(lro, 19)).Formula = "=Numbervalue(l2)"
WS_In.Range(Cells(2, 18), Cells(lro, 18)).Formula = "=IF(K2="""","""",DATEVALUE(K2))"
WS_In.Range(Cells(2, 17), Cells(lro, 17)).Formula = "=IF(J2="""","""",DATEVALUE(J2))"
WS_In.Range(Cells(2, 16), Cells(lro, 16)).Formula = "=IF(I2="""","""",DATEVALUE(I2))"

WS_In.Range(Cells(2, 16), Cells(lro, 20)).Copy
WS_In.Range(Cells(2, 9), Cells(lro, 13)).PasteSpecial Paste:=xlPasteValues

'When calculation column to be deleted.
Columns("O:S").EntireColumn.Delete

WS_In.Range(Cells(2, 10), Cells(lro, 10)).Copy
WS_In.Range("Q2").PasteSpecial Paste:=xlPasteValues
WS_In.Range(Cells(2, 15), Cells(lro, 15)).Value = ""
Application.CutCopyMode = False
Columns(10).EntireColumn.Delete

WS_In.Range("A2:P" & lro).Sort Key1:=Range("G2"), Order1:=xlAscending
WS_In.Range("G2:G" & lro).Select
 
Set Cellda = Selection.Find(What:="ESM", LookIn:=xlValues, lookat:=xlPart, SearchDirection:=xlNext)
    
      If Cellda Is Nothing Then
            
      Else
            j = Cellda.Row
            WS_In.Range(Cells(j, 1), Cells(lro, 16)).Cut Destination:=WS_In.Cells(2, 27)
      End If

lro = WS_In.Cells(Rows.Count, "A").End(xlUp).Row

WB.Sheets(sDa).Activate
Sheets(sDa).Range("A1").Select

'Clearing the contents in the MainData Sheet Cells
lro1 = WS_DA.Cells(WS_DA.Rows.Count, "A").End(xlUp).Row

If lro1 > 4 Then
    WS_DA.Range("A4:Z" & (lro1)).Clear
End If

WS_In.Range("A2:Q" & lro).Copy
WS_DA.Range("A4").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
WS_DA.Range("N4:N" & lro + 2).Value = "Master Card EMO"

Range(Cells(4, 1), Cells(lro + 2, 16)).Select

With Selection
    .Columns.AutoFit
End With

For Each BI In Array(xlEdgeTop, xlEdgeLeft, xlEdgeBottom, xlEdgeRight, xlInsideHorizontal, xlInsideVertical)
    With Range(Cells(4, 1), Cells(2 + lro, 16)).Borders(BI)
         .Weight = xlThin
         .Color = RGB(148, 138, 84)
    End With
 Next BI

End Sub

Sub pMASDD1()

Dim lro As Long
Dim lro1 As Long

Dim WB As Workbook
Dim WS_In As Worksheet
Dim WS_DA As Worksheet

Dim sDa As String
Dim sIn As String
Dim BI As Variant

sIn = "Incident"
sDa = "MainData"

Set WB = ActiveWorkbook
Set WS_In = WB.Sheets(sIn)
Set WS_DA = WB.Sheets(sDa)

WB.Sheets(sIn).Activate
Sheets(sIn).Range("A1").Select

Range("A:Z").EntireColumn.Delete
lro = WS_In.Cells(Rows.Count, "A").End(xlUp).Row

WB.Sheets(sDa).Activate
Sheets(sDa).Range("A1").Select

'Clearing the contents in the MainData Sheet Cells
lro1 = WS_DA.Cells(WS_DA.Rows.Count, "A").End(xlUp).Row

If lro1 > 4 Then
    WS_DA.Range("A4:Z" & (lro1)).Clear
End If

WS_In.Range("A2:Q" & lro).Copy
WS_DA.Range("A4").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
WS_DA.Range("N4:N" & lro + 2).Value = "Master Card ESM"

Range(Cells(4, 1), Cells(lro + 2, 16)).Select

With Selection
    .Columns.AutoFit
End With

For Each BI In Array(xlEdgeTop, xlEdgeLeft, xlEdgeBottom, xlEdgeRight, xlInsideHorizontal, xlInsideVertical)
    With Range(Cells(4, 1), Cells(2 + lro, 16)).Borders(BI)
         .Weight = xlThin
         .Color = RGB(148, 138, 84)
    End With
 Next BI

End Sub
Sub pATICDD()

Application.DisplayAlerts = False

Dim lro As Long
Dim lro1 As Long

Dim WB As Workbook
Dim WS_In As Worksheet
Dim WS_DA As Worksheet

Dim sDa As String
Dim sIn As String
Dim BI As Variant

Dim i As Long

Dim R As Long
Dim C As Integer

sIn = "Incident"
sDa = "MainData"

Set WB = ActiveWorkbook
Set WS_In = WB.Sheets(sIn)
Set WS_DA = WB.Sheets(sDa)

WB.Sheets(sIn).Activate
Sheets(sIn).Select
lro = Cells(Rows.Count, "A").End(xlUp).Row

WB.Sheets(sDa).Activate
Sheets(sDa).Range("A1").Select

'Clearing the contents in the MainData Sheet Cells
lro1 = WS_DA.Cells(WS_DA.Rows.Count, "A").End(xlUp).Row
If lro1 > 4 Then
    WS_DA.Range("A4:Z" & (lro1)).ClearContents
End If

For i = 4 To (lro + 2)
    Sheets(sDa).Cells(i, 1).Value = i - 3
Next i

'Sorting to take out the Problem to the next sheet
'WS_In.Range("A2:P" & lro).Sort Key1:=Range("B2"), Order1:=xlAscending

Sheets(sIn).Activate
Sheets(sIn).Select

lro = WS_In.Cells(Rows.Count, "A").End(xlUp).Row

'Sorting to take out the Problem to the next sheet
'WS_In.Range("A2:P" & lro).Sort Key1:=Range("a2"), Order1:=xlAscending

For i = 2 To lro


    If WS_In.Cells(i, 1).Value = "Bug" Or WS_In.Cells(i, 1).Value = "Data Fixes" Then
        WS_In.Cells(i, 1).Value = "PRB"
    Else
        WS_In.Cells(i, 1).Value = "SRQ"
    End If

    If WS_In.Cells(i, 6).Value = "Normal Queue" Then
        WS_In.Cells(i, 6).Value = 4
    End If
    If WS_In.Cells(i, 6).Value = "Major" Then
        WS_In.Cells(i, 6).Value = 3
    End If
    
'-- Editing The Resolution SLA as Y for Fixed Resolution @Subhankar

    If WS_In.Cells(i, 8).Value = "Fixed" Then
        WS_In.Cells(i, 8).Value = "Y"
    Else
        WS_In.Cells(i, 8).Value = "N"
        WS_In.Cells(i, 10).Value = ""
    End If

Next i
Sheets(sDa).Select
WS_DA.Range(Cells(R + 4, C + 3), Cells(R + lro + 2, C + 3)).Formula = "Y"

'--- Copying Resolution Column to the Resolution SLA @Subhankar
WS_In.Range("H2:H" & lro).Copy
WS_DA.Range("D4").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

WS_DA.Range(Cells(R + 4, C + 12), Cells(R + lro + 2, C + 12)).Formula = "0"

WS_In.Range("A2:A" & lro).Copy
WS_DA.Range("B4").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

WS_In.Range("B2:B" & lro).Copy
WS_DA.Range("E4").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False


WS_In.Range("C2:C" & lro).Copy
WS_DA.Range("F4").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False


WS_In.Range("F2:F" & lro).Copy
WS_DA.Range("K4").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False


WS_In.Range("I2:I" & lro).Copy
WS_DA.Range("I4").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

WS_DA.Range("P4").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

WS_DA.Range(Cells(R + 4, C + 9), Cells(R + lro + 2, C + 9)).NumberFormat = "[$-14009]dd-mm-yyyy;@"
WS_DA.Range(Cells(R + 4, C + 16), Cells(R + lro + 2, C + 16)).NumberFormat = "[$-14009]dd-mm-yyyy;@"

WS_In.Range("J2:J" & lro).Copy
WS_DA.Range("J4").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
WS_DA.Range(Cells(R + 4, C + 10), Cells(R + lro + 2, C + 10)).NumberFormat = "[$-14009]dd-mm-yyyy;@"

WS_In.Range("G2:G" & lro).Copy
WS_DA.Range("M4").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

WS_In.Range("D2:D" & lro).Copy
WS_DA.Range("H4").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

WS_DA.Range("N4:N" & lro + 2).Value = "ATIC"
Call num_Of_Days

Range(Cells(2, 1), Cells(lro + 2, 13)).Select

With Selection
    .Columns.AutoFit
End With

For Each BI In Array(xlEdgeTop, xlEdgeLeft, xlEdgeBottom, xlEdgeRight, xlInsideHorizontal, xlInsideVertical)
    With Range(Cells(4, 1), Cells(2 + lro, 16)).Borders(BI)
         .Weight = xlThin
         .Color = RGB(148, 138, 84)
    End With
 Next BI

End Sub

Sub pIQPCDD()

Application.DisplayAlerts = False

Dim lro As Long
Dim lro1 As Long
Dim Diff As Integer


Dim WB As Workbook
Dim WS_In As Worksheet
Dim WS_DA As Worksheet

Dim sDa As String
Dim sIn As String
Dim BI As Variant
Dim IQPC_TicketType As String

Dim i As Long

Dim R As Long
Dim C As Integer

sIn = "Incident"
sDa = "MainData"

Set WB = ActiveWorkbook
Set WS_In = WB.Sheets(sIn)
Set WS_DA = WB.Sheets(sDa)

WB.Sheets(sIn).Activate
Sheets(sIn).Select
lro = Cells(Rows.Count, "A").End(xlUp).Row

WB.Sheets(sDa).Activate
Sheets(sDa).Range("A1").Select

'Clearing the contents in the MainData Sheet Cells
lro1 = WS_DA.Cells(WS_DA.Rows.Count, "A").End(xlUp).Row
If lro1 > 4 Then
    WS_DA.Range("A4:Z" & (lro1)).ClearContents
End If

For i = 4 To (lro + 2)
    Sheets(sDa).Cells(i, 1).Value = i - 3
Next i

WB.Sheets(sIn).Activate
Sheets(sIn).Select

'--- The Resolution SLA is Calculated on the basis of the status @Subhankar
WS_In.Cells(1, 12).Value = "Resolution SLA"
For i = 2 To lro
    WS_In.Cells(i, 5).Value = "CHG"

    If WS_In.Cells(i, 7).Value = "'-" Then
            WS_In.Cells(i, 7).Value = ""
    End If
 ' Diff = (Date - 1) - WS_In.Cells(i, 10).Value
  
    If WS_In.Cells(i, 7).Value <> "" Then
    
        WS_In.Cells(i, 12).Value = "Y"
    Else
        WS_In.Cells(i, 12).Value = "N"
    End If
Next i

Sheets(sDa).Select
'-- For All received ticket the response SLA is Y
WS_DA.Range(Cells(R + 4, C + 3), Cells(R + lro + 2, C + 3)).Formula = "Y"

'-- Resolution SLA Pasted from incident @Subhankar
WS_In.Range("L2:L" & lro).Copy
WS_DA.Range("D4").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

WS_DA.Range(Cells(R + 4, C + 12), Cells(R + lro + 2, C + 12)).Formula = "0"

WS_In.Range("A2:A" & lro).Copy
WS_DA.Range("E4").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False


WS_In.Range("B2:B" & lro).Copy
WS_DA.Range("M4").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

WS_In.Range("E2:E" & lro).Copy
WS_DA.Range("B4").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

WS_In.Range("F2:F" & lro).Copy
WS_DA.Range("H4").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

WS_In.Range("H2:H" & lro).Copy
WS_DA.Range("K4").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False


WS_In.Range("J2:J" & lro).Copy
WS_DA.Range("I4").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

WS_DA.Range("P4").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

WS_DA.Range(Cells(R + 4, C + 9), Cells(R + lro + 2, C + 9)).NumberFormat = "[$-14009]dd-mm-yyyy;@"
WS_DA.Range(Cells(R + 4, C + 16), Cells(R + lro + 2, C + 16)).NumberFormat = "[$-14009]dd-mm-yyyy;@"

WS_In.Range("G2:G" & lro).Copy
WS_DA.Range("J4").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
WS_DA.Range(Cells(R + 4, C + 10), Cells(R + lro + 2, C + 10)).NumberFormat = "[$-14009]dd-mm-yyyy;@"

WS_DA.Range("N4:N" & lro + 2).Value = "IQPC"
Call num_Of_Days

Range(Cells(2, 1), Cells(lro + 2, 13)).Select

With Selection
    .Columns.AutoFit
End With

For Each BI In Array(xlEdgeTop, xlEdgeLeft, xlEdgeBottom, xlEdgeRight, xlInsideHorizontal, xlInsideVertical)
    With Range(Cells(4, 1), Cells(2 + lro, 16)).Borders(BI)
         .Weight = xlThin
         .Color = RGB(148, 138, 84)
    End With
Next BI
End Sub

Sub pHERDD()
    Dim WB As Workbook
    Dim WS_In As Worksheet
    Dim WS_DA As Worksheet
    Dim R, i As Long
    Dim C As Integer
    Dim temp As String
    Dim BI As Variant
    
    Set WB = ActiveWorkbook
    Set WS_In = WB.Sheets("Incident")
    Set WS_DA = WB.Sheets("MainData")
    
    WS_In.Activate
    WS_In.Select
    
    Dim lro, lro1 As Long
    lro = WS_In.Cells(Rows.Count, "A").End(xlUp).Row
    WS_In.Range("A2:P" & lro).Sort Key1:=Range("E2"), Order1:=xlDescending
    
    'Checking for duplicate
    
    For i = 2 To lro - 1
        temp = WS_In.Cells(i, 5).Value
        If temp = WS_In.Cells(i + 1, 5).Value Then
             WS_In.Range(Cells(i, 1), Cells(i, 80)).Delete Shift:=xlUp
        End If
    Next i
    lro = WS_In.Cells(Rows.Count, "B").End(xlUp).Row

'---- Making Response SLA 'Y' @Subhankar
    WS_In.Range(WS_In.Cells(R + 2, C + 3), WS_In.Cells(R + lro, C + 3)).Formula = "Y"
    
    For i = 2 To lro
        If WS_In.Cells(i, 10).Value = "(null)" Then
            WS_In.Cells(i, 10).Value = ""
        End If
        If WS_In.Cells(i, 12).Value = "(null)" Then
            WS_In.Cells(i, 12).Value = 0
        Else
            Debug.Print WS_In.Cells(i, 12).Value
        End If
'       Resolution sla "Y" only for Resolved Ticket
        If WS_In.Cells(i, 10).Value <> "" Then
            WS_In.Cells(i, 4).Value = UCase(WS_In.Cells(i, 4).Value)
        Else
            WS_In.Cells(i, 4).Value = "N"
        End If
    Next i
    
    WS_DA.Activate
    WS_DA.Select
    
    lro1 = WS_DA.Cells(Rows.Count, "A").End(xlUp).Row
    
    If lro1 >= 4 Then
        WS_DA.Range("A4:Z" & lro1).ClearContents
    End If
    
    WS_In.Activate
    WS_In.Select
    WS_In.Range("B2:N" & lro).Copy
    
    WS_DA.Activate
    WS_DA.Select
    WS_DA.Range("B4:N" & lro).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    For i = 4 To (lro + 2)
        WS_DA.Cells(i, 1).Value = i - 3
    Next i
    
    WS_DA.Range(Cells(R + 4, C + 9), Cells(R + lro + 2, C + 9)).NumberFormat = "[$-14009]dd-mm-yyyy;@"
    WS_DA.Range(Cells(R + 4, C + 16), Cells(R + lro + 2, C + 16)).NumberFormat = "[$-14009]dd-mm-yyyy;@"
    WS_DA.Range(Cells(R + 4, C + 10), Cells(R + lro + 2, C + 10)).NumberFormat = "[$-14009]dd-mm-yyyy;@"

    WS_In.Range("I2:I" & lro).Copy
    WS_DA.Range("P4").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    WS_DA.Range(WS_DA.Cells(R + 4, C + 1), WS_DA.Cells(R + lro + 3, C + 16)).Select
    With Selection
        .Columns.AutoFit
    End With

    For Each BI In Array(xlEdgeTop, xlEdgeLeft, xlEdgeBottom, xlEdgeRight, xlInsideHorizontal, xlInsideVertical)
        With WS_DA.Range(WS_DA.Cells(R + 4, C + 1), WS_DA.Cells(R + 3 + lro, C + 16)).Borders(BI)
         .Weight = xlThin
         .Color = RGB(148, 138, 84)
        End With
    Next BI
    
End Sub

Sub pLM()

Application.DisplayAlerts = False

Dim lro As Long
Dim lro1 As Long

Dim WB As Workbook
Dim WS_In As Worksheet
Dim WS_DA As Worksheet

Dim sDa As String
Dim sIn As String
Dim BI As Variant
Dim TicketType As String

Dim i As Long

Dim R As Long
Dim C As Integer

sIn = "Incident"
sDa = "MainData"

Set WB = ActiveWorkbook
Set WS_In = WB.Sheets(sIn)
Set WS_DA = WB.Sheets(sDa)

WB.Sheets(sIn).Activate
Sheets(sIn).Select
lro = Cells(Rows.Count, "A").End(xlUp).Row

WB.Sheets(sDa).Activate
Sheets(sDa).Range("A1").Select

'Clearing the contents in the MainData Sheet Cells
lro1 = WS_DA.Cells(WS_DA.Rows.Count, "A").End(xlUp).Row
If lro1 > 4 Then
    WS_DA.Range("A4:Z" & (lro1)).ClearContents
End If

WB.Sheets(sIn).Activate
Sheets(sIn).Select

'Sorting to take out the Problem to the next sheet
WS_In.Range("A2:BZ" & lro).Sort Key1:=Range("A2"), Order1:=xlDescending

'Removing Duplicate
WS_In.Range("$A$1:$K$" & lro).RemoveDuplicates Columns:=1, Header:=xlYes

'Checking last row
lro = Cells(Rows.Count, "A").End(xlUp).Row

'--Resolution SLA Column is Added @ Subhankar
WS_In.Cells(1, 13).Value = "Resolution SLA"
For i = 2 To lro
    TicketType = WS_In.Cells(i, 11).Value

    WS_In.Cells(i, 9).Value = WS_In.Cells(i, 9).Value / 60
    
    Select Case UCase(WS_In.Cells(i, 4).Value)
        Case "BLOCKER"
            WS_In.Cells(i, 4).Value = 1
        Case "CRITICAL"
            WS_In.Cells(i, 4).Value = 2
        Case "MAJOR"
            WS_In.Cells(i, 4).Value = 3
        Case "MINOR"
            WS_In.Cells(i, 4).Value = 4
        Case "TRIVIAL"
            WS_In.Cells(i, 4).Value = 5
   End Select
   
   If Left(TicketType, 2) = "P2" Or Left(TicketType, 2) = "P3" Then
        WS_In.Cells(i, 11).Value = "INC"
            If Left(TicketType, 2) = "P2" Then
                 WS_In.Cells(i, 4).Value = 2
            Else
                 WS_In.Cells(i, 4).Value = 3
            End If
    Else
        WS_In.Cells(i, 11).Value = "SRQ"
    End If
    
'--Resolution SLA is Y for the Closed Ticket @ Subhankar
    If Left(UCase(WS_In.Cells(i, 5).Value), 6) <> "CLOSED" Then
        WS_In.Cells(i, 7).Value = ""
        WS_In.Cells(i, 13).Value = "NA"
    Else
        WS_In.Cells(i, 13).Value = "Y"
    End If
    
Next i
WB.Sheets(sIn).Activate
Sheets(sIn).Select

'Checking last row
lro = Cells(Rows.Count, "A").End(xlUp).Row

WB.Sheets(sDa).Activate
Sheets(sDa).Select
'Serial No.
For i = 4 To (lro + 2)
    WS_DA.Cells(i, 1).Value = i - 3
Next i

'SLA Response
WS_DA.Range(Cells(R + 4, C + 3), Cells(R + lro + 2, C + 3)).Formula = "Y"
'SLA Resolution
WS_In.Range("M2:M" & lro).Copy
WS_DA.Range("D4").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

'Ticket Number
WS_In.Range("A2:A" & lro).Copy
WS_DA.Range("E4").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

'Summary
WS_In.Range("B2:B" & lro).Copy
WS_DA.Range("F4").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

'Ticket type
WS_In.Range("K2:K" & lro).Copy
WS_DA.Range("B4").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

'Assignee
WS_In.Range("C2:C" & lro).Copy
WS_DA.Range("H4").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

'Priority
WS_In.Range("D2:D" & lro).Copy
WS_DA.Range("K4").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

'Effort
WS_In.Range("I2:I" & lro).Copy
WS_DA.Range("L4").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

'Start Date
WS_In.Range("F2:F" & lro).Copy
WS_DA.Range("I4").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

WS_DA.Range("P4").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

'End Date
WS_In.Range("G2:G" & lro).Copy
WS_DA.Range("J4").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

WS_DA.Range(Cells(R + 4, C + 9), Cells(R + lro + 2, C + 10)).NumberFormat = "[$-14009]dd-mm-yyyy;@"
WS_DA.Range(Cells(R + 4, C + 16), Cells(R + lro + 2, C + 16)).NumberFormat = "[$-14009]dd-mm-yyyy;@"

WS_In.Range("E2:E" & lro).Copy
WS_DA.Range("M4").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

WS_DA.Range("N4:N" & lro + 2).Value = "LM"
Call num_Of_Days

Range(Cells(2, 1), Cells(lro + 2, 13)).Select

With Selection
    .Columns.AutoFit
End With

For Each BI In Array(xlEdgeTop, xlEdgeLeft, xlEdgeBottom, xlEdgeRight, xlInsideHorizontal, xlInsideVertical)
    With Range(Cells(4, 1), Cells(2 + lro, 16)).Borders(BI)
         .Weight = xlThin
         .Color = RGB(148, 138, 84)
    End With
Next BI
End Sub

Sub pMapToDashboard()
    Call pCleanDB
    
    Dim WB As Workbook
    Dim WS_In As Worksheet
    Dim WS_DB As Worksheet
    Dim R, i As Long
    Dim C As Integer
    Dim temp As String
    Dim BI As Variant
    
    Set WB = ActiveWorkbook
    Set WS_In = WB.Sheets("Incident")
    Set WS_DB = WB.Sheets("Project or Cluster")
    
    WS_In.Activate
    WS_In.Select
    
    Dim INC_opBal_p1 As Long
    Dim INC_opBal_p2 As Long
    Dim INC_opBal_p3 As Long
    Dim INC_opBal_p4 As Long
    Dim INC_opBal_p5 As Long
    
    Dim INC_Recv_p1 As Long
    Dim INC_Recv_p2 As Long
    Dim INC_Recv_p3 As Long
    Dim INC_Recv_p4 As Long
    Dim INC_Recv_p5 As Long
    
    Dim INC_Rspnd_p1 As Long
    Dim INC_Rspnd_p2 As Long
    Dim INC_Rspnd_p3 As Long
    Dim INC_Rspnd_p4 As Long
    Dim INC_Rspnd_p5 As Long
    
    Dim INC_Rsolv_p1 As Long
    Dim INC_Rsolv_p2 As Long
    Dim INC_Rsolv_p3 As Long
    Dim INC_Rsolv_p4 As Long
    Dim INC_Rsolv_p5 As Long
    
    Dim INC_caOvr_p1 As Long
    Dim INC_caOvr_p2 As Long
    Dim INC_caOvr_p3 As Long
    Dim INC_caOvr_p4 As Long
    Dim INC_caOvr_p5 As Long
    Dim INC_OnHold_Array(4) As Long
    
    Dim INC_Queue_Array(4) As Long
    
    Dim INC_Aging_Array(4, 4) As Long
    
    Dim INC_Efrt_p1 As Double
    Dim INC_Efrt_p2 As Double
    Dim INC_Efrt_p3 As Double
    Dim INC_Efrt_p4 As Double
    Dim INC_Efrt_p5 As Double
    
    Dim INC_TeamSize As Long
    
    Dim INC_RspSLA_p1 As Long
    Dim INC_RspSLA_p2 As Long
    Dim INC_RspSLA_p3 As Long
    Dim INC_RspSLA_p4 As Long
    Dim INC_RspSLA_p5 As Long
    
    Dim INC_ResSLA_p1 As Long
    Dim INC_ResSLA_p2 As Long
    Dim INC_ResSLA_p3 As Long
    Dim INC_ResSLA_p4 As Long
    Dim INC_ResSLA_p5 As Long
    
    INC_opBal_p1 = Cells().Value
    
    WS_DB.Activate
    WS_DB.Select
    
    Cells(10, 10).Value = INC_opBal_p1
    Cells(11, 10).Value = INC_Recv_p1
    Cells(12, 10).Value = INC_Rspnd_p1
    Cells(13, 10).Value = INC_Rsolv_p1
    Cells(14, 10).Value = INC_caOvr_p1
    Cells(22, 10).Value = INC_Efrt_p1 / 60
    Cells(24, 10).Value = INC_RspSLA_p1
    Cells(25, 10).Value = INC_ResSLA_p1
    
    Cells(10, 11).Value = INC_opBal_p2
    Cells(11, 11).Value = INC_Recv_p2
    Cells(12, 11).Value = INC_Rspnd_p2
    Cells(13, 11).Value = INC_Rsolv_p2
    Cells(14, 11).Value = INC_caOvr_p2
    Cells(22, 11).Value = INC_Efrt_p2 / 60
    Cells(24, 11).Value = INC_RspSLA_p2
    Cells(25, 11).Value = INC_ResSLA_p2
    
    Cells(10, 12).Value = INC_opBal_p3
    Cells(11, 12).Value = INC_Recv_p3
    Cells(12, 12).Value = INC_Rspnd_p3
    Cells(13, 12).Value = INC_Rsolv_p3
    Cells(14, 12).Value = INC_caOvr_p3
    Cells(22, 12).Value = INC_Efrt_p3 / 60
    Cells(24, 12).Value = INC_RspSLA_p3
    Cells(25, 12).Value = INC_ResSLA_p3
    
    Cells(10, 13).Value = INC_opBal_p4
    Cells(11, 13).Value = INC_Recv_p4
    Cells(12, 13).Value = INC_Rspnd_p4
    Cells(13, 13).Value = INC_Rsolv_p4
    Cells(14, 13).Value = INC_caOvr_p4
    Cells(22, 13).Value = (INC_Efrt_p4 / 60)
    Cells(24, 13).Value = INC_RspSLA_p4
    Cells(25, 13).Value = INC_ResSLA_p4
    
    Cells(10, 14).Value = INC_opBal_p5
    Cells(11, 14).Value = INC_Recv_p5
    Cells(12, 14).Value = INC_Rspnd_p5
    Cells(13, 14).Value = INC_Rsolv_p5
    Cells(14, 14).Value = INC_caOvr_p5
    Cells(22, 14).Value = INC_Efrt_p5 / 60
    Cells(24, 14).Value = INC_RspSLA_p5
    Cells(25, 14).Value = INC_ResSLA_p5
    
    Range("J15:N15").Value = INC_OnHold_Array
    
    Range("J16:N16").Value = INC_Queue_Array
    
    Range("J17:N21").Value = INC_Aging_Array
    
    Cells(23, 10).Value = INC_TeamSize
End Sub
