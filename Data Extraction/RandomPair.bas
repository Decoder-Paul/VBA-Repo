Attribute VB_Name = "RandomPair"
Sub pCleanSingleDB()
    Dim lro As Long
    lro = ActiveSheet.Cells(Rows.count, "F").End(xlUp).Row
    If lro > 3 Then
        ActiveSheet.Range("F4:G" & lro).ClearContents
    End If
End Sub
Sub pCleanDoubleDB()
    Dim lro As Long
    lro = ActiveSheet.Cells(Rows.count, "L").End(xlUp).Row
    If lro > 3 Then
        ActiveSheet.Range("L4:M" & lro).ClearContents
    End If
End Sub
Sub pCleanWomenSingleDB()
    Dim lro As Long
    lro = ActiveSheet.Cells(Rows.count, "I").End(xlUp).Row
    If lro > 3 Then
        ActiveSheet.Range("I4:J" & lro).ClearContents
    End If
End Sub
Sub SinglesPair()
'========================================================================================================
' RandomPair
' -------------------------------------------------------------------------------------------------------
' Purpose   :   To create pair of teams randomly from a list of participants
'
' Author    :   Subhankar Paul
' Date      :   9th November, 2017
' Notes     :   List
' Parameters:   N/A
' Returns   :   N/A
' ---------------------------------------------------------------
' Revision History
'
'========================================================================================================
    Dim wb As Workbook
    Dim filepath As String
    Dim i As Integer
    Dim j As Integer
    Dim lb As Integer
    Dim ub As Integer
    Dim lro As Integer
    Dim l As Integer
    Dim ind As Integer
    Dim teams() As Variant
    Dim randomIndex() As Long
    Dim rng As Range, cel As Range
    Dim pa As Integer
    Dim sname As String
    Static cnt As Long
    cnt = cnt + 1
    sname = cnt & ActiveSheet.Name & "Singles"
    filepath = Application.ActiveWorkbook.Path
    lro = ActiveSheet.Cells(Rows.count, "B").End(xlUp).Row
    If lro < 2 Then
        MsgBox "Please Enter Participants", vbExclamation
        Exit Sub
    End If
    lb = 1
    pa = lro - 1
    l = Round(pa / 2, 0)
    If pa Mod 2 <> 0 Then
        MsgBox "No. of participants is ODD", vbExclamation
        MsgBox "Please enter One Lucky You Participant :)", vbInformation
        Exit Sub
    End If
    'range for single participants
    Call pCleanSingleDB
    Set rng = Range("B2:B" & lro)
    teams = rng
    ReDim randomIndex(lro)
    ub = UBound(teams)
    For i = 1 To ub
loo:    ind = WorksheetFunction.RandBetween(lb, ub)
        If IsError(Application.Match(ind, randomIndex, False)) Then
            randomIndex(i) = ind
            Cells(i + 1, 50).Value = teams(ind, 1)
        Else
            GoTo loo
        End If
    Next i
    'teams updated with randomized set of participants
    Set rng = Range("AX2:AX" & lro)
    teams = rng
    lb = 2
    For i = 1 To l
        Cells(i + 3, 6).Value = teams(i, 1)
        ind = WorksheetFunction.RandBetween(lb, ub)
        Cells(i + 3, 7).Value = teams(ind, 1)
        teams(ind, 1) = "0"
        For j = ind To ub - 1
            teams(j, 1) = teams(j + 1, 1)
        Next j
        teams(ub, 1) = "0"
        lb = lb + 1
        ub = ub - 1
    Next i
    Set wb = Workbooks.Add
    ThisWorkbook.Activate
    ActiveSheet.Copy Before:=wb.Sheets(1)
    wb.Activate
    wb.SaveAs filepath & "\" & sname & ".xlsx"
    wb.Close
End Sub
Sub DoublesPair()
'========================================================================================================
' RandomPair
' -------------------------------------------------------------------------------------------------------
' Purpose   :   To create pair of teams randomly from a list of participants
'
' Author    :   Subhankar Paul
' Date      :   9th November, 2017
' Notes     :   List
' Parameters:   N/A
' Returns   :   N/A
' ---------------------------------------------------------------
' Revision History
'
'========================================================================================================
    Dim i As Integer
    Dim j As Integer
    Dim lb As Integer
    Dim ub As Integer
    Dim lro As Integer
    Dim l As Integer
    Dim ind As Integer
    Dim teams() As Variant
    Dim randomIndex() As Long
    Dim rng As Range, cel As Range
    Dim pa As Integer
    
    Dim wb As Workbook
    Dim filepath As String
    Dim sname As String
    Static cunt As Long
    cunt = cunt + 1
    sname = cunt & ActiveSheet.Name & "Doubles"
    filepath = Application.ActiveWorkbook.Path
    
    lro = ActiveSheet.Cells(Rows.count, "D").End(xlUp).Row
    If lro < 2 Then
        MsgBox "Please Enter Participants", vbExclamation
        Exit Sub
    End If
    lb = 1
    pa = lro - 1
    l = Round(pa / 2, 0)
    If pa Mod 2 <> 0 Then
        MsgBox "No. of participants is ODD", vbExclamation
        MsgBox "Please enter One Lucky You Participant :)", vbInformation
        Exit Sub
    End If
    Call pCleanDoubleDB
    'range for single participants
    Set rng = Range("D2:D" & lro)
    teams = rng
    ReDim randomIndex(lro)
    ub = UBound(teams)
    For i = 1 To ub
loo:    ind = WorksheetFunction.RandBetween(lb, ub)
        If IsError(Application.Match(ind, randomIndex, False)) Then
            randomIndex(i) = ind
            Cells(i + 1, 52).Value = teams(ind, 1)
        Else
            GoTo loo
        End If
    Next i
    'teams updated with randomized set of participants
    Set rng = Range("AZ2:AZ" & lro)
    teams = rng
    lb = 2
    For i = 1 To l
        Cells(i + 3, 12).Value = teams(i, 1)
        ind = WorksheetFunction.RandBetween(lb, ub)
        Cells(i + 3, 13).Value = teams(ind, 1)
        teams(ind, 1) = "0"
        For j = ind To ub - 1
            teams(j, 1) = teams(j + 1, 1)
        Next j
        teams(ub, 1) = "0"
        lb = lb + 1
        ub = ub - 1
    Next i
    
    Set wb = Workbooks.Add
    ThisWorkbook.Activate
    ActiveSheet.Copy Before:=wb.Sheets(1)
    wb.Activate
    wb.SaveAs filepath & "\" & sname & ".xlsx"
    wb.Close
End Sub

Sub WomanSinglesPair()
'========================================================================================================
' RandomPair
' -------------------------------------------------------------------------------------------------------
' Purpose   :   To create pair of teams randomly from a list of participants
'
' Author    :   Subhankar Paul
' Date      :   15th November, 2017
' Notes     :   List
' Parameters:   N/A
' Returns   :   N/A
' ---------------------------------------------------------------
' Revision History
'
'========================================================================================================
    Dim i As Integer
    Dim j As Integer
    Dim lb As Integer
    Dim ub As Integer
    Dim lro As Integer
    Dim l As Integer
    Dim ind As Integer
    Dim teams() As Variant
    Dim randomIndex() As Long
    Dim rng As Range, cel As Range
    Dim pa As Integer
    
    Dim wb As Workbook
    Dim filepath As String
    Dim sname As String
    Static coun As Long
    coun = coun + 1
    sname = coun & ActiveSheet.Name & "Women"
    filepath = Application.ActiveWorkbook.Path
    
    lro = ActiveSheet.Cells(Rows.count, "C").End(xlUp).Row
    If lro < 2 Then
        MsgBox "Please Enter Participants", vbExclamation
        Exit Sub
    End If
    lb = 1
    pa = lro - 1
    l = Round(pa / 2, 0)
    If pa Mod 2 <> 0 Then
        MsgBox "No. of participants is ODD", vbExclamation
        MsgBox "Please enter One Lucky You Participant :)", vbInformation
        Exit Sub
    End If
    'range for single participants
    Call pCleanWomenSingleDB
    Set rng = Range("C2:C" & lro)
    teams = rng
    ReDim randomIndex(lro)
    ub = UBound(teams)
    For i = 1 To ub
loo:    ind = WorksheetFunction.RandBetween(lb, ub)
        If IsError(Application.Match(ind, randomIndex, False)) Then
            randomIndex(i) = ind
            Cells(i + 1, 50).Value = teams(ind, 1)
        Else
            GoTo loo
        End If
    Next i
    'teams updated with randomized set of participants
    Set rng = Range("AX2:AX" & lro)
    teams = rng
    lb = 2
    For i = 1 To l
        Cells(i + 3, 9).Value = teams(i, 1)
        ind = WorksheetFunction.RandBetween(lb, ub)
        Cells(i + 3, 10).Value = teams(ind, 1)
        teams(ind, 1) = "0"
        For j = ind To ub - 1
            teams(j, 1) = teams(j + 1, 1)
        Next j
        teams(ub, 1) = "0"
        lb = lb + 1
        ub = ub - 1
    Next i
    
    Set wb = Workbooks.Add
    ThisWorkbook.Activate
    ActiveSheet.Copy Before:=wb.Sheets(1)
    wb.Activate
    wb.SaveAs filepath & "\" & sname & ".xlsx"
    wb.Close
    
End Sub
