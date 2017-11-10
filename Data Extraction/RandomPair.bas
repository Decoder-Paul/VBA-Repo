Attribute VB_Name = "RandomPair"
Sub pCleanSingleDB()
    Dim lro As Long
    lro = ActiveSheet.Cells(Rows.Count, "E").End(xlUp).Row
    If lro > 1 Then
        ActiveSheet.Range("E5:F" & lro).ClearContents
    End If
End Sub
Sub pCleanDoubleDB()
    Dim lro As Long
    lro = ActiveSheet.Cells(Rows.Count, "H").End(xlUp).Row
    If lro > 1 Then
        ActiveSheet.Range("H5:I" & lro).ClearContents
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
    lro = Sheets("Team Group").Cells(Rows.Count, "A").End(xlUp).Row
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
    Set rng = Range("A2:A" & lro)
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
        Cells(i + 4, 5).Value = teams(i, 1)
        ind = WorksheetFunction.RandBetween(lb, ub)
        Cells(i + 4, 6).Value = teams(ind, 1)
        teams(ind, 1) = "0"
        For j = ind To ub - 1
            teams(j, 1) = teams(j + 1, 1)
        Next j
        teams(ub, 1) = "0"
        lb = lb + 1
        ub = ub - 1
    Next i
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
    lro = Sheets("Team Group").Cells(Rows.Count, "C").End(xlUp).Row
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
    Set rng = Range("C2:C" & lro)
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
        Cells(i + 4, 8).Value = teams(i, 1)
        ind = WorksheetFunction.RandBetween(lb, ub)
        Cells(i + 4, 9).Value = teams(ind, 1)
        teams(ind, 1) = "0"
        For j = ind To ub - 1
            teams(j, 1) = teams(j + 1, 1)
        Next j
        teams(ub, 1) = "0"
        lb = lb + 1
        ub = ub - 1
    Next i
End Sub
