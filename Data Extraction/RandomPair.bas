Attribute VB_Name = "RandomPair"
Sub RandomPair()
'========================================================================================================
' RandomPair
' -------------------------------------------------------------------------------------------------------
' Purpose   :   To create pair of teams randomly from a list of participants
'
' Author    :   Subhankar Paul
' Date      :   9th November, 2017
' Notes     :
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
    Dim rng As Range, cel As Range
    lb = 2
    lro = Sheets("Team Group").Cells(Rows.Count, "A").End(xlUp).Row
    l = Round(lro / 2, 0)
    If l Mod 2 <> 0 Then
        MsgBox "No. of participants is ODD", vbExclamation
        MsgBox "Please enter Even no. of Participants", vbInformation
        Exit Sub
    End If
    Set rng = Range("A2:A" & lro)
    teams = rng
    ub = UBound(teams)
    
    For i = 1 To l
        Cells(i + 1, 4).Value = teams(i, 1)
       ind = WorksheetFunction.RandBetween(lb, ub)
        Cells(i + 1, 5).Value = teams(ind, 1)
        teams(ind, 1) = "0"
        For j = ind To ub - 1
            teams(j, 1) = teams(j + 1, 1)
        Next j
        teams(ub, 1) = "0"
        lb = lb + 1
        ub = ub - 1
    Next i
End Sub

