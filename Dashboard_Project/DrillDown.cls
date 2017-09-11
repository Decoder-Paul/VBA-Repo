VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    Dim client As String
    client = Cells(10, 5).Value
    Dim today As String
    today = Format(Date - 1, "mm-dd-yyyy")
    Sheets("MainData").Visible = True
    Sheets("MainData").Select
    Sheets("MainData").Range("A3:P3").Select
    Selection.AutoFilter
    If client <> "ALL" And client <> "All" Then
        Sheets("MainData").Select
        Sheets("MainData").Range("A3:P3").Select
        With Selection
            .AutoFilter Field:=14, Criteria1:=client
        End With
    End If
    Sheets("MainData").Select
    Sheets("MainData").Range("A3:P3").Select
    R = Target.Row
    C = Target.Column
    With Selection
    Select Case R
        Case 10 'Opening Balance
            .AutoFilter Field:=9, Criteria1:="<" & today
        Case 11 'Received
            .AutoFilter Field:=9, Criteria1:=">=" & today
        Case 12 'Responded
            .AutoFilter Field:=9, Criteria1:=">=" & today
            .AutoFilter Field:=3, Criteria1:="=N", Operator:=xlOr, Criteria2:="=Y"
        Case 13 'Resolved
            .AutoFilter Field:=10, Criteria1:="<>"
        Case 14 'Carry Forward
            .AutoFilter Field:=10, Criteria1:="="
        Case 15 'SLA Hold
            .AutoFilter Field:=13, Criteria1:=Array( _
            "PENDING", "SLAHOLD", "Waiting for Problem", "Waiting for User", _
            "Waiting for Vendor"), Operator:=xlFilterValues
        Case 16 'Queued
            .AutoFilter Field:=15, Criteria1:="="
        Case 18 '1-15D
            .AutoFilter Field:=15, Criteria1:=">=1", Operator:=xlAnd, Criteria2:="<=15"
        Case 19 '16-30D
            .AutoFilter Field:=15, Criteria1:=">=16", Operator:=xlAnd, Criteria2:="<=30"
        Case 20 '31-45D
            .AutoFilter Field:=15, Criteria1:=">=31", Operator:=xlAnd, Criteria2:="<=45"
        Case 21 '46-60D
            .AutoFilter Field:=15, Criteria1:=">=46", Operator:=xlAnd, Criteria2:="<=60"
        Case 22 '>60D
            .AutoFilter Field:=15, Criteria1:=">60"
        Case Else
            If client = "All" Then
                Sheets("Project or Cluster").Select
            Else
                Sheets(client).Select
            End If
    End Select
    Select Case C
        Case 15 'INC
            .AutoFilter Field:=2, Criteria1:="INC"
        Case 22 'SRQ
            .AutoFilter Field:=2, Criteria1:="SRQ"
        Case 29 'CHG
            .AutoFilter Field:=2, Criteria1:="CHG"
        Case 36 'PRB
            .AutoFilter Field:=2, Criteria1:="PRB"
        Case 10
            .AutoFilter Field:=2, Criteria1:="INC"
            .AutoFilter Field:=11, Criteria1:="1"
        Case 11
            .AutoFilter Field:=2, Criteria1:="INC"
            .AutoFilter Field:=11, Criteria1:="2"
        Case 12
            .AutoFilter Field:=2, Criteria1:="INC"
            .AutoFilter Field:=11, Criteria1:="3"
        Case 13
            .AutoFilter Field:=2, Criteria1:="INC"
            .AutoFilter Field:=11, Criteria1:="4"
        Case 14
            .AutoFilter Field:=2, Criteria1:="INC"
            .AutoFilter Field:=11, Criteria1:="5"
        Case 17
            .AutoFilter Field:=2, Criteria1:="SRQ"
            .AutoFilter Field:=11, Criteria1:="1"
        Case 18
            .AutoFilter Field:=2, Criteria1:="SRQ"
            .AutoFilter Field:=11, Criteria1:="2"
        Case 19
            .AutoFilter Field:=2, Criteria1:="SRQ"
            .AutoFilter Field:=11, Criteria1:="3"
        Case 20
            .AutoFilter Field:=2, Criteria1:="SRQ"
            .AutoFilter Field:=11, Criteria1:="4"
        Case 21
            .AutoFilter Field:=2, Criteria1:="SRQ"
            .AutoFilter Field:=11, Criteria1:="5"
        Case 24
            .AutoFilter Field:=2, Criteria1:="CHG"
            .AutoFilter Field:=11, Criteria1:="1"
        Case 25
            .AutoFilter Field:=2, Criteria1:="CHG"
            .AutoFilter Field:=11, Criteria1:="2"
        Case 26
            .AutoFilter Field:=2, Criteria1:="CHG"
            .AutoFilter Field:=11, Criteria1:="3"
        Case 27
            .AutoFilter Field:=2, Criteria1:="CHG"
            .AutoFilter Field:=11, Criteria1:="4"
        Case 28
            .AutoFilter Field:=2, Criteria1:="CHG"
            .AutoFilter Field:=11, Criteria1:="5"
        Case 31
            .AutoFilter Field:=2, Criteria1:="PRB"
            .AutoFilter Field:=11, Criteria1:="1"
        Case 32
            .AutoFilter Field:=2, Criteria1:="PRB"
            .AutoFilter Field:=11, Criteria1:="2"
        Case 33
            .AutoFilter Field:=2, Criteria1:="PRB"
            .AutoFilter Field:=11, Criteria1:="3"
        Case 34
            .AutoFilter Field:=2, Criteria1:="PRB"
            .AutoFilter Field:=11, Criteria1:="4"
        Case 35
            .AutoFilter Field:=2, Criteria1:="PRB"
            .AutoFilter Field:=11, Criteria1:="5"
        Case Else
            ActiveSheet.Select
            Exit Sub
    End Select
    
    End With
End Sub
