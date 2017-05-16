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
    Call incidentFilter(Target, today)
'------------------------------- END OF INCIDENT TICKET ------------------------------
'---------------- Service Request Ticket -------------------
'--------------------------- Priority 1
    'Opening Balance Ticket
    If Not Intersect(Target, Range("P10")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="SRQ"
            .AutoFilter Field:=11, Criteria1:="1"
            .AutoFilter Field:=9, Criteria1:="<" & today
        End With
    End If
    'Received Ticket
    If Not Intersect(Target, Range("P11")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="SRQ"
            .AutoFilter Field:=11, Criteria1:="1"
            .AutoFilter Field:=9, Criteria1:=">=" & today
        End With
    End If
    'Responded Ticket
    If Not Intersect(Target, Range("P12")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="SRQ"
            .AutoFilter Field:=11, Criteria1:="1"
            .AutoFilter Field:=9, Criteria1:=">=" & today
            .AutoFilter Field:=3, Criteria1:="=N", Operator:=xlOr, Criteria2:="=Y"
        End With
    End If
    'Resolved Ticket
    If Not Intersect(Target, Range("P13")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="SRQ"
            .AutoFilter Field:=11, Criteria1:="1"
            .AutoFilter Field:=10, Criteria1:="<>"
        End With
    End If
    'Carried Over Ticket
    If Not Intersect(Target, Range("P14")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="SRQ"
            .AutoFilter Field:=11, Criteria1:="1"
            .AutoFilter Field:=10, Criteria1:="="
        End With
    End If
    'SLA Hold Ticket
    If Not Intersect(Target, Range("P15")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="SRQ"
            .AutoFilter Field:=11, Criteria1:="1"
            .AutoFilter Field:=13, Criteria1:=Array( _
            "PENDING", "SLAHOLD", "Waiting for Problem", "Waiting for User", _
            "Waiting for Vendor"), Operator:=xlFilterValues
        End With
    End If
    'Queued Ticket
    If Not Intersect(Target, Range("P16")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="SRQ"
            .AutoFilter Field:=11, Criteria1:="1"
            .AutoFilter Field:=15, Criteria1:="="
        End With
    End If
    '1-15 D Ticket
    If Not Intersect(Target, Range("P17")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="SRQ"
            .AutoFilter Field:=11, Criteria1:="1"
            .AutoFilter Field:=15, Criteria1:=">=1", Operator:=xlAnd, Criteria2:="<=15"
        End With
    End If
    '16-30 D Ticket
    If Not Intersect(Target, Range("P18")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="SRQ"
            .AutoFilter Field:=11, Criteria1:="1"
            .AutoFilter Field:=15, Criteria1:=">=16", Operator:=xlAnd, Criteria2:="<=30"
        End With
    End If
    '31-45 D Ticket
    If Not Intersect(Target, Range("P19")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="SRQ"
            .AutoFilter Field:=11, Criteria1:="1"
            .AutoFilter Field:=15, Criteria1:=">=31", Operator:=xlAnd, Criteria2:="<=45"
        End With
    End If
    '46-60 D Ticket
    If Not Intersect(Target, Range("P20")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="SRQ"
            .AutoFilter Field:=11, Criteria1:="1"
            .AutoFilter Field:=15, Criteria1:=">=46", Operator:=xlAnd, Criteria2:="<=60"
        End With
    End If
    ' >60 D Ticket
    If Not Intersect(Target, Range("P21")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="SRQ"
            .AutoFilter Field:=11, Criteria1:="1"
            .AutoFilter Field:=15, Criteria1:=">60"
        End With
    End If
'--------------------------- Priority 2
    'Opening Balance Ticket
    If Not Intersect(Target, Range("Q10")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="SRQ"
            .AutoFilter Field:=11, Criteria1:="2"
            .AutoFilter Field:=9, Criteria1:="<" & today
        End With
    End If
    'Received Ticket
    If Not Intersect(Target, Range("Q11")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="SRQ"
            .AutoFilter Field:=11, Criteria1:="2"
            .AutoFilter Field:=9, Criteria1:=">=" & today
        End With
    End If
    'Responded Ticket
    If Not Intersect(Target, Range("Q12")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="SRQ"
            .AutoFilter Field:=11, Criteria1:="2"
            .AutoFilter Field:=9, Criteria1:=">=" & today
            .AutoFilter Field:=3, Criteria1:="=N", Operator:=xlOr, Criteria2:="=Y"
        End With
    End If
    'Resolved Ticket
    If Not Intersect(Target, Range("Q13")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="SRQ"
            .AutoFilter Field:=11, Criteria1:="2"
            .AutoFilter Field:=10, Criteria1:="<>"
        End With
    End If
    'Carried Over Ticket
    If Not Intersect(Target, Range("Q14")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="SRQ"
            .AutoFilter Field:=11, Criteria1:="2"
            .AutoFilter Field:=10, Criteria1:="="
        End With
    End If
    'SLA Hold Ticket
    If Not Intersect(Target, Range("Q15")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="SRQ"
            .AutoFilter Field:=11, Criteria1:="2"
            .AutoFilter Field:=13, Criteria1:=Array( _
            "PENDING", "SLAHOLD", "Waiting for Problem", "Waiting for User", _
            "Waiting for Vendor"), Operator:=xlFilterValues
        End With
    End If
    'Queued Ticket
    If Not Intersect(Target, Range("Q16")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="SRQ"
            .AutoFilter Field:=11, Criteria1:="2"
            .AutoFilter Field:=15, Criteria1:="="
        End With
    End If
    '1-15 D Ticket
    If Not Intersect(Target, Range("Q17")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="SRQ"
            .AutoFilter Field:=11, Criteria1:="2"
            .AutoFilter Field:=15, Criteria1:=">=1", Operator:=xlAnd, Criteria2:="<=15"
        End With
    End If
    '16-30 D Ticket
    If Not Intersect(Target, Range("Q18")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="SRQ"
            .AutoFilter Field:=11, Criteria1:="2"
            .AutoFilter Field:=15, Criteria1:=">=16", Operator:=xlAnd, Criteria2:="<=30"
        End With
    End If
    '31-45 D Ticket
    If Not Intersect(Target, Range("Q19")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="SRQ"
            .AutoFilter Field:=11, Criteria1:="2"
            .AutoFilter Field:=15, Criteria1:=">=31", Operator:=xlAnd, Criteria2:="<=45"
        End With
    End If
    '46-60 D Ticket
    If Not Intersect(Target, Range("Q20")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="SRQ"
            .AutoFilter Field:=11, Criteria1:="2"
            .AutoFilter Field:=15, Criteria1:=">=46", Operator:=xlAnd, Criteria2:="<=60"
        End With
    End If
    ' >60 D Ticket
    If Not Intersect(Target, Range("Q21")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="SRQ"
            .AutoFilter Field:=11, Criteria1:="2"
            .AutoFilter Field:=15, Criteria1:=">60"
        End With
    End If
'--------------------------- Priority 3
    'Opening Balance Ticket
    If Not Intersect(Target, Range("R10")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="SRQ"
            .AutoFilter Field:=11, Criteria1:="3"
            .AutoFilter Field:=9, Criteria1:="<" & today
        End With
    End If
    'Received Ticket
    If Not Intersect(Target, Range("R11")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="SRQ"
            .AutoFilter Field:=11, Criteria1:="3"
            .AutoFilter Field:=9, Criteria1:=">=" & today
        End With
    End If
    'Responded Ticket
    If Not Intersect(Target, Range("R12")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="SRQ"
            .AutoFilter Field:=11, Criteria1:="3"
            .AutoFilter Field:=9, Criteria1:=">=" & today
            .AutoFilter Field:=3, Criteria1:="=N", Operator:=xlOr, Criteria2:="=Y"
        End With
    End If
    'Resolved Ticket
    If Not Intersect(Target, Range("R13")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="SRQ"
            .AutoFilter Field:=11, Criteria1:="3"
            .AutoFilter Field:=10, Criteria1:="<>"
        End With
    End If
    'Carried Over Ticket
    If Not Intersect(Target, Range("R14")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="SRQ"
            .AutoFilter Field:=11, Criteria1:="3"
            .AutoFilter Field:=10, Criteria1:="="
        End With
    End If
    'SLA Hold Ticket
    If Not Intersect(Target, Range("R15")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="SRQ"
            .AutoFilter Field:=11, Criteria1:="3"
            .AutoFilter Field:=13, Criteria1:=Array( _
            "PENDING", "SLAHOLD", "Waiting for Problem", "Waiting for User", _
            "Waiting for Vendor"), Operator:=xlFilterValues
        End With
    End If
    'Queued Ticket
    If Not Intersect(Target, Range("R16")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="SRQ"
            .AutoFilter Field:=11, Criteria1:="3"
            .AutoFilter Field:=15, Criteria1:="="
        End With
    End If
    '1-15 D Ticket
    If Not Intersect(Target, Range("R17")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="SRQ"
            .AutoFilter Field:=11, Criteria1:="3"
            .AutoFilter Field:=15, Criteria1:=">=1", Operator:=xlAnd, Criteria2:="<=15"
        End With
    End If
    '16-30 D Ticket
    If Not Intersect(Target, Range("R18")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="SRQ"
            .AutoFilter Field:=11, Criteria1:="3"
            .AutoFilter Field:=15, Criteria1:=">=16", Operator:=xlAnd, Criteria2:="<=30"
        End With
    End If
    '31-45 D Ticket
    If Not Intersect(Target, Range("R19")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="SRQ"
            .AutoFilter Field:=11, Criteria1:="3"
            .AutoFilter Field:=15, Criteria1:=">=31", Operator:=xlAnd, Criteria2:="<=45"
        End With
    End If
    '46-60 D Ticket
    If Not Intersect(Target, Range("R20")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="SRQ"
            .AutoFilter Field:=11, Criteria1:="3"
            .AutoFilter Field:=15, Criteria1:=">=46", Operator:=xlAnd, Criteria2:="<=60"
        End With
    End If
    ' >60 D Ticket
    If Not Intersect(Target, Range("R21")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="SRQ"
            .AutoFilter Field:=11, Criteria1:="3"
            .AutoFilter Field:=15, Criteria1:=">60"
        End With
    End If
'--------------------------- Priority 4
    'Opening Balance Ticket
    If Not Intersect(Target, Range("S10")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="SRQ"
            .AutoFilter Field:=11, Criteria1:="4"
            .AutoFilter Field:=9, Criteria1:="<" & today
        End With
    End If
    'Received Ticket
    If Not Intersect(Target, Range("S11")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="SRQ"
            .AutoFilter Field:=11, Criteria1:="4"
            .AutoFilter Field:=9, Criteria1:=">=" & today
        End With
    End If
    'Responded Ticket
    If Not Intersect(Target, Range("S12")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="SRQ"
            .AutoFilter Field:=11, Criteria1:="4"
            .AutoFilter Field:=9, Criteria1:=">=" & today
            .AutoFilter Field:=3, Criteria1:="=N", Operator:=xlOr, Criteria2:="=Y"
        End With
    End If
    'Resolved Ticket
    If Not Intersect(Target, Range("S13")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="SRQ"
            .AutoFilter Field:=11, Criteria1:="4"
            .AutoFilter Field:=10, Criteria1:="<>"
        End With
    End If
    'Carried Over Ticket
    If Not Intersect(Target, Range("S14")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="SRQ"
            .AutoFilter Field:=11, Criteria1:="4"
            .AutoFilter Field:=10, Criteria1:="="
        End With
    End If
    'SLA Hold Ticket
    If Not Intersect(Target, Range("S15")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="SRQ"
            .AutoFilter Field:=11, Criteria1:="4"
            .AutoFilter Field:=13, Criteria1:=Array( _
            "PENDING", "SLAHOLD", "Waiting for Problem", "Waiting for User", _
            "Waiting for Vendor"), Operator:=xlFilterValues
        End With
    End If
    'Queued Ticket
    If Not Intersect(Target, Range("S16")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="SRQ"
            .AutoFilter Field:=11, Criteria1:="4"
            .AutoFilter Field:=15, Criteria1:="="
        End With
    End If
    '1-15 D Ticket
    If Not Intersect(Target, Range("S17")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="SRQ"
            .AutoFilter Field:=11, Criteria1:="4"
            .AutoFilter Field:=15, Criteria1:=">=1", Operator:=xlAnd, Criteria2:="<=15"
        End With
    End If
    '16-30 D Ticket
    If Not Intersect(Target, Range("S18")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="SRQ"
            .AutoFilter Field:=11, Criteria1:="4"
            .AutoFilter Field:=15, Criteria1:=">=16", Operator:=xlAnd, Criteria2:="<=30"
        End With
    End If
    '31-45 D Ticket
    If Not Intersect(Target, Range("S19")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="SRQ"
            .AutoFilter Field:=11, Criteria1:="4"
            .AutoFilter Field:=15, Criteria1:=">=31", Operator:=xlAnd, Criteria2:="<=45"
        End With
    End If
    '46-60 D Ticket
    If Not Intersect(Target, Range("S20")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="SRQ"
            .AutoFilter Field:=11, Criteria1:="4"
            .AutoFilter Field:=15, Criteria1:=">=46", Operator:=xlAnd, Criteria2:="<=60"
        End With
    End If
    ' >60 D Ticket
    If Not Intersect(Target, Range("S21")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="SRQ"
            .AutoFilter Field:=11, Criteria1:="4"
            .AutoFilter Field:=15, Criteria1:=">60"
        End With
    End If
    
'--------------------------- Priority 5
    'Opening Balance Ticket
    If Not Intersect(Target, Range("T10")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="SRQ"
            .AutoFilter Field:=11, Criteria1:="5"
            .AutoFilter Field:=9, Criteria1:="<" & today
        End With
    End If
    'Received Ticket
    If Not Intersect(Target, Range("T11")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="SRQ"
            .AutoFilter Field:=11, Criteria1:="5"
            .AutoFilter Field:=9, Criteria1:=">=" & today
        End With
    End If
    'Responded Ticket
    If Not Intersect(Target, Range("T12")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="SRQ"
            .AutoFilter Field:=11, Criteria1:="5"
            .AutoFilter Field:=9, Criteria1:=">=" & today
            .AutoFilter Field:=3, Criteria1:="=N", Operator:=xlOr, Criteria2:="=Y"
        End With
    End If
    'Resolved Ticket
    If Not Intersect(Target, Range("T13")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="SRQ"
            .AutoFilter Field:=11, Criteria1:="5"
            .AutoFilter Field:=10, Criteria1:="<>"
        End With
    End If
    'Carried Over Ticket
    If Not Intersect(Target, Range("T14")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="SRQ"
            .AutoFilter Field:=11, Criteria1:="5"
            .AutoFilter Field:=10, Criteria1:="="
        End With
    End If
    'SLA Hold Ticket
    If Not Intersect(Target, Range("T15")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="SRQ"
            .AutoFilter Field:=11, Criteria1:="5"
            .AutoFilter Field:=13, Criteria1:=Array( _
            "PENDING", "SLAHOLD", "Waiting for Problem", "Waiting for User", _
            "Waiting for Vendor"), Operator:=xlFilterValues
        End With
    End If
    'Queued Ticket
    If Not Intersect(Target, Range("T16")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="SRQ"
            .AutoFilter Field:=11, Criteria1:="5"
            .AutoFilter Field:=15, Criteria1:="="
        End With
    End If
    '1-15 D Ticket
    If Not Intersect(Target, Range("T17")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="SRQ"
            .AutoFilter Field:=11, Criteria1:="5"
            .AutoFilter Field:=15, Criteria1:=">=1", Operator:=xlAnd, Criteria2:="<=15"
        End With
    End If
    '16-30 D Ticket
    If Not Intersect(Target, Range("T18")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="SRQ"
            .AutoFilter Field:=11, Criteria1:="5"
            .AutoFilter Field:=15, Criteria1:=">=16", Operator:=xlAnd, Criteria2:="<=30"
        End With
    End If
    '31-45 D Ticket
    If Not Intersect(Target, Range("T19")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="SRQ"
            .AutoFilter Field:=11, Criteria1:="5"
            .AutoFilter Field:=15, Criteria1:=">=31", Operator:=xlAnd, Criteria2:="<=45"
        End With
    End If
    '46-60 D Ticket
    If Not Intersect(Target, Range("T20")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="SRQ"
            .AutoFilter Field:=11, Criteria1:="5"
            .AutoFilter Field:=15, Criteria1:=">=46", Operator:=xlAnd, Criteria2:="<=60"
        End With
    End If
    ' >60 D Ticket
    If Not Intersect(Target, Range("T21")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="SRQ"
            .AutoFilter Field:=11, Criteria1:="5"
            .AutoFilter Field:=15, Criteria1:=">60"
        End With
    End If
'---------------------------- END OF SERVICE REQUEST TICKET ------------------------------
'---------------- Change Request Ticket -------------------------
'--------------------------- Priority 1
    'Opening Balance Ticket
    If Not Intersect(Target, Range("V10")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="CHG"
            .AutoFilter Field:=11, Criteria1:="1"
            .AutoFilter Field:=9, Criteria1:="<" & today
        End With
    End If
    'Received Ticket
    If Not Intersect(Target, Range("V11")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="CHG"
            .AutoFilter Field:=11, Criteria1:="1"
            .AutoFilter Field:=9, Criteria1:=">=" & today
        End With
    End If
    'Resolved Ticket
    If Not Intersect(Target, Range("V13")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="CHG"
            .AutoFilter Field:=11, Criteria1:="1"
            .AutoFilter Field:=10, Criteria1:="<>"
        End With
    End If
    'Carried Over Ticket
    If Not Intersect(Target, Range("V14")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="CHG"
            .AutoFilter Field:=11, Criteria1:="1"
            .AutoFilter Field:=10, Criteria1:="="
        End With
    End If
    'SLA Hold Ticket
    If Not Intersect(Target, Range("V15")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="CHG"
            .AutoFilter Field:=11, Criteria1:="1"
            .AutoFilter Field:=13, Criteria1:=Array( _
            "PENDING", "SLAHOLD", "Waiting for Problem", "Waiting for User", _
            "Waiting for Vendor"), Operator:=xlFilterValues
        End With
    End If
    'Queued Ticket
    If Not Intersect(Target, Range("V16")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="CHG"
            .AutoFilter Field:=11, Criteria1:="1"
            .AutoFilter Field:=15, Criteria1:="="
        End With
    End If
    '1-15 D Ticket
    If Not Intersect(Target, Range("V17")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="CHG"
            .AutoFilter Field:=11, Criteria1:="1"
            .AutoFilter Field:=15, Criteria1:=">=1", Operator:=xlAnd, Criteria2:="<=15"
        End With
    End If
    '16-30 D Ticket
    If Not Intersect(Target, Range("V18")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="CHG"
            .AutoFilter Field:=11, Criteria1:="1"
            .AutoFilter Field:=15, Criteria1:=">=16", Operator:=xlAnd, Criteria2:="<=30"
        End With
    End If
    '31-45 D Ticket
    If Not Intersect(Target, Range("V19")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="CHG"
            .AutoFilter Field:=11, Criteria1:="1"
            .AutoFilter Field:=15, Criteria1:=">=31", Operator:=xlAnd, Criteria2:="<=45"
        End With
    End If
    '46-60 D Ticket
    If Not Intersect(Target, Range("V20")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="CHG"
            .AutoFilter Field:=11, Criteria1:="1"
            .AutoFilter Field:=15, Criteria1:=">=46", Operator:=xlAnd, Criteria2:="<=60"
        End With
    End If
    ' >60 D Ticket
    If Not Intersect(Target, Range("V21")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="CHG"
            .AutoFilter Field:=11, Criteria1:="1"
            .AutoFilter Field:=15, Criteria1:=">60"
        End With
    End If
'--------------------------- Priority 2
    'Opening Balance Ticket
    If Not Intersect(Target, Range("W10")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="CHG"
            .AutoFilter Field:=11, Criteria1:="2"
            .AutoFilter Field:=9, Criteria1:="<" & today
        End With
    End If
    'Received Ticket
    If Not Intersect(Target, Range("W11")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="CHG"
            .AutoFilter Field:=11, Criteria1:="2"
            .AutoFilter Field:=9, Criteria1:=">=" & today
        End With
    End If
    'Resolved Ticket
    If Not Intersect(Target, Range("W13")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="CHG"
            .AutoFilter Field:=11, Criteria1:="2"
            .AutoFilter Field:=10, Criteria1:="<>"
        End With
    End If
    'Carried Over Ticket
    If Not Intersect(Target, Range("W14")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="CHG"
            .AutoFilter Field:=11, Criteria1:="2"
            .AutoFilter Field:=10, Criteria1:="="
        End With
    End If
    'SLA Hold Ticket
    If Not Intersect(Target, Range("W15")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="CHG"
            .AutoFilter Field:=11, Criteria1:="2"
            .AutoFilter Field:=13, Criteria1:=Array( _
            "PENDING", "SLAHOLD", "Waiting for Problem", "Waiting for User", _
            "Waiting for Vendor"), Operator:=xlFilterValues
        End With
    End If
    'Queued Ticket
    If Not Intersect(Target, Range("W16")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="CHG"
            .AutoFilter Field:=11, Criteria1:="2"
            .AutoFilter Field:=15, Criteria1:="="
        End With
    End If
    '1-15 D Ticket
    If Not Intersect(Target, Range("W17")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="CHG"
            .AutoFilter Field:=11, Criteria1:="2"
            .AutoFilter Field:=15, Criteria1:=">=1", Operator:=xlAnd, Criteria2:="<=15"
        End With
    End If
    '16-30 D Ticket
    If Not Intersect(Target, Range("W18")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="CHG"
            .AutoFilter Field:=11, Criteria1:="2"
            .AutoFilter Field:=15, Criteria1:=">=16", Operator:=xlAnd, Criteria2:="<=30"
        End With
    End If
    '31-45 D Ticket
    If Not Intersect(Target, Range("W19")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="CHG"
            .AutoFilter Field:=11, Criteria1:="2"
            .AutoFilter Field:=15, Criteria1:=">=31", Operator:=xlAnd, Criteria2:="<=45"
        End With
    End If
    '46-60 D Ticket
    If Not Intersect(Target, Range("W20")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="CHG"
            .AutoFilter Field:=11, Criteria1:="2"
            .AutoFilter Field:=15, Criteria1:=">=46", Operator:=xlAnd, Criteria2:="<=60"
        End With
    End If
    ' >60 D Ticket
    If Not Intersect(Target, Range("W21")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="CHG"
            .AutoFilter Field:=11, Criteria1:="2"
            .AutoFilter Field:=15, Criteria1:=">60"
        End With
    End If
'--------------------------- Priority 3
    'Opening Balance Ticket
    If Not Intersect(Target, Range("X10")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="CHG"
            .AutoFilter Field:=11, Criteria1:="3"
            .AutoFilter Field:=9, Criteria1:="<" & today
        End With
    End If
    'Received Ticket
    If Not Intersect(Target, Range("X11")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="CHG"
            .AutoFilter Field:=11, Criteria1:="3"
            .AutoFilter Field:=9, Criteria1:=">=" & today
        End With
    End If
    'Resolved Ticket
    If Not Intersect(Target, Range("X13")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="CHG"
            .AutoFilter Field:=11, Criteria1:="3"
            .AutoFilter Field:=10, Criteria1:="<>"
        End With
    End If
    'Carried Over Ticket
    If Not Intersect(Target, Range("X14")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="CHG"
            .AutoFilter Field:=11, Criteria1:="3"
            .AutoFilter Field:=10, Criteria1:="="
        End With
    End If
    'SLA Hold Ticket
    If Not Intersect(Target, Range("X15")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="CHG"
            .AutoFilter Field:=11, Criteria1:="3"
            .AutoFilter Field:=13, Criteria1:=Array( _
            "PENDING", "SLAHOLD", "Waiting for Problem", "Waiting for User", _
            "Waiting for Vendor"), Operator:=xlFilterValues
        End With
    End If
    'Queued Ticket
    If Not Intersect(Target, Range("X16")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="CHG"
            .AutoFilter Field:=11, Criteria1:="3"
            .AutoFilter Field:=15, Criteria1:="="
        End With
    End If
    '1-15 D Ticket
    If Not Intersect(Target, Range("X17")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="CHG"
            .AutoFilter Field:=11, Criteria1:="3"
            .AutoFilter Field:=15, Criteria1:=">=1", Operator:=xlAnd, Criteria2:="<=15"
        End With
    End If
    '16-30 D Ticket
    If Not Intersect(Target, Range("X18")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="CHG"
            .AutoFilter Field:=11, Criteria1:="3"
            .AutoFilter Field:=15, Criteria1:=">=16", Operator:=xlAnd, Criteria2:="<=30"
        End With
    End If
    '31-45 D Ticket
    If Not Intersect(Target, Range("X19")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="CHG"
            .AutoFilter Field:=11, Criteria1:="3"
            .AutoFilter Field:=15, Criteria1:=">=31", Operator:=xlAnd, Criteria2:="<=45"
        End With
    End If
    '46-60 D Ticket
    If Not Intersect(Target, Range("X20")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="CHG"
            .AutoFilter Field:=11, Criteria1:="3"
            .AutoFilter Field:=15, Criteria1:=">=46", Operator:=xlAnd, Criteria2:="<=60"
        End With
    End If
    ' >60 D Ticket
    If Not Intersect(Target, Range("X21")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="CHG"
            .AutoFilter Field:=11, Criteria1:="3"
            .AutoFilter Field:=15, Criteria1:=">60"
        End With
    End If
'--------------------------- Priority 4
    'Opening Balance Ticket
    If Not Intersect(Target, Range("Y10")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="CHG"
            .AutoFilter Field:=11, Criteria1:="4"
            .AutoFilter Field:=9, Criteria1:="<" & today
        End With
    End If
    'Received Ticket
    If Not Intersect(Target, Range("Y11")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="CHG"
            .AutoFilter Field:=11, Criteria1:="4"
            .AutoFilter Field:=9, Criteria1:=">=" & today
        End With
    End If
    'Resolved Ticket
    If Not Intersect(Target, Range("Y13")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="CHG"
            .AutoFilter Field:=11, Criteria1:="4"
            .AutoFilter Field:=10, Criteria1:="<>"
        End With
    End If
    'Carried Over Ticket
    If Not Intersect(Target, Range("Y14")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="CHG"
            .AutoFilter Field:=11, Criteria1:="4"
            .AutoFilter Field:=10, Criteria1:="="
        End With
    End If
    'SLA Hold Ticket
    If Not Intersect(Target, Range("Y15")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="CHG"
            .AutoFilter Field:=11, Criteria1:="4"
            .AutoFilter Field:=13, Criteria1:=Array( _
            "PENDING", "SLAHOLD", "Waiting for Problem", "Waiting for User", _
            "Waiting for Vendor"), Operator:=xlFilterValues
        End With
    End If
    'Queued Ticket
    If Not Intersect(Target, Range("Y16")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="CHG"
            .AutoFilter Field:=11, Criteria1:="4"
            .AutoFilter Field:=15, Criteria1:="="
        End With
    End If
    '1-15 D Ticket
    If Not Intersect(Target, Range("Y17")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="CHG"
            .AutoFilter Field:=11, Criteria1:="4"
            .AutoFilter Field:=15, Criteria1:=">=1", Operator:=xlAnd, Criteria2:="<=15"
        End With
    End If
    '16-30 D Ticket
    If Not Intersect(Target, Range("Y18")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="CHG"
            .AutoFilter Field:=11, Criteria1:="4"
            .AutoFilter Field:=15, Criteria1:=">=16", Operator:=xlAnd, Criteria2:="<=30"
        End With
    End If
    '31-45 D Ticket
    If Not Intersect(Target, Range("Y19")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="CHG"
            .AutoFilter Field:=11, Criteria1:="4"
            .AutoFilter Field:=15, Criteria1:=">=31", Operator:=xlAnd, Criteria2:="<=45"
        End With
    End If
    '46-60 D Ticket
    If Not Intersect(Target, Range("Y20")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="CHG"
            .AutoFilter Field:=11, Criteria1:="4"
            .AutoFilter Field:=15, Criteria1:=">=46", Operator:=xlAnd, Criteria2:="<=60"
        End With
    End If
    ' >60 D Ticket
    If Not Intersect(Target, Range("Y21")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="CHG"
            .AutoFilter Field:=11, Criteria1:="4"
            .AutoFilter Field:=15, Criteria1:=">60"
        End With
    End If
    
'--------------------------- Priority 5
    'Opening Balance Ticket
    If Not Intersect(Target, Range("Z10")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="CHG"
            .AutoFilter Field:=11, Criteria1:="5"
            .AutoFilter Field:=9, Criteria1:="<" & today
        End With
    End If
    'Received Ticket
    If Not Intersect(Target, Range("Z11")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="CHG"
            .AutoFilter Field:=11, Criteria1:="5"
            .AutoFilter Field:=9, Criteria1:=">=" & today
        End With
    End If
    'Resolved Ticket
    If Not Intersect(Target, Range("Z13")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="CHG"
            .AutoFilter Field:=11, Criteria1:="5"
            .AutoFilter Field:=10, Criteria1:="<>"
        End With
    End If
    'Carried Over Ticket
    If Not Intersect(Target, Range("Z14")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="CHG"
            .AutoFilter Field:=11, Criteria1:="5"
            .AutoFilter Field:=10, Criteria1:="="
        End With
    End If
    'SLA Hold Ticket
    If Not Intersect(Target, Range("Z15")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="CHG"
            .AutoFilter Field:=11, Criteria1:="5"
            .AutoFilter Field:=13, Criteria1:=Array( _
            "PENDING", "SLAHOLD", "Waiting for Problem", "Waiting for User", _
            "Waiting for Vendor"), Operator:=xlFilterValues
        End With
    End If
    'Queued Ticket
    If Not Intersect(Target, Range("Z16")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="CHG"
            .AutoFilter Field:=11, Criteria1:="5"
            .AutoFilter Field:=15, Criteria1:="="
        End With
    End If
    '1-15 D Ticket
    If Not Intersect(Target, Range("Z17")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="CHG"
            .AutoFilter Field:=11, Criteria1:="5"
            .AutoFilter Field:=15, Criteria1:=">=1", Operator:=xlAnd, Criteria2:="<=15"
        End With
    End If
    '16-30 D Ticket
    If Not Intersect(Target, Range("Z18")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="CHG"
            .AutoFilter Field:=11, Criteria1:="5"
            .AutoFilter Field:=15, Criteria1:=">=16", Operator:=xlAnd, Criteria2:="<=30"
        End With
    End If
    '31-45 D Ticket
    If Not Intersect(Target, Range("Z19")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="CHG"
            .AutoFilter Field:=11, Criteria1:="5"
            .AutoFilter Field:=15, Criteria1:=">=31", Operator:=xlAnd, Criteria2:="<=45"
        End With
    End If
    '46-60 D Ticket
    If Not Intersect(Target, Range("Z20")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="CHG"
            .AutoFilter Field:=11, Criteria1:="5"
            .AutoFilter Field:=15, Criteria1:=">=46", Operator:=xlAnd, Criteria2:="<=60"
        End With
    End If
    ' >60 D Ticket
    If Not Intersect(Target, Range("Z21")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="CHG"
            .AutoFilter Field:=11, Criteria1:="5"
            .AutoFilter Field:=15, Criteria1:=">60"
        End With
    End If
'----------------------------- END OF CHANGE REQUEST TICKET ---------------------------------
'----------------- PROBLEM TICKET -------------------------------
'--------------------------- Priority 1
    'Opening Balance Ticket
    If Not Intersect(Target, Range("AB10")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="PRB"
            .AutoFilter Field:=11, Criteria1:="1"
            .AutoFilter Field:=9, Criteria1:="<" & today
        End With
    End If
    'Received Ticket
    If Not Intersect(Target, Range("AB11")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="PRB"
            .AutoFilter Field:=11, Criteria1:="1"
            .AutoFilter Field:=9, Criteria1:=">=" & today
        End With
    End If
    'Responded Ticket
    If Not Intersect(Target, Range("AB12")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="PRB"
            .AutoFilter Field:=11, Criteria1:="1"
            .AutoFilter Field:=9, Criteria1:=">=" & today
            .AutoFilter Field:=3, Criteria1:="=N", Operator:=xlOr, Criteria2:="=Y"
        End With
    End If
    'Resolved Ticket
    If Not Intersect(Target, Range("AB13")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="PRB"
            .AutoFilter Field:=11, Criteria1:="1"
            .AutoFilter Field:=10, Criteria1:="<>"
        End With
    End If
    'Carried Over Ticket
    If Not Intersect(Target, Range("AB14")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="PRB"
            .AutoFilter Field:=11, Criteria1:="1"
            .AutoFilter Field:=10, Criteria1:="="
        End With
    End If
    'SLA Hold Ticket
    If Not Intersect(Target, Range("AB15")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="PRB"
            .AutoFilter Field:=11, Criteria1:="1"
            .AutoFilter Field:=13, Criteria1:=Array( _
            "PENDING", "SLAHOLD", "Waiting for Problem", "Waiting for User", _
            "Waiting for Vendor"), Operator:=xlFilterValues
        End With
    End If
    'Queued Ticket
    If Not Intersect(Target, Range("AB16")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="PRB"
            .AutoFilter Field:=11, Criteria1:="1"
            .AutoFilter Field:=15, Criteria1:="="
        End With
    End If
    '1-15 D Ticket
    If Not Intersect(Target, Range("AB17")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="PRB"
            .AutoFilter Field:=11, Criteria1:="1"
            .AutoFilter Field:=15, Criteria1:=">=1", Operator:=xlAnd, Criteria2:="<=15"
        End With
    End If
    '16-30 D Ticket
    If Not Intersect(Target, Range("AB18")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="PRB"
            .AutoFilter Field:=11, Criteria1:="1"
            .AutoFilter Field:=15, Criteria1:=">=16", Operator:=xlAnd, Criteria2:="<=30"
        End With
    End If
    '31-45 D Ticket
    If Not Intersect(Target, Range("AB19")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="PRB"
            .AutoFilter Field:=11, Criteria1:="1"
            .AutoFilter Field:=15, Criteria1:=">=31", Operator:=xlAnd, Criteria2:="<=45"
        End With
    End If
    '46-60 D Ticket
    If Not Intersect(Target, Range("AB20")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="PRB"
            .AutoFilter Field:=11, Criteria1:="1"
            .AutoFilter Field:=15, Criteria1:=">=46", Operator:=xlAnd, Criteria2:="<=60"
        End With
    End If
    ' >60 D Ticket
    If Not Intersect(Target, Range("AB21")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="PRB"
            .AutoFilter Field:=11, Criteria1:="1"
            .AutoFilter Field:=15, Criteria1:=">60"
        End With
    End If
    'Response SLA
    If Not Intersect(Target, Range("AB24")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="PRB"
            .AutoFilter Field:=11, Criteria1:="1"
            .AutoFilter Field:=9, Criteria1:=">=" & today
            .AutoFilter Field:=3, Criteria1:="=Y"
        End With
    End If
    'Resolution SLA
    If Not Intersect(Target, Range("AB25")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="PRB"
            .AutoFilter Field:=11, Criteria1:="1"
            .AutoFilter Field:=4, Criteria1:="=Y"
        End With
    End If
'--------------------------- Priority 2
    'Opening Balance Ticket
    If Not Intersect(Target, Range("AC10")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="PRB"
            .AutoFilter Field:=11, Criteria1:="2"
            .AutoFilter Field:=9, Criteria1:="<" & today
        End With
    End If
    'Received Ticket
    If Not Intersect(Target, Range("AC11")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="PRB"
            .AutoFilter Field:=11, Criteria1:="2"
            .AutoFilter Field:=9, Criteria1:=">=" & today
        End With
    End If
    'Responded Ticket
    If Not Intersect(Target, Range("AC12")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="PRB"
            .AutoFilter Field:=11, Criteria1:="2"
            .AutoFilter Field:=9, Criteria1:=">=" & today
            .AutoFilter Field:=3, Criteria1:="=N", Operator:=xlOr, Criteria2:="=Y"
        End With
    End If
    'Resolved Ticket
    If Not Intersect(Target, Range("AC13")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="PRB"
            .AutoFilter Field:=11, Criteria1:="2"
            .AutoFilter Field:=10, Criteria1:="<>"
        End With
    End If
    'Carried Over Ticket
    If Not Intersect(Target, Range("AC14")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="PRB"
            .AutoFilter Field:=11, Criteria1:="2"
            .AutoFilter Field:=10, Criteria1:="="
        End With
    End If
    'SLA Hold Ticket
    If Not Intersect(Target, Range("AC15")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="PRB"
            .AutoFilter Field:=11, Criteria1:="2"
            .AutoFilter Field:=13, Criteria1:=Array( _
            "PENDING", "SLAHOLD", "Waiting for Problem", "Waiting for User", _
            "Waiting for Vendor"), Operator:=xlFilterValues
        End With
    End If
    'Queued Ticket
    If Not Intersect(Target, Range("AC16")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="PRB"
            .AutoFilter Field:=11, Criteria1:="2"
            .AutoFilter Field:=15, Criteria1:="="
        End With
    End If
    '1-15 D Ticket
    If Not Intersect(Target, Range("AC17")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="PRB"
            .AutoFilter Field:=11, Criteria1:="2"
            .AutoFilter Field:=15, Criteria1:=">=1", Operator:=xlAnd, Criteria2:="<=15"
        End With
    End If
    '16-30 D Ticket
    If Not Intersect(Target, Range("AC18")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="PRB"
            .AutoFilter Field:=11, Criteria1:="2"
            .AutoFilter Field:=15, Criteria1:=">=16", Operator:=xlAnd, Criteria2:="<=30"
        End With
    End If
    '31-45 D Ticket
    If Not Intersect(Target, Range("AC19")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="PRB"
            .AutoFilter Field:=11, Criteria1:="2"
            .AutoFilter Field:=15, Criteria1:=">=31", Operator:=xlAnd, Criteria2:="<=45"
        End With
    End If
    '46-60 D Ticket
    If Not Intersect(Target, Range("AC20")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="PRB"
            .AutoFilter Field:=11, Criteria1:="2"
            .AutoFilter Field:=15, Criteria1:=">=46", Operator:=xlAnd, Criteria2:="<=60"
        End With
    End If
    ' >60 D Ticket
    If Not Intersect(Target, Range("AC21")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="PRB"
            .AutoFilter Field:=11, Criteria1:="2"
            .AutoFilter Field:=15, Criteria1:=">60"
        End With
    End If
    'Response SLA
    If Not Intersect(Target, Range("AC24")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="PRB"
            .AutoFilter Field:=11, Criteria1:="2"
            .AutoFilter Field:=9, Criteria1:=">=" & today
            .AutoFilter Field:=3, Criteria1:="=Y"
        End With
    End If
    'Resolution SLA
    If Not Intersect(Target, Range("AC25")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="PRB"
            .AutoFilter Field:=11, Criteria1:="2"
            .AutoFilter Field:=4, Criteria1:="=Y"
        End With
    End If
'--------------------------- Priority 3
    'Opening Balance Ticket
    If Not Intersect(Target, Range("AD10")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="PRB"
            .AutoFilter Field:=11, Criteria1:="3"
            .AutoFilter Field:=9, Criteria1:="<" & today
        End With
    End If
    'Received Ticket
    If Not Intersect(Target, Range("AD11")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="PRB"
            .AutoFilter Field:=11, Criteria1:="3"
            .AutoFilter Field:=9, Criteria1:=">=" & today
        End With
    End If
    'Responded Ticket
    If Not Intersect(Target, Range("AD12")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="PRB"
            .AutoFilter Field:=11, Criteria1:="3"
            .AutoFilter Field:=9, Criteria1:=">=" & today
            .AutoFilter Field:=3, Criteria1:="=N", Operator:=xlOr, Criteria2:="=Y"
        End With
    End If
    'Resolved Ticket
    If Not Intersect(Target, Range("AD13")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="PRB"
            .AutoFilter Field:=11, Criteria1:="3"
            .AutoFilter Field:=10, Criteria1:="<>"
        End With
    End If
    'Carried Over Ticket
    If Not Intersect(Target, Range("AD14")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="PRB"
            .AutoFilter Field:=11, Criteria1:="3"
            .AutoFilter Field:=10, Criteria1:="="
        End With
    End If
    'SLA Hold Ticket
    If Not Intersect(Target, Range("AD15")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="PRB"
            .AutoFilter Field:=11, Criteria1:="3"
            .AutoFilter Field:=13, Criteria1:=Array( _
            "PENDING", "SLAHOLD", "Waiting for Problem", "Waiting for User", _
            "Waiting for Vendor"), Operator:=xlFilterValues
        End With
    End If
    'Queued Ticket
    If Not Intersect(Target, Range("AD16")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="PRB"
            .AutoFilter Field:=11, Criteria1:="3"
            .AutoFilter Field:=15, Criteria1:="="
        End With
    End If
    '1-15 D Ticket
    If Not Intersect(Target, Range("AD17")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="PRB"
            .AutoFilter Field:=11, Criteria1:="3"
            .AutoFilter Field:=15, Criteria1:=">=1", Operator:=xlAnd, Criteria2:="<=15"
        End With
    End If
    '16-30 D Ticket
    If Not Intersect(Target, Range("AD18")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="PRB"
            .AutoFilter Field:=11, Criteria1:="3"
            .AutoFilter Field:=15, Criteria1:=">=16", Operator:=xlAnd, Criteria2:="<=30"
        End With
    End If
    '31-45 D Ticket
    If Not Intersect(Target, Range("AD19")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="PRB"
            .AutoFilter Field:=11, Criteria1:="3"
            .AutoFilter Field:=15, Criteria1:=">=31", Operator:=xlAnd, Criteria2:="<=45"
        End With
    End If
    '46-60 D Ticket
    If Not Intersect(Target, Range("AD20")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="PRB"
            .AutoFilter Field:=11, Criteria1:="3"
            .AutoFilter Field:=15, Criteria1:=">=46", Operator:=xlAnd, Criteria2:="<=60"
        End With
    End If
    ' >60 D Ticket
    If Not Intersect(Target, Range("AD21")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="PRB"
            .AutoFilter Field:=11, Criteria1:="3"
            .AutoFilter Field:=15, Criteria1:=">60"
        End With
    End If
    'Response SLA
    If Not Intersect(Target, Range("AD24")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="PRB"
            .AutoFilter Field:=11, Criteria1:="3"
            .AutoFilter Field:=9, Criteria1:=">=" & today
            .AutoFilter Field:=3, Criteria1:="=Y"
        End With
    End If
    'Resolution SLA
    If Not Intersect(Target, Range("AD25")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="PRB"
            .AutoFilter Field:=11, Criteria1:="3"
            .AutoFilter Field:=4, Criteria1:="=Y"
        End With
    End If
'--------------------------- Priority 4
    'Opening Balance Ticket
    If Not Intersect(Target, Range("AE10")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="PRB"
            .AutoFilter Field:=11, Criteria1:="4"
            .AutoFilter Field:=9, Criteria1:="<" & today
        End With
    End If
    'Received Ticket
    If Not Intersect(Target, Range("AE11")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="PRB"
            .AutoFilter Field:=11, Criteria1:="4"
            .AutoFilter Field:=9, Criteria1:=">=" & today
        End With
    End If
    'Responded Ticket
    If Not Intersect(Target, Range("AE12")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="PRB"
            .AutoFilter Field:=11, Criteria1:="4"
            .AutoFilter Field:=9, Criteria1:=">=" & today
            .AutoFilter Field:=3, Criteria1:="=N", Operator:=xlOr, Criteria2:="=Y"
        End With
    End If
    'Resolved Ticket
    If Not Intersect(Target, Range("AE13")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="PRB"
            .AutoFilter Field:=11, Criteria1:="4"
            .AutoFilter Field:=10, Criteria1:="<>"
        End With
    End If
    'Carried Over Ticket
    If Not Intersect(Target, Range("AE14")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="PRB"
            .AutoFilter Field:=11, Criteria1:="4"
            .AutoFilter Field:=10, Criteria1:="="
        End With
    End If
    'SLA Hold Ticket
    If Not Intersect(Target, Range("AE15")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="PRB"
            .AutoFilter Field:=11, Criteria1:="4"
            .AutoFilter Field:=13, Criteria1:=Array( _
            "PENDING", "SLAHOLD", "Waiting for Problem", "Waiting for User", _
            "Waiting for Vendor"), Operator:=xlFilterValues
        End With
    End If
    'Queued Ticket
    If Not Intersect(Target, Range("AE16")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="PRB"
            .AutoFilter Field:=11, Criteria1:="4"
            .AutoFilter Field:=15, Criteria1:="="
        End With
    End If
    '1-15 D Ticket
    If Not Intersect(Target, Range("AE17")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="PRB"
            .AutoFilter Field:=11, Criteria1:="4"
            .AutoFilter Field:=15, Criteria1:=">=1", Operator:=xlAnd, Criteria2:="<=15"
        End With
    End If
    '16-30 D Ticket
    If Not Intersect(Target, Range("AE18")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="PRB"
            .AutoFilter Field:=11, Criteria1:="4"
            .AutoFilter Field:=15, Criteria1:=">=16", Operator:=xlAnd, Criteria2:="<=30"
        End With
    End If
    '31-45 D Ticket
    If Not Intersect(Target, Range("AE19")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="PRB"
            .AutoFilter Field:=11, Criteria1:="4"
            .AutoFilter Field:=15, Criteria1:=">=31", Operator:=xlAnd, Criteria2:="<=45"
        End With
    End If
    '46-60 D Ticket
    If Not Intersect(Target, Range("AE20")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="PRB"
            .AutoFilter Field:=11, Criteria1:="4"
            .AutoFilter Field:=15, Criteria1:=">=46", Operator:=xlAnd, Criteria2:="<=60"
        End With
    End If
    ' >60 D Ticket
    If Not Intersect(Target, Range("AE21")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="PRB"
            .AutoFilter Field:=11, Criteria1:="4"
            .AutoFilter Field:=15, Criteria1:=">60"
        End With
    End If
    'Response SLA
    If Not Intersect(Target, Range("AE24")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="PRB"
            .AutoFilter Field:=11, Criteria1:="4"
            .AutoFilter Field:=9, Criteria1:=">=" & today
            .AutoFilter Field:=3, Criteria1:="=Y"
        End With
    End If
    'Resolution SLA
    If Not Intersect(Target, Range("AE25")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="PRB"
            .AutoFilter Field:=11, Criteria1:="4"
            .AutoFilter Field:=4, Criteria1:="=Y"
        End With
    End If
    
'--------------------------- Priority 5
    'Opening Balance Ticket
    If Not Intersect(Target, Range("AF10")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="PRB"
            .AutoFilter Field:=11, Criteria1:="5"
            .AutoFilter Field:=9, Criteria1:="<" & today
        End With
    End If
    'Received Ticket
    If Not Intersect(Target, Range("AF11")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="PRB"
            .AutoFilter Field:=11, Criteria1:="5"
            .AutoFilter Field:=9, Criteria1:=">=" & today
        End With
    End If
    'Responded Ticket
    If Not Intersect(Target, Range("AF12")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="PRB"
            .AutoFilter Field:=11, Criteria1:="5"
            .AutoFilter Field:=9, Criteria1:=">=" & today
            .AutoFilter Field:=3, Criteria1:="=N", Operator:=xlOr, Criteria2:="=Y"
        End With
    End If
    'Resolved Ticket
    If Not Intersect(Target, Range("AF13")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="PRB"
            .AutoFilter Field:=11, Criteria1:="5"
            .AutoFilter Field:=10, Criteria1:="<>"
        End With
    End If
    'Carried Over Ticket
    If Not Intersect(Target, Range("AF14")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="PRB"
            .AutoFilter Field:=11, Criteria1:="5"
            .AutoFilter Field:=10, Criteria1:="="
        End With
    End If
    'SLA Hold Ticket
    If Not Intersect(Target, Range("AF15")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="PRB"
            .AutoFilter Field:=11, Criteria1:="5"
            .AutoFilter Field:=13, Criteria1:=Array( _
            "PENDING", "SLAHOLD", "Waiting for Problem", "Waiting for User", _
            "Waiting for Vendor"), Operator:=xlFilterValues
        End With
    End If
    'Queued Ticket
    If Not Intersect(Target, Range("AF16")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="PRB"
            .AutoFilter Field:=11, Criteria1:="5"
            .AutoFilter Field:=15, Criteria1:="="
        End With
    End If
    '1-15 D Ticket
    If Not Intersect(Target, Range("AF17")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="PRB"
            .AutoFilter Field:=11, Criteria1:="5"
            .AutoFilter Field:=15, Criteria1:=">=1", Operator:=xlAnd, Criteria2:="<=15"
        End With
    End If
    '16-30 D Ticket
    If Not Intersect(Target, Range("AF18")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="PRB"
            .AutoFilter Field:=11, Criteria1:="5"
            .AutoFilter Field:=15, Criteria1:=">=16", Operator:=xlAnd, Criteria2:="<=30"
        End With
    End If
    '31-45 D Ticket
    If Not Intersect(Target, Range("AF19")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="PRB"
            .AutoFilter Field:=11, Criteria1:="5"
            .AutoFilter Field:=15, Criteria1:=">=31", Operator:=xlAnd, Criteria2:="<=45"
        End With
    End If
    '46-60 D Ticket
    If Not Intersect(Target, Range("AF20")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="PRB"
            .AutoFilter Field:=11, Criteria1:="5"
            .AutoFilter Field:=15, Criteria1:=">=46", Operator:=xlAnd, Criteria2:="<=60"
        End With
    End If
    ' >60 D Ticket
    If Not Intersect(Target, Range("AF21")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="PRB"
            .AutoFilter Field:=11, Criteria1:="5"
            .AutoFilter Field:=15, Criteria1:=">60"
        End With
    End If
    'Response SLA
    If Not Intersect(Target, Range("AF24")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="PRB"
            .AutoFilter Field:=11, Criteria1:="5"
            .AutoFilter Field:=9, Criteria1:=">=" & today
            .AutoFilter Field:=3, Criteria1:="=Y"
        End With
    End If
    'Resolution SLA
    If Not Intersect(Target, Range("AF25")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="PRB"
            .AutoFilter Field:=11, Criteria1:="5"
            .AutoFilter Field:=4, Criteria1:="=Y"
        End With
    End If
'----------------------------- END OF PROBLEM TICKET ------------------------------
End Sub
Sub incidentFilter(ByVal Target As Range, ByVal today As String)

    Sheets("MainData").Select
    Sheets("MainData").Range("A3:P3").Select
'--------------------------- Incident Ticket
'--------------------------- Priority 1
    'Opening Balance Ticket
    If Not Intersect(Target, Range("J10")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="INC"
            .AutoFilter Field:=11, Criteria1:="1"
            .AutoFilter Field:=9, Criteria1:="<" & today
        End With
    End If
    'Received Ticket
    If Not Intersect(Target, Range("J11")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="INC"
            .AutoFilter Field:=11, Criteria1:="1"
            .AutoFilter Field:=9, Criteria1:=">=" & today
        End With
    End If
    'Responded Ticket
    If Not Intersect(Target, Range("J12")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="INC"
            .AutoFilter Field:=11, Criteria1:="1"
            .AutoFilter Field:=9, Criteria1:=">=" & today
            .AutoFilter Field:=3, Criteria1:="=N", Operator:=xlOr, Criteria2:="=Y"
        End With
    End If
    'Resolved Ticket
    If Not Intersect(Target, Range("J13")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="INC"
            .AutoFilter Field:=11, Criteria1:="1"
            .AutoFilter Field:=10, Criteria1:="<>"
        End With
    End If
    'Carried Over Ticket
    If Not Intersect(Target, Range("J14")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="INC"
            .AutoFilter Field:=11, Criteria1:="1"
            .AutoFilter Field:=10, Criteria1:="="
        End With
    End If
    'SLA Hold Ticket
    If Not Intersect(Target, Range("J15")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="INC"
            .AutoFilter Field:=11, Criteria1:="1"
            .AutoFilter Field:=13, Criteria1:=Array( _
            "PENDING", "SLAHOLD", "Waiting for Problem", "Waiting for User", _
            "Waiting for Vendor"), Operator:=xlFilterValues
        End With
    End If
    'Queued Ticket
    If Not Intersect(Target, Range("J16")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="INC"
            .AutoFilter Field:=11, Criteria1:="1"
            .AutoFilter Field:=15, Criteria1:="="
        End With
    End If
    '1-15 D Ticket
    If Not Intersect(Target, Range("J17")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="INC"
            .AutoFilter Field:=11, Criteria1:="1"
            .AutoFilter Field:=15, Criteria1:=">=1", Operator:=xlAnd, Criteria2:="<=15"
        End With
    End If
    '16-30 D Ticket
    If Not Intersect(Target, Range("J18")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="INC"
            .AutoFilter Field:=11, Criteria1:="1"
            .AutoFilter Field:=15, Criteria1:=">=16", Operator:=xlAnd, Criteria2:="<=30"
        End With
    End If
    '31-45 D Ticket
    If Not Intersect(Target, Range("J19")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="INC"
            .AutoFilter Field:=11, Criteria1:="1"
            .AutoFilter Field:=15, Criteria1:=">=31", Operator:=xlAnd, Criteria2:="<=45"
        End With
    End If
    '46-60 D Ticket
    If Not Intersect(Target, Range("J20")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="INC"
            .AutoFilter Field:=11, Criteria1:="1"
            .AutoFilter Field:=15, Criteria1:=">=46", Operator:=xlAnd, Criteria2:="<=60"
        End With
    End If
    ' >60 D Ticket
    If Not Intersect(Target, Range("J21")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="INC"
            .AutoFilter Field:=11, Criteria1:="1"
            .AutoFilter Field:=15, Criteria1:=">60"
        End With
    End If
    'Response SLA
    If Not Intersect(Target, Range("J24")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="INC"
            .AutoFilter Field:=11, Criteria1:="1"
            .AutoFilter Field:=9, Criteria1:=">=" & today
            .AutoFilter Field:=3, Criteria1:="=Y"
        End With
    End If
    'Resolution SLA
    If Not Intersect(Target, Range("J25")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="INC"
            .AutoFilter Field:=11, Criteria1:="1"
            .AutoFilter Field:=4, Criteria1:="=Y"
        End With
    End If
'--------------------------- Priority 2
    'Opening Balance Ticket
    If Not Intersect(Target, Range("K10")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="INC"
            .AutoFilter Field:=11, Criteria1:="2"
            .AutoFilter Field:=9, Criteria1:="<" & today
        End With
    End If
    'Received Ticket
    If Not Intersect(Target, Range("K11")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="INC"
            .AutoFilter Field:=11, Criteria1:="2"
            .AutoFilter Field:=9, Criteria1:=">=" & today
        End With
    End If
    'Responded Ticket
    If Not Intersect(Target, Range("K12")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="INC"
            .AutoFilter Field:=11, Criteria1:="2"
            .AutoFilter Field:=9, Criteria1:=">=" & today
            .AutoFilter Field:=3, Criteria1:="=N", Operator:=xlOr, Criteria2:="=Y"
        End With
    End If
    'Resolved Ticket
    If Not Intersect(Target, Range("K13")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="INC"
            .AutoFilter Field:=11, Criteria1:="2"
            .AutoFilter Field:=10, Criteria1:="<>"
        End With
    End If
    'Carried Over Ticket
    If Not Intersect(Target, Range("K14")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="INC"
            .AutoFilter Field:=11, Criteria1:="2"
            .AutoFilter Field:=10, Criteria1:="="
        End With
    End If
    'SLA Hold Ticket
    If Not Intersect(Target, Range("K15")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="INC"
            .AutoFilter Field:=11, Criteria1:="2"
            .AutoFilter Field:=13, Criteria1:=Array( _
            "PENDING", "SLAHOLD", "Waiting for Problem", "Waiting for User", _
            "Waiting for Vendor"), Operator:=xlFilterValues
        End With
    End If
    'Queued Ticket
    If Not Intersect(Target, Range("K16")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="INC"
            .AutoFilter Field:=11, Criteria1:="2"
            .AutoFilter Field:=15, Criteria1:="="
        End With
    End If
    '1-15 D Ticket
    If Not Intersect(Target, Range("K17")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="INC"
            .AutoFilter Field:=11, Criteria1:="2"
            .AutoFilter Field:=15, Criteria1:=">=1", Operator:=xlAnd, Criteria2:="<=15"
        End With
    End If
    '16-30 D Ticket
    If Not Intersect(Target, Range("K18")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="INC"
            .AutoFilter Field:=11, Criteria1:="2"
            .AutoFilter Field:=15, Criteria1:=">=16", Operator:=xlAnd, Criteria2:="<=30"
        End With
    End If
    '31-45 D Ticket
    If Not Intersect(Target, Range("K19")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="INC"
            .AutoFilter Field:=11, Criteria1:="2"
            .AutoFilter Field:=15, Criteria1:=">=31", Operator:=xlAnd, Criteria2:="<=45"
        End With
    End If
    '46-60 D Ticket
    If Not Intersect(Target, Range("K20")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="INC"
            .AutoFilter Field:=11, Criteria1:="2"
            .AutoFilter Field:=15, Criteria1:=">=46", Operator:=xlAnd, Criteria2:="<=60"
        End With
    End If
    ' >60 D Ticket
    If Not Intersect(Target, Range("K21")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="INC"
            .AutoFilter Field:=11, Criteria1:="2"
            .AutoFilter Field:=15, Criteria1:=">60"
        End With
    End If
    'Response SLA
    If Not Intersect(Target, Range("K24")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="INC"
            .AutoFilter Field:=11, Criteria1:="2"
            .AutoFilter Field:=9, Criteria1:=">=" & today
            .AutoFilter Field:=3, Criteria1:="=Y"
        End With
    End If
    'Resolution SLA
    If Not Intersect(Target, Range("K25")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="INC"
            .AutoFilter Field:=11, Criteria1:="2"
            .AutoFilter Field:=4, Criteria1:="=Y"
        End With
    End If
'--------------------------- Priority 3
    'Opening Balance Ticket
    If Not Intersect(Target, Range("L10")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="INC"
            .AutoFilter Field:=11, Criteria1:="3"
            .AutoFilter Field:=9, Criteria1:="<" & today
        End With
    End If
    'Received Ticket
    If Not Intersect(Target, Range("L11")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="INC"
            .AutoFilter Field:=11, Criteria1:="3"
            .AutoFilter Field:=9, Criteria1:=">=" & today
        End With
    End If
    'Responded Ticket
    If Not Intersect(Target, Range("L12")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="INC"
            .AutoFilter Field:=11, Criteria1:="3"
            .AutoFilter Field:=9, Criteria1:=">=" & today
            .AutoFilter Field:=3, Criteria1:="=N", Operator:=xlOr, Criteria2:="=Y"
        End With
    End If
    'Resolved Ticket
    If Not Intersect(Target, Range("L13")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="INC"
            .AutoFilter Field:=11, Criteria1:="3"
            .AutoFilter Field:=10, Criteria1:="<>"
        End With
    End If
    'Carried Over Ticket
    If Not Intersect(Target, Range("L14")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="INC"
            .AutoFilter Field:=11, Criteria1:="3"
            .AutoFilter Field:=10, Criteria1:="="
        End With
    End If
    'SLA Hold Ticket
    If Not Intersect(Target, Range("L15")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="INC"
            .AutoFilter Field:=11, Criteria1:="3"
            .AutoFilter Field:=13, Criteria1:=Array( _
            "PENDING", "SLAHOLD", "Waiting for Problem", "Waiting for User", _
            "Waiting for Vendor"), Operator:=xlFilterValues
        End With
    End If
    'Queued Ticket
    If Not Intersect(Target, Range("L16")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="INC"
            .AutoFilter Field:=11, Criteria1:="3"
            .AutoFilter Field:=15, Criteria1:="="
        End With
    End If
    '1-15 D Ticket
    If Not Intersect(Target, Range("L17")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="INC"
            .AutoFilter Field:=11, Criteria1:="3"
            .AutoFilter Field:=15, Criteria1:=">=1", Operator:=xlAnd, Criteria2:="<=15"
        End With
    End If
    '16-30 D Ticket
    If Not Intersect(Target, Range("L18")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="INC"
            .AutoFilter Field:=11, Criteria1:="3"
            .AutoFilter Field:=15, Criteria1:=">=16", Operator:=xlAnd, Criteria2:="<=30"
        End With
    End If
    '31-45 D Ticket
    If Not Intersect(Target, Range("L19")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="INC"
            .AutoFilter Field:=11, Criteria1:="3"
            .AutoFilter Field:=15, Criteria1:=">=31", Operator:=xlAnd, Criteria2:="<=45"
        End With
    End If
    '46-60 D Ticket
    If Not Intersect(Target, Range("L20")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="INC"
            .AutoFilter Field:=11, Criteria1:="3"
            .AutoFilter Field:=15, Criteria1:=">=46", Operator:=xlAnd, Criteria2:="<=60"
        End With
    End If
    ' >60 D Ticket
    If Not Intersect(Target, Range("L21")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="INC"
            .AutoFilter Field:=11, Criteria1:="3"
            .AutoFilter Field:=15, Criteria1:=">60"
        End With
    End If
    'Response SLA
    If Not Intersect(Target, Range("L24")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="INC"
            .AutoFilter Field:=11, Criteria1:="3"
            .AutoFilter Field:=9, Criteria1:=">=" & today
            .AutoFilter Field:=3, Criteria1:="=Y"
        End With
    End If
    'Resolution SLA
    If Not Intersect(Target, Range("L25")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="INC"
            .AutoFilter Field:=11, Criteria1:="3"
            .AutoFilter Field:=4, Criteria1:="=Y"
        End With
    End If
'--------------------------- Priority 4
    'Opening Balance Ticket
    If Not Intersect(Target, Range("M10")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="INC"
            .AutoFilter Field:=11, Criteria1:="4"
            .AutoFilter Field:=9, Criteria1:="<" & today
        End With
    End If
    'Received Ticket
    If Not Intersect(Target, Range("M11")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="INC"
            .AutoFilter Field:=11, Criteria1:="4"
            .AutoFilter Field:=9, Criteria1:=">=" & today
        End With
    End If
    'Responded Ticket
    If Not Intersect(Target, Range("M12")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="INC"
            .AutoFilter Field:=11, Criteria1:="4"
            .AutoFilter Field:=9, Criteria1:=">=" & today
            .AutoFilter Field:=3, Criteria1:="=N", Operator:=xlOr, Criteria2:="=Y"
        End With
    End If
    'Resolved Ticket
    If Not Intersect(Target, Range("M13")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="INC"
            .AutoFilter Field:=11, Criteria1:="4"
            .AutoFilter Field:=10, Criteria1:="<>"
        End With
    End If
    'Carried Over Ticket
    If Not Intersect(Target, Range("M14")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="INC"
            .AutoFilter Field:=11, Criteria1:="4"
            .AutoFilter Field:=10, Criteria1:="="
        End With
    End If
    'SLA Hold Ticket
    If Not Intersect(Target, Range("M15")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="INC"
            .AutoFilter Field:=11, Criteria1:="4"
            .AutoFilter Field:=13, Criteria1:=Array( _
            "PENDING", "SLAHOLD", "Waiting for Problem", "Waiting for User", _
            "Waiting for Vendor"), Operator:=xlFilterValues
        End With
    End If
    'Queued Ticket
    If Not Intersect(Target, Range("M16")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="INC"
            .AutoFilter Field:=11, Criteria1:="4"
            .AutoFilter Field:=15, Criteria1:="="
        End With
    End If
    '1-15 D Ticket
    If Not Intersect(Target, Range("M17")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="INC"
            .AutoFilter Field:=11, Criteria1:="4"
            .AutoFilter Field:=15, Criteria1:=">=1", Operator:=xlAnd, Criteria2:="<=15"
        End With
    End If
    '16-30 D Ticket
    If Not Intersect(Target, Range("M18")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="INC"
            .AutoFilter Field:=11, Criteria1:="4"
            .AutoFilter Field:=15, Criteria1:=">=16", Operator:=xlAnd, Criteria2:="<=30"
        End With
    End If
    '31-45 D Ticket
    If Not Intersect(Target, Range("M19")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="INC"
            .AutoFilter Field:=11, Criteria1:="4"
            .AutoFilter Field:=15, Criteria1:=">=31", Operator:=xlAnd, Criteria2:="<=45"
        End With
    End If
    '46-60 D Ticket
    If Not Intersect(Target, Range("M20")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="INC"
            .AutoFilter Field:=11, Criteria1:="4"
            .AutoFilter Field:=15, Criteria1:=">=46", Operator:=xlAnd, Criteria2:="<=60"
        End With
    End If
    ' >60 D Ticket
    If Not Intersect(Target, Range("M21")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="INC"
            .AutoFilter Field:=11, Criteria1:="4"
            .AutoFilter Field:=15, Criteria1:=">60"
        End With
    End If
    'Response SLA
    If Not Intersect(Target, Range("M24")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="INC"
            .AutoFilter Field:=11, Criteria1:="4"
            .AutoFilter Field:=9, Criteria1:=">=" & today
            .AutoFilter Field:=3, Criteria1:="=Y"
        End With
    End If
    'Resolution SLA
    If Not Intersect(Target, Range("M25")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="INC"
            .AutoFilter Field:=11, Criteria1:="4"
            .AutoFilter Field:=4, Criteria1:="=Y"
        End With
    End If
    
'--------------------------- Priority 5
    'Opening Balance Ticket
    If Not Intersect(Target, Range("N10")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="INC"
            .AutoFilter Field:=11, Criteria1:="5"
            .AutoFilter Field:=9, Criteria1:="<" & today
        End With
    End If
    'Received Ticket
    If Not Intersect(Target, Range("N11")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="INC"
            .AutoFilter Field:=11, Criteria1:="5"
            .AutoFilter Field:=9, Criteria1:=">=" & today
        End With
    End If
    'Responded Ticket
    If Not Intersect(Target, Range("N12")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="INC"
            .AutoFilter Field:=11, Criteria1:="5"
            .AutoFilter Field:=9, Criteria1:=">=" & today
            .AutoFilter Field:=3, Criteria1:="=N", Operator:=xlOr, Criteria2:="=Y"
        End With
    End If
    'Resolved Ticket
    If Not Intersect(Target, Range("N13")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="INC"
            .AutoFilter Field:=11, Criteria1:="5"
            .AutoFilter Field:=10, Criteria1:="<>"
        End With
    End If
    'Carried Over Ticket
    If Not Intersect(Target, Range("N14")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="INC"
            .AutoFilter Field:=11, Criteria1:="5"
            .AutoFilter Field:=10, Criteria1:="="
        End With
    End If
    'SLA Hold Ticket
    If Not Intersect(Target, Range("N15")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="INC"
            .AutoFilter Field:=11, Criteria1:="5"
            .AutoFilter Field:=13, Criteria1:=Array( _
            "PENDING", "SLAHOLD", "Waiting for Problem", "Waiting for User", _
            "Waiting for Vendor"), Operator:=xlFilterValues
        End With
    End If
    'Queued Ticket
    If Not Intersect(Target, Range("N16")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="INC"
            .AutoFilter Field:=11, Criteria1:="5"
            .AutoFilter Field:=15, Criteria1:="="
        End With
    End If
    '1-15 D Ticket
    If Not Intersect(Target, Range("N17")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="INC"
            .AutoFilter Field:=11, Criteria1:="5"
            .AutoFilter Field:=15, Criteria1:=">=1", Operator:=xlAnd, Criteria2:="<=15"
        End With
    End If
    '16-30 D Ticket
    If Not Intersect(Target, Range("N18")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="INC"
            .AutoFilter Field:=11, Criteria1:="5"
            .AutoFilter Field:=15, Criteria1:=">=16", Operator:=xlAnd, Criteria2:="<=30"
        End With
    End If
    '31-45 D Ticket
    If Not Intersect(Target, Range("N19")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="INC"
            .AutoFilter Field:=11, Criteria1:="5"
            .AutoFilter Field:=15, Criteria1:=">=31", Operator:=xlAnd, Criteria2:="<=45"
        End With
    End If
    '46-60 D Ticket
    If Not Intersect(Target, Range("N20")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="INC"
            .AutoFilter Field:=11, Criteria1:="5"
            .AutoFilter Field:=15, Criteria1:=">=46", Operator:=xlAnd, Criteria2:="<=60"
        End With
    End If
    ' >60 D Ticket
    If Not Intersect(Target, Range("N21")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="INC"
            .AutoFilter Field:=11, Criteria1:="5"
            .AutoFilter Field:=15, Criteria1:=">60"
        End With
    End If
    'Response SLA
    If Not Intersect(Target, Range("N24")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="INC"
            .AutoFilter Field:=11, Criteria1:="5"
            .AutoFilter Field:=9, Criteria1:=">=" & today
            .AutoFilter Field:=3, Criteria1:="=Y"
        End With
    End If
    'Resolution SLA
    If Not Intersect(Target, Range("N25")) Is Nothing Then
        With Selection
            .AutoFilter Field:=2, Criteria1:="INC"
            .AutoFilter Field:=11, Criteria1:="5"
            .AutoFilter Field:=4, Criteria1:="=Y"
        End With
    End If

End Sub




