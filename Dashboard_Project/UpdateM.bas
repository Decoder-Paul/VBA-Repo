Attribute VB_Name = "UpdateM"
Sub ticketCount()

'========================================================================================================
' TicketCount
' -------------------------------------------------------------------------------------------------------
' Purpose   :   To get no. of Tickets depending on conditions from Data File to Dashboard.
'
' Author    :   Subhankar Paul, 9th February, 2017
' Notes     :   Different Ticket Types: 'INC', 'SRQ', 'CHG', 'PRB' are string constant
'
' Parameter :   N/A
' Returns   :   N/A
' ---------------------------------------------------------------
' Revision History
'
' -> Closed or Resolved Tickets are not counted in Aging analysis
' -> Queued Tickets are counted on the basis of Actual Date Column in MainData
' -> Hertz: SLAHOLD ; MasterCard & LM: PENDING; NYL: Waiting for Users, Vendors etc
'    are considered for OnHold Tickets
' -> Respond and RespondSLA is now calculated for today's recieved ticket only
' -> Event Ticket is added as new Ticket type
'========================================================================================================

Application.ScreenUpdating = False
Application.DisplayAlerts = False


Dim sheetData As String
Dim sheetDbd As String
Dim BI As Variant
Dim R As Long
Dim C As Long

sheetData = "MainData"
sheetDbd = "Project or Cluster"

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

Set clean = Range("J10:N25")
clean.Select
Selection.Cells.ClearContents

Set clean = Range("P10:T25")
clean.Select
Selection.Cells.ClearContents

Set clean = Range("V10:Z25")
clean.Select
Selection.Cells.ClearContents

Set clean = Range("AB10:AF25")
clean.Select
Selection.Cells.ClearContents

Set clean = Range("AH10:AL25")
clean.Select
Selection.Cells.ClearContents

'------------ Dictionary Creation for Distinct Count of Assigned resource --
Dim INC_Dict, CHG_Dict, SRQ_Dict, PRB_Dict, EVT_Dict, Res_count_Dict As Object
Set INC_Dict = CreateObject("scripting.dictionary")
Set CHG_Dict = CreateObject("scripting.dictionary")
Set SRQ_Dict = CreateObject("scripting.dictionary")
Set PRB_Dict = CreateObject("scripting.dictionary")
Set EVT_Dict = CreateObject("scripting.dictionary")
Set Res_count_Dict = CreateObject("scripting.dictionary")

'------------ All Counter for Incident Ticket ---------------
 Dim Res_count_Dict_TeamSize As Long
  Res_count_Dict_TeamSize = 0

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

'------------- All Counter for SERVICE Request -----------------
Dim SRQ_opBal_p1 As Long
Dim SRQ_opBal_p2 As Long
Dim SRQ_opBal_p3 As Long
Dim SRQ_opBal_p4 As Long
Dim SRQ_opBal_p5 As Long

Dim SRQ_Recv_p1 As Long
Dim SRQ_Recv_p2 As Long
Dim SRQ_Recv_p3 As Long
Dim SRQ_Recv_p4 As Long
Dim SRQ_Recv_p5 As Long

Dim SRQ_Rspnd_p1 As Long
Dim SRQ_Rspnd_p2 As Long
Dim SRQ_Rspnd_p3 As Long
Dim SRQ_Rspnd_p4 As Long
Dim SRQ_Rspnd_p5 As Long

Dim SRQ_Rsolv_p1 As Long
Dim SRQ_Rsolv_p2 As Long
Dim SRQ_Rsolv_p3 As Long
Dim SRQ_Rsolv_p4 As Long
Dim SRQ_Rsolv_p5 As Long

Dim SRQ_caOvr_p1 As Long
Dim SRQ_caOvr_p2 As Long
Dim SRQ_caOvr_p3 As Long
Dim SRQ_caOvr_p4 As Long
Dim SRQ_caOvr_p5 As Long
Dim SRQ_OnHold_Array(4) As Long

Dim SRQ_Queue_Array(4) As Long

Dim SRQ_Aging_Array(4, 4) As Long

Dim SRQ_Efrt_p1 As Double
Dim SRQ_Efrt_p2 As Double
Dim SRQ_Efrt_p3 As Double
Dim SRQ_Efrt_p4 As Double
Dim SRQ_Efrt_p5 As Double

Dim SRQ_TeamSize As Long

Dim SRQ_RspSLA_p1 As Long
Dim SRQ_RspSLA_p2 As Long
Dim SRQ_RspSLA_p3 As Long
Dim SRQ_RspSLA_p4 As Long
Dim SRQ_RspSLA_p5 As Long

Dim SRQ_ResSLA_p1 As Long
Dim SRQ_ResSLA_p2 As Long
Dim SRQ_ResSLA_p3 As Long
Dim SRQ_ResSLA_p4 As Long
Dim SRQ_ResSLA_p5 As Long

'------------- All Counter for CHANGES Request -----------------
Dim CHG_opBal_p1 As Long
Dim CHG_opBal_p2 As Long
Dim CHG_opBal_p3 As Long
Dim CHG_opBal_p4 As Long
Dim CHG_opBal_p5 As Long

Dim CHG_Recv_p1 As Long
Dim CHG_Recv_p2 As Long
Dim CHG_Recv_p3 As Long
Dim CHG_Recv_p4 As Long
Dim CHG_Recv_p5 As Long

Dim CHG_Rsolv_p1 As Long
Dim CHG_Rsolv_p2 As Long
Dim CHG_Rsolv_p3 As Long
Dim CHG_Rsolv_p4 As Long
Dim CHG_Rsolv_p5 As Long

Dim CHG_caOvr_p1 As Long
Dim CHG_caOvr_p2 As Long
Dim CHG_caOvr_p3 As Long
Dim CHG_caOvr_p4 As Long
Dim CHG_caOvr_p5 As Long
Dim CHG_OnHold_Array(4) As Long

Dim CHG_Queue_Array(4) As Long

Dim CHG_Aging_Array(4, 4) As Long

Dim CHG_Efrt_p1 As Double
Dim CHG_Efrt_p2 As Double
Dim CHG_Efrt_p3 As Double
Dim CHG_Efrt_p4 As Double
Dim CHG_Efrt_p5 As Double

Dim CHG_TeamSize As Long

Dim CHG_RspSLA_p1 As Long
Dim CHG_RspSLA_p2 As Long
Dim CHG_RspSLA_p3 As Long
Dim CHG_RspSLA_p4 As Long
Dim CHG_RspSLA_p5 As Long

Dim CHG_ResSLA_p1 As Long
Dim CHG_ResSLA_p2 As Long
Dim CHG_ResSLA_p3 As Long
Dim CHG_ResSLA_p4 As Long
Dim CHG_ResSLA_p5 As Long

'------------- All Counter for PROBLEM Ticket -----------------
Dim PRB_opBal_p1 As Long
Dim PRB_opBal_p2 As Long
Dim PRB_opBal_p3 As Long
Dim PRB_opBal_p4 As Long
Dim PRB_opBal_p5 As Long

Dim PRB_Recv_p1 As Long
Dim PRB_Recv_p2 As Long
Dim PRB_Recv_p3 As Long
Dim PRB_Recv_p4 As Long
Dim PRB_Recv_p5 As Long

Dim PRB_Rspnd_p1 As Long
Dim PRB_Rspnd_p2 As Long
Dim PRB_Rspnd_p3 As Long
Dim PRB_Rspnd_p4 As Long
Dim PRB_Rspnd_p5 As Long

Dim PRB_Rsolv_p1 As Long
Dim PRB_Rsolv_p2 As Long
Dim PRB_Rsolv_p3 As Long
Dim PRB_Rsolv_p4 As Long
Dim PRB_Rsolv_p5 As Long

Dim PRB_caOvr_p1 As Long
Dim PRB_caOvr_p2 As Long
Dim PRB_caOvr_p3 As Long
Dim PRB_caOvr_p4 As Long
Dim PRB_caOvr_p5 As Long
Dim PRB_OnHold_Array(4) As Long

Dim PRB_Queue_Array(4) As Long

Dim PRB_Aging_Array(4, 4) As Long

Dim PRB_Efrt_p1 As Double
Dim PRB_Efrt_p2 As Double
Dim PRB_Efrt_p3 As Double
Dim PRB_Efrt_p4 As Double
Dim PRB_Efrt_p5 As Double

Dim PRB_TeamSize As Long

Dim PRB_RspSLA_p1 As Long
Dim PRB_RspSLA_p2 As Long
Dim PRB_RspSLA_p3 As Long
Dim PRB_RspSLA_p4 As Long
Dim PRB_RspSLA_p5 As Long

Dim PRB_ResSLA_p1 As Long
Dim PRB_ResSLA_p2 As Long
Dim PRB_ResSLA_p3 As Long
Dim PRB_ResSLA_p4 As Long
Dim PRB_ResSLA_p5 As Long

'------------ All Counter for  EVENT Ticket ---------------
Dim EVT_opBal_p1 As Long
Dim EVT_opBal_p2 As Long
Dim EVT_opBal_p3 As Long
Dim EVT_opBal_p4 As Long
Dim EVT_opBal_p5 As Long

Dim EVT_Recv_p1 As Long
Dim EVT_Recv_p2 As Long
Dim EVT_Recv_p3 As Long
Dim EVT_Recv_p4 As Long
Dim EVT_Recv_p5 As Long

Dim EVT_Rspnd_p1 As Long
Dim EVT_Rspnd_p2 As Long
Dim EVT_Rspnd_p3 As Long
Dim EVT_Rspnd_p4 As Long
Dim EVT_Rspnd_p5 As Long

Dim EVT_Rsolv_p1 As Long
Dim EVT_Rsolv_p2 As Long
Dim EVT_Rsolv_p3 As Long
Dim EVT_Rsolv_p4 As Long
Dim EVT_Rsolv_p5 As Long

Dim EVT_caOvr_p1 As Long
Dim EVT_caOvr_p2 As Long
Dim EVT_caOvr_p3 As Long
Dim EVT_caOvr_p4 As Long
Dim EVT_caOvr_p5 As Long
Dim EVT_OnHold_Array(4) As Long

Dim EVT_Queue_Array(4) As Long

Dim EVT_Aging_Array(4, 4) As Long

Dim EVT_Efrt_p1 As Double
Dim EVT_Efrt_p2 As Double
Dim EVT_Efrt_p3 As Double
Dim EVT_Efrt_p4 As Double
Dim EVT_Efrt_p5 As Double

Dim EVT_TeamSize As Long

Dim EVT_RspSLA_p1 As Long
Dim EVT_RspSLA_p2 As Long
Dim EVT_RspSLA_p3 As Long
Dim EVT_RspSLA_p4 As Long
Dim EVT_RspSLA_p5 As Long

Dim EVT_ResSLA_p1 As Long
Dim EVT_ResSLA_p2 As Long
Dim EVT_ResSLA_p3 As Long
Dim EVT_ResSLA_p4 As Long
Dim EVT_ResSLA_p5 As Long


'------------- Start of Filtering & Counting Calculation -------------

Sheets(sheetData).Select

Dim Data_rowCount, Data_i, j As Integer
Dim UResoCount As Long

Dim tkt_type, resolution, rspnd, person, status, client As String

Dim prty As Integer
Dim effort As Double
Dim open_date, closed_date, today, element, assigned_date As Variant
Dim age_of_tkt As Variant

today = Date - 1


Data_rowCount = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row

    For Data_i = 4 To Data_rowCount
        
        tkt_type = Cells(Data_i, 2).Value ' Ticket Type
        rspnd = Cells(Data_i, 3).Value ' Response SLA
        resolution = Cells(Data_i, 4).Value ' Resolution SLA
        element = Cells(Data_i, 8).Value ' Assigned Resources
        open_date = Cells(Data_i, 9).Value ' Open Date
        open_date = Int(open_date) ' Converting into integer
        If Cells(Data_i, 10).Value = "" Then
            closed_date = ""
        Else
            closed_date = Cells(Data_i, 10).Value
            closed_date = Int(closed_date)
        End If
        prty = Cells(Data_i, 11).Value ' Priority
        effort = Cells(Data_i, 12).Value 'Effort
        status = Cells(Data_i, 13).Value 'Status
        client = Cells(Data_i, 14).Value 'Client
        age_of_tkt = Cells(Data_i, 15).Value ' Aging
        
        ' For Unique count of Assigned resource of Incident,SRQ,CHG,PRB tickets in MainData using Dictionary
        If Res_count_Dict.Exists(element) Then
            Res_count_Dict.Item(element) = Res_count_Dict.Item(element) + 1
        Else
            Res_count_Dict.Add element, 1
        End If
        
        Select Case tkt_type
        'If Incident ticket type
            Case "INC"
                ' For adding REsource count
                If INC_Dict.Exists(element) Then
                    INC_Dict.Item(element) = INC_Dict.Item(element) + 1
                Else
                    INC_Dict.Add element, 1
                End If
                'Priority Checking
                Select Case prty
                    'Priority 1 and so on same for below also
                    Case 1
                    'Effort is calculated here for P1 incident
                        INC_Efrt_p1 = INC_Efrt_p1 + effort
                        
                        'Resolution is calculated here for P1 incident
                        If resolution = "Y" Then
                            INC_ResSLA_p1 = INC_ResSLA_p1 + 1
                        End If
                        
                        'Opening Balance is calculated here for P1 incident
                        If open_date < today Then
                            INC_opBal_p1 = INC_opBal_p1 + 1
                            
                            'Carried Over and Closed is here for P1 incident
                            If closed_date = "" Then
                                INC_caOvr_p1 = INC_caOvr_p1 + 1
                            ElseIf closed_date = today Or closed_date <> "" Then
                                INC_Rsolv_p1 = INC_Rsolv_p1 + 1
                            Else
                                INC_caOvr_p1 = INC_caOvr_p1 + 1
                            End If
                        Else
                            'if the date is Today then
                            'Received is calculated here for P1 incident
                            
                            'Respond and Respond SLA is calculating here for P1 incident
                            If rspnd = "Y" Then
                                INC_RspSLA_p1 = INC_RspSLA_p1 + 1
                                INC_Rspnd_p1 = INC_Rspnd_p1 + 1
                            ElseIf rspnd = "N" Then
                                INC_Rspnd_p1 = INC_Rspnd_p1 + 1
                            End If
                            
                            INC_Recv_p1 = INC_Recv_p1 + 1
                            If closed_date = "" Then
                                INC_caOvr_p1 = INC_caOvr_p1 + 1
                            ElseIf closed_date = today Or closed_date <> "" Then
                                INC_Rsolv_p1 = INC_Rsolv_p1 + 1
                            Else
                                INC_caOvr_p1 = INC_caOvr_p1 + 1
                            End If
                        End If
                        
                        'Queued data is calculated here for P1 incident
                        'Checking on Actual Date Column
                        If Cells(Data_i, 15).Value = "" Then
                            INC_Queue_Array(0) = INC_Queue_Array(0) + 1
                        End If
                        
                        'Checking for On Hold
                        If status = "SLAHOLD" Or Left(UCase(status), 7) = "WAITING" Or UCase(status) = "PENDING" Then
                            INC_OnHold_Array(0) = INC_OnHold_Array(0) + 1
                        End If
                        
                        'Aging is Calculated here for P1 incident
                        If CStr(age_of_tkt) <> "" And age_of_tkt > 0 And closed_date = "" Then
                            If age_of_tkt <= 15 Then
                                INC_Aging_Array(0, 0) = INC_Aging_Array(0, 0) + 1
                            ElseIf age_of_tkt <= 30 Then
                                INC_Aging_Array(1, 0) = INC_Aging_Array(1, 0) + 1
                            ElseIf age_of_tkt <= 45 Then
                                INC_Aging_Array(2, 0) = INC_Aging_Array(2, 0) + 1
                            ElseIf age_of_tkt <= 60 Then
                                INC_Aging_Array(3, 0) = INC_Aging_Array(3, 0) + 1
                            ElseIf age_of_tkt > 60 Then
                                INC_Aging_Array(4, 0) = INC_Aging_Array(4, 0) + 1
                            End If
                        End If
                        
                    Case 2 ' for Case 2
                        INC_Efrt_p2 = INC_Efrt_p2 + effort
                        
                        If resolution = "Y" Then
                            INC_ResSLA_p2 = INC_ResSLA_p2 + 1
                        End If
                        
                        If open_date < today Then
                            INC_opBal_p2 = INC_opBal_p2 + 1
                            If closed_date = "" Then
                                INC_caOvr_p2 = INC_caOvr_p2 + 1
                            ElseIf closed_date = today Or closed_date <> "" Then
                                INC_Rsolv_p2 = INC_Rsolv_p2 + 1
                            Else
                                INC_caOvr_p2 = INC_caOvr_p2 + 1
                            End If
                        Else
                            'Respond SLA calculated here
                            If rspnd = "Y" Then
                                INC_RspSLA_p2 = INC_RspSLA_p2 + 1
                                INC_Rspnd_p2 = INC_Rspnd_p2 + 1
                            ElseIf rspnd = "N" Then
                                INC_Rspnd_p2 = INC_Rspnd_p2 + 1
                            End If
                            
                            INC_Recv_p2 = INC_Recv_p2 + 1
                            If closed_date = "" Then
                                INC_caOvr_p2 = INC_caOvr_p2 + 1
                            ElseIf closed_date = today Or closed_date <> "" Then
                                INC_Rsolv_p2 = INC_Rsolv_p2 + 1
                            Else
                                INC_caOvr_p2 = INC_caOvr_p2 + 1
                            End If
                        End If
                        If Cells(Data_i, 15).Value = "" Then
                            INC_Queue_Array(1) = INC_Queue_Array(1) + 1
                        End If
                        'Checking for On Hold
                        If status = "SLAHOLD" Or Left(UCase(status), 7) = "WAITING" Or UCase(status) = "PENDING" Then
                            INC_OnHold_Array(1) = INC_OnHold_Array(1) + 1
                        End If
                        If CStr(age_of_tkt) <> "" And age_of_tkt > 0 And closed_date = "" Then
                            If age_of_tkt <= 15 Then
                                INC_Aging_Array(0, 1) = INC_Aging_Array(0, 1) + 1
                            ElseIf age_of_tkt <= 30 Then
                                INC_Aging_Array(1, 1) = INC_Aging_Array(1, 1) + 1
                            ElseIf age_of_tkt <= 45 Then
                                INC_Aging_Array(2, 1) = INC_Aging_Array(2, 1) + 1
                            ElseIf age_of_tkt <= 60 Then
                                INC_Aging_Array(3, 1) = INC_Aging_Array(3, 1) + 1
                            ElseIf age_of_tkt > 60 Then
                                INC_Aging_Array(4, 1) = INC_Aging_Array(4, 1) + 1
                            End If
                        End If
                    Case 3
                        INC_Efrt_p3 = INC_Efrt_p3 + effort
                        
                        If resolution = "Y" Then
                            INC_ResSLA_p3 = INC_ResSLA_p3 + 1
                        End If
                        If open_date < today Then
                            INC_opBal_p3 = INC_opBal_p3 + 1
                            If closed_date = "" Then
                                INC_caOvr_p3 = INC_caOvr_p3 + 1
                            ElseIf closed_date = today Or closed_date <> "" Then
                                INC_Rsolv_p3 = INC_Rsolv_p3 + 1
                            Else
                                INC_caOvr_p3 = INC_caOvr_p3 + 1
                            End If
                        Else
                            
                            If rspnd = "Y" Then
                                INC_RspSLA_p3 = INC_RspSLA_p3 + 1
                                INC_Rspnd_p3 = INC_Rspnd_p3 + 1
                            ElseIf rspnd = "N" Then
                                INC_Rspnd_p3 = INC_Rspnd_p3 + 1
                            End If
                                
                            INC_Recv_p3 = INC_Recv_p3 + 1
                            If closed_date = "" Then
                                INC_caOvr_p3 = INC_caOvr_p3 + 1
                            ElseIf closed_date = today Or closed_date <> "" Then
                                INC_Rsolv_p3 = INC_Rsolv_p3 + 1
                            Else
                                INC_caOvr_p3 = INC_caOvr_p3 + 1
                            End If
                        End If
                        If Cells(Data_i, 15).Value = "" Then
                            INC_Queue_Array(2) = INC_Queue_Array(2) + 1
                        End If
                        'Checking for On Hold
                        If status = "SLAHOLD" Or Left(UCase(status), 7) = "WAITING" Or UCase(status) = "PENDING" Then
                            INC_OnHold_Array(2) = INC_OnHold_Array(2) + 1
                        End If
                        If CStr(age_of_tkt) <> "" And age_of_tkt > 0 And closed_date = "" Then
                            If age_of_tkt <= 15 Then
                                INC_Aging_Array(0, 2) = INC_Aging_Array(0, 2) + 1
                            ElseIf age_of_tkt <= 30 Then
                                INC_Aging_Array(1, 2) = INC_Aging_Array(1, 2) + 1
                            ElseIf age_of_tkt <= 45 Then
                                INC_Aging_Array(2, 2) = INC_Aging_Array(2, 2) + 1
                            ElseIf age_of_tkt <= 60 Then
                                INC_Aging_Array(3, 2) = INC_Aging_Array(3, 2) + 1
                            ElseIf age_of_tkt > 60 Then
                                INC_Aging_Array(4, 2) = INC_Aging_Array(4, 2) + 1
                            End If
                        End If
                    Case 4
                        INC_Efrt_p4 = INC_Efrt_p4 + effort
                        
                        If resolution = "Y" Then
                            INC_ResSLA_p4 = INC_ResSLA_p4 + 1
                        End If
                        
                        If open_date < today Then
                            INC_opBal_p4 = INC_opBal_p4 + 1
                            If closed_date = "" Then
                                INC_caOvr_p4 = INC_caOvr_p4 + 1
                            ElseIf closed_date = today Or closed_date <> "" Then
                                INC_Rsolv_p4 = INC_Rsolv_p4 + 1
                            Else
                                INC_caOvr_p4 = INC_caOvr_p4 + 1
                            End If
                        Else
                            If rspnd = "Y" Then
                                INC_RspSLA_p4 = INC_RspSLA_p4 + 1
                                INC_Rspnd_p4 = INC_Rspnd_p4 + 1
                            ElseIf rspnd = "N" Then
                                INC_Rspnd_p4 = INC_Rspnd_p4 + 1
                            End If
                            
                            INC_Recv_p4 = INC_Recv_p4 + 1
                            If closed_date = "" Then
                                INC_caOvr_p4 = INC_caOvr_p4 + 1
                            ElseIf closed_date = today Or closed_date <> "" Then
                                INC_Rsolv_p4 = INC_Rsolv_p4 + 1
                            Else
                                INC_caOvr_p4 = INC_caOvr_p4 + 1
                            End If
                        End If
                        If Cells(Data_i, 15).Value = "" Then
                            INC_Queue_Array(3) = INC_Queue_Array(3) + 1
                        End If
                        'Checking for On Hold
                        If status = "SLAHOLD" Or Left(UCase(status), 7) = "WAITING" Or UCase(status) = "PENDING" Then
                            INC_OnHold_Array(3) = INC_OnHold_Array(3) + 1
                        End If
                        If CStr(age_of_tkt) <> "" And age_of_tkt > 0 And closed_date = "" Then
                            If age_of_tkt <= 15 Then
                                INC_Aging_Array(0, 3) = INC_Aging_Array(0, 3) + 1
                            ElseIf age_of_tkt <= 30 Then
                                INC_Aging_Array(1, 3) = INC_Aging_Array(1, 3) + 1
                            ElseIf age_of_tkt <= 45 Then
                                INC_Aging_Array(2, 3) = INC_Aging_Array(2, 3) + 1
                            ElseIf age_of_tkt <= 60 Then
                                INC_Aging_Array(3, 3) = INC_Aging_Array(3, 3) + 1
                            ElseIf age_of_tkt > 60 Then
                                INC_Aging_Array(4, 3) = INC_Aging_Array(4, 3) + 1
                            End If
                        End If
                    Case 5
                        INC_Efrt_p5 = INC_Efrt_p5 + effort
                        If resolution = "Y" Then
                            INC_ResSLA_p5 = INC_ResSLA_p5 + 1
                        End If
                        If open_date < today Then
                            INC_opBal_p5 = INC_opBal_p5 + 1
                            If closed_date = "" Then
                                INC_caOvr_p5 = INC_caOvr_p5 + 1
                            ElseIf closed_date = today Or closed_date <> "" Then
                                INC_Rsolv_p5 = INC_Rsolv_p5 + 1
                            Else
                                INC_caOvr_p5 = INC_caOvr_p5 + 1
                            End If
                        Else
                            If rspnd = "Y" Then
                                INC_RspSLA_p5 = INC_RspSLA_p5 + 1
                                INC_Rspnd_p5 = INC_Rspnd_p5 + 1
                            ElseIf rspnd = "N" Then
                                INC_Rspnd_p5 = INC_Rspnd_p5 + 1
                            End If
                            
                            INC_Recv_p5 = INC_Recv_p5 + 1
                            If closed_date = "" Then
                                INC_caOvr_p5 = INC_caOvr_p5 + 1
                            ElseIf closed_date = today Or closed_date <> "" Then
                                INC_Rsolv_p5 = INC_Rsolv_p5 + 1
                            Else
                                INC_caOvr_p5 = INC_caOvr_p5 + 1
                            End If
                        End If
                        If Cells(Data_i, 15).Value = "" Then
                            INC_Queue_Array(4) = INC_Queue_Array(4) + 1
                        End If
                        'Checking for On Hold
                        If status = "SLAHOLD" Or Left(UCase(status), 7) = "WAITING" Or UCase(status) = "PENDING" Then
                            INC_OnHold_Array(4) = INC_OnHold_Array(4) + 1
                        End If
                        If CStr(age_of_tkt) <> "" And age_of_tkt > 0 And closed_date = "" Then
                            If age_of_tkt <= 15 Then
                                INC_Aging_Array(0, 4) = INC_Aging_Array(0, 4) + 1
                            ElseIf age_of_tkt <= 30 Then
                                INC_Aging_Array(1, 4) = INC_Aging_Array(1, 4) + 1
                            ElseIf age_of_tkt <= 45 Then
                                INC_Aging_Array(2, 4) = INC_Aging_Array(2, 4) + 1
                            ElseIf age_of_tkt <= 60 Then
                                INC_Aging_Array(3, 4) = INC_Aging_Array(3, 4) + 1
                            ElseIf age_of_tkt > 60 Then
                                INC_Aging_Array(4, 4) = INC_Aging_Array(4, 4) + 1
                            End If
                        End If
                End Select
            Case "SRQ"
                If SRQ_Dict.Exists(element) Then
                    SRQ_Dict.Item(element) = SRQ_Dict.Item(element) + 1
                Else
                    SRQ_Dict.Add element, 1
                End If
                Select Case prty
                    Case 1
                        SRQ_Efrt_p1 = SRQ_Efrt_p1 + effort
                        
                        If resolution = "Y" Then
                            SRQ_ResSLA_p1 = SRQ_ResSLA_p1 + 1
                        End If
                        If open_date < today Then
                            SRQ_opBal_p1 = SRQ_opBal_p1 + 1
                            If closed_date = "" Then
                                SRQ_caOvr_p1 = SRQ_caOvr_p1 + 1
                            ElseIf closed_date = today Or closed_date <> "" Then
                                SRQ_Rsolv_p1 = SRQ_Rsolv_p1 + 1
                            Else
                                SRQ_caOvr_p1 = SRQ_caOvr_p1 + 1
                            End If
                        Else
                            If rspnd = "Y" Then
                                SRQ_RspSLA_p1 = SRQ_RspSLA_p1 + 1
                                SRQ_Rspnd_p1 = SRQ_Rspnd_p1 + 1
                            ElseIf rspnd = "N" Then
                                SRQ_Rspnd_p1 = SRQ_Rspnd_p1 + 1
                            End If
                            
                            SRQ_Recv_p1 = SRQ_Recv_p1 + 1
                            If closed_date = "" Then
                                SRQ_caOvr_p1 = SRQ_caOvr_p1 + 1
                            ElseIf closed_date = today Or closed_date <> "" Then
                                SRQ_Rsolv_p1 = SRQ_Rsolv_p1 + 1
                            Else
                                SRQ_caOvr_p1 = SRQ_caOvr_p1 + 1
                            End If
                        End If
                        If Cells(Data_i, 15).Value = "" Then
                            SRQ_Queue_Array(0) = SRQ_Queue_Array(0) + 1
                        End If
                        'Checking for On Hold
                        If status = "SLAHOLD" Or Left(UCase(status), 7) = "WAITING" Or UCase(status) = "PENDING" Then
                            SRQ_OnHold_Array(0) = SRQ_OnHold_Array(0) + 1
                        End If
                        If CStr(age_of_tkt) <> "" And age_of_tkt > 0 And closed_date = "" Then
                            If age_of_tkt <= 15 Then
                                SRQ_Aging_Array(0, 0) = SRQ_Aging_Array(0, 0) + 1
                            ElseIf age_of_tkt <= 30 Then
                                SRQ_Aging_Array(1, 0) = SRQ_Aging_Array(1, 0) + 1
                            ElseIf age_of_tkt <= 45 Then
                                SRQ_Aging_Array(2, 0) = SRQ_Aging_Array(2, 0) + 1
                            ElseIf age_of_tkt <= 60 Then
                                SRQ_Aging_Array(3, 0) = SRQ_Aging_Array(3, 0) + 1
                            ElseIf age_of_tkt > 60 Then
                                SRQ_Aging_Array(4, 0) = SRQ_Aging_Array(4, 0) + 1
                            End If
                        End If
                    Case 2
                        SRQ_Efrt_p2 = SRQ_Efrt_p2 + effort
                        If resolution = "Y" Then
                            SRQ_ResSLA_p2 = SRQ_ResSLA_p2 + 1
                        End If
                        If open_date < today Then
                            SRQ_opBal_p2 = SRQ_opBal_p2 + 1
                            If closed_date = "" Then
                                SRQ_caOvr_p2 = SRQ_caOvr_p2 + 1
                            ElseIf closed_date = today Or closed_date <> "" Then
                                SRQ_Rsolv_p2 = SRQ_Rsolv_p2 + 1
                            Else
                                SRQ_caOvr_p2 = SRQ_caOvr_p2 + 1
                            End If
                        Else
                            If rspnd = "Y" Then
                                SRQ_RspSLA_p2 = SRQ_RspSLA_p2 + 1
                                SRQ_Rspnd_p2 = SRQ_Rspnd_p2 + 1
                            ElseIf rspnd = "N" Then
                                SRQ_Rspnd_p2 = SRQ_Rspnd_p2 + 1
                            End If
                            
                            SRQ_Recv_p2 = SRQ_Recv_p2 + 1
                            If closed_date = "" Then
                                SRQ_caOvr_p2 = SRQ_caOvr_p2 + 1
                            ElseIf closed_date = today Or closed_date <> "" Then
                                SRQ_Rsolv_p2 = SRQ_Rsolv_p2 + 1
                            Else
                                SRQ_caOvr_p2 = SRQ_caOvr_p2 + 1
                            End If
                        End If
                        If Cells(Data_i, 15).Value = "" Then
                            SRQ_Queue_Array(1) = SRQ_Queue_Array(1) + 1
                        End If
                        'Checking for On Hold
                        If status = "SLAHOLD" Or Left(UCase(status), 7) = "WAITING" Or UCase(status) = "PENDING" Then
                            SRQ_OnHold_Array(1) = SRQ_OnHold_Array(1) + 1
                        End If
                        If CStr(age_of_tkt) <> "" And age_of_tkt > 0 And closed_date = "" Then
                            If age_of_tkt <= 15 Then
                                SRQ_Aging_Array(0, 1) = SRQ_Aging_Array(0, 1) + 1
                            ElseIf age_of_tkt <= 30 Then
                                SRQ_Aging_Array(1, 1) = SRQ_Aging_Array(1, 1) + 1
                            ElseIf age_of_tkt <= 45 Then
                                SRQ_Aging_Array(2, 1) = SRQ_Aging_Array(2, 1) + 1
                            ElseIf age_of_tkt <= 60 Then
                                SRQ_Aging_Array(3, 1) = SRQ_Aging_Array(3, 1) + 1
                            ElseIf age_of_tkt > 60 Then
                                SRQ_Aging_Array(4, 1) = SRQ_Aging_Array(4, 1) + 1
                            End If
                        End If
                    Case 3
                        SRQ_Efrt_p3 = SRQ_Efrt_p3 + effort
                        If resolution = "Y" Then
                            SRQ_ResSLA_p3 = SRQ_ResSLA_p3 + 1
                        End If
                        If open_date < today Then
                            SRQ_opBal_p3 = SRQ_opBal_p3 + 1
                            If closed_date = "" Then
                                SRQ_caOvr_p3 = SRQ_caOvr_p3 + 1
                            ElseIf closed_date = today Or closed_date <> "" Then
                                SRQ_Rsolv_p3 = SRQ_Rsolv_p3 + 1
                            Else
                                SRQ_caOvr_p3 = SRQ_caOvr_p3 + 1
                            End If
                        Else
                            If rspnd = "Y" Then
                                SRQ_RspSLA_p3 = SRQ_RspSLA_p3 + 1
                                SRQ_Rspnd_p3 = SRQ_Rspnd_p3 + 1
                            ElseIf rspnd = "N" Then
                                SRQ_Rspnd_p3 = SRQ_Rspnd_p3 + 1
                            End If
                            
                            SRQ_Recv_p3 = SRQ_Recv_p3 + 1
                            If closed_date = "" Then
                                SRQ_caOvr_p3 = SRQ_caOvr_p3 + 1
                            ElseIf closed_date = today Or closed_date <> "" Then
                                SRQ_Rsolv_p3 = SRQ_Rsolv_p3 + 1
                            Else
                                SRQ_caOvr_p3 = SRQ_caOvr_p3 + 1
                            End If
                        End If
                        If Cells(Data_i, 15).Value = "" Then
                            SRQ_Queue_Array(2) = SRQ_Queue_Array(2) + 1
                        End If
                        'Checking for On Hold
                        If status = "SLAHOLD" Or Left(UCase(status), 7) = "WAITING" Or UCase(status) = "PENDING" Then
                            SRQ_OnHold_Array(2) = SRQ_OnHold_Array(2) + 1
                        End If
                        If CStr(age_of_tkt) <> "" And age_of_tkt > 0 And closed_date = "" Then
                            If age_of_tkt <= 15 Then
                                SRQ_Aging_Array(0, 2) = SRQ_Aging_Array(0, 2) + 1
                            ElseIf age_of_tkt <= 30 Then
                                SRQ_Aging_Array(1, 2) = SRQ_Aging_Array(1, 2) + 1
                            ElseIf age_of_tkt <= 45 Then
                                SRQ_Aging_Array(2, 2) = SRQ_Aging_Array(2, 2) + 1
                            ElseIf age_of_tkt <= 60 Then
                                SRQ_Aging_Array(3, 2) = SRQ_Aging_Array(3, 2) + 1
                            ElseIf age_of_tkt > 60 Then
                                SRQ_Aging_Array(4, 2) = SRQ_Aging_Array(4, 2) + 1
                            End If
                        End If
                    Case 4
                        SRQ_Efrt_p4 = SRQ_Efrt_p4 + effort
                        If resolution = "Y" Then
                            SRQ_ResSLA_p4 = SRQ_ResSLA_p4 + 1
                        End If
                        If open_date < today Then
                            SRQ_opBal_p4 = SRQ_opBal_p4 + 1
                            If closed_date = "" Then
                                SRQ_caOvr_p4 = SRQ_caOvr_p4 + 1
                            ElseIf closed_date = today Or closed_date <> "" Then
                                SRQ_Rsolv_p4 = SRQ_Rsolv_p4 + 1
                            Else
                                SRQ_caOvr_p4 = SRQ_caOvr_p4 + 1
                            End If
                        Else
                            If rspnd = "Y" Then
                                SRQ_RspSLA_p4 = SRQ_RspSLA_p4 + 1
                                SRQ_Rspnd_p4 = SRQ_Rspnd_p4 + 1
                            ElseIf rspnd = "N" Then
                                SRQ_Rspnd_p4 = SRQ_Rspnd_p4 + 1
                            End If
                            
                            SRQ_Recv_p4 = SRQ_Recv_p4 + 1
                            If closed_date = "" Then
                                SRQ_caOvr_p4 = SRQ_caOvr_p4 + 1
                            ElseIf closed_date = today Or closed_date <> "" Then
                                SRQ_Rsolv_p4 = SRQ_Rsolv_p4 + 1
                            Else
                                SRQ_caOvr_p4 = SRQ_caOvr_p4 + 1
                            End If
                        End If
                        If Cells(Data_i, 15).Value = "" Then
                            SRQ_Queue_Array(3) = SRQ_Queue_Array(3) + 1
                        End If
                        'Checking for On Hold
                        If status = "SLAHOLD" Or Left(UCase(status), 7) = "WAITING" Or UCase(status) = "PENDING" Then
                            SRQ_OnHold_Array(3) = SRQ_OnHold_Array(3) + 1
                        End If
                        If CStr(age_of_tkt) <> "" And age_of_tkt > 0 And closed_date = "" Then
                            If age_of_tkt <= 15 Then
                                SRQ_Aging_Array(0, 3) = SRQ_Aging_Array(0, 3) + 1
                            ElseIf age_of_tkt <= 30 Then
                                SRQ_Aging_Array(1, 3) = SRQ_Aging_Array(1, 3) + 1
                            ElseIf age_of_tkt <= 45 Then
                                SRQ_Aging_Array(2, 3) = SRQ_Aging_Array(2, 3) + 1
                            ElseIf age_of_tkt <= 60 Then
                                SRQ_Aging_Array(3, 3) = SRQ_Aging_Array(3, 3) + 1
                            ElseIf age_of_tkt > 60 Then
                                SRQ_Aging_Array(4, 3) = SRQ_Aging_Array(4, 3) + 1
                            End If
                        End If
                    Case 5
                        SRQ_Efrt_p5 = SRQ_Efrt_p5 + effort
                        If resolution = "Y" Then
                            SRQ_ResSLA_p5 = SRQ_ResSLA_p5 + 1
                        End If
                        If open_date < today Then
                            SRQ_opBal_p5 = SRQ_opBal_p5 + 1
                            If closed_date = "" Then
                                SRQ_caOvr_p5 = SRQ_caOvr_p5 + 1
                            ElseIf closed_date = today Or closed_date <> "" Then
                                SRQ_Rsolv_p5 = SRQ_Rsolv_p5 + 1
                            Else
                                SRQ_caOvr_p5 = SRQ_caOvr_p5 + 1
                            End If
                        Else
                            If rspnd = "Y" Then
                                SRQ_RspSLA_p5 = SRQ_RspSLA_p5 + 1
                                SRQ_Rspnd_p5 = SRQ_Rspnd_p5 + 1
                            ElseIf rspnd = "N" Then
                                SRQ_Rspnd_p5 = SRQ_Rspnd_p5 + 1
                            End If
                            
                            SRQ_Recv_p5 = SRQ_Recv_p5 + 1
                            If closed_date = "" Then
                                SRQ_caOvr_p5 = SRQ_caOvr_p5 + 1
                            ElseIf closed_date = today Or closed_date <> "" Then
                                SRQ_Rsolv_p5 = SRQ_Rsolv_p5 + 1
                            Else
                                SRQ_caOvr_p5 = SRQ_caOvr_p5 + 1
                            End If
                        End If
                        If Cells(Data_i, 15).Value = "" Then
                            SRQ_Queue_Array(4) = SRQ_Queue_Array(4) + 1
                        End If
                        'Checking for On Hold
                        If status = "SLAHOLD" Or Left(UCase(status), 7) = "WAITING" Or UCase(status) = "PENDING" Then
                            SRQ_OnHold_Array(4) = SRQ_OnHold_Array(4) + 1
                        End If
                        If CStr(age_of_tkt) <> "" And age_of_tkt > 0 And closed_date = "" Then
                            If age_of_tkt <= 15 Then
                                SRQ_Aging_Array(0, 4) = SRQ_Aging_Array(0, 4) + 1
                            ElseIf age_of_tkt <= 30 Then
                                SRQ_Aging_Array(1, 4) = SRQ_Aging_Array(1, 4) + 1
                            ElseIf age_of_tkt <= 45 Then
                                SRQ_Aging_Array(2, 4) = SRQ_Aging_Array(2, 4) + 1
                            ElseIf age_of_tkt <= 60 Then
                                SRQ_Aging_Array(3, 4) = SRQ_Aging_Array(3, 4) + 1
                            ElseIf age_of_tkt > 60 Then
                                SRQ_Aging_Array(4, 4) = SRQ_Aging_Array(4, 4) + 1
                            End If
                        End If
                End Select
            Case "PRB"
                If PRB_Dict.Exists(element) Then
                    PRB_Dict.Item(element) = PRB_Dict.Item(element) + 1
                Else
                    PRB_Dict.Add element, 1
                End If
                Select Case prty
                    Case 1
                        PRB_Efrt_p1 = PRB_Efrt_p1 + effort
                        If resolution = "Y" Then
                            PRB_ResSLA_p1 = PRB_ResSLA_p1 + 1
                        End If
                        If open_date < today Then
                            PRB_opBal_p1 = PRB_opBal_p1 + 1
                            If closed_date = "" Then
                                PRB_caOvr_p1 = PRB_caOvr_p1 + 1
                            ElseIf closed_date = today Or closed_date <> "" Then
                                PRB_Rsolv_p1 = PRB_Rsolv_p1 + 1
                            Else
                                PRB_caOvr_p1 = PRB_caOvr_p1 + 1
                            End If
                        Else
                            If rspnd = "Y" Then
                                PRB_RspSLA_p1 = PRB_RspSLA_p1 + 1
                                PRB_Rspnd_p1 = PRB_Rspnd_p1 + 1
                            ElseIf rspnd = "N" Then
                                PRB_Rspnd_p1 = PRB_Rspnd_p1 + 1
                            End If
                            
                            PRB_Recv_p1 = PRB_Recv_p1 + 1
                            If closed_date = "" Then
                                PRB_caOvr_p1 = PRB_caOvr_p1 + 1
                            ElseIf closed_date = today Or closed_date <> "" Then
                                PRB_Rsolv_p1 = PRB_Rsolv_p1 + 1
                            Else
                                PRB_caOvr_p1 = PRB_caOvr_p1 + 1
                            End If
                        End If
                        If Cells(Data_i, 15).Value = "" Then
                            PRB_Queue_Array(0) = PRB_Queue_Array(0) + 1
                        End If
                        'Checking for On Hold
                        If status = "SLAHOLD" Or Left(UCase(status), 7) = "WAITING" Or UCase(status) = "PENDING" Then
                            PRB_OnHold_Array(0) = PRB_OnHold_Array(0) + 1
                        End If
                        If CStr(age_of_tkt) <> "" And age_of_tkt > 0 And closed_date = "" Then
                            If age_of_tkt <= 15 Then
                                PRB_Aging_Array(0, 0) = PRB_Aging_Array(0, 0) + 1
                            ElseIf age_of_tkt <= 30 Then
                                PRB_Aging_Array(1, 0) = PRB_Aging_Array(1, 0) + 1
                            ElseIf age_of_tkt <= 45 Then
                                PRB_Aging_Array(2, 0) = PRB_Aging_Array(2, 0) + 1
                            ElseIf age_of_tkt <= 60 Then
                                PRB_Aging_Array(3, 0) = PRB_Aging_Array(3, 0) + 1
                            ElseIf age_of_tkt > 60 Then
                                PRB_Aging_Array(4, 0) = PRB_Aging_Array(4, 0) + 1
                            End If
                        End If
                    Case 2
                        PRB_Efrt_p2 = PRB_Efrt_p2 + effort
                        If resolution = "Y" Then
                            PRB_ResSLA_p2 = PRB_ResSLA_p2 + 1
                        End If
                        If open_date < today Then
                            PRB_opBal_p2 = PRB_opBal_p2 + 1
                            If closed_date = "" Then
                                PRB_caOvr_p2 = PRB_caOvr_p2 + 1
                            ElseIf closed_date = today Or closed_date <> "" Then
                                PRB_Rsolv_p2 = PRB_Rsolv_p2 + 1
                            Else
                                PRB_caOvr_p2 = PRB_caOvr_p2 + 1
                            End If
                        Else
                            If rspnd = "Y" Then
                                PRB_RspSLA_p2 = PRB_RspSLA_p2 + 1
                                PRB_Rspnd_p2 = PRB_Rspnd_p2 + 1
                            ElseIf rspnd = "N" Then
                                PRB_Rspnd_p2 = PRB_Rspnd_p2 + 1
                            End If
                            
                            PRB_Recv_p2 = PRB_Recv_p2 + 1
                            If closed_date = "" Then
                                PRB_caOvr_p2 = PRB_caOvr_p2 + 1
                            ElseIf closed_date = today Or closed_date <> "" Then
                                PRB_Rsolv_p2 = PRB_Rsolv_p2 + 1
                            Else
                                PRB_caOvr_p2 = PRB_caOvr_p2 + 1
                            End If
                        End If
                        If Cells(Data_i, 15).Value = "" Then
                            PRB_Queue_Array(1) = PRB_Queue_Array(1) + 1
                        End If
                        'Checking for On Hold
                        If status = "SLAHOLD" Or Left(UCase(status), 7) = "WAITING" Or UCase(status) = "PENDING" Then
                            PRB_OnHold_Array(1) = PRB_OnHold_Array(1) + 1
                        End If
                        If CStr(age_of_tkt) <> "" And age_of_tkt > 0 And closed_date = "" Then
                            If age_of_tkt <= 15 Then
                                PRB_Aging_Array(0, 1) = PRB_Aging_Array(0, 1) + 1
                            ElseIf age_of_tkt <= 30 Then
                                PRB_Aging_Array(1, 1) = PRB_Aging_Array(1, 1) + 1
                            ElseIf age_of_tkt <= 45 Then
                                PRB_Aging_Array(2, 1) = PRB_Aging_Array(2, 1) + 1
                            ElseIf age_of_tkt <= 60 Then
                                PRB_Aging_Array(3, 1) = PRB_Aging_Array(3, 1) + 1
                            ElseIf age_of_tkt > 60 Then
                                PRB_Aging_Array(4, 1) = PRB_Aging_Array(4, 1) + 1
                            End If
                        End If
                    Case 3
                        PRB_Efrt_p3 = PRB_Efrt_p3 + effort
                        If resolution = "Y" Then
                            PRB_ResSLA_p3 = PRB_ResSLA_p3 + 1
                        End If
                        If open_date < today Then
                            PRB_opBal_p3 = PRB_opBal_p3 + 1
                            If closed_date = "" Then
                                PRB_caOvr_p3 = PRB_caOvr_p3 + 1
                            ElseIf closed_date = today Or closed_date <> "" Then
                                PRB_Rsolv_p3 = PRB_Rsolv_p3 + 1
                            Else
                                PRB_caOvr_p3 = PRB_caOvr_p3 + 1
                            End If
                        Else
                            If rspnd = "Y" Then
                                PRB_RspSLA_p3 = PRB_RspSLA_p3 + 1
                                PRB_Rspnd_p3 = PRB_Rspnd_p3 + 1
                            ElseIf rspnd = "N" Then
                                PRB_Rspnd_p3 = PRB_Rspnd_p3 + 1
                            End If
                            
                            PRB_Recv_p3 = PRB_Recv_p3 + 1
                            If closed_date = "" Then
                                PRB_caOvr_p3 = PRB_caOvr_p3 + 1
                            ElseIf closed_date = today Or closed_date <> "" Then
                                PRB_Rsolv_p3 = PRB_Rsolv_p3 + 1
                            Else
                                PRB_caOvr_p3 = PRB_caOvr_p3 + 1
                            End If
                        End If
                        If Cells(Data_i, 15).Value = "" Then
                            PRB_Queue_Array(2) = PRB_Queue_Array(2) + 1
                        End If
                        'Checking for On Hold
                        If status = "SLAHOLD" Or Left(UCase(status), 7) = "WAITING" Or UCase(status) = "PENDING" Then
                            PRB_OnHold_Array(2) = PRB_OnHold_Array(2) + 1
                        End If
                        If CStr(age_of_tkt) <> "" And age_of_tkt > 0 And closed_date = "" Then
                            If age_of_tkt <= 15 Then
                                PRB_Aging_Array(0, 2) = PRB_Aging_Array(0, 2) + 1
                            ElseIf age_of_tkt <= 30 Then
                                PRB_Aging_Array(1, 2) = PRB_Aging_Array(1, 2) + 1
                            ElseIf age_of_tkt <= 45 Then
                                PRB_Aging_Array(2, 2) = PRB_Aging_Array(2, 2) + 1
                            ElseIf age_of_tkt <= 60 Then
                                PRB_Aging_Array(3, 2) = PRB_Aging_Array(3, 2) + 1
                            ElseIf age_of_tkt > 60 Then
                                PRB_Aging_Array(4, 2) = PRB_Aging_Array(4, 2) + 1
                            End If
                        End If
                    Case 4
                        PRB_Efrt_p4 = PRB_Efrt_p4 + effort
                        If resolution = "Y" Then
                            PRB_ResSLA_p4 = PRB_ResSLA_p4 + 1
                        End If
                        If open_date < today Then
                            PRB_opBal_p4 = PRB_opBal_p4 + 1
                            If closed_date = "" Then
                                PRB_caOvr_p4 = PRB_caOvr_p4 + 1
                            ElseIf closed_date = today Or closed_date <> "" Then
                                PRB_Rsolv_p4 = PRB_Rsolv_p4 + 1
                            Else
                                PRB_caOvr_p4 = PRB_caOvr_p4 + 1
                            End If
                        Else
                            If rspnd = "Y" Then
                                PRB_RspSLA_p4 = PRB_RspSLA_p4 + 1
                                PRB_Rspnd_p4 = PRB_Rspnd_p4 + 1
                            ElseIf rspnd = "N" Then
                                PRB_Rspnd_p4 = PRB_Rspnd_p4 + 1
                            End If
                            
                            PRB_Recv_p4 = PRB_Recv_p4 + 1
                            If closed_date = "" Then
                                PRB_caOvr_p4 = PRB_caOvr_p4 + 1
                            ElseIf closed_date = today Or closed_date <> "" Then
                                PRB_Rsolv_p4 = PRB_Rsolv_p4 + 1
                            Else
                                PRB_caOvr_p4 = PRB_caOvr_p4 + 1
                            End If
                        End If
                        If Cells(Data_i, 15).Value = "" Then
                            PRB_Queue_Array(3) = PRB_Queue_Array(3) + 1
                        End If
                        'Checking for On Hold
                        If status = "SLAHOLD" Or Left(UCase(status), 7) = "WAITING" Or UCase(status) = "PENDING" Then
                            PRB_OnHold_Array(3) = PRB_OnHold_Array(3) + 1
                        End If
                        If CStr(age_of_tkt) <> "" And age_of_tkt > 0 And closed_date = "" Then
                            If age_of_tkt <= 15 Then
                                PRB_Aging_Array(0, 3) = PRB_Aging_Array(0, 3) + 1
                            ElseIf age_of_tkt <= 30 Then
                                PRB_Aging_Array(1, 3) = PRB_Aging_Array(1, 3) + 1
                            ElseIf age_of_tkt <= 45 Then
                                PRB_Aging_Array(2, 3) = PRB_Aging_Array(2, 3) + 1
                            ElseIf age_of_tkt <= 60 Then
                                PRB_Aging_Array(3, 3) = PRB_Aging_Array(3, 3) + 1
                            ElseIf age_of_tkt > 60 Then
                                PRB_Aging_Array(4, 3) = PRB_Aging_Array(4, 3) + 1
                            End If
                        End If
                    Case 5
                        PRB_Efrt_p5 = PRB_Efrt_p5 + effort
                        If resolution = "Y" Then
                            PRB_ResSLA_p5 = PRB_ResSLA_p5 + 1
                        End If
                        If open_date < today Then
                            PRB_opBal_p5 = PRB_opBal_p5 + 1
                            If closed_date = "" Then
                                PRB_caOvr_p5 = PRB_caOvr_p5 + 1
                            ElseIf closed_date = today Or closed_date <> "" Then
                                PRB_Rsolv_p5 = PRB_Rsolv_p5 + 1
                            Else
                                PRB_caOvr_p5 = PRB_caOvr_p5 + 1
                            End If
                        Else
                            If rspnd = "Y" Then
                                PRB_RspSLA_p5 = PRB_RspSLA_p5 + 1
                                PRB_Rspnd_p5 = PRB_Rspnd_p5 + 1
                            ElseIf rspnd = "N" Then
                                PRB_Rspnd_p5 = PRB_Rspnd_p5 + 1
                            End If
                            
                            PRB_Recv_p5 = PRB_Recv_p5 + 1
                            If closed_date = "" Then
                                PRB_caOvr_p5 = PRB_caOvr_p5 + 1
                            ElseIf closed_date = today Or closed_date <> "" Then
                                PRB_Rsolv_p5 = PRB_Rsolv_p5 + 1
                            Else
                                PRB_caOvr_p5 = PRB_caOvr_p5 + 1
                            End If
                        End If
                        If Cells(Data_i, 15).Value = "" Then
                            PRB_Queue_Array(4) = PRB_Queue_Array(4) + 1
                        End If
                        'Checking for On Hold
                        If status = "SLAHOLD" Or Left(UCase(status), 7) = "WAITING" Or UCase(status) = "PENDING" Then
                            PRB_OnHold_Array(4) = PRB_OnHold_Array(4) + 1
                        End If
                        If CStr(age_of_tkt) <> "" And age_of_tkt > 0 And closed_date = "" Then
                            If age_of_tkt <= 15 Then
                                PRB_Aging_Array(0, 4) = PRB_Aging_Array(0, 4) + 1
                            ElseIf age_of_tkt <= 30 Then
                                PRB_Aging_Array(1, 4) = PRB_Aging_Array(1, 4) + 1
                            ElseIf age_of_tkt <= 45 Then
                                PRB_Aging_Array(2, 4) = PRB_Aging_Array(2, 4) + 1
                            ElseIf age_of_tkt <= 60 Then
                                PRB_Aging_Array(3, 4) = PRB_Aging_Array(3, 4) + 1
                            ElseIf age_of_tkt > 60 Then
                                PRB_Aging_Array(4, 4) = PRB_Aging_Array(4, 4) + 1
                            End If
                        End If
                End Select
            Case "CHG"
                If CHG_Dict.Exists(element) Then
                    CHG_Dict.Item(element) = CHG_Dict.Item(element) + 1
                Else
                    CHG_Dict.Add element, 1
                End If
                Select Case prty
                    Case 1
                        CHG_Efrt_p1 = CHG_Efrt_p1 + effort
'***********************************************************************************
'********************** WHAT TO DO IN CASE RspSLA & ResSLA *************************
'***********************************************************************************
                        If open_date < today Then
                            CHG_opBal_p1 = CHG_opBal_p1 + 1
                            If closed_date = "" Then
                                CHG_caOvr_p1 = CHG_caOvr_p1 + 1
                            ElseIf closed_date = today Or closed_date <> "" Then
                                CHG_Rsolv_p1 = CHG_Rsolv_p1 + 1
                            Else
                                CHG_caOvr_p1 = CHG_caOvr_p1 + 1
                            End If
                        Else
                            CHG_Recv_p1 = CHG_Recv_p1 + 1
                            If closed_date = "" Then
                                CHG_caOvr_p1 = CHG_caOvr_p1 + 1
                            ElseIf closed_date = today Or closed_date <> "" Then
                                CHG_Rsolv_p1 = CHG_Rsolv_p1 + 1
                            Else
                                CHG_caOvr_p1 = CHG_caOvr_p1 + 1
                            End If
                        End If
                        If Cells(Data_i, 15).Value = "" Then
                            CHG_Queue_Array(0) = CHG_Queue_Array(0) + 1
                        End If
                        'Checking for On Hold
                        If status = "SLAHOLD" Or Left(UCase(status), 7) = "WAITING" Or UCase(status) = "PENDING" Then
                            CHG_OnHold_Array(0) = CHG_OnHold_Array(0) + 1
                        End If
                        If CStr(age_of_tkt) <> "" And age_of_tkt > 0 And closed_date = "" Then
                            If age_of_tkt <= 15 Then
                                CHG_Aging_Array(0, 0) = CHG_Aging_Array(0, 0) + 1
                            ElseIf age_of_tkt <= 30 Then
                                CHG_Aging_Array(1, 0) = CHG_Aging_Array(1, 0) + 1
                            ElseIf age_of_tkt <= 45 Then
                                CHG_Aging_Array(2, 0) = CHG_Aging_Array(2, 0) + 1
                            ElseIf age_of_tkt <= 60 Then
                                CHG_Aging_Array(3, 0) = CHG_Aging_Array(3, 0) + 1
                            ElseIf age_of_tkt > 60 Then
                                CHG_Aging_Array(4, 0) = CHG_Aging_Array(4, 0) + 1
                            End If
                        End If
                    Case 2
                        CHG_Efrt_p2 = CHG_Efrt_p2 + effort
                        If open_date < today Then
                            CHG_opBal_p2 = CHG_opBal_p2 + 1
                            If closed_date = "" Then
                                CHG_caOvr_p2 = CHG_caOvr_p2 + 1
                            ElseIf closed_date = today Or closed_date <> "" Then
                                CHG_Rsolv_p2 = CHG_Rsolv_p2 + 1
                            Else
                                CHG_caOvr_p2 = CHG_caOvr_p2 + 1
                            End If
                        Else
                            CHG_Recv_p2 = CHG_Recv_p2 + 1
                            If closed_date = "" Then
                                CHG_caOvr_p2 = CHG_caOvr_p2 + 1
                            ElseIf closed_date = today Or closed_date <> "" Then
                                CHG_Rsolv_p2 = CHG_Rsolv_p2 + 1
                            Else
                                CHG_caOvr_p2 = CHG_caOvr_p2 + 1
                            End If
                        End If
                        If Cells(Data_i, 15).Value = "" Then
                            CHG_Queue_Array(1) = CHG_Queue_Array(1) + 1
                        End If
                        'Checking for On Hold
                        If status = "SLAHOLD" Or Left(UCase(status), 7) = "WAITING" Or UCase(status) = "PENDING" Then
                            CHG_OnHold_Array(1) = CHG_OnHold_Array(1) + 1
                        End If
                        If CStr(age_of_tkt) <> "" And age_of_tkt > 0 And closed_date = "" Then
                            If age_of_tkt <= 15 Then
                                CHG_Aging_Array(0, 1) = CHG_Aging_Array(0, 1) + 1
                            ElseIf age_of_tkt <= 30 Then
                                CHG_Aging_Array(1, 1) = CHG_Aging_Array(1, 1) + 1
                            ElseIf age_of_tkt <= 45 Then
                                CHG_Aging_Array(2, 1) = CHG_Aging_Array(2, 1) + 1
                            ElseIf age_of_tkt <= 60 Then
                                CHG_Aging_Array(3, 1) = CHG_Aging_Array(3, 1) + 1
                            ElseIf age_of_tkt > 60 Then
                                CHG_Aging_Array(4, 1) = CHG_Aging_Array(4, 1) + 1
                            End If
                        End If
                    Case 3
                        CHG_Efrt_p3 = CHG_Efrt_p3 + effort
                        If open_date < today Then
                            CHG_opBal_p3 = CHG_opBal_p3 + 1
                            If closed_date = "" Then
                                CHG_caOvr_p3 = CHG_caOvr_p3 + 1
                            ElseIf closed_date = today Or closed_date <> "" Then
                                CHG_Rsolv_p3 = CHG_Rsolv_p3 + 1
                            Else
                                CHG_caOvr_p3 = CHG_caOvr_p3 + 1
                            End If
                        Else
                            CHG_Recv_p3 = CHG_Recv_p3 + 1
                            If closed_date = "" Then
                                CHG_caOvr_p3 = CHG_caOvr_p3 + 1
                            ElseIf closed_date = today Or closed_date <> "" Then
                                CHG_Rsolv_p3 = CHG_Rsolv_p3 + 1
                            Else
                                CHG_caOvr_p3 = CHG_caOvr_p3 + 1
                            End If
                        End If
                        If Cells(Data_i, 15).Value = "" Then
                            CHG_Queue_Array(2) = CHG_Queue_Array(2) + 1
                        End If
                        'Checking for On Hold
                        If status = "SLAHOLD" Or Left(UCase(status), 7) = "WAITING" Or UCase(status) = "PENDING" Then
                            CHG_OnHold_Array(2) = CHG_OnHold_Array(2) + 1
                        End If
                        If CStr(age_of_tkt) <> "" And age_of_tkt > 0 And closed_date = "" Then
                            If age_of_tkt <= 15 Then
                                CHG_Aging_Array(0, 2) = CHG_Aging_Array(0, 2) + 1
                            ElseIf age_of_tkt <= 30 Then
                                CHG_Aging_Array(1, 2) = CHG_Aging_Array(1, 2) + 1
                            ElseIf age_of_tkt <= 45 Then
                                CHG_Aging_Array(2, 2) = CHG_Aging_Array(2, 2) + 1
                            ElseIf age_of_tkt <= 60 Then
                                CHG_Aging_Array(3, 2) = CHG_Aging_Array(3, 2) + 1
                            ElseIf age_of_tkt > 60 Then
                                CHG_Aging_Array(4, 2) = CHG_Aging_Array(4, 2) + 1
                            End If
                        End If
                    Case 4
                        CHG_Efrt_p4 = CHG_Efrt_p4 + effort
                        If open_date < today Then
                            CHG_opBal_p4 = CHG_opBal_p4 + 1
                            If closed_date = "" Then
                                CHG_caOvr_p4 = CHG_caOvr_p4 + 1
                            ElseIf closed_date = today Or closed_date <> "" Then
                                CHG_Rsolv_p4 = CHG_Rsolv_p4 + 1
                            Else
                                CHG_caOvr_p4 = CHG_caOvr_p4 + 1
                            End If
                        Else
                            CHG_Recv_p4 = CHG_Recv_p4 + 1
                            If closed_date = "" Then
                                CHG_caOvr_p4 = CHG_caOvr_p4 + 1
                            ElseIf closed_date = today Or closed_date <> "" Then
                                CHG_Rsolv_p4 = CHG_Rsolv_p4 + 1
                            Else
                                CHG_caOvr_p4 = CHG_caOvr_p4 + 1
                            End If
                        End If
                        If Cells(Data_i, 15).Value = "" Then
                            CHG_Queue_Array(3) = CHG_Queue_Array(3) + 1
                        End If
                        'Checking for On Hold
                        If status = "SLAHOLD" Or Left(UCase(status), 7) = "WAITING" Or UCase(status) = "PENDING" Then
                            CHG_OnHold_Array(3) = CHG_OnHold_Array(3) + 1
                        End If
                        If CStr(age_of_tkt) <> "" And age_of_tkt > 0 And closed_date = "" Then
                            If age_of_tkt <= 15 Then
                                CHG_Aging_Array(0, 3) = CHG_Aging_Array(0, 3) + 1
                            ElseIf age_of_tkt <= 30 Then
                                CHG_Aging_Array(1, 3) = CHG_Aging_Array(1, 3) + 1
                            ElseIf age_of_tkt <= 45 Then
                                CHG_Aging_Array(2, 3) = CHG_Aging_Array(2, 3) + 1
                            ElseIf age_of_tkt <= 60 Then
                                CHG_Aging_Array(3, 3) = CHG_Aging_Array(3, 3) + 1
                            ElseIf age_of_tkt > 0 Then
                                CHG_Aging_Array(4, 3) = CHG_Aging_Array(4, 3) + 1
                            End If
                        End If
                    Case 5
                        CHG_Efrt_p5 = CHG_Efrt_p5 + effort
                        If open_date < today Then
                            CHG_opBal_p5 = CHG_opBal_p5 + 1
                            If closed_date = "" Then
                                CHG_caOvr_p5 = CHG_caOvr_p5 + 1
                            ElseIf closed_date = today Or closed_date <> "" Then
                                CHG_Rsolv_p5 = CHG_Rsolv_p5 + 1
                            Else
                                CHG_caOvr_p5 = CHG_caOvr_p5 + 1
                            End If
                        Else
                            CHG_Recv_p5 = CHG_Recv_p5 + 1
                            If closed_date = "" Then
                                CHG_caOvr_p5 = CHG_caOvr_p5 + 1
                            ElseIf closed_date = today Or closed_date <> "" Then
                                CHG_Rsolv_p5 = CHG_Rsolv_p5 + 1
                            Else
                                CHG_caOvr_p5 = CHG_caOvr_p5 + 1
                            End If
                        End If
                        If Cells(Data_i, 15).Value = "" Then
                            CHG_Queue_Array(4) = CHG_Queue_Array(4) + 1
                        End If
                        'Checking for On Hold
                        If status = "SLAHOLD" Or Left(UCase(status), 7) = "WAITING" Or UCase(status) = "PENDING" Then
                            CHG_OnHold_Array(4) = CHG_OnHold_Array(4) + 1
                        End If
                        If CStr(age_of_tkt) <> "" And age_of_tkt > 0 And closed_date = "" Then
                            If age_of_tkt <= 15 Then
                                CHG_Aging_Array(0, 4) = CHG_Aging_Array(0, 4) + 1
                            ElseIf age_of_tkt <= 30 Then
                                CHG_Aging_Array(1, 4) = CHG_Aging_Array(1, 4) + 1
                            ElseIf age_of_tkt <= 45 Then
                                CHG_Aging_Array(2, 4) = CHG_Aging_Array(2, 4) + 1
                            ElseIf age_of_tkt <= 60 Then
                                CHG_Aging_Array(3, 4) = CHG_Aging_Array(3, 4) + 1
                            ElseIf age_of_tkt > 60 Then
                                CHG_Aging_Array(4, 4) = CHG_Aging_Array(4, 4) + 1
                            End If
                        End If
                End Select
            Case "EVT"
                If EVT_Dict.Exists(element) Then
                    EVT_Dict.Item(element) = EVT_Dict.Item(element) + 1
                Else
                    EVT_Dict.Add element, 1
                End If
                Select Case prty
                    Case 1
                        EVT_Efrt_p1 = EVT_Efrt_p1 + effort
                        'Respond and Respond SLA is calculating here for P1 incident
                        
                        
                        'Resolution is calculated here for P1 incident
                        If resolution = "Y" Then
                            EVT_ResSLA_p1 = EVT_ResSLA_p1 + 1
                        End If
    
                        If open_date < today Then
                            EVT_opBal_p1 = EVT_opBal_p1 + 1
                            If closed_date = "" Then
                                EVT_caOvr_p1 = EVT_caOvr_p1 + 1
                            ElseIf closed_date = today Or closed_date <> "" Then
                                EVT_Rsolv_p1 = EVT_Rsolv_p1 + 1
                            Else
                                EVT_caOvr_p1 = EVT_caOvr_p1 + 1
                            End If
                        Else
                            If rspnd = "Y" Then
                                EVT_RspSLA_p1 = EVT_RspSLA_p1 + 1
                                EVT_Rspnd_p1 = EVT_Rspnd_p1 + 1
                            ElseIf rspnd = "N" Then
                                EVT_Rspnd_p1 = EVT_Rspnd_p1 + 1
                            End If
                            
                            EVT_Recv_p1 = EVT_Recv_p1 + 1
                            If closed_date = "" Then
                                EVT_caOvr_p1 = EVT_caOvr_p1 + 1
                            ElseIf closed_date = today Or closed_date <> "" Then
                                EVT_Rsolv_p1 = EVT_Rsolv_p1 + 1
                            Else
                                EVT_caOvr_p1 = EVT_caOvr_p1 + 1
                            End If
                        End If
                        If Cells(Data_i, 15).Value = "" Then
                            EVT_Queue_Array(0) = EVT_Queue_Array(0) + 1
                        End If
                        'Checking for On Hold
                        If status = "SLAHOLD" Or Left(UCase(status), 7) = "WAITING" Or UCase(status) = "PENDING" Then
                            EVT_OnHold_Array(0) = EVT_OnHold_Array(0) + 1
                        End If
                        If CStr(age_of_tkt) <> "" And age_of_tkt > 0 And closed_date = "" Then
                            If age_of_tkt <= 15 Then
                                EVT_Aging_Array(0, 0) = EVT_Aging_Array(0, 0) + 1
                            ElseIf age_of_tkt <= 30 Then
                                EVT_Aging_Array(1, 0) = EVT_Aging_Array(1, 0) + 1
                            ElseIf age_of_tkt <= 45 Then
                                EVT_Aging_Array(2, 0) = EVT_Aging_Array(2, 0) + 1
                            ElseIf age_of_tkt <= 60 Then
                                EVT_Aging_Array(3, 0) = EVT_Aging_Array(3, 0) + 1
                            ElseIf age_of_tkt > 60 Then
                                EVT_Aging_Array(4, 0) = EVT_Aging_Array(4, 0) + 1
                            End If
                        End If
                    Case 2
                        EVT_Efrt_p2 = EVT_Efrt_p2 + effort
                        If resolution = "Y" Then
                            EVT_ResSLA_p2 = EVT_ResSLA_p2 + 1
                        End If
                        If open_date < today Then
                            EVT_opBal_p2 = EVT_opBal_p2 + 1
                            If closed_date = "" Then
                                EVT_caOvr_p2 = EVT_caOvr_p2 + 1
                            ElseIf closed_date = today Or closed_date <> "" Then
                                EVT_Rsolv_p2 = EVT_Rsolv_p2 + 1
                            Else
                                EVT_caOvr_p2 = EVT_caOvr_p2 + 1
                            End If
                        Else
                            If rspnd = "Y" Then
                                EVT_RspSLA_p2 = EVT_RspSLA_p2 + 1
                                EVT_Rspnd_p2 = EVT_Rspnd_p2 + 1
                            ElseIf rspnd = "N" Then
                                EVT_Rspnd_p2 = EVT_Rspnd_p2 + 1
                            End If
                            
                            EVT_Recv_p2 = EVT_Recv_p2 + 1
                            If closed_date = "" Then
                                EVT_caOvr_p2 = EVT_caOvr_p2 + 1
                            ElseIf closed_date = today Or closed_date <> "" Then
                                EVT_Rsolv_p2 = EVT_Rsolv_p2 + 1
                            Else
                                EVT_caOvr_p2 = EVT_caOvr_p2 + 1
                            End If
                        End If
                        If Cells(Data_i, 15).Value = "" Then
                            EVT_Queue_Array(1) = EVT_Queue_Array(1) + 1
                        End If
                        'Checking for On Hold
                        If status = "SLAHOLD" Or Left(UCase(status), 7) = "WAITING" Or UCase(status) = "PENDING" Then
                            EVT_OnHold_Array(1) = EVT_OnHold_Array(1) + 1
                        End If
                        If CStr(age_of_tkt) <> "" And age_of_tkt > 0 And closed_date = "" Then
                            If age_of_tkt <= 15 Then
                                EVT_Aging_Array(0, 1) = EVT_Aging_Array(0, 1) + 1
                            ElseIf age_of_tkt <= 30 Then
                                EVT_Aging_Array(1, 1) = EVT_Aging_Array(1, 1) + 1
                            ElseIf age_of_tkt <= 45 Then
                                EVT_Aging_Array(2, 1) = EVT_Aging_Array(2, 1) + 1
                            ElseIf age_of_tkt <= 60 Then
                                EVT_Aging_Array(3, 1) = EVT_Aging_Array(3, 1) + 1
                            ElseIf age_of_tkt > 60 Then
                                EVT_Aging_Array(4, 1) = EVT_Aging_Array(4, 1) + 1
                            End If
                        End If
                    Case 3
                        EVT_Efrt_p3 = EVT_Efrt_p3 + effort
                        If resolution = "Y" Then
                            EVT_ResSLA_p3 = EVT_ResSLA_p3 + 1
                        End If
                        If open_date < today Then
                            EVT_opBal_p3 = EVT_opBal_p3 + 1
                            If closed_date = "" Then
                                EVT_caOvr_p3 = EVT_caOvr_p3 + 1
                            ElseIf closed_date = today Or closed_date <> "" Then
                                EVT_Rsolv_p3 = EVT_Rsolv_p3 + 1
                            Else
                                EVT_caOvr_p3 = EVT_caOvr_p3 + 1
                            End If
                        Else
                            If rspnd = "Y" Then
                                EVT_RspSLA_p3 = EVT_RspSLA_p3 + 1
                                EVT_Rspnd_p3 = EVT_Rspnd_p3 + 1
                            ElseIf rspnd = "N" Then
                                EVT_Rspnd_p3 = EVT_Rspnd_p3 + 1
                            End If
                            
                            EVT_Recv_p3 = EVT_Recv_p3 + 1
                            If closed_date = "" Then
                                EVT_caOvr_p3 = EVT_caOvr_p3 + 1
                            ElseIf closed_date = today Or closed_date <> "" Then
                                EVT_Rsolv_p3 = EVT_Rsolv_p3 + 1
                            Else
                                EVT_caOvr_p3 = EVT_caOvr_p3 + 1
                            End If
                        End If
                        If Cells(Data_i, 15).Value = "" Then
                            EVT_Queue_Array(2) = EVT_Queue_Array(2) + 1
                        End If
                        'Checking for On Hold
                        If status = "SLAHOLD" Or Left(UCase(status), 7) = "WAITING" Or UCase(status) = "PENDING" Then
                            EVT_OnHold_Array(2) = EVT_OnHold_Array(2) + 1
                        End If
                        If CStr(age_of_tkt) <> "" And age_of_tkt > 0 And closed_date = "" Then
                            If age_of_tkt <= 15 Then
                                EVT_Aging_Array(0, 2) = EVT_Aging_Array(0, 2) + 1
                            ElseIf age_of_tkt <= 30 Then
                                EVT_Aging_Array(1, 2) = EVT_Aging_Array(1, 2) + 1
                            ElseIf age_of_tkt <= 45 Then
                                EVT_Aging_Array(2, 2) = EVT_Aging_Array(2, 2) + 1
                            ElseIf age_of_tkt <= 60 Then
                                EVT_Aging_Array(3, 2) = EVT_Aging_Array(3, 2) + 1
                            ElseIf age_of_tkt > 60 Then
                                EVT_Aging_Array(4, 2) = EVT_Aging_Array(4, 2) + 1
                            End If
                        End If
                    Case 4
                        EVT_Efrt_p4 = EVT_Efrt_p4 + effort
                        If resolution = "Y" Then
                            EVT_ResSLA_p4 = INC_ResSLA_p4 + 1
                        End If
                        If open_date < today Then
                            EVT_opBal_p4 = EVT_opBal_p4 + 1
                            If closed_date = "" Then
                                EVT_caOvr_p4 = EVT_caOvr_p4 + 1
                            ElseIf closed_date = today Or closed_date <> "" Then
                                EVT_Rsolv_p4 = EVT_Rsolv_p4 + 1
                            Else
                                EVT_caOvr_p4 = EVT_caOvr_p4 + 1
                            End If
                        Else
                            If rspnd = "Y" Then
                                EVT_RspSLA_p4 = EVT_RspSLA_p4 + 1
                                EVT_Rspnd_p4 = EVT_Rspnd_p4 + 1
                            ElseIf rspnd = "N" Then
                                EVT_Rspnd_p4 = EVT_Rspnd_p4 + 1
                            End If
                            
                            EVT_Recv_p4 = EVT_Recv_p4 + 1
                            If closed_date = "" Then
                                EVT_caOvr_p4 = EVT_caOvr_p4 + 1
                            ElseIf closed_date = today Or closed_date <> "" Then
                                EVT_Rsolv_p4 = EVT_Rsolv_p4 + 1
                            Else
                                EVT_caOvr_p4 = EVT_caOvr_p4 + 1
                            End If
                        End If
                        If Cells(Data_i, 15).Value = "" Then
                            EVT_Queue_Array(3) = EVT_Queue_Array(3) + 1
                        End If
                        'Checking for On Hold
                        If status = "SLAHOLD" Or Left(UCase(status), 7) = "WAITING" Or UCase(status) = "PENDING" Then
                            EVT_OnHold_Array(3) = EVT_OnHold_Array(3) + 1
                        End If
                        If CStr(age_of_tkt) <> "" And age_of_tkt > 0 And closed_date = "" Then
                            If age_of_tkt <= 15 Then
                                EVT_Aging_Array(0, 3) = EVT_Aging_Array(0, 3) + 1
                            ElseIf age_of_tkt <= 30 Then
                                EVT_Aging_Array(1, 3) = EVT_Aging_Array(1, 3) + 1
                            ElseIf age_of_tkt <= 45 Then
                                EVT_Aging_Array(2, 3) = EVT_Aging_Array(2, 3) + 1
                            ElseIf age_of_tkt <= 60 Then
                                EVT_Aging_Array(3, 3) = EVT_Aging_Array(3, 3) + 1
                            ElseIf age_of_tkt > 0 Then
                                EVT_Aging_Array(4, 3) = EVT_Aging_Array(4, 3) + 1
                            End If
                        End If
                    Case 5
                        EVT_Efrt_p5 = EVT_Efrt_p5 + effort
                        If resolution = "Y" Then
                            EVT_ResSLA_p5 = EVT_ResSLA_p5 + 1
                        End If
                        If open_date < today Then
                            EVT_opBal_p5 = EVT_opBal_p5 + 1
                            If closed_date = "" Then
                                EVT_caOvr_p5 = EVT_caOvr_p5 + 1
                            ElseIf closed_date = today Or closed_date <> "" Then
                                EVT_Rsolv_p5 = EVT_Rsolv_p5 + 1
                            Else
                                EVT_caOvr_p5 = EVT_caOvr_p5 + 1
                            End If
                        Else
                            If rspnd = "Y" Then
                                EVT_RspSLA_p5 = EVT_RspSLA_p5 + 1
                                EVT_Rspnd_p5 = EVT_Rspnd_p5 + 1
                            ElseIf rspnd = "N" Then
                                EVT_Rspnd_p5 = EVT_Rspnd_p5 + 1
                            End If
                            
                            EVT_Recv_p5 = EVT_Recv_p5 + 1
                            If closed_date = "" Then
                                EVT_caOvr_p5 = EVT_caOvr_p5 + 1
                            ElseIf closed_date = today Or closed_date <> "" Then
                                EVT_Rsolv_p5 = EVT_Rsolv_p5 + 1
                            Else
                                EVT_caOvr_p5 = EVT_caOvr_p5 + 1
                            End If
                        End If
                        If Cells(Data_i, 15).Value = "" Then
                            EVT_Queue_Array(4) = EVT_Queue_Array(4) + 1
                        End If
                        'Checking for On Hold
                        If status = "SLAHOLD" Or Left(UCase(status), 7) = "WAITING" Or UCase(status) = "PENDING" Then
                            EVT_OnHold_Array(4) = EVT_OnHold_Array(4) + 1
                        End If
                        If CStr(age_of_tkt) <> "" And age_of_tkt > 0 And closed_date = "" Then
                            If age_of_tkt <= 15 Then
                                EVT_Aging_Array(0, 4) = EVT_Aging_Array(0, 4) + 1
                            ElseIf age_of_tkt <= 30 Then
                                EVT_Aging_Array(1, 4) = EVT_Aging_Array(1, 4) + 1
                            ElseIf age_of_tkt <= 45 Then
                                EVT_Aging_Array(2, 4) = EVT_Aging_Array(2, 4) + 1
                            ElseIf age_of_tkt <= 60 Then
                                EVT_Aging_Array(3, 4) = EVT_Aging_Array(3, 4) + 1
                            ElseIf age_of_tkt > 60 Then
                                EVT_Aging_Array(4, 4) = EVT_Aging_Array(4, 4) + 1
                            End If
                        End If
                End Select
            End Select
    Next Data_i
    
    INC_TeamSize = INC_Dict.Count
    SRQ_TeamSize = SRQ_Dict.Count
    CHG_TeamSize = CHG_Dict.Count
    PRB_TeamSize = PRB_Dict.Count
    EVT_TeamSize = EVT_Dict.Count
    Res_count_Dict_TeamSize = Res_count_Dict.Count
    
    
    If Res_count_Dict.Exists("") Then
        Res_count_Dict_TeamSize = Res_count_Dict.Count - 1
    End If
    If INC_Dict.Exists("") Then
        INC_TeamSize = INC_Dict.Count - 1
    End If
    If SRQ_Dict.Exists("") Then
        SRQ_TeamSize = SRQ_Dict.Count - 1
    End If
    If CHG_Dict.Exists("") Then
        CHG_TeamSize = CHG_Dict.Count - 1
    End If
    If PRB_Dict.Exists("") Then
        PRB_TeamSize = PRB_Dict.Count - 1
    End If
    If EVT_Dict.Exists("") Then
        EVT_TeamSize = EVT_Dict.Count - 1
    End If
   
Sheets(sheetDbd).Select

Cells(10, 2).Value = Date - 1
Cells(10, 8).Value = Res_count_Dict_TeamSize
    
'--------------- printing INCIDENT value to the respective cells ------------

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
Cells(22, 13).Value = INC_Efrt_p4 / 60
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

'--------------- printing SERVICE Request value to the respective cells ------------

Cells(10, 16).Value = SRQ_opBal_p1
Cells(11, 16).Value = SRQ_Recv_p1
Cells(12, 16).Value = SRQ_Rspnd_p1
Cells(13, 16).Value = SRQ_Rsolv_p1
Cells(14, 16).Value = SRQ_caOvr_p1
Cells(22, 16).Value = SRQ_Efrt_p1 / 60
Cells(24, 16).Value = SRQ_RspSLA_p1
Cells(25, 16).Value = SRQ_ResSLA_p1

Cells(10, 17).Value = SRQ_opBal_p2
Cells(11, 17).Value = SRQ_Recv_p2
Cells(12, 17).Value = SRQ_Rspnd_p2
Cells(13, 17).Value = SRQ_Rsolv_p2
Cells(14, 17).Value = SRQ_caOvr_p2
Cells(22, 17).Value = SRQ_Efrt_p2 / 60
Cells(24, 17).Value = SRQ_RspSLA_p2
Cells(25, 17).Value = SRQ_ResSLA_p2

Cells(10, 18).Value = SRQ_opBal_p3
Cells(11, 18).Value = SRQ_Recv_p3
Cells(12, 18).Value = SRQ_Rspnd_p3
Cells(13, 18).Value = SRQ_Rsolv_p3
Cells(14, 18).Value = SRQ_caOvr_p3
Cells(22, 18).Value = SRQ_Efrt_p3 / 60
Cells(24, 18).Value = SRQ_RspSLA_p3
Cells(25, 18).Value = SRQ_ResSLA_p3

Cells(10, 19).Value = SRQ_opBal_p4
Cells(11, 19).Value = SRQ_Recv_p4
Cells(12, 19).Value = SRQ_Rspnd_p4
Cells(13, 19).Value = SRQ_Rsolv_p4
Cells(14, 19).Value = SRQ_caOvr_p4
Cells(22, 19).Value = SRQ_Efrt_p4 / 60
Cells(24, 19).Value = SRQ_RspSLA_p4
Cells(25, 19).Value = SRQ_ResSLA_p4

Cells(10, 20).Value = SRQ_opBal_p5
Cells(11, 20).Value = SRQ_Recv_p5
Cells(12, 20).Value = SRQ_Rspnd_p5
Cells(13, 20).Value = SRQ_Rsolv_p5
Cells(14, 20).Value = SRQ_caOvr_p5
Cells(22, 20).Value = SRQ_Efrt_p5 / 60
Cells(24, 20).Value = SRQ_RspSLA_p5
Cells(25, 20).Value = SRQ_ResSLA_p5
Range("P15:T15").Value = SRQ_OnHold_Array

Range("P16:T16").Value = SRQ_Queue_Array

Range("P17:T21").Value = SRQ_Aging_Array

Cells(23, 16).Value = SRQ_TeamSize

'--------------- printing CHANGE Request value to the respective cells ------------

Cells(10, 22).Value = CHG_opBal_p1
Cells(11, 22).Value = CHG_Recv_p1
'Cells(12, 22).Value = NULL
Cells(13, 22).Value = CHG_Rsolv_p1
Cells(14, 22).Value = CHG_caOvr_p1
Cells(22, 22).Value = CHG_Efrt_p1 / 60
Cells(24, 22).Value = CHG_RspSLA_p1
Cells(25, 22).Value = CHG_ResSLA_p1

Cells(10, 23).Value = CHG_opBal_p2
Cells(11, 23).Value = CHG_Recv_p2
'Cells(12, 23).Value = NULL
Cells(13, 23).Value = CHG_Rsolv_p2
Cells(14, 23).Value = CHG_caOvr_p2
Cells(22, 23).Value = CHG_Efrt_p2 / 60
Cells(24, 23).Value = CHG_RspSLA_p2
Cells(25, 23).Value = CHG_ResSLA_p2

Cells(10, 24).Value = CHG_opBal_p3
Cells(11, 24).Value = CHG_Recv_p3
'Cells(12, 24).Value = NULL
Cells(13, 24).Value = CHG_Rsolv_p3
Cells(14, 24).Value = CHG_caOvr_p3
Cells(22, 24).Value = CHG_Efrt_p3 / 60
Cells(24, 24).Value = CHG_RspSLA_p3
Cells(25, 24).Value = CHG_ResSLA_p3

Cells(10, 25).Value = CHG_opBal_p4
Cells(11, 25).Value = CHG_Recv_p4
'Cells(12, 25).Value = NULL
Cells(13, 25).Value = CHG_Rsolv_p4
Cells(14, 25).Value = CHG_caOvr_p4
Cells(22, 25).Value = CHG_Efrt_p4 / 60
Cells(24, 25).Value = CHG_RspSLA_p4
Cells(25, 25).Value = CHG_ResSLA_p4

Cells(10, 26).Value = CHG_opBal_p5
Cells(11, 26).Value = CHG_Recv_p5
'Cells(12, 26).Value = NULL
Cells(13, 26).Value = CHG_Rsolv_p5
Cells(14, 26).Value = CHG_caOvr_p5
Cells(22, 26).Value = CHG_Efrt_p5 / 60
Cells(24, 26).Value = CHG_RspSLA_p5
Cells(25, 26).Value = CHG_ResSLA_p5
Range("V15:Z15").Value = CHG_OnHold_Array

Range("V16:Z16").Value = CHG_Queue_Array
Range("V17:Z21").Value = CHG_Aging_Array
Cells(23, 22).Value = CHG_TeamSize

'--------------- printing PROBLEM Ticket value to the respective cells ------------

Cells(10, 28).Value = PRB_opBal_p1
Cells(11, 28).Value = PRB_Recv_p1
Cells(12, 28).Value = PRB_Rspnd_p1
Cells(13, 28).Value = PRB_Rsolv_p1
Cells(14, 28).Value = PRB_caOvr_p1
Cells(22, 28).Value = PRB_Efrt_p1 / 60
Cells(24, 28).Value = PRB_RspSLA_p1
Cells(25, 28).Value = PRB_ResSLA_p1

Cells(10, 29).Value = PRB_opBal_p2
Cells(11, 29).Value = PRB_Recv_p2
Cells(12, 29).Value = PRB_Rspnd_p2
Cells(13, 29).Value = PRB_Rsolv_p2
Cells(14, 29).Value = PRB_caOvr_p2
Cells(22, 29).Value = PRB_Efrt_p2 / 60
Cells(24, 29).Value = PRB_RspSLA_p2
Cells(25, 29).Value = PRB_ResSLA_p2

Cells(10, 30).Value = PRB_opBal_p3
Cells(11, 30).Value = PRB_Recv_p3
Cells(12, 30).Value = PRB_Rspnd_p3
Cells(13, 30).Value = PRB_Rsolv_p3
Cells(14, 30).Value = PRB_caOvr_p3
Cells(22, 30).Value = PRB_Efrt_p3 / 60
Cells(24, 30).Value = PRB_RspSLA_p3
Cells(25, 30).Value = PRB_ResSLA_p3

Cells(10, 31).Value = PRB_opBal_p4
Cells(11, 31).Value = PRB_Recv_p4
Cells(12, 31).Value = PRB_Rspnd_p4
Cells(13, 31).Value = PRB_Rsolv_p4
Cells(14, 31).Value = PRB_caOvr_p4
Cells(22, 31).Value = PRB_Efrt_p4 / 60
Cells(24, 31).Value = PRB_RspSLA_p4
Cells(25, 31).Value = PRB_ResSLA_p4

Cells(10, 32).Value = PRB_opBal_p5
Cells(11, 32).Value = PRB_Recv_p5
Cells(12, 32).Value = PRB_Rspnd_p5
Cells(13, 32).Value = PRB_Rsolv_p5
Cells(14, 32).Value = PRB_caOvr_p5
Cells(22, 32).Value = PRB_Efrt_p5 / 60
Cells(24, 32).Value = PRB_RspSLA_p5
Cells(25, 32).Value = PRB_ResSLA_p5
Range("AB15:AF15").Value = PRB_OnHold_Array

Range("AB16:AF16").Value = PRB_Queue_Array
Range("AB17:AF21").Value = PRB_Aging_Array
Cells(23, 28).Value = PRB_TeamSize


'--------------- printing Event value to the respective cells ------------

Cells(10, 34).Value = EVT_opBal_p1
Cells(11, 34).Value = EVT_Recv_p1
Cells(12, 34).Value = EVT_Rspnd_p1
Cells(13, 34).Value = EVT_Rsolv_p1
Cells(14, 34).Value = EVT_caOvr_p1
Cells(22, 34).Value = EVT_Efrt_p1 / 60
Cells(24, 34).Value = EVT_RspSLA_p1
Cells(25, 34).Value = EVT_ResSLA_p1

Cells(10, 35).Value = EVT_opBal_p2
Cells(11, 35).Value = EVT_Recv_p2
Cells(12, 35).Value = EVT_Rspnd_p2
Cells(13, 35).Value = EVT_Rsolv_p2
Cells(14, 35).Value = EVT_caOvr_p2
Cells(22, 35).Value = EVT_Efrt_p2 / 60
Cells(24, 35).Value = EVT_RspSLA_p2
Cells(25, 35).Value = EVT_ResSLA_p2

Cells(10, 36).Value = EVT_opBal_p3
Cells(11, 36).Value = EVT_Recv_p3
Cells(12, 36).Value = EVT_Rspnd_p3
Cells(13, 36).Value = EVT_Rsolv_p3
Cells(14, 36).Value = EVT_caOvr_p3
Cells(22, 36).Value = EVT_Efrt_p3 / 60
Cells(24, 36).Value = EVT_RspSLA_p3
Cells(25, 36).Value = EVT_ResSLA_p3

Cells(10, 37).Value = EVT_opBal_p4
Cells(11, 37).Value = EVT_Recv_p4
Cells(12, 37).Value = EVT_Rspnd_p4
Cells(13, 37).Value = EVT_Rsolv_p4
Cells(14, 37).Value = EVT_caOvr_p4
Cells(22, 37).Value = (EVT_Efrt_p4 / 60)
Cells(24, 37).Value = EVT_RspSLA_p4
Cells(25, 37).Value = EVT_ResSLA_p4

Cells(10, 38).Value = EVT_opBal_p5
Cells(11, 38).Value = EVT_Recv_p5
Cells(12, 38).Value = EVT_Rspnd_p5
Cells(13, 38).Value = EVT_Rsolv_p5
Cells(14, 38).Value = EVT_caOvr_p5
Cells(22, 38).Value = EVT_Efrt_p5 / 60
Cells(24, 38).Value = EVT_RspSLA_p5
Cells(25, 38).Value = EVT_ResSLA_p5

Range("AH15:AL15").Value = EVT_OnHold_Array

Range("AH16:AL16").Value = EVT_Queue_Array

Range("AH17:AL21").Value = EVT_Aging_Array

Cells(23, 34).Value = EVT_TeamSize

End Sub



