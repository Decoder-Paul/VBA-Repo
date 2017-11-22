Sub Merge()
'========================================================================================================
' Merge
' -------------------------------------------------------------------------------------------------------
' Purpose   :   To Merge the record of two different sheet by Activity_no
'
' Author    :   Subhankar Paul
' Date      :   22nd November, 2017
' Notes     :
' Parameters:   N/A
' Returns   :   N/A
' ---------------------------------------------------------------
' Revision History
'
'========================================================================================================
    Dim WB As Workbook
    Dim WS_output As Worksheet
    Dim WS_log As Worksheet
    Dim WS_lab As Worksheet
    Set WB = ActiveWorkbook
    Set WS_output = WB.Sheets("Final Output")
    Set WS_log = WB.Sheets("Worklog record")
    Set WS_lab = WB.Sheets("Labtrans record")
    
    Dim lab_lro As Integer
    Dim log_lro As Integer
    log_lro = WS_log.Cells(Row.Count, "B").End(xlUp).Row
    lab_lro = WS_lab.Cells(Rows.Count, "B").End(xlUp).Row
    
    Dim i As Integer    'Iterator for picking activity no. from lab sheet
    Dim log_i As Integer 'matching row index in log sheet by activity no. of lab sheet
    Dim lab As Integer  'counting repeataion of same activity no. in lab sheet
    Dim log As Integer  'counting repeataion of same activity no. in log sheet
    Dim j As Integer    'iterator for output sheet
    
    Dim site As String
    Dim srq As String
    Dim ssr_stat As String
    Dim ssr_tts As Double
    Dim act_no As String
    Dim actvty_stat As String
    
    Dim actstart As Date
    Dim actfinsh As Date
    Dim actlabh As Double
    
    Dim staff_id As String
    Dim staff_name As String
    j = 2
    For i = 2 To lab_lro
        act_no = WS_lab.Cells(i, 2).Value
        
        lab = 0
        While act_no = WS_lab.Cells(i + lab + 1, 2).Value
            lab = lab + 1
        Wend
        j = j + lab
        log_i = Application.Match(act_no, WS_lab.Columns(2), 0)
        If Not IsError(log_i) Then
            j = j + 1
            log = 0
            While act_no = WS_lab.Cells(log_i + log + 1, 2).Value
                log = log + 1
            Wend
        End If
        WS_lab.Range("E" & log_i & ":H" & log_i + log).Copy
        WS_output.Range("M" & j).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Application.CutCopyMode = False
        j = j + log
    Next i
End Sub
