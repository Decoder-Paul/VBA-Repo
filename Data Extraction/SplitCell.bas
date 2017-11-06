Attribute VB_Name = "SplitCell"
'========================================================================================================
' SplitCells
' -------------------------------------------------------------------------------------------------------
' Purpose   :   To splitting the values in a cell by carriage return and putting the values
'               into consequent rows
' Author    :   Subhankar Paul
' Date      :   6th November, 2017
' Notes     :
' Parameters:   N/A
' Returns   :   N/A
' ---------------------------------------------------------------
' Revision History
'
'========================================================================================================
Dim WB As Workbook
Dim WS_Main As Worksheet
Dim WS_Slave As Worksheet
Sub pCleanDB()
    Dim lro As Long
    
    lro = WB.Sheets("Sheet1").Cells(Rows.Count, "A").End(xlUp).Row
    If lro > 1 Then
        WB.Sheets("Sheet1").Range("A2:H" & lro).ClearContents
    End If
    
    lro = WB.Sheets("Sheet2").Cells(Rows.Count, "A").End(xlUp).Row
    If lro > 1 Then
        WB.Sheets("Sheet2").Range("A2:H" & lro).ClearContents
    End If
    
    lro = WB.Sheets("Sheet3").Cells(Rows.Count, "A").End(xlUp).Row
    If lro > 1 Then
        WB.Sheets("Sheet3").Range("A2:H" & lro).ClearContents
    End If
    
    lro = WB.Sheets("Sheet4").Cells(Rows.Count, "A").End(xlUp).Row
    If lro > 1 Then
        WB.Sheets("Sheet4").Range("A2:H" & lro).ClearContents
    End If
End Sub
Sub SplitCells(ByVal i As Integer)
    Dim comNo As String     'Computer Name
    Dim dType As String     'Device Type
    Dim OSver As String     'OS Version
    Dim TLS As String       'Time OF Last Scan(BFI)
    Dim ip As String        'IP Address
    Dim OS As String        'OS
    Dim LRT As String       'Last Report Time
    Dim j As Integer
    
    'list of apps to be stored in str
    Dim str() As String
    Dim lro As Long
    comNo = WS_Main.Cells(i, 1)
    dType = WS_Main.Cells(i, 2)
    OSver = WS_Main.Cells(i, 3)
    TLS = WS_Main.Cells(i, 4)
    ip = WS_Main.Cells(i, 5)
    OS = WS_Main.Cells(i, 7)
    LRT = WS_Main.Cells(i, 8)
    lro = WS_Slave.Cells(Rows.Count, "F").End(xlUp).Row
    
    If Len(WS_Main.Cells(i, 6).Value) Then
        str = VBA.Split(WS_Main.Cells(i, 6).Value, vbLf)
        For j = 0 To UBound(str)
            WS_Slave.Cells(lro + j + 1, 6).Value = str(j)
            WS_Slave.Cells(lro + j + 1, 1).Value = comNo
            WS_Slave.Cells(lro + j + 1, 2).Value = dType
            WS_Slave.Cells(lro + j + 1, 3).Value = OSver
            WS_Slave.Cells(lro + j + 1, 4).Value = TLS
            WS_Slave.Cells(lro + j + 1, 5).Value = ip
            WS_Slave.Cells(lro + j + 1, 7).Value = OS
            WS_Slave.Cells(lro + j + 1, 8).Value = LRT
        Next j
    End If
    
End Sub
Sub main()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Dim StartTime As Double
    Dim SecondsElapsed As Double
    'Remember time when macro starts
    StartTime = Timer
    
    Dim i As Long
    Dim j As Integer
    
    Dim Data_rowCount As Long
    Set WB = ActiveWorkbook
    Set WS_Main = WB.Sheets("Main")
    Set WS_Slave = WB.Sheets("Sheet1")
    
    Call pCleanDB
    
    Data_rowCount = WS_Main.Cells(Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To Data_rowCount
    'switching sheets for every 8000 rows of Main sheet
        If i Mod 8000 = 0 Then
            j = (i \ 8000) + 1
            Set WS_Slave = WB.Sheets("Sheet" & j)
        End If
        Call SplitCells(i)
    Next
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    SecondsElapsed = Round(Timer - StartTime, 2)
    MsgBox "Code ran Successfully in " & SecondsElapsed & "Seconds", vbInformation
End Sub
