Attribute VB_Name = "MainProgram"
Public sCFilNam As String
Public StartTime As Double
Public SecondsElapsed As Double
Public dateOfAnalysis As Date

Option Explicit

Sub pMain()

'Remember time when macro starts
StartTime = Timer
Application.ScreenUpdating = False
Application.DisplayAlerts = False
dateOfAnalysis = Date - 1

Calls.pOpenApp

Call pDataFromFolder
Calls.pCloseApp

Application.ScreenUpdating = True
Application.DisplayAlerts = True

' Determine how many seconds this code will take to run
  SecondsElapsed = Round(Timer - StartTime, 2)
'Notify user in seconds
  MsgBox "This code ran successfully in " & SecondsElapsed & " seconds", vbInformation

End Sub

Sub pDataFromFolder()
' ===================================================================================================
' pDataFromFolder
' --------------------------------------------------------------------------------------------------
' Purpose of the Programm : To extract the dump file into the Data1 sheet and other folder details
'
' Author : Mathews Jacob 6th February, 2017
' Notes  : N/A
' Parameters : N/A
' Returns : N/A
' ---------------------------------------------------------------------------------------------------
' Revision History
' ====================================================================================================
' Variable declaration for various Purpose
' On Error GoTo ErrorHandler
 
  Dim oFSO As Object, oFSO1 As Object
  Dim oFolder As Object, objFolder1 As Object, objSubFolder1 As Object
  Dim oFile As Object, oFile1 As Object
    
  Dim sDaIn As String
  Dim sDa As String
  Dim sSTR As String
  Dim sIn As String
  Dim sP1 As String
  Dim sPa As String
  Dim flag As Integer
  
  Dim iCounter1 As Integer
  Dim iCounter2 As Integer
  Dim iCountF As Integer
  Dim i As Long
  Dim iFold As Integer
    
  Dim WB As Workbook
  Dim WS_DaIn As Worksheet
  Dim WS_Rep As Worksheet

  Dim NewBook As Workbook
  Dim sPath As String ' Default Path of the File
  Dim sPath1 As String
  Dim sDate As Variant
  
  Dim Fname As String
  Dim lro As Long
  
' Initial Assignment of value to variables
  sDaIn = "DataInf"
  sIn = "Incident"
  sP1 = "Page 1"
  sDa = "MainData"
  sPa = "REP"
  
  Set WB = ActiveWorkbook
  Set WS_DaIn = WB.Sheets(sDaIn)
  Set WS_Rep = WB.Sheets(sPa)
  
'Collecting the current folder path details getting the path to pass it to the Browser
  sPath1 = Application.ActiveWorkbook.Path
  sPath = sPath1
  sPath = Left(sPath, InStrRev(sPath, "\") - 1)
  
  sCFilNam = Application.ActiveWorkbook.Name
  sCFilNam = Right(sCFilNam, Len(sCFilNam) - InStrRev(sCFilNam, "\"))

'Checking if Sheet DataInfo is available if Available then delete all the details in the sheet.

  If fSheetExists(sDaIn) = False Then
    
    Call pSheetCreate(sDaIn)
  
  Else
    Sheets(sDaIn).Activate
    Sheets(sDaIn).Range("a1").Select
    Cells.Select
    Selection.Cells.Clear
  
  End If
    
    Sheets(sDaIn).Activate
    Sheets(sDaIn).Range("a1").Select
  
    'Heading to the cells.
 
    WS_DaIn.Cells(1, 1).Value = "File Name"
    WS_DaIn.Cells(1, 2).Value = "Path of File"
    WS_DaIn.Cells(1, 3).Value = "Name of the Folder"
    WS_DaIn.Cells(1, 4).Value = "Path of the Folder"
    WS_DaIn.Cells(1, 5).Value = "Default Folder Path"
    WS_DaIn.Cells(1, 6).Value = "Total files in the Folder"
    WS_DaIn.Cells(2, 4).Value = sPath1
    WS_DaIn.Cells(2, 3).Value = "Main Folder"
  
    'Create an instance of the FileSystemObject
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oFSO1 = CreateObject("Scripting.FileSystemObject")
    Set objFolder1 = oFSO1.GetFolder(sPath1)
    'Get the folder object
    'Set oFolder = oFSO.GetFolder(sPath)
    
    iCounter1 = 0
    iCounter2 = 2
    iFold = 2
    
        For Each objSubFolder1 In objFolder1.SubFolders

           WS_DaIn.Cells(iFold + 1, 3) = objSubFolder1.Name
           'print folder path
           WS_DaIn.Cells(iFold + 1, 4) = objSubFolder1.Path
           iFold = iFold + 1
              
        Next objSubFolder1
    
        iCountF = WS_DaIn.Cells(WS_DaIn.Rows.Count, "d").End(xlUp).Row

        iCounter1 = iCountF + 1
        ' How many folders were exracted accourding to that we need to find out the files
        
        For iCounter2 = 3 To iCountF
          sPath1 = Cells(iCounter2, 4).Value
         
          Set oFolder = oFSO.GetFolder(sPath1)
          
            For Each oFile In oFolder.Files
                sSTR = oFile.Name
                
                'Checking if excel files exist and copy only exile format files
                'Any other files it should not copy.
                
                    If Right(sSTR, 4) = "xlsx" Or Right(sSTR, 3) = "xls" Or Right(sSTR, 4) = "xlsm" Or Right(sSTR, 4) = "CSV" Then
                        WS_DaIn.Cells(iCounter1, 1) = Left(sSTR, Len(sSTR))
                        WS_DaIn.Cells(iCounter1, 2) = oFile.Path
                        WS_DaIn.Cells(iCounter1, 3).Value = WS_DaIn.Cells(iCounter2, 3).Value
                        WS_DaIn.Cells(iCounter1, 4).Value = WS_DaIn.Cells(iCounter2, 4).Value
                        WS_DaIn.Cells(iCounter1, 7).Value = WS_DaIn.Cells(iCounter2, 4).Value & "\Archive\"
                    End If
                    iCounter1 = iCounter1 + 1
            Next oFile
            
        Next iCounter2
        
    'Few Basic formating doing to the data sheet
    WS_DaIn.Range(Cells(2, 1), Cells(2, 6)).Select
    Selection.Columns.AutoFit
    WS_DaIn.Range(Cells(1, 1), Cells(1, 6)).Select
    Selection.Font.Bold = True
    Selection.Interior.Color = RGB(192, 192, 192)

    Set oFSO = Nothing
    
    ' Starting the loop to extract the data from each folder
    ' if there are no folders then is empty row skip extracting with a message.
    lro = WS_DaIn.Cells(WS_DaIn.Rows.Count, "d").End(xlUp).Row

    Sheets(sDaIn).Activate
    Sheets(sDaIn).Range("a1").Select
 
 iCounter1 = 0
 
    If WS_DaIn.Cells(iCountF + 1, 2).Value = "" Then
        Calls.pCloseApp
        MsgBox "Folders not available for extracting data in the root directory.", , "Folder Selection"
        End
    Else
        
        Sheets(sPa).Activate
        WS_Rep.Cells(3, 3).Value = "Clients Included: "
        flag = 1
        
        Sheets(sDaIn).Activate
        For i = iCountF + 1 To lro
            
            iFold = WorksheetFunction.CountIf(WS_DaIn.Range(WS_DaIn.Cells(iCountF + 1, 3), WS_DaIn.Cells(lro, 3)), WS_DaIn.Cells(i, 3).Value)
            
            Select Case Left(UCase(WS_DaIn.Cells(i, 3).Value), 3)

'Checking for NYL
                
                Case "NYL"
                    
                     iCounter1 = iCounter1 + 1
                     If iCounter1 = 1 Then
                        Call pInClean
                     End If
                     
                     Call pNYL(i)
                     
                        If iCounter1 = iFold Then
                            Call pNYLDD
                            Call num_Of_Days
                            Call ticketCount
                                If flag <> 1 Then
                                    WS_Rep.Cells(3, 3).Value = WS_Rep.Cells(3, 3).Value & ", " & "NYL"
                                Else
                                    WS_Rep.Cells(3, 3).Value = WS_Rep.Cells(3, 3).Value & "NYL"
                                    flag = 2
                                End If
                            Call pCopytoMSheet("NYL")
                            iCounter1 = 0
                        End If
                        
'Checking for Master Card

                 Case "MAS"
                    iCounter1 = iCounter1 + 1
                     If iCounter1 = 1 Then
                        Call pInClean
                     End If
                    Call pMAS(i)
                   
                        If iCounter1 = iFold Then
                            Call pMASDD
                            Call num_Of_Days
                            Call ticketCount
                                If flag <> 1 Then
                                    WS_Rep.Cells(3, 3).Value = WS_Rep.Cells(3, 3).Value & ", " & "MASTER CARD EMO"
                                Else
                                    WS_Rep.Cells(3, 3).Value = WS_Rep.Cells(3, 3).Value & "MASTER CARD EMO"
                                    flag = 2
                                End If
                            Call pCopytoMSheet("Master Card EMO")
                            'For Second Master Card
                            Call pMASDD1
                            Call num_Of_Days
                            Call ticketCount
                                If flag <> 1 Then
                                    WS_Rep.Cells(3, 3).Value = WS_Rep.Cells(3, 3).Value & ", " & "MASTER CARD ESM"
                                Else
                                    WS_Rep.Cells(3, 3).Value = WS_Rep.Cells(3, 3).Value & "MASTER CARD ESM"
                                    flag = 2
                                End If
                            Call pCopytoMSheet("Master Card ESM")
                            iCounter1 = 0
                        End If
                        
                        
'Checking for ATIC

                Case "ATI"
                    iCounter1 = iCounter1 + 1
                     If iCounter1 = 1 Then
                        Call pInClean
                     End If
                    Call pMAS(i)
                   
                        If iCounter1 = iFold Then
                            Call pATICDD
                          '  Call num_Of_Days it is inside the program
                            Call ticketCount
                                If flag <> 1 Then
                                    WS_Rep.Cells(3, 3).Value = WS_Rep.Cells(3, 3).Value & ", " & "ATIC"
                                Else
                                    WS_Rep.Cells(3, 3).Value = WS_Rep.Cells(3, 3).Value & "ATIC"
                                    flag = 2
                                End If
                            Call pCopytoMSheet("ATIC")
                            iCounter1 = 0
                        End If
                        
'Checking for IQPC

                 Case "IQP"
                    iCounter1 = iCounter1 + 1
                     If iCounter1 = 1 Then
                        Call pInClean
                     End If
                    Call pMAS(i)
                   
                        If iCounter1 = iFold Then
                            Call pIQPCDD
                          ' Call num_Of_Days it is inside the program
                            Call ticketCount
                                If flag <> 1 Then
                                    WS_Rep.Cells(3, 3).Value = WS_Rep.Cells(3, 3).Value & ", " & "IQPC"
                                Else
                                    WS_Rep.Cells(3, 3).Value = WS_Rep.Cells(3, 3).Value & "IQPC"
                                    flag = 2
                                End If
                            Call pCopytoMSheet("IQPC")
                            iCounter1 = 0
                        End If
                        
'Checking for HERTZ
                  Case "HER"
                    iCounter1 = iCounter1 + 1
                     If iCounter1 = 1 Then
                        Call pInClean
                     End If

                     Call pHER(i)
                     
                        If iCounter1 = iFold Then
                            Call pHERDD
                            Call num_Of_Days
                            Call ticketCount
                                If flag <> 1 Then
                                    WS_Rep.Cells(3, 3).Value = WS_Rep.Cells(3, 3).Value & ", " & "HERTZ"
                                Else
                                    WS_Rep.Cells(3, 3).Value = WS_Rep.Cells(3, 3).Value & "HERTZ"
                                    flag = 2
                                End If
                            Call pCopytoMSheet("Hertz")
                            iCounter1 = 0
                        End If
'Checking for Liberty
                  Case "LIB"
                    iCounter1 = iCounter1 + 1
                     If iCounter1 = 1 Then
                        Call pInClean
                     End If

                     Call pHER(i)
                     
                        If iCounter1 = iFold Then
                            Call pLM
                            Call num_Of_Days
                            Call ticketCount
                                If flag <> 1 Then
                                    WS_Rep.Cells(3, 3).Value = WS_Rep.Cells(3, 3).Value & ", " & "LIBERTY MUTUAL"
                                Else
                                    WS_Rep.Cells(3, 3).Value = WS_Rep.Cells(3, 3).Value & "LIBERTY MUTUAL"
                                    flag = 2
                                End If
                            Call pCopytoMSheet("LM")
                            iCounter1 = 0
                        End If
                End Select
'Checking for Equinix - RIMS
'                Case "RIM"
'                    iCounter1 = iCounter1 + 1
'                    If iCounter1 = 1 Then
'                        Call pInClean
'                    End If
'
'                    Call pEquinix(i)
'
'                    If iCounter1 = iFold Then
'                        Call pMapToDashboard
'                        If flag <> 1 Then
'                            WS_Rep.Cells(3, 3).Value = WS_Rep.Cells(3, 3).Value & ", " & "LIBERTY MUTUAL"
'                        Else
'                            WS_Rep.Cells(3, 3).Value = WS_Rep.Cells(3, 3).Value & "LIBERTY MUTUAL"
'                            flag = 2
'                        End If
'                        Call pCopytoMSheet("RIMS")
'                        iCounter1 = 0
'                    End If
'                End Select
        Next i
    
    End If
    
'Last Combined Calculation

Call pCopytoMSheet("last")
Call ticketCount
Calls.pCopsDB
Calls.pCopyToEmail

End Sub

Sub pCopytoMSheet(sSh As String)

' ===================================================================================================
' pCopytoMSheet
' --------------------------------------------------------------------------------------------------
' Purpose of the Programm : To Copy the Data from the Main Sheet to Master Sheet
' Author : Mathews Jacob 17th February, 2017
' Notes  : N/A
' Parameters : Sheet Name
' Returns : N/A
' ---------------------------------------------------------------------------------------------------
' Revision History
' ====================================================================================================
' Variable declaration for various Purpose
' On Error GoTo ErrorHandler

Dim sMDa As String
Dim sMSh As String
Dim lro As Long
Dim iCh As Byte
Dim sPCl As String
Dim BI As Variant

Dim WB As Workbook
Dim WS_DA As Worksheet
Dim WS_MS As Worksheet

sMDa = "MainData"
sMSh = "MasterSheet"
sPCl = "Project or Cluster"

Set WB = ActiveWorkbook
Set WS_DA = WB.Sheets(sMDa)
Set WS_MS = WB.Sheets(sMSh)

' Combined Report
' This section is called when it is last called
If sSh = "last" Then
    Sheets(sMDa).Activate
    Sheets(sMDa).Range("a1").Select

    lro = WS_DA.Cells(Rows.Count, "A").End(xlUp).Row
    
    If lro > 4 Then
        WS_DA.Range("A4:Z" & (lro)).Clear
    End If

    lro = WS_MS.Cells(Rows.Count, "A").End(xlUp).Row
    Sheets(sMSh).Activate
    WS_MS.Range(Cells(2, 1), Cells(lro, 25)).Copy
    
    Sheets(sMDa).Activate
    WS_DA.Range("a4").PasteSpecial Paste:=xlPasteValues
    
'This is to format all the main data details
'Project dashboard Workbook  selecting then MainData sheet Selecting Formating
    Windows(sCFilNam).Activate
    
    If fSheetExists("MainData") = True Then
        Sheets("MainData").Select
        lro = Cells(Rows.Count, "A").End(xlUp).Row
    'Maindata formating
        Sheets("MainData").Range(Cells(2, 1), Cells(lro, 16)).Select
        With Selection
            .Font.Size = 10
            .Font.Name = "Calibri"
        End With
    
        For Each BI In Array(xlEdgeTop, xlEdgeLeft, xlEdgeBottom, xlEdgeRight, xlInsideHorizontal, xlInsideVertical)
            With Range(Cells(4, 1), Cells(lro, 16)).Borders(BI)
            .Weight = xlThin
            .Color = RGB(148, 138, 84)
            End With
        Next BI
    
        lro = lro + 1
        Sheets("MainData").Range(Cells(lro, 1), Cells(lro + 5000, 30)).Delete Shift:=xlUp
    
        Columns("I:I").Select
        Selection.NumberFormat = "[$-14009]dd-mm-yyyy;@"
        Columns("J:J").Select
        Selection.NumberFormat = "[$-14009]dd-mm-yyyy;@"
        Columns("P:P").Select
        Selection.NumberFormat = "[$-14009]dd-mm-yyyy;@"
    End If
    
    'After this program go to the end of this program without Execuiting the Normal report
    
    GoTo mm
    
End If

' This section is Applicable when each client is called
' Normal Report

    If fSheetExists(sMDa) = False Then
    
        Call pSheetCreate(sMDa)
        iCh = 1
    Else
        iCh = 2
    End If

If iCh = 1 Then

    Sheets(sMDa).Activate
    Sheets(sMDa).Select

    lro = Sheets(sMDa).Cells(Rows.Count, "A").End(xlUp).Row
    Sheets(sMDa).Range(Cells(3, 1), Cells(lro, 25)).Copy
    Sheets(sMSh).Activate
    Sheets(sMSh).Range("a1").PasteSpecial Paste:=xlPasteValues
    Sheets(sMSh).Range("a1").PasteSpecial Paste:=xlPasteFormats

Else

    Sheets(sMDa).Activate
    Sheets(sMDa).Select
    
    lro = Sheets(sMDa).Cells(Rows.Count, "A").End(xlUp).Row
    Sheets(sMDa).Range(Cells(4, 1), Cells(lro, 25)).Copy
    Sheets(sMSh).Activate
    lro = Sheets(sMSh).Cells(Rows.Count, "A").End(xlUp).Row

    Sheets(sMSh).Cells(lro + 1, 1).PasteSpecial Paste:=xlPasteValues
    Sheets(sMSh).Cells(lro + 1, 1).PasteSpecial Paste:=xlPasteFormats
End If

Sheets(sPCl).Activate
ActiveWindow.Zoom = 75
'Making a duplicate Copy of project or cluster with Client Name
Sheets(sPCl).Copy after:=Sheets(sMSh)
ActiveSheet.Name = sSh

ActiveSheet.Cells(10, 3).Value = sSh
ActiveSheet.Cells(10, 5).Value = sSh
mm:


End Sub


