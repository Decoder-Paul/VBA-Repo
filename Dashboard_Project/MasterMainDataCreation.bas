Attribute VB_Name = "MasterMainDataCreation"
 Public sCFilNam As String
 Sub Master_MainData_Page()
 
 'Remember time when macro starts
StartTime = Timer
Application.ScreenUpdating = False
Application.DisplayAlerts = False
 
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
  Dim lroInc As Long
  
  Dim iCounter1 As Integer
  Dim iCounter2 As Integer
  Dim iCountF As Integer
  Dim i As Long
  Dim iFold As Integer
    
  Dim WB As Workbook
  Dim WS_DaIn As Worksheet
  Dim WS_sIn As Worksheet
  Dim WS_Rep As Worksheet

  Dim wkb As Workbook
  Dim sht As Worksheet
  
  Dim NewBook As Workbook
  Dim sPath As String ' Default Path of the File
  Dim sPath1 As String
  Dim sDate As Variant
  Dim fold As Variant
  Dim SubFolders As Variant
  
  Dim Fname As String
  Dim lro As Long
  Dim lro1 As Long
  
' Initial Assignment of value to variables
  sDaIn = "MainDataInf"
  sIn = "MainDataBackup"
  sP1 = "MainData"
  
  If fSheetExists(sDaIn) = False Then
    
    Call pSheetCreate(sDaIn)
  
  Else
    Sheets(sDaIn).Activate
    Sheets(sDaIn).Range("a1").Select
    Cells.Select
    Selection.Cells.Clear
  
  End If
  
  If fSheetExists(sIn) = False Then
    
    Call pSheetCreate(sIn)
  
  Else
    Sheets(sIn).Activate
    Sheets(sIn).Range("a1").Select
    Cells.Select
    Selection.Cells.Clear
  
  End If
  
  Set WB = ActiveWorkbook
  Set WS_DaIn = WB.Sheets(sDaIn)
  Set WS_sIn = WB.Sheets(sIn)
  
'Collecting the current folder path details getting the path to pass it to the Browser
  sPath1 = Application.ActiveWorkbook.Path
  sPath = sPath1
  sPath = Left(sPath, InStrRev(sPath, "\") - 1)
  sPath = sPath & "\DashBoardBackup\"
  
  sCFilNam = Application.ActiveWorkbook.Name
  sCFilNam = Right(sCFilNam, Len(sCFilNam) - InStrRev(sCFilNam, "\"))

'Checking if Sheet DataInfo is available if Available then delete all the details in the sheet.

    
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
    Set objFolder1 = oFSO1.GetFolder(sPath)
    'Get the folder object
    'Set oFolder = oFSO.GetFolder(sPath)
    
    iCounter1 = 0
    iCounter2 = 2
    iFold = 2
    
        For Each objSubFolder1 In objFolder1.SubFolders
           If objSubFolder1 = sPath & "MainData backup" Then
                WS_DaIn.Cells(iFold + 1, 3) = objSubFolder1.Name
                'print folder path
                WS_DaIn.Cells(iFold + 1, 4) = objSubFolder1.Path
                iFold = iFold + 1
            End If
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
                        WS_DaIn.Cells(iCounter1, 7).Value = WS_DaIn.Cells(iCounter2, 4).Value
                        WS_DaIn.Cells(iCounter1, 8).Value = Mid(WS_DaIn.Cells(iCounter1, 1).Value, 9, 11)
        
                    End If
                    iCounter1 = iCounter1 + 1
            Next oFile
            
        Next iCounter2
       
        lro1 = WS_DaIn.Cells(WS_DaIn.Rows.Count, "d").End(xlUp).Row
        
        Sheets(sDaIn).Range(Cells(2, 8), Cells(lro1, 8)).NumberFormat = "[$-14009]dd-mm-yyyy;@"
        
        Call pMasterMainDataClean
        
    Sheets("MainDataBackup").Activate
    lroInc = WS_sIn.Cells(WS_sIn.Rows.Count, "a").End(xlUp).Row
    
    
 For i = iCountF + 1 To lro1
    'Opening the Sheet
    Workbooks.Open (WS_DaIn.Cells(i, 2).Value)
    Workbooks(WS_DaIn.Cells(i, 1).Value).Activate
    'if sheet page 1 is avalable or not select
    If fSheetExists(sP1) = False Then
        Workbooks(WS_DaIn.Cells(icount, 1).Value).Close
        MsgBox "Page 1 Sheet is missing in InFlow File."
    End
    Else
       If lroInc = 1 Then
            
            Sheets(sP1).Activate
            Sheets(sP1).Range("a1").Select
            lro = Cells(Rows.Count, "A").End(xlUp).Row
            lco = Cells(4, Columns.Count).End(xlToRight).Column
            Range(Cells(1, 1), Cells(lro, lco)).Copy
            Workbooks(sCFilNam).Activate
            Sheets(sIn).Activate
            Sheets(sIn).Range("a1").PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            Sheets(sIn).Range("a1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            Sheets(sIn).Range("a3").Copy
            Sheets(sIn).Range("Q3").PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            Sheets(sIn).Cells(3, 17).Value = "Date"
            Sheets(sDaIn).Activate
            Cells(i, 8).Copy
            Sheets(sIn).Activate
            Sheets(sIn).Range(Cells(4, 17), Cells(lro, 17)).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            Sheets(sIn).Range(Cells(4, 17), Cells(lro, 17)).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            Application.CutCopyMode = False
            Workbooks(WS_DaIn.Cells(i, 1).Value).Close
        Else
            Sheets(sP1).Activate
            Sheets(sP1).Range("a1").Select
            lro = Cells(Rows.Count, "A").End(xlUp).Row
            lco = Cells(1, Columns.Count).End(xlToRight).Column
            Range(Cells(4, 1), Cells(lro, lco)).Copy
            Workbooks(sCFilNam).Activate
            Sheets(sIn).Activate
            Sheets(sIn).Cells(lroInc + 1, 1).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            Sheets(sIn).Cells(lroInc + 1, 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            Sheets(sDaIn).Activate
            Cells(i, 8).Copy
            Sheets(sIn).Activate
            Sheets(sIn).Range(Cells(lroInc + 1, 17), Cells(lroInc + lro - 3, 17)).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            Sheets(sIn).Range(Cells(lroInc + 1, 17), Cells(lroInc + lro - 3, 17)).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            Workbooks(WS_DaIn.Cells(i, 1).Value).Close
        
       End If
    End If
    lroInc = WS_sIn.Cells(WS_sIn.Rows.Count, "a").End(xlUp).Row
    
Next i

With Selection
    .Columns.AutoFit
End With

For Each BI In Array(xlEdgeTop, xlEdgeLeft, xlEdgeBottom, xlEdgeRight, xlInsideHorizontal, xlInsideVertical)
    With Range(Cells(4, 17), Cells(lroInc, 17)).Borders(BI)
         .Weight = xlThin
         .Color = RGB(148, 138, 84)
    End With
 Next BI
 
Dim sNam As String
Dim sWSNam As String
sPath = Application.ActiveWorkbook.Path

sCFilNam = Application.ActiveWorkbook.Name
sCFilNam = Right(sCFilNam, Len(sCFilNam) - InStrRev(sCFilNam, "\"))
  
 'Adding New Workbook
Set wkb = Workbooks.Add
'Saving the Workbook
sNam = sPath & "\" & "MainDatabackup " & ".xlsx"
wkb.SaveAs sNam, AccessMode:=xlExclusive, ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges
sWSNam = "MainDatabackup " & ".xlsx"
 
 Windows(sCFilNam).Activate
Sheets("MainDataBackup").Select
Sheets("MainDataBackup").Move Before:=Workbooks(sWSNam).Sheets(1)
  
Workbooks(sWSNam).Save
Workbooks(sWSNam).Close

Sheets("MainDataInf").Delete
 
 ' Determine how many seconds this code will take to run
  SecondsElapsed = Round(Timer - StartTime, 2)
'Notify user in seconds
  MsgBox "This code ran successfully in " & SecondsElapsed & " seconds", vbInformation

    
 End Sub
    

