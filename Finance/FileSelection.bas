Attribute VB_Name = "FileSelection"
Option Explicit
Public sCFilNam As String

Sub pFilePreCheck()

'========================================================================================================
' pFilePreCheck
' -------------------------------------------------------------------------------------------------------
' Purpose of this programm :  To Check all the input files,Sheets, Columns if any mismatch.
' Each and every columns and Rows needed for the extraction of the relevant data
' Will also be Checked by calling a Procedure pColumnPreCheck.
'
' Author : Prasanna kumar 3rd October, 2016
' Notes  : N/A
' Parameters : N/A
' Returns : N/A
' ---------------------------------------------------------------
' Revision History
'
'========================================================================================================

On Error GoTo ErrorHandler

Application.ScreenUpdating = False
Application.DisplayAlerts = False

  Dim Wb As Workbook  ' For the current work book.
  Dim WS_DA As Worksheet ' For Data file
  Dim WS_SDH As Worksheet
  Dim sPath_st As String  ' To find out the program path of the current folder.
  Dim sPath_br As String ' To Put the browsed folder selected items.
  Dim sFile_path As String ' To put the File path to open the Location.
  Dim sFileNamClose As String ' To put the value the file name to close the file.
  Dim sAssignFileNam As String  ' To find out the file name, in future if they change.
  Dim sSDH As String  ' To Store the sheet name Installbase Data
  Dim sDa As String  ' To store the sheet name of Data
  Dim rcellda As Range 'To store the Range
  Dim lro As Long ' Dynamic Rows
  Dim lroda As Integer
  Dim iCo As Integer  ' Dynamic Columns
  Dim iresult As Integer ' Dialog box checking
  Dim iCounter As Long ' For loop
  Dim i As Integer
  Dim j As Long
  Dim sSum As Long
  Dim sPc1 As String ' For C1 Details
  Dim sBt As String
  Dim sLDs As String
  Dim sPRl As String
  Dim FSO As Object
  Dim sC1FilePath As String
  Dim B1 As Shape
  Dim B2 As Shape              ' button enabled in this sheet
  Dim bInstallCheck As Boolean
  Dim iMsg As Integer
  Dim aColumn(5) As Integer
  Dim iCount As Integer
  Dim iNum As Integer
  Dim sNBl As String
  Dim sPLs As String
  Dim sTs As String
  
  
  Set FSO = New Scripting.FileSystemObject
  Set Wb = ActiveWorkbook     'Master workbook
  sSDH = "Std Hrs"
  sDa = "Data"
  sPc1 = "PC"
  sBt = "BT"
  sLDs = "Leadership"
  sPRl = "Payroll"
  sNBl = "Non Billable"
  sPLs = " Perdium in Liue of Salary "
  sTs = "Timesheet"
  sAssignFileNam = "Assignment Tracker"

  'Collecting the current folder path details getting the path to pass it to the Browser
  sPath_st = Application.ActiveWorkbook.Path
  sCFilNam = Application.ActiveWorkbook.FullName
  sCFilNam = Right(sCFilNam, Len(sCFilNam) - InStrRev(sAssignFileNam, "\"))
  
       
    If fSheetExists(sSDH) = True Then
      
      Workbooks(sAssignFileNam).Sheets(sSDH).Activate
      Sheets(sSDH).Select
      Cells.Select
      Selection.Cells.Clear
    Else
      Call pWorksheetcheck(sSDH)
    End If
  
   
  'Code to Get/Browse the folder
  'Changes the folder to the program path
  Application.FileDialog(msoFileDialogFilePicker).InitialFileName = sPath_st
  
  'The dialog is displayed to the user
  iresult = Application.FileDialog(msoFileDialogFilePicker).Show
  
  'Checks if user has cancelld the dialog
  If iresult <> 0 Then
      'Store information
      sPath_br = Application.FileDialog(msoFileDialogFilePicker).SelectedItems(1)
      
  Else
      MsgBox "Select a desired File folder to proceed.", , "Folder Selection"
      'Call closeapp
      End
  End If
    
  'Checking if SDH Sheet is available or not in the Pricing Estimator.
  'If not available Create one.
    
  'Call pWorksheetCheck(sSDH)
  Workbooks(sCFilNam).Sheets(sSDH).Activate
  Sheets(sSDH).Select
  Cells.Select
  Selection.Cells.Clear
  
  'Checking if PC Sheet is available or not in the Finance folder.
  'If not available Create one.

  'Call pWorksheetCheck(sPc1)
  Workbooks(sCFilNam).Sheets(sPc1).Activate
  Sheets(sPc1).Select
  Cells.Select
  Selection.Cells.Clear
  
  'Checking if BT Details Sheet is available or not in the Assignment Tracker.
  'If not available Create one.
  
  'Call pWorksheetCheck(sBt)
  Workbooks(sCFilNam).Sheets(sBt).Activate
  Sheets(sBt).Select
  Cells.Select
  Selection.Cells.Clear
   
  Set WS_DA = Wb.Sheets(sLDs)    'Sheet you store all the manipulation data.
    
  Workbooks(sCFilNam).Sheets(sLDs).Activate
  Sheets(sLDs).Select
  Cells(2, 5).Value = sPath_st
  Cells(3, 5).Value = sCFilNam
    
  'Check for the total extracted files
  lroda = Cells(Rows.Count, "A").End(xlUp).Row
    
  'Storing the value in the data sheet so that we don't have to check again and again
  Cells(2, 6).Value = lroda
  If lroda >= 2 Then '2 and above then
     Else
        MsgBox "Required input files are not available in selected folder.", , "Folder Selection"
        'Call pCloseApp
        End
        
  End If
  
  'Data file is checking
     
        'Checking if Std Hrs is available in Master file.

        Workbooks(sAssignFileNam).Activate
               
        If fSheetExists(sSDH) = False Then
          MsgBox "Assignment Tracker.xlsx File - " & sSDH & " Sheet is Missing or Renamed", , "File Selection"
          
          ElseIf fSheetExists(sSDH) > 1 Then
          MsgBox "Program terminates. Multiple [Row Name] exists in Master file"
          
          Call pCloseApp
          End
        Else
            Workbooks(sAssignFileNam).Activate
            Sheets(sSDH).Select
            Cells(1, 1).Select
    
            'Copy the last row and last column
            Range(Cells(6, 2), Cells(10, 3)).Select
            Selection.Copy
    
            'Copy the details from Installbase to
            Workbooks(sCFilNam).Sheets(sSDH).Activate
            Sheets(sSDH).Select
            Cells(2, 27).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            
        'CheckData.pRowPreCheck (sSDH)
       
        End If
        
        Workbooks(sPc1).Activate
        
        If fSheetExists(sPc1) = True Then
            Sheets(sPc1).Select
            'Before copy check if the sheet is in filtered mode.
            Cells(1, 1).Select
    
            'Find the last row and last column
            lro = Cells(Rows.Count, "b").End(xlUp).Row
            iCo = Cells(7, Columns.Count).End(xlToLeft).Column
            Range(Cells(7, 2), Cells(lro, iCo)).Select
            Selection.Copy
    
            'Copy the details from Master file to data sheet
            Workbooks(sAssignFileNam).Sheets(sPc1).Activate
            Sheets(sPc1).Select
            Cells(1, 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            
        End If
       
        ' Closing the file
        Workbooks(sAssignFileNam).Close
                    
MM2:
 'CheckData.pColumnPreCheck (sSDH)
 
 Workbooks(sAssignFileNam).Sheets(sSDH).Activate
 Sheets(sSDH).Range("a1").Select
 
 aColumn(1) = WS_DA.Cells(7, 22).Value
 aColumn(2) = WS_DA.Cells(8, 22).Value
 aColumn(3) = WS_DA.Cells(9, 22).Value
  
'Checking if there is value in the SDH File Requirement sheet.

 Workbooks(sAssignFileNam).Sheets(sSDH).Activate
 Sheets(sSDH).Range("a1").Select
 lro = Cells(Rows.Count, "a").End(xlUp).Row
  For iCounter = 1 To 3
    For j = 1 To lro
      If IsNumeric(WS_SDH.Cells(j, aColumn(iCounter)).Value) Then
        sSum = WS_SDH.Cells(j, aColumn(iCounter)).Value
      Else
        sSum = 0
      End If
      
      If sSum > 0 Then
      Workbooks(sAssignFileNam).Sheets(sSDH).Activate
      Sheets(sSDH).Range("a1").Select
      WS_DA.Cells(29, 21).Value = "Yes"
      WS_DA.Cells(1, 21).Value = "Yes"
      Workbooks(sAssignFileNam).Sheets(sSDH).Activate
      Sheets(sSDH).Range("a1").Select
      
        GoTo RR
      End If
    Next j
  Next iCounter

         MsgBox "File checking Done "
RR:
Application.ScreenUpdating = True
Application.DisplayAlerts = True
        
ErrorHandler:
End Sub

