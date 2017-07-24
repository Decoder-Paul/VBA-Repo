Attribute VB_Name = "Calls"
Public Function fSheetExists(sheetToFind As String) As Boolean
'========================================================================================================
' fSheetExists
' -------------------------------------------------------------------------------------------------------
' Purpose of this Function : To check if a sheet is existing or not
'
' Author : Prasanna Kumar 03rd October, 2016
' Notes  : N/A
' Parameters : sSheetNameIN - Sheet Name
' Returns : N/A
' ---------------------------------------------------------------
' Revision History
'
'========================================================================================================
 On Error GoTo ErrorHandler
 
  Dim sheet As Worksheet

    fSheetExists = False
    For Each sheet In Worksheets
        If sheetToFind = sheet.Name Then
            fSheetExists = True
            Exit Function
        End If
    Next sheet
    
ErrorHandler:
  'MsgBox "Error: Function Sheet Check"

End Function

 Sub pWorksheetcheck(ByVal sSheetNameINPASS As String)
'========================================================================================================
' pWorkseetCheck
' -------------------------------------------------------------------------------------------------------
' Purpose of this programm : ' To check if a particular sheet is available or not
' if not then call the Pocedure to create a sheet.
'
' Author : Prasanna Kumar 3rd October, 2016
' Notes  : N/A
' Parameters : sSheetNameINPASS - Sheet Name
' Returns : N/A
' ---------------------------------------------------------------
' Revision History
'
'========================================================================================================
 On Error GoTo ErrorHandler
  
  Dim iCounter As Integer
  Dim bExists As Boolean
  
' A loop for Total count of sheets and check if the above sheet is available or not.
  For iCounter = 1 To Worksheets.Count
    If Worksheets(iCounter).Name = sSheetNameINPASS Then
       Sheets(sSheetNameINPASS).Visible = True
       'Boolean operation making it to True
       bExists = True
       Exit For
    End If
  Next iCounter

  If Not bExists Then
  
  MsgBox "sSheetNameINPASS is Missing or Renamed"
  
  End If
  
ErrorHandler:
  'MsgBox "Error: Worksheet check"

End Sub


Sub pCloseApp()

On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    
    'Now you hide later Delete
    If fSheetExists("AssignFilename") = True Then
      Sheets("AssignFilename").Activate
      Sheets("AssignFilename").Range("a1").Select
      Sheets("AssignFilename").Visible = xlSheetVeryHidden

    End If
    
    If fSheetExists("SDH") = True Then
      Sheets("SDH").Activate
      Sheets("SDH").Range("a1").Select
      Sheets("SDH").Visible = xlSheetVeryHidden
    End If
    
    If fSheetExists("Master_Price_List") = True Then
      Sheets("Master_Price_List").Activate
      Sheets("Master_Price_List").Range("a1").Select
      Sheets("Master_Price_List").Visible = xlSheetVeryHidden

    End If
    
    If fSheetExists("Data") = True Then
      Sheets("Data").Activate
      Sheets("Data").Range("a1").Select
      Sheets("data").Visible = xlSheetVeryHidden
    End If
    
    If fSheetExists("SWSS Adjustment") = True Then
      Sheets("SWSS Adjustment").Activate
      Sheets("SWSS Adjustment").Range("a1").Select
      Sheets("SWSS Adjustment").Visible = xlSheetVeryHidden
    End If
ErrorHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub

