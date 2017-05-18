Attribute VB_Name = "Calls"
Option Explicit

Public Function fSheetExists(sheetToFind As String) As Boolean
'========================================================================================================
' fSheetExists
' -------------------------------------------------------------------------------------------------------
' Purpose of this Function : To check if a sheet is existing or not
'
' Author : Mathews Jacob 16th August, 2016
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
 
End Function

Sub pSheetCreate(sSheetName As String)
'========================================================================================================
' pSheetCreate
' -------------------------------------------------------------------------------------------------------
' Purpose of this programm : To create a new sheet
'
' Author : Mathews Jacob 6th February, 2017
' Notes  : N/A
' Parameters : sSheetNameIN - Sheet Name
' Returns : N/A
' ---------------------------------------------------------------
' Revision History
'
'========================================================================================================
On Error GoTo ErrorHandler
  Dim ws As Worksheet
  
    With ThisWorkbook
        Set ws = .Sheets.Add(after:=.Sheets(.Sheets.Count))
        ws.Name = sSheetName
    End With
    
ErrorHandler:

End Sub

Sub pCloseApp()
'========================================================================================================
' pCloseApp
' -------------------------------------------------------------------------------------------------------
' Purpose of this programm : To create a new sheet
'
' Author : Mathews Jacob 10th February, 2017
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
    
    'Hidding the File
    If fSheetExists("DataInf") = True Then
      Sheets("DataInf").Activate
      Sheets("DataInf").Range("a1").Select
      Sheets("DataInf").Visible = xlSheetVeryHidden
    End If

    If fSheetExists("MainData") = True Then
      Sheets("MainData").Activate
      Sheets("MainData").Range("a1").Select
      Sheets("MainData").Visible = xlSheetHidden
    End If

    If fSheetExists("Incident") = True Then
      Sheets("Incident").Activate
      Sheets("Incident").Range("a1").Select
      Sheets("Incident").Visible = xlSheetVeryHidden
    End If

    If fSheetExists("MasterSheet") = True Then
      Sheets("MasterSheet").Activate
      Sheets("MasterSheet").Range("a1").Select
      Sheets("MasterSheet").Visible = xlSheetVeryHidden
    End If
    
'   If fSheetExists("REP") = True Then
'      Sheets("REP").Activate
'      Sheets("REP").Range("a1").Select
'      Sheets("REP").Visible = xlSheetVeryHidden
'    End If

    Sheets("Project or Cluster").Activate
    Sheets("Project or Cluster").Range("j10").Select
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub

Sub pOpenApp()
'========================================================================================================
' pOpenApp
' -------------------------------------------------------------------------------------------------------
' Purpose of this programm : To create a new sheet
'
' Author : Mathews Jacob 10th February, 2017
' Notes  : N/A
' Parameters : N/A
' Returns : N/A
' ---------------------------------------------------------------
' Revision History
'
'========================================================================================================

'On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Call pVisibleallSheets
    
    If fSheetExists("MasterSheet") = True Then
        Sheets("MasterSheet").Activate
        Sheets("MasterSheet").Range("a1").Select
        Sheets("MasterSheet").Delete
        Call pSheetCreate("MasterSheet")
        Else
        Call pSheetCreate("MasterSheet")
        
    End If
    
    If fSheetExists("Incident") = True Then
        Sheets("Incident").Activate
        Sheets("Incident").Range("a1").Select
        Sheets("Incident").Delete
        Call pSheetCreate("Incident")
        Else
        Call pSheetCreate("Incident")
    End If
    
   If fSheetExists("NYL") = True Then
      Sheets("NYL").Activate
      Sheets("NYL").Range("a1").Select
      Sheets("NYL").Delete
    End If
    
   If fSheetExists("Master Card ESM") = True Then
      Sheets("Master Card ESM").Activate
      Sheets("Master Card ESM").Range("a1").Select
      Sheets("Master Card ESM").Delete
    End If
    
    If fSheetExists("Master Card EMO") = True Then
      Sheets("Master Card EMO").Activate
      Sheets("Master Card EMO").Range("a1").Select
      Sheets("Master Card EMO").Delete
    End If
    
   If fSheetExists("ATIC") = True Then
      Sheets("ATIC").Activate
      Sheets("ATIC").Range("a1").Select
      Sheets("ATIC").Delete
    End If
    
   If fSheetExists("IQPC") = True Then
      Sheets("IQPC").Activate
      Sheets("IQPC").Range("a1").Select
      Sheets("IQPC").Delete
    End If
    
   If fSheetExists("Hertz") = True Then
      Sheets("Hertz").Activate
      Sheets("Hertz").Range("a1").Select
      Sheets("Hertz").Delete
    End If
    
    If fSheetExists("LM") = True Then
      Sheets("LM").Activate
      Sheets("LM").Range("a1").Select
      Sheets("LM").Delete
    End If
    
ErrorHandler:
 
End Sub

Sub pCleanDB()
'This program cleans the Dashboard

    Sheets("Project or Cluster").Range("J10:N25").ClearContents
    Sheets("Project or Cluster").Range("P10:T25").ClearContents
    Sheets("Project or Cluster").Range("V10:Z25").ClearContents
    Sheets("Project or Cluster").Range("AB10:AF25").ClearContents
    Sheets("Project or Cluster").Range("AH10:AL25").ClearContents

End Sub

Sub pInCreate()
  
  'Checking if Incident sheet is available or not
    If fSheetExists(sIn) = False Then
        Call pSheetCreate(sIn)
    End If

End Sub


Sub pVisibleallSheets()

'Unhide all sheets in workbook.

Dim ws As Worksheet
For Each ws In ActiveWorkbook.Worksheets
 ws.Visible = xlSheetVisible
Next ws

End Sub
Sub pInClean()
    'Cleaning
    If fSheetExists("Incident") = True Then
        Sheets("Incident").Activate
        Sheets("Incident").Range("a1").Select
        Cells.Select
        Selection.Cells.Clear
        Cells(1, 1).Value = "a"
    End If
    
End Sub

Sub pCopyToEmail()

Call pCopyChart("Project or Cluster", "REP")
'emailing the template
Call colorIndicator
Call pAutoEmail

End Sub
Sub colorIndicator()

'Purpose of this program is to put the color code in the Project master Report Email format.

Dim WB As Workbook
Dim WS_Rep As Worksheet
Dim WS_Pro As Worksheet
Dim sPa As String
Dim sheetDbd As String
 sPa = "REP"
sheetDbd = "Project or Cluster"

Set WB = ActiveWorkbook
Set WS_Rep = WB.Sheets(sPa)
Set WS_Pro = WB.Sheets(sheetDbd)
WS_Rep.Activate
WS_Pro.Activate

With Worksheets("REP")
If (WS_Pro.Cells(14, 10).Value - WS_Pro.Cells(15, 10).Value) = 0 And (WS_Pro.Cells(14, 11).Value - WS_Pro.Cells(15, 11).Value) = 0 Then
    .Shapes("Oval 1").Fill.ForeColor.RGB = RGB(76, 153, 0)
ElseIf (WS_Pro.Cells(14, 10).Value - WS_Pro.Cells(15, 10).Value) = 0 And (WS_Pro.Cells(14, 11).Value - WS_Pro.Cells(15, 11).Value) > 0 Then
    .Shapes("Oval 1").Fill.ForeColor.RGB = RGB(255, 255, 0)
ElseIf (WS_Pro.Cells(14, 10).Value - WS_Pro.Cells(15, 10).Value) > 0 And (WS_Pro.Cells(14, 11).Value - WS_Pro.Cells(15, 11).Value) = 0 Then
    .Shapes("Oval 1").Fill.ForeColor.RGB = RGB(255, 0, 0)
Else
    .Shapes("Oval 1").Fill.ForeColor.RGB = RGB(255, 0, 0)
End If

End With

End Sub
Sub pCopsDB()

' Purpose of this program is to copy the details into cops dashboard file which goes along with email has an attachment

Dim sPath As String
Dim wkb As Workbook
Dim sNam As String
Dim dNam As String
Dim sWSNam As String
Dim newName As String
Dim Path As String
Dim oldName As String

sPath = Application.ActiveWorkbook.Path

sNam = sPath & "\" & "COPS Dashboard" & ".xlsm"
dNam = sPath & "Backup" & "\COPS DashBoard Backup\" & "COPS Dashboard" & ".xlsm"
Path = sPath & "\"
newName = sPath & "Backup" & "\COPS DashBoard Backup\" & "COPS Dashboard " & dateOfAnalysis & ".xlsm"
    
'Copy COPS Dashboard file in Dashboard folder to COPS DashBoard Backup folder and renaming with date
If Dir(newName) = "" Then
    FileCopy sNam, dNam
    Name dNam As newName
End If

Set wkb = Workbooks.Open(newName)
sWSNam = "COPS Dashboard " & ".xlsm"

'creating sheet1 if not exists in COPS Dashboard
wkb.Activate
If fSheetExists("Sheet1") = False Then
     wkb.Sheets.Add Before:=Worksheets(Worksheets.Count)
End If

'-----In the begining ckecking if any sheets exists in COPS Dashboard file,if exists deletting all the sheet-----
    wkb.Activate
If fSheetExists("Project or Cluster") = True Then
    Sheets("Project or Cluster").Delete
End If
If fSheetExists("MainData") = True Then
    Sheets("MainData").Delete
End If
If fSheetExists("NYL") = True Then
    Sheets("NYL").Delete
End If
If fSheetExists("Master Card EMO") = True Then
    Sheets("Master Card EMO").Delete
End If
If fSheetExists("Master Card ESM") = True Then
    Sheets("Master Card ESM").Delete
End If
If fSheetExists("ATIC") = True Then
    Sheets("ATIC").Delete
End If
If fSheetExists("IQPC") = True Then
    Sheets("IQPC").Delete
End If
If fSheetExists("Hertz") = True Then
    Sheets("Hertz").Delete
End If
If fSheetExists("LM") = True Then
    Sheets("LM").Delete
End If
'-----------------------------------------------------------------------------------------------------------

'copying the sheets in Activeworksheet to the COPS Dashboard
    Windows(sCFilNam).Activate
If fSheetExists("Project or Cluster") = True Then
    Sheets("Project or Cluster").Select
    Sheets("Project or Cluster").Copy Before:=wkb.Sheets("Sheet1")
End If

    Windows(sCFilNam).Activate
If fSheetExists("MainData") = True Then
    Sheets("MainData").Select
    ActiveWindow.Zoom = 75
    Sheets("MainData").Copy Before:=wkb.Sheets("Sheet1")
End If

    Windows(sCFilNam).Activate
If fSheetExists("NYL") = True Then
    Sheets("NYL").Select
    Sheets("NYL").Copy Before:=wkb.Sheets("Sheet1")
End If

    Windows(sCFilNam).Activate
If fSheetExists("Master Card EMO") = True Then
    Sheets("Master Card EMO").Select
    Sheets("Master Card EMO").Copy Before:=wkb.Sheets("Sheet1")
End If

    Windows(sCFilNam).Activate
If fSheetExists("Master Card ESM") = True Then
    Sheets("Master Card ESM").Select
    Sheets("Master Card ESM").Copy Before:=wkb.Sheets("Sheet1")
End If

    Windows(sCFilNam).Activate
If fSheetExists("ATIC") = True Then
    Sheets("ATIC").Select
    Sheets("ATIC").Copy Before:=wkb.Sheets("Sheet1")
End If

    Windows(sCFilNam).Activate
If fSheetExists("IQPC") = True Then
    Sheets("IQPC").Select
    Sheets("IQPC").Copy Before:=wkb.Sheets("Sheet1")
End If

    Windows(sCFilNam).Activate
If fSheetExists("Hertz") = True Then
    Sheets("Hertz").Select
    Sheets("Hertz").Copy Before:=wkb.Sheets("Sheet1")
End If

    Windows(sCFilNam).Activate
If fSheetExists("LM") = True Then
    Sheets("LM").Select
    Sheets("LM").Copy Before:=wkb.Sheets("Sheet1")
End If

'deleting sheet1 from COPS Dashboard
  wkb.Activate
If fSheetExists("Sheet1") = True Then
    Sheets("Sheet1").Delete
End If

'Making default page as Project or Cluster
Sheets("Project or Cluster").Activate
Sheets("Project or Cluster").Range("J10").Select

wkb.Save

wkb.Close
Call pCopyMainData


End Sub

Sub pCopyMainData()

'Purpose Main data into file.
Dim sPath As String
Dim wkb As Workbook
Dim sNam As String
Dim sWSNam As String
Dim lro As Long
Dim BI As Variant

sPath = Application.ActiveWorkbook.Path

'Adding New Workbook
Set wkb = Workbooks.Add
'Saving the Workbook
sNam = sPath & "Backup\MainData backup\" & "MainData " & dateOfAnalysis & ".xlsx"
wkb.SaveAs sNam, AccessMode:=xlExclusive, ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges
sWSNam = "MainData " & dateOfAnalysis & ".xlsx"

    Windows(sCFilNam).Activate
        
    Sheets("MainData").Copy Before:=Workbooks(sWSNam).Sheets(1)


Workbooks(sWSNam).Save
Workbooks(sWSNam).Close

End Sub

Sub num_Of_Days()
'========================================================================================================
' fSheetExists
' -------------------------------------------------------------------------------------------------------
' Purpose of this Function : To count number of days between created and resolved dates for Aging
'
' Author : Shambhavi B M, 27th February, 2017
' Notes  : N/A
' Parameters : N/A
' Returns : N/A
' ---------------------------------------------------------------
' Revision History
'
'========================================================================================================

Dim WB As Workbook
Dim WS_DA As Worksheet
Dim sDa As String
Dim lro As Long
Dim tkt_type As String
Dim opened_date As Long
Dim i As Long

sDa = "MainData"

Set WB = ActiveWorkbook
Set WS_DA = WB.Sheets(sDa)
WB.Sheets(sDa).Activate
Sheets(sDa).Select
lro = Cells(Rows.Count, "A").End(xlUp).Row

Cells(1, 1).Value = dateOfAnalysis
If lro >= 4 Then
    Range(Cells(4, 15), Cells(lro, 15)).Formula = "=IFERROR(IF(P4="""","""",IF(J4="""",DATEDIF(P4,$A$1,""d""),DATEDIF(P4,J4,""d""))),0)"
End If

For i = 4 To lro

'if actual date is empty then we need to Calculate
    If Cells(i, 16).Value = "" Then
        tkt_type = Cells(i, 2).Value
        opened_date = Cells(i, 9).Value
        
        If tkt_type = "INC" And (dateOfAnalysis - opened_date) >= 2 Then
            Cells(i, 15).Value = dateOfAnalysis - opened_date
        Else
            Cells(i, 15).Value = ""
        End If
        
        If (tkt_type = "SRQ" Or tkt_type = "CHG" Or tkt_type = "PRB") And (dateOfAnalysis - opened_date) > 5 Then
            Cells(i, 15).Value = dateOfAnalysis - opened_date
        Else
            Cells(i, 15).Value = ""
        End If
        
    End If
            
Next i
  
End Sub

Sub moveFileToFolder()
Dim sDaIn As String
Dim FSO As Object
Dim lro As Long
Dim sourceFileName As String
Dim destFileName As String
Dim i As Integer

sDaIn = "DataInf"
Sheets(sDaIn).Visible = True
Sheets(sDaIn).Select

lro = Cells(Rows.Count, "A").End(xlUp).Row
  
For i = 3 To lro

    Set FSO = CreateObject("scripting.filesystemobject")
    sourceFileName = Sheets(sDaIn).Cells(i, 2).Value
    destFileName = Sheets(sDaIn).Cells(i, 7).Value
    If sourceFileName <> "" Then
        FSO.MoveFile sourceFileName, destFileName
    End If
    
Next i
Sheets(sDaIn).Visible = False

End Sub
