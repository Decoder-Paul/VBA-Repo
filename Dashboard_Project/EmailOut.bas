Attribute VB_Name = "EmailOut"
Sub MailToCloud()

'========================================================================================================
' AutoEmail
' -------------------------------------------------------------------------------------------------------
' Purpose   :   To email the template to ' @trianz.com'
'
' Author    :   Subhankar Paul, 21st February, 2017
' Notes     :   N/A
'
' Parameter :   N/A
' Returns   :   N/A
' ---------------------------------------------------------------
' Revision History
'========================================================================================================

    Dim submsg As String
    submsg = "Main DATA :: As On : " & Format(Date - 1, "mmmm dd, yyyy")
    Dim OutApp As Object
    Dim OutMail As Object
    Dim strbody As String
    Dim TempFilePath As String
    Dim sNam As String
    
    'Create a new Microsoft Outlook session
    Set OutApp = CreateObject("outlook.application")
    'create a new message
    Set OutMail = OutApp.CreateItemFromTemplate(ThisWorkbook.Path & "\DashboardTemplate.oft")

    On Error Resume Next
    With OutMail
        .To = "Sangeetha.Anand@trianz.com;Badhrinath.S@trianz.com;RajKiran.Akkera@trianz.com"
        .CC = "Rakesh.Vijendra@trianz.com; mathews.jacob@trianz.com"
        .BCC = "Subhankar.Paul@trianz.com;Shambhavi.BM@trianz.com"
        .Subject = submsg
        '.Body = strbody
        Call pCreateJpg("REP", "B1:F17", "DashboardFile.jpg")
        'add the image in hidden manner, position at 0 will make it hidden
        TempFilePath = ThisWorkbook.Path & "\"
        .Attachments.Add TempFilePath & "DashboardFile.jpg", olByValue, 0

        .HTMLBody = "<br>" & strbody & "<br><br>" _
            & "<div style='position:absolute; vertical-align:middle; text-align:center; width:100%; height:100%'>" _
            & "<img src='cid:DashboardFile.jpg'" & "align=center width=width height=heigth><br><br>" _
            & "<br><br></font></span>" & .HTMLBody
        'Attachment of COPS Dashboard Excel File
        sNam = "\" & "COPS Dashboard " & dateOfAnalysis & ".xlsx"
        .Attachments.Add TempFilePath & sNam
        TempFilePath = "D:\Project COPS DashBoard\DashBoardBackup\MainData backup\MainData " & dateOfAnalysis & ".xlsx"
        .Attachments.Add TempFilePath
        .Send
    End With
    On Error GoTo 0
    
    Set OutMail = Nothing
    Set OutApp = Nothing

End Sub

Sub pCopyChart(source As String, target As String)
'========================================================================================================
' DeleteEmbeddedCharts
' -------------------------------------------------------------------------------------------------------
' Purpose   :   To copy selected charts to the sheet 'REP' in specific location
'
' Author    :   Subhankar Paul, 21st February, 2017
' Notes     :   Copying charts as chart object not as picture, to make the delete
'               procedure work correctly.
'
' Parameter :   N/A
' Returns   :   N/A
' ---------------------------------------------------------------
' Revision History
'========================================================================================================
    Call pSizeChart
    Dim ws As Worksheet
    Dim cht As Chart
    
'   Deleting old charts
    Call pDelCharts(target)
    
'   Copying New Charts
    Worksheets(source).Activate
    ThisWorkbook.Worksheets(source).ChartObjects(1).Select
    ThisWorkbook.Worksheets(source).ChartObjects(1).Copy
    Worksheets(target).Activate
    ThisWorkbook.Worksheets(target).Range("C5").Select
    ActiveSheet.Paste
    
    Worksheets(source).Activate
    ThisWorkbook.Worksheets(source).ChartObjects(6).Select
    ThisWorkbook.Worksheets(source).ChartObjects(6).Copy
    Worksheets(target).Activate
    ThisWorkbook.Worksheets(target).Range("D5").Select
    ActiveSheet.Paste
    
    Worksheets(source).Activate
    ThisWorkbook.Worksheets(source).ChartObjects(2).Select
    ThisWorkbook.Worksheets(source).ChartObjects(2).Copy
    Worksheets(target).Activate
    ThisWorkbook.Worksheets(target).Range("C9").Select
    ActiveSheet.Paste
    
    Worksheets(source).Activate
    ThisWorkbook.Worksheets(source).ChartObjects(7).Select
    ThisWorkbook.Worksheets(source).ChartObjects(7).Copy
    Worksheets(target).Activate
    ThisWorkbook.Worksheets(target).Range("D9").Select
    ActiveSheet.Paste
    
    Worksheets(source).Activate
    ThisWorkbook.Worksheets(source).ChartObjects(3).Select
    ThisWorkbook.Worksheets(source).ChartObjects(3).Copy
    Worksheets(target).Activate
    ThisWorkbook.Worksheets(target).Range("C12").Select
    ActiveSheet.Paste
    
    Worksheets(source).Activate
    ThisWorkbook.Worksheets(source).ChartObjects(9).Select
    ThisWorkbook.Worksheets(source).ChartObjects(9).Copy
    Worksheets(target).Activate
    ThisWorkbook.Worksheets(target).Range("D12").Select
    ActiveSheet.Paste
    
    Worksheets(source).Activate
    ThisWorkbook.Worksheets(source).ChartObjects(4).Select
    ThisWorkbook.Worksheets(source).ChartObjects(4).Copy
    Worksheets(target).Activate
    ThisWorkbook.Worksheets(target).Range("C14").Select
    ActiveSheet.Paste
    
    Worksheets(source).Activate
    ThisWorkbook.Worksheets(source).ChartObjects(8).Select
    ThisWorkbook.Worksheets(source).ChartObjects(8).Copy
    Worksheets(target).Activate
    ThisWorkbook.Worksheets(target).Range("D14").Select
    ActiveSheet.Paste
    
    Worksheets(source).Activate
    ThisWorkbook.Worksheets(source).ChartObjects(10).Select
    ThisWorkbook.Worksheets(source).ChartObjects(10).Copy
    Worksheets(target).Activate
    ThisWorkbook.Worksheets(target).Range("E5").Select
    ActiveSheet.Paste
    
    Worksheets(source).Activate
    ThisWorkbook.Worksheets(source).ChartObjects(11).Select
    ThisWorkbook.Worksheets(source).ChartObjects(11).Copy
    Worksheets(target).Activate
    ThisWorkbook.Worksheets(target).Range("E9").Select
    ActiveSheet.Paste
    
    Worksheets(source).Activate
    ThisWorkbook.Worksheets(source).ChartObjects(12).Select
    ThisWorkbook.Worksheets(source).ChartObjects(12).Copy
    Worksheets(target).Activate
    ThisWorkbook.Worksheets(target).Range("E12").Select
    ActiveSheet.Paste

    Worksheets(source).Activate
    ThisWorkbook.Worksheets(source).ChartObjects(13).Select
    ThisWorkbook.Worksheets(source).ChartObjects(13).Copy
    Worksheets(target).Activate
    ThisWorkbook.Worksheets(target).Range("E14").Select
    ActiveSheet.Paste
    
End Sub
Sub pAutoEmail()
'========================================================================================================
' AutoEmail
' -------------------------------------------------------------------------------------------------------
' Purpose   :   To email the template to ' @trianz.com'
'
' Author    :   Subhankar Paul, 21st February, 2017
' Notes     :   N/A
'
' Parameter :   N/A
' Returns   :   N/A
' ---------------------------------------------------------------
' Revision History
'========================================================================================================

    Dim submsg As String
    submsg = "Client Operations Dashboard (COPS) :: As On : " & Format(dateOfAnalysis, "mmmm dd, yyyy")
    Dim OutApp As Object
    Dim OutMail As Object
    Dim strbody As String
    Dim TempFilePath As String
    Dim sNam As String
    
    'Create a new Microsoft Outlook session
    Set OutApp = CreateObject("outlook.application")
    'create a new message
    Set OutMail = OutApp.CreateItemFromTemplate(ThisWorkbook.Path & "\DashboardTemplate.oft")

    On Error Resume Next
    With OutMail
        .To = "Cops@trianz.com"
        .CC = ""
        .BCC = ""
        .Subject = submsg
        '.Body = strbody
        Call pCreateJpg("REP", "B1:F17", "DashboardFile.jpg")
        'add the image in hidden manner, position at 0 will make it hidden
        TempFilePath = ThisWorkbook.Path & "\"
        .Attachments.Add TempFilePath & "DashboardFile.jpg", olByValue, 0

        .HTMLBody = "<br>" & strbody & "<br><br>" _
            & "<div style='position:absolute; vertical-align:middle; text-align:center; width:100%; height:100%'>" _
            & "<img src='cid:DashboardFile.jpg'" & "align=center width=width height=heigth><br><br>" _
            & "<br><br></font></span>" & .HTMLBody
        'Attachment of COPS Dashboard Excel File
        sNam = "\" & "COPS Dashboard " & dateOfAnalysis & ".xlsx"
        .Attachments.Add TempFilePath & sNam
        .Send
    End With
    On Error GoTo 0
    
    Set OutMail = Nothing
    Set OutApp = Nothing
    
End Sub
Sub pCreateJpg(Namesheet As String, nameRange As String, nameFile As String)

Dim w, h As Long
Dim sFilePath As String

ThisWorkbook.Activate
Worksheets(Namesheet).Activate
Set Plage = ThisWorkbook.Worksheets(Namesheet).Range(nameRange)
Plage.CopyPicture xlScreen, xlPicture

sFilePath = ThisWorkbook.Path & "\" & nameFile

w = Plage.Width
h = Plage.Height

With ThisWorkbook.ActiveSheet

    .Activate

    Dim chtObj As ChartObject
    Set chtObj = .ChartObjects.Add(100, 30, 400, 250)
    chtObj.Name = "TemporaryPictureChart"

    'resize obj to picture size
    chtObj.Width = w
    chtObj.Height = h

    ActiveSheet.ChartObjects("TemporaryPictureChart").Activate
    ActiveChart.Paste

    ActiveChart.Export sFilePath, FilterName:="JPG"

    chtObj.Delete

End With

Set Plage = Nothing

End Sub

Sub pSizeChart()
'========================================================================================================
' TicketCount
' -------------------------------------------------------------------------------------------------------
' Purpose   :   To Resize the charts by a fixed size and position in the project / cluster sheet
'               right after creation of it.
'
' Author    :   Subhankar Paul, 27th February, 2017
' Notes     :
' ---------------------------------------------------------------
' Revision History
'========================================================================================================

Dim chtObj As ChartObject
Dim sheetDbd As String

sheetDbd = "Project or Cluster"

Sheets(sheetDbd).Activate
Sheets(sheetDbd).Select

'Same width and Height
For Each chtObj In ActiveSheet.ChartObjects
    chtObj.Width = 291
    chtObj.Height = 185
Next

End Sub

Sub pDelCharts(target As String)
' Purpose   :   To delete all the available charts on the sheet 'REP'

    Dim wsItem As Worksheet
    Dim chtObj As ChartObject
        For Each chtObj In ThisWorkbook.Worksheets(target).ChartObjects
            chtObj.Delete
        Next
End Sub
