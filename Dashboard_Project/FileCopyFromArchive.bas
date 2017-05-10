Attribute VB_Name = "FileCopyFromArchive"
'-- Date of Analysis is basicall a global variable which hold the default value Today()-1
Public dateOfAnalysis As Date
Sub archiveFileExtraction()
    'Remember time when macro starts
    StartTime = Timer
    
    Dim WB As Workbook
    Dim WS_COPS As Worksheet
    Dim WS_DA As Worksheet
    Set WB = ActiveWorkbook
    Set WS_COPS = WB.Sheets("Cops DashBoard")
    
    Dim startDate As Date
    Dim endDate As Date
    
    WS_COPS.Activate
    WS_COPS.Select
    
    startDate = WS_COPS.Cells(14, 7).Value
    endDate = WS_COPS.Cells(14, 9).Value
    dateOfAnalysis = startDate
    
    Do While dateOfAnalysis <> endDate + 1
        Call copyFileOnDateMatch
        Call pMain
        dateOfAnalysis = dateOfAnalysis + 1
    Loop
    
    ' Determine how many seconds this code will take to run
    SecondsElapsed = Round(Timer - StartTime, 2)
    'Notify user in seconds
    MsgBox "This code ran successfully in " & SecondsElapsed & " seconds", vbInformation

End Sub
Sub copyFileOnDateMatch()
'========================================================================================================
' copyFileOnDateMatch
' -------------------------------------------------------------------------------------------------------
' Purpose of this programm : To copy raw file from archive folder of each client folder to client folder
'                            on the basis of Date given
' Author : Subhankar Paul 25th April, 2017
' Notes  : dateOfAnalysis is a global variable
' Parameters : N/A
' Returns : N/A
' ---------------------------------------------------------------
' Revision History
'
'========================================================================================================

    Dim s, file, fold, fa, archiveFile, subFold, f, fi, fObj, subArchive As Variant
    Dim filePath, deletePath, masterPath, toPath, sourceDate, sourceMonth, sourceYear As String
    Dim pos As Integer
    
    filePath = Application.ActiveWorkbook.Path
    
    
    Set fObj = CreateObject("Scripting.FileSystemObject")
    Set fold = fObj.GetFolder(filePath)
    Set subFold = fold.SubFolders 'subFold is a Collection of Client folder
    For Each f In subFold
        s = s & f.Name
        s = s & vbCrLf
'        Debug.Print f.Name
        
'------ Renaming the Archive files of Master Card like "Opening 06-04-2017"
        If f.Name = "MASTER" Then
            masterPath = filePath & "\" & f.Name & "\" & "Archive\"
            Call renameFile(masterPath)
        End If
        
'------ These below files are in the folder of client for Dashboard Analysis
        Set file = f.Files
'        For Each fi In file
'            Debug.Print Tab(5); fi.Name
'        Next
        
'------ If File Exist it will be deleted and next Date raw file will arrive from Archive
        deletePath = filePath & "\" & f.Name & "\*.xlsx*"
        If Dir(deletePath) <> "" Then
            fObj.DeleteFile deletePath, True
        End If
        deletePath = filePath & "\" & f.Name & "\*.xls*"
        If Dir(deletePath) <> "" Then
            fObj.DeleteFile deletePath, True
        End If
        
        Set subArchive = f.SubFolders 'Archive subfolder
'------ Getting into Archive Folder
        For Each fa In subArchive
            Set archiveFile = fa.Files
            For Each fi In archiveFile
 
'-------------- Source Year Extraction
                pos = InStr(fi.Name, ".")
                sourceYear = Mid(fi.Name, pos - 4, 4)
                
'-------------- Source Month Extraction
                sourceMonth = Mid(fi.Name, pos - 7, 2)
                
                If Left(sourceMonth, 1) <> "0" Or Left(sourceMonth, 1) <> "1" Then
                    sourceMonth = "0" & Right(sourceMonth, 1)
                End If
                                
'-------------- Source Date Extraction
                sourceDate = Mid(fi.Name, pos - 10, 3)
'-------------- Checking for 4 Different Cases
'--------------     01-03-2017
'--------------     01-3-2017
'--------------     1-03-2017
'--------------     1-3-2017

                If IsNumeric(Left(sourceDate, 1)) = True Then
                    sourceDate = Left(sourceDate, 2)
                ElseIf IsNumeric(Right(sourceDate, 1)) = True And IsNumeric(Mid(sourceDate, 2, 1)) = True Then
                    sourceDate = Right(sourceDate, 2)
                ElseIf IsNumeric(Mid(sourceDate, 2, 1)) = True Then
                    sourceDate = "0" & Mid(sourceDate, 2, 1)
                ElseIf IsNumeric(Right(sourceDate, 1)) = True Then
                    sourceDate = "0" & Right(sourceDate, 1)
                End If
                sourceDate = sourceDate & "-" & sourceMonth & "-" & sourceYear
                
'------ If the Date is matched then the file will be copied to Client folder
                If sourceDate = dateOfAnalysis Then
                    toPath = filePath & "\" & f.Name & "\"
                    fi.Copy toPath
                End If
            Next
        Next
    Next
    'MsgBox s
End Sub
Sub renameFile(sourcePath)
    Dim MySource, MyObj As Object
    Dim file As Variant
    Dim newName As String
    Set MyObj = CreateObject("scripting.filesystemobject")
    Set MySource = MyObj.GetFolder(sourcePath)
    For Each file In MySource.Files
        
        If Len(file.Name) = 28 Then
            newName = file.Name
            dateUsed = Mid(newName, 9, 2) & "/" & Mid(newName, 11, 3) & "/" & Mid(newName, 14, 4)
            date_format = Format(dateUsed, "DD-MM-YYYY")
            newName = "Opening " & date_format & ".xls"
            Name sourcePath & "\" & file.Name As sourcePath & "\" & newName
        End If
    Next file
End Sub

