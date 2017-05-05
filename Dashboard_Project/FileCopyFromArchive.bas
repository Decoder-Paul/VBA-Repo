Attribute VB_Name = "FileCopyFromArchive"
Sub archiveFileExtraction()
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
    
End Sub
Sub copyFileOnDateMatch()
'========================================================================================================
' copyFileOnDateMatch
' -------------------------------------------------------------------------------------------------------
' Purpose of this programm : To copy raw file from archive folder of each client folder to client folder
'                            on the basis of Date given
' Author : Subhankar Paul 25th April, 2017
' Notes  : N/A
' Parameters : N/A
' Returns : N/A
' ---------------------------------------------------------------
' Revision History
'
'========================================================================================================

    Dim s, file, fold, fa, archiveFile, subFold, f, fi, fObj, subArchive As Variant
    Dim filePath, deletePath, masterPath, toPath As String
    
    filePath = Application.ActiveWorkbook.Path
    
    Debug.Print startDate
    
    Set fObj = CreateObject("Scripting.FileSystemObject")
    Set fold = fObj.GetFolder(filePath)
    Set subFold = fold.SubFolders 'subFold is a Collection of Subfolders
    For Each f In subFold
        s = s & f.Name
        s = s & vbCrLf
        Set subArchive = f.SubFolders
        Debug.Print f.Name
'------ Renaming the Archive files of Master Card like "Opening 06-04-2017"
        If f.Name = "MASTER" Then
            masterPath = filePath & "\" & f.Name & "\" & "Archive\"
            Call renameFile(masterPath)
        End If
'------ These below files are in the folder of client for Dashboard Analysis
        Set file = f.Files
        For Each fi In file
            Debug.Print Tab(5); fi.Name
        Next
'------ If File Exist it will be deleted and next Date raw file will arrive from Archive
        deletePath = filePath & "\" & f.Name & "\*.xlsx*"
        
        fObj.DeleteFile deletePath, True

'------ Getting into Archive Folder
        For Each fa In subArchive
            Debug.Print Tab(5); fa.Name
            Set archiveFile = fa.Files
            For Each fi In archiveFile
                Debug.Print Tab(10); fi.Name
'------ If the Date is matched then the file will be copied to Client folder
                If criteria Then
                    toPath = filePath & "\" & f.Name & "\"
                    fi.Copy toPath
                End If
            Next
        Next
    Next
    MsgBox s
End Sub
Sub renameFile(sourcePath)
    Dim MySource, MyObj As Object
    Dim file As Variant
    Dim newName As String
    
    Set MySource = MyObj.GetFolder(sourcePath)
    For Each file In MySource.Files
        newName = file.Name
        Debug.Print newName
        newName = Right(newName, 8)
        Name sourcePath & "\" & file.Name As sourcePath & "\" & newName
    Next file
End Sub

