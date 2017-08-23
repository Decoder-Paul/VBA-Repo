Attribute VB_Name = "BinarySearch"
Sub pSearchCustName()

Dim Op As String
Dim Searchitem As String
Dim Totalid As Long
Dim i As Long
Dim j As Long
Dim cellda As Range


Dim WB As Workbook
Dim WSOp As Worksheet
Dim WSBR As Worksheet

Op = "Sheet1"

Set WB = ActiveWorkbook
Set WSOp = WB.Sheets(Op)

'Sheet Booking
Sheets(Op).Activate
Sheets(Op).Select

lro_BR = Cells(Rows.Count, "C").End(xlUp).Row
lro = Cells(Rows.Count, "A").End(xlUp).Row

   
'-------------------Sorting Customer name Opportunity
    WSOp.Range(Cells(1, 1), Cells(lro, 1)).Sort key1:=Range("A1:A" & lro), _
   order1:=xlAscending, Header:=xlYes
   
For i = 1 To 1000
    Searchitem = Cells(i, 1).Value
    Sheets(Op).Range(Cells(1, 3), Cells(lro_BR, 3)).Select
    
  Set cellda = Selection.Find(What:=Searchitem, LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext)

    If cellda Is Nothing Then
        Cells(i, 2).Value = "N"
    Else
        Cells(i, 2).Value = "Y"
    End If
    j = i
    Do While Searchitem = Cells(i + 1, 1).Value
        i = i + 1
        Cells(i, 2).Value = "Copy"
    Loop
Next i
    'Opportunity Customer Name no. is 58 - BF
            'Booking Customer name No. is 39 - AM
            'Totalid = Application.CountIfs(WSBR.Range(WSBR.Cells(2, 39), WSBR.Cells(lro_BR, 39)), Searchitem)
        
End Sub
