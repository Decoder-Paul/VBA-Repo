Attribute VB_Name = "Clients"
Option Explicit

Sub pHER(icount As Long)
    
    Dim WB As Workbook
    Dim WS_DaIn As Worksheet
    Dim WS_In As Worksheet
    
    Dim sDaIn As String
    Dim sIn As String
    Dim sP1 As String
    
    Dim lroInc As Long
    Dim lco As Long
    Dim lro As Long
      
    sDaIn = "DataInf"
    sIn = "Incident"
    sP1 = "Report"
    
    Set WB = ActiveWorkbook
    Set WS_DaIn = WB.Sheets(sDaIn)
    Set WS_In = WB.Sheets(sIn)
    
    'Checking if Incident sheet is available or not
    
    Workbooks(sCFilNam).Activate
    Sheets(sIn).Activate
    lroInc = WS_In.Cells(WS_In.Rows.Count, "a").End(xlUp).Row
    
    Sheets(sDaIn).Activate
    Sheets(sDaIn).Range("a1").Select
    'InFlow
If Left(LCase(WS_DaIn.Cells(icount, 1).Value), 6) = "inflow" Or _
    Left(LCase(WS_DaIn.Cells(icount, 1).Value), 7) = "in flow" Then
            
    'Opening the Sheet
    Workbooks.Open (WS_DaIn.Cells(icount, 2).Value)
    Workbooks(WS_DaIn.Cells(icount, 1).Value).Activate
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
            lco = Cells(1, Columns.Count).End(xlToRight).Column
            'copy and paste dump only if lro is greater than or equal to 2
            If lro >= 2 Then
                Range(Cells(1, 1), Cells(lro, lco)).Copy
                Workbooks(sCFilNam).Activate
                Sheets(sIn).Activate
                Sheets(sIn).Range("a1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
                Application.CutCopyMode = False
            End If
            Workbooks(WS_DaIn.Cells(icount, 1).Value).Close
        Else
            Sheets(sP1).Activate
            Sheets(sP1).Range("a1").Select
            lro = Cells(Rows.Count, "A").End(xlUp).Row
            lco = Cells(1, Columns.Count).End(xlToRight).Column
            If lro >= 2 Then
                Range(Cells(2, 1), Cells(lro, lco)).Copy
                Workbooks(sCFilNam).Activate
                Sheets(sIn).Activate
                Sheets(sIn).Cells(lroInc + 1, 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
                Application.CutCopyMode = False
            End If
            Workbooks(WS_DaIn.Cells(icount, 1).Value).Close
       End If
    End If
                
    'Outflow
ElseIf Left(LCase(WS_DaIn.Cells(icount, 1).Value), 7) = "outflow" Or _
       Left(LCase(WS_DaIn.Cells(icount, 1).Value), 8) = "out flow" Then
                    
   'Opening the Sheet
  Workbooks.Open (WS_DaIn.Cells(icount, 2).Value)
  Workbooks(WS_DaIn.Cells(icount, 1).Value).Activate
   'if sheet page 1 is avalable or not select
  If fSheetExists(sP1) = False Then
        Workbooks(WS_DaIn.Cells(icount, 1).Value).Close
        MsgBox "Page 1 Sheet is missing in OutFlow File."
        End
  Else
       If lroInc = 1 Then
            Sheets(sP1).Activate
            Sheets(sP1).Range("a1").Select
            lro = Cells(Rows.Count, "A").End(xlUp).Row
            lco = Cells(1, Columns.Count).End(xlToRight).Column
            If lro >= 2 Then
                Range(Cells(1, 1), Cells(lro, lco)).Copy
                Workbooks(sCFilNam).Activate
                Sheets(sIn).Activate
                Sheets(sIn).Range("a1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
                Application.CutCopyMode = False
            End If
           Workbooks(WS_DaIn.Cells(icount, 1).Value).Close
        Else
            Sheets(sP1).Activate
            Sheets(sP1).Range("a1").Select
            lro = Cells(Rows.Count, "A").End(xlUp).Row
            lco = Cells(1, Columns.Count).End(xlToRight).Column
            If lro >= 2 Then
                Range(Cells(2, 1), Cells(lro, lco)).Copy
                Workbooks(sCFilNam).Activate
                Sheets(sIn).Activate
                Sheets(sIn).Cells(lroInc + 1, 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
                Application.CutCopyMode = False
            End If
            Workbooks(WS_DaIn.Cells(icount, 1).Value).Close
       End If
  End If
  
ElseIf Left(LCase(WS_DaIn.Cells(icount, 1).Value), 7) = "opening" Then
                    
 
   'Opening the Sheet
  Workbooks.Open (WS_DaIn.Cells(icount, 2).Value)
  Workbooks(WS_DaIn.Cells(icount, 1).Value).Activate
   'if sheet page 1 is avalable or not select
  If fSheetExists(sP1) = False Then
        Workbooks(Cells(icount, 1).Value).Close
        MsgBox "Page 1 Sheet is missing in OutFlow File."
        End
  Else
       If lroInc = 1 Then
            Sheets(sP1).Activate
            Sheets(sP1).Range("a1").Select
            lro = Cells(Rows.Count, "A").End(xlUp).Row
            lco = Cells(1, Columns.Count).End(xlToRight).Column
            If lro >= 2 Then
                Range(Cells(1, 1), Cells(lro, lco)).Copy
                Workbooks(sCFilNam).Activate
                Sheets(sIn).Activate
                Sheets(sIn).Range("a1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
                Application.CutCopyMode = False
            End If
            Workbooks(WS_DaIn.Cells(icount, 1).Value).Close
        Else
            Sheets(sP1).Activate
            Sheets(sP1).Range("a1").Select
            lro = Cells(Rows.Count, "A").End(xlUp).Row
            lco = Cells(1, Columns.Count).End(xlToRight).Column
            If lro >= 2 Then
                Range(Cells(2, 1), Cells(lro, lco)).Copy
                Workbooks(sCFilNam).Activate
                Sheets(sIn).Activate
                Sheets(sIn).Cells(lroInc + 1, 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
                Application.CutCopyMode = False
            End If
            Workbooks(WS_DaIn.Cells(icount, 1).Value).Close
       End If
   End If
End If
End Sub
'------------------------------------------------------------------------------
 Sub pNYL(icount As Long)
  
  Dim WB As Workbook
  Dim WS_DaIn As Worksheet
  Dim WS_In As Worksheet
  
  Dim sDaIn As String
  Dim sIn As String
  Dim sP1 As String
  
  Dim lroInc As Long
  Dim lco As Long
  Dim lro As Long
    
  sDaIn = "DataInf"
  sIn = "Incident"
  sP1 = "Page 1"
  
  Set WB = ActiveWorkbook
  Set WS_DaIn = WB.Sheets(sDaIn)
  Set WS_In = WB.Sheets(sIn)
  
    'Checking if Incident sheet is available or not

Workbooks(sCFilNam).Activate
Sheets(sIn).Activate
lroInc = WS_In.Cells(WS_In.Rows.Count, "a").End(xlUp).Row
  
Sheets(sDaIn).Activate
Sheets(sDaIn).Range("a1").Select
  
    'InFlow
    If Left(LCase(WS_DaIn.Cells(icount, 1).Value), 6) = "inflow" Or _
    Left(LCase(WS_DaIn.Cells(icount, 1).Value), 7) = "in flow" Then
            
    'Opening the Sheet
    Workbooks.Open (WS_DaIn.Cells(icount, 2).Value)
    Workbooks(WS_DaIn.Cells(icount, 1).Value).Activate
    'if sheet page 1 is avalable or not select
  If fSheetExists(sP1) = False Then
        Workbooks(WS_DaIn.Cells(icount, 1).Value).Close
        MsgBox "Page 1 Sheet is missing in InFlow File."
        End
  Else
       If lroInc = 1 Then
            Sheets("Page 1").Activate
            Sheets("Page 1").Range("a1").Select
            lro = Cells(Rows.Count, "A").End(xlUp).Row
            lco = Cells(1, Columns.Count).End(xlToRight).Column
            If lro >= 2 Then
                Range(Cells(1, 1), Cells(lro, lco)).Copy
                Workbooks(sCFilNam).Activate
                Sheets(sIn).Activate
                Sheets(sIn).Range("a1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
                Application.CutCopyMode = False
            End If
            Workbooks(WS_DaIn.Cells(icount, 1).Value).Close
        Else
            Sheets("Page 1").Activate
            Sheets("Page 1").Range("a1").Select
            lro = Cells(Rows.Count, "A").End(xlUp).Row
            lco = Cells(1, Columns.Count).End(xlToRight).Column
            If lro >= 2 Then
                Range(Cells(2, 1), Cells(lro, lco)).Copy
                Workbooks(sCFilNam).Activate
                Sheets(sIn).Activate
                Sheets(sIn).Cells(lroInc + 1, 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
                Application.CutCopyMode = False
            End If
            Workbooks(WS_DaIn.Cells(icount, 1).Value).Close
       End If
  End If
                
    'Outflow
ElseIf Left(LCase(WS_DaIn.Cells(icount, 1).Value), 7) = "outflow" Or _
       Left(LCase(WS_DaIn.Cells(icount, 1).Value), 8) = "out flow" Then
                    
   'Opening the Sheet
  Workbooks.Open (WS_DaIn.Cells(icount, 2).Value)
  Workbooks(WS_DaIn.Cells(icount, 1).Value).Activate
   'if sheet page 1 is avalable or not select
  If fSheetExists(sP1) = False Then
        Workbooks(WS_DaIn.Cells(icount, 1).Value).Close
        MsgBox "Page 1 Sheet is missing in OutFlow File."
        End
  Else
       If lroInc = 1 Then
            Sheets("Page 1").Activate
            Sheets("Page 1").Range("a1").Select
            lro = Cells(Rows.Count, "A").End(xlUp).Row
            lco = Cells(1, Columns.Count).End(xlToRight).Column
            If lro >= 2 Then
                Range(Cells(1, 1), Cells(lro, lco)).Copy
                Workbooks(sCFilNam).Activate
                Sheets(sIn).Activate
                Sheets(sIn).Range("a1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
                Application.CutCopyMode = False
            End If
            Workbooks(WS_DaIn.Cells(icount, 1).Value).Close
        Else
            Sheets("Page 1").Activate
            Sheets("Page 1").Range("a1").Select
            lro = Cells(Rows.Count, "A").End(xlUp).Row
            lco = Cells(1, Columns.Count).End(xlToRight).Column
            If lro >= 2 Then
                Range(Cells(2, 1), Cells(lro, lco)).Copy
                Workbooks(sCFilNam).Activate
                Sheets(sIn).Activate
                Sheets(sIn).Cells(lroInc + 1, 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
                Application.CutCopyMode = False
            End If
            Workbooks(WS_DaIn.Cells(icount, 1).Value).Close
       End If
  End If
  
ElseIf Left(LCase(WS_DaIn.Cells(icount, 1).Value), 7) = "opening" Then
                    
 
   'Opening the Sheet
  Workbooks.Open (WS_DaIn.Cells(icount, 2).Value)
  Workbooks(WS_DaIn.Cells(icount, 1).Value).Activate
   'if sheet page 1 is avalable or not select
  If fSheetExists(sP1) = False Then
        Workbooks(Cells(icount, 1).Value).Close
        MsgBox "Page 1 Sheet is missing in OutFlow File."
        End
  Else
       If lroInc = 1 Then
            Sheets("Page 1").Activate
            Sheets("Page 1").Range("a1").Select
            lro = Cells(Rows.Count, "A").End(xlUp).Row
            lco = Cells(1, Columns.Count).End(xlToRight).Column
            If lro >= 2 Then
                Range(Cells(1, 1), Cells(lro, lco)).Copy
                Workbooks(sCFilNam).Activate
                Sheets(sIn).Activate
                Sheets(sIn).Range("a1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
                Application.CutCopyMode = False
            End If
            Workbooks(WS_DaIn.Cells(icount, 1).Value).Close
        Else
            Sheets("Page 1").Activate
            Sheets("Page 1").Range("a1").Select
            lro = Cells(Rows.Count, "A").End(xlUp).Row
            lco = Cells(1, Columns.Count).End(xlToRight).Column
            If lro >= 2 Then
                Range(Cells(2, 1), Cells(lro, lco)).Copy
                Workbooks(sCFilNam).Activate
                Sheets(sIn).Activate
                Sheets(sIn).Cells(lroInc + 1, 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
                Application.CutCopyMode = False
            End If
            Workbooks(WS_DaIn.Cells(icount, 1).Value).Close
       End If
   End If
 End If
 
End Sub
'----------------------------------------------------------------------------------------------
Sub pMAS(icount As Long)
 Dim WB As Workbook
  Dim WS_DaIn As Worksheet
  Dim WS_In As Worksheet
  
  Dim sDaIn As String
  Dim sIn As String
  Dim sR2 As String
  
  Dim lroInc As Long
  Dim lco As Long
  Dim lro As Long
    
  sDaIn = "DataInf"
  sIn = "Incident"
  sR2 = "Report"
  
  Set WB = ActiveWorkbook
  Set WS_DaIn = WB.Sheets(sDaIn)
 

  Set WS_In = WB.Sheets(sIn)
Workbooks(sCFilNam).Activate
Sheets(sIn).Activate
lroInc = WS_In.Cells(WS_In.Rows.Count, "a").End(xlUp).Row
  
Sheets(sDaIn).Activate
Sheets(sDaIn).Range("a1").Select
  
    'Outflow
If Left(LCase(WS_DaIn.Cells(icount, 1).Value), 7) = "outflow" Or _
       Left(LCase(WS_DaIn.Cells(icount, 1).Value), 8) = "out flow" Then
                    
   'Opening the Sheet
  Workbooks.Open (Cells(icount, 2).Value)
  Workbooks(WS_DaIn.Cells(icount, 1).Value).Activate
   'if sheet page 1 is avalable or not select
  If fSheetExists(sR2) = False Then
        Workbooks(WS_DaIn.Cells(icount, 1).Value).Close
        MsgBox "Report Sheet is missing in OutFlow File."
        End
  Else
       If lroInc = 1 Then
            Sheets(sR2).Activate
            Sheets(sR2).Range("a1").Select
            lro = Cells(Rows.Count, "A").End(xlUp).Row
            lco = Cells(1, Columns.Count).End(xlToRight).Column
            If lro >= 2 Then
                Range(Cells(1, 1), Cells(lro, lco)).Copy
                Workbooks(sCFilNam).Activate
                Sheets(sIn).Activate
                Sheets(sIn).Range("a1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
                Application.CutCopyMode = False
            End If
            Workbooks(WS_DaIn.Cells(icount, 1).Value).Close
        Else
            Sheets(sR2).Activate
            Sheets(sR2).Range("a1").Select
            lro = Cells(Rows.Count, "A").End(xlUp).Row
            lco = Cells(1, Columns.Count).End(xlToRight).Column
            If lro >= 2 Then
                Range(Cells(2, 1), Cells(lro, lco)).Copy
                Workbooks(sCFilNam).Activate
                Sheets(sIn).Activate
                Sheets(sIn).Cells(lroInc + 1, 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
                Application.CutCopyMode = False
            End If
            Workbooks(WS_DaIn.Cells(icount, 1).Value).Close
       End If
  End If
  
ElseIf Left(LCase(WS_DaIn.Cells(icount, 1).Value), 7) = "opening" Then
                    
   'Opening the Sheet
  Workbooks.Open (WS_DaIn.Cells(icount, 2).Value)
  Workbooks(WS_DaIn.Cells(icount, 1).Value).Activate
            
   'if sheet page 1 is avalable or not select
  If fSheetExists(sR2) = False Then
        Workbooks(WS_DaIn.Cells(icount, 1).Value).Close
        MsgBox "Report Sheet is missing in Opening File."
        End
  Else
       If lroInc = 1 Then
            Sheets(sR2).Activate
            Sheets(sR2).Range("a1").Select
            lro = Cells(Rows.Count, "A").End(xlUp).Row
            lco = Cells(1, Columns.Count).End(xlToRight).Column
            If lro >= 2 Then
                Range(Cells(1, 1), Cells(lro, lco)).Copy
                Workbooks(sCFilNam).Activate
                Sheets(sIn).Activate
                Sheets(sIn).Range("a1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
                Application.CutCopyMode = False
            End If
            Workbooks(WS_DaIn.Cells(icount, 1).Value).Close
        Else
            Sheets(sR2).Activate
            Sheets(sR2).Range("a1").Select
            lro = Cells(Rows.Count, "A").End(xlUp).Row
            lco = Cells(1, Columns.Count).End(xlToRight).Column
            If lro >= 2 Then
                Range(Cells(2, 1), Cells(lro, lco)).Copy
                Workbooks(sCFilNam).Activate
                Sheets(sIn).Activate
                Sheets(sIn).Cells(lroInc + 1, 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
                Application.CutCopyMode = False
            End If
            Workbooks(WS_DaIn.Cells(icount, 1).Value).Close
       End If
   End If
 End If

End Sub

'------------------------------------------------------------------------------------------
Sub pEquinix(icount As Long)
 Dim WB As Workbook
  Dim WS_DaIn As Worksheet
  Dim WS_In As Worksheet
  
  Dim sDaIn As String
  Dim sIn As String
  Dim sR2 As String
  
  Dim lco As Long
  Dim lro As Long
    
  sDaIn = "DataInf"
  sIn = "Incident"
  sR2 = "Project Progress Metrics"
  
  Set WB = ActiveWorkbook
  Set WS_DaIn = WB.Sheets(sDaIn)
 

  Set WS_In = WB.Sheets(sIn)
Workbooks(sCFilNam).Activate
Sheets(sIn).Activate

Sheets(sDaIn).Activate
Sheets(sDaIn).Range("a1").Select
  
If Left(LCase(WS_DaIn.Cells(icount, 1).Value), 7) = "opening" Then
                    
   'Opening the Sheet
  Workbooks.Open (WS_DaIn.Cells(icount, 2).Value)
  Workbooks(WS_DaIn.Cells(icount, 1).Value).Activate
            
   'if sheet Project Progress Metrics is avalable or not select
  If fSheetExists(sR2) = False Then
        Workbooks(WS_DaIn.Cells(icount, 1).Value).Close
        MsgBox "Project Progress Metrics Sheet is missing in Opening File."
        End
  Else
        Sheets(sR2).Activate
        Sheets(sR2).Range("a1").Select
        lro = Cells(Rows.Count, "A").End(xlUp).Row
        lco = Cells(1, Columns.Count).End(xlToRight).Column
        If lro >= 2 Then
            Range(Cells(1, 1), Cells(lro, lco)).Copy
            Workbooks(sCFilNam).Activate
            Sheets(sIn).Activate
            Sheets(sIn).Range("a1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            Application.CutCopyMode = False
        End If
        Workbooks(WS_DaIn.Cells(icount, 1).Value).Close
   End If
End If

End Sub




