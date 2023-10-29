Attribute VB_Name = "Module1"
Sub Masterfile()
Attribute Masterfile.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Masterfile Macro
'
'Delete data in current import sheet
    Application.ScreenUpdating = False

    Worksheets("Reg export Sub").Activate
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.CurrentRegion.Select
    Selection.Delete
'Open a workbook

'Open method requires full file path to be referenced.
    Workbooks.Open "G:\Shared drives\Scottish Power - PD\Community management\Panel Management\MASTER Panel List\Xlookup Macro\Registration files\reg export (subscribes).xlsx"
  
'Open method has additional parameters
'Workbooks.Open(FileName, UpdateLinks, ReadOnly, Format, Password, WriteResPassword, IgnoreReadOnlyRecommended, Origin, Delimiter, Editable, Notify, Converter, AddToMru, Local, CorruptLoad)
'Help page: https://docs.microsoft.com/en-us/office/vba/api/excel.workbooks.open

'Copy range to clipboard
    Workbooks("reg export (subscribes).xlsx").Worksheets("export").Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.CurrentRegion.Select
    Selection.Copy
'paste
    Workbooks("SP Masterlist template (formula autofill) 2.xlsm").Worksheets("Reg export Sub").Range("A1").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
'close extra workbook
    Workbooks("reg export (subscribes).xlsx").Close SaveChanges:=True
'



'unsubscribes


    Worksheets("Reg export Unsub").Activate
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.CurrentRegion.Select
    Selection.Delete
'Open a workbook

'Open method requires full file path to be referenced.
    Workbooks.Open "G:\Shared drives\Scottish Power - PD\Community management\Panel Management\MASTER Panel List\Xlookup Macro\Registration files\reg export (unsubscribes).xlsx"
  
'Open method has additional parameters
'Workbooks.Open(FileName, UpdateLinks, ReadOnly, Format, Password, WriteResPassword, IgnoreReadOnlyRecommended, Origin, Delimiter, Editable, Notify, Converter, AddToMru, Local, CorruptLoad)
'Help page: https://docs.microsoft.com/en-us/office/vba/api/excel.workbooks.open

'Copy range to clipboard
    Workbooks("reg export (unsubscribes).xlsx").Worksheets("export").Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.CurrentRegion.Select
    Selection.Copy
'paste
    Workbooks("SP Masterlist template (formula autofill) 2.xlsm").Worksheets("Reg export Unsub").Range("A1").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
'close extra workbook
    Workbooks("reg export (unsubscribes).xlsx").Close SaveChanges:=True



'copy in IDs

    Sheets("Reg export Sub").Select
    Columns("A:A").Select
    Selection.Copy
    Sheets("Main Panel Live").Select
    Columns("B:B").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "ID"
    Sheets("Reg export Unsub").Select
    Columns("A:A").Select
    Selection.Copy
    Sheets("Unsubscribed").Select
    Columns("B:B").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "ID"

'formulas

    Sheets("Main Panel Live").Select
    Range("C2").Select
    ActiveCell.FormulaR1C1 = "="""""
    Range("D2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(R1C=R1C17,R1C="""",RC2=""""),"""",IFERROR(IF(VLOOKUP(RC2,'Reg export Sub'!R1:R1048576," & Chr(10) & "MATCH(" & Chr(10) & "XLOOKUP(" & Chr(10) & "R1C," & Chr(10) & "Datamap!C1," & Chr(10) & "Datamap!C2," & Chr(10) & """error"",0),'Reg export Sub'!R1,0),FALSE)<>0," & Chr(10) & "" & Chr(10) & "VLOOKUP(RC2,'Reg export Sub'!R1:R1048576," & Chr(10) & "MATCH(" & Chr(10) & "XLOOKUP(" & Chr(10) & "R1C," & Chr(10) & "Datamap!C1," & Chr(10) & "Datamap!C2," & Chr(10) & """error"",0),'Reg export Sub'!R1,0),0)," & Chr(10) & "" & Chr(10) & "IF(" & Chr(10) & "" & Chr(10) & "VLOOKUP(RC2,'Reg export Sub'!R1:R1048576," & Chr(10) & "MATCH(" & Chr(10) & "XLOOKUP(" & _
        "" & Chr(10) & "R1C," & Chr(10) & "Datamap!C1," & Chr(10) & "Datamap!C3," & Chr(10) & """error"",0),'Reg export Sub'!R1,0),0)<>0," & Chr(10) & "" & Chr(10) & "VLOOKUP(RC2,'Reg export Sub'!R1:R1048576," & Chr(10) & "MATCH(" & Chr(10) & "XLOOKUP(" & Chr(10) & "R1C," & Chr(10) & "Datamap!C1," & Chr(10) & "Datamap!C3," & Chr(10) & """error"",0),'Reg export Sub'!R1,0),0)," & Chr(10) & "" & Chr(10) & "IF(" & Chr(10) & "" & Chr(10) & "VLOOKUP(RC2,'Reg export Sub'!R1:R1048576," & Chr(10) & "MATCH(" & Chr(10) & "XLOOKUP(" & Chr(10) & "R1C," & Chr(10) & "Datamap!C1," & Chr(10) & "Datamap!C4," & Chr(10) & """error"",0),'Reg export Sub'!R1,0), 0)<>0," & Chr(10) & "" & Chr(10) & "VLOOKUP(RC2,'Reg export Sub'!" & _
        "R1:R1048576," & Chr(10) & "MATCH(" & Chr(10) & "XLOOKUP(" & Chr(10) & "R1C," & Chr(10) & "Datamap!C1," & Chr(10) & "Datamap!C4," & Chr(10) & """error"",0),'Reg export Sub'!R1,0),0)," & Chr(10) & "" & Chr(10) & "VLOOKUP(RC2,'Reg export Sub'!R1:R1048576," & Chr(10) & "MATCH(" & Chr(10) & "XLOOKUP(" & Chr(10) & "R1C," & Chr(10) & "Datamap!C1," & Chr(10) & "Datamap!C5," & Chr(10) & """error"",0),'Reg export Sub'!R1,0),0))" & Chr(10) & "" & Chr(10) & ")),0))" & _
        ""
    Range("D3").Select

    Sheets("Unsubscribed").Select
    Range("C2").Select
    ActiveCell.FormulaR1C1 = "="""""
    Range("D2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(AND(R1C=""Unsubscribed"",RC2<>0),""Yes"",IF(OR(R1C=R1C17,R1C="""",RC2=""""),"""",IFERROR(IF(VLOOKUP(RC2,'Reg export Unsub'!R1:R1048576," & Chr(10) & "MATCH(" & Chr(10) & "XLOOKUP(" & Chr(10) & "R1C," & Chr(10) & "Datamap!C1," & Chr(10) & "Datamap!C2," & Chr(10) & """error"",0),'Reg export Unsub'!R1,0),0)<>0," & Chr(10) & "" & Chr(10) & "VLOOKUP(RC2,'Reg export Unsub'!R1:R1048576," & Chr(10) & "MATCH(" & Chr(10) & "XLOOKUP(" & Chr(10) & "R1C," & Chr(10) & "Datamap!C1," & Chr(10) & "Datamap!C2," & Chr(10) & """error"",0),'Reg export Unsub'!R1,0),0)," & Chr(10) & "" & Chr(10) & "IF(" & Chr(10) & "" & Chr(10) & "" & _
        "VLOOKUP(RC2,'Reg export Unsub'!R1:R1048576," & Chr(10) & "MATCH(" & Chr(10) & "XLOOKUP(" & Chr(10) & "R1C," & Chr(10) & "Datamap!C1," & Chr(10) & "Datamap!C3," & Chr(10) & """error"",0),'Reg export Unsub'!R1,0),0)<>0," & Chr(10) & "" & Chr(10) & "VLOOKUP(RC2,'Reg export Unsub'!R1:R1048576," & Chr(10) & "MATCH(" & Chr(10) & "XLOOKUP(" & Chr(10) & "R1C," & Chr(10) & "Datamap!C1," & Chr(10) & "Datamap!C3," & Chr(10) & """error"",0),'Reg export Unsub'!R1,0),0)," & Chr(10) & "" & Chr(10) & "IF(" & Chr(10) & "" & Chr(10) & "VLOOKUP(RC2,'Reg export Unsub'!R1:R1048576," & Chr(10) & "MATCH(" & Chr(10) & "XLOOKUP(" & Chr(10) & "R1C," & Chr(10) & "Datamap!C1," & Chr(10) & "Datamap!C4," & Chr(10) & """error" & _
        """,0),'Reg export Unsub'!R1,0),0)<>0," & Chr(10) & "" & Chr(10) & "VLOOKUP(RC2,'Reg export Unsub'!R1:R1048576," & Chr(10) & "MATCH(" & Chr(10) & "XLOOKUP(" & Chr(10) & "R1C," & Chr(10) & "Datamap!C1," & Chr(10) & "Datamap!C4," & Chr(10) & """error"",0),'Reg export Unsub'!R1,0),0)," & Chr(10) & "" & Chr(10) & "VLOOKUP(RC2,'Reg export Unsub'!R1:R1048576," & Chr(10) & "MATCH(" & Chr(10) & "XLOOKUP(" & Chr(10) & "R1C," & Chr(10) & "Datamap!C1," & Chr(10) & "Datamap!C5," & Chr(10) & """error"",0),'Reg export Unsub'!R1,0),0))" & Chr(10) & "" & Chr(10) & ")),0)))" & _
        ""
    Range("D3").Select

'expand formula across

    Set SourceRange = Worksheets("Main panel Live").Range("d2")
    Set fillRange = Worksheets("Main panel Live").Range("d2:Ai2")
    SourceRange.AutoFill Destination:=fillRange
    Set SourceRange = Worksheets("Unsubscribed").Range("d2")
    Set fillRange = Worksheets("Unsubscribed").Range("d2:s2")
    SourceRange.AutoFill Destination:=fillRange

'format time

    Sheets("Unsubscribed").Select
    Columns("F:F").Select
    Selection.NumberFormat = "m/d/yyyy"
    
'Expand formula down

    Set SourceRange = Worksheets("Main panel Live").Range("d2:ai2")
    Set fillRange = Worksheets("Main panel Live").Range("d2:Ai40000")
    SourceRange.AutoFill Destination:=fillRange
    Set SourceRange = Worksheets("Unsubscribed").Range("d2:s2")
    Set fillRange = Worksheets("Unsubscribed").Range("d2:s40000")
    SourceRange.AutoFill Destination:=fillRange
    
'Missing BPID
    Sheets("Missing BPID").Select
    Cells.Select
    Selection.ClearContents
    Sheets("Main Panel Live").Select
    Columns("A:P").Select
    Selection.Copy
    Sheets("Missing BPID").Select
    Columns("A:A").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("I:J").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
    Range("H2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-1]=0,TRUE,FALSE)"
    Range("H2").Select
    Selection.AutoFill Destination:=Range("H2:H99999")
    Range("H2:H23748").Select
    Range("H2").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$N$99999").AutoFilter Field:=8, Criteria1:="FALSE"
    Rows("2:2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    ActiveWindow.SmallScroll Down:=-6
    Selection.Delete Shift:=xlUp
    Range("H7199").Select
    ActiveSheet.Range("$A$1:$N$7197").AutoFilter Field:=8
    Columns("H:H").Delete
    
'copy over to sample review sheet

    Workbooks.Open "G:\Shared drives\Scottish Power - PD\Community management\Panel Management\MASTER Panel List\Xlookup Macro\To Review\Masterfile to send to client.xlsx"

    Windows("SP Masterlist template (formula autofill) 2.xlsm").Activate
    Sheets("Main Panel Live").Select
    Cells.Select
    Selection.Copy
    Windows("Masterfile to send to client.xlsx").Activate
    Sheets("Main Panel Live").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Windows("SP Masterlist template (formula autofill) 2.xlsm").Activate
    Sheets("Unsubscribed").Select
    Cells.Select
    Application.CutCopyMode = False
    Selection.Copy
    
    Windows("Masterfile to send to client.xlsx").Activate
    Sheets("Unsubscribed").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Windows("SP Masterlist template (formula autofill) 2.xlsm").Activate
    Sheets("Missing BPID").Select
    Cells.Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("Masterfile to send to client.xlsx").Activate
    Sheets("Missing BPID").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Application.CutCopyMode = False
    Sheets("Main Panel Live").Select
    Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Unsubscribed").Select
    Cells.Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    ActiveWorkbook.Save
    ActiveWindow.Close
    Application.ScreenUpdating = True

'paste to final sheet

    
End Sub
Sub formula_paste()
Attribute formula_paste.VB_ProcData.VB_Invoke_Func = " \n14"
'
' formula_paste Macro
'

'
    Sheets("Main Panel Live").Select
    Range("C2").Select
    ActiveCell.FormulaR1C1 = "="""""
    Range("D2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(R1C=R1C17,R1C=""""),"""",IFERROR(IF(VLOOKUP(RC2,'Reg export Sub'!R1:R1048576," & Chr(10) & "MATCH(" & Chr(10) & "XLOOKUP(" & Chr(10) & "R1C," & Chr(10) & "Datamap!C1," & Chr(10) & "Datamap!C2," & Chr(10) & """error"",0),'Reg export Sub'!R1,0),FALSE)<>0," & Chr(10) & "" & Chr(10) & "VLOOKUP(RC2,'Reg export Sub'!R1:R1048576," & Chr(10) & "MATCH(" & Chr(10) & "XLOOKUP(" & Chr(10) & "R1C," & Chr(10) & "Datamap!C1," & Chr(10) & "Datamap!C2," & Chr(10) & """error"",0),'Reg export Sub'!R1,0),0)," & Chr(10) & "" & Chr(10) & "IF(" & Chr(10) & "" & Chr(10) & "VLOOKUP(RC2,'Reg export Sub'!R1:R1048576," & Chr(10) & "MATCH(" & Chr(10) & "XLOOKUP(" & _
        "" & Chr(10) & "R1C," & Chr(10) & "Datamap!C1," & Chr(10) & "Datamap!C3," & Chr(10) & """error"",0),'Reg export Sub'!R1,0),0)<>0," & Chr(10) & "" & Chr(10) & "VLOOKUP(RC2,'Reg export Sub'!R1:R1048576," & Chr(10) & "MATCH(" & Chr(10) & "XLOOKUP(" & Chr(10) & "R1C," & Chr(10) & "Datamap!C1," & Chr(10) & "Datamap!C3," & Chr(10) & """error"",0),'Reg export Sub'!R1,0),0)," & Chr(10) & "" & Chr(10) & "IF(" & Chr(10) & "" & Chr(10) & "VLOOKUP(RC2,'Reg export Sub'!R1:R1048576," & Chr(10) & "MATCH(" & Chr(10) & "XLOOKUP(" & Chr(10) & "R1C," & Chr(10) & "Datamap!C1," & Chr(10) & "Datamap!C4," & Chr(10) & """error"",0),'Reg export Sub'!R1,0), 0)<>0," & Chr(10) & "" & Chr(10) & "VLOOKUP(RC2,'Reg export Sub'!" & _
        "R1:R1048576," & Chr(10) & "MATCH(" & Chr(10) & "XLOOKUP(" & Chr(10) & "R1C," & Chr(10) & "Datamap!C1," & Chr(10) & "Datamap!C4," & Chr(10) & """error"",0),'Reg export Sub'!R1,0),0)," & Chr(10) & "" & Chr(10) & "VLOOKUP(RC2,'Reg export Sub'!R1:R1048576," & Chr(10) & "MATCH(" & Chr(10) & "XLOOKUP(" & Chr(10) & "R1C," & Chr(10) & "Datamap!C1," & Chr(10) & "Datamap!C5," & Chr(10) & """error"",0),'Reg export Sub'!R1,0),0))" & Chr(10) & "" & Chr(10) & ")),0))" & _
        ""
    Range("D3").Select
End Sub
Sub unsubformula()
Attribute unsubformula.VB_ProcData.VB_Invoke_Func = " \n14"
'
' unsubformula Macro
'

'
    Sheets("Unsubscribed").Select
    Range("C2").Select
    ActiveCell.FormulaR1C1 = "="""""
    Range("D2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(R1C=""Unsubscribed"",""Yes"",IF(OR(R1C=R1C17,R1C=""""),"""",IFERROR(IF(VLOOKUP(RC2,'Reg export Unsub'!R1:R1048576," & Chr(10) & "MATCH(" & Chr(10) & "XLOOKUP(" & Chr(10) & "R1C," & Chr(10) & "Datamap!C1," & Chr(10) & "Datamap!C2," & Chr(10) & """error"",0),'Reg export Unsub'!R1,0),0)<>0," & Chr(10) & "" & Chr(10) & "VLOOKUP(RC2,'Reg export Unsub'!R1:R1048576," & Chr(10) & "MATCH(" & Chr(10) & "XLOOKUP(" & Chr(10) & "R1C," & Chr(10) & "Datamap!C1," & Chr(10) & "Datamap!C2," & Chr(10) & """error"",0),'Reg export Unsub'!R1,0),0)," & Chr(10) & "" & Chr(10) & "IF(" & Chr(10) & "" & Chr(10) & "VLOOKUP(RC2,'Reg expo" & _
        "rt Unsub'!R1:R1048576," & Chr(10) & "MATCH(" & Chr(10) & "XLOOKUP(" & Chr(10) & "R1C," & Chr(10) & "Datamap!C1," & Chr(10) & "Datamap!C3," & Chr(10) & """error"",0),'Reg export Unsub'!R1,0),0)<>0," & Chr(10) & "" & Chr(10) & "VLOOKUP(RC2,'Reg export Unsub'!R1:R1048576," & Chr(10) & "MATCH(" & Chr(10) & "XLOOKUP(" & Chr(10) & "R1C," & Chr(10) & "Datamap!C1," & Chr(10) & "Datamap!C3," & Chr(10) & """error"",0),'Reg export Unsub'!R1,0),0)," & Chr(10) & "" & Chr(10) & "IF(" & Chr(10) & "" & Chr(10) & "VLOOKUP(RC2,'Reg export Unsub'!R1:R1048576," & Chr(10) & "MATCH(" & Chr(10) & "XLOOKUP(" & Chr(10) & "R1C," & Chr(10) & "Datamap!C1," & Chr(10) & "Datamap!C4," & Chr(10) & """error"",0),'Reg export Uns" & _
        "ub'!R1,0),0)<>0," & Chr(10) & "" & Chr(10) & "VLOOKUP(RC2,'Reg export Unsub'!R1:R1048576," & Chr(10) & "MATCH(" & Chr(10) & "XLOOKUP(" & Chr(10) & "R1C," & Chr(10) & "Datamap!C1," & Chr(10) & "Datamap!C4," & Chr(10) & """error"",0),'Reg export Unsub'!R1,0),0)," & Chr(10) & "" & Chr(10) & "VLOOKUP(RC2,'Reg export Unsub'!R1:R1048576," & Chr(10) & "MATCH(" & Chr(10) & "XLOOKUP(" & Chr(10) & "R1C," & Chr(10) & "Datamap!C1," & Chr(10) & "Datamap!C5," & Chr(10) & """error"",0),'Reg export Unsub'!R1,0),0))" & Chr(10) & "" & Chr(10) & ")),0)))" & _
        ""
    Range("D3").Select
End Sub
Sub formulaexpand()
Attribute formulaexpand.VB_ProcData.VB_Invoke_Func = " \n14"
'
' formulaexpand Macro
'

'
    Set SourceRange = Worksheets("Main panel Live").Range("d2")
    Set fillRange = Worksheets("Main panel Live").Range("d2:Ai2")
    SourceRange.AutoFill Destination:=fillRange
    Sheets("Unsubscribed").Select
    Selection.AutoFill Destination:=Range("D2:S2"), Type:=xlFillDefault
    Range("D2:S2").Select
End Sub
Sub unsubdate()
Attribute unsubdate.VB_ProcData.VB_Invoke_Func = " \n14"
'
' unsubdate Macro
'

'
    Sheets("Unsubscribed").Select
    Columns("F:F").Select
    Selection.NumberFormat = "m/d/yyyy"
End Sub


Sub expandforumlaDown()
Attribute expandforumlaDown.VB_ProcData.VB_Invoke_Func = " \n14"
'
' expandforumlaDown Macro
'

'
    Range("b2").Select
    Range(Selection, Selection.End(xlDown)).Select
End Sub
