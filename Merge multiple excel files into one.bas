Attribute VB_Name = "Module1"

Sub CombineMultipleFiles()

Application.ScreenUpdating = False

Dim oFSO As Object
Dim oFolder As Object
Dim oFile As Object
Dim i As Integer
Dim master As String
Dim masterwb As String


Set oFSO = CreateObject("Scripting.FileSystemObject")

 ' folder containing files to merge here
 
Set oFolder = oFSO.GetFolder("C:\Users\Jnrai\Downloads\rrh forums")
master = "C:\Users\Jnrai\Downloads\rrh forums"

 ' name of workbook here
 ' name of sheet to paste to should be Sheet1
 
masterwb = "Master Merger.xlsm"

For Each oFile In oFolder.Files

' go to the bottom of master
' open targetfile

    Dim targetfile As String

    targetfile = oFile

    Workbooks.Open Filename:=targetfile
    
    Set targetwb = Workbooks.Open(targetfile)
    
    Dim stringtargetwb As String
    
    stringtargetwb = targetwb.Name
    
    Debug.Print stringtargetwb

' copy targetfile data
    If IsEmpty(Range("A1").Value) = False Then
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy

' go to the bottom of master
   ' If i = 1 Then
    
' paste file into master
        Workbooks(masterwb).Activate
        Sheets("Sheet1").Select

        Sheets("Sheet1").Range("A" & Rows.Count).End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues
        
   ' ElseIf i > 1 Then
    
       ' Workbooks(1).Activate
       ' Range("A1").Select
       ' Range(Selection, Selection.End(xlToRight)).Select
       ' Range(Selection, Selection.End(xlDown)).Select
        
       ' ActiveCell.Offset(1).Select
       ' ActiveSheet.Paste
    



    
    Else
    End If
    Application.CutCopyMode = False
    Debug.Print stringtargetwb
    targetwb.Close SaveChanges:=False
    
    
' Workbooks.Close Filename:=targetfile


Next oFile

Application.ScreenUpdating = True
End Sub



