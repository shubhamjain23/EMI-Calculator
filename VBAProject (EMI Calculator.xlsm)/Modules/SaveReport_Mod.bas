Attribute VB_Name = "SaveReport_Mod"
'This means all variables, object or anything must be declared explicitly
Option Explicit

Sub SaveReport()

    Call CreateGraph_Mod.EMIGraphs
    
    Sheets("EMI Graphs").Select
    
    Dim lastRow As Long
'    Select all, Move up from the last row to find the last row where data is present
    lastRow = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row
    
'    Export to pdf format and saving in the current folder
'    Suppose lastRow = 10, then range = A1:H30
'    Used H row as it is the width limit of one A4 sheet for printing
    ActiveSheet.Range("A1" & ":H" & lastRow + 20).ExportAsFixedFormat xlTypePDF, ThisWorkbook.Path & "\Report_" & VBA.Format(Now, "ddmmyyyy-hhmm") & ".pdf"

End Sub
