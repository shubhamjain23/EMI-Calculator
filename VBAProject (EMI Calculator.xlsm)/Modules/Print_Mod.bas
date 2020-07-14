Attribute VB_Name = "Print_Mod"
'This means all variables, object or anything must be declared explicitly
Option Explicit

Sub PrintReport()

'    Methods to print by sheet number and sheet name

'    Dim num As Integer
'    num = Application.InputBox("Enter sheet number")
'    Worksheets(num).PrintOut

'    Dim name As Integer
'    name = Application.InputBox("Enter sheet name")
'    Worksheets(name).PrintOut


    Call CreateGraph_Mod.EMIGraphs
    
'   Hiding form to show print preview
    calc_ufm.Hide
    
    Dim worksheet1 As Worksheet
    Set worksheet1 = ThisWorkbook.Sheets("EMI Graphs")
    
    ActiveSheet.Range("A1").Select
    
    Dim lastRow As Long
    
'    Select all, Move up from the last row to find the last row where data is present
    lastRow = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row
    
    Dim printObj As Object
    
'    Assigning print range to an object
'    Suppose lastRow = 10, then range = A1:H30
'    Used H row as it is the width limit of one A4 sheet for printing
    Set printObj = worksheet1.Range("A1" & ":H" & lastRow + 20)
    
'    Setting print area as the range assigned above
    worksheet1.PageSetup.PrintArea = printObj.Address
    
'    command for only print preview
    worksheet1.PrintPreview
    
    calc_ufm.Show
    
'    command for direct printing
'    worksheet1.PrintOut
    
End Sub
