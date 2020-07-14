Attribute VB_Name = "UpdateBanks_Mod"
'This means all variables, object or anything must be declared explicitly
Option Explicit

Sub UpdateBanks()

'    Opening workbook "Bank Details" which must be enclosed in the same folder
    Workbooks.Open Filename:= _
        (Application.ActiveWorkbook.Path & "\Bank Details.xlsx")
        
    Sheets("Source Data").Select
    Range("A1").Select
    
'    Selecting current region where data is present
'    Alternatively, we can select data by using last row and last column concept
    Selection.CurrentRegion.Select
    
'    Copying data and closing the workbook
    Selection.Copy
    ActiveWindow.Close

'    Activating the main file, no need to open
    Windows("EMI Calculator.xlsm").Activate
    
    Call createSheet_Mod.createSheet("Bank Details")
    
'    Clearing all contents from 2nd row
    Sheets("Bank Details").Select
    Rows("2:" & Rows.Count).ClearContents
    
'    Pasting data from cell A1 in
    With ActiveSheet
        .Range("A1").Select
        .Paste
    End With
    
    Call fitAndFormat_Mod.FitAndFormat("Bank Details")
    
End Sub
