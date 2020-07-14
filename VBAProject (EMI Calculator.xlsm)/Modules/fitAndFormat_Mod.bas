Attribute VB_Name = "fitAndFormat_Mod"
'This means all variables, object or anything must be declared explicitly
Option Explicit
Sub FitAndFormat(sheetName As String)

    Sheets(sheetName).Select
    With ActiveSheet
    
'       Setting all coumns and all rows to fit by the text present in them
        .Cells.EntireColumn.AutoFit
        .Cells.EntireRow.AutoFit
        
'       Selecting 1st row from A1 to last column
        .Range("A1").Select
        .Range(Selection, Selection.End(xlToRight)).Select
        
        With Selection
        
'           Setting font color to white
            .Font.ThemeColor = xlThemeColorDark1
            
'           Setting background color to black
            .Interior.ThemeColor = xlThemeColorLight1
            
        End With

    End With
End Sub
