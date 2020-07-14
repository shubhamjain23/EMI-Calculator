Attribute VB_Name = "CreateGraph_Mod"
'This means all variables, object or anything must be declared explicitly
Option Explicit
Sub EMIGraphs()
    
    Call CalcEMI_Mod.CalcEMI
    
'   Declaring arrays with no dimension
    Dim banksName() As String, EMI() As Variant
    Dim lastRow As Long
    Sheets("Selected Banks").Select

'   Select all, Move up from the last row to find the last row where data is present
    lastRow = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row
    
    Dim i As Long
    
'   Redeclaring arrays with dimension
    ReDim banksName(1 To lastRow)
    ReDim EMI(1 To lastRow)
    
'   Loop to select bank name and EMI from current worksheet
    For i = 1 To lastRow
        With ActiveSheet
            banksName(i) = .Range("A" & i)
            EMI(i) = .Range("D" & (i))
        End With
    Next i

    Call createSheet_Mod.createSheet("EMI Graphs")
    
    Sheets("EMI Graphs").Select
'   Clearing contents
    ActiveSheet.Cells.ClearContents
    
    Dim chtObj As Object

'   Deleting all charts and their objects if present from current sheet
    For Each chtObj In ActiveSheet.ChartObjects
        chtObj.Delete
    Next
    
'   Pasting data
    For i = 1 To lastRow
        With ActiveSheet
            .Range("A" & (i)) = banksName(i)
            .Range("B" & (i)) = EMI(i)
        End With
    Next i
    
    Call fitAndFormat_Mod.FitAndFormat("EMI Graphs")
    
    Dim lastRow1 As Long, chartRange As Object, chartArea As Object
    
    lastRow1 = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row
    
'   Dynamically selecting chart range
    Set chartRange = ActiveSheet.Range("A1:B" & lastRow1)
    
'   Suppose lastRow = 10, then chartArea range = A13:H30
'   Used H row as it is the width limit of one A4 sheet for printing
    Set chartArea = Range("A" & lastRow1 + 3 & ":H" & lastRow1 + 20)
  
'   Adding chart
    ActiveSheet.Shapes.AddChart.Select
    With ActiveChart
'       Setting chart type
        .ChartType = xlColumnClustered
'       Setting chart range
        .SetSourceData Source:=chartRange
'       Changing chart element color to orange
        .ChartStyle = 8
'       Adding Rotated Vertical Axis
        .SetElement (msoElementPrimaryValueAxisTitleRotated)
'       Changing value of axis to "EMI"
        .Axes(xlValue, xlPrimary).AxisTitle.text = "EMI"
'       Selecting chart title and changing its value
        .ChartTitle.Select
        .ChartTitle.text = "EMI Comparison"
    End With

'   Deleting gridlines from chart
    ActiveChart.Axes(xlValue).MajorGridlines.Select
    Selection.Delete
    
'   Deleting legend from chart
    ActiveChart.Legend.Select
    Selection.Delete
    
    ActiveChart.chartArea.Select

    Dim chartObj As Object
'   Assigning chart object to active chart
    Set chartObj = ActiveChart.Parent
    
'   Setting position of chart (Top, Left, Width, Height)
    chartObj.Top = chartArea.Top
    chartObj.Left = chartArea.Left
    chartObj.Width = chartArea.Width
    chartObj.Height = chartArea.Height
    
End Sub

