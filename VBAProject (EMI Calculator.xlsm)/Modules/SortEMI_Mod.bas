Attribute VB_Name = "SortEMI_Mod"
'This means all variables, object or anything must be declared explicitly
Option Explicit

Sub SortEMI()

    Call CalcEMI_Mod.CalcEMI
    
'    Declaring arrays with no dimension
    Dim banksName() As Variant, interestRate() As Variant
    Dim processingChargesPercent() As Variant, EMI() As Variant, processingCharges() As Variant
    
    Dim i As Long, lastRow As Long

    Sheets("Selected Banks").Select
'    Select all, Move up from the last row to find the last row where data is present
    lastRow = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row

'    Redeclaring arrays with dimension
    ReDim banksName(1 To lastRow)
    ReDim interestRate(1 To lastRow)
    ReDim processingChargesPercent(1 To lastRow)
    ReDim EMI(1 To lastRow)
    ReDim processingCharges(1 To lastRow)

'    Loop to select bank name, EMI and other columns from current worksheet
    For i = 1 To lastRow
        With ActiveSheet
            banksName(i) = .Range("A" & i + 1)
            interestRate(i) = .Range("B" & (i + 1))
            processingChargesPercent(i) = .Range("C" & (i + 1))
            EMI(i) = .Range("D" & (i + 1))
            processingCharges(i) = .Range("E" & (i + 1))
        End With
    Next i

'    Sorting all values (all rows) but only by EMI
    Call Quicksort(EMI(), LBound(EMI), UBound(EMI), banksName(), interestRate(), processingChargesPercent(), processingCharges())
    
    Call createSheet_Mod.createSheet("Sorted EMI")
    
    Sheets("Sorted EMI").Select
    With ActiveSheet
        .Cells.ClearContents
'        Setting first row of current sheet
        .Cells(1, 1).Value = "Banks"
        .Cells(1, 2).Value = "Interest Rate (%)"
        .Cells(1, 3).Value = "Processing Charges (%)"
        .Cells(1, 4).Value = "EMI"
        .Cells(1, 5).Value = "Processing Charges"
    End With
    
    Dim j As Long
    For j = 1 To lastRow - 1
        With ActiveSheet
'            Populating "Sorted EMI" sheet from 2nd row
            .Range("A" & (j + 1)) = banksName(j + 1)
            .Range("B" & (j + 1)) = interestRate(j + 1)
            .Range("C" & (j + 1)) = processingChargesPercent(j + 1)
            .Range("D" & (j + 1)) = EMI(j + 1)
            .Range("E" & (j + 1)) = processingCharges(j + 1)
        End With
    Next j

    Call fitAndFormat_Mod.FitAndFormat("Sorted EMI")

    End Sub

