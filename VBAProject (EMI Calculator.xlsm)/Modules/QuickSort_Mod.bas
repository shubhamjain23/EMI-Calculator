Attribute VB_Name = "QuickSort_Mod"
'This means all variables, object or anything must be declared explicitly
Option Explicit

'Sorts a one-dimensional VBA array from smallest to largest
'Fastest sorting algorithm for smaller arrays
Sub Quicksort(EMI As Variant, arrLbound As Long, arrUbound As Long, banksName As Variant, interestRate As Variant, processingChargesPercent As Variant, processingCharges As Variant)

    Dim pivotVal As Variant
    Dim swap    As Variant
    Dim swap1    As Variant
    Dim swap2    As Variant
    Dim swap3    As Variant
    Dim swap4    As Variant
    Dim tmpLow   As Long
    Dim tmpHi    As Long
 
    tmpLow = arrLbound
    tmpHi = arrUbound
    
'    Choosing pivot as the middle value of the array
    pivotVal = EMI((arrLbound + arrUbound) \ 2)
 
    While (tmpLow <= tmpHi) 'divide
        While (EMI(tmpLow) < pivotVal And tmpLow < arrUbound)
            tmpLow = tmpLow + 1
        Wend
  
        While (pivotVal < EMI(tmpHi) And tmpHi > arrLbound)
            tmpHi = tmpHi - 1
        Wend
 
        If (tmpLow <= tmpHi) Then
            
'            Swapping all column values together Or swapping one row with another
            Call swap_Mod.swap(EMI, tmpLow, tmpHi)
            Call swap_Mod.swap(interestRate, tmpLow, tmpHi)
            Call swap_Mod.swap(processingChargesPercent, tmpLow, tmpHi)
            Call swap_Mod.swap(processingCharges, tmpLow, tmpHi)
            Call swap_Mod.swap(banksName, tmpLow, tmpHi)
    
            tmpLow = tmpLow + 1
            tmpHi = tmpHi - 1
            
        End If
    Wend
    
'    Run again for the arrays divided by pivot value
    If (arrLbound < tmpHi) Then Quicksort EMI, arrLbound, tmpHi, banksName, interestRate, processingChargesPercent, processingCharges 'conquer
    If (tmpLow < arrUbound) Then Quicksort EMI, tmpLow, arrUbound, banksName, interestRate, processingChargesPercent, processingCharges 'conquer
    
End Sub
