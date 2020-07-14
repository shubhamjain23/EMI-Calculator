Attribute VB_Name = "CalcEMI_Mod"
'This means all variables, object or anything must be declared explicitly
'Option Base 1 : Change default base for arrays from 0 to 1
Option Explicit

Sub CalcEMI()
    
    Dim tenor_month As Long, loan_amount As Long
    
'   Coverting year value in months
    tenor_month = calc_ufm.tenor_tbx.Value * 12
    loan_amount = calc_ufm.loan_amt_tbx.Value
    
    Dim lastRow, i, r, c, lastRowAddress, lastColumnAddress, bank_details As Variant
    
    Sheets("Bank Details").Select
    
'   Populating 4th and 5th column as it is not present in source file, and also to find last column
    With ActiveSheet
        .Cells(1, 4).Value = "EMI"
        .Cells(1, 5).Value = "Processing Charges"
    End With

'   Select all, Move up from the last row to find the last row where data is present
    lastRow = Range("A10000").End(xlUp).Row
    
'   Populating EMI and Processing Charges by given formulae
    For i = 2 To lastRow
        Cells(i, 4).Value = "=-PMT(RC[-2]/100/12, " & tenor_month & "," & loan_amount & ",0,0)"
        Cells(i, 4).NumberFormat = "0.00"
        Cells(i, 5).Value = "=RC[-2]/100* " & loan_amount & ""
        Cells(i, 5).NumberFormat = "0.00"
    Next i
    
    Call fitAndFormat_Mod.FitAndFormat("Bank Details")
    
'   Finding last row and column addresses
    lastRowAddress = Range("A10000").End(xlUp).Address
    lastColumnAddress = Range("XFD1").End(xlToLeft).Address
    
'   Assign range to a variabe
    bank_details = Range(lastRowAddress, lastColumnAddress)
    
    Call createSheet_Mod.createSheet("Selected Banks")
    
    Sheets("Selected Banks").Select
    
'   Clearing contents (not first row as it has headers)
    Rows("2:" & Rows.Count).ClearContents
    
'   Populating Headers
    With ActiveSheet
        .Cells(1, 1).Value = "Banks"
        .Cells(1, 2).Value = "Interest Rate (%)"
        .Cells(1, 3).Value = "Processing Charges (%)"
        .Cells(1, 4).Value = "EMI"
        .Cells(1, 5).Value = "Processing Charges"
    End With
    
    r = 2
    c = 1
    
'   Loop to populate values in Selected Banks worsheet by their corresponding bank name using vlookup
    For i = 0 To calc_ufm.banks_lbx.ListCount - 1
        If calc_ufm.banks_lbx.Selected(i) = True Then
            With ActiveSheet
                .Cells(r, c).Value = calc_ufm.banks_lbx.List(i)
                .Cells(r, c + 1).Value = Application.WorksheetFunction.VLookup(calc_ufm.banks_lbx.List(i), [bank_details], 2, 0)
                .Cells(r, c + 2).Value = Application.WorksheetFunction.VLookup(calc_ufm.banks_lbx.List(i), [bank_details], 3, 0)
                .Cells(r, c + 3).Value = Application.WorksheetFunction.VLookup(calc_ufm.banks_lbx.List(i), [bank_details], 4, 0)
                .Cells(r, c + 4).Value = Application.WorksheetFunction.VLookup(calc_ufm.banks_lbx.List(i), [bank_details], 5, 0)
            End With
            r = r + 1
        End If
    Next i
    
    Call fitAndFormat_Mod.FitAndFormat("Selected Banks")
    
End Sub
