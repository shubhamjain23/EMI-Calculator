Attribute VB_Name = "UpdateBanksLbx_Mod"
'This means all variables, object or anything must be declared explicitly
Option Explicit
Sub UpdateBanksLbx()
    
    Call UpdateBanks_Mod.UpdateBanks
    Sheets("Bank Details").Select
    
'    Setting list of list box "banks_lbx" from range A2:(lastrow)2
    With ActiveSheet
        calc_ufm.banks_lbx.RowSource = Range(.Range("A2"), .Cells(Rows.Count, "A").End(xlUp)).Address(, , , True)
    End With
    
    calc_ufm.select_All_Cbx.Value = False
    
End Sub
