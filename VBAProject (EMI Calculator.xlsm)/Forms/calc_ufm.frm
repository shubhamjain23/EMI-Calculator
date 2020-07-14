VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} calc_ufm 
   Caption         =   "EMI Calculator"
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7425
   OleObjectBlob   =   "calc_ufm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "calc_ufm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub calc_EMI_btn_Click()
    
    Call CalcEMI_Mod.CalcEMI
    
End Sub

Private Sub loan_amt_tbx_AfterUpdate()

    If (IsEmpty(loan_amt_tbx.Value) Or (loan_amt_tbx.Value = "") Or (loan_amt_tbx.Value = 0)) Then
        loan_amt_tbx.Value = 1000
    End If
    
End Sub

Private Sub loan_amt_tbx_Change()

    If (IsEmpty(loan_amt_tbx.Value) Or (loan_amt_tbx.Value = "") Or (loan_amt_tbx.Value = 0)) Then
        submit_btn.Enabled = False
        calc_EMI_btn.Enabled = False
        sort_EMI_btn.Enabled = False
        make_graphs_btn.Enabled = False
    Else
        If Not (IsEmpty(name_tbx.Value) Or (name_tbx.Value = "")) Then
            submit_btn.Enabled = True
            calc_EMI_btn.Enabled = True
            sort_EMI_btn.Enabled = True
            make_graphs_btn.Enabled = True
        End If
    End If

End Sub

Private Sub make_graphs_btn_Click()

    Call CreateGraph_Mod.EMIGraphs

End Sub

Private Sub name_tbx_Change()
    If (IsEmpty(name_tbx.Value) Or (name_tbx.Value = "")) Then
        submit_btn.Enabled = False
        calc_EMI_btn.Enabled = False
        sort_EMI_btn.Enabled = False
        make_graphs_btn.Enabled = False
    Else
        submit_btn.Enabled = True
        calc_EMI_btn.Enabled = True
        sort_EMI_btn.Enabled = True
        make_graphs_btn.Enabled = True
    End If
End Sub

Private Sub name_tbx_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    'Limiting characters to 40
    If Not (Len(tenor_tbx.Value) < 40) Then
        KeyAscii = 0
    End If
    Select Case KeyAscii
        '32 = [SPACE], 97-122 = a-z, 65 to 90 = A-Z, 48 to 57 = 0-9, 46 = . , 64 = @
        Case 97 To 122, 65 To 90, 32
            
        Case Else
            KeyAscii = 0
    End Select
    
End Sub

Private Sub loan_amt_tbx_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Not (Len(tenor_tbx.Value) < 10) Then
        KeyAscii = 0
    End If
    Select Case KeyAscii
        '32 = [SPACE], 97-122 = a-z, 65 to 90 = A-Z, 48 to 57 = 0-9, 46 = . , 64 = @
        Case 48 To 57
            
        Case Else
            KeyAscii = 0
    End Select
End Sub

'Subroutine to select and deselect all list box entries
Private Sub select_All_Cbx_Click()

    Dim i As Integer
    If (select_All_Cbx.Value = True) Then
        For i = 0 To banks_lbx.ListCount - 1
            banks_lbx.Selected(i) = True
        Next i
    Else
        For i = 0 To banks_lbx.ListCount - 1
            banks_lbx.Selected(i) = False
        Next i
    End If
End Sub

Private Sub sort_EMI_btn_Click()

    Call SortEMI_Mod.SortEMI
    
End Sub

Private Sub tenor_tbx_AfterUpdate()

    If (IsEmpty(tenor_tbx.Value) Or (tenor_tbx.Value = "") Or (tenor_tbx.Value = 0)) Then
        tenor_tbx.Value = 5
    End If
    
End Sub

Private Sub tenor_tbx_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    
'    Limiting tenor value to 3 digits
    If Not (Len(tenor_tbx.Value) < 3) Then
        KeyAscii = 0
    End If
    Select Case KeyAscii
        '32 = [SPACE], 97-122 = a-z, 65 to 90 = A-Z, 48 to 57 = 0-9, 46 = . , 64 = @
        Case 48 To 57
            
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub submit_btn_Click()
    
    If (email_report_rdbtn.Value = True) Then
        Mail_Ufm.Show
    Else
    If (print_report_rdbtn.Value = True) Then
        Call Print_Mod.PrintReport
    Else
    If (save_report_rdbtn.Value = True) Then
        Call SaveReport_Mod.SaveReport
    End If
    End If
    End If
    
End Sub

Private Sub tenor_tbx_Change()
    
    Dim x As Single
    
    If (IsEmpty(tenor_tbx.Value) Or (tenor_tbx.Value = "")) Then
        tenor_months_lbl.Caption = "0 month"
        submit_btn.Enabled = False
        calc_EMI_btn.Enabled = False
        sort_EMI_btn.Enabled = False
        make_graphs_btn.Enabled = False
    Else
'        Convert text to number
        x = CSng(tenor_tbx.Value)
    End If
    
    If (x = 0) Then
        tenor_months_lbl.Caption = x & " month"
    Else
        tenor_months_lbl.Caption = x * 12 & " months"
        
        If Not (IsEmpty(name_tbx.Value) Or (name_tbx.Value = "")) Then
            submit_btn.Enabled = True
            calc_EMI_btn.Enabled = True
            sort_EMI_btn.Enabled = True
            make_graphs_btn.Enabled = True
        End If

    End If
    
End Sub

Private Sub update_banks_btn_Click()

    Call UpdateBanksLbx_Mod.UpdateBanksLbx
    
End Sub

Private Sub UserForm_Initialize()

'    Check if Oulook library present or not
    If (References_Mod.isReferenceLoaded("Microsoft Outlook \d+\.\d+ Object Library")) = False Then
        MsgBox ("Microsoft Outlook is not installed or 'Outlook' library is not present" & vbCrLf & vbCrLf & vbCrLf & "Mail Options won't work")
        calc_ufm.email_report_rdbtn.Enabled = False
    Else
        calc_ufm.email_report_rdbtn.Enabled = True
    End If
    
'    Code for removing any reference which starts with MISSING
    Dim reurnedRef As String
    Dim ref As Variant
    returnedRef = References_Mod.returnReference("^MISSING")
    If Not (returnedRef = "") Then
    
'        "ThisWorkbook.VBProject.References" works only if user has given access to VBA project object model in Trust Center Settings
        Set ref = ThisWorkbook.VBProject.References(returnedRef)
        
'        Removing reference
        ThisWorkbook.VBProject.References.Remove ref

    End If
    
    Call UpdateBanksLbx_Mod.UpdateBanksLbx
    
End Sub
