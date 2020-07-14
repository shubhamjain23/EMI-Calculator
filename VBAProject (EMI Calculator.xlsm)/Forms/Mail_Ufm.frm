VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Mail_Ufm 
   Caption         =   "Mail Options"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5535
   OleObjectBlob   =   "Mail_Ufm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Mail_Ufm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This means all variables, object or anything must be declared explicitly
Option Explicit

Private Sub send_btn_Click()
    
'    Check if mail entered is valid or not
    If (PatternMatching_Mod.IsValidEmail(Mail_Ufm.mailto_tbx.Value) = False) Then
    
        MsgBox "Please enter a valid email address"
    Else
        MsgBox ("Microsoft Outlook must be configured for mail options.")
        Call Email_Mod.EmailReport
        
'        Hide Mail_Ufm to activate previous form
        Mail_Ufm.Hide
        
    End If
    
End Sub


Private Sub UserForm_Initialize()

'    Setting subject at Mail_Ufm userform initialize
    Mail_Ufm.subject_tbx.Value = "EMI Report from " & calc_ufm.name_tbx.Value & "."
    
End Sub
