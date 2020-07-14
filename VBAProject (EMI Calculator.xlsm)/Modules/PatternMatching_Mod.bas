Attribute VB_Name = "PatternMatching_Mod"
'This means all variables, object or anything must be declared explicitly
Option Explicit

'Function to check if Email entered is valid or not
Function IsValidEmail(sEmailAddress As String) As Boolean
    
    'Code from Officetricks
    Dim sEmailPattern As String
    Dim oRegEx As Object
    Dim bReturn As Boolean
    
    'Use the below regular expressions for email checking
    sEmailPattern = "^\w+([\.-]?\w+)*@\w+([\.-]?\w+)*(\.\w{2,3})+$" 'or
    sEmailPattern = "^([a-zA-Z0-9_\-\.]+)@[a-z0-9-]+(\.[a-z0-9-]+)*(\.[a-z]{2,3})$"
    
    'Create Regular Expression Object
    Set oRegEx = CreateObject("VBScript.RegExp")
    oRegEx.Global = True
    oRegEx.IgnoreCase = True
    oRegEx.Pattern = sEmailPattern
    bReturn = False
    
    'Check if Email match regex pattern
    If oRegEx.Test(sEmailAddress) Then
        'Debug.Print "Valid Email ('" & sEmailAddress & "')"
        bReturn = True
    Else
        'Debug.Print "Invalid Email('" & sEmailAddress & "')"
        bReturn = False
    End If

    IsValidEmail = bReturn
End Function

'Function to check if a particular string matched with the partiuar RegEx pattern
Function IsPatternMatch(regexPattern As String, textToMatch As String) As Boolean

    Dim obRegEx As Object
    Dim bReturn As Boolean
    
'    Create Regular Expression Object
    Set obRegEx = CreateObject("VBScript.RegExp")
    obRegEx.Global = True
    obRegEx.IgnoreCase = True
    obRegEx.Pattern = regexPattern
    bReturn = False
    
'    Check if textToMatch match regex pattern
    If obRegEx.Test(textToMatch) Then
    
'        Debug.Print "Valid Pattern ('" & sEmailAddress & "')"
        bReturn = True
    
    Else
'        Debug.Print "Invalid Pattern('" & sEmailAddress & "')"
        bReturn = False
    End If

    IsPatternMatch = bReturn
End Function
