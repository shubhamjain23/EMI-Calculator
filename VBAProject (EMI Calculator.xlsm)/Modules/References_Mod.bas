Attribute VB_Name = "References_Mod"
'This means all variables, object or anything must be declared explicitly
Option Explicit

'Function to check if a particular reference is loaded in project or not
Function isReferenceLoaded(referenceName As String) As Boolean

    Dim xRef As Variant
    Dim bReturn As Boolean
    bReturn = False
'    "ThisWorkbook.VBProject.References" works only if user has given access to VBA project object model in Trust Center Settings
    For Each xRef In ThisWorkbook.VBProject.References
        If (PatternMatching_Mod.IsPatternMatch(referenceName, xRef.Description)) = True Then
            bReturn = True
            Exit For
        End If
    Next xRef
    
    isReferenceLoaded = bReturn
    
End Function

'Function to check if a particular reference is loaded. If loaded, return that reference name
Function returnReference(referenceDescription As String) As String

    Dim xRef As Variant
    Dim refReturn As String
    refReturn = ""
'    "ThisWorkbook.VBProject.References" works only if user has given access to VBA project object model in Trust Center Settings
    For Each xRef In ThisWorkbook.VBProject.References
        If (PatternMatching_Mod.IsPatternMatch(referenceDescription, xRef.Description)) = True Then
            refReturn = xRef.Name
            Exit For
        End If
    Next xRef
    
    returnReference = refReturn
    
End Function
