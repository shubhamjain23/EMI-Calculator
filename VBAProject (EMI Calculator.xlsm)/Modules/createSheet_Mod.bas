Attribute VB_Name = "createSheet_Mod"
'This means all variables, object or anything must be declared explicitly
Option Explicit

Sub createSheet(sheetName As String)

    Dim wsTest As Worksheet
'   When Nothing is assigned to an object, variable no longer refers to any actual object
    Set wsTest = Nothing
    
'   If error, execute the next line
    On Error Resume Next
    Set wsTest = ActiveWorkbook.Worksheets(sheetName)
    
'   If error, show excel runtime error
    On Error GoTo 0
    
'   If worksheet not present
    If wsTest Is Nothing Then
    
'   Add worksheet
        Worksheets.Add.Name = sheetName
    End If

End Sub
