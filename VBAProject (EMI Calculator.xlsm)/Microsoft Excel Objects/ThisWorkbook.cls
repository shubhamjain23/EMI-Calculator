VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()

'    MsgBox ("Make sure 'Trust Access to the VBA project object model' is ticked." & vbCrLf & vbCrLf & "File " & Chr(26) & " Options " & Chr(26) & " Trust Center " & Chr(26) & " Trust Center Settings " & Chr(26) & " Macro Settings")
    MsgBox ("Make sure 'Trust Access to the VBA project object model' is ticked." & vbCrLf & vbCrLf & "File -> Options -> Trust Center -> Trust Center Settings -> Macro Settings")

End Sub
