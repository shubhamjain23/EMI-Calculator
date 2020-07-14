Attribute VB_Name = "swap_Mod"
'This means all variables, object or anything must be declared explicitly
Option Explicit

'Swapping one variabe with another
Sub swap(arr As Variant, var1 As Long, var2 As Long)
    
    Dim temp As Variant

    temp = arr(var1)
    arr(var1) = arr(var2)
    arr(var2) = temp

End Sub
