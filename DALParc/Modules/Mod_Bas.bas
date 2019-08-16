Attribute VB_Name = "Mod_Bas"
Option Explicit

Function SQLText(txt As Variant) As String
    Dim N As Integer
    N = InStr(1, txt, "'")
    If N > 0 Then SQLText = "'" & Left$(txt, N) & SQLText(Right$(txt, Len(txt) - N)) Else SQLText = "'" & txt & "'"
End Function
