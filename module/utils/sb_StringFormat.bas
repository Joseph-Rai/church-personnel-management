Attribute VB_Name = "sb_StringFormat"
Option Explicit

Public Function StringFormat(ByVal mask As String, ParamArray tokens()) As String

    Dim i As Long
    For i = LBound(tokens) To UBound(tokens)
        mask = Replace(mask, "{" & i & "}", tokens(i))
    Next
    StringFormat = mask

End Function
