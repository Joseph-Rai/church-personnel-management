Attribute VB_Name = "sb_space"
Option Explicit

Public Function space(cnt As Integer) As String

    Dim str As String
    Dim i As Integer
    
    For i = 1 To cnt
        str = str & " "
    Next
    
    space = str

End Function
