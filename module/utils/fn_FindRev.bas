Attribute VB_Name = "fn_FindRev"
Option Explicit

Function FindRevWIS(str As String, ByRef txtOriginal As Range)
    
    Application.Volatile False
    FindRevWIS = InStrRev(txtOriginal, str)

    If FindRevWIS = 0 Then FindRevWIS = xlErrNA

End Function

