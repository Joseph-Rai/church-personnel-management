Attribute VB_Name = "z_SubModule"
Option Explicit
Function FileSequence(filePath As String, Optional Sequence As Long = 0)

    Dim Ext As String: Dim Path As String: Dim newPath As String
    Dim Pnt As Long
    
    Pnt = InStrRev(filePath, ".")
    If Pnt <> 0 Then
        Path = Left(filePath, Pnt - 1)
        Ext = Right(filePath, Len(filePath) - Pnt + 1)
    Else
        Path = filePath
        Ext = ""
    End If
    
    newPath = Path & Ext
    
    Do Until FileExists(newPath) = False
        Sequence = Sequence + 1
        newPath = Path & "(" & Sequence & ")" & Ext
    Loop
    
    FileSequence = newPath

End Function
Public Function FileExists(ByVal path_ As String) As Boolean
    
    FileExists = (Dir(path_, vbDirectory) <> "")

End Function
