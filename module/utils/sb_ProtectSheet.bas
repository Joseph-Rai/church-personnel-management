Attribute VB_Name = "sb_ProtectSheet"
Option Explicit
'------------------------------------
'��Ʈ��ȣ
'------------------------------------
Public Sub shProtect(PW As String)
    ActiveSheet.Protect PW
End Sub
Public Sub shUnprotect(PW As String)
    ActiveSheet.Unprotect PW
End Sub
