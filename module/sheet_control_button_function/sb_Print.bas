Attribute VB_Name = "sb_Print"
Option Explicit
'----------------------------------------------------------------------
'������ ��Ȳ ������ ���� �°� �μ�
'----------------------------------------------------------------------
Sub sbPrint_PStaff()

Dim i As Long
Dim j As Long

i = MsgBox("�μ⸦ �����ұ��?", vbYesNo)
If i = vbYes Then
    ActiveWindow.SelectedSheets.PrintOut
End If

Call shUnprotect(globalSheetPW)
Range("PStaff_rngPrint") = Range("PStaff_rngPrint") + 1
Call shProtect(globalSheetPW)

End Sub
