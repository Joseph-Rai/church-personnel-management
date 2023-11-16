Attribute VB_Name = "sb_Print"
Option Explicit
'----------------------------------------------------------------------
'선지자 현황 페이지 수에 맞게 인쇄
'----------------------------------------------------------------------
Sub sbPrint_PStaff()

Dim i As Long
Dim j As Long

i = MsgBox("인쇄를 시작할까요?", vbYesNo)
If i = vbYes Then
    ActiveWindow.SelectedSheets.PrintOut
End If

Call shUnprotect(globalSheetPW)
Range("PStaff_rngPrint") = Range("PStaff_rngPrint") + 1
Call shProtect(globalSheetPW)

End Sub
