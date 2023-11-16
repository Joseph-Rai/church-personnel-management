Attribute VB_Name = "sb_ShowEngNameInPStaffReport"
Option Explicit
Sub ShowEngNameInPStaffReport()
Attribute ShowEngNameInPStaffReport.VB_ProcData.VB_Invoke_Func = " \n14"
    
    ArrangeGroupingInPStaffReport Range("C6")
    
End Sub

Sub ShowVisaInPStaffReport()

    ArrangeGroupingInPStaffReport Range("C7")

End Sub

Sub ArrangeGroupingInPStaffReport(ByRef targetRange As Range)

    Application.ScreenUpdating = False

    '--//시트 비밀번호 해제
    Call shUnprotect(globalSheetPW)
    
    Dim pageHeight As Integer: pageHeight = COUNT_PAGE_HEIGHT_CELLS
    Dim pageWidth As Integer: pageWidth = COUNT_PAGE_WIDTH_CELLS
    
    Dim lineHeight As Integer: lineHeight = (COUNT_PAGE_HEIGHT_CELLS - 1) / 3 '--//제목 1 빼고 3으로 나눔
    
    '--//그룹 펼치고
'    ActiveSheet.Outline.ShowLevels RowLevels:=2
    
    '--//영문이름 보이기/숨기기
    Dim i As Integer
On Error Resume Next
    If targetRange.EntireRow.Hidden = True Then
        ActiveSheet.Outline.ShowLevels RowLevels:=2
        For i = 0 To 2
            targetRange.Offset(lineHeight * i).EntireRow.Ungroup
        Next
        ActiveSheet.Outline.ShowLevels RowLevels:=1
    Else
        ActiveSheet.Outline.ShowLevels RowLevels:=2
        For i = 0 To 2
            targetRange.Offset(lineHeight * i).EntireRow.Ungroup
            targetRange.Offset(lineHeight * i).EntireRow.Group
        Next
        ActiveSheet.Outline.ShowLevels RowLevels:=1
    End If
    
On Error GoTo 0
    
    '--//시트 비밀번호 잠금
    Call shProtect(globalSheetPW)
    
    Application.ScreenUpdating = True

End Sub
