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

    '--//��Ʈ ��й�ȣ ����
    Call shUnprotect(globalSheetPW)
    
    Dim pageHeight As Integer: pageHeight = COUNT_PAGE_HEIGHT_CELLS
    Dim pageWidth As Integer: pageWidth = COUNT_PAGE_WIDTH_CELLS
    
    Dim lineHeight As Integer: lineHeight = (COUNT_PAGE_HEIGHT_CELLS - 1) / 3 '--//���� 1 ���� 3���� ����
    
    '--//�׷� ��ġ��
'    ActiveSheet.Outline.ShowLevels RowLevels:=2
    
    '--//�����̸� ���̱�/�����
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
    
    '--//��Ʈ ��й�ȣ ���
    Call shProtect(globalSheetPW)
    
    Application.ScreenUpdating = True

End Sub
