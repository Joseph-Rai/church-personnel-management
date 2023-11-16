Attribute VB_Name = "Atten_Date_Move"
Option Explicit

Sub sbMoveRight_Atten()

    If Range("Atten_rngDate") < Range("Atten_MaxDate") Then
        Range("Atten_rngDate") = WorksheetFunction.EDate(Range("Atten_rngDate"), 1)
    End If
    Call shUnprotect(globalSheetPW)
    Call sbArrangeChart_Atten
    Call shProtect(globalSheetPW)
End Sub

Sub sbMoveLeft_Atten()

    If WorksheetFunction.EDate(Range("Atten_rngDate"), -12) > Range("Atten_MinDate") Then
        Range("Atten_rngDate") = WorksheetFunction.EDate(Range("Atten_rngDate"), -1)
    End If
    Call shUnprotect(globalSheetPW)
    Call sbArrangeChart_Atten
    Call shProtect(globalSheetPW)
End Sub

Public Sub sbArrangeChart_Atten()

Dim noMax As Integer
    Dim noMin As Integer
    Dim i As Long
    Dim j As Long
    Dim Term As Long
    
On Error Resume Next
    noMax = WorksheetFunction.Max(Range("F17:R17")) '--//�л��̻� 1ȸ�⼮ �ִ밪
    noMin = WorksheetFunction.Min(Range("F19:R19")) '--//�л��̻� 4ȸ�⼮ �ּҰ�
    i = 1: j = 1
      

      
      '--//�⼮ �׷��� ��Ʈ����
      With Sheets("��ȸ�� �⼮��Ȳ").ChartObjects(1).Chart.Axes(xlValue)
        '--//�ı� �Ը� ���� �������� �޸� �մϴ�.
        Select Case noMax
            Case Is <= 100: Term = 10
            Case Is <= 500: Term = 50
            Case Is <= 1000: Term = 100
            Case Else: Term = 100
        End Select
        
        '--//������ �ִ밪�� ���մϴ�..
        Do
            If Term * i > noMax Then
                .MaximumScale = Term * i
                Exit Do
            End If
            i = i + 1
        Loop
        
        '--//������ �ּҰ��� ���մϴ�.
        Do
            If Term * j >= noMin * 0.9 Then
                .MinimumScale = Term * (j - 1)
                Exit Do
            End If
            j = j + 1
        Loop
        
        '--//������ �ִ밪�� �ּҰ��� ���̰� 4�� ����� �ƴϸ� �ִ밪 ����
        Do
            If (.MaximumScale - .MinimumScale) Mod 4 = 0 Then Exit Do
            i = i + 1
            .MaximumScale = Term * i
        Loop
        
        .MajorUnit = (.MaximumScale - .MinimumScale) / 4
        
      End With

      With Sheets("��ȸ�� �⼮��Ȳ").ChartObjects(1).Chart.Axes(xlValue, xlSecondary)
'        .MaximumScale = Application.Max(WorksheetFunction.RoundUp(WorksheetFunction.Max(Range("F26:R26")), 1), WorksheetFunction.RoundUp(WorksheetFunction.Max(Range("F27:R27")), 1))
'        .MinimumScale = Application.Min(WorksheetFunction.RoundDown(WorksheetFunction.Min(Range("F26:R26")), 1), WorksheetFunction.RoundDown(WorksheetFunction.Min(Range("F27:R27")), 1))
        .MaximumScale = 3
        .MinimumScale = 0
      End With
      
On Error GoTo 0

End Sub
