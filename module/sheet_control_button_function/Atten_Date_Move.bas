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
    noMax = WorksheetFunction.Max(Range("F17:R17")) '--//학생이상 1회출석 최대값
    noMin = WorksheetFunction.Min(Range("F19:R19")) '--//학생이상 4회출석 최소값
    i = 1: j = 1
      

      
      '--//출석 그래프 차트에서
      With Sheets("교회별 출석현황").ChartObjects(1).Chart.Axes(xlValue)
        '--//식구 규모에 따라 스케일을 달리 합니다.
        Select Case noMax
            Case Is <= 100: Term = 10
            Case Is <= 500: Term = 50
            Case Is <= 1000: Term = 100
            Case Else: Term = 100
        End Select
        
        '--//범위의 최대값을 구합니다..
        Do
            If Term * i > noMax Then
                .MaximumScale = Term * i
                Exit Do
            End If
            i = i + 1
        Loop
        
        '--//범위의 최소값을 구합니다.
        Do
            If Term * j >= noMin * 0.9 Then
                .MinimumScale = Term * (j - 1)
                Exit Do
            End If
            j = j + 1
        Loop
        
        '--//범위의 최대값과 최소값의 차이가 4의 배수가 아니면 최대값 수정
        Do
            If (.MaximumScale - .MinimumScale) Mod 4 = 0 Then Exit Do
            i = i + 1
            .MaximumScale = Term * i
        Loop
        
        .MajorUnit = (.MaximumScale - .MinimumScale) / 4
        
      End With

      With Sheets("교회별 출석현황").ChartObjects(1).Chart.Axes(xlValue, xlSecondary)
'        .MaximumScale = Application.Max(WorksheetFunction.RoundUp(WorksheetFunction.Max(Range("F26:R26")), 1), WorksheetFunction.RoundUp(WorksheetFunction.Max(Range("F27:R27")), 1))
'        .MinimumScale = Application.Min(WorksheetFunction.RoundDown(WorksheetFunction.Min(Range("F26:R26")), 1), WorksheetFunction.RoundDown(WorksheetFunction.Min(Range("F27:R27")), 1))
        .MaximumScale = 3
        .MinimumScale = 0
      End With
      
On Error GoTo 0

End Sub
