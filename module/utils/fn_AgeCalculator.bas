Attribute VB_Name = "fn_AgeCalculator"
Option Explicit

'--//만 나이 계산기
Public Function CalculateOnlyAge(ByVal argBirthday As Date) As Integer

    Dim yearDiff As Integer
    
    yearDiff = DateDiff("YYYY", argBirthday, Date)
    
    If WorksheetFunction.EDate(Date, -1 * yearDiff) > argBirthday Then
        CalculateOnlyAge = yearDiff
    Else
        CalculateOnlyAge = yearDiff - 1
    End If

End Function

'--//한국 나이 계산기
Public Function CalculateKoreanAge(ByVal argBirthday As Date) As Integer

    CalculateKoreanAge = DateDiff("YYYY", argBirthday, Date) + 1

End Function
