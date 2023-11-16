Attribute VB_Name = "Atten_CompareStandard"
Option Explicit

Sub sbCompareWithYearStart()

    Call shUnprotect(globalSheetPW)
    Call sbChangeStandardCode(2)
    Range("Atten_CompareArea").Interior.color = RGB(237, 237, 237)
    Call shProtect(globalSheetPW)
    
End Sub

Sub sbCompareWithSameMonth()

    Call shUnprotect(globalSheetPW)
    Call sbChangeStandardCode(1)
    Range("Atten_CompareArea").Interior.color = RGB(255, 243, 203)
    Call shProtect(globalSheetPW)
    
End Sub

Sub sbChangeStandardCode(standardCode As Integer)

    Range("Atten_CompareStandard") = standardCode
    
End Sub
