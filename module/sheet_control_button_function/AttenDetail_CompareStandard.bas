Attribute VB_Name = "AttenDetail_CompareStandard"
Option Explicit

Sub sbCompareWithYearStart_AttenDetail()

    Call shUnprotect(globalSheetPW)
    Call sbChangeStandardCode_AttenDetail(2)
    
    Dim i As Integer
    For i = 0 To Range("AttenDetail_ChurchCount") - 1
        Range("AttenDetail_CompareArea").Offset(i * 12).Interior.color = RGB(237, 237, 237)
    Next
    Call shProtect(globalSheetPW)
    
End Sub

Sub sbCompareWithSameMonth_AttenDetail()

    Call shUnprotect(globalSheetPW)
    Call sbChangeStandardCode_AttenDetail(1)
    
    Dim i As Integer
    For i = 0 To Range("AttenDetail_ChurchCount") - 1
        Range("AttenDetail_CompareArea").Offset(i * 12).Interior.color = RGB(255, 243, 203)
    Next
    Call shProtect(globalSheetPW)
    
End Sub

Sub sbChangeStandardCode_AttenDetail(standardCode As Integer)

    Range("AttenDetail_CompareStandard") = standardCode
    
End Sub

