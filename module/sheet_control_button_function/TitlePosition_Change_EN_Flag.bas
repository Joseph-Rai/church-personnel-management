Attribute VB_Name = "TitlePosition_Change_EN_Flag"
Option Explicit

Sub changeEnFlag_TitlePosition()

    '--//Flag: 0: 한글 / 1: 영어

    Dim flag As Range
    
    Set flag = Range("TitlePosition_EnFlag")
    
    Call shUnprotect(globalSheetPW)
    
    If flag = 0 Then
        flag = 1
    Else
        flag = 0
    End If
    
    Call shProtect(globalSheetPW)

End Sub
