Attribute VB_Name = "UserFormMethodCaller"
Option Explicit

Public Sub call_SearchByTitlePosition_InsertPic()

    Dim ws As Worksheet
    
    Call shUnprotect(globalSheetPW)

    Call frm_Search_by_TitlePosition.sbInsertPic
    
    Call shProtect(globalSheetPW)

End Sub
