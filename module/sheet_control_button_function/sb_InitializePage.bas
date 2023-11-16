Attribute VB_Name = "sb_InitializePage"
Option Explicit

Public Sub InitializeA3AppointmentPage()

    frm_Search_Appointment.InitializeReportPage

End Sub

Public Sub initializeAttendanceDetailPage()

    Call shUnprotect(globalSheetPW)
    frm_Search_AttendanceDetail.initCurrentPage
    Range("AttenDetail_ChurchCount") = 1
    frm_Search_AttendanceDetail.attenDetailInsertPicture
    Call shProtect(globalSheetPW)
    
End Sub

Public Sub initializeTitlePositionPage()

    Call shUnprotect(globalSheetPW)
    Call frm_Search_by_TitlePosition.sbInitialize_From
    ActiveSheet.Pictures.Delete
    Call shProtect(globalSheetPW)

End Sub
