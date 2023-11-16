Attribute VB_Name = "a_Ribbon"
Option Explicit

'-----------------------------------------------------
'  리본 메뉴의 Button ID에 대한 처리 프로시저
'-----------------------------------------------------
Sub run_RibbonControl_WIS(Button As Office.IRibbonControl)
    Select Case Button.Id
        '//프로그램
        Case "btnUpdate":   Call frm_Update_Appointment_Show
        Case "Update_Attendance":   Call frm_Update_Attendance_Show
        Case "Update_Transfer": frm_Update_Appointment.optTransfer.Value = True:    Call frm_Update_Appointment_Show
        Case "Update_Title":    frm_Update_Appointment.optTitle.Value = True:    Call frm_Update_Appointment_Show
        Case "Update_Position": frm_Update_Appointment.optPosition.Value = True:    Call frm_Update_Appointment_Show
        Case "Update_Theological": Call frm_Update_Theological_Show
        Case "Update_PStaff":   Call frm_Update_PInformation_Show
        Case "Update_Hostory_Church":   Call frm_Update_History_Show
        Case "Update_Church_Matching":   Call frm_Update_Church_Esta_Show
        Case "Update_Flight_Schedule":   Call frm_Update_Flight_Show
        Case "Update_BCManager": Call frm_Update_BCLeader_Show
        Case "Search_phone": Call frm_Search_Phone_Show
        Case "Update_Union":    Call frm_Update_Union_Show
        Case "Update_Sermon":    Call frm_Update_Sermon_Show
        Case "Update_Visa":    Call frm_Update_Visa_Show
        Case "Update_Counsel":    Call frm_Update_Counsel_Show
        Case "UserSettings":    Call frm_Update_User_Show
        Case "UserAuthority":    Call frm_Update_User_Authority_Show
                      
        '//공통
        Case "LogIn":     Call LogIn
        Case "LogOut":     Call LogOut
        Case "AddinUninstall":     Call AddinUninstall
        
        Case Else:     Call RibbonButton_Error(Button.Id)
    End Select
End Sub

'-------------------------------------------------------------------
'  Button ID에 대한 처리 프로시저가 없는 경우 오류 메시지
'-------------------------------------------------------------------
Sub RibbonButton_Error(sbID As String)
   MsgBox "선택하신 메뉴(" & sbID & ")는 아직 준비가 되어 있지 않습니다.", vbCritical, banner
End Sub

'-----------
'  로그인
'-----------
Sub LogIn()
    If checkLogin = 1 Then
        MsgBox Application.UserName & "님 이미 로그인 되어 있습니다.", vbInformation, banner
        Exit Sub
    End If
    f_login.Show
    
End Sub

'------------
'  로그아웃
'------------
Sub LogOut()
    
    '--//선지자 상세정보 시트 조정
    Sheets("선지자 상세정보").Visible = False
    Sheets("A3인사발령").Visible = False
    
    If checkLogin = 0 Then
        MsgBox Application.UserName & "님 이미 로그아웃 되어 있습니다.", vbInformation, banner
        Exit Sub
    End If
    checkLogin = 0 '로그아웃 상태
    '//전역변수 초기화
    connIP = Empty
    connDB = Empty
    connUN = Empty
    connPW = Empty
    USER_ID = Empty
    USER_GB = Empty
    USER_DEPT = Empty
    MsgBox "로그아웃 되었습니다." & space(7), vbInformation, banner
    
End Sub

'------------------------------------
'  현재 파일을 닫아 리본 탭 닫음
'------------------------------------
Sub AddinUninstall()
   ThisWorkbook.Close False
End Sub

'---------------
'  환율조회기
'---------------
Sub FX_Calculator()
    f_currency_cal.Show vbModeless
End Sub
