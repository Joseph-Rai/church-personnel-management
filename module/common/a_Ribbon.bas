Attribute VB_Name = "a_Ribbon"
Option Explicit

'-----------------------------------------------------
'  ���� �޴��� Button ID�� ���� ó�� ���ν���
'-----------------------------------------------------
Sub run_RibbonControl_WIS(Button As Office.IRibbonControl)
    Select Case Button.Id
        '//���α׷�
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
                      
        '//����
        Case "LogIn":     Call LogIn
        Case "LogOut":     Call LogOut
        Case "AddinUninstall":     Call AddinUninstall
        
        Case Else:     Call RibbonButton_Error(Button.Id)
    End Select
End Sub

'-------------------------------------------------------------------
'  Button ID�� ���� ó�� ���ν����� ���� ��� ���� �޽���
'-------------------------------------------------------------------
Sub RibbonButton_Error(sbID As String)
   MsgBox "�����Ͻ� �޴�(" & sbID & ")�� ���� �غ� �Ǿ� ���� �ʽ��ϴ�.", vbCritical, banner
End Sub

'-----------
'  �α���
'-----------
Sub LogIn()
    If checkLogin = 1 Then
        MsgBox Application.UserName & "�� �̹� �α��� �Ǿ� �ֽ��ϴ�.", vbInformation, banner
        Exit Sub
    End If
    f_login.Show
    
End Sub

'------------
'  �α׾ƿ�
'------------
Sub LogOut()
    
    '--//������ ������ ��Ʈ ����
    Sheets("������ ������").Visible = False
    Sheets("A3�λ�߷�").Visible = False
    
    If checkLogin = 0 Then
        MsgBox Application.UserName & "�� �̹� �α׾ƿ� �Ǿ� �ֽ��ϴ�.", vbInformation, banner
        Exit Sub
    End If
    checkLogin = 0 '�α׾ƿ� ����
    '//�������� �ʱ�ȭ
    connIP = Empty
    connDB = Empty
    connUN = Empty
    connPW = Empty
    USER_ID = Empty
    USER_GB = Empty
    USER_DEPT = Empty
    MsgBox "�α׾ƿ� �Ǿ����ϴ�." & space(7), vbInformation, banner
    
End Sub

'------------------------------------
'  ���� ������ �ݾ� ���� �� ����
'------------------------------------
Sub AddinUninstall()
   ThisWorkbook.Close False
End Sub

'---------------
'  ȯ����ȸ��
'---------------
Sub FX_Calculator()
    f_currency_cal.Show vbModeless
End Sub
