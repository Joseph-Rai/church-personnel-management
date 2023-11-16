Attribute VB_Name = "sb_Show_Userform"
Option Explicit

Dim LISTDATA() As String '--//DB���� �޾ƿ� rs�� �迭�� ����
Dim LISTFIELD() As String '--//DB���� �޾ƿ� rs�� �ʵ带 �迭�� ����
Dim cntRecord As Integer '--//DB���� �޾ƿ� ���ڵ��� ����
Dim strSql As String '--//SQL ������

Sub frm_A3_Appointment_Form_Show()
    frm_Search_Appointment.Show vbModeless
End Sub

Sub frmPStaff_History_Show()
    frm_Search_History.Show vbModeless
End Sub
Sub frmPStaff_Detail_Show()
    frm_Search_PStaff_Detail.Show vbModeless
End Sub
Sub frmPStaff_Show()
    frm_Search_PStaff.Show vbModeless
End Sub
Sub frmAttendance_Show()
    frm_Search_Attendance.Show vbModeless
End Sub
Sub frmAttendance_Detail_Show()
    frm_Search_AttendanceDetail.Show vbModeless
End Sub
Sub frm_Update_Attendance_Show()
    frm_Update_Attendance.Show vbModeless
End Sub
Sub frm_Update_Appointment_Show()
    frm_Update_Appointment.Show vbModeless
End Sub
Sub frm_Update_PInformation_Show()
    frm_Update_PInformation.Show vbModeless
End Sub
Sub frm_Update_Theological_Show()
    frm_Update_Theological.Show vbModeless
End Sub
Sub frm_Update_History_Show()
    frm_Update_History.Show vbModeless
End Sub
Sub frm_Update_Church_Esta_Show()
    frm_Update_Church_Esta.Show vbModeless
End Sub
Sub frm_Update_Flight_Show()
    frm_Update_Flight.Show vbModeless
End Sub
Sub frm_Update_BCLeader_Show()
    frm_Update_BCLeader.Show vbModeless
End Sub
Sub frm_Search_Phone_Show()
    frm_Search_Phone.Show vbModeless
End Sub
Sub frm_Search_Statistic_Country_Show()
    SEARCH_CODE = 1 '--//������ ���
    frm_Search_Statistic.Show vbModeless
End Sub
Sub frm_Search_Statistic_Church_Show()
    SEARCH_CODE = 2 '--//��ȸ���
    frm_Search_Statistic.Show vbModeless
End Sub
Sub frm_Search_Statistic_PStaff_Show()
    SEARCH_CODE = 3 '--//��ȸ�����
    frm_Search_Statistic.Show vbModeless
End Sub
Sub frm_Search_Statistic_ChurchDetail_Show()
    SEARCH_CODE = 4 '--//��ȸ���
    frm_Search_Statistic.Show vbModeless
End Sub
Sub frm_Update_Union_Show()
    frm_Update_Union.Show vbModeless
End Sub
Sub frm_Update_Union_1_Show()
    frm_Update_Union_1.Show vbModeless
End Sub
Sub frm_Search_by_TitlePosition_Show()
    frm_Search_by_TitlePosition.Show vbModeless
End Sub
Sub frm_Update_Sermon_Show()
    frm_Update_Sermon.Show vbModeless
End Sub
Sub frm_Update_Visa_Show()
    frm_Update_Visa.Show vbModeless
End Sub
Sub frm_Update_User_Show()

    Call checkLoginStatus
    
    '--//����üũ(���� ���Ѹ� ���Ӱ���)
    Call GetUserAuthorities
    If IsInArray("USER_EDIT", LISTDATA) = -1 And IsInArray("DEPT_NUM_CHANGE", LISTDATA) = -1 Then
        MsgBox "������ �����ϴ�.", vbCritical, "���ѿ���"
        Exit Sub
    End If

    frm_Update_User.Show vbModeless
End Sub
Sub frm_Update_User_Authority_Show()
    
    Call checkLoginStatus
    
    '--//���ѿ� ���� ����
    Call GetUserAuthorities
    If IsInArray("USER_EDIT", LISTDATA) = -1 And IsInArray("SECTION_CHIEF", LISTDATA) = -1 Then
        MsgBox "������ �����ϴ�.", vbCritical, "���ѿ���"
        Exit Sub
    End If
    
    frm_Update_User_Authority.Show vbModeless
End Sub
Sub frm_Update_Counsel_Show()
    
    Call checkLoginStatus
    
    '--//����üũ(���� ���Ѹ� ���Ӱ���)
    Call GetUserAuthorities
    If IsInArray("COUNSEL", LISTDATA) = -1 Then
        MsgBox "��� ������ �����ϴ�. �����ڿ��� �����ϼ���.", vbCritical, "���ѿ���"
        Exit Sub
    End If
    
    frm_Update_Counsel.Show vbModeless
End Sub

Private Sub GetUserAuthorities()

    Dim sql As String
    
    sql = "SELECT b.authority FROM op_system.a_auth_table a" & _
          " LEFT JOIN op_system.a_authority b " & _
          "     ON a.authority_id = b.id" & _
          " WHERE a.user_id = " & USER_ID & ";"
    Call makeListData(sql, "op_system.a_auth_table")
    
End Sub


Private Sub makeListData(ByVal strSql As String, ByVal tableNM As String)

    Dim i As Integer, j As Integer
    
    '//DB���� �ڷ� ȣ���Ͽ� ���ڵ�¿� ��ȯ
    connectTaskDB
    callDBtoRS "makeListData", tableNM, strSql
    
    '//���ڵ���� �����͸� listData �迭�� ��ȯ
    If Not rs.EOF Then
        ReDim LISTDATA(0 To rs.RecordCount - 1, 0 To rs.Fields.Count - 1) '//DB���� ��ȯ�� �迭�� ũ�� ����: ���ڵ���� ���ڵ� ��, �ʵ� ��
        rs.MoveFirst
        For i = 0 To rs.RecordCount - 1
            For j = 0 To rs.Fields.Count - 1
                If IsNull(rs.Fields(j).Value) = True Then
                    LISTDATA(i, j) = ""
                Else
                    LISTDATA(i, j) = rs.Fields(j).Value
                End If
            Next j
            rs.MoveNext
        Next i
    End If
    
    '--//�ʵ�� �迭 ä���
    ReDim LISTFIELD(0 To rs.Fields.Count - 1)
    For i = 0 To rs.Fields.Count - 1
        LISTFIELD(i) = rs.Fields(i).Name
    Next i
    cntRecord = rs.RecordCount '--//���ڵ� �� ����
    disconnectALL
    
End Sub

Private Sub checkLoginStatus()

    If connIP = Empty Then setGlobalVariant '//�������� �缳��
    If checkLogin = 0 Then f_login.Show '//�α���üũ
    If checkLogin = 0 Then End '--//�α��� ���� �� ���ν��� ����

End Sub
