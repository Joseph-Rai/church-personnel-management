Attribute VB_Name = "sb_setGlobalVariant"
Option Explicit

'-----------------------------------------------------------------
'  �������� ����
'    - Error ��ȣ 3709, -2147217843
'    - �������� ���ν��� �� ����
'-----------------------------------------------------------------
Sub setGlobalVariant(Optional procedureNM As String = "NULL")
    Dim strSql As String
    
    '//�������� ��ȸ
    connectCommonDB
    strSql = "SELECT * FROM common.users WHERE user_id = (SELECT user_id FROM common.users WHERE user_nm = " & SText(Application.UserName) & ");"
    callDBtoRS "setGlobalVariant", "common.users", strSql, , "����������ȸ"
    
    '//�۾�DB������ ���� �������� ����
    connIP = rs("argIP").Value
    connDB = rs("argDB").Value
    connUN = rs("argUN").Value
    connPW = rs("argPW").Value
    
    '//�������� �������� ����
    USER_ID = rs("user_id").Value
    USER_NM = rs("user_nm").Value
    USER_GB = rs("user_gb").Value
    USER_DEPT = rs("user_dept").Value
    
    disconnectALL
    
    connectTaskDB
    strSql = "SELECT * FROM op_system.db_ovs_dept WHERE dept_id = " & SText(USER_DEPT)
    callDBtoRS "setGlobalVariant", "op_system.db_ovs_dept", strSql, , "����������ȸ"
    
    PICPATH_PSTAFF = rs("dept_picpath").Value
    Set WB_ORIGIN = ThisWorkbook
    
    disconnectALL
    
    '//�����߻����� �������� �� ���� �� ���� ���ν��� ����
    If procedureNM <> "NULL" Then Application.run procedureNM
End Sub

'--------------------------------------------------------------------
'  SA�� ������� ȯ�� �ľ��� ���� ������� �̸����� �α���
'--------------------------------------------------------------------
Sub setGlobalVariant2(userNM As String)
    Dim strSql As String
    
    '//Ư�� ����� �������� ��ȸ
    connectCommonDB
    strSql = "SELECT * FROM common.users WHERE user_nm = " & SText(userNM) & ";"
    callDBtoRS "setGlobalVariant2", "common.users", strSql, "Ư�����������������ȸ"
    
    '//�۾�DB������ ���� �������� ����
    connIP = rs("argIP").Value
    connDB = rs("argDB").Value
    connUN = rs("argUN").Value
    connPW = rs("argPW").Value
    
    '//�����ڵ� ����
    USER_ID = rs("user_id").Value
    USER_NM = rs("user_nm").Value
    USER_GB = rs("user_gb").Value
    USER_DEPT = rs("user_dept").Value
    
    disconnectALL
End Sub
