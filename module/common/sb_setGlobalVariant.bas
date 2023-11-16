Attribute VB_Name = "sb_setGlobalVariant"
Option Explicit

'-----------------------------------------------------------------
'  전역변수 설정
'    - Error 번호 3709, -2147217843
'    - 실행중인 프로시저 재 실행
'-----------------------------------------------------------------
Sub setGlobalVariant(Optional procedureNM As String = "NULL")
    Dim strSql As String
    
    '//전역변수 조회
    connectCommonDB
    strSql = "SELECT * FROM common.users WHERE user_id = (SELECT user_id FROM common.users WHERE user_nm = " & SText(Application.UserName) & ");"
    callDBtoRS "setGlobalVariant", "common.users", strSql, , "전역변수조회"
    
    '//작업DB연결을 위한 전역변수 설정
    connIP = rs("argIP").Value
    connDB = rs("argDB").Value
    connUN = rs("argUN").Value
    connPW = rs("argPW").Value
    
    '//유저관련 전역변수 설정
    USER_ID = rs("user_id").Value
    USER_NM = rs("user_nm").Value
    USER_GB = rs("user_gb").Value
    USER_DEPT = rs("user_dept").Value
    
    disconnectALL
    
    connectTaskDB
    strSql = "SELECT * FROM op_system.db_ovs_dept WHERE dept_id = " & SText(USER_DEPT)
    callDBtoRS "setGlobalVariant", "op_system.db_ovs_dept", strSql, , "전역변수조회"
    
    PICPATH_PSTAFF = rs("dept_picpath").Value
    Set WB_ORIGIN = ThisWorkbook
    
    disconnectALL
    
    '//오류발생으로 전역변수 재 설정 시 기존 프로시저 실행
    If procedureNM <> "NULL" Then Application.run procedureNM
End Sub

'--------------------------------------------------------------------
'  SA가 사용자의 환경 파악을 위해 사용자의 이름으로 로그인
'--------------------------------------------------------------------
Sub setGlobalVariant2(userNM As String)
    Dim strSql As String
    
    '//특정 사용자 전역변수 조회
    connectCommonDB
    strSql = "SELECT * FROM common.users WHERE user_nm = " & SText(userNM) & ";"
    callDBtoRS "setGlobalVariant2", "common.users", strSql, "특정사용자전역변수조회"
    
    '//작업DB연결을 위한 전역변수 설정
    connIP = rs("argIP").Value
    connDB = rs("argDB").Value
    connUN = rs("argUN").Value
    connPW = rs("argPW").Value
    
    '//유저코드 설정
    USER_ID = rs("user_id").Value
    USER_NM = rs("user_nm").Value
    USER_GB = rs("user_gb").Value
    USER_DEPT = rs("user_dept").Value
    
    disconnectALL
End Sub
