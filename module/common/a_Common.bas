Attribute VB_Name = "a_Common"
Option Explicit

Public Const programv As String = "V20201230" '프로그램 버전 관리★★
Public Const banner As String = "해외국 통합관리 프로그램 " & programv '★★
Public Const ODBCDriver As String = "MariaDB ODBC 3.1 Driver" 'Client PC에 설치된 ODBC Driver★★
Public Const IPAddress As String = "172.17.109.45" 'DB IP Address★★
Public Const commonDB As String = "common" 'Common DB명★★
Public Const commonID As String = "common" 'Common DB UN★★
Public Const commonPW As String = "12345" 'Common DB 비밀번호★★
Public Const globalSheetPW As String = "12345" '시트비밀번호
Public Const updateDay As Integer = 7 '매월 업데이트하는 날짜

Public Const SCHEMA_OPSYSTEM = "op_system." '스키마명: op_system
Public Const SCHEME_COMMON = "common." '스키마명: common

Public conn As ADODB.Connection 'ADO Connection 개체 변수
Public rs As New ADODB.RecordSet 'ADO Recordset 개체 변수
Public connIP As String, connDB As String, connUN As String, connPW As String 'Task DB 연결 정보
Public USER_ID As Integer '사용자코드
Public USER_GB As String '사용자구분(SA, AM, MG, WP)
Public USER_NM As String '사용자이름
Public USER_DEPT As String '사용자부서
Public checkLogin As Integer '로그인 여부 0: 로그인 안함, 1 = 로그인
Public today As Date '콤보박스의 날짜 옵션이 비었을 경우 오늘 날짜로 조회
Public TASK_CODE As Integer '--//1. 교회발령, 2. 직분임명, 3. 직책임명
Public SEARCH_CODE As Integer '--//1. 국가별 통계, 2. 교회통계, 3. 목회자통계
Public argShow As Integer '--//1. frm_Update_Appointment, 2. frm_Update_PInformation, 3. frm_Search_Appointment
                          '--//1. BCLeader, 2. FamilyInfo
Public argShow2 As Integer '--//1. 선지자 가족 업데이트, 2. 배우자 가족 업데이트
Public argShow3 As Integer '--//1. frm_Update_Appointment, 2. frm_Update_FamilyInfo, 3. frm_Search_Appointment
Public PICPATH_PSTAFF As String '--//부서별 사진파일 경로
Public WB_ORIGIN As Workbook '--//로그인 시 현재문서 받아오기

Public Const COUNT_PAGE_WIDTH_CELLS = 15 '--//교회별 선지자현황 페이지너비
Public Const COUNT_PAGE_HEIGHT_CELLS = 37 '--//교회별 선지자현황 페이지높이


'-------------------------------
'  Common DB연결
'    - 로그인 체크
'    - TaskDB접속 정보 반환
'-------------------------------
Sub connectCommonDB()
    connectDB IPAddress, commonDB, commonID, commonPW
End Sub

'-------------------
'  Task DB연결
'-------------------
Sub connectTaskDB()
    connectDB connIP, connDB, connUN, connPW
End Sub

'-----------------------------------------------
'  DB연결 프로시저
'    - connectDB(서버 IP, 스키마, ID, PW)
'-----------------------------------------------
Sub connectDB(argIP As String, argDB As String, argId As String, argPW As String)
    
    If argIP = "" Then
        MsgBox "로그인 되어 있지 않습니다.", vbCritical + vbOKOnly, banner
        End
    End If
    
    Set conn = New ADODB.Connection
    conn.ConnectionString = "Driver={" & ODBCDriver & "};Server=" & argIP & ";Port=3306;Database=" & argDB & ";User=" & argId & ";Password=" & argPW & ";Option=2;"
    conn.Open
End Sub

'---------------------------------------------------------------------
'  레코드셋 설정 및 데이터 반환
'    - calDBtoRS(프로시저명, 테이블명, SQL문, 폼이름, 잡이름)
'    - 오류발생 시 에러 핸들링 및 로그 기록
'    - 오류발생 안하면 잡 수행 프로시저에서 로그 기록(필요 시)
'---------------------------------------------------------------------
Sub callDBtoRS(procedureNM As String, tableNM As String, SQLScript As String, Optional formNM As String = "NULL", Optional jobNM As String = "데이터 조회")
On Error GoTo ErrHandler
    Set rs = New ADODB.RecordSet
    rs.CursorLocation = adUseClient
    rs.Open Source:=SQLScript, ActiveConnection:=conn, CursorType:=adOpenForwardOnly, LockType:=adLockReadOnly, Options:=adCmdText
    Exit Sub
ErrHandler:
    ErrHandle procedureNM, tableNM, SQLScript, formNM, jobNM
    writeLog procedureNM, tableNM, Replace(SQLScript, ";", ""), 1, formNM, jobNM '//오류코드 1
End Sub

'-------------------------------------------------------------------------------------
'  SQL문을 실행하고 로그기록
'    - executeSqlWithLog(SQL문, 프로시져명, 테이블명, 잡이름(옵션))
'    - 작업성공 시 로그기록
'    - 오류발생 시 executeSQL 실행 시 로그기록
'-------------------------------------------------------------------------------------
Public Sub executeSqlWithLog(sql As String, procedureNM As String, tableNM As String, Optional jobNM As String = "NULL")
    Dim result As T_RESULT
    
    '--//DB에 연결
    connectTaskDB
    
    result.strSql = sql
    result.affectedCount = executeSQL(procedureNM, tableNM, sql, , jobNM)
    writeLog procedureNM, tableNM, Replace(sql, ";", ""), 0, , jobNM, result.affectedCount
    
    '--//DB연결 해제
    disconnectALL
End Sub

'-------------------------------------------------------------------------------------
'  SQL문을 실행하고 실행결과 영향을 받은 레코드 수를 반환
'    - executeSQL(프로시져명, 테이블명, SQL문, 폼이름(옵션), 잡이름(옵션))
'    - SQL문 실행 결과 성공 여부를 알기 위해 영향 받은 레코드 수 검토
'    - 오류발생 시 에러 핸들링 및 로그 기록
'    - 오류발생 안하면 잡 수행 프로시저에서 로그 기록
'-------------------------------------------------------------------------------------
Public Function executeSQL(procedureNM As String, tableNM As String, SQLScript As String, Optional formNM As String = "NULL", Optional jobNM As String = "기타") As Long
On Error GoTo ErrHandler
    Dim affectedCount As Long
    Dim sql As Variant
    
    For Each sql In Split(SQLScript, ";")
        If Len(sql) > 0 Then
            Debug.Print sql
            conn.Execute CommandText:=sql, recordsaffected:=affectedCount
        End If
    Next
    
    executeSQL = affectedCount
    Exit Function
ErrHandler:
    ErrHandle procedureNM, tableNM, SQLScript, formNM, jobNM
    writeLog procedureNM, tableNM, Replace(SQLScript, ";", ""), 1, formNM, jobNM '//오류코드 1
End Function

'--------------------------
'  DB 및 RS 연결 해제
'--------------------------
Sub disconnectRS()
    On Error Resume Next
        rs.Close
        Set rs = Nothing
    On Error GoTo 0
End Sub
Sub disconnectDB()
    On Error Resume Next
        conn.Close
        Set conn = Nothing
    On Error GoTo 0
End Sub
Sub disconnectALL()
    On Error Resume Next
        rs.Close
        Set rs = Nothing
        conn.Close
        Set conn = Nothing
    On Error GoTo 0
End Sub

'------------------------------------------------
'  SQL 패턴매칭 검색어 처리('%검색어%')
'------------------------------------------------
Public Function PText(argString As Variant) As String
    If argString = "" Or Len(argString) = 0 Then
        PText = "'%%'"
    Else
        PText = "'%" & Trim(Replace(Replace(argString, "%", "\%"), "'", "''")) & "%'"
    End If
End Function

'---------------------------------------------
'  SQL 스칼라매칭 검색어 처리('검색어')
'---------------------------------------------
Public Function SText(argString As Variant) As String
    If argString = "" Or Len(argString) = 0 Then
        SText = "''"
    Else
        SText = "'" & Trim(Replace(Replace(argString, "%", "\%"), "'", "''")) & "'"
    End If
End Function

'--------------------
'  매크로 최적화
'--------------------
Sub Optimization()
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
'        .Calculation = xlCalculationManual
    End With
End Sub

'-------------------------
'  매크로 최적화 원복
'-------------------------
Sub Normal()
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
'        .Calculation = xlCalculationAutomatic
    End With
End Sub

'----------------
'  전체화면On
'----------------
Sub FullscreenOn()
    ActiveWindow.DisplayHeadings = False
    ActiveWindow.DisplayWorkbookTabs = False
    Application.DisplayFullScreen = True
    Application.DisplayFormulaBar = False
End Sub

'----------------
'  전체화면Off
'----------------
Sub FullscreenOff()
    ActiveWindow.DisplayHeadings = True
    ActiveWindow.DisplayWorkbookTabs = True
    Application.DisplayFullScreen = False
    Application.DisplayFormulaBar = True
End Sub

'---------------------
'  엑셀화면 숨기기
'---------------------
Sub HideExcel()
    Application.Visible = False
End Sub

'---------------------
'  엑셀화면 보이기
'---------------------
Sub ShowExcel()
    Application.Visible = True
End Sub
Public Function GetDesktopPath(Optional BackSlash As Boolean = True)
    Dim oWSHShell As Object
    
    Set oWSHShell = CreateObject("WScript.Shell")
    If BackSlash = True Then
        GetDesktopPath = oWSHShell.SpecialFolders("Desktop") & "\"
    Else
        GetDesktopPath = oWSHShell.SpecialFolders("Desktop")
    End If
    
    Set oWSHShell = Nothing
End Function
