VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Search_Statistic 
   Caption         =   "국가별 통계데이터 조회"
   ClientHeight    =   1470
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3210
   OleObjectBlob   =   "frm_Search_Statistic.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "frm_Search_Statistic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cntlst1Col As Integer '//리스트 컬럼 수
Dim TB1 As String, TB2 As String, TB3 As String, TB4 As String '//폼에 연결된 DB 테이블
Const strOrderBy As String = "currency_un ASC" '//DB에서 sort_order 필드
Dim caseSave As Integer '//1: 추가, 2: 수정, 3: 삭제(SUSPEND), 4: 완전삭제
Dim queryKey As Integer '//리스트 위치 반환에 사용될 id
Dim LISTDATA() As String '--//DB에서 받아온 rs를 배열로 저장
Dim LISTFIELD() As String '--//DB에서 받아온 rs의 필드를 배열로 저장
Dim cntRecord As Integer '--//DB에서 받아온 레코드의 개수
Dim strSql As String '--//SQL 쿼리문
Dim txtBox_Focus As MSForms.control
Dim rngA As Range, rngB As Range, rngC As Range
Dim ws As Worksheet

Private Sub cboYear_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Un'HookListBoxScroll
End Sub

Private Sub cboYear_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Me.cboYear.ListCount > 0 Then
        'HookListBoxScroll Me, Me.cboYear
    End If
End Sub
Private Sub cboMonth_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Un'HookListBoxScroll
End Sub

Private Sub cboMonth_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Me.cboMonth.ListCount > 0 Then
        'HookListBoxScroll Me, Me.cboMonth
    End If
End Sub

Private Sub UserForm_Initialize()
    
    Dim i As Long
    
    If connIP = Empty Then setGlobalVariant '//전역변수 재설정
    If checkLogin = 0 Then f_login.Show '//로그인체크
    If checkLogin = 0 Then End '--//로그인 실패 시 프로시저 종료
    
    '--//시트설정
    Select Case SEARCH_CODE
    Case 1
        Set ws = WB_ORIGIN.Sheets("국가별 통계") '--//국가별 통계 시트
    Case 2
        Set ws = WB_ORIGIN.Sheets("교회통계") '--//교회통계 시트
    Case 3
        Set ws = WB_ORIGIN.Sheets("목회자통계") '--//목회자통계 시트
    Case 4
        Set ws = WB_ORIGIN.Sheets("교회통계상세") '--//교회통계상세 시트
    Case Else
    End Select
    
    '--//기초설정
    Me.cmdClose.Cancel = True
    TB1 = "op_system.temp_statistic_by_country" '--//국가별통계
    TB2 = "op_system.temp_statistic_by_church" '--//교회별통계
    TB3 = "op_system.temp_statistic_by_pstaff" '--//교회별 목회자통계
    TB4 = "op_system.temp_statistic_by_church_all" '--//교회별통계상세
    
    '--//콤보박스 월 추가
    For i = 1 To 12
        With Me.cboMonth
            .AddItem i
        End With
    Next
    
    '--//콤보박스 년도 추가
    For i = year(Date) To year(Date) - 10 Step -1
        With Me.cboYear
            .AddItem i
        End With
    Next
    
    '--//콤보박스 최신데이터 날짜삽입
    Me.cboYear = IIf(Day(Date) < 10, year(DateAdd("m", -2, Date)), year(DateAdd("m", -1, Date)))
    Me.cboMonth = IIf(Day(Date) < 10, month(DateAdd("m", -2, Date)), month(DateAdd("m", -1, Date)))
    
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer, j As Integer
    Dim result As T_RESULT
    
    '--//시트 활성화 및 잠금해제
    WB_ORIGIN.Activate
    ws.Activate
    Call shUnprotect(globalSheetPW)
    
    Select Case SEARCH_CODE
    Case 1
        WB_ORIGIN.Activate
        Sheets("국가별 통계").Activate
        Optimization
        '--//기존 데이터 삭제
        Call initializeRepport
        
        '--//temp_churchlist_by_time 테이블 업데이트
        strSql = "CALL `Routine_churchlist_by_time`(" & SText(DateSerial(Me.cboYear, Me.cboMonth, 1)) & "," & SText(USER_DEPT) & ");"
        connectTaskDB
        result.strSql = strSql
        result.affectedCount = executeSQL("cmdADD_Clikc", "temp_pstaff_by_time", strSql, Me.Name, "임시조회 테이블 생성")
    '    writeLog "cmdADD_Click", "temp_pstaff_by_time", strSQL, 0, Me.Name, "임시조회 테이블 생성", result.affectedCount
        disconnectALL
        
        '--//temp_pstaff_by_time 테이블 업데이트
        strSql = "CALL `Routine_pstaff_by_time`(" & SText(DateSerial(Me.cboYear, Me.cboMonth, 1)) & "," & SText(USER_DEPT) & ");"
        connectTaskDB
        result.strSql = strSql
        result.affectedCount = executeSQL("cmdADD_Clikc", "temp_pstaff_by_time", strSql, Me.Name, "임시조회 테이블 생성")
    '    writeLog "cmdADD_Click", "temp_pstaff_by_time", strSQL, 0, Me.Name, "임시조회 테이블 생성", result.affectedCount
        disconnectALL
        
        '--//temp_statistic_by_country 테이블 업데이트
        strSql = "CALL `Routine_statistic_by_Country`(" & SText(DateSerial(Me.cboYear, Me.cboMonth, 1)) & ", " & SText(USER_DEPT) & ");"
        connectTaskDB
        result.strSql = strSql
        result.affectedCount = executeSQL("cmdADD_Clikc", "temp_statistic_by_country", strSql, Me.Name, "임시조회 테이블 생성")
    '    writeLog "cmdADD_Click", "temp_pstaff_by_time", strSQL, 0, Me.Name, "임시조회 테이블 생성", result.affectedCount
        disconnectALL
        
        '--//SQL문
        strSql = makeSelectSQL(TB1)
        
        '--//DB에서 자료 호출하여 레코드셋 반환
        Call makeListData(strSql, TB1)
        
        '--//리포트 포맷설정
        Call makeReport
        
        '--//반환된 ListData를 보고서 시트에 삽입
        If cntRecord > 0 Then
            Range("Stat_Country_Start").Resize(, UBound(LISTFIELD) + 1) = LISTFIELD
            Range("Stat_Country_Start").Offset(1).Resize(cntRecord, UBound(LISTFIELD) + 1) = LISTDATA
        End If
        Range("A2").Copy
        Range(Cells(4, "C"), Cells(4 + cntRecord - 1, "AG")).PasteSpecial Paste:=xlPasteValues, operation:=xlPasteSpecialOperationAdd
        Application.CutCopyMode = False
        
        '--//변수 초기화
        Call sbClearVariant
        
        '--//조회기준일 등록
        Range("Stat_Country_Date") = DateSerial(Me.cboYear, Me.cboMonth, 1)
        
        strSql = "SELECT a.dept_nm FROM op_system.db_ovs_dept a WHERE a.dept_id = " & SText(USER_DEPT)
        Call makeListData(strSql, "op_system.db_ovs_dept")
        
        Sheets("국가별 통계").Range("A1").FormulaR1C1 = _
                            "=""" & LISTDATA(0, 0) & " 국가별 출석현황 및 목회자 통계표 [""&TEXT(Stat_Country_Date,""yyyy년 mm월"")&""]"""
        Call sbClearVariant
        
        Range("A2").Select
        Normal
    Case 2
        WB_ORIGIN.Activate
        Sheets("교회통계").Activate
        Optimization
        '--//기존 데이터 삭제
        Call initializeRepport
        
        '--//temp_churchlist_by_time 테이블 업데이트
        strSql = "CALL `Routine_churchlist_by_time`(" & SText(DateSerial(Me.cboYear, Me.cboMonth, 1)) & "," & SText(USER_DEPT) & ");"
        connectTaskDB
        result.strSql = strSql
        result.affectedCount = executeSQL("cmdADD_Clikc", "temp_pstaff_by_time", strSql, Me.Name, "임시조회 테이블 생성")
    '    writeLog "cmdADD_Click", "temp_pstaff_by_time", strSQL, 0, Me.Name, "임시조회 테이블 생성", result.affectedCount
        disconnectALL
        
        '--//temp_pstaff_by_time 테이블 업데이트
        strSql = "CALL `Routine_pstaff_by_time`(" & SText(Format(WorksheetFunction.EoMonth(DateSerial(Me.cboYear, Me.cboMonth, 1), 0), "yyyy-mm-dd")) & ", " & SText(USER_DEPT) & ");"
        connectTaskDB
        result.strSql = strSql
        result.affectedCount = executeSQL("cmdADD_Clikc", "temp_pstaff_by_time", strSql, Me.Name, "임시조회 테이블 생성")
    '    writeLog "cmdADD_Click", "temp_pstaff_by_time", strSQL, 0, Me.Name, "임시조회 테이블 생성", result.affectedCount
        disconnectALL
        
        '--//temp_statistic_by_church 테이블 업데이트
        strSql = "CALL `Routine_statistic_by_Church`(" & SText(DateSerial(Me.cboYear, Me.cboMonth, 1)) & ");"
        connectTaskDB
        result.strSql = strSql
        result.affectedCount = executeSQL("cmdADD_Clikc", "temp_statistic_by_country", strSql, Me.Name, "임시조회 테이블 생성")
    '    writeLog "cmdADD_Click", "temp_pstaff_by_time", strSQL, 0, Me.Name, "임시조회 테이블 생성", result.affectedCount
        disconnectALL
        
        '--//SQL문
        strSql = makeSelectSQL(TB2)
        
        '--//DB에서 자료 호출하여 레코드셋 반환
        Call makeListData(strSql, TB2)
        
        '--//리포트 포맷설정
        Call makeReport
        
        '--//반환된 ListData를 보고서 시트에 삽입
        If cntRecord > 0 Then
            Range("Stat_Church_Start").Resize(, UBound(LISTFIELD) + 1) = LISTFIELD
            Range("Stat_Church_Start").Offset(1).Resize(cntRecord, UBound(LISTFIELD) + 1) = LISTDATA
        End If
        Range("A2").Copy
        Range(Cells(4, "F"), Cells(4 + cntRecord - 1, "AF")).PasteSpecial Paste:=xlPasteValues, operation:=xlPasteSpecialOperationAdd
        Application.CutCopyMode = False
        
        '--//연합회 순으로 정렬
'        Select Case USER_DEPT
'        Case 10
'            ActiveWorkbook.Worksheets("교회통계").Sort.SortFields.Clear
'            ActiveWorkbook.Worksheets("교회통계").Sort.SortFields.Add Key:=Range("Stat_Church_Start").Offset(1, 1).Resize(cntRecord), _
'                SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder:= _
'                "중동,카트만두,네팔동부,네팔중부,네팔서부,북인도,인도차이나", DataOption:=xlSortNormal
'            With ActiveWorkbook.Worksheets("교회통계").Sort
'                .SetRange Range("Stat_Church_Start").Offset(1, -1).Resize(cntRecord, UBound(listField) + 2)
'                .Header = xlGuess
'                .MatchCase = False
'                .Orientation = xlTopToBottom
'                .SortMethod = xlPinYin
'                .Apply
'            End With
'        Case Else
'        End Select
        
        '--//변수 초기화
        Call sbClearVariant
        
        '--//조회기준일 등록
        Range("Stat_Church_Date") = DateSerial(Me.cboYear, Me.cboMonth, 1)
        
        strSql = "SELECT a.dept_nm FROM op_system.db_ovs_dept a WHERE a.dept_id = " & SText(USER_DEPT)
        Call makeListData(strSql, "op_system.db_ovs_dept")
        
        Sheets("교회통계").Range("A1").FormulaR1C1 = _
                        "=""" & LISTDATA(0, 0) & " 교회별 출석현황 및 목회자 통계표  [""&TEXT(Stat_Church_Date,""yyyy년 mm월"")&"" 기준]"""
        Call sbClearVariant
        
        Range("A2").Select
        Normal
    Case 3
        WB_ORIGIN.Activate
        Sheets("목회자통계").Activate
        Optimization
        '--//기존 데이터 삭제
        Call initializeRepport
        
        '--//temp_churchlist_by_time 테이블 업데이트
        strSql = "CALL `Routine_churchlist_by_time`(" & SText(DateSerial(Me.cboYear, Me.cboMonth, 1)) & "," & SText(USER_DEPT) & ");"
        connectTaskDB
        result.strSql = strSql
        result.affectedCount = executeSQL("cmdADD_Clikc", "temp_pstaff_by_time", strSql, Me.Name, "임시조회 테이블 생성")
    '    writeLog "cmdADD_Click", "temp_pstaff_by_time", strSQL, 0, Me.Name, "임시조회 테이블 생성", result.affectedCount
        disconnectALL
        
        '--//temp_pstaff_by_time 테이블 업데이트
        strSql = "CALL `Routine_pstaff_by_time`(" & SText(Format(DateSerial(Me.cboYear, Me.cboMonth, 1), "yyyy-mm-dd")) & ", " & SText(USER_DEPT) & ");"
        connectTaskDB
        result.strSql = strSql
        result.affectedCount = executeSQL("cmdADD_Clikc", "temp_pstaff_by_time", strSql, Me.Name, "임시조회 테이블 생성")
    '    writeLog "cmdADD_Click", "temp_pstaff_by_time", strSQL, 0, Me.Name, "임시조회 테이블 생성", result.affectedCount
        disconnectALL
        
        '--//temp_statistic_by_pstaff 테이블 업데이트
        strSql = "CALL `Routine_statistic_by_pstaff`(" & SText(DateSerial(Me.cboYear, Me.cboMonth, 1)) & ");"
        connectTaskDB
        result.strSql = strSql
        result.affectedCount = executeSQL("cmdADD_Clikc", "temp_statistic_by_country", strSql, Me.Name, "임시조회 테이블 생성")
    '    writeLog "cmdADD_Click", "temp_pstaff_by_time", strSQL, 0, Me.Name, "임시조회 테이블 생성", result.affectedCount
        disconnectALL
        
        '--//SQL문
        strSql = makeSelectSQL(TB3)
        
        '--//DB에서 자료 호출하여 레코드셋 반환
        Call makeListData(strSql, TB3)
        
        '--//리포트 포맷설정
        Call makeReport
        
        '--//반환된 ListData를 보고서 시트에 삽입
        If cntRecord > 0 Then
            Range("Stat_PStaff_Start").Resize(, UBound(LISTFIELD) + 1) = LISTFIELD
            Range("Stat_PStaff_Start").Offset(1).Resize(cntRecord, UBound(LISTFIELD) + 1) = LISTDATA
        End If
        Range("A2").Copy
        Range(Cells(5, "F"), Cells(5 + cntRecord - 1, "S")).PasteSpecial Paste:=xlPasteValues, operation:=xlPasteSpecialOperationAdd
        Application.CutCopyMode = False
        
        '--//연합회 순으로 정렬
'        Select Case USER_DEPT
'        Case 10
'            ActiveWorkbook.Worksheets("목회자통계").Sort.SortFields.Clear
'            ActiveWorkbook.Worksheets("목회자통계").Sort.SortFields.Add Key:=Range("Stat_PStaff_Start").Offset(1, 1).Resize(cntRecord), _
'                SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder:= _
'                "중동,카트만두,네팔동부,네팔중부,네팔서부,북인도,인도차이나", DataOption:=xlSortNormal
'            With ActiveWorkbook.Worksheets("목회자통계").Sort
'                .SetRange Range("Stat_PStaff_Start").Offset(1, -1).Resize(cntRecord, UBound(listField) + 2)
'                .Header = xlGuess
'                .MatchCase = False
'                .Orientation = xlTopToBottom
'                .SortMethod = xlPinYin
'                .Apply
'            End With
'        Case Else
'        End Select
        
        '--//변수 초기화
        Call sbClearVariant
        
        '--//조회기준일 등록
        Range("Stat_PStaff_Date") = DateSerial(Me.cboYear, Me.cboMonth, 1)
        
        strSql = "SELECT a.dept_nm FROM op_system.db_ovs_dept a WHERE a.dept_id = " & SText(USER_DEPT)
        Call makeListData(strSql, "op_system.db_ovs_dept")
        
        Sheets("목회자통계").Range("A1").FormulaR1C1 = _
                        "=""" & LISTDATA(0, 0) & " 교회별 목회자 통계표 [""&TEXT(Stat_PStaff_Date,""yyyy년 mm월"")&"" 기준]"""
        Call sbClearVariant
        
        Range("A2").Select
        Normal
    Case 4
        WB_ORIGIN.Activate
        Sheets("교회통계상세").Activate
        Optimization
        '--//기존 데이터 삭제
        Call initializeRepport
        
        '--//temp_churchlist_by_time 테이블 업데이트
        strSql = "CALL `Routine_churchlist_by_time`(" & SText(Format(Application.EoMonth(DateSerial(Me.cboYear, Me.cboMonth, 1), 0), "yyyy-mm-dd")) & "," & SText(USER_DEPT) & ");"
        connectTaskDB
        result.strSql = strSql
        result.affectedCount = executeSQL("cmdADD_Clikc", "temp_pstaff_by_time", strSql, Me.Name, "임시조회 테이블 생성")
    '    writeLog "cmdADD_Click", "temp_pstaff_by_time", strSQL, 0, Me.Name, "임시조회 테이블 생성", result.affectedCount
        disconnectALL
        
        '--//temp_statistic_by_church 테이블 업데이트
        strSql = "CALL `Routine_statistic_by_Church_all`(" & SText(DateSerial(Me.cboYear, Me.cboMonth, 1)) & ");"
        connectTaskDB
        result.strSql = strSql
        result.affectedCount = executeSQL("cmdADD_Clikc", "temp_statistic_by_country", strSql, Me.Name, "임시조회 테이블 생성")
    '    writeLog "cmdADD_Click", "temp_pstaff_by_time", strSQL, 0, Me.Name, "임시조회 테이블 생성", result.affectedCount
        disconnectALL
        
        '--//SQL문
        strSql = makeSelectSQL(TB4)
        
        '--//DB에서 자료 호출하여 레코드셋 반환
        Call makeListData(strSql, TB4)
        
        '--//리포트 포맷설정
        Call makeReport
        
        '--//반환된 ListData를 보고서 시트에 삽입
        If cntRecord > 0 Then
            Range("Stat_ChurchDetail_Start").Resize(, UBound(LISTFIELD) + 1) = LISTFIELD
            Range("Stat_ChurchDetail_Start").Offset(1).Resize(cntRecord, UBound(LISTFIELD) + 1) = LISTDATA
        End If
        Range("A2").Copy
        Range(Cells(4, "D"), Cells(4 + cntRecord - 1, "T")).PasteSpecial Paste:=xlPasteValues, operation:=xlPasteSpecialOperationAdd
        Application.CutCopyMode = False
        
        '--//총회관리지교회 MM 관리자 삽입
        Dim indexR As Integer
        For indexR = 3 To Cells(Rows.Count, "B").End(xlUp).Row
            If Cells(indexR, "C") = "MM" And Cells(indexR - 1, "C") = "HBC" Then
                Cells(indexR, "E") = Cells(indexR - 1, "E")
                Cells(indexR, "F") = Cells(indexR - 1, "F")
            End If
        Next

        
        '--//변수 초기화
        Call sbClearVariant
        
        '--//조회기준일 등록
        Range("Stat_ChurchDetail_Date") = DateSerial(Me.cboYear, Me.cboMonth, 1)
        
        strSql = "SELECT a.dept_nm FROM op_system.db_ovs_dept a WHERE a.dept_id = " & SText(USER_DEPT)
        Call makeListData(strSql, "op_system.db_ovs_dept")
        
        Sheets("교회통계상세").Range("A1").FormulaR1C1 = _
                        "=""" & LISTDATA(0, 0) & " 교회별 출석현황상세 [""&TEXT(Stat_ChurchDetail_Date,""yyyy년 mm월"")&"" 기준]"""
        Call sbClearVariant
        
        Range("A2").Select
        Normal
    
    Case Else
    End Select
    
    '--//시트보호
    Call shProtect(globalSheetPW)
    
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub makeListData(ByVal strSql As String, ByVal tableNM As String)

    Dim i As Integer, j As Integer
    
    '//DB에서 자료 호출하여 레코드셋에 반환
    connectTaskDB
    callDBtoRS "makeListData", tableNM, strSql, Me.Name
    
    '//레코드셋의 데이터를 listData 배열에 반환
    If Not rs.EOF Then
        ReDim LISTDATA(0 To rs.RecordCount - 1, 0 To rs.Fields.Count - 1) '//DB에서 반환할 배열의 크기 지정: 레코드셋의 레코드 수, 필드 수
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
    
    '--//필드명 배열 채우기
    ReDim LISTFIELD(0 To rs.Fields.Count - 1)
    For i = 0 To rs.Fields.Count - 1
        LISTFIELD(i) = rs.Fields(i).Name
    Next i
    
    cntRecord = rs.RecordCount
    
    disconnectALL
    
End Sub
'-----------------------------------------
'  조건별 Select SQL문 작성
'    - makeSelectSQL(검색어, 필터)
'    - DB에서 반환할 리스트 필드수정
'-----------------------------------------
Private Function makeSelectSQL(ByVal tableNM As String) As String
    
    Dim strOrderByClause As String
    Dim strTemp As Variant
    
    Select Case tableNM
    Case TB1
        strSql = "SELECT * FROM " & TB1 & " a WHERE a.`관리부서` = " & SText(USER_DEPT) & ";"
    Case TB2
        '--//해당 부서의 연합회 목록 추출
        strSql = "SELECT a.union_nm FROM " & "op_system.a_union" & " a WHERE a.ovs_dept = " & SText(USER_DEPT) & " ORDER BY a.sort_order;"
        Call makeListData(strSql, "op_system.a_union")
        
        If cntRecord > 0 Then '--//등록된 연합회가 있으면
            '--//ORDER BY 구문생성
            For Each strTemp In LISTDATA
                If strOrderByClause = "" Then
                    strOrderByClause = SText(strTemp)
                Else
                    strOrderByClause = strOrderByClause & "," & SText(strTemp)
                End If
            Next
            strOrderByClause = "FIELD(`연합회`," & strOrderByClause & ")"
            
            '--//최종 strSQL문 생성
            strSql = "SELECT * FROM " & TB2 & " a WHERE a.`관리부서` = " & SText(USER_DEPT) & " ORDER BY " & strOrderByClause & ",`정렬순서`;"
        Else
            '--//최종 strSQL문 생성
            strSql = "SELECT * FROM " & TB2 & " a WHERE a.`관리부서` = " & SText(USER_DEPT) & " ORDER BY `정렬순서`;"
        End If
    Case TB3
        '--//해당 부서의 연합회 목록 추출
        strSql = "SELECT a.union_nm FROM " & "op_system.a_union" & " a WHERE a.ovs_dept = " & SText(USER_DEPT) & " ORDER BY a.sort_order;"
        Call makeListData(strSql, "op_system.a_union")
        
        If cntRecord > 0 Then '--//등록된 연합회가 있으면
            '--//ORDER BY 구문생성
            For Each strTemp In LISTDATA
                If strOrderByClause = "" Then
                    strOrderByClause = SText(strTemp)
                Else
                    strOrderByClause = strOrderByClause & "," & SText(strTemp)
                End If
            Next
            strOrderByClause = "FIELD(`연합회`," & strOrderByClause & ")"
            
            '--//최종 strSQL문 생성
            strSql = "SELECT * FROM " & TB3 & " a WHERE a.`관리부서` = " & SText(USER_DEPT) & " ORDER BY " & strOrderByClause & ",`정렬순서`;"
        Else
            '--//최종 strSQL문 생성
            strSql = "SELECT * FROM " & TB3 & " a WHERE a.`관리부서` = " & SText(USER_DEPT) & " ORDER BY `정렬순서`;"
        End If
    Case TB4
        '--//최종 strSQL문 생성
        strSql = "SELECT '부서전체',NULL '교회형태',NULL '설립일',NULL '직책',NULL '관리자',SUM(a.`전체1회`) '전체1회',SUM(a.`전체4회`) '전체4회',SUM(a.`학생1회`) '학생1회',SUM(a.`학생4회`) '학생4회',SUM(a.`학생반차`) '학생반차',SUM(a.`전체침례`) '전체침례',SUM(a.`전도인`) '전도인',SUM(a.`지역장`) '지역장',SUM(a.`구역장`) '구역장',NULL '관리부서',NULL '연합회',NULL '본교회명',NULL '정렬순서','교회개수' FROM (SELECT * FROM " & TB4 & " a WHERE a.`관리부서` = " & SText(USER_DEPT) & " AND a.`교회형태` IN ('MC','HBC') ORDER BY `정렬순서`) a" & _
                " UNION SELECT d.* FROM (SELECT b.`연합회`,NULL '교회형태',NULL '설립일',NULL '직책',NULL '관리자',SUM(b.`전체1회`),SUM(b.`전체4회`),SUM(b.`학생1회`),SUM(b.`학생4회`),SUM(b.`학생반차`),SUM(b.`전체침례`),SUM(b.`전도인`),SUM(b.`지역장`),SUM(b.`구역장`),NULL '관리부서',NULL '연합회명',NULL '본교회명',NULL '정렬순서','교회개수' FROM (SELECT * FROM " & TB4 & " a WHERE a.`관리부서` = " & SText(USER_DEPT) & " AND a.`교회형태` IN ('MC','HBC') ORDER BY `정렬순서`) b LEFT JOIN op_system.a_union union_order ON union_order.union_nm=b.`연합회` GROUP BY `연합회` ORDER BY union_order.sort_order) d" & _
                " UNION SELECT c.* FROM (SELECT * FROM " & TB4 & " a WHERE a.`관리부서` = " & SText(USER_DEPT) & " ORDER BY `정렬순서`) c;"
    Case Else
    End Select
    makeSelectSQL = strSql
End Function

Private Sub sbClearVariant()
    Erase LISTFIELD
    Erase LISTDATA
    cntRecord = Empty
    strSql = vbNullString
End Sub
Private Function fnData_Validation()
'---------------------------------------
'유저폼 입력값에 대한 데이터 유효성 검사
'TRUE: 이상없음, FALSE: 잘못됨.
'---------------------------------------
    fnData_Validation = True '데이터가 유효하다는 가정 하에 시작
    
    
    If Not IsNumeric(Me.cboYear.Value) Then
        fnData_Validation = False
        MsgBox "년도 입력 값이 유효하지 않습니다. 다시 확인 해주세요.", vbCritical, banner
        Set txtBox_Focus = Me.cboYear:  Exit Function
    End If
    If Me.cboYear < 1900 Or Me.cboYear > 2100 Then
        fnData_Validation = False
        MsgBox "년도 입력 값이 유효하지 않습니다. 다시 확인 해주세요.", vbCritical, banner
        Set txtBox_Focus = Me.cboYear: Exit Function
    End If
    If Not IsNumeric(Me.cboMonth.Value) Then
        fnData_Validation = False
        MsgBox "월 입력 값이 유효하지 않습니다. 다시 확인 해주세요.", vbCritical, banner
        Set txtBox_Focus = Me.cboMonth: Exit Function
    End If
    If Me.cboMonth > 12 Or Me.cboMonth < 1 Then
        fnData_Validation = False
        MsgBox "월 입력 값이 유효하지 않습니다. 다시 확인 해주세요.", vbCritical, banner
        Set txtBox_Focus = Me.cboMonth: Exit Function
    End If
End Function

Private Sub initializeRepport()
    Select Case SEARCH_CODE
    Case 1
        With Sheets("국가별 통계")
            
            '[영역설정]
            Set rngA = Range("Stat_Country_Start")
            Set rngB = .Columns("A").Find("합계", lookat:=xlWhole)
            
            '[입력내용초기화]
            rngA.Offset(1).Resize(rngB.Row - rngA.Row - 1, 32).ClearContents
            
            '[찌꺼기영역 제거]
            rngB.Offset(1).Resize(Rows.Count - rngB.Row - 1).EntireRow.Delete Shift:=xlUp
            
            '[마무리]
            Set rngA = Nothing
            Set rngB = Nothing
        End With
    Case 2
        With Sheets("교회통계")
            
            '[영역설정]
            Set rngA = Range("Stat_Church_Start")
            Set rngB = .Columns("A").Find("합계", lookat:=xlWhole)
            
            '[입력내용초기화]
            rngA.Offset(1).Resize(rngB.Row - rngA.Row - 1, 32).ClearContents
            
            '[찌꺼기영역 제거]
            rngB.Offset(1).Resize(Rows.Count - rngB.Row - 1).EntireRow.Delete Shift:=xlUp
            
            '[마무리]
            Set rngA = Nothing
            Set rngB = Nothing
        End With
    Case 3
        With Sheets("목회자통계")
            
            '[영역설정]
            Set rngA = Range("Stat_PStaff_Start")
            Set rngB = .Columns("A").Find("합계", lookat:=xlWhole)
            
            '[입력내용초기화]
            rngA.Offset(1).Resize(rngB.Row - rngA.Row - 1, 18).ClearContents
            
            '[찌꺼기영역 제거]
            rngB.Offset(1).Resize(Rows.Count - rngB.Row - 1).EntireRow.Delete Shift:=xlUp
            
            '[마무리]
            Set rngA = Nothing
            Set rngB = Nothing
        End With
    Case 4
        With Sheets("교회통계상세")
            
            '[영역설정]
            Set rngA = Range("Stat_ChurchDetail_Start")
            Set rngB = .Columns("A").Find("합계", lookat:=xlWhole)
            
            '[입력내용초기화]
            rngA.Offset(1).Resize(rngB.Row - rngA.Row - 1, 18).ClearContents
            
            '[찌꺼기영역 제거]
            rngB.Offset(1).Resize(Rows.Count - rngB.Row - 1).EntireRow.Delete Shift:=xlUp
            
            '[마무리]
            Set rngA = Nothing
            Set rngB = Nothing
        End With
    Case Else
    End Select
End Sub

Private Sub makeReport()
    Dim i As Integer
    Dim iRow As Integer, jRow As Integer
    Dim cntColumn As Integer
    
    Select Case SEARCH_CODE
    Case 1
        With Sheets("국가별 통계")
            '//영역설정
            Set rngA = Range("Stat_Country_Start")
            Set rngB = .Columns("A").Find("합계", lookat:=xlWhole)
            cntColumn = 33
            
            '//소제목1 리포트 작성
            '[보고행수]
            i = cntRecord
            
            '[보고서 영역 조정]
            iRow = rngB.Row - rngA.Row - 1 '현재 리포트 영역
            jRow = i - iRow '초과 리포트 영역
            
            If jRow > 0 Then '데이터가 제공된 영역보다 많은 경우
'                .Rows(rngB.Row - 1 & ":" & rngB.Row - 1 + jRow - 1).Insert Shift:=xlDown
                .Rows(rngB.Row).Resize(cntRecord - iRow).Insert Shift:=xlDown
                rngA.Offset(1).Resize(1, cntColumn).Copy .Range(rngA.Offset(1), rngA.Offset(2 + i - 2))
            ElseIf jRow < 0 And i <> 0 Then '데이터가 제공된 영역보다 적은 경우
                .Rows(rngB.Row - 2 & ":" & rngB.Row - 1 + jRow).Delete Shift:=xlUp
            ElseIf jRow < 0 And i = 0 And iRow > 1 Then '조회 데이터가 없는 경우
                .Rows(rngB.Row - 2 & ":" & rngB.Row + jRow).Delete Shift:=xlUp
            End If
            
            '//찌꺼기 영역 제거
            i = Cells(Rows.Count, "A").End(xlUp).Row
            Cells(Rows.Count, "A").End(xlUp).Offset(1).Resize(Rows.Count - i).EntireRow.Delete Shift:=xlUp
            With Range(rngA.Offset(1), rngB.Offset(-1)).Resize(, cntColumn).Borders(xlInsideHorizontal)
                .LineStyle = xlContinuous
                .ColorIndex = xlAutomatic
                .TintAndShade = 0
                .Weight = xlHairline
            End With
            With Range(rngA.Offset(1), rngB.Offset(-1)).Resize(, cntColumn).Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            
            '//함수삽입
            Set rngB = .Columns("A").Find("합계", lookat:=xlWhole)
            Range(rngA.Offset(1, -1), rngB.Offset(-1)).Formula = "=row()-3"
            rngB.Offset(, 2).Resize(, cntColumn - 2).FormulaR1C1 = "=SUM(R[-" & cntRecord & "]C:R[-1]C)"
            
            '//마무리
            Set rngA = Nothing
            Set rngB = Nothing
        End With
    Case 2
        With Sheets("교회통계")
            '//영역설정
            Set rngA = Range("Stat_Church_Start")
            Set rngB = .Columns("A").Find("합계", lookat:=xlWhole)
            cntColumn = 32
            
            '//소제목1 리포트 작성
            '[보고행수]
            i = cntRecord
            
            '[보고서 영역 조정]
            iRow = rngB.Row - rngA.Row - 1 '현재 리포트 영역
            jRow = i - iRow '초과 리포트 영역
            
            If jRow > 0 Then '데이터가 제공된 영역보다 많은 경우
'                .Rows(rngB.Row - 1 & ":" & rngB.Row - 1 + jRow - 1).Insert Shift:=xlDown
                .Rows(rngB.Row).Resize(cntRecord - iRow).Insert Shift:=xlDown
                rngA.Offset(1).Resize(1, cntColumn).Copy .Range(rngA.Offset(1), rngA.Offset(2 + i - 2))
            ElseIf jRow < 0 And i <> 0 Then '데이터가 제공된 영역보다 적은 경우
                .Rows(rngB.Row - 2 & ":" & rngB.Row - 1 + jRow).Delete Shift:=xlUp
            ElseIf jRow < 0 And i = 0 And iRow > 1 Then '조회 데이터가 없는 경우
                .Rows(rngB.Row - 2 & ":" & rngB.Row + jRow).Delete Shift:=xlUp
            End If
            
            '//찌꺼기 영역 제거
            i = Cells(Rows.Count, "A").End(xlUp).Row
            Cells(Rows.Count, "A").End(xlUp).Offset(1).Resize(Rows.Count - i).EntireRow.Delete Shift:=xlUp
            With Range(rngA.Offset(1), rngB.Offset(-1)).Resize(, cntColumn).Borders(xlInsideHorizontal)
                .LineStyle = xlContinuous
                .ColorIndex = xlAutomatic
                .TintAndShade = 0
                .Weight = xlHairline
            End With
            With Range(rngA.Offset(1), rngB.Offset(-1)).Resize(, cntColumn).Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            
            '//함수삽입
            Set rngB = .Columns("A").Find("합계", lookat:=xlWhole)
            Range(rngA.Offset(1, -1), rngB.Offset(-1)).Formula = "=row()-3"
            rngB.Offset(, 5).Resize(, cntColumn - 5).FormulaR1C1 = "=SUM(R[-" & cntRecord & "]C:R[-1]C)"
            
            '//마무리
            Set rngA = Nothing
            Set rngB = Nothing
        End With
    Case 3
        With Sheets("목회자통계")
            '//영역설정
            Set rngA = Range("Stat_PStaff_Start")
            Set rngB = .Columns("A").Find("합계", lookat:=xlWhole)
            cntColumn = 19
            
            '//소제목1 리포트 작성
            '[보고행수]
            i = cntRecord
            
            '[보고서 영역 조정]
            iRow = rngB.Row - rngA.Row - 1 '현재 리포트 영역
            jRow = i - iRow '초과 리포트 영역
            
            If jRow > 0 Then '데이터가 제공된 영역보다 많은 경우
'                .Rows(rngB.Row - 1 & ":" & rngB.Row - 1 + jRow - 1).Insert Shift:=xlDown
                .Rows(rngB.Row).Resize(cntRecord - iRow).Insert Shift:=xlDown
                rngA.Offset(1).Resize(1, cntColumn).Copy .Range(rngA.Offset(1), rngA.Offset(2 + i - 2))
            ElseIf jRow < 0 And i <> 0 Then '데이터가 제공된 영역보다 적은 경우
                .Rows(rngB.Row - 2 & ":" & rngB.Row - 1 + jRow).Delete Shift:=xlUp
            ElseIf jRow < 0 And i = 0 And iRow > 1 Then '조회 데이터가 없는 경우
                .Rows(rngB.Row - 2 & ":" & rngB.Row + jRow).Delete Shift:=xlUp
            End If
            
            '//찌꺼기 영역 제거
            i = Cells(Rows.Count, "A").End(xlUp).Row
            Cells(Rows.Count, "A").End(xlUp).Offset(1).Resize(Rows.Count - i).EntireRow.Delete Shift:=xlUp
            With Range(rngA.Offset(1), rngB.Offset(-1)).Resize(, cntColumn).Borders(xlInsideHorizontal)
                .LineStyle = xlContinuous
                .ColorIndex = xlAutomatic
                .TintAndShade = 0
                .Weight = xlHairline
            End With
            With Range(rngA.Offset(1), rngB.Offset(-1)).Resize(, cntColumn).Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            
            '//함수삽입
            Set rngB = .Columns("A").Find("합계", lookat:=xlWhole)
            Range(rngA.Offset(1, -1), rngB.Offset(-1)).Formula = "=row()-4"
            rngB.Offset(, 5).Resize(, cntColumn - 5).FormulaR1C1 = "=SUM(R[-" & cntRecord & "]C:R[-1]C)"
            
            '//마무리
            Set rngA = Nothing
            Set rngB = Nothing
        End With
    Case 4
        With Sheets("교회통계상세")
            '//영역설정
            Set rngA = Range("Stat_ChurchDetail_Start")
            Set rngB = .Columns("A").Find("합계", lookat:=xlWhole)
            cntColumn = 19
            
            '//소제목1 리포트 작성
            '[보고행수]
            i = cntRecord
            
            '[보고서 영역 조정]
            iRow = rngB.Row - rngA.Row - 1 '현재 리포트 영역
            jRow = i - iRow '초과 리포트 영역
            
            If jRow > 0 Then '데이터가 제공된 영역보다 많은 경우
'                .Rows(rngB.Row - 1 & ":" & rngB.Row - 1 + jRow - 1).Insert Shift:=xlDown
                .Rows(rngB.Row).Resize(cntRecord - iRow).Insert Shift:=xlDown
                rngA.Offset(1).Resize(1, cntColumn).Copy .Range(rngA.Offset(1), rngA.Offset(2 + i - 2))
            ElseIf jRow < 0 And i <> 0 Then '데이터가 제공된 영역보다 적은 경우
                .Rows(rngB.Row - 2 & ":" & rngB.Row - 1 + jRow).Delete Shift:=xlUp
            ElseIf jRow < 0 And i = 0 And iRow > 1 Then '조회 데이터가 없는 경우
                .Rows(rngB.Row - 2 & ":" & rngB.Row + jRow).Delete Shift:=xlUp
            End If
            
            '//찌꺼기 영역 제거
            i = Cells(Rows.Count, "A").End(xlUp).Row
            Cells(Rows.Count, "A").End(xlUp).Offset(1).Resize(Rows.Count - i).EntireRow.Delete Shift:=xlUp
            With Range(rngA.Offset(1), rngB.Offset(-1)).Resize(, cntColumn).Borders(xlInsideHorizontal)
                .LineStyle = xlContinuous
                .ColorIndex = xlAutomatic
                .TintAndShade = 0
                .Weight = xlHairline
            End With
            With Range(rngA.Offset(1), rngB.Offset(-1)).Resize(, cntColumn).Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            
            '//함수삽입
            Set rngB = .Columns("A").Find("합계", lookat:=xlWhole)
            Range(rngA.Offset(1, -1), rngB.Offset(-1)).Formula = "=row()-3"
            rngB.Offset(, 6).Resize(, cntColumn - 10).FormulaR1C1 = "=SUMIF(R4C3:R" & cntRecord + 3 & "C3,""*MC*"",R[-" & cntRecord & "]C:R[-1]C)+SUMIF(R4C3:R" & cntRecord + 3 & "C3,""*HBC*"",R[-" & cntRecord & "]C:R[-1]C)"
            
            '//마무리
            Set rngA = Nothing
            Set rngB = Nothing
        End With
    Case Else
    End Select
End Sub



