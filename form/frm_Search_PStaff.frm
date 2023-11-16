VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Search_PStaff 
   Caption         =   "교회 검색"
   ClientHeight    =   8280.001
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5250
   OleObjectBlob   =   "frm_Search_PStaff.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "frm_Search_PStaff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cntlst1Col As Integer '//리스트 컬럼 수
Dim TB1 As String, TB2 As String '//폼에 연결된 DB 테이블
Const strOrderBy As String = "currency_un ASC" '//DB에서 sort_order 필드
Dim caseSave As Integer '//1: 추가, 2: 수정, 3: 삭제(SUSPEND), 4: 완전삭제
Dim queryKey As Integer '//리스트 위치 반환에 사용될 id
Dim LISTDATA() As String '--//DB에서 받아온 rs를 배열로 저장
Dim LISTFIELD() As String '--//DB에서 받아온 rs의 필드를 배열로 저장
Dim cntRecord As Integer '--//DB에서 받아온 레코드의 개수
Dim strSql As String '--//SQL 쿼리문
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

Private Sub chkAll_Change()
    If Me.chkAll.Value = True Then
        Me.chkManager.Enabled = False
        Me.chkManager.Value = True
        Me.chkOther.Enabled = False
        Me.chkOther.Value = True
        Me.chkPastoral.Enabled = False
        Me.chkPastoral.Value = True
        Me.chkTheological.Enabled = False
        Me.chkTheological.Value = True
        Me.chkAll.SetFocus
    Else
        Me.chkManager.Enabled = True
        Me.chkManager.Value = False
        Me.chkOther.Enabled = True
        Me.chkOther.Value = False
        Me.chkPastoral.Enabled = True
        Me.chkPastoral.Value = False
        Me.chkTheological.Enabled = True
        Me.chkTheological.Value = False
    End If
End Sub

Private Sub chkDate_Click()
    If Me.chkDate.Value = True Then
        Me.cboYear.Enabled = True
        Me.cboMonth.Enabled = True
        Me.cboYear = year(Range("PStaff_rngDate"))
        Me.cboMonth = month(Range("PStaff_rngDate"))
    Else
        Me.cboYear.Enabled = False
        Me.cboMonth.Enabled = False
        Me.cboYear = year(Date)
        Me.cboMonth = month(Date)
    End If
End Sub

Private Sub chkManager_Click()
    If Me.chkPastoral.Value = True And Me.chkOther.Value = True And Me.chkTheological.Value = True And Me.chkManager.Value = True Then
        Me.chkAll.Value = True
    End If
End Sub

Private Sub chkOther_Click()
    If Me.chkManager.Value = True And Me.chkPastoral.Value = True And Me.chkTheological.Value = True And Me.chkOther.Value = True Then
        Me.chkAll.Value = True
    End If
End Sub

Private Sub chkPastoral_Click()
    If Me.chkManager.Value = True And Me.chkOther.Value = True And Me.chkTheological.Value = True And Me.chkPastoral.Value = True Then
        Me.chkAll.Value = True
    End If
End Sub

Private Sub chkTheological_Click()
    If Me.chkManager.Value = True And Me.chkOther.Value = True And Me.chkPastoral.Value = True And Me.chkTheological.Value = True Then
        Me.chkAll.Value = True
    End If
End Sub

Private Sub cmdPrint_Click()
    Call sbPrint_PStaff
End Sub

Private Sub cmdPrintPDF_Click()
    Dim filePath As String
    
    filePath = fnPrintAsPDF
    MsgBox "작업이 완료 되었습니다." & vbNewLine & vbNewLine & _
            "파일저장경로: " & filePath, , banner
    
End Sub

Private Sub cmdPrintPDFAllList_Click()
    Dim i As Integer
    Dim filePath As String
    Dim FileName As String
    
    Call Optimization
    
    If Me.lstChurch.ListCount <= 0 Then
        MsgBox "검색된 교회가 없습니다." & vbNewLine & _
                "먼저 내보내기 할 교회목록을 조회하세요.", vbCritical, banner
        GoTo Here
    End If
    
    '--//시간이 오래 소요될 수 있으므로 경고 메세지 띄우기
    If MsgBox("해당 기능은 검색된 교회목록 전체를 한 번씩 조회하면서" & vbNewLine & _
                "PDF로 출력하므로 상당한 시간이 소요될 수 있습니다." & vbNewLine & vbNewLine & _
                "계속 진행 하시겠습니까?", vbYesNo + vbInformation, banner) = vbNo Then
        GoTo Here
    End If
    
    filePath = GetDesktopPath '--//파일 경로 설정
    filePath = filePath & "ExportByPDF"
    filePath = FileSequence(filePath) & Application.PathSeparator
    
    With Me.lstChurch
        For i = 0 To .ListCount - 1
            Me.lstChurch.listIndex = i
            Call cmdOK_Click
            filePath = fnPrintAsPDF(filePath)
        Next
    End With
    
    MsgBox "작업이 완료 되었습니다." & vbNewLine & vbNewLine & _
            "파일저장경로: " & filePath, , banner
Here:
    Call Normal
    
End Sub

Private Sub lstChurch_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Un'HookListBoxScroll
End Sub

Private Sub lstChurch_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Me.lstChurch.ListCount > 0 Then
        'HookListBoxScroll Me, Me.lstChurch
    End If
End Sub

Private Sub UserForm_Initialize()
    
    Dim i As Long
    
    If connIP = Empty Then setGlobalVariant '//전역변수 재설정
    If checkLogin = 0 Then f_login.Show '//로그인체크
    If checkLogin = 0 Then End '--//로그인 실패 시 프로시저 종료
    
    '--//시트설정
    Set ws = ActiveSheet
    
    '--//기초설정
    Me.cmdClose.Cancel = True
    TB1 = "op_system.db_churchlist_custom" '--//교회리스트
    TB2 = "op_system.temp_pstaff_by_time" '--//선지자현황
    
    '--//컨트롤 설정
    Me.cmdClose.Cancel = True
    Me.lblStatus.Visible = False
    Me.chkManager.Enabled = False
    Me.chkManager.Value = True
    Me.chkOther.Enabled = False
    Me.chkOther.Value = True
    Me.chkPastoral.Enabled = False
    Me.chkPastoral.Value = True
    Me.chkTheological.Enabled = False
    Me.chkTheological.Value = True
    Me.chkAll.Value = True
'    Me.cmdOK.Visible = False
    Me.cboYear.Enabled = False
    Me.cboMonth.Enabled = False
    Me.optSort1.Value = True
    
    '--//콤보박스 현재 년월로 내용추가
    Me.cboYear = year(Date)
    Me.cboMonth = month(Date)
    
    '--//콤보박스 아이템 추가
    With Me.cboYear
        For i = year(Date) To year(Date) - 10 Step -1
            .AddItem i
        Next
    End With
    
    With Me.cboMonth
        For i = 1 To 12
            .AddItem i
        Next
    End With
    
    '--//리스트박스 설정
    With Me.lstChurch
        .ColumnCount = 2
        .ColumnHeads = False
        .ColumnWidths = "0,120" '교회코드, 교회명
        .Width = 241.45
        .TextAlign = fmTextAlignLeft
        .Font = "굴림"
    End With
    
    Me.txtChurch.SetFocus
End Sub

Private Sub cmdSearch_Click()
    strSql = makeSelectSQL(TB1)
    Call makeListData(strSql, TB1)
    If cntRecord > 0 Then
        Me.lstChurch.List = LISTDATA
    End If
    Call sbClearVariant
End Sub

Private Sub txtChurch_Change()
    Me.txtChurch.BackColor = RGB(255, 255, 255)
End Sub

Private Sub lstChurch_Click()
    cmdOk.Enabled = True
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub
Private Sub cmdOK_Click()
    Dim i As Integer, j As Integer
    Dim result As T_RESULT
    Dim page As Long
    Dim t As Single
    
    t = Timer
    
    '--//교회 선택여부 판단
    If Me.lstChurch.listIndex = -1 Then
        MsgBox "교회를 선택하세요.", vbCritical, "오류"
        Exit Sub
    End If
    
    '--//시트 활성화 및 잠금해제
    WB_ORIGIN.Activate
    ws.Activate
    Call shUnprotect(globalSheetPW)
    
'Debug.Print "시트활성화 및 잠금해제: " & Format(Timer - t, "#0.00")
    
Application.Calculation = xlCalculationManual '자동계산 죽이기
    
    '--//기존 데이터 삭제
    Range("PStaff_rngTarget").CurrentRegion.ClearContents
    
    '--//temp_pstaff_by_time 테이블 업데이트
    strSql = "CALL `Routine_pstaff_by_time`(" & SText(Format(WorksheetFunction.EoMonth(DateSerial(Me.cboYear, Me.cboMonth, 1), 0), "yyyy-mm-dd")) & ", " & SText(USER_DEPT) & ");"
    connectTaskDB
    result.strSql = strSql
    result.affectedCount = executeSQL("cmdADD_Click", "temp_pstaff_by_time", strSql, Me.Name, "임시조회 테이블 생성")
    disconnectALL

    '--//SQL문
    strSql = makeSelectSQL2
    
    '--//DB에서 자료 호출하여 레코드셋 반환
    Call makeListData(strSql, TB2)
    
    '--//반환된 ListData를 보고서 시트에 삽입
    Optimization
    If cntRecord > 0 Then
        Range("PStaff_rngTarget").Resize(cntRecord, UBound(LISTFIELD) + 1) = LISTFIELD
        Range("PStaff_rngTarget").Offset(1).Resize(cntRecord, UBound(LISTFIELD) + 1) = LISTDATA
    End If
    
    Normal
    
'Debug.Print "DB에서 자료 가져오기: " & Format(Timer - t, "#0.00")
    
'    Application.Wait (Now + TimeValue("0:00:02")) '--//정렬이 끝날 때까지 잠시 대기
    Optimization
    '--//변수 초기화
    page = Int(WorksheetFunction.Quotient(cntRecord - 1, 9)) + 1
    Call sbClearVariant
    
    '--//레코드 수에 따른 페이지 만들기
    Call sbClearPic '사진 초기화
    Call sbMakePage(page) '--필요한 개수만큼 페이지 생성
    
Application.Calculation = xlCalculationAutomatic '자동계산 살리기
'Application.CalculateFullRebuild
'Debug.Print "페이지 생성: " & Format(Timer - t, "#0.00")
    
    '--//사진삽입
    Sheets("교회별 선지자현황").Range("B1").Select
    Me.lblStatus.Visible = True
    Me.Repaint
    Call sbInsertPic
    Me.lblStatus.Visible = False
    Me.Repaint
'Application.CalculateFullRebuild
'Debug.Print "사진삽입: " + Format(Timer - t, "#0.00")
    
    '--//조회기준일 등록
    Range("PStaff_rngDate") = DateSerial(Me.cboYear, Me.cboMonth, 1)
    
    '--//인원통계 '0명/0명' 가리기
    ActiveSheet.Outline.ShowLevels RowLevels:=0, ColumnLevels:=2
    On Error Resume Next
    Range(Range("PStaff_Stat_cntByPosition").Offset(, 0), Range("PStaff_Stat_cntByPosition").Offset(, 20)).Columns.Ungroup
    On Error GoTo 0
    If Not IsError(Range("PStaff_Stat_cntByPosition").Offset(, 1)) Then
        If Range("PStaff_Stat_cntByPosition").Offset(, 1) = "0명/0명" Then
            Range(Range("PStaff_Stat_cntByPosition").Offset(, 0), Range("PStaff_Stat_cntByPosition").Offset(, 1)).Columns.Group
        End If
    Else
        Range(Range("PStaff_Stat_cntByPosition").Offset(, 0), Range("PStaff_Stat_cntByPosition").Offset(, 1)).Columns.Group
    End If
    If Not IsError(Range("PStaff_Stat_cntByPosition").Offset(, 3)) Then
        If Range("PStaff_Stat_cntByPosition").Offset(, 3) = "0명/0명" Then
            Range(Range("PStaff_Stat_cntByPosition").Offset(, 2), Range("PStaff_Stat_cntByPosition").Offset(, 3)).Columns.Group
        End If
    Else
        Range(Range("PStaff_Stat_cntByPosition").Offset(, 2), Range("PStaff_Stat_cntByPosition").Offset(, 3)).Columns.Group
    End If
    If Not IsError(Range("PStaff_Stat_cntByPosition").Offset(, 5)) Then
        If Range("PStaff_Stat_cntByPosition").Offset(, 5) = "0명/0명" Then
            Range(Range("PStaff_Stat_cntByPosition").Offset(, 4), Range("PStaff_Stat_cntByPosition").Offset(, 5)).Columns.Group
        End If
    Else
        Range(Range("PStaff_Stat_cntByPosition").Offset(, 4), Range("PStaff_Stat_cntByPosition").Offset(, 5)).Columns.Group
    End If
    If Not IsError(Range("PStaff_Stat_cntByPosition").Offset(, 7)) Then
        If Range("PStaff_Stat_cntByPosition").Offset(, 7) = "0명/0명" Then
            Range(Range("PStaff_Stat_cntByPosition").Offset(, 6), Range("PStaff_Stat_cntByPosition").Offset(, 7)).Columns.Group
        End If
    Else
        Range(Range("PStaff_Stat_cntByPosition").Offset(, 6), Range("PStaff_Stat_cntByPosition").Offset(, 7)).Columns.Group
    End If
    ActiveSheet.Outline.ShowLevels RowLevels:=0, ColumnLevels:=1
    
    '--//인쇄회수 초기화
    Range("PStaff_rngPrint").ClearContents
'Debug.Print "인원통계작업: " + Format(Timer - t, "#0.00")
    Normal
    shProtect globalSheetPW
    
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
    
    '//리스팅할 레코드 수 검토
    If cntRecord = 0 Then
        MsgBox "반환할 DB 데이터가 없습니다.", vbInformation, banner
        Exit Sub
    End If
    
End Sub

'-----------------------------------------
'  조건별 Select SQL문 작성
'    - makeSelectSQL(검색어, 필터)
'    - DB에서 반환할 리스트 필드수정
'-----------------------------------------
Private Function makeSelectSQL(ByVal tableNM As String) As String
    Select Case tableNM
    Case TB1
        '--//교회코드, 교회명
        If Me.chkAllChurch Then
            strSql = "SELECT a.church_sid,a.church_nm " & _
                        "FROM " & TB1 & " a " & _
                        "WHERE a.ovs_dept = " & USER_DEPT & " AND a.church_gb IN ('MC','HBC') AND a.church_nm LIKE '%" & Me.txtChurch & "%' ORDER BY a.sort_order;"
        Else
            strSql = "SELECT a.church_sid,a.church_nm " & _
                        "FROM " & TB1 & " a " & _
                        "WHERE a.ovs_dept = " & USER_DEPT & " AND a.suspend = 0 AND a.church_gb IN ('MC','HBC') AND a.church_nm LIKE '%" & Me.txtChurch & "%' ORDER BY a.sort_order;"
        End If
    Case Else
        '//화폐id, 화폐약칭, 조회일, 원화환율, 달러화환율
        'strSQL = "SELECT a.currency_id, a.currency_un, a.refer_dt, a.fx_rate_krw, a.fx_rate_usd " & _
                      "FROM " & TB1 & " a WHERE a.user_id = " & user_id & ";"
    End Select
    makeSelectSQL = strSql
End Function
'-----------------------------------------
'  조건별 Select SQL문 작성
'    - makeSelectSQL(검색어, 필터)
'    - DB에서 반환할 리스트 필드수정
'-----------------------------------------
Private Function makeSelectSQL2() As String
    
    '조건별 sql문 생성
    strSql = makeSelectSqlByCondition()
    
    'Order By절 생성
    strSql = addOrderByClause(strSql)
    
    makeSelectSQL2 = strSql
End Function
Private Function addOrderByClause(query As String)

    If Me.optSort1.Value Then
        query = query & " Order By a.`직책` IS NULL ASC,FIELD(a.`직책`,'당회장','당사모','당회장대리','당대리사모','동역','동사모','지교회관리자','지교회관리자사모','예배소관리자','예비생도3단계','예비생도2단계','예비생도1단계','예배소관리자사모', " & getPosition2Joining & ", '장지역장','부장지역장','(임)장지역장','장구역장','부장구역장','(임)장구역장','청지역장','부청지역장','(임)청지역장','청구역장','부청구역장','(임)청구역장','학지역장','부학지역장','(임)학지역장','학구역장','부학구역장','(임)학구역장','직책없음',NULL),a.`직분` IS NULL ASC,FIELD(a.`직분`," & getTitleJoining & ",NULL),a.`최초발령일`,a.`생년월일`;"
    End If
    
    If Me.optSort2.Value Then
        query = query & " Order By a.`직책` IS NULL ASC,FIELD(a.`직책`,'당회장','당사모','당회장대리','당대리사모','동역','동사모','지교회관리자','지교회관리자사모','예배소관리자','예비생도3단계','예비생도2단계','예비생도1단계','예배소관리자사모', " & getPosition2Joining & ", '장지역장','부장지역장','(임)장지역장','장구역장','부장구역장','(임)장구역장','청지역장','부청지역장','(임)청지역장','청구역장','부청구역장','(임)청구역장','학지역장','부학지역장','(임)학지역장','학구역장','부학구역장','(임)학구역장','직책없음',NULL),a.`지전체1회` IS NULL ASC,a.`지전체1회` DESC,a.`최초발령일`;"
    End If
    
    If Me.optSort3.Value Then
        query = query & " Order By a.`지전체1회` IS NULL DESC,a.`직책` IS NULL ASC,FIELD(a.`직책`,'당회장','당사모','당회장대리','당대리사모','동역','동사모','지교회관리자','지교회관리자사모','예배소관리자','예비생도3단계','예비생도2단계','예비생도1단계','예배소관리자사모', " & getPosition2Joining & ", '장지역장','부장지역장','(임)장지역장','장구역장','부장구역장','(임)장구역장','청지역장','부청지역장','(임)청지역장','청구역장','부청구역장','(임)청구역장','학지역장','부학지역장','(임)학지역장','학구역장','부학구역장','(임)학구역장','직책없음',NULL),a.`직분` IS NULL ASC,FIELD(a.`직분`, " & getTitleJoining & ",NULL),a.`지전체1회` DESC,a.`최초발령일`,a.`생년월일`;"
    End If
    
    If Me.optSort4.Value Then
        query = query & " Order By a.`지전체1회` IS NULL DESC,a.`직책` IS NULL ASC,FIELD(a.`직책`,'당회장','당사모','당회장대리','당대리사모','동역','동사모','지교회관리자','지교회관리자사모','예배소관리자','예비생도3단계','예비생도2단계','예비생도1단계','예배소관리자사모', " & getPosition2Joining & ", '장지역장','부장지역장','(임)장지역장','장구역장','부장구역장','(임)장구역장','청지역장','부청지역장','(임)청지역장','청구역장','부청구역장','(임)청구역장','학지역장','부학지역장','(임)학지역장','학구역장','부학구역장','(임)학구역장','직책없음',NULL),a.`지전체1회` DESC,a.`최초발령일`,a.`생년월일`;"
    End If
    
    addOrderByClause = query

End Function
Private Function makeSelectSqlByCondition()

    Dim chk1st As Boolean
    Dim result As String
    
    chk1st = True
    result = "SELECT * FROM " & TB2 & " a WHERE a.`교회명` = " & SText(Me.lstChurch.List(Me.lstChurch.listIndex, 1))
    
    If Me.chkAll.Value = False Then
        result = result & " AND ("
        If Me.chkPastoral.Value = True Then
            result = result & "a.`직책` LIKE '%당%' OR a.`직책` LIKE '%동%'"
            chk1st = False
        End If
        
        If Me.chkTheological.Value = True Then
            If chk1st = False Then
                result = result & " OR a.`생도기수` LIKE '%생도%'"
            Else
                result = result & " a.`생도기수` LIKE '%생도%'"
            End If
            chk1st = False
        End If
        
        If Me.chkManager.Value = True Then
            If chk1st = False Then
                result = result & " OR a.`직책` LIKE '%관리자%'"
            Else
                result = result & " a.`직책` LIKE '%관리자%'"
            End If
            chk1st = False
        End If
        
        If Me.chkOther.Value = True Then
            If chk1st = False Then
                result = result & " OR a.`직책2` IS NOT NULL"
            Else
                result = result & " a.`직책2` IS NOT NULL"
            End If
            chk1st = False
        End If
        
        result = result & ")"
        result = Replace(result, " AND ()", "") '기타만 선택 했을 시 나오는 에러 수정
        
    End If
    
    If Me.optNoLeaderExclude.Value Then
        result = result & " AND a.`생명번호` IS NOT NULL"
    End If
    '--//플래그 설정
    Range("PStaff_Stat_flagNoLeader") = Me.optNoLeaderInclude.Value
    
    makeSelectSqlByCondition = result

End Function
Public Sub sbInsertPic()

    Dim lifeNo As String
    Dim pageHeight As Integer: pageHeight = COUNT_PAGE_HEIGHT_CELLS
    Dim pageWidth As Integer: pageWidth = COUNT_PAGE_WIDTH_CELLS
    Dim lineHeight As Integer: lineHeight = (COUNT_PAGE_HEIGHT_CELLS - 1) / 3 '--//제목행 1 빼고 3으로 나눔
    Dim targetRange As Range: Set targetRange = Range("C8")
    
    '--//사진 초기화
    Call sbClearPic

    '--//사진 넣기 프로세스
On Error Resume Next
    Dim i As Long, j As Long
    Dim tmpRange As Range
    For j = targetRange.Row To targetRange.Offset(lineHeight * 2).Row Step lineHeight
        For i = targetRange.Column To Range("A1").Value '--//Range("A1"): 페이지 끝 열번호 추출
            '--//변수설정
            lifeNo = Cells(j, 1).Offset(, i - 1).Value
            Set tmpRange = Range(Cells(j, 1).Offset(, i - 1), Cells(j, 1).Offset(, i))
    
            '--//사진삽입
            If Not (lifeNo = "" Or lifeNo = "0") Then
                InsertPStaffPic lifeNo, tmpRange
            End If
        Next i
    Next j
    
    '--//사진 틀어짐 방지를 위한 마지막 사진 삭제
    InsertPStaffPic "", Range("A1")
    
    If ActiveSheet.Pictures.Count > 0 Then
        If Not ActiveSheet.Pictures(ActiveSheet.Pictures.Count).Name Like "*Stat_Shp*" Then
            ActiveSheet.Pictures(ActiveSheet.Pictures.Count).Delete
        End If
    End If
    
On Error GoTo 0

End Sub
Private Sub sbClearPic()
    Dim PicImage As Shape
    
    For Each PicImage In ActiveSheet.Shapes
        If PicImage.Name Like "*Stat_Shp*" Or PicImage.Name Like "*직사각형*" Or PicImage.Name Like "*Option*" Or PicImage.Name Like "*Rectangle*" Then
        Else
            PicImage.Delete
        End If
    Next
End Sub
Private Sub sbClearPicInRange(targetRange As Range)
    Dim PicImage As Shape
    
    For Each PicImage In ActiveSheet.Shapes
        If Not Intersect(PicImage.TopLeftCell, targetRange) Is Nothing Then
            PicImage.Delete
        End If
    Next
End Sub

Private Sub sbMakePage(page As Long)

    Dim targetCol As Long
    Dim i As Long
    Dim curPage As Long
    Dim PageStandard As Range
    Dim targetRange As Range
    
    '--// 전역상수
    '--// COUNT_PAGE_WIDTH_CELLS
    '--// COUNT_PAGE_HEIGHT_CELLS
    
    '--//필요페이지가 0이면 프로시저 종료
    If page = 0 Then Exit Sub
    
    '--//첫페이지 기준셀 설정
    Set PageStandard = Range("C1")
    
    '--//인쇄영역 정돈하기 -> 영문이름 표시될 경우 페이지가 2개로 나뉘는 오류 해결목적
On Error Resume Next
    ActiveSheet.HPageBreaks(1).DragOff Direction:=xlDown, RegionIndex:=1
On Error GoTo 0

    '--//현재 페이지 수 가져오기
    curPage = Application.ExecuteExcel4Macro("Get.Document(50)")
    
    '--//필요한 페이지 개수 맞추기 프로세스
    If curPage > page Then '기존 페이지가 필요 페이지보다 많을 때
        '기존 페이지 삭제 프로세스
        With Range(PageStandard.Offset(, COUNT_PAGE_WIDTH_CELLS * (page)), PageStandard.Offset(, COUNT_PAGE_WIDTH_CELLS * (curPage) - 1))
            .EntireColumn.Delete Shift:=xlLeft
        End With
    ElseIf curPage < page Then '기존 페이지가 필요 페이지보다 적을 때
        '신규 페이지 추가 프로세스
        Set targetRange = PageStandard.Offset(, 15 * curPage).Resize(, 15 * (page - curPage)).EntireColumn
        targetRange.Insert Shift:=xlRight
        PageStandard.Resize(COUNT_PAGE_HEIGHT_CELLS, COUNT_PAGE_WIDTH_CELLS).Copy
        For i = curPage To page - 1
            With PageStandard.Offset(, 15 * i).Resize(COUNT_PAGE_HEIGHT_CELLS, COUNT_PAGE_WIDTH_CELLS)
                .PasteSpecial Paste:=xlPasteFormats
                .PasteSpecial Paste:=xlPasteColumnWidths
                .PasteSpecial Paste:=xlPasteFormulas
            End With
        Next
        
        '잡다한 것 정리하기
        Set targetRange = PageStandard.Resize(COUNT_PAGE_HEIGHT_CELLS, COUNT_PAGE_WIDTH_CELLS)
        Call sbClearPicInRange(targetRange.Offset(, 15).Resize(2, 15 * (page - 1)))
    Dim rngUnion1 As Range
    Dim rngUnion2 As Range
    Dim rngUnion3 As Range
        For i = 1 To page - 1
            If rngUnion1 Is Nothing Then
                Set rngUnion1 = targetRange.Offset(, 15).Resize(1, 3).Offset(1, 15 * i - 4)
            Else
                Set rngUnion1 = UNION(rngUnion1, targetRange.Offset(, 15).Resize(1, 3).Offset(1, 15 * i - 4))
            End If
            
            If rngUnion2 Is Nothing Then
                Set rngUnion2 = targetRange.Offset(, 15).Resize(1, 2).Offset(, 15 * i - 3)
            Else
                Set rngUnion2 = UNION(rngUnion2, targetRange.Offset(, 15).Resize(1, 2).Offset(, 15 * i - 3))
            End If
            
            If rngUnion3 Is Nothing Then
                Set rngUnion3 = targetRange.Offset(, 15).Resize(2, 1).Offset(2, 15 * (i - 1))
            Else
                Set rngUnion3 = UNION(rngUnion3, targetRange.Offset(, 15).Resize(2, 1).Offset(2, 15 * (i - 1)))
            End If
            
        Next
        targetRange.Offset(, 15).Resize(1, 3).Offset(1, 6).Copy
        rngUnion1.PasteSpecial Paste:=xlPasteAll
        rngUnion2.ClearContents
        rngUnion3.FormatConditions.Delete
        
        '--//프린트영역 설정
        Set targetRange = PageStandard.Resize(COUNT_PAGE_HEIGHT_CELLS, COUNT_PAGE_WIDTH_CELLS)
        ActiveSheet.PageSetup.PrintArea = targetRange.Resize(, COUNT_PAGE_WIDTH_CELLS * page).Address
        For i = 1 To page - 1
            Set ActiveSheet.VPageBreaks(i).Location = PageStandard.Offset(, COUNT_PAGE_WIDTH_CELLS * i)
        Next
        
    Else
        '아무것도 안하기
    End If
    
End Sub

Private Sub sbClearVariant()
    Erase LISTFIELD
    Erase LISTDATA
    cntRecord = Empty
    strSql = vbNullString
End Sub
Private Sub sbSortData_PStaff()

    ActiveWorkbook.Worksheets("교회별 선지자현황").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("교회별 선지자현황").Sort.SortFields.Add key:=Range("PStaff_rngTarget").Offset(1, 8).Resize(Range("PStaff_rngTarget").CurrentRegion.Rows.Count - 1) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder:= _
        "당회장,당사모,당회장대리,당대리사모,동역,동사모,지교회관리자,지교회관리자사모,예배소관리자,예비생도3단계,예비생도2단계,예비생도1단계,예배소관리자사모,장지역장,(임)장지역장,장구역장,(임)장구역장,청지역장,(임)청지역장,청구역장,(임)청구역장,학지역장,(임)학지역장,학구역장,(임)학구역장" _
        , DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("교회별 선지자현황").Sort.SortFields.Add key:=Range("PStaff_rngTarget").Offset(1, 12).Resize(Range("PStaff_rngTarget").CurrentRegion.Rows.Count - 1) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("교회별 선지자현황").Sort
        .SetRange Range("PStaff_rngTarget").CurrentRegion
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

Private Function fnPrintAsPDF(Optional filePath As String)

    Dim FileName As String
    
    '--//시트 활성화 및 잠금해제
    WB_ORIGIN.Activate
    ws.Activate
    
    '--//PDF로 내보내기
    If filePath = "" Then
        filePath = GetDesktopPath '--//파일 경로 설정
        filePath = filePath & "ExportByPDF"
        filePath = FileSequence(filePath) & Application.PathSeparator
    End If
    FileName = Range("E1") & ".pdf" '--//파일명은 교회이름
    If (Len(Dir(filePath, vbDirectory)) <= 0) Then
        MkDir (filePath)
    End If
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:=filePath & Me.lstChurch.listIndex + 1 & ". " & FileName
    
    fnPrintAsPDF = filePath
    
End Function
Private Function getPosition2Joining()

    Dim strQuery As String
    strQuery = "SELECT * FROM op_system.a_position2;"
    Call makeListData(strQuery, "op_system.a_position2")
        
    Dim result As String
    Dim i As Integer
    For i = 0 To cntRecord - 1
        If i < cntRecord - 1 Then
            result = result & "'" & LISTDATA(i, 0) & "', "
        Else
            result = result & "'" & LISTDATA(i, 0) & "'"
        End If
    Next
    
    getPosition2Joining = result

End Function
Private Function getTitleJoining()

    Dim strQuery As String
    strQuery = "SELECT * FROM op_system.a_title;"
    Call makeListData(strQuery, "op_system.a_title")
        
    Dim result As String
    Dim i As Integer
    For i = 0 To cntRecord - 1
        If i < cntRecord - 1 Then
            result = result & "'" & LISTDATA(i, 0) & "', "
        Else
            result = result & "'" & LISTDATA(i, 0) & "'"
        End If
    Next
    
    getTitleJoining = result

End Function
