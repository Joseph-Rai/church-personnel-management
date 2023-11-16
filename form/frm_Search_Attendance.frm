VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Search_Attendance 
   Caption         =   "교회 검색"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7155
   OleObjectBlob   =   "frm_Search_Attendance.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "frm_Search_Attendance"
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
Dim ws As Worksheet

Private Sub chkAllSelect_Change()
    If Me.chkAllSelect Then
        Me.Height = 220
        setValueSelectCheckBox (True)
    Else
        Me.Height = 276
        setValueSelectCheckBox (False)
    End If
    
    validateAnySelectionCheckBox
    
End Sub

Private Sub chkBC_Change()
    validateAnySelectionCheckBox
End Sub

Private Sub chkMC_Change()
    validateAnySelectionCheckBox
End Sub


Private Sub chkMM_Change()
    validateAnySelectionCheckBox
End Sub

Private Sub chkPBC_Change()
    validateAnySelectionCheckBox
End Sub

Private Sub cmdPrintPDFAllList_Click()
    Dim i As Integer
    Dim filePath As String
    Dim FileName As String
    Dim selectList As Object
    
    Set selectList = getSelectList
    
    Call Optimization
    
    If Me.lstChurch.ListCount <= 0 Then
        MsgBox "검색된 교회가 없습니다." & vbNewLine & _
                "먼저 내보내기 할 교회목록을 조회하세요.", vbCritical, banner
        GoTo Here
    End If
    
    If Me.chkAllSelect = False And Not isAllSelectedCheckBox Then
        
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
            Dim arrList As Variant
            arrList = selectList.ToArray
            If IsInArray(.List(.listIndex, 2), arrList, True, rtnValue) <> -1 Then
                Call cmdOK_Click
                filePath = fnPrintAsPDF(filePath)
            End If
        Next
    End With
    
    MsgBox "작업이 완료 되었습니다." & vbNewLine & vbNewLine & _
            "파일저장경로: " & filePath, , banner
Here:
    Call Normal
End Sub

Private Sub cmdPrintPDFCurrentPage_Click()
    Dim filePath As String
    
    filePath = fnPrintAsPDF
    MsgBox "작업이 완료 되었습니다." & vbNewLine & vbNewLine & _
            "파일저장경로: " & filePath, , banner
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
    
    Dim intYear As Integer
    Dim intMonth As Integer
    
    If connIP = Empty Then setGlobalVariant '//전역변수 재설정
    If checkLogin = 0 Then f_login.Show '//로그인체크
    If checkLogin = 0 Then End '--//로그인 실패 시 프로시저 종료
    
    '--//시트설정
    Set ws = ActiveSheet
    
    '--//기초설정
    Me.cmdClose.Cancel = True
    TB1 = "op_system.db_churchlist_custom" '--//교회리스트
    TB2 = "op_system.temp_pstaff_by_time" '--//교역자 정보
    TB3 = "op_system.v_history_church" '--//교회연혁
    TB4 = "op_system.db_attendance" '--//출석현황
    
    '--//리스트박스 설정
    With Me.lstChurch
        .ColumnCount = 4
        .ColumnHeads = False
        .ColumnWidths = "0,150,50,150,0" '교회코드, 교회명, 교회구분, 관리교회명, 커스텀코드
        .Width = 265.5
        .TextAlign = fmTextAlignLeft
        .Font = "굴림"
    End With
    
    '--//콤보박스 아이템 추가
    For intYear = year(Date) To 2005 Step -1 '--//년도 채우기
        Me.cboYear.AddItem intYear
    Next
    For intMonth = 12 To 1 Step -1 '--//월 채우기
        Me.cboMonth.AddItem intMonth
    Next
    
    '--//콤보박스 현재 설정된 날짜로 내용추가
    If Range("Atten_rngDate") = "" Then
        Me.cboYear = year(Date)
        Me.cboMonth = month(Date)
    Else
        If Range("Atten_MaxDate") <> WorksheetFunction.EDate(DateSerial(year(Date), month(Date), 1), -1) Then
            Me.cboYear = year(WorksheetFunction.EDate(Date, -1))
            Me.cboMonth = month(WorksheetFunction.EDate(Date, -1))
        Else
            Me.cboYear = year(Range("Atten_rngDate"))
            Me.cboMonth = month(Range("Atten_rngDate"))
        End If
    End If
    
    '--//체크박스 초기화
    Me.chkAllSelect = True
    
    '--//높이 초기화
    Me.Height = 220
    
    Me.txtChurch.SetFocus
End Sub

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
    Me.cmdOk.Enabled = True
End Sub
Private Sub cmdClose_Click()
    Unload Me
End Sub
Private Sub cmdPrint_Click()
    ActiveWindow.SelectedSheets.PrintOut FROM:=1, to:=1, Copies:=1
End Sub
Private Sub cmdOK_Click()
    
    Dim filePath As String
    Dim FileName As String
    Dim result As T_RESULT
    Dim shp As Shape
    
    '--//시트 활성화 및 잠금해제
    WB_ORIGIN.Activate
    ws.Activate
    Call shUnprotect(globalSheetPW)
    
    '--//교회 선택여부 판단
    If Me.lstChurch.listIndex = -1 Then
        MsgBox "교회를 선택하세요.", vbCritical, "오류"
        Exit Sub
    End If
    
    '--//날짜 입력여부
    If Me.cboYear <> "" Or Me.cboMonth <> "" Then
        If Me.cboYear = "" And Me.cboMonth <> "" Then
            MsgBox "검색 년을 입력해 주세요.", vbCritical, banner
            Exit Sub
        End If
        
        If Me.cboYear <> "" And Me.cboMonth = "" Then
            MsgBox "검색 월을 입력해 주세요.", vbCritical, banner
            Exit Sub
        End If
        
        Range("Atten_rngDate") = DateSerial(Me.cboYear, Me.cboMonth, 1)
    End If
    
    '--//교회정보 입력
    With Me.lstChurch
        Range("E2:O2").ClearContents '--//기존 값 초기화
        Range("P2:R2").ClearContents
        Range("P3:R3").ClearContents
        
        '--//교회명 입력
        If .List(.listIndex, 2) = "MC" Then
            Range("E2") = .List(.listIndex, 1) & " 전체"
        Else
            Range("E2") = .List(.listIndex, 1)
        End If
        
        '--//교회구분 입력
        Range("P2:R2").ClearContents
        If .List(.listIndex, 2) = "BC" Then
            Range("P2") = "지교회"
        ElseIf .List(.listIndex, 2) = "PBC" Then
            Range("P2") = "예배소"
        End If
        
        '--//관리교회 입력
        Range("P3:R3").ClearContents
        If Range("P2") <> "" Then
            Range("P3") = .List(.listIndex, 3)
        End If
    End With
    
    Call Optimization
    
    '--//목회자 정보 삽입
    Range("Atten_rngTarget").CurrentRegion.ClearContents '--//기존 데이터 삭제
    
    strSql = "CALL `Routine_pstaff_by_time`(" & SText(Format(WorksheetFunction.EoMonth(DateSerial(Me.cboYear, Me.cboMonth, 1), 0), "yyyy-mm-dd")) & ", " & SText(USER_DEPT) & ");"
    connectTaskDB
    result.strSql = strSql
    result.affectedCount = executeSQL("cmdOK_Click", "temp_pstaff_by_time", strSql, Me.Name, "임시조회 테이블 생성")
'    writeLog "cmdADD_Click", "temp_pstaff_by_time", strSQL, 0, Me.Name, "임시조회 테이블 생성", result.affectedCount
    disconnectALL
    
    strSql = makeSelectSQL(TB2) '--//SQL문
    Call makeListData(strSql, TB2) '--//ListData 만들기
    
    '--//반환된 ListData를 보고서 시트에 삽입
    If cntRecord > 0 Then
        Range("Atten_rngTarget").Resize(1, UBound(LISTFIELD) + 1) = LISTFIELD
        Range("Atten_rngTarget").Offset(1).Resize(cntRecord, UBound(LISTFIELD) + 1) = LISTDATA
    End If
    Call sbClearVariant
    
    '--//교회연혁 삽입
'    Range("Atten_rngHistory").Resize(10, 2) = vbNullString
    Range("Atten_rngHistory_Data").CurrentRegion.ClearContents
    strSql = makeSelectSQL(TB3)
    Call makeListData(strSql, TB3)
    If cntRecord > 0 Then
        Range("Atten_rngHistory_Data").Resize(1, UBound(LISTFIELD) + 1) = LISTFIELD
        Range("Atten_rngHistory_Data").Offset(1).Resize(cntRecord, UBound(LISTFIELD) + 1) = LISTDATA
        Range("Atten_rngHistory_cntRecord") = cntRecord
        Range("Atten_rngHistory_Index") = IIf(cntRecord - 9 < 1, 1, cntRecord - 9)
        
        For Each shp In ActiveSheet.Shapes
            If shp.Name Like "*Move*" Then
                If cntRecord > 10 Then
                    shp.Visible = msoTrue
                Else
                    shp.Visible = msoFalse
                End If
            End If
        Next
    End If
    
    Call sbClearVariant
    
    '--//출석데이터 삽입
    Range("Atten_rngAttendance_Data").CurrentRegion.ClearContents
    strSql = makeSelectSQL(TB4)
    Call makeListData(strSql, TB4)
    Range("Atten_rngAttendance_Data").Resize(, 10) = LISTFIELD
    If cntRecord > 0 Then
        Range("Atten_rngAttendance_Data").Offset(1).Resize(cntRecord, 10) = LISTDATA
        Range("Atten_cntRecord") = cntRecord
        Range("A1").Copy
        Range("Atten_rngAttendance_Data").Offset(1).Resize(cntRecord, 10).PasteSpecial Paste:=xlPasteValues, operation:=xlPasteSpecialOperationAdd
    End If
    Application.CutCopyMode = False
    Call sbClearVariant
    
    If Range("Atten_rngDate") > Range("Atten_MaxDate") And IsDate(Range("Atten_MaxDate")) Then
        Range("Atten_rngDate") = Range("Atten_MaxDate")
    End If
    
'------------------------------------사진삽입--------------------------------------------
On Error Resume Next
    '--//기존 사진 삭제
    ActiveSheet.Pictures.Delete
    
    InsertPStaffPic Range("Atten_LifeNo"), Range("Atten_Pic_M") '--//선지자 사진삽입
    If Not (Range("Atten_LifeNo_Spouse") = "" Or Range("Atten_LifeNo_Spouse") = 0) Then
        InsertPStaffPic Range("Atten_LifeNo_Spouse"), Range("Atten_Pic_F") '--//배우자 사진삽입
    End If
    InsertChurchMap Range("Atten_ChurchCode"), Range("Atten_Church_Map") '--//교회지도 삽입
    
    '--//마지막에 가짜사진 하나 추가 후 삭제하여 뒤틀어짐 방지
    InsertPStaffPic "", Range("T17")
    If ActiveSheet.Pictures.Count > 0 Then
        ActiveSheet.Pictures(ActiveSheet.Pictures.Count).Delete
    End If
On Error GoTo 0

    '--//차트조정
    sbArrangeChart_Atten
    
    Normal
    
    Range("A2").Select
    
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
'        MsgBox "반환할 DB 데이터가 없습니다.", vbInformation, banner
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
        If Me.chkAll Then
            strSql = "SELECT a.church_sid,a.church_nm,a.church_gb,b.church_nm,c.church_sid_custom " & _
                        "FROM " & TB1 & " a LEFT JOIN op_system.db_churchlist_custom b ON a.main_church_cd = b.church_sid " & _
                        "LEFT JOIN op_system.db_history_church_establish c ON c.church_sid=REPLACE(a.church_sid,'MM','MC') " & _
                        "WHERE a.ovs_dept = " & USER_DEPT & " AND (a.church_nm LIKE '%" & Me.txtChurch & "%' OR b.church_nm LIKE '%" & Me.txtChurch & "%') ORDER BY a.sort_order;"
        Else
            strSql = "SELECT a.church_sid,a.church_nm,a.church_gb,b.church_nm,c.church_sid_custom " & _
                        "FROM " & TB1 & " a LEFT JOIN op_system.db_churchlist_custom b ON a.main_church_cd = b.church_sid " & _
                        "LEFT JOIN op_system.db_history_church_establish c ON c.church_sid=REPLACE(a.church_sid,'MM','MC') " & _
                        "WHERE a.ovs_dept = " & USER_DEPT & " AND a.suspend = 0 AND (a.church_nm LIKE '%" & Me.txtChurch & "%' OR b.church_nm LIKE '%" & Me.txtChurch & "%') ORDER BY a.sort_order;"
        End If
    Case TB2
        With Me.lstChurch
'            strSQL = "SELECT b.`생명번호`,b.`한글이름(직분)`,b.`배우자생번`,b.`사모한글이름(직분)`,b.`최초발령일`,b.`현당회발령일` FROM " & TB2 & " b WHERE b.`관리부서` = " & SText(USER_DEPT) & " AND b.`지교회명` = " & Replace(SText(.List(.ListIndex, 1)), " 본교회", "") & IIf(InStr(.List(.ListIndex, 2), "M") > 0, " AND b.`직책` LIKE '%당%'", "") & ";"
            strSql = "SELECT b.* FROM " & TB2 & " b WHERE b.`관리부서` = " & SText(USER_DEPT) & " AND b.`지교회명` = " & Replace(SText(.List(.listIndex, 1)), " 본교회", "") & IIf(InStr(.List(.listIndex, 2), "M") > 0, " AND b.`직책` LIKE '%당%'", "") & ";"
        End With
    Case TB3
'        strSQL = "SELECT DATE_FORMAT(b.`날짜`,'%y년%c월') '날짜',b.`교회연혁` FROM (SELECT a.`날짜`,a.`교회연혁` FROM " & TB3 & " a WHERE a.`커스텀코드` = " & Replace(SText(Me.lstChurch.List(Me.lstChurch.ListIndex, 4)), "MM", "MC") & " ORDER BY a.`날짜` DESC LIMIT 10) b ORDER BY b.`날짜`;"
        strSql = "SELECT DATE_FORMAT(b.`날짜`,'%y년%c월') '날짜',b.`교회연혁` FROM (SELECT a.`날짜`,a.`교회연혁` FROM " & TB3 & " a WHERE a.`날짜` <= " & SText(Format(WorksheetFunction.EoMonth(DateSerial(Me.cboYear, Me.cboMonth, 1), 0), "yyyy-mm-dd")) & " AND a.`커스텀코드` = " & Replace(SText(Me.lstChurch.List(Me.lstChurch.listIndex, 4)), "MM", "MC") & " ORDER BY a.`날짜`) b ORDER BY b.`날짜`;"
    Case TB4
        If Left(Me.lstChurch.List(Me.lstChurch.listIndex), 2) = "MM" Then
            '--//순수 본교회 출석데이터
'            strSQL = "SELECT a.attendance_dt,a.once_all,a.forth_all,a.once_stu,a.forth_stu,a.tithe_stu,a.baptism_all,a.evangelist,a.ul,a.gl FROM " & TB4 & " a WHERE a.church_sid = " & SText(Me.lstChurch.List(Me.lstChurch.ListIndex)) & _
'                            " AND attendance_dt BETWEEN ADDDATE(" & SText(Range("Atten_rngDate")) & ", INTERVAL -1 year) AND " & SText(Range("Atten_rngDate")) & " ORDER BY a.attendance_dt;"
            strSql = "SELECT a.attendance_dt,a.once_all,a.forth_all,a.once_stu,a.forth_stu,a.tithe_stu,a.baptism_all,a.evangelist,a.ul,a.gl FROM " & TB4 & " a WHERE a.church_sid = " & SText(Me.lstChurch.List(Me.lstChurch.listIndex)) & _
                            " ORDER BY a.attendance_dt;"
            Call makeListData(strSql, TB4)
            
            If Not cntRecord > 0 Then '--//순수 본교회 출석이 없는 교회들은 전체 출석데이터 추출
                strSql = "SELECT a.church_sid_custom FROM op_system.db_history_church_establish a WHERE a.church_sid = " & SText(Replace(Me.lstChurch.List(Me.lstChurch.listIndex), "MM", "MC"))
                Call makeListData(strSql, "op_system.db_history_church_establish")
                
                If cntRecord > 0 Then
'                strSQL = "SELECT a.attendance_dt,a.once_all,a.forth_all,a.once_stu,a.forth_stu,a.tithe_stu,a.baptism_all,a.evangelist,a.ul,a.gl FROM " & TB4 & " a LEFT JOIN op_system.db_history_church_establish b ON a.church_sid = b.church_sid WHERE b.church_sid_custom = " & SText(listData(0, 0)) & " AND attendance_dt BETWEEN ADDDATE(" & _
'                            SText(Range("Atten_rngDate")) & ", INTERVAL -1 year) AND " & SText(Range("Atten_rngDate")) & " GROUP BY a.attendance_dt ORDER BY a.attendance_dt;"
                strSql = "SELECT a.attendance_dt,MAX(a.once_all),MAX(a.forth_all),MAX(a.once_stu),MAX(a.forth_stu),MAX(a.tithe_stu),MAX(a.baptism_all),MAX(a.evangelist),MAX(a.ul),MAX(a.gl) FROM " & TB4 & " a LEFT JOIN op_system.db_history_church_establish b ON a.church_sid = b.church_sid WHERE b.church_sid_custom = " & SText(LISTDATA(0, 0)) & _
                            "GROUP BY a.attendance_dt ORDER BY a.attendance_dt;"
                End If
            End If
        Else
            '--//그 외 출석 데이터
            strSql = "SELECT a.church_sid_custom FROM op_system.db_history_church_establish a WHERE a.church_sid = " & SText(Replace(Me.lstChurch.List(Me.lstChurch.listIndex), "MM", "MC"))
            Call makeListData(strSql, "op_system.db_history_church_establish")
            
            If cntRecord > 0 Then
'                strSQL = "SELECT a.attendance_dt,a.once_all,a.forth_all,a.once_stu,a.forth_stu,a.tithe_stu,a.baptism_all,a.evangelist,a.ul,a.gl FROM " & TB4 & " a LEFT JOIN op_system.db_history_church_establish b ON a.church_sid = b.church_sid WHERE b.church_sid_custom = " & SText(listData(0, 0)) & " AND attendance_dt BETWEEN ADDDATE(" & _
'                            SText(Range("Atten_rngDate")) & ", INTERVAL -1 year) AND " & SText(Range("Atten_rngDate")) & " GROUP BY a.attendance_dt ORDER BY a.attendance_dt;"
                strSql = "SELECT a.attendance_dt,MAX(a.once_all),MAX(a.forth_all),MAX(a.once_stu),MAX(a.forth_stu),MAX(a.tithe_stu),MAX(a.baptism_all),MAX(a.evangelist),MAX(a.ul),MAX(a.gl) FROM " & TB4 & " a LEFT JOIN op_system.db_history_church_establish b ON a.church_sid = b.church_sid WHERE b.church_sid_custom = " & SText(LISTDATA(0, 0)) & _
                            "GROUP BY a.attendance_dt ORDER BY a.attendance_dt;"
            End If
        End If
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

'Private Sub sbArrangeChart_Atten()
'
'Dim noMax As Integer
'    Dim noMin As Integer
'    Dim i As Long
'    Dim j As Long
'    Dim Term As Long
'
'On Error Resume Next
'    noMax = WorksheetFunction.Max(Range("F17:R17")) '--//학생이상 1회출석 최대값
'    noMin = WorksheetFunction.Min(Range("F19:R19")) '--//학생이상 4회출석 최소값
'    i = 1: j = 1
'
'
'
'      '--//출석 그래프 차트에서
'      With Sheets("교회별 출석현황").ChartObjects(1).Chart.Axes(xlValue)
'        '--//식구 규모에 따라 스케일을 달리 합니다.
'        Select Case noMax
'            Case Is <= 100: Term = 10
'            Case Is <= 500: Term = 50
'            Case Is <= 1000: Term = 100
'            Case Else: Term = 100
'        End Select
'
'        '--//범위의 최대값을 구합니다..
'        Do
'            If Term * i > noMax Then
'                .MaximumScale = Term * i
'                Exit Do
'            End If
'            i = i + 1
'        Loop
'
'        '--//범위의 최소값을 구합니다.
'        Do
'            If Term * j >= noMin * 0.9 Then
'                .MinimumScale = Term * (j - 1)
'                Exit Do
'            End If
'            j = j + 1
'        Loop
'
'        '--//범위의 최대값과 최소값의 차이가 4의 배수가 아니면 최대값 수정
'        Do
'            If (.MaximumScale - .MinimumScale) Mod 4 = 0 Then Exit Do
'            i = i + 1
'            .MaximumScale = Term * i
'        Loop
'
'        .MajorUnit = (.MaximumScale - .MinimumScale) / 4
'
'      End With
'
'      With Sheets("교회별 출석현황").ChartObjects(1).Chart.Axes(xlValue, xlSecondary)
'        .MaximumScale = Application.Max(WorksheetFunction.RoundUp(WorksheetFunction.Max(Range("F26:R26")), 1), WorksheetFunction.RoundUp(WorksheetFunction.Max(Range("F27:R27")), 1))
'        .MinimumScale = Application.Min(WorksheetFunction.RoundDown(WorksheetFunction.Min(Range("F26:R26")), 1), WorksheetFunction.RoundDown(WorksheetFunction.Min(Range("F27:R27")), 1))
'      End With
'
'On Error GoTo 0
'
'End Sub

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
    FileName = Me.lstChurch.listIndex + 1 & ". " & Range("E2") & ".pdf" '--//파일명은 교회이름
    If (Len(Dir(filePath, vbDirectory)) <= 0) Then
        MkDir (filePath)
    End If
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:=filePath & FileName
    
    fnPrintAsPDF = filePath
    
End Function

Private Function getSelectList() As Object
    Dim selectList As Object
    Set selectList = CreateObject("System.Collections.ArrayList")
    
    If Me.chkMC Then
        selectList.Add "MC"
    End If
    
    If Me.chkMM Then
        selectList.Add "MM"
    End If
    
    If Me.chkBC Then
        selectList.Add "BC"
    End If
    
    If Me.chkPBC Then
        selectList.Add "PBC"
    End If
    
    Set getSelectList = selectList
    
End Function

Private Sub setValueSelectCheckBox(arg As Boolean)

    Me.chkMC = arg
    Me.chkMM = arg
    Me.chkBC = arg
    Me.chkPBC = arg

End Sub

Private Function isAllSelectedCheckBox() As Boolean

    Dim result As Boolean

    result = False
    If Me.chkMC And Me.chkMM And Me.chkBC And Me.chkPBC Then
        result = True
    End If
    
    isAllSelectedCheckBox = result

End Function

Private Function isNothingSelectedCheckBox() As Boolean

    Dim result As Boolean

    result = False
    If Not (Me.chkMC Or Me.chkMM Or Me.chkBC Or Me.chkPBC Or Me.chkAllSelect) Then
        result = True
    End If
    
    isNothingSelectedCheckBox = result

End Function

Private Sub validateAnySelectionCheckBox()

    If isNothingSelectedCheckBox Then
        Me.cmdPrintPDFAllList.Enabled = False
    Else
        Me.cmdPrintPDFAllList.Enabled = True
    End If

End Sub
