VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Update_Attendance 
   Caption         =   "출석데이터 관리마법사"
   ClientHeight    =   6960
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   8310.001
   OleObjectBlob   =   "frm_Update_Attendance.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "frm_Update_Attendance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cntlst1Col As Integer '//리스트 컬럼 수
Dim TB1 As String, TB2 As String, TB3 As String, TB4 As String '//폼에 연결된 DB 테이블
Const strOrderBy As String = "currency_un ASC" '//DB에서 sort_order 필드
Dim caseSave As Integer '//1: 추가, 2: 수정, 3: 삭제(SUSPEND), 4: 완전삭제
Dim queryKey As Variant '//리스트 위치 반환에 사용될 id
Dim LISTDATA() As String '--//DB에서 받아온 rs를 배열로 저장
Dim LISTFIELD() As String '--//DB에서 받아온 rs의 필드를 배열로 저장
Dim cntRecord As Integer '--//DB에서 받아온 레코드의 개수
Dim strSql As String '--//SQL 쿼리문
Dim txtBox_Focus As MSForms.control

Private Sub lstAttendance_Click()
    
    '--//컨트롤 설정
    If Me.lstAttendance.listIndex <> -1 Then
        Me.cmdAdd.Enabled = True
        Me.cmdDelete.Enabled = True
        Me.cmdEdit.Enabled = True
        Me.txtOnce.Enabled = True
        Me.txtForth.Enabled = True
        Me.txtOnce_Stu.Enabled = True
        Me.txtForth_Stu.Enabled = True
        Me.txtTithe_All.Enabled = True
        Me.txtTithe_Stu.Enabled = True
        Me.txtBaptism.Enabled = True
        Me.txtEvangelist.Enabled = True
        Me.txtGL.Enabled = True
        Me.txtUL.Enabled = True
        Me.cboYear.Enabled = True
        Me.cboMonth.Enabled = False
    Else
        Call sbtxtBox_Init
        Me.cmdAdd.Enabled = False
        Me.cmdDelete.Enabled = False
        Me.cmdEdit.Enabled = False
        Me.txtOnce.Enabled = False
        Me.txtForth.Enabled = False
        Me.txtOnce_Stu.Enabled = False
        Me.txtForth_Stu.Enabled = False
        Me.txtTithe_All.Enabled = False
        Me.txtTithe_Stu.Enabled = False
        Me.txtBaptism.Enabled = False
        Me.txtEvangelist.Enabled = False
        Me.txtGL.Enabled = False
        Me.txtUL.Enabled = False
        Me.cboYear.Enabled = False
        Me.cboMonth.Enabled = False
    End If
    
    '--//내용삽입
    If Me.lstAttendance.listIndex <> -1 Then
        With Me.lstAttendance
            Me.txtOnce = .List(.listIndex, 2)
            Me.txtOnce_Stu = .List(.listIndex, 3)
            Me.txtForth = .List(.listIndex, 4)
            Me.txtForth_Stu = .List(.listIndex, 5)
            Me.txtTithe_All = .List(.listIndex, 6)
            Me.txtTithe_Stu = .List(.listIndex, 7)
            Me.txtBaptism = .List(.listIndex, 8)
            Me.txtEvangelist = .List(.listIndex, 9)
            Me.txtGL = .List(.listIndex, 10)
            Me.txtUL = .List(.listIndex, 11)
            Me.cboYear = Left(.List(.listIndex, 1), 4)
            Me.cboMonth = Right(.List(.listIndex, 1), 2)
        End With
    End If
End Sub

Private Sub lstAttendance_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Un'HookListBoxScroll
End Sub

Private Sub lstAttendance_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Me.lstAttendance.ListCount > 0 Then
        'HookListBoxScroll Me, Me.lstAttendance
    End If
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
    
    '--//기초설정
    Me.cmdClose.Cancel = True
    TB1 = "op_system.db_churchlist_custom" '--//교회리스트
    TB2 = "op_system.db_attendance" '--//출석현황
    
    '--//권한에 따른 컨트롤 설정
    Call HideDeleteButtonByUserAuth
    
    '--//컨트롤 설정
    Me.cmdAdd.Enabled = False
    Me.cmdAdd.Visible = False
    Me.cmdCancel.Visible = False
    Me.cmdNew.Enabled = False
    Me.cmdDelete.Enabled = False
    Me.cmdEdit.Enabled = False
    Me.txtOnce.Enabled = False
    Me.txtForth.Enabled = False
    Me.txtOnce_Stu.Enabled = False
    Me.txtForth_Stu.Enabled = False
    Me.txtTithe_All.Enabled = False
    Me.txtTithe_Stu.Enabled = False
    Me.txtBaptism.Enabled = False
    Me.txtEvangelist.Enabled = False
    Me.txtGL.Enabled = False
    Me.txtUL.Enabled = False
    Me.cboYear.Enabled = False
    Me.cboMonth.Enabled = False
    
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
    
    '--//리스트박스 설정
    With Me.lstChurch
        .ColumnCount = 4
        .ColumnHeads = False
        .ColumnWidths = "0,150,50,150" '교회코드, 교회명, 교회구분, 관리교회명
'        .Width = 401
        .TextAlign = fmTextAlignLeft
        .Font = "굴림"
    End With
    
    With Me.lstAttendance
        .ColumnCount = 12
        .ColumnHeads = False
        .ColumnWidths = "0,45,30,30,30,40,40,30,30,35,30,100" '교회코드, 날짜,1회(전),1회(학↑),4회(전),4회(학↑),전체반차,학생반차,침례,전도인,지역장,구역장
        .Width = 401
        .TextAlign = fmTextAlignLeft
        .Font = "굴림"
    End With
    
    Me.txtChurch.SetFocus
    Call WaitFor(0.005)
End Sub
Private Sub cmdSearch_Click()
    strSql = makeSelectSQL(TB1)
    Call makeListData(strSql, TB1)
    If cntRecord > 0 Then
        Me.lstChurch.List = LISTDATA
    Else
        Me.lstChurch.Clear
    End If
    Call sbClearVariant
End Sub
Private Sub txtChurch_Change()
    Me.txtChurch.BackColor = RGB(255, 255, 255)
End Sub
Private Sub cmdCancel_Click()
    'Call cmdbtn_visible
    Call INPUTMODE(False)
    Call sbtxtBox_Init
    Call HideDeleteButtonByUserAuth
    Call lstAttendance_Click
'    Me.cboYear.Enabled = False
'    Me.cboMonth.Enabled = False
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    
    Dim result As T_RESULT
    
    If MsgBox("선택한 데이터를 삭제하시겠습니까?", vbQuestion + vbYesNo, banner) = vbNo Then
        Exit Sub
    End If
    
    strSql = makeDeleteSQL(TB2)
    connectTaskDB
    result.strSql = strSql
    result.affectedCount = executeSQL("cmdDelete_Click", TB2, strSql, Me.Name, "출석이력 삭제")
    writeLog "cmdDelete_Click", TB2, strSql, 0, Me.Name, "출석이력 삭제"
    disconnectALL
    
    '--//메세지박스
    MsgBox "해당 데이터가 삭제되었습니다.", , banner
    
    '--//리스트박스 새로고침
    Call lstChurch_Click
    Me.lstAttendance.listIndex = -1
    
End Sub

Private Sub cmdEdit_Click()
    
    Dim result As T_RESULT
    
    '--//수정된 내용 있는지 체크
    With Me.lstAttendance
        If Me.txtOnce = .List(.listIndex, 2) And Me.txtOnce_Stu = .List(.listIndex, 3) And Me.txtForth = .List(.listIndex, 4) And _
            Me.txtForth_Stu = .List(.listIndex, 5) And Me.txtTithe_All = .List(.listIndex, 6) And Me.txtTithe_Stu = .List(.listIndex, 7) And Me.txtBaptism = .List(.listIndex, 8) And _
            Me.txtEvangelist = .List(.listIndex, 9) And Me.txtGL = .List(.listIndex, 10) And Me.txtUL = .List(.listIndex, 11) Then
            Exit Sub
        End If
    End With
    
    '--//데이터 유효성 검사
    If fnData_Validation = False Then
On Error Resume Next
        txtBox_Focus.SetFocus
        txtBox_Focus.SelStart = 0
        txtBox_Focus.SelLength = Len(txtBox_Focus)
On Error GoTo 0
        Exit Sub
    End If
    
    '--//SQL문 생성, 실행, 로그기록

    strSql = makeUpdateSQL(TB4)
    result.strSql = strSql
    connectTaskDB
    result.affectedCount = executeSQL("cmdEdit_Click", TB4, strSql, Me.Name, "출석이력 업데이트")
    writeLog "cmdEdit_Click", TB4, strSql, 0, Me.Name, "출석이력 업데이트", result.affectedCount
    disconnectALL

    '--//메세지박스
    MsgBox "업데이트 되었습니다.", , banner
    
    '--//리스트박스 초기화
    Call lstChurch_Click
    Me.lstAttendance.listIndex = queryKey
    
End Sub

Private Sub cmdNew_Click()
    'Call cmdbtn_visible
    Call INPUTMODE(True)
    Call HideDeleteButtonByUserAuth
    Me.lstAttendance.listIndex = Me.lstAttendance.ListCount - 1
    Call sbtxtBox_Init
    Me.cboYear = year(Date)
    Me.cboMonth = month(WorksheetFunction.EDate(Date, -1))
'    Me.cboYear.Enabled = True
'    Me.cboMonth.Enabled = True
End Sub

Private Sub cmdADD_Click()
    
    Dim argData As T_ATTENDANCE
    Dim result As T_RESULT
    
    '--//중복체크
    With Me.lstChurch
        strSql = "SELECT * FROM " & TB2 & " a WHERE a.church_sid = " & SText(.List(.listIndex)) & _
                " AND a.attendance_dt = " & SText(DateSerial(Me.cboYear, Me.cboMonth, 1)) & ";"
        Call makeListData(strSql, TB2)
    End With
    
    If cntRecord > 0 Then
        MsgBox "중복된 기간은 존재할 수 없습니다. 입력 값을 다시 확인해주세요.", vbCritical, banner
        queryKey = Format(LISTDATA(0, 1), "yyyy-mm")
        Call returnListPosition2(Me, Me.lstAttendance.Name, queryKey)
        Exit Sub
    End If
    Call sbClearVariant
    
    '--//데이터 유효성 검사
    If fnData_Validation = False Then
On Error Resume Next
        txtBox_Focus.SetFocus
        txtBox_Focus.SelStart = 0
        txtBox_Focus.SelLength = Len(txtBox_Focus)
On Error GoTo 0
        Exit Sub
    End If
    
    '--//작업에 따라 구조체에 값 추가
    argData.church_sid = Me.lstChurch.List(Me.lstChurch.listIndex)
    argData.ATTENDANCE_DT = DateSerial(Me.cboYear, Me.cboMonth, 1)
    argData.ONCE_ALL = IIf(Me.txtOnce = "", 0, Me.txtOnce)
    argData.ONCE_STU = IIf(Me.txtOnce_Stu = "", 0, Me.txtOnce_Stu)
    argData.FORTH_ALL = IIf(Me.txtForth = "", 0, Me.txtForth)
    argData.FORTH_STU = IIf(Me.txtForth_Stu = "", 0, Me.txtForth_Stu)
    argData.TITHE_ALL = IIf(Me.txtTithe_All = "", 0, Me.txtTithe_All)
    argData.TITHE_STU = IIf(Me.txtTithe_Stu = "", 0, Me.txtTithe_Stu)
    argData.BAPTISM_ALL = IIf(Me.txtBaptism = "", 0, Me.txtBaptism)
    argData.Evangelist = IIf(Me.txtEvangelist = "", 0, Me.txtEvangelist)
    argData.GL = IIf(Me.txtGL = "", 0, Me.txtGL)
    argData.UL = IIf(Me.txtUL = "", 0, Me.txtUL)
    
    If WorksheetFunction.Sum(argData.BAPTISM_ALL, argData.Evangelist, argData.FORTH_ALL, argData.FORTH_STU, argData.GL, argData.ONCE_ALL, argData.ONCE_STU, argData.TITHE_STU, argData.UL) = 0 Then
        MsgBox "값을 입력해 주세요.", vbCritical, banner
        Exit Sub
    End If
    
    '--//쿼리문 실행 및 로그기록
    strSql = makeInsertSQL(TB2, argData)
    connectTaskDB
    result.strSql = strSql
    result.affectedCount = executeSQL("cmdADD_Clikc", TB2, strSql, Me.Name, "출석이력 추가")
    writeLog "cmdADD_Click", TB2, strSql, 0, Me.Name, "출석이력 추가", result.affectedCount
    disconnectALL

    '--//메세지박스
    MsgBox "추가 되었습니다.", , banner
    Call lstChurch_Click
    Me.lstAttendance.listIndex = Me.lstAttendance.ListCount - 1
    
    '--//버튼설정 원래대로
    'Call cmdbtn_visible
    Call INPUTMODE(False)
    Call HideDeleteButtonByUserAuth
    Me.cboYear.Enabled = False
    Me.cboMonth.Enabled = False
    
End Sub
Private Sub lstChurch_Click()
    
    '--//교회 선택여부 판단
    If Me.lstChurch.listIndex = -1 Then
        MsgBox "교회를 선택하세요.", vbCritical, "오류"
        Exit Sub
    End If
    
    '--//출석데이터 추가
    strSql = makeSelectSQL(TB2)
    Call makeListData(strSql, TB2)
On Error Resume Next
    Me.lstAttendance.List = LISTDATA
    If err.Number <> 0 Then
        Me.lstAttendance.Clear
    End If
On Error GoTo 0
    Call sbClearVariant
    
    '--//마지막 데이터 선택
    With Me.lstAttendance
        .listIndex = .ListCount - 1
    End With
    
    '--//컨트롤 설정
    If Me.lstChurch.listIndex <> -1 Then
        Me.cmdNew.Enabled = True
    Else
        Me.cmdNew.Enabled = False
    End If
    
'    If Me.lstAttendance.listIndex <> -1 Then
'        Me.cmdADD.Enabled = True
'        Me.cmdDelete.Enabled = True
'        Me.cmdEdit.Enabled = True
'        Me.txtOnce.Enabled = True
'        Me.txtForth.Enabled = True
'        Me.txtOnce_Stu.Enabled = True
'        Me.txtForth_Stu.Enabled = True
'        Me.txtTithe_All.Enabled = True
'        Me.txtTithe_Stu.Enabled = True
'        Me.txtBaptism.Enabled = True
'        Me.txtEvangelist.Enabled = True
'        Me.txtGL.Enabled = True
'        Me.txtUL.Enabled = True
'        Me.cboYear.Enabled = True
'        Me.cboMonth.Enabled = True
'    Else
'        Call sbtxtBox_Init
'        Me.cmdADD.Enabled = False
'        Me.cmdDelete.Enabled = False
'        Me.cmdEdit.Enabled = False
'        Me.txtOnce.Enabled = False
'        Me.txtForth.Enabled = False
'        Me.txtOnce_Stu.Enabled = False
'        Me.txtForth_Stu.Enabled = False
'        Me.txtTithe_All.Enabled = False
'        Me.txtTithe_Stu.Enabled = False
'        Me.txtBaptism.Enabled = False
'        Me.txtEvangelist.Enabled = False
'        Me.txtGL.Enabled = False
'        Me.txtUL.Enabled = False
'        Me.cboYear.Enabled = False
'        Me.cboMonth.Enabled = False
'    End If
    
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
    
    Select Case tableNM
    Case TB1
        '--//교회코드, 교회명
        If Me.chkAll Then
            strSql = "SELECT a.church_sid,a.church_nm,a.church_gb,b.church_nm " & _
                        "FROM " & TB1 & " a LEFT JOIN op_system.db_churchlist_custom b ON a.main_church_cd = b.church_sid " & _
                        "WHERE a.ovs_dept = " & USER_DEPT & " AND (a.church_nm LIKE '%" & Me.txtChurch & "%' OR b.church_nm LIKE '%" & Me.txtChurch & "%') ORDER BY a.sort_order;"
        Else
            strSql = "SELECT a.church_sid,a.church_nm,a.church_gb,b.church_nm " & _
                        "FROM " & TB1 & " a LEFT JOIN op_system.db_churchlist_custom b ON a.main_church_cd = b.church_sid " & _
                        "WHERE a.ovs_dept = " & USER_DEPT & " AND a.suspend = 0 AND (a.church_nm LIKE '%" & Me.txtChurch & "%' OR b.church_nm LIKE '%" & Me.txtChurch & "%') ORDER BY a.sort_order;"
        End If
    Case TB2
        strSql = "SELECT a.church_sid,DATE_FORMAT(a.attendance_dt,'%Y-%m'),a.once_all,a.once_stu,a.forth_all,a.forth_stu,a.tithe_all,a.tithe_stu,a.baptism_all,a.evangelist,a.gl,a.ul FROM " & TB2 & " a WHERE church_sid = " & SText(Me.lstChurch.List(Me.lstChurch.listIndex)) & ";"
    Case Else
    End Select
    makeSelectSQL = strSql
End Function
Private Function makeUpdateSQL(ByVal tableNM As String) As String
    
    With Me.lstAttendance
        strSql = "UPDATE " & TB2 & " a " & _
                "SET a.once_all=" & SText(Me.txtOnce) & ",a.forth_all=" & SText(Me.txtForth) & ",a.once_stu=" & SText(Me.txtOnce_Stu) & _
                ",a.forth_stu=" & SText(Me.txtForth_Stu) & ",a.tithe_all=" & SText(Me.txtTithe_All) & " ,a.tithe_stu=" & SText(Me.txtTithe_Stu) & ",a.baptism_all=" & SText(Me.txtBaptism) & ",a.evangelist=" & SText(Me.txtEvangelist) & ",a.ul=" & SText(Me.txtUL) & ",a.gl=" & SText(Me.txtGL) & _
                " WHERE a.church_sid=" & SText(.List(.listIndex)) & " AND a.attendance_dt=" & SText(.List(.listIndex, 1) & "-01") & ";"
        queryKey = .listIndex
    End With
        
    makeUpdateSQL = strSql
End Function
Private Function makeInsertSQL(ByVal tableNM As String, _
                                argData As T_ATTENDANCE) As String
    With Me.lstAttendance
        strSql = "INSERT INTO " & TB2 & " VALUES(" & _
                    SText(argData.church_sid) & "," & _
                    SText(argData.ATTENDANCE_DT) & "," & _
                    SText(argData.ONCE_ALL) & "," & _
                    SText(argData.FORTH_ALL) & "," & _
                    SText(argData.ONCE_STU) & "," & _
                    SText(argData.FORTH_STU) & "," & _
                    SText(argData.TITHE_ALL) & "," & _
                    SText(argData.TITHE_STU) & "," & _
                    SText(argData.BAPTISM_ALL) & "," & _
                    SText(argData.Evangelist) & "," & _
                    SText(argData.GL) & "," & _
                    SText(argData.UL) & ");"
    End With
    queryKey = Me.lstAttendance.ListCount - 1
    makeInsertSQL = strSql
End Function
Private Function makeDeleteSQL(ByVal tableNM As String) As String
    With Me.lstAttendance
        strSql = "DELETE FROM " & TB2 & " WHERE church_sid = " & SText(.List(.listIndex)) & " AND attendance_dt = " & SText(.List(.listIndex, 1) & "-01") & ";"
    End With
    makeDeleteSQL = strSql
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
    If Not IsNumeric(Me.txtOnce.Value) Then
        fnData_Validation = False
        MsgBox "1회 출석(전체) 입력 값이 유효하지 않습니다. 다시 확인 해주세요.", vbCritical, banner
        Set txtBox_Focus = Me.txtOnce: Exit Function
    End If
    If Not IsNumeric(Me.txtOnce_Stu.Value) Then
        fnData_Validation = False
        MsgBox "1회 출석(학생이상) 입력 값이 유효하지 않습니다. 다시 확인 해주세요.", vbCritical, banner
        Set txtBox_Focus = Me.txtOnce_Stu: Exit Function
    End If
    If Not IsNumeric(Me.txtForth.Value) Then
        fnData_Validation = False
        MsgBox "4회 출석(전체) 입력 값이 유효하지 않습니다. 다시 확인 해주세요.", vbCritical, banner
        Set txtBox_Focus = Me.txtForth: Exit Function
    End If
    If Not IsNumeric(Me.txtForth_Stu.Value) Then
        fnData_Validation = False
        MsgBox "4회 출석(학생이상) 입력 값이 유효하지 않습니다. 다시 확인 해주세요.", vbCritical, banner
        Set txtBox_Focus = Me.txtForth_Stu: Exit Function
    End If
    If Not IsNumeric(Me.txtTithe_All.Value) Then
        fnData_Validation = False
        MsgBox "전체반차 입력 값이 유효하지 않습니다. 다시 확인 해주세요.", vbCritical, banner
        Set txtBox_Focus = Me.txtTithe_All: Exit Function
    End If
    If Not IsNumeric(Me.txtTithe_Stu.Value) Then
        fnData_Validation = False
        MsgBox "학생이상 반차 입력 값이 유효하지 않습니다. 다시 확인 해주세요.", vbCritical, banner
        Set txtBox_Focus = Me.txtTithe_Stu: Exit Function
    End If
    If Not IsNumeric(Me.txtBaptism.Value) Then
        fnData_Validation = False
        MsgBox "침례 입력 값이 유효하지 않습니다. 다시 확인 해주세요.", vbCritical, banner
        Set txtBox_Focus = Me.txtBaptism: Exit Function
    End If
    If Not IsNumeric(Me.txtEvangelist.Value) Then
        fnData_Validation = False
        MsgBox "고정전도인 입력 값이 유효하지 않습니다. 다시 확인 해주세요.", vbCritical, banner
        Set txtBox_Focus = Me.txtEvangelist: Exit Function
    End If
    If Not IsNumeric(Me.txtGL.Value) Then
        fnData_Validation = False
        MsgBox "지역장 입력 값이 유효하지 않습니다. 다시 확인 해주세요.", vbCritical, banner
        Set txtBox_Focus = Me.txtGL: Exit Function
    End If
    If Not IsNumeric(Me.txtUL.Value) Then
        fnData_Validation = False
        MsgBox "구역장 입력 값이 유효하지 않습니다. 다시 확인 해주세요.", vbCritical, banner
        Set txtBox_Focus = Me.txtUL: Exit Function
    End If
End Function
Sub sbtxtBox_Init()
    Me.txtOnce = ""
    Me.txtForth = ""
    Me.txtOnce_Stu = ""
    Me.txtForth_Stu = ""
    Me.txtTithe_All = ""
    Me.txtTithe_Stu = ""
    Me.txtBaptism = ""
    Me.txtEvangelist = ""
    Me.txtGL = ""
    Me.txtUL = ""
    Me.cboYear = ""
    Me.cboMonth = ""
End Sub
Private Sub HideDeleteButtonByUserAuth()
    Call GetUserAuthorities
    
    If cntRecord < 1 Then
        Exit Sub
    End If
    
    If IsInArray("DELETE_ITEM", LISTDATA) = -1 Then
        Me.cmdDelete.Visible = False
    End If
End Sub

Private Sub GetUserAuthorities()

    Dim sql As String
    
    sql = "SELECT b.authority FROM op_system.a_auth_table a" & _
          " LEFT JOIN op_system.a_authority b " & _
          "     ON a.authority_id = b.id" & _
          " WHERE a.user_id = " & USER_ID & ";"
    Call makeListData(sql, "op_system.a_auth_table")
    
End Sub

Private Sub cmdbtn_visible()
    Me.cmdNew.Visible = Not Me.cmdNew.Visible
    Me.cmdEdit.Visible = Not Me.cmdEdit.Visible
    Me.cmdDelete.Visible = Not Me.cmdDelete.Visible
    Me.cmdCancel.Visible = Not Me.cmdCancel.Visible
    Me.cmdAdd.Visible = Not Me.cmdAdd.Visible
End Sub
Private Sub INPUTMODE(ByVal argBoolean As Boolean)
    '--//버튼 활성화/비활성화
    Me.cmdNew.Visible = Not argBoolean
    Me.cmdEdit.Visible = Not argBoolean
    Me.cmdDelete.Visible = Not argBoolean
    Me.cmdCancel.Visible = argBoolean
    Me.cmdCancel.Enabled = argBoolean
    Me.cmdAdd.Visible = argBoolean
    Me.cmdAdd.Enabled = argBoolean
    
    '--//컨트롤 활성화/비활성화
    Me.txtChurch.Enabled = Not argBoolean
    Me.cmdSearch.Enabled = Not argBoolean
    Me.lstChurch.Enabled = Not argBoolean
    Me.lstAttendance.Enabled = Not argBoolean
    Me.chkAll.Enabled = Not argBoolean
    
    Me.txtOnce.Enabled = argBoolean
    Me.txtOnce_Stu.Enabled = argBoolean
    Me.txtForth.Enabled = argBoolean
    Me.txtForth_Stu.Enabled = argBoolean
    Me.txtTithe_All.Enabled = argBoolean
    Me.txtTithe_Stu.Enabled = argBoolean
    Me.txtBaptism.Enabled = argBoolean
    Me.txtEvangelist.Enabled = argBoolean
    Me.txtGL.Enabled = argBoolean
    Me.txtUL.Enabled = argBoolean
    Me.cboYear.Enabled = argBoolean
    Me.cboMonth.Enabled = argBoolean
    Me.cmdAdd.Enabled = argBoolean
    
End Sub
Sub WaitFor(NumOfSeconds As Single)

    Dim SngSec As Single

    SngSec = Timer + NumOfSeconds

Do While Timer < SngSec

        DoEvents

   Loop

End Sub
