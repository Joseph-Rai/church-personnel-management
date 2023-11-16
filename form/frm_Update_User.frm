VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Update_User 
   Caption         =   "사용자 관리 마법사"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5325
   OleObjectBlob   =   "frm_Update_User.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "frm_Update_User"
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
Dim txtBox_Focus As MSForms.control

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdEdit_Click()
    
    Dim result As T_RESULT
    Dim argData As T_USERS
    
    '--//수정된 내용 있는지 체크
    With Me.lstUser
        If Me.txtUsername = .List(.listIndex, 1) And Me.cboDepartment = .List(.listIndex, 6) Then
            Exit Sub
        End If
    
        If Me.cboDepartment <> .List(.listIndex, 6) Then
            Call GetUserAuthorities
            If IsInArray("DEPT_NUM_CHANGE", LISTDATA) = -1 Then
                MsgBox "부서 수정 권한이 없습니다.", vbCritical, "권한오류"
                Me.cboDepartment.text = .List(.listIndex, 7)
                Exit Sub
            End If
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
    Dim listIndex As Integer
    With Me.lstUser
        strSql = "SELECT * FROM " & TB1 & " a WHERE a.user_id = " & SText(.List(.listIndex)) & ";"
        Call makeListData(strSql, TB1)
    
        If cntRecord > 0 Then
            strSql = makeUpdateSQL(TB1)
        End If
    End With
    
    result.strSql = strSql
    connectTaskDB
    result.affectedCount = executeSQL("cmdEdit_Click", TB1, strSql, Me.Name, "사용자 업데이트")
    writeLog "cmdEdit_Click", TB1, strSql, 0, Me.Name, "사용자 업데이트", result.affectedCount
    disconnectALL
    
    Call sbClearVariant
    
    '--//메세지박스
    MsgBox "업데이트 되었습니다.", , banner
    
    '--//리스트박스 초기화
    Call cmdSearch_Click
    Call lstUser_Click
    Call setGlobalVariant
    
End Sub

Private Sub lstUser_Click()
    
    Me.cmdEdit.Enabled = True
    Me.txtUsername.Enabled = True
    Me.cboDepartment.Enabled = True
    
    '--//텍스트박스 초기화
    Me.txtUsername = ""
    Me.cboDepartment = ""
    
    '--//컨트롤 세팅
    If Me.lstUser.listIndex >= 0 Then
        Me.cmdDelete.Enabled = True
    Else
        Me.cmdDelete.Enabled = False
    End If
    
    '--//텍스트박스 내용추가
    Dim i As Integer
    With Me.lstUser
        If .listIndex < 0 Then
            .listIndex = .ListCount - 1
        End If
        Me.txtUsername = .List(.listIndex, 1)
        For i = 0 To Me.cboDepartment.ListCount - 1
            If Me.cboDepartment.List(i, 1) = .List(.listIndex, 7) Then
                Me.cboDepartment.listIndex = i
            End If
        Next
    End With
    
End Sub

Private Sub lstUser_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Un'HookListBoxScroll
End Sub

Private Sub lstUser_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    'HookListBoxScroll Me, Me.lstUser
End Sub

Private Sub txtSearchName_Change()
    Me.txtSearchName.BackColor = RGB(255, 255, 255)
End Sub

Private Sub UserForm_Initialize()
    If connIP = Empty Then setGlobalVariant '//전역변수 재설정
    If checkLogin = 0 Then f_login.Show '//로그인체크
    If checkLogin = 0 Then End '--//로그인 실패 시 프로시저 종료
    
    '--//기초설정
    Me.cmdClose.Cancel = True
    TB1 = "common.users" '--//사용자 정보
    TB2 = "op_system.db_ovs_dept" '--//총회부서
    
    '--//권한에 따른 컨트롤 설정
    Call HideDeleteButtonByUserAuth
    
    '--//컨트롤 설정
    Me.txtUsername.Enabled = False
    Me.cboDepartment.Enabled = False
    Me.cmdEdit.Enabled = False
    Me.cmdAdd.Enabled = False
    Me.cmdAdd.Visible = False
    Me.cmdCancel.Visible = False
    Me.cmdDelete.Enabled = False
    Me.cmdNew.Enabled = False
    
    '--//권한에 따른 설정
    Call GetUserAuthorities
    If IsInArray("USER_EDIT", LISTDATA) <> -1 Then
        Me.cmdNew.Enabled = True
    End If
    
    '--//리스트박스 설정
    With Me.lstUser
        .ColumnCount = 8
        .ColumnHeads = False
        .ColumnWidths = "0,120,0,0,0,0,0,250" '유저id, 유저명, 유저구분, 비밀번호, 초기화, IP주소, 부서id, 부서명
        .TextAlign = fmTextAlignLeft
        .Font = "굴림"
    End With
    
    '--//콤보박스 초기화
    With Me.cboDepartment
        .ColumnCount = 2
        .ColumnHeads = False
        .ColumnWidths = "0, 100" '부서id, 부서명
    End With
    
    strSql = makeSelectSQL(TB2)
    Call makeListData(strSql, TB2)
    
    Me.cboDepartment.List = LISTDATA
    
    Call cmdSearch_Click
    Me.txtSearchName.SetFocus
    
End Sub
Private Sub cmdSearch_Click()
    
    strSql = makeSelectSQL(TB1)
    Call makeListData(strSql, TB1)
    
    If cntRecord = 0 Then
        MsgBox "반환할 DB 데이터가 없습니다.", vbInformation, banner
        Call sbClearVariant
        Exit Sub
    End If
    
    Me.lstUser.List = LISTDATA
    Call sbClearVariant
    Me.lstUser.Enabled = True
    
End Sub
Private Sub cmdCancel_Click()
    Call INPUTMODE(False)
    Call HideDeleteButtonByUserAuth
    Call sbtxtBox_Init
    Call lstUser_Click
End Sub

Private Sub cmdDelete_Click()
    
    Dim result As T_RESULT
    
    If MsgBox("선택한 데이터를 삭제하시겠습니까?", vbQuestion + vbYesNo, banner) = vbNo Then
        Exit Sub
    End If
    
    strSql = makeDeleteSQL(TB1)
    connectTaskDB
    result.strSql = strSql
    result.affectedCount = executeSQL("cmdDelete_Click", TB1, strSql, Me.Name, "사용자 삭제")
    writeLog "cmdDelete_Click", TB1, strSql, 0, Me.Name, "사용자 삭제"
    disconnectALL
    
    '--//메세지박스
    MsgBox "해당 데이터가 삭제되었습니다.", , banner
    
    '--//리스트박스 새로고침
    Call cmdSearch_Click
    Call lstUser_Click
    Me.lstUser.listIndex = -1
    
End Sub

Private Sub cmdNew_Click()
    Call cmdSearch_Click
    Call lstUser_Click
    Call INPUTMODE(True)
    Call HideDeleteButtonByUserAuth
    Me.lstUser.listIndex = Me.lstUser.ListCount - 1
    Call sbtxtBox_Init
End Sub

Private Sub cmdADD_Click()
    
    Dim argData As T_USERS
    Dim result As T_RESULT
    
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
    With Me.lstUser
        argData.USER_NM = Me.txtUsername
        argData.USER_DEPT = Me.cboDepartment.List(Me.cboDepartment.listIndex, 0)
    End With
    
    '--//작업에 따라 쿼리문 실행 및 로그기록
    strSql = makeInsertSQL(TB1, argData)
    connectTaskDB
    result.strSql = strSql
    result.affectedCount = executeSQL("cmdADD_Clikc", TB1, strSql, Me.Name, "사용자 추가")
    writeLog "cmdADD_Click", TB1, strSql, 0, Me.Name, "사용자 추가", result.affectedCount
    disconnectALL
    
    '--//메세지박스
    MsgBox "추가 되었습니다.", , banner
    Call lstUser_Click
    Me.lstUser.listIndex = Me.lstUser.ListCount - 1
    
    '--//버튼설정 원래대로
    Call INPUTMODE(False)
    Call HideDeleteButtonByUserAuth
    Me.cmdCancel.Visible = False
    
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
    
    cntRecord = rs.RecordCount '--//레코드 수 검토
    
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
        '--//사용자 목록 불러오기
        Call GetUserAuthorities
                 
        If cntRecord > 0 And IsInArray("USER_EDIT", LISTDATA) <> -1 Then
            strSql = "SELECT u.user_id, u.user_nm, u.user_gb, u.user_pw, u.pw_initialize, u.user_ip, u.user_dept, d.dept_nm" & _
                      " FROM " & TB1 & " u" & _
                      " LEFT JOIN " & TB2 & " d" & _
                      "    ON u.user_dept = d.dept_id" & _
                      " WHERE u.user_nm LIKE '%" & Me.txtSearchName & "%';"
        ElseIf IsInArray("SECTION_CHIEF", LISTDATA) <> -1 Then
            strSql = "SELECT u.user_id, u.user_nm, u.user_gb, u.user_pw, u.pw_initialize, u.user_ip, u.user_dept, d.dept_nm" & _
                      " FROM " & TB1 & " u" & _
                      " LEFT JOIN " & TB2 & " d" & _
                      "    ON u.user_dept = d.dept_id" & _
                      " WHERE u.user_nm LIKE '%" & Me.txtSearchName & "%'" & _
                      "     AND u.user_dept = " & USER_DEPT & ";"
        Else
            strSql = "SELECT u.user_id, u.user_nm, u.user_gb, u.user_pw, u.pw_initialize, u.user_ip, u.user_dept, d.dept_nm" & _
                      " FROM " & TB1 & " u" & _
                      " LEFT JOIN " & TB2 & " d" & _
                      "    ON u.user_dept = d.dept_id" & _
                      " WHERE u.user_nm LIKE '%" & USER_NM & "%';"
        End If
    Case TB2
        '--//부서 목록 불러오기
        strSql = "SELECT d.dept_id, d.dept_nm FROM " & TB2 & " d WHERE d.dept_lv1 = '해외선교국';"
    Case Else
    End Select
    makeSelectSQL = strSql
End Function

Private Function makeUpdateSQL(ByVal tableNM As String) As String
    
    Select Case tableNM
    Case TB1
        strSql = "UPDATE " & TB1 & " a" & _
                 " SET a.user_nm = " & SText(Me.txtUsername) & ", a.user_dept = " & Me.cboDepartment.List(Me.cboDepartment.listIndex, 0) & _
                 " WHERE a.user_id = " & Me.lstUser.List(Me.lstUser.listIndex, 0) & ";"
    Case Else
    End Select
    makeUpdateSQL = strSql
End Function

Private Function makeInsertSQL(ByVal tableNM As String, argData As T_USERS) As String
    
    Select Case tableNM
    Case TB1
        strSql = "INSERT INTO " & TB1 & " (user_nm, user_dept) VALUES(" & _
                 SText(argData.USER_NM) & ", " & _
                 SText(argData.USER_DEPT) & ");"
    Case Else
    End Select
    makeInsertSQL = strSql
End Function

Private Function makeDeleteSQL(ByVal tableNM As String) As String
    
    Select Case tableNM
    Case TB1
        With Me.lstUser
            strSql = "DELETE a.* FROM " & TB1 & " a WHERE a.user_id = " & SText(.List(.listIndex)) & ";"
        End With
    Case Else
    End Select
    makeDeleteSQL = strSql
End Function

Private Sub sbClearVariant()
    Erase LISTFIELD
    Erase LISTDATA
    cntRecord = Empty
    strSql = vbNullString
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

Sub sbtxtBox_Init()
    Me.txtUsername = ""
    Me.cboDepartment.listIndex = -1
End Sub

Private Sub INPUTMODE(ByVal argBoolean As Boolean)
    Call sbtxtBox_Init
    Me.cmdNew.Visible = Not argBoolean
    Me.cmdEdit.Visible = Not argBoolean
    Me.cmdDelete.Visible = Not argBoolean
    Me.cmdClose.Visible = Not argBoolean
    Me.cmdCancel.Visible = argBoolean
    Me.cmdCancel.Enabled = argBoolean
    Me.cmdAdd.Visible = argBoolean
    Me.cmdAdd.Enabled = argBoolean
    
    Me.lstUser.Enabled = Not argBoolean
    Me.txtSearchName.Enabled = Not argBoolean
    Me.cmdSearch.Enabled = Not argBoolean
    
    Me.txtUsername.Enabled = argBoolean
    Me.cboDepartment.Enabled = argBoolean
    Me.cmdAdd.Enabled = argBoolean
End Sub

Private Function fnData_Validation()
'---------------------------------------
'유저폼 입력값에 대한 데이터 유효성 검사
'TRUE: 이상없음, FALSE: 잘못됨.
'---------------------------------------
    fnData_Validation = True '데이터가 유효하다는 가정 하에 시작
    
    If IsInArray(Me.cboDepartment.Value, Me.cboDepartment.List) = -1 Then
        MsgBox "부서 선택이 잘못 되었습니다.. 다시 확인 해주세요.", vbCritical, banner
        Set txtBox_Focus = Me.cboDepartment: fnData_Validation = False: Exit Function
    End If
    
End Function

Private Sub GetUserAuthorities()

    Dim sql As String
    
    sql = "SELECT b.authority FROM op_system.a_auth_table a" & _
          " LEFT JOIN op_system.a_authority b " & _
          "     ON a.authority_id = b.id" & _
          " WHERE a.user_id = " & USER_ID & ";"
    Call makeListData(sql, "op_system.a_auth_table")
    
End Sub


