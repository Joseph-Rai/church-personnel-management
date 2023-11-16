VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Update_Flight 
   Caption         =   "출입국이력 관리마법사"
   ClientHeight    =   8175
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8130
   OleObjectBlob   =   "frm_Update_Flight.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "frm_Update_Flight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim cntlst1Col As Integer '//리스트 컬럼 수
Dim TB1 As String, TB2 As String, TB3 As String, TB4 As String, TB5 As String, TB6 As String, TB7 As String '//폼에 연결된 DB 테이블
Const strOrderBy As String = "currency_un ASC" '//DB에서 sort_order 필드
Dim caseSave As Integer '//1: 추가, 2: 수정, 3: 삭제(SUSPEND), 4: 완전삭제
Dim queryKey As Integer '//리스트 위치 반환에 사용될 id
Dim LISTDATA() As String '--//DB에서 받아온 rs를 배열로 저장
Dim LISTFIELD() As String '--//DB에서 받아온 rs의 필드를 배열로 저장
Dim cntRecord As Integer '--//DB에서 받아온 레코드의 개수
Dim strSql As String '--//SQL 쿼리문
Dim txtBox_Focus As MSForms.textBox

Private Sub cmdCancel_Click()
    'Call cmdbtn_visible
    Call INPUTMODE(False)
    Call HideDeleteButtonByUserAuth
    Call lstHistory_Click
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
    result.affectedCount = executeSQL("cmdDelete_Click", TB2, strSql, Me.Name, "출입국 이력 삭제")
    writeLog "cmdDelete_Click", TB2, strSql, 0, Me.Name, "출입국 이력 삭제"
    disconnectALL
    
    '--//메세지박스
    MsgBox "해당 데이터가 삭제되었습니다.", , banner
    
    '--//리스트박스 새로고침
    Call lstPStaff_Click
    Me.lstHistory.listIndex = -1
    
End Sub

Private Sub cmdEdit_Click()
    
    Dim result As T_RESULT
    
    '--//수정된 내용 있는지 체크
    With Me.lstHistory
        If Me.txtDate = .List(.listIndex, 2) And Me.txtDeparture = .List(.listIndex, 3) And Me.txtDestination = .List(.listIndex, 4) And Me.txtPurpose = .List(.listIndex, 5) Then
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
    
'    '--//중복체크
'    With Me.lstHistory
'        strSQL = "SELECT * FROM " & TB2 & " a WHERE a.lifeno = " & SText(.List(.ListIndex, 1)) & _
'                " AND a.flight_dt = " & SText(Me.txtDate) & ";"
'        Call makeListData(strSQL, TB2)
'    End With
'
'    If cntRecord > 0 Then
'        If MsgBox("동일한 날짜에 항공스케줄이 이미 존재 합니다. 정말 진행 하시겠습니까?", vbQuestion + vbYesNo, banner) = vbNo Then
'            queryKey = listData(0, 0)
'            Call returnListPosition(Me, Me.lstHistory.Name, queryKey)
'            Exit Sub
'        End If
'    End If
    
    Call sbClearVariant
    
    '--//SQL문 생성, 실행, 로그기록
    strSql = makeUpdateSQL(TB2)
    result.strSql = strSql
    connectTaskDB
    result.affectedCount = executeSQL("cmdEdit_Click", TB2, strSql, Me.Name, "출입국 이력 업데이트")
    writeLog "cmdEdit_Click", TB2, strSql, 0, Me.Name, "출입국 이력 업데이트", result.affectedCount
    disconnectALL

    '--//메세지박스
    MsgBox "업데이트 되었습니다.", , banner
    
    '--//리스트박스 초기화
    Call lstPStaff_Click
    Me.lstHistory.listIndex = queryKey
    
End Sub

Private Sub cmdNew_Click()
    
    '--//이력 리스트 박스를 선택하지 않았으면 텍스트 박스들이 활성화 되지 않으므로 클릭처리 하기
    If lstHistory.ListCount = 0 Then
        Call lstHistory_Click
    End If
    
    'Call cmdbtn_visible
    Call INPUTMODE(True)
    Call HideDeleteButtonByUserAuth
    Me.lstHistory.listIndex = Me.lstHistory.ListCount - 1
    Call sbtxtBox_Init
End Sub

Private Sub cmdADD_Click()
    
    Dim argData As T_FLIGHT_SCHEDULE
    Dim result As T_RESULT
    
    '--//데이터 유효성 검사
    If fnData_Validation = False Then
On Error Resume Next
        txtBox_Focus.SetFocus
        txtBox_Focus.SelStart = 0
        txtBox_Focus.Setlength = Len(txtBox_Focus)
On Error GoTo 0
        Exit Sub
    End If
    
    '--//중복체크
    With Me.lstPStaff
        strSql = "SELECT * FROM " & TB2 & " a WHERE a.lifeno = " & SText(.List(.listIndex)) & _
                " AND a.flight_dt = " & SText(Me.txtDate) & ";"
        Call makeListData(strSql, TB2)
    End With
    
    If cntRecord > 0 Then
        If MsgBox("동일한 날짜에 항공스케줄이 이미 존재 합니다. 정말 진행 하시겠습니까?", vbQuestion + vbYesNo, banner) = vbNo Then
            queryKey = LISTDATA(0, 0)
            Call returnListPosition(Me, Me.lstHistory.Name, queryKey)
            Exit Sub
        End If
    End If
    Call sbClearVariant
    
    '--//작업에 따라 구조체에 값 추가
    With Me.lstPStaff
        argData.lifeNo = .List(.listIndex)
        argData.FLIGHT_DT = Me.txtDate
        argData.DEPARTURE = Replace(Me.txtDeparture, "한국", "대한민국")
        argData.Destination = Replace(Me.txtDestination, "한국", "대한민국")
        argData.VISIT_PURPOSE = Me.txtPurpose
    End With

    
    '--//작업에 따라 쿼리문 실행 및 로그기록
    strSql = makeInsertSQL(TB2, argData)
    connectTaskDB
    result.strSql = strSql
    result.affectedCount = executeSQL("cmdADD_Clikc", TB2, strSql, Me.Name, "출입국 이력 추가")
    writeLog "cmdADD_Click", TB2, strSql, 0, Me.Name, "출입국 이력 추가", result.affectedCount
    disconnectALL
    
    '--//메세지박스
    MsgBox "추가 되었습니다.", , banner
    Call lstPStaff_Click
    Me.lstHistory.listIndex = Me.lstHistory.ListCount - 1
    
    '--//버튼설정 원래대로
    'Call cmdbtn_visible
    Call INPUTMODE(False)
    Call HideDeleteButtonByUserAuth
    
End Sub

Private Sub lstHistory_Click()
    
    '--//컨트롤 설정
    If Me.lstHistory.listIndex <> -1 Then
        Me.txtDate.Enabled = True
        Me.txtDeparture.Enabled = True
        Me.txtDestination.Enabled = True
        Me.txtPurpose.Enabled = True
        Me.cmdEdit.Enabled = True
        Me.cmdAdd.Enabled = True
        Me.cmdDelete.Enabled = True
    Else
        Me.txtDate.Enabled = False
        Me.txtDeparture.Enabled = False
        Me.txtDestination.Enabled = False
        Me.txtPurpose.Enabled = False
        Me.cmdEdit.Enabled = False
        Me.cmdAdd.Enabled = False
        Me.cmdDelete.Enabled = False
    End If
    
    '--//리스트 클릭 시 시작일, 종료일, 내용 표시
    If Me.lstHistory.listIndex <> -1 Then
        With Me.lstHistory
            Me.txtDate = .List(.listIndex, 2)
            Me.txtDeparture = .List(.listIndex, 3)
            Me.txtDestination = .List(.listIndex, 4)
            Me.txtPurpose = .List(.listIndex, 5)
        End With
    End If
    
End Sub

Private Sub lstHistory_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Un'HookListBoxScroll
End Sub

Private Sub lstHistory_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Me.lstHistory.ListCount > 0 Then
        'HookListBoxScroll Me, Me.lstHistory
    End If
End Sub

Private Sub lstPStaff_Click()
    
    Dim filePath As String
    Dim FileName As String
    
    If Me.lstPStaff.listIndex <> -1 Then
        Me.lstHistory.Enabled = True
        Me.cmdNew.Enabled = True
    Else
        Me.lstHistory.Enabled = False
        Me.cmdNew.Enabled = False
    End If
    
    '--//이력 목록상자 조정
    With Me.lstHistory
        .ColumnCount = 6
        .ColumnHeads = False
        .ColumnWidths = "0,0,60,70,70,180" '출입국코드, 생명번호, 날짜, 출발지, 목적지, 출입국 목적
        .Width = 380
        .TextAlign = fmTextAlignLeft
        .Font = "굴림"
    End With
    
    '--//출입국 이력 리스트박스 리스팅
    Call makeSelectSQL(TB2)
    Call makeListData(strSql, TB2)
    If cntRecord > 0 Then
        Me.lstHistory.List = LISTDATA
    Else
        Me.lstHistory.Clear
    End If
    Call sbClearVariant
    
    '--//텍스트박스 초기화
    Call sbtxtBox_Init
    Me.txtDate = ""
    
    '--//사진추가
    Dim strLifeNo As String
    With Me.lstPStaff
        strLifeNo = .List(.listIndex)
    End With
    InsertPicToLabel Me.lblPic, strLifeNo
    
End Sub

Private Sub lstPStaff_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Un'HookListBoxScroll
End Sub

Private Sub lstPStaff_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Me.lstPStaff.ListCount > 0 Then
        'HookListBoxScroll Me, Me.lstPStaff
    End If
End Sub

Private Sub txtChurchNM_Change()
    Me.txtChurchNM.BackColor = RGB(255, 255, 255)
End Sub

Private Sub txtDate_Change()
    Call Date_Format(Me.txtDate)
End Sub

Private Sub txtDeparture_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    argShow = 1
    frm_Search_Country.Show
End Sub

Private Sub txtDeparture_Enter()
    If Me.txtDeparture = "" Then
        argShow = 1
        frm_Search_Country.Show
    End If
End Sub

Private Sub txtDestination_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    argShow = 2
    frm_Search_Country.Show
End Sub

Private Sub txtDestination_Enter()
    If Me.txtDestination = "" Then
        argShow = 2
        frm_Search_Country.Show
    End If
End Sub

Private Sub UserForm_Initialize()
    If connIP = Empty Then setGlobalVariant '//전역변수 재설정
    If checkLogin = 0 Then f_login.Show '//로그인체크
    If checkLogin = 0 Then End '--//로그인 실패 시 프로시저 종료
    
    '--//기초설정
    Me.cmdClose.Cancel = True
    TB1 = "op_system.v0_pstaff_information_all" '--//선지자 리스트(전체)
    TB2 = "op_system.db_flight_schedule" '--//출입국 이력 테이블
    TB3 = "op_system.v0_pstaff_information" '--//선지자 리스트
    
    '--//권한에 따른 컨트롤 설정
    Call HideDeleteButtonByUserAuth
    
    '--//컨트롤 설정
    Me.lstPStaff.Enabled = False
    Me.lstHistory.Enabled = False
    Me.txtDate.Enabled = False
    Me.txtDeparture.Enabled = False
    Me.txtDestination.Enabled = False
    Me.txtPurpose.Enabled = False
    Me.cmdEdit.Enabled = False
    Me.cmdAdd.Enabled = False
    Me.cmdAdd.Visible = False
    Me.cmdCancel.Visible = False
    Me.cmdNew.Enabled = False
    Me.cmdDelete.Enabled = False
    
    '--//리스트박스 설정
    With Me.lstPStaff
        .ColumnCount = 4
        .ColumnHeads = False
        .ColumnWidths = "0,120,70,50" '생명번호, 교회명, 한글이름(직분), 직책
        .TextAlign = fmTextAlignLeft
        .Font = "굴림"
    End With
    
End Sub
Private Sub cmdSearch_Click()
    
    If Me.chkAll.Value Then
        Me.lstHistory.Clear '--//기존의 출입국 이력 리스트박스 초기화
        
        strSql = makeSelectSQL(TB1)
        Call makeListData(strSql, TB1)
        
        '--//반환할 데이터가 없으면 메세지 박스 후 종료
        If cntRecord = 0 Then
            MsgBox "반환할 DB 데이터가 없습니다.", vbInformation, banner
            Call sbClearVariant
            Exit Sub
        End If
        
        '--//선지자 리스트에 검색한 목록 띄우기
        Me.lstPStaff.List = LISTDATA
        Call sbClearVariant '--//변수 초기화
        Me.lstPStaff.Enabled = True
    Else
        Me.lstHistory.Clear '--//기존의 출입국 이력 리스트박스 초기화
        
        strSql = makeSelectSQL(TB3)
        Call makeListData(strSql, TB3)
        
        '--//반환할 데이터가 없으면 메세지 박스 후 종료
        If cntRecord = 0 Then
            MsgBox "반환할 DB 데이터가 없습니다.", vbInformation, banner
            Call sbClearVariant
            Exit Sub
        End If
        
        '--//선지자 리스트에 검색한 목록 띄우기
        Me.lstPStaff.List = LISTDATA
        Call sbClearVariant '--//변수 초기화
        Me.lstPStaff.Enabled = True
    End If
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
        '--//교회코드, 교회명
        strSql = "SELECT a.`생명번호`,a.`교회명`,a.`한글이름(직분)`,a.`직책` " & _
                    "FROM " & TB1 & " a " & _
                    "WHERE (a.`한글이름(직분)` LIKE '%" & Me.txtChurchNM & "%' OR a.`교회명` LIKE '%" & Me.txtChurchNM & "%'" & _
                    " OR a.`영문이름` LIKE '%" & Me.txtChurchNM & "%'" & _
                    " OR a.`생명번호` LIKE '%" & Me.txtChurchNM & "%')" & _
                    " AND a.`관리부서` = " & SText(USER_DEPT) & _
                 " UNION " & _
                 "SELECT b.`배우자생번`,b.`교회명`,b.`사모한글이름(직분)`,b.`사모직책` " & _
                    "FROM " & TB1 & " b " & _
                    "WHERE (b.`사모한글이름(직분)` LIKE '%" & Me.txtChurchNM & "%' OR b.`교회명` LIKE '%" & Me.txtChurchNM & "%'" & _
                    " OR b.`사모영문이름` LIKE '%" & Me.txtChurchNM & "%'" & _
                    " OR b.`배우자생번` LIKE '%" & Me.txtChurchNM & "%')" & _
                    " AND b.`관리부서` = " & SText(USER_DEPT) & ";"
    Case TB2
        strSql = "SELECT * FROM " & TB2 & " a " & _
                "WHERE a.lifeno = " & SText(Me.lstPStaff.List(Me.lstPStaff.listIndex)) & _
                "ORDER BY a.flight_dt;"
    Case TB3
        '--//교회코드, 교회명
        strSql = "SELECT a.`생명번호`,a.`교회명`,a.`한글이름(직분)`,a.`직책` " & _
                    "FROM " & TB3 & " a " & _
                    "WHERE (a.`한글이름(직분)` LIKE '%" & Me.txtChurchNM & "%' OR a.`교회명` LIKE '%" & Me.txtChurchNM & "%'" & _
                    " OR a.`영문이름` LIKE '%" & Me.txtChurchNM & "%'" & _
                    " OR a.`생명번호` LIKE '%" & Me.txtChurchNM & "%')" & _
                    " AND a.`관리부서` = " & SText(USER_DEPT) & _
                 " UNION " & _
                 "SELECT b.`배우자생번`,b.`교회명`,b.`사모한글이름(직분)`,b.`사모직책` " & _
                    "FROM " & TB3 & " b " & _
                    "WHERE (b.`사모한글이름(직분)` LIKE '%" & Me.txtChurchNM & "%' OR b.`교회명` LIKE '%" & Me.txtChurchNM & "%'" & _
                    " OR b.`사모영문이름` LIKE '%" & Me.txtChurchNM & "%'" & _
                    " OR b.`배우자생번` LIKE '%" & Me.txtChurchNM & "%')" & _
                    " AND b.`관리부서` = " & SText(USER_DEPT) & ";"
    Case Else
    End Select
    makeSelectSQL = strSql
End Function
Private Function makeUpdateSQL(ByVal tableNM As String) As String
    
    Select Case tableNM
    Case TB1
    Case TB2
        With Me.lstHistory
            strSql = "UPDATE " & TB2 & " a " & _
                    "SET a.flight_dt = " & SText(Me.txtDate) & ", a.departure = " & SText(Me.txtDeparture) & ",a.destination = " & SText(Me.txtDestination) & ",a.visit_purpose = " & SText(Me.txtPurpose) & _
                    " WHERE a.flight_cd = " & SText(.List(.listIndex)) & ";"
        queryKey = .listIndex
        End With
    Case Else
    End Select
    makeUpdateSQL = strSql
End Function
Private Function makeInsertSQL(ByVal tableNM As String, argData As T_FLIGHT_SCHEDULE) As String
    
    Select Case tableNM
    Case TB1
    Case TB2
        strSql = "INSERT INTO " & TB2 & " VALUES(DEFAULT," & _
                    SText(argData.lifeNo) & "," & _
                    SText(argData.FLIGHT_DT) & "," & _
                    SText(argData.DEPARTURE) & "," & _
                    SText(argData.Destination) & "," & _
                    SText(argData.VISIT_PURPOSE) & ");"
        queryKey = Me.lstHistory.ListCount - 1
    Case Else
    End Select
    makeInsertSQL = strSql
End Function
Private Function makeDeleteSQL(ByVal tableNM As String) As String
    
    Select Case tableNM
    Case TB1
    Case TB2
        With Me.lstHistory
            strSql = "DELETE FROM " & TB2 & " WHERE flight_cd = " & SText(.List(.listIndex)) & ";"
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
Sub sbtxtBox_Init()
    Me.txtDate.Value = Date
    Me.txtDeparture.Value = ""
    Me.txtDestination.Value = ""
    Me.txtPurpose.Value = ""
End Sub
Private Function fnData_Validation()
'---------------------------------------
'유저폼 입력값에 대한 데이터 유효성 검사
'TRUE: 이상없음, FALSE: 잘못됨.
'---------------------------------------
    fnData_Validation = True '데이터가 유효하다는 가정 하에 시작
    
    strSql = "SELECT ctry_nm FROM op_system.db_country"
    Call makeListData(strSql, "op_system.db_country")
    
    If IsInArray(Me.txtDeparture, LISTDATA) = -1 Then
        MsgBox "출발지를 잘못 입력하였습니다. 다시 확인 해주세요.", vbCritical, banner
        Set txtBox_Focus = Me.txtDeparture: fnData_Validation = False: Exit Function
    End If
    
    If IsInArray(Me.txtDestination, LISTDATA) = -1 Then
        MsgBox "목적지를 잘못 입력하였습니다. 다시 확인 해주세요.", vbCritical, banner
        Set txtBox_Focus = Me.txtDestination: fnData_Validation = False: Exit Function
    End If
    
    If Not IsDate(Me.txtDate) Then
        MsgBox "올바른 날짜 형태가 아닙니다. 날짜를 다시 확인 해주세요.", vbCritical, banner
        Set txtBox_Focus = Me.txtDate: fnData_Validation = False: Exit Function
    End If
    
    If Me.txtPurpose = "" Then
        MsgBox "출입국 목적을 기록해 주세요.", vbCritical, banner
        Set txtBox_Focus = Me.txtPurpose: fnData_Validation = False: Exit Function
    End If
    
End Function
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
    Me.txtChurchNM.Enabled = Not argBoolean
    Me.cmdSearch.Enabled = Not argBoolean
    Me.lstPStaff.Enabled = Not argBoolean
    Me.lstHistory.Enabled = Not argBoolean
    Me.chkAll.Enabled = Not argBoolean
    
    Me.txtDate.Enabled = argBoolean
    Me.txtDeparture.Enabled = argBoolean
    Me.txtDestination.Enabled = argBoolean
    Me.txtPurpose.Enabled = argBoolean
    
    
End Sub
Private Sub Date_Format(textBox As MSForms.textBox)
    Dim strDate As String
    
    If Len(Replace(textBox, "-", "")) <= 3 Then
        strDate = Replace(textBox, "-", "")
        strDate = strDate
    End If
    
    If Len(Replace(textBox, "-", "")) >= 4 And Len(Replace(textBox, "-", "")) <= 6 Then
        strDate = Replace(textBox, "-", "")
        strDate = Left(strDate, 4) & "-" & Mid(strDate, 5, Len(strDate))
    End If
    
    If Len(Replace(textBox, "-", "")) > 6 Then
        strDate = Replace(textBox, "-", "")
        strDate = Left(strDate, 4) & "-" & Mid(strDate, 5, 2) & "-" & Mid(strDate, 7, 2)
    End If
    
    textBox = strDate
End Sub

