VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Update_Counsel 
   Caption         =   "상담 관리대장"
   ClientHeight    =   9825.001
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14670
   OleObjectBlob   =   "frm_Update_Counsel.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "frm_Update_Counsel"
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
Dim txtBox_Focus As MSForms.control

Private Sub cboSearchArg_Category_Change()
    If lstPStaff.listIndex <> -1 Then
        Call lstPStaff_Click
    End If
End Sub

Private Sub cboSearchArg_Duration_Change()
    If lstPStaff.listIndex <> -1 Then
        Call lstPStaff_Click
    End If
End Sub

Private Sub cboSearchArg_Status_Change()
    If lstPStaff.listIndex <> -1 Then
        Call lstPStaff_Click
    End If
End Sub

Private Sub cmdCancel_Click()
    'Call cmdbtn_visible
    Call sbtxtBox_Init
    Call INPUTMODE(False)
    Call HideDeleteButtonByUserAuth
    Call lstCounsel_Click
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    
    Dim result As T_RESULT
    
    '--//리스트박스를 선택하지 않았으면 프로시저 종료
    If Me.lstCounsel.listIndex = -1 Then
        MsgBox "삭제할 데이터를 선택해 주세요.", vbCritical, banner
        Exit Sub
    End If
    
    '--//삭제여부 재확인
    If MsgBox("선택한 데이터를 삭제하시겠습니까?", vbQuestion + vbYesNo, banner) = vbNo Then
        Exit Sub
    End If
    
    strSql = makeDeleteSQL(TB2)
    connectTaskDB
    result.strSql = strSql
    result.affectedCount = executeSQL("cmdDelete_Click", TB2, strSql, Me.Name, "상담이력 삭제")
    writeLog "cmdDelete_Click", TB2, strSql, 0, Me.Name, "상담이력 삭제"
    disconnectALL
    
    '--//메세지박스
    MsgBox "해당 데이터가 삭제되었습니다.", , banner
    
    '--//리스트박스 새로고침
    Call lstPStaff_Click
    Me.lstCounsel.listIndex = -1
    Call lstCounsel_Click
    
End Sub

Private Sub cmdEdit_Click()
    
    Dim result As T_RESULT
    
    '--//리스트박스가 선택되어 있지 않으면 프로시저 종료
    If Me.lstCounsel.listIndex = -1 Then
        MsgBox "수정할 데이터를 선택해 주세요.", vbCritical, banner
        Exit Sub
    End If
    
    '--//수정된 내용 있는지 체크
    With Me.lstCounsel
        If Me.txtInputDate = .List(.listIndex, 2) And Me.cboCategory = .List(.listIndex, 3) And Me.txtTitle = .List(.listIndex, 4) And Me.txtContent = .List(.listIndex, 5) And _
            Me.txtResult = .List(.listIndex, 6) And Me.txtRemark = .List(.listIndex, 7) And Me.cboStatus = .List(.listIndex, 8) Then
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
    strSql = makeUpdateSQL(TB2)
    result.strSql = strSql
    connectTaskDB
    result.affectedCount = executeSQL("cmdEdit_Click", TB2, strSql, Me.Name, "상담이력 업데이트")
    writeLog "cmdEdit_Click", TB2, strSql, 0, Me.Name, "상담이력 업데이트", result.affectedCount
    disconnectALL
    
    '--//메세지박스
    MsgBox "업데이트 되었습니다.", , banner
    
    '--//리스트박스 초기화
    Call lstPStaff_Click
    If Me.lstCounsel.ListCount > 0 Then
        Me.lstCounsel.listIndex = queryKey
    End If
    
End Sub

Private Sub cmdNew_Click()
    
    '--//커맨드 버튼 활성화를 위한 이력 리스트박스 클릭
    If lstCounsel.ListCount = 0 Then
        Call lstCounsel_Click
    End If
    
    'Call cmdbtn_visible
    Me.lstCounsel.listIndex = Me.lstCounsel.ListCount - 1
    Call sbtxtBox_Init
    Call INPUTMODE(True)
    Call HideDeleteButtonByUserAuth
End Sub

Private Sub cmdADD_Click()
    
    Dim argData As T_COUNSEL
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
    
    '--//구조체에 값 추가
    argData.LIFE_NO = Me.lstPStaff.List(Me.lstPStaff.listIndex, 6)
    argData.COUNSEL_DT = Me.txtInputDate
    argData.CATEGORY = Me.cboCategory
    argData.title = Me.txtTitle
    argData.CONTENT = Me.txtContent
    argData.result = Me.txtResult
    argData.REMARK = Me.txtRemark
    argData.STATUS = Me.cboStatus
    
    '--//쿼리문 실행 및 로그기록
    strSql = makeInsertSQL(TB2, argData)
    connectTaskDB
    result.strSql = strSql
    result.affectedCount = executeSQL("cmdADD_Clikc", TB2, strSql, Me.Name, "상담이력 추가")
    writeLog "cmdADD_Click", TB2, strSql, 0, Me.Name, "상담이력 추가", result.affectedCount
    disconnectALL
    
    '--//버튼설정 원래대로
    Call INPUTMODE(False)
    Call HideDeleteButtonByUserAuth
    
    '--//메세지박스
    MsgBox "추가 되었습니다.", , banner
    Call lstPStaff_Click
    Me.lstCounsel.listIndex = Me.lstCounsel.ListCount - 1
    
End Sub
Private Sub lstCounsel_Click()
    
    '--//컨트롤 설정
    Call controlSettingByClickListCouncel
    
    '--//리스트 클릭 시 시작일, 종료일, 내용 표시
    If Me.lstCounsel.listIndex <> -1 Then
        With Me.lstCounsel
            Me.txtTitle = .List(.listIndex, 4)
            Me.txtContent = .List(.listIndex, 5)
            Me.txtResult = .List(.listIndex, 6)
            Me.txtRemark = .List(.listIndex, 7)
            Me.txtInputDate = .List(.listIndex, 2)
            Me.cboCategory = .List(.listIndex, 3)
            Me.cboStatus = .List(.listIndex, 8)
        End With
    End If
    
End Sub
Private Sub controlSettingByClickListCouncel()

    If Me.lstCounsel.listIndex <> -1 Then
        Me.txtInputDate.Enabled = True
        Me.cmdEdit.Enabled = True
        Me.cmdAdd.Enabled = True
        Me.cmdDelete.Enabled = True
        Me.txtTitle.Enabled = True
        Me.txtContent.Enabled = True
        Me.txtResult.Enabled = True
        Me.txtRemark.Enabled = True
        Me.txtInputDate.Enabled = True
        Me.cboCategory.Enabled = True
        Me.cboStatus.Enabled = True
    Else
        Me.txtInputDate.Enabled = False
        Me.cmdEdit.Enabled = False
        Me.cmdAdd.Enabled = False
        Me.cmdDelete.Enabled = False
        Me.txtTitle.Enabled = False
        Me.txtContent.Enabled = False
        Me.txtResult.Enabled = False
        Me.txtRemark.Enabled = False
        Me.txtInputDate.Enabled = False
        Me.cboCategory.Enabled = False
        Me.cboStatus.Enabled = False
    End If

End Sub

Private Sub lstCounsel_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Un'HookListBoxScroll
End Sub

Private Sub lstCounsel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Me.lstCounsel.ListCount > 0 Then
        'HookListBoxScroll Me, Me.lstCounsel
    End If
End Sub

Private Sub lstPStaff_Click()
    
    Dim filePath As String
    Dim FileName As String
    
    '--//컨트롤 설정
    If Me.lstPStaff.listIndex <> -1 Then
        Me.lstCounsel.Enabled = True
        Me.cmdNew.Enabled = True
    Else
        Me.lstCounsel.Enabled = False
        Me.cmdNew.Enabled = False
    End If
    Call controlSettingByClickListCouncel
    
    '--//클릭 시 텍스트박스 초기화
    Call sbtxtBox_Init
    
    '--//이력 목록상자 조정
    With Me.lstCounsel
        .ColumnCount = 9
        .ColumnHeads = False
        .ColumnWidths = "0,0,80,80,120,0,0,0,80" '상담코드, 생명번호, 상담일, 카테고리, 제목, 내용, 결과, 비고, 상태
        .Width = 351.75
        .Height = 91.95
        .TextAlign = fmTextAlignLeft
        .Font = "굴림"
    End With
    
    '--//사진추가
    Dim strLifeNo As String
    With Me.lstPStaff
        strLifeNo = .List(.listIndex, 6)
    End With
    InsertPicToLabel Me.lblPic, strLifeNo
    
    '--//이력목록 데이터 채우기
    Call makeSelectSQL(TB2)
    Call makeListData(strSql, TB2)
    If cntRecord > 0 Then
        Me.lstCounsel.List = LISTDATA
    Else
        Call truncateControlContent
    End If
    Call sbClearVariant
    
End Sub

Private Sub truncateControlContent()

    Me.lstCounsel.Clear
    Me.txtTitle = ""
    Me.txtContent = ""
    Me.txtResult = ""
    Me.txtRemark = ""
    Me.txtInputDate = ""
    Me.cboCategory = ""
    Me.cboStatus = ""

End Sub

Private Sub lstPStaff_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Un'HookListBoxScroll
End Sub

Private Sub lstPStaff_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Me.lstPStaff.ListCount > 0 Then
        'HookListBoxScroll Me, Me.lstPStaff
    End If
End Sub

Private Sub cboSearchArg_Duration_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Un'HookListBoxScroll
End Sub

Private Sub cboSearchArg_Duration_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Me.cboSearchArg_Duration.ListCount > 0 Then
        'HookListBoxScroll Me, Me.cboSearchArg_Duration
    End If
End Sub

Private Sub cboSearchArg_Category_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Un'HookListBoxScroll
End Sub

Private Sub cboSearchArg_Category_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Me.cboSearchArg_Category.ListCount > 0 Then
        'HookListBoxScroll Me, Me.cboSearchArg_Category
    End If
End Sub

Private Sub cboSearchArg_Status_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Un'HookListBoxScroll
End Sub

Private Sub cboSearchArg_Status_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Me.cboSearchArg_Status.ListCount > 0 Then
        'HookListBoxScroll Me, Me.cboSearchArg_Status
    End If
End Sub

Private Sub txtChurch_Change()
    Me.txtChurch.BackColor = RGB(255, 255, 255)
End Sub

Private Sub txtInputDate_Change()
    Call Date_Format(Me.txtInputDate)
End Sub

Private Sub UserForm_Initialize()
    If connIP = Empty Then setGlobalVariant '//전역변수 재설정
    If checkLogin = 0 Then f_login.Show '//로그인체크
    If checkLogin = 0 Then End '--//로그인 실패 시 프로시저 종료
    
    '--//기초설정
    Me.cmdClose.Cancel = True
    TB1 = "op_system.v0_pstaff_information" '--//목회자리스트
    TB2 = "op_system.db_counsel" '--//상담이력
    TB3 = "op_system.v0_pstaff_information_all" '--//목회자리스트 전체
    
    '--//권한에 따른 컨트롤 설정
    Call HideDeleteButtonByUserAuth
    
    '--//컨트롤 설정
    Call controlSettingByClickListCouncel
    Me.cmdCancel.Visible = False
    Me.cmdAdd.Visible = False
    Me.cmdNew.Enabled = False
    
    '--//콤보박스 설정
    Me.cboStatus.Clear
    Me.cboStatus.AddItem "진행"
    Me.cboStatus.AddItem "완료"
    Me.cboStatus.AddItem "취소"
    
    Me.cboCategory.Clear
    Me.cboSearchArg_Category.Clear
    Call makeListData("select * from op_system.a_counsel_category;", "op_system.a_counsel_category")
    If cntRecord > 0 Then
        Me.cboCategory.List = LISTDATA
        Me.cboSearchArg_Category.List = LISTDATA
        Me.cboSearchArg_Category.AddItem "전체", 0
    End If
    Call sbClearVariant
    
    Me.cboSearchArg_Duration.Clear
    Me.cboSearchArg_Duration.AddItem "전체기간"
    Me.cboSearchArg_Duration.AddItem "최근 1주"
    Me.cboSearchArg_Duration.AddItem "최근 1개월"
    Me.cboSearchArg_Duration.AddItem "최근 3개월"
    
    Me.cboSearchArg_Status.Clear
    Me.cboSearchArg_Status.AddItem "전체"
    Me.cboSearchArg_Status.AddItem "진행"
    Me.cboSearchArg_Status.AddItem "완료"
    Me.cboSearchArg_Status.AddItem "취소"
    
    Me.cboSearchArg_Duration.listIndex = 0
    Me.cboSearchArg_Category.listIndex = 0
    Me.cboSearchArg_Status.listIndex = 0
    
    '--//리스트박스 설정
    With Me.lstPStaff
        .ColumnCount = 10
        .ColumnHeads = False
        .ColumnWidths = "0,120,0,0,0,0,0,80,0,60" '교회코드, 교회명, 영문교회명, 지교회명, 영문지교회명, 선교국가, 생명번호, 한글이름(직분), 영문이름, 직책
'        .Width = 330
        .TextAlign = fmTextAlignLeft
        .Font = "굴림"
    End With
    
End Sub
Private Sub cmdSearch_Click()
    
    Me.lstCounsel.Clear
    
    If Me.chkAll Then
        strSql = makeSelectSQL(TB3)
        Call makeListData(strSql, TB3)
    Else
        strSql = makeSelectSQL(TB1)
        Call makeListData(strSql, TB1)
    End If
    
    If cntRecord = 0 Then
        MsgBox "반환할 DB 데이터가 없습니다.", vbInformation, banner
        Call sbClearVariant
        Exit Sub
    End If
    
    Me.lstPStaff.List = LISTDATA
    Call sbClearVariant
    Me.lstPStaff.Enabled = True
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
        strSql = "SELECT `교회코드`,`교회명`,`영문교회명`,`지교회명`,`영문지교회명`,`선교국가`,`생명번호`,`한글이름(직분)`,`영문이름`,`직책`,`관리부서`" & _
                " FROM " & TB1 & _
                " WHERE (`관리부서` = " & USER_DEPT & ") AND (`교회명` is not null)" & _
                " AND (`교회명` LIKE '%" & Me.txtName & "%' OR `지교회명` LIKE '%" & Me.txtName & "%'" & _
                " OR `영문교회명` LIKE '%" & Me.txtName & "%' OR `영문지교회명` LIKE '%" & Me.txtName & "%'" & _
                " OR `생명번호` LIKE '%" & Me.txtName & "%' OR `한글이름(직분)` LIKE '%" & Me.txtName & "%'" & _
                " OR `영문이름` LIKE '%" & Me.txtName & "%')" & _
                " UNION" & _
                " SELECT `교회코드`,`교회명`,`영문교회명`,`지교회명`,`영문지교회명`,`선교국가`,`배우자생번`,`사모한글이름(직분)`,`사모영문이름`,`사모직책`,`관리부서`" & _
                " FROM " & TB1 & _
                " WHERE (`관리부서` = " & USER_DEPT & ") AND (`교회명` is not null)" & _
                " AND (`교회명` LIKE '%" & Me.txtName & "%' OR `지교회명` LIKE '%" & Me.txtName & "%'" & _
                " OR `영문교회명` LIKE '%" & Me.txtName & "%' OR `영문지교회명` LIKE '%" & Me.txtName & "%'" & _
                " OR `배우자생번` LIKE '%" & Me.txtName & "%' OR `사모한글이름(직분)` LIKE '%" & Me.txtName & "%'" & _
                " OR `사모영문이름` LIKE '%" & Me.txtName & "%')" & _
                " ORDER BY `직책` IS NULL ASC, FIELD(`직책`,'당회장','당회장대리','당사모','당대리사모','동역','동사모','지교회관리자','지관자사모','예배소관리자','예관자사모','예비생도1단계','예비생도2단계','예비생도3단계','생도사모'," & getPosition2Joining & ",'');"
    Case TB2
        
        Select Case Me.cboSearchArg_Duration.listIndex
        Case 0:
            strSql = "SELECT *" & _
                " FROM " & TB2 & _
                " WHERE life_no = " & SText(Me.lstPStaff.List(Me.lstPStaff.listIndex, 6)) & _
                " AND (category LIKE '%" & Replace(Me.cboSearchArg_Category, "전체", "") & "%'" & _
                " AND status LIKE '%" & Replace(Me.cboSearchArg_Status, "전체", "") & "%')" & _
                " ORDER BY counsel_dt DESC;"
        Case 1:
            strSql = "SELECT *" & _
                " FROM " & TB2 & _
                " WHERE life_no = " & SText(Me.lstPStaff.List(Me.lstPStaff.listIndex, 6)) & _
                " AND counsel_dt BETWEEN ADDDATE(CURDATE(), INTERVAL -1 WEEK) AND CURDATE()" & _
                " AND (category LIKE '%" & Replace(Me.cboSearchArg_Category, "전체", "") & "%'" & _
                " AND status LIKE '%" & Replace(Me.cboSearchArg_Status, "전체", "") & "%')" & _
                " ORDER BY counsel_dt DESC;"
        Case 2:
            strSql = "SELECT *" & _
                " FROM " & TB2 & _
                " WHERE life_no = " & SText(Me.lstPStaff.List(Me.lstPStaff.listIndex, 6)) & _
                " AND counsel_dt BETWEEN ADDDATE(CURDATE(), INTERVAL -1 MONTH) AND CURDATE()" & _
                " AND (category LIKE '%" & Replace(Me.cboSearchArg_Category, "전체", "") & "%'" & _
                " AND status LIKE '%" & Replace(Me.cboSearchArg_Status, "전체", "") & "%')" & _
                " ORDER BY counsel_dt DESC;"
        Case 3:
            strSql = "SELECT *" & _
                " FROM " & TB2 & _
                " WHERE life_no = " & SText(Me.lstPStaff.List(Me.lstPStaff.listIndex, 6)) & _
                " AND counsel_dt BETWEEN ADDDATE(CURDATE(), INTERVAL -3 MONTH) AND CURDATE()" & _
                " AND (category LIKE '%" & Replace(Me.cboSearchArg_Category, "전체", "") & "%'" & _
                " AND status LIKE '%" & Replace(Me.cboSearchArg_Status, "전체", "") & "%')" & _
                " ORDER BY counsel_dt DESC;"
        End Select
        
    Case TB3
        strSql = "SELECT `교회코드`,`교회명`,`영문교회명`,`지교회명`,`영문지교회명`,`선교국가`,`생명번호`,`한글이름(직분)`,`영문이름`,`직책`,`관리부서`" & _
                " FROM " & TB3 & _
                " WHERE (`관리부서` = " & USER_DEPT & ") AND (`교회명` is not null)" & _
                " AND (`교회명` LIKE '%" & Me.txtName & "%' OR `지교회명` LIKE '%" & Me.txtName & "%'" & _
                " OR `영문교회명` LIKE '%" & Me.txtName & "%' OR `영문지교회명` LIKE '%" & Me.txtName & "%'" & _
                " OR `생명번호` LIKE '%" & Me.txtName & "%' OR `한글이름(직분)` LIKE '%" & Me.txtName & "%'" & _
                " OR `영문이름` LIKE '%" & Me.txtName & "%')" & _
                " UNION" & _
                " SELECT `교회코드`,`교회명`,`영문교회명`,`지교회명`,`영문지교회명`,`선교국가`,`배우자생번`,`사모한글이름(직분)`,`사모영문이름`,`사모직책`,`관리부서`" & _
                " FROM " & TB3 & _
                " WHERE (`관리부서` = " & USER_DEPT & ") AND (`교회명` is not null)" & _
                " AND (`교회명` LIKE '%" & Me.txtName & "%' OR `지교회명` LIKE '%" & Me.txtName & "%'" & _
                " OR `영문교회명` LIKE '%" & Me.txtName & "%' OR `영문지교회명` LIKE '%" & Me.txtName & "%'" & _
                " OR `배우자생번` LIKE '%" & Me.txtName & "%' OR `사모한글이름(직분)` LIKE '%" & Me.txtName & "%'" & _
                " OR `사모영문이름` LIKE '%" & Me.txtName & "%')" & _
                " ORDER BY `직책` IS NULL ASC, FIELD(`직책`,'당회장','당회장대리','당사모','당대리사모','동역','동사모','지교회관리자','지관자사모','예배소관리자','예관자사모','예비생도1단계','예비생도2단계','예비생도3단계','생도사모'," & getPosition2Joining & ",'');"
    Case Else
    End Select
    makeSelectSQL = strSql
End Function
Private Function makeUpdateSQL(ByVal tableNM As String) As String
    
    Select Case tableNM
    Case TB1
    Case TB2
        With Me.lstCounsel
            strSql = "UPDATE " & TB2 & " a " & _
                    "SET a.counsel_dt = " & SText(Me.txtInputDate) & _
                    ",a.category = " & SText(Me.cboCategory) & ",a.title = " & SText(Me.txtTitle) & _
                    ",a.content = " & SText(Me.txtContent) & ",a.result = " & SText(Me.txtResult) & _
                    ",a.remark = " & SText(Me.txtRemark) & ",a.status = " & SText(Me.cboStatus) & _
                    " WHERE a.counsel_id = " & SText(.List(.listIndex)) & ";"
        queryKey = .listIndex
        End With
    Case Else
    End Select
    makeUpdateSQL = strSql
End Function
Private Function makeInsertSQL(ByVal tableNM As String, argData As T_COUNSEL) As String
    
    Select Case tableNM
    Case TB1
    Case TB2
        strSql = "INSERT INTO " & TB2 & " VALUES(DEFAULT," & _
                    SText(argData.LIFE_NO) & "," & _
                    SText(argData.COUNSEL_DT) & "," & _
                    SText(argData.CATEGORY) & "," & _
                    SText(argData.title) & "," & _
                    SText(argData.CONTENT) & "," & _
                    SText(argData.result) & "," & _
                    SText(argData.REMARK) & "," & _
                    SText(argData.STATUS) & ");"
        queryKey = Me.lstCounsel.ListCount - 1
    Case Else
    End Select
    makeInsertSQL = strSql
End Function
Private Function makeDeleteSQL(ByVal tableNM As String) As String
    
    Select Case tableNM
    Case TB1
    Case TB2
        With Me.lstCounsel
            strSql = "DELETE FROM " & TB2 & " WHERE counsel_id = " & SText(.List(.listIndex)) & ";"
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
    Me.txtInputDate.Value = ""
    Me.txtTitle.Value = ""
    Me.txtContent.Value = ""
    Me.txtResult.Value = ""
    Me.txtRemark.Value = ""
    Me.cboCategory.Value = ""
    Me.cboStatus.Value = ""
End Sub
Private Function fnData_Validation()
'---------------------------------------
'유저폼 입력값에 대한 데이터 유효성 검사
'TRUE: 이상없음, FALSE: 잘못됨.
'---------------------------------------
    fnData_Validation = True '데이터가 유효하다는 가정 하에 시작
    
    '--//날짜 형식체크
    If Not IsDate(Me.txtInputDate) Then
        MsgBox "올바른 날짜 형태가 아닙니다. 상담일을 다시 확인 해주세요.", vbCritical, banner
        Set txtBox_Focus = Me.txtInputDate: fnData_Validation = False: Exit Function
    End If
    
    '--//콤보박스 값 검토
    If Not (Me.cboStatus = "진행" Or Me.cboStatus = "완료" Or Me.cboStatus = "취소") Then
        MsgBox "상태 값이 올바르지 않습니다. 다시 한 번 확인해 주세요.", vbCritical, banner
        Set txtBox_Focus = Me.cboStatus: fnData_Validation = False: Exit Function
    End If
    
    strSql = "SELECT * FROM op_system.a_counsel_category"
    makeListData strSql, "op_system.a_counsel_category"
    If IsInArray(Me.cboCategory, LISTDATA, True, rtnValue) = -1 Then
        MsgBox "상담 카테고리가 올바르지 않습니다. 다시 한 번 확인해 주세요.", vbCritical, banner
        Set txtBox_Focus = Me.cboCategory: fnData_Validation = False: Exit Function
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
    Me.lstPStaff.Enabled = Not argBoolean
    Me.lstCounsel.Enabled = Not argBoolean
    Me.cboSearchArg_Duration.Enabled = Not argBoolean
    Me.cboSearchArg_Category.Enabled = Not argBoolean
    Me.cboSearchArg_Status.Enabled = Not argBoolean
    Me.txtName.Enabled = Not argBoolean
    Me.chkAll.Enabled = Not argBoolean
    Me.cmdSearch.Enabled = Not argBoolean
    
    Me.txtTitle.Enabled = argBoolean
    Me.txtContent.Enabled = argBoolean
    Me.txtResult.Enabled = argBoolean
    Me.txtRemark.Enabled = argBoolean
    Me.txtInputDate.Enabled = argBoolean
    Me.cboCategory.Enabled = argBoolean
    Me.cboStatus.Enabled = argBoolean
    
    '--//기본값 설정
    If argBoolean = True Then
        Me.cboStatus.listIndex = 0
        Me.txtInputDate = Date
    End If

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
