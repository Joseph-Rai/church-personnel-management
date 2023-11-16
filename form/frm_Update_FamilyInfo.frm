VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Update_FamilyInfo 
   Caption         =   "가족정보 관리마법사"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13260
   OleObjectBlob   =   "frm_Update_FamilyInfo.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "frm_Update_FamilyInfo"
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
Dim DUPLICATION As Boolean '--//관리대상 중 중복이 있으면 가족코드만 업데이트

Private Sub cboPosition_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Un'HookListBoxScroll
End Sub

Private Sub cboPosition_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Me.cboPosition.ListCount > 0 Then
        'HookListBoxScroll Me, Me.cboPosition
    End If
End Sub

Private Sub cboRelations_Change()
    
    '--//만약 가족관계 콤보박스의 값이 '부' 혹은 '형제'가 아니면
    If IsInArray(Me.cboRelations.Value, Array("부", "형제"), , rtnSequence) = -1 Then
        '--//여성직분으로 세팅
        strSql = "SELECT * FROM op_system.a_title a WHERE a.title NOT IN ('목사', '장로')"
    Else
        '--//남성직분으로 세팅
        strSql = "SELECT * FROM op_system.a_title a WHERE a.title NOT IN ('권사')"
    End If
    
    Call makeListData(strSql, "op_system.a_title")
    
    Me.cboTitle.Clear
    Me.cboTitle.List = LISTDATA
    
    Call sbClearVariant
    
End Sub

Private Sub cboReligion_Change()
        
    Dim CtrlBox As MSForms.control
    
    If Me.cboReligion = "본교성도" Then
        For Each CtrlBox In Me.controls
            If InStr(CtrlBox.Name, "Title") > 0 Or InStr(CtrlBox.Name, "Position") > 0 Or InStr(CtrlBox.Name, "Church") > 0 Then
                CtrlBox.Visible = True
                On Error Resume Next
                Me.controls("lbl" & Right(CtrlBox.Name, Len(CtrlBox.Name) - 3)).Visible = True
                Me.controls("cmd" & Right(CtrlBox.Name, Len(CtrlBox.Name) - 3)).Visible = True
                On Error GoTo 0
            End If
        Next
    Else
        For Each CtrlBox In Me.controls
            If InStr(CtrlBox.Name, "Title") > 0 Or InStr(CtrlBox.Name, "Position") > 0 Or InStr(CtrlBox.Name, "Church") > 0 Then
                CtrlBox.Visible = False
                On Error Resume Next
                Me.controls("lbl" & Right(CtrlBox.Name, Len(CtrlBox.Name) - 3)).Visible = False
                Me.controls("cmd" & Right(CtrlBox.Name, Len(CtrlBox.Name) - 3)).Visible = False
                On Error GoTo 0
            End If
        Next
    End If
End Sub

Private Sub chkActivate_Click()
    Call Search_Mode(Me.chkActivate.Value)
End Sub

Private Sub cmdCancel_Click()
    Call Input_Mode(False)
    Call HideDeleteButtonByUserAuth
    Call lstFamily_Click
End Sub

Private Sub cmdChurch_Click()
    argShow = 1 '--//본교회만 검색
    argShow3 = 2 '--//확인 버튼 누를 시 frm_Update_FamilyInfo에 자료 삽입
    frm_Update_Appointment_1.Show
End Sub

Private Sub cmdClose_Click()
    
    Dim result As T_RESULT
    
    '--//선지자 정보 폼에 가족코드 넣기
    Select Case argShow2
    Case 1
        With Me.lstFamily
            frm_Update_PInformation.txtFamily = .List(IIf(.listIndex = -1, .ListCount - 1, .listIndex), 1)
        
            strSql = "SELECT a.family FROM op_system.db_pastoralstaff a WHERE a.lifeno = " & SText(frm_Update_PInformation.txtLifeNo) & ";"
            Call makeListData(strSql, "op_system.db_pastoralstaff")
            
            If LISTDATA(0, 0) <> .List(IIf(.listIndex = -1, .ListCount - 1, .listIndex), 1) Then
                strSql = "UPDATE op_system.db_pastoralstaff a SET a.family = " & SText(.List(.listIndex, 1)) & " WHERE a.lifeno = " & SText(frm_Update_PInformation.txtLifeNo) & ";"
                connectTaskDB
                result.strSql = strSql
                result.affectedCount = executeSQL("cmdClose_Click", "op_system.db_pastoralstaff", strSql, Me.Name, "가족코드 업데이트")
                writeLog "cmdClose_Click", "op_system.db_pastoralstaff", strSql, 0, Me.Name, "가족코드 업데이트"
                disconnectALL
            End If
        End With
    Case 2
        With Me.lstFamily
            frm_Update_PInformation.txtFamily_Spouse = .List(IIf(.listIndex = -1, .ListCount - 1, .listIndex), 1)
        
            strSql = "SELECT a.family FROM op_system.db_pastoralwife a WHERE a.lifeno = " & SText(frm_Update_PInformation.txtLifeNo_Spouse) & ";"
            Call makeListData(strSql, "op_system.db_pastoralwife")
            
            If LISTDATA(0, 0) <> .List(IIf(.listIndex = -1, .ListCount - 1, .listIndex), 1) Then
                strSql = "UPDATE op_system.db_pastoralwife a SET a.family = " & SText(.List(.listIndex, 1)) & " WHERE a.lifeno = " & SText(frm_Update_PInformation.txtLifeNo_Spouse) & ";"
                connectTaskDB
                result.strSql = strSql
                result.affectedCount = executeSQL("cmdClose_Click", "op_system.db_pastoralwife", strSql, Me.Name, "가족코드 업데이트")
                writeLog "cmdClose_Click", "op_system.db_pastoralwife", strSql, 0, Me.Name, "가족코드 업데이트"
                disconnectALL
            End If
        End With
    Case Else
    End Select
    
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    
    Dim result As T_RESULT
    
    '--//리스트박스를 선택하지 않았으면 프로시저 종료
    If Me.lstFamily.listIndex = -1 Then
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
    result.affectedCount = executeSQL("cmdDelete_Click", TB2, strSql, Me.Name, "가족구성원 삭제")
    writeLog "cmdDelete_Click", TB2, strSql, 0, Me.Name, "가족구성원 삭제"
    disconnectALL
    
    '--//메세지박스
    MsgBox "해당 데이터가 삭제되었습니다.", , banner
    
    '--//리스트박스 새로고침
    Call sbtxtBox_Init
    Call UserForm_Initialize
    
End Sub

Private Sub cmdEdit_Click()
    
    Dim result As T_RESULT
    
    DUPLICATION = False '--//중복이 없다는 가정 하에
    
    '--//리스트박스가 선택되어 있지 않으면 프로시저 종료
    If Me.lstFamily.listIndex = -1 Then
        MsgBox "수정할 데이터를 선택해 주세요.", vbCritical, banner
        Exit Sub
    End If
    
    '--//수정된 내용 있는지 체크
    '--//listindex: 0:family_id, 1:family_cd, 2:가족관계, 3:생명번호 4:한글이름, 5:영문이름,6:교회코드, 7:소속교회, 8:직분, 9:직책, 10:생년월일, 11:최종학력, 12:종교, 13:본교인식, 14:메모, 15:별세여부
    With Me.lstFamily
        If Me.cboRelations = Replace(.List(.listIndex, 2), "(별세)", "") And Me.txtName_ko = .List(.listIndex, 4) And Me.txtName_en = .List(.listIndex, 5) And Me.txtChurch_Sid = .List(.listIndex, 6) And Me.cboTitle = .List(.listIndex, 8) And _
            Me.cboPosition = .List(.listIndex, 9) And Me.txtBirthday = .List(.listIndex, 10) And Me.txtEducation = .List(.listIndex, 11) And Me.cboReligion = .List(.listIndex, 12) And _
            Me.cboRecognition = .List(.listIndex, 13) And Me.txtMemo = .List(.listIndex, 14) And Int(Me.chkDecedent) * -1 = Int(.List(.listIndex, 15)) Then
            Exit Sub
        End If
    End With
    
    '--//중복체크: 부친과 모친은 한 분만 입력 가능
    If Me.cboRelations = "부" Or Me.cboRelations = "모" Then
    '--//listindex: 0:family_id, 1:family_cd, 2:가족관계, 3:생명번호 4:한글이름, 5:영문이름,6:교회코드, 7:소속교회, 8:직분, 9:직책, 10:생년월일, 11:최종학력, 12:종교, 13:본교인식, 14:메모, 15:별세여부
        With Me.lstFamily
            strSql = "SELECT * FROM " & TB2 & " a WHERE a.family_cd = " & SText(.List(.listIndex, 1)) & " AND a.relations = " & SText(Me.cboRelations) & ";"
            Call makeListData(strSql, TB2)
        End With
        
        If cntRecord > 0 Then
            If LISTDATA(0, 0) <> Me.lstFamily.List(Me.lstFamily.listIndex) Then '--//id까지 다르면 오류발생
                MsgBox Me.cboRelations & "친은 중복될 수 없습니다. 다시 확인해주세요.", vbCritical, banner
                Me.cboRelations.SetFocus
                Me.cboRelations.SelStart = 0
                Me.cboRelations.SelLength = Len(Me.cboRelations)
                Exit Sub
            End If
        End If
        Call sbClearVariant
    Else
        If Me.chkActivate.Value = True And Me.txtLifeNo <> "" Then
            With Me.lstFamily
                strSql = "SELECT * FROM " & TB2 & " a WHERE a.family_cd <> " & SText(.List(.listIndex, 1)) & " AND a.lifeno = " & SText(Me.txtLifeNo) & ";"
                Call makeListData(strSql, TB2)
            End With
            
            If cntRecord > 0 Then
                DUPLICATION = True '--//중복
            End If
        End If
    End If
    
    '--//데이터 유효성 검사
    If fnData_Validation = False Then
On Error Resume Next
        txtBox_Focus.SetFocus '--//에러난 지점에 포커싱
        txtBox_Focus.SelStart = 0
        txtBox_Focus.SelLength = Len(txtBox_Focus)
On Error GoTo 0
        Exit Sub
    End If
    
    '--//SQL문 생성, 실행, 로그기록
    If DUPLICATION Then
        strSql = makeUpdateSQL2(TB2)
        result.strSql = strSql
        connectTaskDB
        result.affectedCount = executeSQL("cmdEdit_Click", TB2, strSql, Me.Name, "가족구성원 업데이트")
        writeLog "cmdEdit_Click", TB2, strSql, 0, Me.Name, "가족구성원 업데이트", result.affectedCount
        disconnectALL
    Else
        strSql = makeUpdateSQL(TB2)
        result.strSql = strSql
        connectTaskDB
        result.affectedCount = executeSQL("cmdEdit_Click", TB2, strSql, Me.Name, "가족구성원 업데이트")
        writeLog "cmdEdit_Click", TB2, strSql, 0, Me.Name, "가족구성원 업데이트", result.affectedCount
        disconnectALL
    End If
    
    '--//메세지박스
    MsgBox "업데이트 되었습니다.", , banner
    
    '--//리스트박스 초기화
    If DUPLICATION Then
        strSql = "SELECT a.family_id FROM op_system.db_familyinfo a WHERE a.lifeno = " & SText(Me.txtLifeNo)
        Call makeListData(strSql, TB2)
        queryKey = LISTDATA(0, 0)
        Call sbClearVariant
        
        Call UserForm_Initialize
        Call returnListPosition(Me, Me.lstFamily.Name, queryKey)
    Else
        Call UserForm_Initialize
        Me.lstFamily.listIndex = queryKey
    End If
    
End Sub

Private Sub cmdNew_Click()
    
    '--//커맨드 버튼 활성화를 위한 이력 리스트박스 클릭
    If lstFamily.ListCount = 0 Then
        Call lstFamily_Click
    End If
    
    Me.lstFamily.listIndex = Me.lstFamily.ListCount - 1
    Call Input_Mode(True)
    Call HideDeleteButtonByUserAuth
'    Call sbtxtBox_Init
    Call Search_Mode(False)
    
    Me.cboRelations.SetFocus
End Sub

Private Sub cmdADD_Click()
    
    Dim argData As T_FAMILY
    Dim result As T_RESULT
    Dim i As Integer
    
    DUPLICATION = False
    
    '--//중복체크: 부친과 모친은 한 분만 입력 가능
    If Me.cboRelations = "부" Or Me.cboRelations = "모" Then
    '--//listindex: 0:family_id, 1:family_cd, 2:가족관계, 3:생명번호 4:한글이름, 5:영문이름,6:교회코드, 7:소속교회, 8:직분, 9:직책, 10:생년월일, 11:최종학력, 12:종교, 13:본교인식, 14:메모, 15:별세여부
        With Me.lstFamily
            strSql = "SELECT * FROM " & TB2 & " a WHERE a.family_cd = " & SText(.List(.listIndex, 1)) & " AND a.relations = " & SText(Me.cboRelations) & ";"
            Call makeListData(strSql, TB2)
        End With
        
        If cntRecord > 0 Then
            MsgBox Me.cboRelations & "친은 중복될 수 없습니다. 다시 확인해주세요.", vbCritical, banner
            Me.cboRelations.SetFocus
            Me.cboRelations.SelStart = 0
            Me.cboRelations.SelLength = Len(Me.cboRelations)
            Exit Sub
        End If
        Call sbClearVariant
    Else
        If Me.chkActivate.Value = True And Me.txtLifeNo <> "" Then
            With Me.lstFamily
                strSql = "SELECT * FROM " & TB2 & " a WHERE a.family_cd <> " & SText(.List(.listIndex, 1)) & " AND a.lifeno = " & SText(Me.txtLifeNo) & ";"
                Call makeListData(strSql, TB2)
            End With
            
            If cntRecord > 0 Then
                DUPLICATION = True '--//중복
            End If
        End If
    End If
    
    '--//데이터 유효성 검사
    If fnData_Validation = False Then
On Error Resume Next
        txtBox_Focus.SetFocus '--//에러난 지점에 포커싱
        txtBox_Focus.SelStart = 0
        txtBox_Focus.SelLength = Len(txtBox_Focus)
On Error GoTo 0
        Exit Sub
    End If
    
    If DUPLICATION Then
        '--//해당 생명번호에 대한 가족코드 업데이트
        strSql = makeUpdateSQL2(TB2)
        result.strSql = strSql
        connectTaskDB
        result.affectedCount = executeSQL("cmdEdit_Click", TB2, strSql, Me.Name, "가족구성원 업데이트")
        writeLog "cmdEdit_Click", TB2, strSql, 0, Me.Name, "가족구성원 업데이트", result.affectedCount
        disconnectALL
    Else
        '--//구조체에 값 추가
        With Me.lstFamily
            argData.FAMILY_CD = .List(.listIndex, 1)
        End With
        If Me.txtLifeNo <> "" Then
            argData.lifeNo = Me.txtLifeNo
        End If
        argData.RELATIONS = Me.cboRelations.Value
        argData.NAME_KO = Me.txtName_ko
        argData.name_en = Me.txtName_en
        argData.church_sid = Me.txtChurch_Sid
        argData.title = Me.cboTitle
        argData.position = Me.cboPosition
        argData.Birthday = IIf(Me.txtBirthday = "", "1900-01-01", Me.txtBirthday)
        argData.Education = Me.txtEducation
        argData.RELIGION = Me.cboReligion.Value
        argData.memo = Me.txtMemo
        argData.RECOGNITION = Me.cboRecognition.Value
        argData.Suspend = Me.chkDecedent.Value
        
        '--//쿼리문 실행 및 로그기록
        strSql = makeInsertSQL(TB2, argData)
        connectTaskDB
        result.strSql = strSql
        result.affectedCount = executeSQL("cmdADD_Clikc", TB2, strSql, Me.Name, "가족구성원 추가")
        writeLog "cmdADD_Click", TB2, strSql, 0, Me.Name, "가족구성원 추가", result.affectedCount
        disconnectALL
    End If
    
    '--//메세지박스
    MsgBox "추가 되었습니다.", , banner
    
    If DUPLICATION Then
        strSql = "SELECT a.family_id FROM op_system.db_familyinfo a WHERE a.lifeno = " & SText(Me.txtLifeNo)
        Call makeListData(strSql, TB2)
        queryKey = LISTDATA(0, 0)
        Call sbClearVariant
        
        Call UserForm_Initialize
        Call returnListPosition(Me, Me.lstFamily.Name, queryKey)
    Else
        Call UserForm_Initialize
        With Me.lstFamily
            strSql = "SELECT a.family_id FROM op_system.db_familyinfo a WHERE a.family_cd = " & SText(.List(.listIndex, 1))
            Call makeListData(strSql, TB2)
            For i = 0 To UBound(LISTDATA)
                queryKey = WorksheetFunction.Max(queryKey, LISTDATA(i, 0))
            Next
            Call sbClearVariant
        End With
        Call returnListPosition(Me, Me.lstFamily.Name, queryKey)
    End If
    
    '--//버튼설정 원래대로
    Call Input_Mode(False)
    Me.chkActivate.Value = False
'    If Me.chkActivate.Value = True Then
'        Me.chkActivate.Value = False
'        Call Search_Mode(Me.chkActivate.Value)
'    End If
'    Call cmdbtn_visible
'    Call HideDeleteButtonByUserAuth
    Call lstFamily_Click
    
End Sub
Private Sub cmdSearch_Click()
    argShow = 2
    frm_Update_BCLeader_1.Show
    
    '--//생명번호가 비어있지 않으면
    If Me.txtLifeNo <> "" Then
        Me.cboRelations.Enabled = True
        Me.cboRelations.BackColor = RGB(255, 255, 255)
        Me.chkDecedent.Enabled = True
        Me.chkDecedent.BackColor = RGB(255, 255, 255)
    End If
End Sub

Private Sub Frame1_Click()

End Sub

Private Sub lstFamily_Click()
    
    Dim CtrlBox As MSForms.control
    
    '--//컨트롤 설정
    Call sbtxtBox_Init
    Me.cmdSearch.Enabled = True
    Me.cmdEdit.Enabled = True
    Me.cmdDelete.Enabled = True
    
    '--//리스트 클릭 시 텍스트박스, 콤보박스에 내용추가
    If Me.lstFamily.listIndex <> -1 Then
        With Me.lstFamily
            Me.txtLifeNo = .List(.listIndex, 3)
            If InStr(.List(.listIndex, 2), "별세") > 0 Then
                Me.cboRelations = Left(.List(.listIndex, 2), InStr(.List(.listIndex, 2), "(") - 1)
            Else
                Me.cboRelations = .List(.listIndex, 2)
            End If
            Me.txtName_ko = .List(.listIndex, 4)
            Me.txtName_en = .List(.listIndex, 5)
            Me.txtChurch_Sid = .List(.listIndex, 6)
            Me.txtChurch = .List(.listIndex, 16)
            Me.cboTitle = .List(.listIndex, 8)
            Me.cboPosition = .List(.listIndex, 9)
            Me.txtBirthday = .List(.listIndex, 10)
            Me.txtEducation = .List(.listIndex, 11)
            Me.cboReligion = .List(.listIndex, 12)
            Me.cboRecognition = .List(.listIndex, 13)
            Me.txtMemo = .List(.listIndex, 14)
            Me.chkDecedent.Value = CBool(.List(.listIndex, 15))
        End With
    End If
    
    '--//txtChurch 내용에 따른 배경색 변경
    If Me.txtChurch = "" Then
        Me.txtChurch.BackColor = &HC0FFFF
    Else
        Me.txtChurch.BackColor = &HE0E0E0
    End If
    
    '--//관리대상 여부에 따른 컨트롤 박스 설정
    '1. 생명번호로 등록된 가족은 생명번호만 수정가능
    '2. 생명번호 없이 직접 입력된 가족은 생명번호 수정불가
    '3. 본인이면 아무것도 수정불가
    With Me.lstFamily
        If .List(.listIndex, 3) <> "" Then
            If .List(.listIndex, 2) = "본인" Then
                Call Search_Mode(True, False)
                Me.txtLifeNo.Enabled = False
                Me.txtLifeNo.BackColor = &HE0E0E0
                Me.cmdSearch.Enabled = False
                Me.cmdDelete.Enabled = False
'                Me.cboRecognition.Enabled = False
'                Me.cboRecognition.BackColor = &HE0E0E0
            Else
                Call Search_Mode(True, False)
                Me.cmdSearch.Enabled = True
                Me.cmdDelete.Enabled = True
            End If
        Else
            Call Search_Mode(False, False)
        End If
    End With
    
End Sub

Private Sub lstfamily_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Un'HookListBoxScroll
End Sub

Private Sub lstfamily_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Me.lstFamily.ListCount > 0 Then
        'HookListBoxScroll Me, Me.lstFamily
    End If
End Sub

Private Sub txtBirthday_Change()
    Call Date_Format(Me.txtBirthday)
End Sub

Private Sub UserForm_Initialize()
    
    Dim CtrlBox As MSForms.control
    Dim result As T_RESULT
    
    If connIP = Empty Then setGlobalVariant '//전역변수 재설정
    If checkLogin = 0 Then f_login.Show '//로그인체크
    If checkLogin = 0 Then End '--//로그인 실패 시 프로시저 종료
    
    '--//기초설정
    Me.cmdClose.Cancel = True
    TB1 = "op_system.v_familyinfo" '--//가족정보 뷰
    TB2 = "op_system.db_familyinfo" '--//가족정보 테이블
    
    '--//권한에 따른 컨트롤 설정
    Call HideDeleteButtonByUserAuth
    
    '--//컨트롤 설정
    Me.cmdEdit.Enabled = False
    Me.cmdAdd.Visible = False
    Me.cmdCancel.Visible = False
    Me.cmdSearch.Enabled = False
    Me.cmdDelete.Enabled = False
    Me.chkActivate.Visible = False
    Me.lblInfo1.Visible = False
    For Each CtrlBox In Me.controls
        If TypeName(CtrlBox) = "TextBox" Or TypeName(CtrlBox) = "ComboBox" Or TypeName(CtrlBox) = "CheckBox" Then
                CtrlBox.Enabled = False
                CtrlBox.BackColor = &HE0E0E0
        End If
    Next
'    Me.cboPosition.List = Array("당회장", "당사모", "당회장대리", "당대리사모", "동역", "임시동역", "동사모", "생도", "예비생도", "생도사모", "지교회관리자", "지관지사모", "예배소관리자", "예관자사모", "장지역장", "(임)장지역장", "장구역장", "(임)장구역장", "청지역장", "(임)청지역장", "청구역장", "(임)청구역장", "학지역장", "(임)학지역장", "학구역장", "(임)학구역장", "부장지역장", "부장구역장", "부청지역장", "부청구역장", "부학지역장", "부학구역장")
    strSql = "SELECT * FROM op_system.a_position;"
    Call makeListData(strSql, "a_position")
    Me.cboPosition.List = LISTDATA
    Me.cboPosition.AddItem "예비생도"
    Me.cboPosition.AddItem "당사모"
    Me.cboPosition.AddItem "당대리사모"
    Me.cboPosition.AddItem "동사모"
    Me.cboPosition.AddItem "생도사모"
    Me.cboPosition.AddItem "지관자사모"
    Me.cboPosition.AddItem "예관자사모"
    sbClearVariant
    
    '--//리스트박스 채우기
    Call Make_FirstRecord '--//첫 데이터 없으면 생성
    
    Select Case argShow2 '--//생명번호 기준으로 family_cd 가져오기
    Case 1
        strSql = "SELECT a.family_cd FROM " & TB1 & " a WHERE a.lifeno = " & SText(frm_Update_PInformation.txtLifeNo) & " AND a.relations NOT IN ('부','모');"
    Case 2
        strSql = "SELECT a.family_cd FROM " & TB1 & " a WHERE a.lifeno = " & SText(frm_Update_PInformation.txtLifeNo_Spouse) & " AND a.relations NOT IN ('부','모');"
    Case Else
    End Select
    Call makeListData(strSql, TB1)
    
    Call makeSelectSQL2(TB1) '--//family_cd기준으로 가족구성원 불러오기
    Call makeListData(strSql, TB1)
    If cntRecord > 0 Then
        Me.lstFamily.List = LISTDATA
    End If
    Call sbClearVariant
    
    '--//리스트박스 설정
    With Me.lstFamily
        If .listIndex = -1 Or Me.lstFamily.Width < 500 Then '--//유저폼 처음 열었을 때에만 실시
            .ColumnCount = 17
            .ColumnHeads = False
            .ColumnWidths = "0,0,50,0,70,80,0,100,40,50,70,0,70,70,0,0,0" '구성원id,가족코드,가족관계,생명번호,한글이름,영문이름,교회코드,소속교회,직분,직책,생년월일,최종학력,종교,본교인식,메모,별세여부,교회풀네임
            .Width = 624.45
            .TextAlign = fmTextAlignLeft
            .Font = "굴림"
        
            '--//콤보박스 채우기
            Me.cboRecognition.Clear
            Me.cboRelations.Clear
            Me.cboReligion.Clear
            Me.cboRecognition.List = Array("우호", "보통", "나쁨")
            Me.cboRelations.List = Array("부", "모", "형제", "자매")
            Me.cboReligion.List = Array("본교성도", "기독교", "천주교", "힌두교", "불교", "이슬람", "무신론", "기타")
        End If
    End With
    
    '--//자신을 가리키도록 리스트인덱스 설정
    Select Case argShow2
    Case 1
        strSql = "SELECT a.family_id FROM op_system.db_familyinfo a WHERE a.lifeno = " & SText(frm_Update_PInformation.txtLifeNo) & " AND a.relations NOT IN ('부','모');"
    Case 2
        strSql = "SELECT a.family_id FROM op_system.db_familyinfo a WHERE a.lifeno = " & SText(frm_Update_PInformation.txtLifeNo_Spouse) & " AND a.relations NOT IN ('부','모');"
    Case Else
    End Select
    
    Call makeListData(strSql, "op_system.db_familyinfo")
    If Me.lstFamily.listIndex = -1 Then
        queryKey = LISTDATA(0, 0)
        Call returnListPosition(Me, "lstFamily", queryKey)
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
        strSql = "SELECT * FROM " & TB1 & " a WHERE a.lifeno = " & SText(Int(LISTDATA(0, 0))) & " ORDER BY a.birthday;"
    Case TB2
    
    Case Else
    End Select
    makeSelectSQL = strSql
End Function
Private Function makeSelectSQL2(ByVal tableNM As String) As String
    
    Select Case tableNM
    Case TB1
        Select Case argShow2
        Case 1
            strSql = "SELECT a.family_id,a.family_cd,IF(a.lifeno= " & SText(frm_Update_PInformation.txtLifeNo) & ",'본인',a.relations),a.lifeno,a.name_ko,a.name_en,a.church_sid,a.church_nm,a.title,a.position,a.birthday,a.education,a.religion,a.recognition,a.memo,a.suspend, a.churchFullName FROM " & TB1 & " a WHERE a.family_cd = " & SText(Int(LISTDATA(0, 0))) & " ORDER BY a.birthday;"
        Case 2
            strSql = "SELECT a.family_id,a.family_cd,IF(a.lifeno= " & SText(frm_Update_PInformation.txtLifeNo_Spouse) & ",'본인',a.relations),a.lifeno,a.name_ko,a.name_en,a.church_sid,a.church_nm,a.title,a.position,a.birthday,a.education,a.religion,a.recognition,a.memo,a.suspend, a.churchFullName FROM " & TB1 & " a WHERE a.family_cd = " & SText(Int(LISTDATA(0, 0))) & " ORDER BY a.birthday;"
        Case Else
        End Select
    Case TB2
    
    Case Else
    End Select
    makeSelectSQL2 = strSql
End Function
Private Function makeUpdateSQL(ByVal tableNM As String) As String
    
    Select Case tableNM
    Case TB1
    Case TB2
        With Me.lstFamily
            '--//생명번호가 있는 관리대상이면
            If .List(.listIndex, 3) <> "" Then
                strSql = "UPDATE " & TB2 & " a " & _
                        "SET a.recognition = " & SText(Me.cboRecognition.Value) & ",a.suspend = " & SText(Int(Me.chkDecedent) * -1) & ",a.memo = " & SText(Me.txtMemo) & _
                        " WHERE a.family_id = " & SText(.List(.listIndex)) & ";"
            Else
                strSql = "UPDATE " & TB2 & " a " & _
                        "SET a.relations = " & SText(Me.cboRelations.Value) & ", a.lifeno = " & SText(Me.txtLifeNo) & ", a.name_ko = " & SText(Me.txtName_ko) & ", a.name_en = " & SText(Me.txtName_en) & _
                        ",a.title = " & SText(Me.cboTitle) & ",a.position = " & SText(Me.cboPosition) & ",a.birthday = " & IIf(Me.txtBirthday = "", "NULL", SText(Me.txtBirthday)) & ",a.education = " & SText(Me.txtEducation) & _
                        ",a.religion = " & SText(Me.cboReligion.Value) & ",a.recognition = " & SText(Me.cboRecognition.Value) & ",a.suspend = " & SText(Int(Me.chkDecedent) * -1) & ",a.church_sid = " & SText(Me.txtChurch_Sid) & ",a.memo = " & SText(Me.txtMemo) & _
                        " WHERE a.family_id = " & SText(.List(.listIndex)) & ";"
            End If
        queryKey = .listIndex
        End With
    Case Else
    End Select
    makeUpdateSQL = strSql
End Function
Private Function makeUpdateSQL2(ByVal tableNM As String) As String
    
    Select Case tableNM
    Case TB1
    Case TB2
        With Me.lstFamily
            '--//생명번호가 있는 관리대상이면
            strSql = "UPDATE " & TB2 & " a " & _
                    "SET a.family_cd = " & .List(.listIndex, 1) & " ,a.recognition = " & SText(Me.cboRecognition.Value) & ",a.suspend = " & SText(Int(Me.chkDecedent) * -1) & ",a.memo = " & SText(Me.txtMemo) & _
                    " WHERE a.lifeno = " & SText(Me.txtLifeNo) & ";"
'        queryKey = .ListIndex
        End With
    Case Else
    End Select
    makeUpdateSQL2 = strSql
End Function
Private Function makeInsertSQL(ByVal tableNM As String, argData As T_FAMILY) As String
    
    Select Case tableNM
    Case TB1
    Case TB2
        If Me.txtLifeNo = "" Then
            strSql = "INSERT INTO " & TB2 & " VALUES(DEFAULT," & _
                        SText(argData.FAMILY_CD) & "," & _
                        SText(argData.RELATIONS) & "," & _
                        SText(argData.lifeNo) & "," & _
                        SText(argData.NAME_KO) & "," & _
                        SText(argData.name_en) & "," & _
                        SText(argData.church_sid) & "," & _
                        SText(argData.title) & "," & _
                        SText(argData.position) & "," & _
                        IIf(argData.Birthday = "1900-01-01", "NULL", SText(argData.Birthday)) & "," & _
                        SText(argData.Education) & "," & _
                        SText(argData.RELIGION) & "," & _
                        SText(argData.RECOGNITION) & "," & _
                        SText(argData.memo) & "," & _
                        SText(Int(argData.Suspend) * -1) & ");"
        Else
            strSql = "INSERT INTO " & TB2 & "(family_id,family_cd,relations,lifeno,recognition,suspend) VALUES(DEFAULT," & _
                        SText(argData.FAMILY_CD) & "," & _
                        SText(argData.RELATIONS) & "," & _
                        SText(argData.lifeNo) & "," & _
                        SText(argData.RECOGNITION) & "," & _
                        SText(Int(argData.Suspend) * -1) & ");"

        End If
'        queryKey = Me.lstFamily.ListCount - 1
    Case Else
    End Select
    makeInsertSQL = strSql
End Function
Private Function makeDeleteSQL(ByVal tableNM As String) As String
    
    Select Case tableNM
    Case TB1
    Case TB2
        With Me.lstFamily
            strSql = "DELETE FROM " & TB2 & " WHERE family_id = " & SText(.List(.listIndex)) & ";"
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

Private Function fnData_Validation()
'---------------------------------------
'유저폼 입력값에 대한 데이터 유효성 검사
'TRUE: 이상없음, FALSE: 잘못됨.
'---------------------------------------
    Dim CtrlBox As MSForms.control
    
    fnData_Validation = True '데이터가 유효하다는 가정 하에 시작
    
    '--//생명번호 데이터 입력여부 확인
    If Me.chkActivate.Value = True Then
        If Me.txtLifeNo = "" Then
            If Me.chkActivate = True Then
                MsgBox "생명번호를 입력해 주세요.", vbCritical, banner
                fnData_Validation = False: Exit Function
            End If
        End If
    End If
    
    '--//생명번호 형식체크
    If Me.txtLifeNo <> "" Then
        If Not IsNumeric(fnExtract(Me.txtLifeNo)) Then
            fnData_Validation = False
            MsgBox "선지자 생명번호가 잘못되었습니다. 다시 확인해 주세요.", vbCritical, banner
            Set txtBox_Focus = Me.txtLifeNo
            Exit Function
        ElseIf Mid(Me.txtLifeNo, 4, 1) <> "-" Or Mid(Me.txtLifeNo, 11, 1) <> "-" Then
            fnData_Validation = False
            MsgBox "선지자 생명번호가 잘못되었습니다. 다시 확인해 주세요.", vbCritical, banner
            Set txtBox_Focus = Me.txtLifeNo
            Exit Function
        End If
    End If
    
    '--//날짜 형식체크
    If Not IsDate(Me.txtBirthday) And Me.txtBirthday <> "" Then
        MsgBox "올바른 날짜 형태가 아닙니다. 시작일을 다시 확인 해주세요.", vbCritical, banner
        Set txtBox_Focus = Me.txtBirthday: fnData_Validation = False: Exit Function
    End If
    
    '--//콤보박스 유효성 검사(리스트에 없는 값 선택 시 오류)
    For Each CtrlBox In Me.controls
        If TypeName(CtrlBox) = "ComboBox" Then
            If IsInArray(CtrlBox.Value, CtrlBox.List, , rtnSequence) = -1 And CtrlBox <> "" And Me.chkActivate = False Then
                If Not (CtrlBox.Name = "cboRelations" And CtrlBox.Value = "본인") Then
                    If CtrlBox.Name Like "*Religion*" Then
                        MsgBox "종교를 잘못 선택하셨습니다. 다시 확인해 주세요.", vbCritical, banner
                        Set txtBox_Focus = CtrlBox: fnData_Validation = False: Exit Function
                    End If
                    If CtrlBox.Name Like "*Title*" Then
                        MsgBox "직분을 잘못 선택하셨습니다. 다시 확인해 주세요.", vbCritical, banner
                        Set txtBox_Focus = CtrlBox: fnData_Validation = False: Exit Function
                    End If
                    If CtrlBox.Name Like "*Position*" Then
                        MsgBox "직책을 잘못 선택하셨습니다. 다시 확인해 주세요.", vbCritical, banner
                        Set txtBox_Focus = CtrlBox: fnData_Validation = False: Exit Function
                    End If
                End If
            End If
        End If
    Next
    
    '--//이름 유효성 검사
    If fnExtract(Me.txtName_ko, "E") <> "" Then
        Set txtBox_Focus = Me.txtName_ko: fnData_Validation = False: Exit Function
    End If
    If fnExtract(Me.txtName_en, "H") <> "" Then
        Set txtBox_Focus = Me.txtName_en: fnData_Validation = False: Exit Function
    End If
    
    '--//필수값 체크
    If Me.cboRelations = "" Then
        MsgBox "가족관계를 입력해 주세요.", vbCritical, banner
        Set txtBox_Focus = Me.cboRelations: fnData_Validation = False: Exit Function
    End If
    
End Function

Sub sbtxtBox_Init()
    
    Dim CtrlBox As MSForms.control
    
    For Each CtrlBox In Me.controls
        If TypeName(CtrlBox) = "TextBox" Or TypeName(CtrlBox) = "ComboBox" Or TypeName(CtrlBox) = "CheckBox" Then
            If CtrlBox.Name <> Me.txtChurch.Name Then '--//txtChurch는 제외
                CtrlBox.Enabled = True
            End If
            If TypeName(CtrlBox) <> "CheckBox" Then
                CtrlBox.Value = ""
            End If
            CtrlBox.BackColor = RGB(255, 255, 255)
        End If
    Next
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

Private Sub Search_Mode(ByVal argBoolean As Boolean, Optional blnClear As Boolean = True)
    
    Dim CtrlBox As MSForms.control
    
    For Each CtrlBox In Me.Frame1.controls
        If TypeName(CtrlBox) = "TextBox" Or TypeName(CtrlBox) = "ComboBox" Or TypeName(CtrlBox) = "CheckBox" Then
            If CtrlBox.Name <> Me.cboRecognition.Name Then '--//cboRecognition은 제외
                
                If CtrlBox.Name <> Me.txtChurch.Name Then '--//txtChurch는 제외
                    CtrlBox.Enabled = Not argBoolean
                    CtrlBox.BackColor = IIf(argBoolean, &HE0E0E0, RGB(255, 255, 255))
                End If
                
                If blnClear Then
                    If TypeName(CtrlBox) = "CheckBox" Then
                        CtrlBox.Value = 0
                    Else
                        CtrlBox.Value = ""
                    End If
                End If
            End If
        End If
    Next
    Me.cmdSearch.Enabled = argBoolean
    Me.txtLifeNo.Enabled = False
    Me.txtLifeNo.BackColor = IIf(argBoolean, &HC0FFFF, &HE0E0E0)
End Sub
Private Sub Input_Mode(ByVal argBoolean As Boolean)
    Call sbtxtBox_Init
    Me.cmdNew.Visible = Not argBoolean
    Me.cmdEdit.Visible = Not argBoolean
    Me.cmdDelete.Visible = Not argBoolean
    Me.cmdClose.Visible = Not argBoolean
    Me.cmdCancel.Visible = argBoolean
    Me.cmdCancel.Enabled = argBoolean
    Me.cmdAdd.Visible = argBoolean
    Me.cmdAdd.Enabled = argBoolean
    
    Me.chkActivate.Value = Not argBoolean
    Me.chkActivate.Visible = argBoolean
    Me.chkActivate.Enabled = argBoolean
    
    Me.lblInfo1.Visible = argBoolean
    
    Me.lstFamily.Enabled = Not argBoolean
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

Private Sub Make_FirstRecord()

    Dim result As T_RESULT

    Select Case argShow2
    Case 1
        strSql = "SELECT a.family_cd FROM " & TB1 & " a WHERE a.lifeno = " & SText(frm_Update_PInformation.txtLifeNo) & " AND a.relations NOT IN ('부','모');"
    Case 2
        strSql = "SELECT a.family_cd FROM " & TB1 & " a WHERE a.lifeno = " & SText(frm_Update_PInformation.txtLifeNo_Spouse) & " AND a.relations NOT IN ('부','모');"
    Case Else
    End Select
            Call makeListData(strSql, TB1)
    If cntRecord = 0 Then '--//가족코드가 없으면 신규추가
        strSql = "SELECT MAX(a.family_cd) FROM " & TB1 & " a;"
        Call makeListData(strSql, TB2)
        Select Case argShow2
        Case 1
            strSql = "INSERT INTO " & TB2 & "(family_id,family_cd,lifeno,relations) VALUES (DEFAULT," & SText(Int(IIf(LISTDATA(0, 0) = "", 0, LISTDATA(0, 0))) + 1) & "," & SText(frm_Update_PInformation.txtLifeNo) & _
                        "," & SText(IIf(Mid(frm_Update_PInformation.txtLifeNo, 12, 1) = 1, "형제", "자매")) & ");"
        Case 2
            strSql = "INSERT INTO " & TB2 & "(family_id,family_cd,lifeno,relations) VALUES (DEFAULT," & SText(Int(IIf(LISTDATA(0, 0) = "", 0, LISTDATA(0, 0))) + 1) & "," & SText(frm_Update_PInformation.txtLifeNo_Spouse) & _
                        "," & SText(IIf(Mid(frm_Update_PInformation.txtLifeNo_Spouse, 12, 1) = 1, "형제", "자매")) & ");"
        Case Else
        End Select
        result.strSql = strSql
        connectTaskDB
        result.affectedCount = executeSQL("cmdEdit_Click", TB2, strSql, Me.Name, "가족구성원 추가")
        writeLog "cmdEdit_Click", TB2, strSql, 0, Me.Name, "가족구성원 추가", result.affectedCount
        disconnectALL
    End If
    sbClearVariant

End Sub
