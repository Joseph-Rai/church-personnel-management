VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Update_BCLeader 
   Caption         =   "관리자 이력관리 마법사"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8280.001
   OleObjectBlob   =   "frm_Update_BCLeader.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "frm_Update_BCLeader"
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

Private Sub chkDirect_Click()
    Call Direct_Mode(Me.chkDirect.Value)
End Sub

Private Sub chkPresent_Change()
    Select Case Me.chkPresent.Value
        Case True
            Me.txtEnd.BackColor = &HE0E0E0
            Me.txtEnd.Value = "현재"
            Me.txtEnd.Enabled = False
        Case False
            Me.txtEnd.Enabled = True
            Me.txtEnd.BackColor = RGB(255, 255, 255)
            If Me.lstHistory.listIndex = -1 Then
                Me.txtEnd = Date - 1
            Else
                If Me.txtEnd = "현재" Then
                    Me.txtEnd.Value = Date - 1
                End If
            End If
    Case Else
    End Select
End Sub

Private Sub cmdCancel_Click()
    'Call cmdbtn_visible
    Call sbtxtBox_Init
    Call INPUTMODE(False)
    Call HideDeleteButtonByUserAuth
    Call lstHistory_Click
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    
    Dim result As T_RESULT
    
    '--//리스트박스를 선택하지 않았으면 프로시저 종료
    If Me.lstHistory.listIndex = -1 Then
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
    result.affectedCount = executeSQL("cmdDelete_Click", TB2, strSql, Me.Name, "관리자이력 삭제")
    writeLog "cmdDelete_Click", TB2, strSql, 0, Me.Name, "관리자이력 삭제"
    disconnectALL
    
    '--//메세지박스
    MsgBox "해당 데이터가 삭제되었습니다.", , banner
    
    '--//리스트박스 새로고침
    Call lstChurch_Click
    Me.lstHistory.listIndex = -1
    
End Sub

Private Sub cmdEdit_Click()
    
    Dim result As T_RESULT
    
    '--//리스트박스가 선택되어 있지 않으면 프로시저 종료
    If Me.lstHistory.listIndex = -1 Then
        MsgBox "수정할 데이터를 선택해 주세요.", vbCritical, banner
        Exit Sub
    End If
    
    '--//수정된 내용 있는지 체크
    With Me.lstHistory
        If Me.txtStart = .List(.listIndex, 2) And Me.txtEnd = .List(.listIndex, 3) And Me.txtLifeNo = .List(.listIndex, 4) And Me.cboResponsibility = .List(.listIndex, 7) Then
            Exit Sub
        End If
    End With
    
    '--//중복체크
    With Me.lstHistory
        If Me.cboResponsibility = "관리자" Then
'            strSQL = "SELECT * FROM " & TB2 & " a WHERE a.church_sid = " & SText(Me.lstChurch.List(Me.lstChurch.ListIndex)) & _
'                    " AND a.responsibility = '관리자'" & _
'                    " AND ((a.start_dt BETWEEN " & SText(Me.txtStart) & " AND " & SText(IIf(Me.txtEnd = "현재", DateSerial(9999, 12, 31), Me.txtEnd)) & _
'                    ") OR (a.end_dt BETWEEN " & SText(Me.txtStart) & " AND " & SText(IIf(Me.txtEnd = "현재", DateSerial(9999, 12, 31), Me.txtEnd)) & _
'                    ") OR (a.start_dt <= " & SText(Me.txtStart) & " AND a.end_dt >= " & SText(IIf(Me.txtEnd = "현재", DateSerial(9999, 12, 31), Me.txtEnd)) & _
'                    ")) AND a.bcleader_cd <> " & SText(.List(.ListIndex)) & ";"
            strSql = "SELECT * FROM " & TB2 & " a WHERE a.church_sid = " & SText(Me.lstChurch.List(Me.lstChurch.listIndex)) & _
                    " AND a.responsibility = '관리자'" & _
                    " AND a.bcleader_cd <> " & SText(.List(.listIndex, 0)) & _
                    " AND IF(a.start_dt > " & SText(Me.txtStart) & ", a.start_dt, " & SText(Me.txtStart) & ") <= " & _
                    " IF(a.end_dt < " & SText(IIf(Me.txtEnd = "현재", DateSerial(9999, 12, 31), Me.txtEnd)) & ", a.end_dt, " & _
                        SText(IIf(Me.txtEnd = "현재", DateSerial(9999, 12, 31), Me.txtEnd)) & ");"
            Call makeListData(strSql, TB2)
        Else
            cntRecord = 0
        End If
    End With
    
    If cntRecord > 0 Then
        MsgBox "중복된 기간은 존재할 수 없습니다. 입력 값을 다시 확인해주세요.", vbCritical, banner
        queryKey = LISTDATA(0, 0)
        Call returnListPosition(Me, Me.lstHistory.Name, queryKey)
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
    
    '--//SQL문 생성, 실행, 로그기록
    strSql = makeUpdateSQL(TB2)
    result.strSql = strSql
    connectTaskDB
    result.affectedCount = executeSQL("cmdEdit_Click", TB2, strSql, Me.Name, "관리자이력 업데이트")
    writeLog "cmdEdit_Click", TB2, strSql, 0, Me.Name, "관리자이력 업데이트", result.affectedCount
    disconnectALL
    
    '--//메세지박스
    MsgBox "업데이트 되었습니다.", , banner
    
    '--//리스트박스 초기화
    Call lstChurch_Click
    Me.lstHistory.listIndex = queryKey
    
End Sub

Private Sub cmdNew_Click()
    
    '--//커맨드 버튼 활성화를 위한 이력 리스트박스 클릭
    If lstHistory.ListCount = 0 Then
        Call lstHistory_Click
    End If
    
    'Call cmdbtn_visible
    Me.lstHistory.listIndex = Me.lstHistory.ListCount - 1
    Call sbtxtBox_Init
    Call INPUTMODE(True)
    Me.txtEnd = "현재"
    If Me.txtEnd = "현재" Then
        Me.chkPresent.Value = True
    End If
    Call HideDeleteButtonByUserAuth
End Sub

Private Sub cmdADD_Click()
    
    Dim argData As T_BC_LEADER
    Dim result As T_RESULT
    
    '--//중복체크
    With Me.lstHistory
        If Me.cboResponsibility = "관리자" Then
            strSql = "SELECT * FROM " & TB2 & " a WHERE a.church_sid = " & SText(Me.lstChurch.List(Me.lstChurch.listIndex)) & _
                    " AND a.responsibility = '관리자'" & _
                    " AND IF(a.start_dt > " & SText(Me.txtStart) & ", a.start_dt, " & SText(Me.txtStart) & ") <= " & _
                    " IF(a.end_dt < " & SText(IIf(Me.txtEnd = "현재", DateSerial(9999, 12, 31), Me.txtEnd)) & ", a.end_dt, " & _
                        SText(IIf(Me.txtEnd = "현재", DateSerial(9999, 12, 31), Me.txtEnd)) & ");"
            Call makeListData(strSql, TB2)
        Else
            cntRecord = 0
        End If
        
    End With
    
    If cntRecord > 0 And Me.lstHistory.ListCount > 0 Then
        MsgBox "중복된 기간은 존재할 수 없습니다. 입력 값을 다시 확인해주세요.", vbCritical, banner
        queryKey = LISTDATA(0, 0)
        Call returnListPosition(Me, Me.lstHistory.Name, queryKey)
        'Call cmdbtn_visible
        Call INPUTMODE(False)
        Call HideDeleteButtonByUserAuth
        Exit Sub
    End If
    Call sbClearVariant
    
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
    argData.church_sid = Me.lstChurch.List(Me.lstChurch.listIndex)
    argData.START_DT = Me.txtStart
    argData.END_DT = IIf(Me.txtEnd = "현재", DateSerial(9999, 12, 31), Me.txtEnd)
    argData.lifeNo = Me.txtLifeNo
    argData.RESPONSIBILITY = Me.cboResponsibility
    
    '--//쿼리문 실행 및 로그기록
    strSql = makeInsertSQL(TB2, argData)
    connectTaskDB
    result.strSql = strSql
    result.affectedCount = executeSQL("cmdADD_Clikc", TB2, strSql, Me.Name, "관리자이력 추가")
    writeLog "cmdADD_Click", TB2, strSql, 0, Me.Name, "관리자이력 추가", result.affectedCount
    disconnectALL
    
    '--//메세지박스
    MsgBox "추가 되었습니다.", , banner
    Call lstChurch_Click
    Me.lstHistory.listIndex = Me.lstHistory.ListCount - 1
    
    '--//버튼설정 원래대로
    If Me.chkDirect.Value = True Then
        Me.chkDirect.Value = False
        Call Direct_Mode(Me.chkDirect.Value)
    End If
    'Call cmdbtn_visible
    Call INPUTMODE(False)
    Call HideDeleteButtonByUserAuth
    
End Sub
Private Sub cmdSearch_manager_Click()
    argShow = 1
    frm_Update_BCLeader_1.Show
End Sub
Private Sub lstHistory_Click()
    
    '--//컨트롤 설정
    If Me.lstHistory.listIndex <> -1 Then
        Me.txtStart.Enabled = True
        Me.txtEnd.Enabled = True
        Me.cmdEdit.Enabled = True
        Me.cmdAdd.Enabled = True
        Me.cmdDelete.Enabled = True
        Me.chkPresent.Visible = True
        Me.chkDirect.Visible = True
        Me.cmdSearch_Manager.Enabled = True
        Me.cboResponsibility.Enabled = True
    Else
        Me.txtStart.Enabled = False
        Me.txtEnd.Enabled = False
        Me.cmdEdit.Enabled = False
        Me.cmdAdd.Enabled = False
        Me.cmdDelete.Enabled = False
        Me.chkPresent.Visible = False
        Me.chkDirect.Visible = False
        Me.cmdSearch_Manager.Enabled = False
        Me.cboResponsibility.Enabled = False
    End If
    
    '--//리스트 클릭 시 시작일, 종료일, 내용 표시
    If Me.lstHistory.listIndex <> -1 Then
        With Me.lstHistory
            Me.txtStart = .List(.listIndex, 2)
            Me.txtEnd = .List(.listIndex, 3)
            Me.txtLifeNo = .List(.listIndex, 4)
            Me.txtManager = .List(.listIndex, 5)
            Me.cboResponsibility = .List(.listIndex, 7)
        End With
    End If
    
    If Me.txtEnd = "현재" Then
        Me.chkPresent.Value = True
        Me.txtEnd.Enabled = False
    Else
        Me.chkPresent.Value = False
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

Private Sub lstChurch_Click()
    
    Dim filePath As String
    Dim FileName As String
    
    '--//컨트롤 설정
    If Me.lstChurch.listIndex <> -1 Then
        Me.lstHistory.Enabled = True
        Me.cmdNew.Enabled = True
    Else
        Me.lstHistory.Enabled = False
        Me.cmdNew.Enabled = False
    End If
    
    If Me.lstHistory.listIndex <> -1 Then
        Me.txtStart.Enabled = True
        Me.txtEnd.Enabled = True
        Me.cmdEdit.Enabled = True
        Me.cmdAdd.Enabled = True
        Me.cmdDelete.Enabled = True
        Me.chkPresent.Visible = True
        Me.chkDirect.Visible = True
        Me.cmdSearch_Manager.Enabled = True
        Me.cboResponsibility.Enabled = True
    Else
        Me.txtStart.Enabled = False
        Me.txtEnd.Enabled = False
        Me.cmdEdit.Enabled = False
        Me.cmdAdd.Enabled = False
        Me.cmdDelete.Enabled = False
        Me.chkPresent.Visible = False
        Me.chkDirect.Visible = False
        Me.cmdSearch_Manager.Enabled = False
        Me.cboResponsibility.Enabled = False
    End If
    
    '--//클릭 시 텍스트박스 초기화
    Call sbtxtBox_Init
    
    '--//이력 목록상자 조정
    With Me.lstHistory
        .ColumnCount = 7
        .ColumnHeads = False
        .ColumnWidths = "0,0,80,93,0,100,200" '관리자코드, 교회코드, 시작일, 종료일, 생명번호, 이름, 소속교회
'        .Width = 345
        .TextAlign = fmTextAlignLeft
        .Font = "굴림"
    End With
    
    '--//이력목록 데이터 채우기
    Call makeSelectSQL(TB2)
    Call makeListData(strSql, TB2)
    If cntRecord > 0 Then
        Me.lstHistory.List = LISTDATA
    Else
        Me.lstHistory.Clear
        Me.txtStart = ""
        Me.txtEnd = ""
        Me.txtManager = ""
        Me.txtLifeNo = ""
    End If
    Call sbClearVariant
    
End Sub

Private Sub lstChurch_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Un'HookListBoxScroll
End Sub

Private Sub lstChurch_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Me.lstChurch.ListCount > 0 Then
        'HookListBoxScroll Me, Me.lstChurch
    End If
End Sub

Private Sub txtChurch_Change()
    Me.txtChurch.BackColor = RGB(255, 255, 255)
End Sub

Private Sub txtEnd_Change()
    Call Date_Format(Me.txtEnd)
End Sub

Private Sub txtStart_Change()
    Call Date_Format(Me.txtStart)
End Sub

Private Sub UserForm_Initialize()
    If connIP = Empty Then setGlobalVariant '//전역변수 재설정
    If checkLogin = 0 Then f_login.Show '//로그인체크
    If checkLogin = 0 Then End '--//로그인 실패 시 프로시저 종료
    
    '--//기초설정
    Me.cmdClose.Cancel = True
    TB1 = "op_system.db_churchlist_custom" '--//교회리스트
    TB2 = "op_system.db_branchleader" '--//관리자 이력
    
    '--//권한에 따른 컨트롤 설정
    Call HideDeleteButtonByUserAuth
    
    '--//컨트롤 설정
    Me.lstChurch.Enabled = False
    Me.lstHistory.Enabled = False
    Me.txtStart.Enabled = False
    Me.txtEnd.Enabled = False
    Me.txtManager.Enabled = False
    Me.cmdEdit.Enabled = False
    Me.cmdAdd.Enabled = False
    Me.cmdAdd.Visible = False
    Me.cmdCancel.Visible = False
    Me.cmdNew.Enabled = False
    Me.txtLifeNo.Enabled = False
    Me.cmdSearch_Manager.Enabled = False
    Me.cmdDelete.Enabled = False
    Me.chkPresent.Visible = False
    Me.chkDirect.Visible = False
    
    '--//콤보박스 설정
    Me.cboResponsibility.Clear
    Me.cboResponsibility.AddItem "관리자"
    Me.cboResponsibility.AddItem "단순소속"
    
    '--//리스트박스 설정
    With Me.lstChurch
        .ColumnCount = 4
        .ColumnHeads = False
        .ColumnWidths = "0,150,50,200" '교회코드, 교회명, 교회구분, 관리교회명
'        .Width = 330
        .TextAlign = fmTextAlignLeft
        .Font = "굴림"
    End With
    
End Sub
Private Sub cmdSearch_Click()
    
    Me.lstHistory.Clear
    
    strSql = makeSelectSQL(TB1)
    Call makeListData(strSql, TB1)
    
    If cntRecord = 0 Then
        MsgBox "반환할 DB 데이터가 없습니다.", vbInformation, banner
        Call sbClearVariant
        Exit Sub
    End If
    
    Me.lstChurch.List = LISTDATA
    Call sbClearVariant
    Me.lstChurch.Enabled = True
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
        If Me.chkOld.Value = False Then
            strSql = "SELECT a.church_sid,a.church_nm,a.church_gb,b.church_nm " & _
                        "FROM " & TB1 & " a LEFT JOIN op_system.db_churchlist_custom b ON a.main_church_cd = b.church_sid " & _
                        "WHERE a.ovs_dept = " & USER_DEPT & " AND a.suspend = 0 AND (a.church_gb NOT LIKE '%M%' AND a.church_gb NOT LIKE '%H%') AND (a.church_nm LIKE '%" & Me.txtChurch & "%' OR b.church_nm LIKE '%" & Me.txtChurch & "%') ORDER BY a.sort_order;"
        Else
            strSql = "SELECT a.church_sid,a.church_nm,a.church_gb,b.church_nm " & _
                        "FROM " & TB1 & " a LEFT JOIN op_system.db_churchlist_custom b ON a.main_church_cd = b.church_sid " & _
                        "WHERE a.ovs_dept = " & USER_DEPT & " AND (a.church_gb NOT LIKE '%M%' AND a.church_gb NOT LIKE '%H%') AND (a.church_nm LIKE '%" & Me.txtChurch & "%' OR b.church_nm LIKE '%" & Me.txtChurch & "%') ORDER BY a.sort_order;"
        End If
    Case TB2
        strSql = "SELECT a.bcleader_cd,a.church_sid,a.start_dt,if(a.end_dt='9999-12-31','현재',a.end_dt),a.lifeno,concat(If(isnull(b.name_ko),a.lifeno,b.name_ko),ifnull(concat('(',left(c.Title,1),')'),'')),e.church_nm,a.responsibility " & _
                "FROM " & TB2 & " a LEFT JOIN op_system.db_pastoralstaff b " & _
                "ON a.lifeno = b.lifeno " & _
                "LEFT JOIN op_system.db_title c ON a.lifeno = c.lifeno AND (CURRENT_DATE BETWEEN c.Start_dt AND c.End_dt) " & _
                "LEFT JOIN op_system.db_transfer d ON a.lifeno = d.lifeno AND CURDATE() BETWEEN d.start_dt AND d.end_dt LEFT JOIN op_system.db_churchlist e ON d.church_sid = e.church_sid " & _
                "WHERE a.church_sid = " & SText(Me.lstChurch.List(Me.lstChurch.listIndex)) & _
                " ORDER BY a.start_dt;"
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
                    "SET a.start_dt = " & SText(Me.txtStart) & ", a.end_dt = " & SText(IIf(Me.txtEnd = "현재", DateSerial(9999, 12, 31), Me.txtEnd)) & _
                    ",a.lifeno = " & SText(Me.txtLifeNo) & ",a.responsibility = " & SText(Me.cboResponsibility) & " WHERE a.bcleader_cd = " & SText(.List(.listIndex)) & ";"
        queryKey = .listIndex
        End With
    Case Else
    End Select
    makeUpdateSQL = strSql
End Function
Private Function makeInsertSQL(ByVal tableNM As String, argData As T_BC_LEADER) As String
    
    Select Case tableNM
    Case TB1
    Case TB2
        strSql = "INSERT INTO " & TB2 & " VALUES(DEFAULT," & _
                    SText(argData.church_sid) & "," & _
                    SText(argData.START_DT) & "," & _
                    SText(argData.END_DT) & "," & _
                    SText(argData.lifeNo) & "," & _
                    SText(argData.RESPONSIBILITY) & ");"
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
            strSql = "DELETE FROM " & TB2 & " WHERE bcleader_cd = " & SText(.List(.listIndex)) & ";"
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
    Me.txtStart.Value = ""
    Me.chkPresent.Value = False
    Me.txtEnd.Value = ""
    Me.txtManager.Value = ""
    Me.txtLifeNo.Value = ""
    Me.cboResponsibility.Value = ""
End Sub
Private Function fnData_Validation()
'---------------------------------------
'유저폼 입력값에 대한 데이터 유효성 검사
'TRUE: 이상없음, FALSE: 잘못됨.
'---------------------------------------
    fnData_Validation = True '데이터가 유효하다는 가정 하에 시작
    
    '--//생명번호 데이터 입력여부 확인
    If Me.txtLifeNo = "" Then
        If Me.chkDirect = False Then
            MsgBox "관리자를 입력해 주세요.", vbCritical, banner
            Exit Function
        Else
            MsgBox "생명번호를 입력해 주세요.", vbCritical, banner
            Exit Function
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
    If Not IsDate(Me.txtStart) Then
        MsgBox "올바른 날짜 형태가 아닙니다. 시작일을 다시 확인 해주세요.", vbCritical, banner
        Set txtBox_Focus = Me.txtStart: fnData_Validation = False: Exit Function
    End If
    If Not IsDate(Me.txtEnd) And Me.txtEnd <> "현재" Then
        MsgBox "올바른 날짜 형태가 아닙니다. 종료일을 다시 확인 해주세요.", vbCritical, banner
        Set txtBox_Focus = Me.txtEnd: fnData_Validation = False: Exit Function
    End If
    
    '--//날짜 유효성 여부 검토
    If Me.txtEnd <> "현재" Then
        If CDate(Me.txtEnd) <= CDate(Me.txtStart) Then
            MsgBox "종료일은 시작일보다 작거나 같을 수 없습니다.", vbCritical, banner
            fnData_Validation = False: Exit Function
        End If
    End If
    
    '--//콤보박스 값 검토
    If Not (Me.cboResponsibility = "관리자" Or Me.cboResponsibility = "단순소속") Then
        MsgBox "관리자 혹은 단순소속 중에서 선택해 주세요.", vbCritical, banner
        Set txtBox_Focus = Me.cboResponsibility: fnData_Validation = False: Exit Function
    End If
    
    '--//지관자, 예관자 직책자만 관리자로 등록 가능
    Dim lifeNo As String
    Dim objPosition As position
    Dim objPositionDao As New PositionDao
    Dim availablePositionList As Object
    
    If Me.cboResponsibility = "관리자" And Me.chkDirect.Value = False Then
        Set availablePositionList = CreateObject("System.Collections.ArrayList")
        availablePositionList.Add "동역"
        availablePositionList.Add "지교회관리자"
        availablePositionList.Add "예배소관리자"
        
        lifeNo = Me.txtLifeNo
        Set objPosition = objPositionDao.FindPositionByLifeNoAndDate(lifeNo, Now)
        
        If objPosition Is Nothing Then Set objPosition = New position
        If Not availablePositionList.Contains(objPosition.position) Then
            MsgBox "동역, 지교회관리자, 예배소관리자 직책자만 관리자로 등록할 수 있습니다." & vbNewLine & _
                    "생명번호를 직접 입력하거나 동역 및 관리자 직책을 먼저 등록해주세요."
            fnData_Validation = False: Exit Function
        End If
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
    Me.cmdSearch_Manager.Enabled = argBoolean
    
    '--//컨트롤 활성화/비활성화
    Me.txtChurch.Enabled = Not argBoolean
    Me.cmdSearch.Enabled = Not argBoolean
    Me.lstChurch.Enabled = Not argBoolean
    Me.lstHistory.Enabled = Not argBoolean
    Me.chkOld.Enabled = Not argBoolean
    
    Me.txtStart.Enabled = argBoolean
    Me.txtEnd.Enabled = argBoolean
    Me.chkPresent.Visible = argBoolean
    Me.chkDirect.Visible = argBoolean
    Me.cboResponsibility.Enabled = argBoolean
    
    If argBoolean = True Then
        Me.cboResponsibility.listIndex = 0
    End If
End Sub

Private Sub Direct_Mode(ByVal argBoolean As Boolean)
    Me.txtManager.Visible = Not argBoolean
    Me.txtLifeNo.Enabled = argBoolean
    Me.cmdSearch_Manager.Visible = Not argBoolean
    If argBoolean Then
        Me.lblKind2 = "생명번호"
    Else
        Me.lblKind2 = "이름"
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
