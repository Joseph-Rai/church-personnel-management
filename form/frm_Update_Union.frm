VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Update_Union 
   Caption         =   "연합회 관리마법사"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5355
   OleObjectBlob   =   "frm_Update_Union.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "frm_Update_Union"
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

Private Sub cboUnion_Change()
    If Me.cboUnion <> "" And Me.cboUnion.listIndex <> -1 Then
        strSql = "SELECT * FROM op_system.a_union a WHERE a.union_nm = " & SText(Me.cboUnion) & ";"
        Call makeListData(strSql, "op_system.a_union")
        Me.txtUnion_cd = LISTDATA(0, 0)
    End If
End Sub

Private Sub cboUnion_Enter()
    '--//콤보박스 아이템 추가
    strSql = "SELECT a.union_nm FROM op_system.a_union a WHERE a.suspend = 0 AND a.ovs_dept = " & SText(USER_DEPT) & " ORDER BY a.sort_order;"
    Call makeListData(strSql, "op_system.a_union")
'    Me.cboUnion.Clear
    If cntRecord > 0 Then
        Me.cboUnion.List = LISTDATA
    Else
        Me.cboUnion.Clear
    End If
End Sub

Private Sub cboUnion_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Un'HookListBoxScroll
End Sub

Private Sub cboUnion_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Me.cboUnion.ListCount > 0 Then
        'HookListBoxScroll Me, Me.cboUnion
    End If
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
                Me.txtEnd = ""
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
    Call INPUTMODE(False)
    Call HideDeleteButtonByUserAuth
    Call sbtxtBox_Init
    Me.txtEnd = ""
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
    result.affectedCount = executeSQL("cmdDelete_Click", TB2, strSql, Me.Name, "연합회 이력 삭제")
    writeLog "cmdDelete_Click", TB2, strSql, 0, Me.Name, "연합회 이력 삭제"
    disconnectALL
    
    '--//메세지박스
    MsgBox "해당 데이터가 삭제되었습니다.", , banner
    
    '--//리스트박스 새로고침
    Call lstChurch_Click
    Me.lstHistory.listIndex = -1
    
End Sub

Private Sub cmdEdit_Click()
    
    Dim result As T_RESULT
    
    '--//수정된 내용 있는지 체크
    With Me.lstHistory
        If Me.cboUnion = .List(.listIndex, 6) And Me.txtStart = .List(.listIndex, 3) And Me.txtEnd = .List(.listIndex, 4) Then
            Exit Sub
        End If
    End With
    
    '--//중복체크
    With Me.lstHistory
        strSql = "SELECT * FROM " & TB2 & " a WHERE a.church_sid_custom = " & SText(.List(.listIndex, 1)) & _
                " AND ((a.start_dt BETWEEN " & SText(Me.txtStart) & " AND " & SText(IIf(Me.txtEnd = "현재", DateSerial(9999, 12, 31), Me.txtEnd)) & _
                ") OR (a.end_dt BETWEEN " & SText(Me.txtStart) & " AND " & SText(IIf(Me.txtEnd = "현재", DateSerial(9999, 12, 31), Me.txtEnd)) & _
                ") OR (a.start_dt <= " & SText(Me.txtStart) & " AND a.end_dt >= " & SText(IIf(Me.txtEnd = "현재", DateSerial(9999, 12, 31), Me.txtEnd)) & _
                ")) AND a.union_cd <> " & SText(.List(.listIndex)) & ";"
        Call makeListData(strSql, TB2)
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
    result.affectedCount = executeSQL("cmdEdit_Click", TB2, strSql, Me.Name, "연합회 이력 업데이트")
    writeLog "cmdEdit_Click", TB2, strSql, 0, Me.Name, "연합회 이력 업데이트", result.affectedCount
    disconnectALL
    
    '--//메세지박스
    MsgBox "업데이트 되었습니다.", , banner
    
    '--//리스트박스 초기화
    Call lstChurch_Click
    Call lstHistory_Click
    Me.lstHistory.listIndex = queryKey
    
End Sub

Private Sub cmdNew_Click()
    Call lstHistory_Click
    'Call cmdbtn_visible
    Call INPUTMODE(True)
    Call HideDeleteButtonByUserAuth
    Me.lstHistory.listIndex = Me.lstHistory.ListCount - 1
    Call sbtxtBox_Init
    Me.chkPresent.Value = True
    Me.txtEnd.Enabled = False
End Sub

Private Sub cmdADD_Click()
    
    Dim argData As T_UNION
    Dim result As T_RESULT
    
    '--//중복체크
    
    strSql = "SELECT * FROM " & TB2 & " a WHERE a.church_sid_custom = " & SText(Me.lstChurch.List(Me.lstChurch.listIndex, 1)) & _
            " AND ((a.start_dt BETWEEN " & SText(Me.txtStart) & " AND " & SText(IIf(Me.txtEnd = "현재", DateSerial(9999, 12, 31), Me.txtEnd)) & _
            ") OR (a.end_dt BETWEEN " & SText(Me.txtStart) & " AND " & SText(IIf(Me.txtEnd = "현재", DateSerial(9999, 12, 31), Me.txtEnd)) & _
            ") OR (a.start_dt <= " & SText(Me.txtStart) & " AND a.end_dt >= " & SText(IIf(Me.txtEnd = "현재", DateSerial(9999, 12, 31), Me.txtEnd)) & "));"
    Call makeListData(strSql, TB2)
   
    
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
    
    '--//작업에 따라 구조체에 값 추가
    With Me.lstHistory
        argData.CHURCH_SID_CUSTOM = Me.lstChurch.List(Me.lstChurch.listIndex)
        argData.START_DT = Me.txtStart
        argData.END_DT = IIf(Me.txtEnd = "현재", DateSerial(9999, 12, 31), Me.txtEnd)
        argData.UNION = Me.txtUnion_cd
    End With
    
    '--//작업에 따라 쿼리문 실행 및 로그기록
    strSql = makeInsertSQL(TB2, argData)
    connectTaskDB
    result.strSql = strSql
    result.affectedCount = executeSQL("cmdADD_Clikc", TB2, strSql, Me.Name, "연합회 이력 추가")
    writeLog "cmdADD_Click", TB2, strSql, 0, Me.Name, "연합회 이력 추가", result.affectedCount
    disconnectALL
    
    '--//메세지박스
    MsgBox "추가 되었습니다.", , banner
    Call lstChurch_Click
    Me.lstHistory.listIndex = Me.lstHistory.ListCount - 1
    
    '--//버튼설정 원래대로
    'Call cmdbtn_visible
    Call INPUTMODE(False)
    Call HideDeleteButtonByUserAuth
    Me.cmdCancel.Visible = False
    
    '--//리스트박스 초기화
    Call lstChurch_Click
    Call lstHistory_Click
    
End Sub

Private Sub cmdUnion_Click()
    Call frm_Update_Union_1_Show
End Sub

Private Sub lstHistory_Click()
    
    '--//컨트롤 설정
    If Me.lstHistory.listIndex <> -1 Then
        Me.txtStart.Enabled = True
        Me.txtEnd.Enabled = True
        Me.cboUnion.Enabled = True
        Me.cmdEdit.Enabled = True
        Me.cmdAdd.Enabled = True
        Me.cmdDelete.Enabled = True
        Me.chkPresent.Visible = True
        Me.cmdUnion.Enabled = True
    Else
        Me.txtStart.Enabled = False
        Me.txtEnd.Enabled = False
        Me.cboUnion.Enabled = False
        Me.cmdEdit.Enabled = False
        Me.cmdAdd.Enabled = False
        Me.cmdDelete.Enabled = False
        Me.chkPresent.Visible = False
        Me.cmdUnion.Enabled = False
    End If
    
    '--//리스트 클릭 시 시작일, 종료일, 내용 표시
    With Me.lstHistory
        If .ListCount > 0 And .listIndex <> -1 Then
            Me.cboUnion = .List(.listIndex, 6)
            Me.txtStart = .List(.listIndex, 3)
            Me.txtEnd = IIf(.List(.listIndex, 4) = "9999-12-31", "현재", .List(.listIndex, 4))
            Me.txtUnion_cd = .List(.listIndex, 5)
        End If
    End With
    
    
    If Me.txtEnd = "현재" Then
        Me.chkPresent.Value = True
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
    
    Call UserForm_Initialize
    
    If Me.lstChurch.listIndex <> -1 Then
        Me.lstHistory.Enabled = True
        Me.cmdNew.Enabled = True
    Else
        Me.lstHistory.Enabled = False
        Me.cmdNew.Enabled = False
    End If
    
    '--//이력 목록상자 조정
    With Me.lstHistory
        .ColumnCount = 8
        .ColumnHeads = False
        .ColumnWidths = "0,0,0,65,65,0,200" '연합회 이력코드, 교회코드, 교회명, 시작일, 종료일, 연합회코드, 연합회명
        .Width = 250
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
    End If
    Call sbClearVariant
    
    '--//텍스트박스 초기화
    Call sbtxtBox_Init
    Me.txtEnd.Value = ""
    Me.chkPresent.Value = False
    Me.chkPresent.Visible = False
    
    '--//이력 리스트박스가 비어있지 않으면 마지막 데이터 클릭
    If Me.lstHistory.ListCount > 0 Then
        Me.lstHistory.listIndex = Me.lstHistory.ListCount - 1
    End If
    
End Sub

Private Sub lstChurch_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Un'HookListBoxScroll
End Sub

Private Sub lstChurch_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    'HookListBoxScroll Me, Me.lstChurch
End Sub

Private Sub txtChurchNM_Change()
    Me.txtChurchNM.BackColor = RGB(255, 255, 255)
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
    TB1 = "op_system.v_churchlist_final" '--//선지자정보
    TB2 = "op_system.db_union" '--//예비생도 이력
    
    '--//권한에 따른 컨트롤 설정
    Call HideDeleteButtonByUserAuth
    
    '--//컨트롤 설정
    Me.lstHistory.Enabled = False
    Me.txtStart.Enabled = False
    Me.txtEnd.Enabled = False
    Me.cboUnion.Enabled = False
    Me.txtUnion_cd.Enabled = False
    Me.cmdEdit.Enabled = False
    Me.cmdAdd.Enabled = False
    Me.cmdAdd.Visible = False
    Me.cmdCancel.Visible = False
    Me.cmdNew.Enabled = False
    Me.cmdDelete.Enabled = False
    Me.chkPresent.Visible = False
    Me.cmdUnion.Enabled = False
    
    '--//콤보박스 아이템 추가
    strSql = "SELECT a.union_nm FROM op_system.a_union a WHERE a.suspend = 0 AND a.ovs_dept = " & SText(USER_DEPT) & " ORDER BY a.sort_order;"
    Call makeListData(strSql, "op_system.a_union")
    If cntRecord > 0 Then
        Me.cboUnion.List = LISTDATA
    Else
        Me.cboUnion.Clear
    End If
    
    '--//리스트박스 설정
    If Me.lstChurch.listIndex < 0 Then
        With Me.lstChurch
            .ColumnCount = 4
            .ColumnHeads = False
            .ColumnWidths = "0,150,200" '커스텀 교회코드, 한글교회명, 영문교회명
            .TextAlign = fmTextAlignLeft
            .Font = "굴림"
        End With
    End If
'    Me.Width = 270
    If Me.txtChurchNM.Enabled = True Then
        Me.txtChurchNM.SetFocus
    End If
    
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
        If Me.chkAll Then
            strSql = "SELECT a.`교회커스텀코드`,a.`교회명(ko)`,a.`교회명(en)` " & _
                    "FROM " & TB1 & " a " & _
                    "WHERE (a.`교회명(ko)` LIKE '%" & Me.txtChurchNM & "%' OR a.`교회명(en)` LIKE '%" & Me.txtChurchNM & "%') " & _
                    "AND a.`교회구분` in ('MC','HBC') AND a.`관리부서` = " & SText(USER_DEPT) & ";"
        Else
            strSql = "SELECT a.`교회커스텀코드`,a.`교회명(ko)`,a.`교회명(en)` " & _
                        "FROM " & TB1 & " a " & _
                        "WHERE (a.`교회명(ko)` LIKE '%" & Me.txtChurchNM & "%' OR a.`교회명(en)` LIKE '%" & Me.txtChurchNM & "%') " & _
                        "AND a.`교회구분` in ('MC','HBC') AND a.`논리삭제` = 0 AND a.`관리부서` = " & SText(USER_DEPT) & ";"
        End If
    Case TB2
        With Me.lstChurch
            strSql = "SELECT a.union_cd,a.church_sid_custom,b.`교회명(ko)`,a.start_dt,if(a.end_dt='9999-12-31','현재',a.end_dt),a.`union`,c.union_nm " & _
                    "FROM " & TB2 & " a " & _
                    "LEFT JOIN op_system.v_churchlist_final b ON a.church_sid_custom = b.`교회커스텀코드` AND b.`교회구분` IN ('MC','HBC') " & _
                    "LEFT JOIN op_system.a_union c ON a.union = c.union_cd " & _
                    "WHERE (b.`교회명(ko)` = " & SText(.List(.listIndex, 1)) & " OR b.`교회명(en)` = " & SText(.List(.listIndex, 1)) & ") " & _
                    "AND a.church_sid_custom = " & SText(Me.lstChurch.List(Me.lstChurch.listIndex)) & ";"
        End With
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
                    "SET a.start_dt = " & SText(Me.txtStart) & _
                    ", a.end_dt = " & SText(IIf(Me.txtEnd = "현재", DateSerial(9999, 12, 31), Me.txtEnd)) & _
                    ", a.union = " & SText(Me.txtUnion_cd) & _
                    " WHERE a.union_cd = " & SText(.List(.listIndex)) & ";"
        queryKey = .listIndex
        End With
    Case Else
    End Select
    makeUpdateSQL = strSql
End Function
Private Function makeInsertSQL(ByVal tableNM As String, argData As T_UNION) As String
    
    Select Case tableNM
    Case TB1
    Case TB2
        strSql = "INSERT INTO " & TB2 & " VALUES(DEFAULT," & _
                    SText(argData.CHURCH_SID_CUSTOM) & "," & _
                    SText(argData.START_DT) & "," & _
                    SText(argData.END_DT) & "," & _
                    SText(argData.UNION) & ");"
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
            strSql = "DELETE FROM " & TB2 & " WHERE union_cd = " & SText(.List(.listIndex)) & ";"
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
    Me.cboUnion = ""
    Me.txtStart.Value = ""
    Me.txtEnd.Value = "현재"
    Me.txtUnion_cd = ""
End Sub
Private Function fnData_Validation()
'---------------------------------------
'유저폼 입력값에 대한 데이터 유효성 검사
'TRUE: 이상없음, FALSE: 잘못됨.
'---------------------------------------
    fnData_Validation = True '데이터가 유효하다는 가정 하에 시작
    
    strSql = "SELECT a.union_nm FROM op_system.a_union a WHERE a.suspend = 0;"
    Call makeListData(strSql, "op_system.a_union")
    
    If IsInArray(Me.cboUnion, LISTDATA, True, rtnValue) = -1 Then
        MsgBox "연합회를 잘못 입력하였습니다. 다시 확인 해주세요.", vbCritical, banner
        Set txtBox_Focus = Me.cboUnion: fnData_Validation = False: Exit Function
    End If
    
    If Not IsDate(Me.txtStart) Then
        MsgBox "날짜 형식이 잘못 되었습니다. 시작일을 다시 확인해 주세요.", vbCritical, banner
        Set txtBox_Focus = Me.txtStart: fnData_Validation = False: Exit Function
    End If
    
    If Not IsDate(Me.txtEnd) And Me.txtEnd <> "현재" Then
        MsgBox "날짜 형식이 잘못 되었습니다. 종료일을 다시 확인해 주세요.", vbCritical, banner
        Set txtBox_Focus = Me.txtEnd: fnData_Validation = False: Exit Function
    End If
    
    If Me.txtEnd <> "현재" Then
        If CDate(Me.txtEnd) <= CDate(Me.txtStart) Then
            MsgBox "종료일은 시작일보다 작거나 같을 수 없습니다.", vbCritical, banner
            Set txtBox_Focus = Me.txtStart: fnData_Validation = False: Exit Function
        End If
    End If
    
    If Me.txtUnion_cd = "" Or Me.txtStart = "" Or Me.txtEnd = "" Then
        MsgBox "필수 입력값이 누락되었습니다. 다시 확인해주세요.", vbCritical, banner
        If Me.txtUnion_cd = "" Then Set txtBox_Focus = Me.cboUnion: fnData_Validation = False: Exit Function
        If Me.txtStart = "" Then Set txtBox_Focus = Me.txtStart: fnData_Validation = False: Exit Function
        If Me.txtEnd = "" Then Set txtBox_Focus = Me.txtEnd: fnData_Validation = False: Exit Function
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
    Call sbtxtBox_Init
    Me.cmdNew.Visible = Not argBoolean
    Me.cmdEdit.Visible = Not argBoolean
    Me.cmdDelete.Visible = Not argBoolean
    Me.cmdClose.Visible = Not argBoolean
    Me.cmdCancel.Visible = argBoolean
    Me.cmdCancel.Enabled = argBoolean
    Me.cmdAdd.Visible = argBoolean
    Me.cmdAdd.Enabled = argBoolean
    
    Me.chkPresent.Value = argBoolean
    
    Me.lstChurch.Enabled = Not argBoolean
    Me.lstHistory.Enabled = Not argBoolean
    Me.txtChurchNM.Enabled = Not argBoolean
    Me.cmdSearch.Enabled = Not argBoolean
    Me.chkAll.Enabled = Not argBoolean
    
    Me.txtStart.Enabled = argBoolean
    Me.txtEnd.Enabled = argBoolean
    Me.cboUnion.Enabled = argBoolean
    Me.cmdUnion.Enabled = argBoolean
    Me.chkPresent.Visible = argBoolean
    Me.cmdAdd.Enabled = argBoolean
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



