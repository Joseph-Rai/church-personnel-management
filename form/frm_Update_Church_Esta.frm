VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Update_Church_Esta 
   Caption         =   "교회설립 이력관리 마법사"
   ClientHeight    =   9765.001
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   19440
   OleObjectBlob   =   "frm_Update_Church_Esta.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "frm_Update_Church_Esta"
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
Dim txtBox_Focus As MSForms.control '--//포커스가 가야하는 컨트롤 설정

Private Sub chkPresent_Click()
    Select Case Me.chkPresent.Value
        Case True '--//종료일 현재로 바꾸고 음영처리
            Me.txtEnd.BackColor = &HE0E0E0
            Me.txtEnd.Value = "현재"
            Me.txtEnd.Enabled = False
        Case False '--//종료일 음영처리 원상복구 및 오늘날짜 삽입
            Me.txtEnd.Enabled = True
            Me.txtEnd.BackColor = RGB(255, 255, 255)
            If Me.lstNomatch.listIndex = -1 Then
                Me.txtEnd = "" '--//노매치 리스트박스가 선택되어 있지 않으면
            Else
                If Me.txtEnd = "현재" Then
                    Me.txtEnd.Value = Date - 1 '--//노매치 리스트박스가 선택되어 있으면
                End If
            End If
    Case Else
    End Select
End Sub

Private Sub chkPresent_Nomatch_Click()
    Select Case Me.chkPresent_Nomatch.Value
        Case True '--//종료일 현재로 바꾸고 음영처리
            Me.txtEnd_Nomatch.BackColor = &HE0E0E0
            Me.txtEnd_Nomatch.Value = "현재"
            Me.txtEnd_Nomatch.Enabled = False
        Case False '--//종료일 음영처리 원상복구 및 오늘날짜 삽입
            Me.txtEnd_Nomatch.Enabled = True
            Me.txtEnd_Nomatch.BackColor = RGB(255, 255, 255)
            If Me.lstNomatch.listIndex = -1 Then
                Me.txtEnd_Nomatch = "" '--//노매치 리스트박스가 선택되어 있지 않으면
            Else
                With Me.lstNomatch
                    If .List(.listIndex, 5) = "" Then
                        Me.txtEnd_Nomatch.Value = Date - 1 '--//노매치 리스트박스가 선택되어 있으면
                    Else
                        Me.txtEnd_Nomatch.Value = .List(.listIndex, 5)
                    End If
                End With
            End If
    Case Else
    End Select
End Sub

Private Sub cmdADD_Click()
    
    Dim argData As T_CHURCH_ESTA
    Dim result As T_RESULT
    
    '--//노매치 교회리스트박스 미선택 시 프로시저 종료
    If Me.lstNomatch.listIndex = -1 Then
        MsgBox "매칭 하고자 하는 교회를 선택해 주세요."
        If Me.lstNomatch.ListCount > 0 Then
            Me.lstNomatch.listIndex = 0
        End If
        Exit Sub
    End If
    
    '--//교회목록 리스트박스 미선택 시 프로시저 종료
    If Me.lstChurch.listIndex = -1 Then
        MsgBox "어떤 교회와 매칭할지 선택해 주세요."
        If Me.lstChurch.ListCount > 0 Then
            Me.lstChurch.listIndex = 0
        End If
        Exit Sub
    End If
    
    '--//중복검사
    strSql = "SELECT * FROM " & TB3 & " a WHERE a.church_sid_custom = " & SText(Me.lstChurch.List(Me.lstChurch.listIndex)) & _
            " AND IF(a.start_dt > " & SText(Me.txtStart_Nomatch) & ", a.start_dt, " & SText(Me.txtStart_Nomatch) & ") <= " & _
            "IF(a.end_dt < " & SText(IIf(Me.txtEnd_Nomatch = "현재", DateSerial(9999, 12, 31), Me.txtEnd_Nomatch)) & ", a.end_dt, " & _
                SText(IIf(Me.txtEnd_Nomatch = "현재", DateSerial(9999, 12, 31), Me.txtEnd_Nomatch)) & ");"
    Call makeListData(strSql, TB2)

    If cntRecord > 0 Then
        MsgBox "중복된 기간은 존재할 수 없습니다. 시작일 혹은 종료일을 다시 확인해 주세요.", vbCritical, banner
        queryKey = LISTDATA(0, 0)
        Call returnListPosition(Me, Me.lstHistory.Name, queryKey)
        Exit Sub
    End If
    Call sbClearVariant
    
    '--//데이터 유효성 검사
    TASK_CODE = 2
    If fnData_Validation = False Then
On Error Resume Next '--//컨트롤이 비활성화 되어 에러가 나면 아래 과정 생략
        txtBox_Focus.SetFocus
        txtBox_Focus.SelStart = 0
        txtBox_Focus.SelLength = Len(txtBox_Focus)
On Error GoTo 0
        Exit Sub
    End If
    
    '--//데이터 삽입을 위해 구조체 설정
    With Me.lstChurch
        argData.CHURCH_SID_CUSTOM = .List(.listIndex)
        argData.START_DT = Me.txtStart_Nomatch
        argData.END_DT = IIf(Me.txtEnd_Nomatch = "현재", DateSerial(9999, 12, 31), Me.txtEnd_Nomatch)
        argData.church_sid = Me.lstNomatch.List(Me.lstNomatch.listIndex)
    End With
    
    '--//선택된 교회 교회설립이력DB에 추가
    connectTaskDB
    strSql = makeInsertSQL(TB3, argData)
    result.affectedCount = executeSQL("cmdAdd_Click", TB3, strSql, Me.Name, "교회 설립이력 추가")
    writeLog "cmdADD_Click", TB3, strSql, 0, Me.Name, "교회 설립이력 추가", result.affectedCount
    disconnectALL
    Call sbClearVariant
    
    '--//메세지 박스
    MsgBox "교회 설립이력이 추가 되었습니다.", , banner
    
    '--//매칭이력 리스트박스 새로고침
    Call lstChurch_Click
    
    '--//리스트박스 마지막 값 선택
    With Me.lstHistory
        .listIndex = .ListCount - 1
    End With
    
    '--//노매칭 리스트박스 새로고침
    Call cmdSearch_Nomatch_Click
    
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    
    Dim result As T_RESULT
    
    '--//교회 설립이력 리스트박스 선택여부 확인
    If Me.lstHistory.listIndex = -1 Then
        MsgBox "설립이력을 신규생성 하고자 하는 교회를 선택해 주세요."
        If Me.lstHistory.ListCount > 0 Then
            Me.lstHistory.listIndex = 0
        End If
        Exit Sub
    End If
    
    '--//정말 삭제할 것인지 재확인
    If MsgBox("정말 삭제하시겠습니까?", vbQuestion + vbYesNo, banner) = vbNo Then
        Exit Sub
    End If
    
    '--//삭제 쿼리문 실행
    connectTaskDB
    strSql = makeDeleteSQL(TB3)
    result.affectedCount = executeSQL("cmdDelete_Click", TB3, strSql, Me.Name, "교회 설립이력 삭제")
    writeLog "cmdDelete_Click", TB3, strSql, 0, Me.Name, "교회 설립이력 삭제", result.affectedCount
    disconnectALL
    Call sbClearVariant
    
    '--//메세지박스
    MsgBox "선택하신 데이터가 삭제 되었습니다.", , banner
    
    '--//노매칭 리스트박스 새로고침
    With Me.lstNomatch
        queryKey = .listIndex
    End With
    Call cmdSearch_Nomatch_Click
    
    '--//노매칭 리스트박스 기존 선택되어 있던 곳 선택
    Me.lstNomatch.listIndex = queryKey
    
    If Me.lstHistory.ListCount = 1 Then '--//완전삭제일 경우
        Me.lstHistory.Clear
        Me.lstChurch.Clear
        Me.txtChurch = ""
        Me.txtStart = ""
        Me.txtEnd = ""
    Else '--//안전 삭제가 아닐 경우
        '--//설립이력 리스트박스 새로고침
        Call lstChurch_Click
    End If

End Sub

Private Sub cmdEdit_Click()
    
    Dim result As T_RESULT
    
    '--//교회 설립이력 리스트박스 선택여부 확인
    If Me.lstHistory.listIndex = -1 Then
        MsgBox "수정 하고자 하는 교회를 선택해 주세요."
        If Me.lstHistory.ListCount > 0 Then
            Me.lstHistory.listIndex = 0
        End If
        Exit Sub
    End If
    
    '--//중복검사
    With Me.lstHistory
        strSql = "SELECT * FROM " & TB3 & " a WHERE a.church_sid_custom = " & SText(Me.lstChurch.List(Me.lstChurch.listIndex)) & _
                " AND IF(a.start_dt > " & SText(Me.txtStart) & ", a.start_dt, " & SText(Me.txtStart) & ") <= " & _
                "IF(a.end_dt < " & SText(IIf(Me.txtEnd = "현재", DateSerial(9999, 12, 31), Me.txtEnd)) & ", a.end_dt, " & _
                    SText(IIf(Me.txtEnd = "현재", DateSerial(9999, 12, 31), Me.txtEnd)) & ")" & _
                " AND a.church_sid <> " & SText(.List(.listIndex, 4)) & ";"
    End With
    Call makeListData(strSql, TB3)

    If cntRecord > 0 Then
        MsgBox "중복된 기간은 존재할 수 없습니다. 시작일 혹은 종료일을 다시 확인해 주세요.", vbCritical, banner
        queryKey = LISTDATA(0, 0)
        Call returnListPosition(Me, Me.lstHistory.Name, queryKey)
        Exit Sub
    End If
    Call sbClearVariant
    
    '--//데이터 유효성 검사
    TASK_CODE = 3
    If fnData_Validation = False Then
On Error Resume Next '--//컨트롤이 비활성화 되어 에러가 나면 아래 과정 생략
        txtBox_Focus.SetFocus
        txtBox_Focus.SelStart = 0
        txtBox_Focus.SelLength = Len(txtBox_Focus)
On Error GoTo 0
        Exit Sub
    End If
    
    '--//수정 쿼리문 실행
    connectTaskDB
    strSql = makeUpdateSQL(TB3)
    result.affectedCount = executeSQL("cmdEdit_Click", TB3, strSql, Me.Name, "교회 설립이력 수정")
    writeLog "cmdEdit_Click", TB3, strSql, 0, Me.Name, "교회 설립이력 수정", result.affectedCount
    disconnectALL
    Call sbClearVariant
    
    '--//메세지박스
    MsgBox "내용이 수정되었습니다.", , banner
    
    '--//매칭이력 리스트박스 새로고침
    Call lstChurch_Click
    
End Sub

Private Sub cmdNew_Click()
    
    Dim result As T_RESULT
    Dim argData As T_CHURCH_ESTA
    
    '--//노매치 교회리스트박스 미선택 시 프로시저 종료
    If Me.lstNomatch.listIndex = -1 Then
        MsgBox "설립이력을 신규생성 하고자 하는 교회를 선택해 주세요."
        If Me.lstNomatch.ListCount > 0 Then
            Me.lstNomatch.listIndex = 0
        End If
        Exit Sub
    End If
    
    '--//데이터 유효성 검사
    TASK_CODE = 1
    If fnData_Validation = False Then
On Error Resume Next '--//컨트롤이 비활성화 되어 에러가 나면 아래 과정 생략
        txtBox_Focus.SetFocus
        txtBox_Focus.SelStart = 0
        txtBox_Focus.SelLength = Len(txtBox_Focus)
On Error GoTo 0
        Exit Sub
    End If
    
    '--//데이터 삽입을 위해 구조체 설정
    strSql = "SELECT max(a.`교회커스텀코드`) FROM op_system.v_churchlist_final a;"
    Call makeListData(strSql, TB3)
    With Me.lstChurch
        argData.CHURCH_SID_CUSTOM = LISTDATA(0, 0) + 1
        argData.START_DT = Me.txtStart_Nomatch
        argData.END_DT = IIf(Me.txtEnd_Nomatch = "현재", DateSerial(9999, 12, 31), Me.txtEnd_Nomatch)
        argData.church_sid = Me.lstNomatch.List(Me.lstNomatch.listIndex)
    End With
    sbClearVariant
    
    '--//선택된 교회 교회설립이력DB에 삽입(커스텀코드 신규생성)
    connectTaskDB
    strSql = makeInsertSQL(TB3, argData)
    result.affectedCount = executeSQL("cmdNew_Click", TB3, strSql, Me.Name, "교회 설립이력 추가")
    writeLog "cmdADD_Click", TB3, strSql, 0, Me.Name, "교회 설립이력 추가", result.affectedCount
    disconnectALL
    Call sbClearVariant
    
    '--//메세지박스
    MsgBox "교회 설립 이력이 신규 추가 되었습니다.", , banner
    
    '--//txtChurchNM에 신규 생성한 교회명 삽입
    With Me.lstNomatch
        Me.txtChurch = .List(.listIndex, 1)
    End With
    
    '--//매칭 교회리스트 새로고침
    Call cmdSearch_Click
    
    '--//매칭이력 리스트박스 새로고침
    Me.lstChurch.listIndex = 0
    Call lstChurch_Click
    
    '--//리스트박스 마지막 값 선택
    With Me.lstHistory
        .listIndex = .ListCount - 1
    End With
    
    '--//노매칭 리스트박스 새로고침
    With Me.lstNomatch
        queryKey = .listIndex
    End With
    Call cmdSearch_Nomatch_Click
    If queryKey >= Me.lstNomatch.ListCount Then
        Me.lstNomatch.listIndex = Me.lstNomatch.ListCount - 1
    Else
        Me.lstNomatch.listIndex = queryKey
    End If
    
End Sub

Private Sub cmdSearch_Click()
    '--//매칭 교회리스트 리스트박스 새로고침
    connectTaskDB
    strSql = makeSelectSQL(TB2)
    Call makeListData(strSql, TB2)
    If cntRecord > 0 Then
        Me.lstChurch.List = LISTDATA
    End If
    disconnectALL
    sbClearVariant
End Sub

Private Sub cmdSearch_Nomatch_Click()
    '--//노매치 교회리스트 리스트박스 새로고침
    connectTaskDB
    strSql = makeSelectSQL(TB1)
    Call makeListData(strSql, TB1)
    If cntRecord > 0 Then
        Me.lstNomatch.List = LISTDATA
    Else
        Me.lstNomatch.Clear
    End If
    disconnectALL
    sbClearVariant
End Sub

Private Sub lstChurch_Click()
    '--//신규버튼 활성화
    If Me.lstNomatch.listIndex <> -1 Then
        Me.cmdNew.Enabled = True
        Me.cmdAdd.Enabled = True
    End If
    
    '--//매칭이력 리스트박스 새로고침
    If Me.lstChurch.listIndex <> -1 Then
        connectTaskDB
        strSql = makeSelectSQL(TB3)
        Call makeListData(strSql, TB3)
        If cntRecord > 0 Then
            Me.lstHistory.List = LISTDATA
        End If
        disconnectALL
        sbClearVariant
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

Private Sub lstHistory_Click()
    '--//시작일, 종료일 텍스트박스 활성화
    Me.txtStart.Enabled = True
    Me.txtStart.BackColor = vbWhite
    Me.txtEnd.Enabled = True
    Me.txtEnd.BackColor = vbWhite
    
    '--//수정버튼 활성화
    Me.cmdEdit.Enabled = True
    
    '--//추가,삭제 버튼 활성화
    Me.cmdAdd.Enabled = True
    Me.cmdDelete.Enabled = True
    
    '--//시작일, 종료일에 내용채우기
    If Me.lstHistory.ListCount > 0 Then
        With Me.lstHistory
            Me.txtStart = .List(.listIndex, 2)
            Me.txtEnd = .List(.listIndex, 3)
        End With
    End If
    
    '--//종료일이 9999-12-31이면?
    If Me.txtEnd = "현재" Then
        Me.chkPresent.Value = True
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

Private Sub lstNomatch_Click()
    
    '--//신규, 추가버튼 활성화
    If Me.lstChurch.listIndex <> -1 Then
        Me.cmdAdd.Enabled = True
    End If
    Me.cmdNew.Enabled = True
    
    '--//시작일, 종료일에 내용채우기
    If Me.lstNomatch.ListCount > 0 Then
        With Me.lstNomatch
            Me.txtStart_Nomatch = .List(.listIndex, 4)
            Me.txtEnd_Nomatch = .List(.listIndex, 5)
            If Me.txtEnd_Nomatch = "현재" Then
                Me.chkPresent_Nomatch.Value = True
                Me.txtEnd_Nomatch.Enabled = False
            Else
                Me.chkPresent_Nomatch.Value = False
            End If
        End With
    End If
    
    '--//시작일, 종료일 텍스트박스 활성화
    Me.txtStart_Nomatch.Enabled = True
    Me.txtStart_Nomatch.BackColor = vbWhite
    If Me.txtEnd_Nomatch <> "현재" Then
        Me.txtEnd_Nomatch.Enabled = True
        Me.txtEnd_Nomatch.BackColor = vbWhite
    End If
    
End Sub

Private Sub lstNomatch_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Un'HookListBoxScroll
End Sub

Private Sub lstNomatch_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Me.lstNomatch.ListCount > 0 Then
        'HookListBoxScroll Me, Me.lstNomatch
    End If
End Sub

Private Sub txtChurch_Change()
    Me.txtChurch.BackColor = RGB(255, 255, 255)
End Sub

Private Sub txtChurch_Nomatch_Change()
    Me.txtChurch_Nomatch.BackColor = RGB(255, 255, 255)
End Sub

Private Sub txtEnd_Change()
    Call Date_Format(Me.txtEnd)
End Sub

Private Sub txtEnd_Nomatch_Change()
    Call Date_Format(Me.txtEnd_Nomatch)
End Sub

Private Sub txtStart_Change()
    Call Date_Format(Me.txtStart)
End Sub

Private Sub txtStart_Nomatch_Change()
    Call Date_Format(Me.txtStart_Nomatch)
End Sub

Private Sub UserForm_Initialize()
    If connIP = Empty Then setGlobalVariant '//전역변수 재설정
    If checkLogin = 0 Then f_login.Show '//로그인체크
    If checkLogin = 0 Then End '--//로그인 실패 시 프로시저 종료
    
    '--//기초설정
    Me.cmdClose.Cancel = True
    TB1 = "op_system.v_churchlist_nomatch" '--//노매치 교회리스트
    TB2 = "op_system.v_churchlist_final" '--//매칭완료 교회리스트
    TB3 = "op_system.db_history_church_establish" '--//매칭이력
    
    '--//컨트롤 설정
    Me.cmdNew.Enabled = False
    Me.cmdAdd.Enabled = False
    Me.cmdDelete.Enabled = False
    Me.txtStart.Enabled = False
    Me.txtStart.BackColor = &HE0E0E0
    Me.txtEnd.Enabled = False
    Me.txtEnd.BackColor = &HE0E0E0
    Me.txtStart_Nomatch.Enabled = False
    Me.txtStart_Nomatch.BackColor = &HE0E0E0
    Me.txtEnd_Nomatch.Enabled = False
    Me.txtEnd_Nomatch.BackColor = &HE0E0E0
    
    '--//리스트박스 설정
    With Me.lstNomatch
        .ColumnCount = 6
        .ColumnHeads = False
        .ColumnWidths = "0,150,50,150,70,70" '교회코드, 교회명, 교회구분, 관리교회명, 시작일, 종료일
        .Width = 531
        .TextAlign = fmTextAlignLeft
        .Font = "굴림"
    End With
    With Me.lstChurch
        .ColumnCount = 7
        .ColumnHeads = False
        .ColumnWidths = "0,0,120,40,0,100,70" '커스텀교회코드.교회코드,교회명(ko),교회구분,본교회코드,본교회명,시작일
        .Width = 344
        .TextAlign = fmTextAlignLeft
        .Font = "굴림"
    End With
    With Me.lstHistory
        .ColumnCount = 11
        .ColumnHeads = False
        .ColumnWidths = "0,0,70,70,0,130,20" 'DBKey값, 커스텀교회코드,시작일,종료일,교회코드,교회명
        .Width = 344
        .TextAlign = fmTextAlignLeft
        .Font = "굴림"
    End With
    
    '--//노매치 교회리스트 새로고침
    Call cmdSearch_Nomatch_Click
    
    '--//노매치 교회검색 포커스
    Me.txtChurch.SetFocus

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
        strSql = "SELECT * FROM " & TB1 & " a WHERE a.ovs_dept = " & SText(USER_DEPT) & " AND (a.church_nm LIKE '%" & Me.txtChurch_Nomatch & "%' OR a.main_church LIKE '%" & Me.txtChurch_Nomatch & "%');"
    Case TB2
        strSql = "SELECT a.`교회커스텀코드`,`교회코드`,a.`교회명(ko)`,a.`교회구분`,a.`본교회코드`,a.`본교회명`,b.`start_dt` " & _
                    "FROM " & TB2 & " a " & _
                    "LEFT JOIN " & TB3 & " b ON a.`교회코드` = b.church_sid " & _
                    "WHERE a.`관리부서` = " & SText(USER_DEPT) & " AND (a.`교회명(ko)` LIKE '%" & Me.txtChurch & "%' OR a.`교회명(en)` LIKE '%" & Me.txtChurch & "%' OR a.`본교회명` = '%" & Me.txtChurch & "%');"
    Case TB3
        With Me.lstChurch
            strSql = "SELECT a.church_esta_cd,a.church_sid_custom,a.start_dt,replace(a.end_dt,'9999-12-31','현재'),a.church_sid,b.church_nm,b.church_gb FROM " & TB3 & _
                        " a LEFT JOIN op_system.db_churchlist b ON a.church_sid = b.church_sid WHERE a.church_sid_custom = " & SText(.List(.listIndex)) & "ORDER BY a.start_dt;"
        End With
    Case Else
    End Select
    makeSelectSQL = strSql
End Function
Private Function makeUpdateSQL(ByVal tableNM As String) As String
    
    With Me.lstHistory
        strSql = "UPDATE " & TB3 & " a " & _
                    "SET a.start_dt = " & SText(Me.txtStart) & ", a.end_dt = " & SText(IIf(Me.txtEnd = "현재", DateSerial(9999, 12, 31), Me.txtEnd)) & _
                    " WHERE a.church_esta_cd = " & SText(.List(.listIndex)) & ";"
    End With
        
    makeUpdateSQL = strSql
End Function
Private Function makeInsertSQL(ByVal tableNM As String, _
                                argData As T_CHURCH_ESTA) As String
    strSql = "INSERT INTO " & TB3 & " VALUES(DEFAULT," & _
                SText(argData.CHURCH_SID_CUSTOM) & "," & _
                SText(argData.START_DT) & "," & _
                SText(argData.END_DT) & "," & _
                SText(argData.church_sid) & ");"
    makeInsertSQL = strSql
End Function
Private Function makeDeleteSQL(ByVal tableNM As String) As String
    With Me.lstHistory
        strSql = "DELETE FROM " & TB3 & " WHERE church_esta_cd = " & SText(.List(.listIndex)) & ";"
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
    
    Select Case TASK_CODE
        Case 1, 2 '--//신규,추가일 때
            If Not IsDate(Me.txtStart_Nomatch) Then
                MsgBox "날짜 형식이 잘못 되었습니다. 시작일을 다시 확인해 주세요.", vbCritical, banner
                Set txtBox_Focus = Me.txtStart_Nomatch: fnData_Validation = False: Exit Function
            End If
            If Not IsDate(Me.txtEnd_Nomatch) And Me.txtEnd_Nomatch <> "현재" Then
                MsgBox "날짜 형식이 잘못 되었습니다. 종료일일을 다시 확인해 주세요.", vbCritical, banner
                Set txtBox_Focus = Me.txtEnd_Nomatch: fnData_Validation = False: Exit Function
            End If
        
        Case 3  '--//설립이력 수정일 때일 때
            If Not IsDate(Me.txtStart) Then
                MsgBox "날짜 형식이 잘못 되었습니다. 시작일을 다시 확인해 주세요.", vbCritical, banner
                Set txtBox_Focus = Me.txtStart: fnData_Validation = False: Exit Function
            End If
            If Not IsDate(Me.txtEnd) And Me.txtEnd <> "현재" Then
                MsgBox "날짜 형식이 잘못 되었습니다. 종료일일을 다시 확인해 주세요.", vbCritical, banner
                Set txtBox_Focus = Me.txtEnd: fnData_Validation = False: Exit Function
            End If
    End Select
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
