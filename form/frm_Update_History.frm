VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Update_History 
   Caption         =   "교회이력 관리마법사"
   ClientHeight    =   6750
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   7380
   OleObjectBlob   =   "frm_Update_History.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "frm_Update_History"
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
Dim txtBox_Focus As MSForms.textBox

Private Sub lstChurch_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Un'HookListBoxScroll
End Sub

Private Sub lstChurch_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Me.lstChurch.ListCount > 0 Then
        'HookListBoxScroll Me, Me.lstChurch
    End If
End Sub

Private Sub lstHistory_Click()
    
    '--//컨트롤설정
    If Me.lstHistory.listIndex <> -1 Then
        Me.cmdAdd.Enabled = True
        Me.cmdDelete.Enabled = True
        Me.cmdEdit.Enabled = True
        Me.txtDate.Enabled = True
        Me.txtHistory.Enabled = True
    Else
        Me.cmdAdd.Enabled = False
        Me.cmdDelete.Enabled = False
        Me.cmdEdit.Enabled = False
        Me.txtDate.Enabled = False
        Me.txtHistory.Enabled = False
    End If
    
    '--//컨트롤 내용 채우기
    If Me.lstHistory.listIndex <> -1 Then
        With Me.lstHistory
            Me.txtDate = .List(.listIndex, 2)
            Me.txtHistory = .List(.listIndex, 3)
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

Private Sub txtDate_Change()
    Call Date_Format(Me.txtDate)
End Sub

Private Sub UserForm_Initialize()
    
    If connIP = Empty Then setGlobalVariant '//전역변수 재설정
    If checkLogin = 0 Then f_login.Show '//로그인체크
    If checkLogin = 0 Then End '--//로그인 실패 시 프로시저 종료
    
    '--//기초설정
    Me.cmdClose.Cancel = True
    TB1 = "op_system.db_churchlist_custom" '--//교회리스트
    TB2 = "op_system.db_history_church" '--//교회현황
    
    '--//권한에 따른 컨트롤 설정
    Call HideDeleteButtonByUserAuth
    
    '--//컨트롤 설정
    Me.cmdAdd.Enabled = False
    Me.cmdAdd.Visible = False
    Me.cmdCancel.Visible = False
    Me.cmdNew.Enabled = False
    Me.cmdDelete.Enabled = False
    Me.cmdEdit.Enabled = False
    Me.txtDate.Enabled = False
    Me.txtHistory.Enabled = False
    
    '--//리스트박스 설정
    With Me.lstChurch
        .ColumnCount = 4
        .ColumnHeads = False
        .ColumnWidths = "0,150,50,150" '교회코드, 교회명, 교회구분, 관리교회명
'        .Width = 330
        .TextAlign = fmTextAlignLeft
        .Font = "굴림"
    End With
    
    With Me.lstHistory
        .ColumnCount = 4
        .ColumnHeads = False
        .ColumnWidths = "0,0,60,280" '이력코드, 교회코드, 날짜, 교회이력
'        .Width = 330
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
    result.affectedCount = executeSQL("cmdDelete_Click", TB2, strSql, Me.Name, "교회이력 삭제")
    writeLog "cmdDelete_Click", TB2, strSql, 0, Me.Name, "교회이력 삭제"
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
        If .listIndex > -1 Then
            If Me.txtDate = .List(.listIndex, 2) And Me.txtHistory = .List(.listIndex, 3) Then
                Exit Sub
            End If
        Else
            Exit Sub '--//리스트가 선택되지 않았으면 프로시저 종료
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
    result.affectedCount = executeSQL("cmdEdit_Click", TB2, strSql, Me.Name, "교회이력 업데이트")
    writeLog "cmdEdit_Click", TB2, strSql, 0, Me.Name, "교회이력 업데이트", result.affectedCount
    disconnectALL

    
    '--//메세지박스
    MsgBox "업데이트 되었습니다.", , banner
    
    '--//리스트박스 초기화
    Call lstChurch_Click
    Me.lstHistory.listIndex = queryKey
    
End Sub

Private Sub cmdNew_Click()
    'Call cmdbtn_visible
    Call INPUTMODE(True)
    Call HideDeleteButtonByUserAuth
    Me.lstHistory.listIndex = Me.lstHistory.ListCount - 1
    Call sbtxtBox_Init
End Sub

Private Sub cmdADD_Click()
    
    Dim argData As T_HISTORY_CHURCH
    Dim result As T_RESULT
    
    '--//중복체크
    With Me.lstChurch
        strSql = "SELECT * FROM " & TB2 & " a WHERE a.church_sid = " & SText(.List(.listIndex)) & _
                    "AND a.his_dt = " & SText(Me.txtDate) & " AND a.history = " & SText(Me.txtHistory) & ";"
        Call makeListData(strSql, TB2)
    End With
    
    If cntRecord > 0 Then
        MsgBox "중복된 내용이 존재합니다. 입력 값을 다시 확인해주세요.", vbCritical, banner
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
    argData.church_sid = Me.lstChurch.List(Me.lstChurch.listIndex)
    argData.HIS_DT = IIf(Me.txtDate = "", "1900-01-01", Me.txtDate)
    argData.HISTORY = Me.txtHistory
    
    If Me.txtDate = "" And Me.txtHistory = "" Then
        MsgBox "값을 입력해 주세요.", vbCritical, banner
        Exit Sub
    End If
    
    '--//쿼리문 실행 및 로그기록
    strSql = makeInsertSQL(TB2, argData)
    connectTaskDB
    result.strSql = strSql
    result.affectedCount = executeSQL("cmdADD_Clikc", TB2, strSql, Me.Name, "교회이력 추가")
    writeLog "cmdADD_Click", TB2, strSql, 0, Me.Name, "교회이력 추가", result.affectedCount
    disconnectALL

    '--//메세지박스
    MsgBox "추가 되었습니다.", , banner
    Call lstChurch_Click
    Me.lstHistory.listIndex = Me.lstHistory.ListCount - 1
    
    '--//버튼설정 원래대로
    'Call cmdbtn_visible
    Call INPUTMODE(False)
    Call HideDeleteButtonByUserAuth
    
End Sub
Private Sub lstChurch_Click()
    
    '--//교회 선택여부 판단
    If Me.lstChurch.listIndex = -1 Then
        MsgBox "교회를 선택하세요.", vbCritical, "오류"
        Exit Sub
    End If
    
    '--//컨트롤 설정
    If Me.lstChurch.listIndex <> -1 Then
        Me.cmdNew.Enabled = True
        Me.lstHistory.Enabled = True
    Else
        Me.cmdNew.Enabled = False
        Me.lstHistory.Enabled = False
    End If
    
    '--//텍스트박스 초기화
    Call sbtxtBox_Init
    
    '--//교회이력데이터 추가
    Erase LISTDATA
    strSql = makeSelectSQL(TB2)
    Call makeListData(strSql, TB2)
On Error Resume Next '--//이력 데이터가 없어 오류나면 내용 모두 비우기
    Me.lstHistory.List = LISTDATA
    If err.Number <> 0 Then
        Me.lstHistory.Clear
    End If
On Error GoTo 0
    Call sbClearVariant
    
    With Me.lstHistory
        .listIndex = .ListCount - 1
    End With
    
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
                        "WHERE a.ovs_dept = " & USER_DEPT & " AND a.church_gb <> 'MM' AND (a.church_nm LIKE '%" & Me.txtChurch & "%' OR b.church_nm LIKE '%" & Me.txtChurch & "%') ORDER BY a.sort_order;"
        Else
            strSql = "SELECT a.church_sid,a.church_nm,a.church_gb,b.church_nm " & _
                        "FROM " & TB1 & " a LEFT JOIN op_system.db_churchlist_custom b ON a.main_church_cd = b.church_sid " & _
                        "WHERE a.ovs_dept = " & USER_DEPT & " AND a.suspend = 0 AND a.church_gb <> 'MM' AND (a.church_nm LIKE '%" & Me.txtChurch & "%' OR b.church_nm LIKE '%" & Me.txtChurch & "%') ORDER BY a.sort_order;"
        End If
    Case TB2
        strSql = "SELECT * FROM " & TB2 & " a WHERE church_sid = " & SText(Me.lstChurch.List(Me.lstChurch.listIndex)) & ";"
    Case Else
    End Select
    makeSelectSQL = strSql
End Function
Private Function makeUpdateSQL(ByVal tableNM As String) As String
    
    With Me.lstHistory
        strSql = "UPDATE " & TB2 & " a " & _
                "SET a.his_dt = " & IIf(Me.txtDate = "", "NULL", SText(Me.txtDate)) & ",a.history = " & SText(Me.txtHistory) & _
                " WHERE a.his_cd=" & SText(.List(.listIndex)) & ";"
        queryKey = .listIndex
    End With
        
    makeUpdateSQL = strSql
End Function
Private Function makeInsertSQL(ByVal tableNM As String, _
                                argData As T_HISTORY_CHURCH) As String
    With Me.lstHistory
        strSql = "INSERT INTO " & TB2 & " VALUES(DEFAULT," & _
                    SText(argData.church_sid) & "," & _
                    IIf(argData.HIS_DT = "1900-01-01", "NULL", SText(argData.HIS_DT)) & "," & _
                    SText(argData.HISTORY) & ");"
    End With
    queryKey = Me.lstHistory.ListCount - 1
    makeInsertSQL = strSql
End Function
Private Function makeDeleteSQL(ByVal tableNM As String) As String
    With Me.lstHistory
        strSql = "DELETE FROM " & TB2 & " WHERE his_cd = " & SText(.List(.listIndex)) & ";"
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
    
    If Not IsDate(Me.txtDate) Then
        MsgBox "올바른 날짜 형태가 아닙니다. 다시 확인 해주세요.", vbCritical, banner
        Set txtBox_Focus = Me.txtDate: fnData_Validation = False: Exit Function
    End If
End Function
Sub sbtxtBox_Init()
    Me.txtDate = ""
    Me.txtHistory = ""
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
    Me.lstHistory.Enabled = Not argBoolean
    Me.chkAll.Enabled = Not argBoolean
    
    Me.txtDate.Enabled = argBoolean
    Me.txtHistory.Enabled = argBoolean
    
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


