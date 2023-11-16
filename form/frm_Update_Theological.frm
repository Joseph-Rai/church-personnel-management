VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Update_Theological 
   Caption         =   "예비생도 이력 관리마법사"
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7215
   OleObjectBlob   =   "frm_Update_Theological.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "frm_Update_Theological"
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

Private Sub chkPresent_Change()
    Select Case Me.chkPresent.Value
        Case True
            Me.txtEnd.BackColor = &HE0E0E0
            Me.txtEnd.Value = "현재"
            Me.txtEnd.Enabled = False
            Me.chkResign.Value = False
            Me.chkResign.Visible = False
            Me.txtResign.Enabled = False
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
            
            Me.chkResign.Visible = True
            
    Case Else
    End Select
End Sub

Private Sub chkResign_Click()
    Select Case Me.chkResign.Value
        Case True
            Me.txtResign.BackColor = RGB(255, 255, 255)
            Me.txtResign.Value = Date - 1
            Me.txtResign.Enabled = True
        Case False
            Me.txtResign.Enabled = False
            Me.txtResign.BackColor = &HE0E0E0
            Me.txtResign.Value = ""
    Case Else
    End Select
End Sub

Private Sub cmdCancel_Click()
'    Call cmdbtn_visible
    Call Input_Mode(False)
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
    result.affectedCount = executeSQL("cmdDelete_Click", TB2, strSql, Me.Name, "예비생도 이력 삭제")
    writeLog "cmdDelete_Click", TB2, strSql, 0, Me.Name, "예비생도 이력 삭제"
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
        If Me.cboStep = .List(.listIndex, 2) And Me.txtStart = .List(.listIndex, 3) And Me.txtEnd = .List(.listIndex, 4) And Me.txtResign = .List(.listIndex, 5) Then
            Exit Sub
        End If
    End With
    
    '--//중복체크
    With Me.lstHistory
        strSql = "SELECT * FROM " & TB2 & " a WHERE a.lifeno = " & SText(.List(.listIndex, 1)) & _
                " AND IF(a.start_dt > " & SText(Me.txtStart) & ", a.start_dt, " & SText(Me.txtStart) & ") <= " & _
                        "IF(a.end_dt < " & SText(IIf(Me.txtEnd = "현재", DateSerial(9999, 12, 31), Me.txtEnd)) & ", a.end_dt, " & _
                            SText(IIf(Me.txtEnd = "현재", DateSerial(9999, 12, 31), Me.txtEnd)) & ")" & _
                " AND a.theological_cd <> " & SText(.List(.listIndex)) & ";"
        Call makeListData(strSql, TB2)
    End With
    
    If cntRecord > 0 Then
        MsgBox "중복된 기간은 존재할 수 없습니다. 입력 값을 다시 확인해주세요.", vbCritical, banner
        queryKey = LISTDATA(0, 0)
        Call returnListPosition(Me, Me.lstHistory.Name, queryKey)
        Exit Sub
    End If
    Call sbClearVariant
    
    If Me.cboStep <> Me.lstHistory.List(Me.lstHistory.listIndex, 2) Then
        strSql = "SELECT a.Level FROM " & TB2 & " a WHERE a.lifeno = " & SText(Me.lstPStaff.List(Me.lstPStaff.listIndex))
        Call makeListData(strSql, TB2)
        
        If IsInArray(Me.cboStep, LISTDATA, , rtnSequence) <> -1 Then
            MsgBox "해당 단계로는 수정이 불가능 합니다. 다시 설정해 주세요.", vbCritical, banner
            Me.cboStep = Me.lstHistory.List(Me.lstHistory.listIndex, 2)
            Exit Sub
        End If
    End If
    
    '--//데이터 유효성 검사
    If fnData_Validation = False Then
On Error Resume Next
        txtBox_Focus.SetFocus
        txtBox_Focus.SelStart = 0
        txtBox_Focus.SelLength = Len(txtBox_Focus)
On Error GoTo 0
        Exit Sub
    End If
    
    If Me.txtResign <> "" And Me.txtEnd = "" Then
        Me.txtEnd = Me.txtResign
    End If
    
    '--//SQL문 생성, 실행, 로그기록

    strSql = makeUpdateSQL(TB2)
    result.strSql = strSql
    connectTaskDB
    result.affectedCount = executeSQL("cmdEdit_Click", TB2, strSql, Me.Name, "예비생도 이력 업데이트")
    writeLog "cmdEdit_Click", TB2, strSql, 0, Me.Name, "예비생도 이력 업데이트", result.affectedCount
    disconnectALL
    
    '--//메세지박스
    MsgBox "업데이트 되었습니다.", , banner
    
    '--//리스트박스 초기화
    Call lstPStaff_Click
    Call lstHistory_Click
    Me.lstHistory.listIndex = queryKey
    
End Sub

Private Sub cmdNew_Click()
    Call lstHistory_Click
    Call Input_Mode(True)
    Call HideDeleteButtonByUserAuth
    Me.lstHistory.listIndex = Me.lstHistory.ListCount - 1
    Call sbtxtBox_Init
    Me.chkPresent.Value = False
    Me.chkPresent.Value = True
End Sub

Private Sub cmdADD_Click()
    
    Dim argData As T_THEOLOGICAL
    Dim result As T_RESULT
    
    '--//중복체크
    strSql = "SELECT * FROM " & TB2 & " a WHERE a.lifeno = " & SText(Me.lstPStaff.List(Me.lstPStaff.listIndex)) & _
            " AND IF(a.start_dt > " & SText(Me.txtStart) & ", a.start_dt, " & SText(Me.txtStart) & ") <= " & _
            "IF(a.end_dt < " & SText(IIf(Me.txtEnd = "현재", DateSerial(9999, 12, 31), Me.txtEnd)) & ", a.end_dt, " & _
                SText(IIf(Me.txtEnd = "현재", DateSerial(9999, 12, 31), Me.txtEnd)) & ");"
    Call makeListData(strSql, TB2)
   
    
    If cntRecord > 0 Then
        MsgBox "중복된 기간은 존재할 수 없습니다. 입력 값을 다시 확인해주세요.", vbCritical, banner
        queryKey = LISTDATA(0, 0)
        Call returnListPosition(Me, Me.lstHistory.Name, queryKey)
        Exit Sub
    End If
    Call sbClearVariant
    
    strSql = "SELECT a.Level FROM " & TB2 & " a WHERE a.resign_dt IS NULL AND a.lifeno = " & SText(Me.lstPStaff.List(Me.lstPStaff.listIndex))
    Call makeListData(strSql, TB2)
    
    If cntRecord > 0 Then
        If IsInArray(Me.cboStep, LISTDATA, , rtnSequence) <> -1 Then
            MsgBox "이미 등록된 단계 입니다. 다시 설정해 주세요.", vbCritical, banner
            Me.cboStep.SetFocus
            Me.cboStep.SelStart = 0
            Me.cboStep.SelLength = Len(Me.cboStep)
            Exit Sub
        End If
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
        argData.lifeNo = Me.lstPStaff.List(Me.lstPStaff.listIndex)
        argData.LEVEL = Me.cboStep
        argData.START_DT = Me.txtStart
        argData.END_DT = IIf(Me.txtEnd = "현재", DateSerial(9999, 12, 31), Me.txtEnd)
        argData.RESIGN_DT = IIf(Me.txtResign = "", "1900-01-01", Me.txtResign)
        
        strSql = "SELECT a.church_sid,a.church_nm,a.church_gb,b.church_nm " & _
                    "FROM " & "op_system.db_churchlist_custom" & " a LEFT JOIN op_system.db_churchlist_custom b ON a.main_church_cd = b.church_sid " & _
                    "WHERE a.church_gb = 'MC' AND (a.church_nm =" & SText(Me.lstPStaff.List(Me.lstPStaff.listIndex, 1)) & "OR b.church_nm = " & SText(Me.lstPStaff.List(Me.lstPStaff.listIndex, 1)) & ") ORDER BY a.sort_order;"
                    'a.ovs_dept = " & USER_DEPT & " AND / AND a.suspend = 0
        Call makeListData(strSql, TB1)
        If cntRecord > 0 Then
            argData.RECOMMAND_CHURCH = LISTDATA(0, 0)
        Else
            MsgBox "현재 소속된 교회가 없습니다. 발령 이력을 먼저 확인해주세요.", vbCritical, banner
            Exit Sub
        End If
    End With
    
    '--//작업에 따라 쿼리문 실행 및 로그기록
    strSql = makeInsertSQL(TB2, argData)
    connectTaskDB
    result.strSql = strSql
    result.affectedCount = executeSQL("cmdADD_Clikc", TB2, strSql, Me.Name, "예비생도 이력 추가")
    writeLog "cmdADD_Click", TB2, strSql, 0, Me.Name, "예비생도 이력 추가", result.affectedCount
    disconnectALL
    
    '--//메세지박스
    MsgBox "추가 되었습니다.", , banner
    Call lstPStaff_Click
    Me.lstHistory.listIndex = Me.lstHistory.ListCount - 1
    
    '--//버튼설정 원래대로
    Call Input_Mode(False)
    Me.controls("cmdCancel").Visible = False
'    For Each txtBox_Focus In Me.Controls
'        If TypeName(txtBox_Focus) = "CommandButton" Then
'            txtBox_Focus.Visible = True
'        End If
'    Next
    Call HideDeleteButtonByUserAuth
    
End Sub
Private Sub lstHistory_Click()
    
    '--//컨트롤 설정
    If Me.lstHistory.listIndex <> -1 Then
        Me.txtStart.Enabled = True
        Me.txtEnd.Enabled = True
        Me.cboStep.Enabled = True
        Me.cmdEdit.Enabled = True
        Me.cmdAdd.Enabled = True
        Me.cmdDelete.Enabled = True
        Me.chkPresent.Visible = True
    Else
        Me.txtStart.Enabled = False
        Me.txtEnd.Enabled = False
        Me.cboStep.Enabled = False
        Me.cmdEdit.Enabled = False
        Me.cmdAdd.Enabled = False
        Me.cmdDelete.Enabled = False
        Me.chkPresent.Visible = False
    End If
    
    '--//리스트 클릭 시 시작일, 종료일, 내용 표시
    With Me.lstHistory
        If .ListCount > 0 And .listIndex <> -1 Then
            Me.cboStep = .List(.listIndex, 2)
            Me.txtStart = .List(.listIndex, 3)
            Me.txtEnd = .List(.listIndex, 4)
            Me.txtResign = .List(.listIndex, 5)
        End If
    End With
    
    
    If Me.txtEnd = "현재" Then
        Me.chkPresent.Value = True
    Else
        Me.chkPresent.Value = False
    End If
    
    If Me.txtEnd <> "현재" And Me.txtEnd <> "" Then
        Me.chkResign.Visible = True
    End If
    
    '--//현재상태 추가
    If Me.lstHistory.ListCount > 0 Then
        With Me.lstHistory
            Me.txtPresent = .List(.listIndex, 2)
            If .List(.listIndex, 5) = "" Then
                If .List(.listIndex, 4) = "9999-12-31" Or .List(.listIndex, 4) = "현재" Then
                    Me.txtPresent = Me.txtPresent & " 진행 중"
                Else
                    Me.txtPresent = Me.txtPresent & " 수료"
                End If
            Else
                Me.txtPresent = Me.txtPresent & " 성도복귀"
            End If
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
    
    Call UserForm_Initialize
    
    '--//컨트롤설정
    If Me.lstPStaff.listIndex <> -1 Then
        Me.lstHistory.Enabled = True
        Me.cmdNew.Enabled = True
    Else
        Me.lstHistory.Enabled = False
        Me.cmdNew.Enabled = False
    End If
    Me.txtPresent = ""
    
    '--//이력 목록상자 조정
    With Me.lstHistory
        .ColumnCount = 8
        .ColumnHeads = False
        .ColumnWidths = "0,0,70,65,65,0,0,120" '예비생도코드, 생명번호, 생도단계, 시작일, 종료일, 성도복귀일, 교회코드, 교회명
        .Width = 337
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
    Me.chkResign.Visible = False
    
    '--//이력 리스트박스가 비어있지 않으면 마지막 데이터 클릭
    If Me.lstHistory.ListCount > 0 Then
        Me.lstHistory.listIndex = Me.lstHistory.ListCount - 1
    End If
    
    '--//사진추가
    InsertPicToLabel Me.lblPic, Me.lstPStaff.List(Me.lstPStaff.listIndex)
    
End Sub

Private Sub lstPStaff_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Un'HookListBoxScroll
End Sub

Private Sub lstPStaff_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    'HookListBoxScroll Me, Me.lstPStaff
End Sub

Private Sub txtChurchNM_Change()
    Me.txtChurchNM.BackColor = RGB(255, 255, 255)
End Sub

Private Sub txtEnd_Change()
    Call Date_Format(Me.txtEnd)
End Sub

Private Sub txtPresent_Change()

End Sub

Private Sub txtResign_Change()
    Call Date_Format(Me.txtResign)
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
    TB1 = "op_system.v0_pstaff_information_all" '--//선지자정보(전체검색)
    TB2 = "op_system.db_theological" '--//예비생도 이력
    TB3 = "op_system.v0_pstaff_information" '--//선지자정보
    
    '--//권한에 따른 컨트롤 설정
    Call HideDeleteButtonByUserAuth
    
    '--//컨트롤 설정
    Me.lstHistory.Enabled = False
    Me.txtStart.Enabled = False
    Me.txtEnd.Enabled = False
    Me.txtResign.Enabled = False
    Me.cboStep.Enabled = False
    Me.cmdEdit.Enabled = False
    Me.cmdAdd.Enabled = False
    Me.cmdAdd.Visible = False
    Me.cmdCancel.Visible = False
    Me.cmdNew.Enabled = False
    Me.cmdDelete.Enabled = False
    Me.chkPresent.Visible = False
    Me.txtPresent.Enabled = False
    Me.chkResign.Visible = False
    Me.txtResign.BackColor = &HE0E0E0
    Me.cboStep.Clear
    Me.cboStep.AddItem "예비생도1단계"
    Me.cboStep.AddItem "예비생도2단계"
    Me.cboStep.AddItem "예비생도3단계"
    
    '--//리스트박스 설정
    With Me.lstPStaff
        .ColumnCount = 4
        .ColumnHeads = False
        .ColumnWidths = "0,120,70,50" '생명번호, 교회명, 한글이름(직분), 직책
        .TextAlign = fmTextAlignLeft
        .Font = "굴림"
    End With
    
    If Me.txtChurchNM.Enabled = True Then
        Me.txtChurchNM.SetFocus
    End If
    
End Sub
Private Sub cmdSearch_Click()
    
    If Me.chkAll.Value Then
        Me.lstHistory.Clear
        
        strSql = makeSelectSQL(TB1)
        Call makeListData(strSql, TB1)
        
        If cntRecord = 0 Then
            MsgBox "반환할 DB 데이터가 없습니다.", vbInformation, banner
            Call sbClearVariant
            Exit Sub
        End If
        
        Me.lstPStaff.List = LISTDATA
        Call sbClearVariant
        Me.lstPStaff.Enabled = True
    Else
        Me.lstHistory.Clear
        
        strSql = makeSelectSQL(TB3)
        Call makeListData(strSql, TB3)
        
        If cntRecord = 0 Then
            MsgBox "반환할 DB 데이터가 없습니다.", vbInformation, banner
            Call sbClearVariant
            Exit Sub
        End If
        
        Me.lstPStaff.List = LISTDATA
        Call sbClearVariant
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
                    " OR a.`영문이름` LIKE '%" & Me.txtChurchNM & "%' OR a.`지교회명` LIKE '%" & Me.txtChurchNM & "%'" & _
                    " OR a.`영문교회명` LIKE '%" & Me.txtChurchNM & "%' OR a.`영문지교회명` LIKE '%" & Me.txtChurchNM & "%'" & _
                    " OR a.`생명번호` LIKE '%" & Me.txtChurchNM & "%')" & _
                    " AND a.`관리부서` = " & SText(USER_DEPT) & ";"
    Case TB2
        strSql = "SELECT a.theological_cd,a.lifeno,a.level,a.start_dt,if(a.end_dt='9999-12-31','현재',a.end_dt),a.resign_dt,a.church_sid,b.church_nm " & _
                "FROM " & TB2 & " a LEFT JOIN op_system.db_churchlist_custom b " & _
                "ON a.church_sid = b.church_sid WHERE a.lifeno = " & SText(Me.lstPStaff.List(Me.lstPStaff.listIndex)) & _
                "ORDER BY a.start_dt;"
    Case TB3
        strSql = "SELECT a.`생명번호`,a.`교회명`,a.`한글이름(직분)`,a.`직책` " & _
                    "FROM " & TB3 & " a " & _
                    "WHERE (a.`한글이름(직분)` LIKE '%" & Me.txtChurchNM & "%' OR a.`교회명` LIKE '%" & Me.txtChurchNM & "%'" & _
                    " OR a.`영문이름` LIKE '%" & Me.txtChurchNM & "%' OR a.`지교회명` LIKE '%" & Me.txtChurchNM & "%'" & _
                    " OR a.`영문교회명` LIKE '%" & Me.txtChurchNM & "%' OR a.`영문지교회명` LIKE '%" & Me.txtChurchNM & "%'" & _
                    " OR a.`생명번호` LIKE '%" & Me.txtChurchNM & "%')" & _
                    " AND a.`관리부서` = " & SText(USER_DEPT) & ";"
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
                    "SET a.lifeno = " & SText(.List(.listIndex, 1)) & ",a.level = " & SText(Me.cboStep) & ", a.start_dt = " & SText(Me.txtStart) & _
                    ", a.end_dt = " & SText(IIf(Me.txtEnd = "현재", DateSerial(9999, 12, 31), Me.txtEnd)) & ",a.resign_dt = " & IIf(Me.txtResign = "", "NULL", SText(Me.txtResign)) & ",a.church_sid = " & SText(.List(.listIndex, 6)) & " WHERE a.theological_cd = " & SText(.List(.listIndex)) & ";"
        queryKey = .listIndex
        End With
    Case Else
    End Select
    makeUpdateSQL = strSql
End Function
Private Function makeInsertSQL(ByVal tableNM As String, argData As T_THEOLOGICAL) As String
    
    Select Case tableNM
    Case TB1
    Case TB2
        strSql = "INSERT INTO " & TB2 & " VALUES(DEFAULT," & _
                    SText(argData.lifeNo) & "," & _
                    SText(argData.LEVEL) & "," & _
                    SText(argData.START_DT) & "," & _
                    SText(argData.END_DT) & "," & _
                    IIf(argData.RESIGN_DT = "1900-01-01", "NULL", SText(argData.RESIGN_DT)) & "," & _
                    SText(argData.RECOMMAND_CHURCH) & ");"
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
            strSql = "DELETE FROM " & TB2 & " WHERE theological_cd = " & SText(.List(.listIndex)) & ";"
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
    Me.cboStep = ""
    Me.txtStart.Value = ""
    Me.txtEnd.Value = "현재"
    Me.txtResign = ""
End Sub
Private Function fnData_Validation()
'---------------------------------------
'유저폼 입력값에 대한 데이터 유효성 검사
'TRUE: 이상없음, FALSE: 잘못됨.
'---------------------------------------
    fnData_Validation = True '데이터가 유효하다는 가정 하에 시작
    
    If IsInArray(Me.cboStep, Array("예비생도1단계", "예비생도2단계", "예비생도3단계"), True, rtnValue) = -1 Then
        MsgBox "예비생도 단계를 잘못 입력하였습니다. 다시 확인 해주세요.", vbCritical, banner
        Set txtBox_Focus = Me.cboStep: fnData_Validation = False: Exit Function
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
    
    If Me.txtResign <> "" And Me.txtEnd <> "" Then
        If CDate(Me.txtResign) < IIf(Me.txtEnd = "현재", CDate("9999-12-31"), CDate(Me.txtEnd)) Then
            MsgBox "성도 복귀일은 종료일보다 크거나 같아야 합니다. 다시 확인해 주세요.", vbCritical, banner
            Set txtBox_Focus = Me.txtResign: fnData_Validation = False: Exit Function
        End If
    End If
    
    If Me.cboStep = "" Or Me.txtStart = "" Or Me.txtEnd = "" Then
        MsgBox "필수 입력값이 누락되었습니다. 다시 확인해주세요.", vbCritical, banner
        If Me.cboStep = "" Then Set txtBox_Focus = Me.cboStep: fnData_Validation = False: Exit Function
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
    
    Me.chkPresent.Value = argBoolean
    
    Me.lstHistory.Enabled = Not argBoolean
    Me.lstPStaff.Enabled = Not argBoolean
    Me.txtChurchNM.Enabled = Not argBoolean
    Me.cmdSearch.Enabled = Not argBoolean
    Me.chkAll.Enabled = Not argBoolean
    
    Me.cboStep.Enabled = argBoolean
    Me.txtStart.Enabled = argBoolean
    Me.txtEnd.Enabled = argBoolean
    Me.txtResign.Enabled = argBoolean
    Me.chkPresent.Visible = argBoolean
    Me.chkResign.Visible = argBoolean
    
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

