VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Update_Visa 
   Caption         =   "비자정보 관리마법사"
   ClientHeight    =   8295.001
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14625
   OleObjectBlob   =   "frm_Update_Visa.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "frm_Update_Visa"
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
    result.affectedCount = executeSQL("cmdDelete_Click", TB2, strSql, Me.Name, "비자이력 삭제")
    writeLog "cmdDelete_Click", TB2, strSql, 0, Me.Name, "비자이력 삭제"
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
        If Me.cboVisa = .List(.listIndex, 4) And Me.txtStart = .List(.listIndex, 2) And Me.txtEnd = .List(.listIndex, 3) And Me.txtMemo = .List(.listIndex, 5) Then
            Exit Sub
        End If
    End With
    
    '--//중복체크
    With Me.lstHistory
        strSql = "SELECT * FROM " & TB2 & " a WHERE a.lifeno = " & SText(.List(.listIndex, 1)) & _
                " AND ((a.start_dt BETWEEN " & SText(Me.txtStart) & " AND " & SText(IIf(Me.txtEnd = "현재", DateSerial(9999, 12, 31), Me.txtEnd)) & _
                ") OR (a.end_dt BETWEEN " & SText(Me.txtStart) & " AND " & SText(IIf(Me.txtEnd = "현재", DateSerial(9999, 12, 31), Me.txtEnd)) & _
                ") OR (a.start_dt <= " & SText(Me.txtStart) & " AND a.end_dt >= " & SText(IIf(Me.txtEnd = "현재", DateSerial(9999, 12, 31), Me.txtEnd)) & _
                ")) AND a.visa_cd <> " & SText(.List(.listIndex)) & ";"
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
    result.affectedCount = executeSQL("cmdEdit_Click", TB2, strSql, Me.Name, "비자이력 업데이트")
    writeLog "cmdEdit_Click", TB2, strSql, 0, Me.Name, "비자이력 업데이트", result.affectedCount
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
    Call INPUTMODE(True)
    Call HideDeleteButtonByUserAuth
    Me.lstHistory.listIndex = Me.lstHistory.ListCount - 1
    Call sbtxtBox_Init
    Me.chkPresent.Value = True
End Sub

Private Sub cmdADD_Click()
    
    Dim argData As T_VISA
    Dim result As T_RESULT
    
    '--//중복체크
    
    strSql = "SELECT * FROM " & TB2 & " a WHERE a.lifeno = " & SText(Me.lstPStaff.List(Me.lstPStaff.listIndex)) & _
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
        argData.lifeNo = Me.lstPStaff.List(Me.lstPStaff.listIndex)
        argData.START_DT = Me.txtStart
        argData.END_DT = IIf(Me.txtEnd = "현재", DateSerial(9999, 12, 31), Me.txtEnd)
        argData.Visa = Me.cboVisa
        argData.memo = Me.txtMemo
    End With
    
    '--//작업에 따라 쿼리문 실행 및 로그기록
    strSql = makeInsertSQL(TB2, argData)
    connectTaskDB
    result.strSql = strSql
    result.affectedCount = executeSQL("cmdADD_Clikc", TB2, strSql, Me.Name, "비자이력 추가")
    writeLog "cmdADD_Click", TB2, strSql, 0, Me.Name, "비자이력 추가", result.affectedCount
    disconnectALL
    
    '--//메세지박스
    MsgBox "추가 되었습니다.", , banner
    Call lstPStaff_Click
    Me.lstHistory.listIndex = Me.lstHistory.ListCount - 1
    
    '--//버튼설정 원래대로
    Call INPUTMODE(False)
    Call HideDeleteButtonByUserAuth
    
End Sub

Private Sub cmdSavePassportPhoto_Click()
    Dim filePath As String
    Dim lifeNo As String
    With Me.lstPStaff
        lifeNo = .List(.listIndex)
    End With
    
    filePath = Application.GetOpenFilename("Image File Format, *.jpg;*.jpeg;", , "Photo Select")
    
    If filePath = "False" Then
        Exit Sub
    End If
    
    '--//DB에 사진 저장
    savePassportPhoto lifeNo, filePath
    
    '--//성공메시지
    MsgBox "사진이 성공적으로 저장되었습니다.", , banner
    
    '--//라벨에 사진삽입
    InsertPicToLabel Me.lblPassportPhoto, lifeNo, PASSPORT_PHOTO
End Sub

Private Sub savePassportPhoto(lifeNo As String, filePath As String)

    Dim ppPhoto As New PassportPhoto
    Dim ppPhotoDao As New PassportPhotoDao
    Dim stream As New ADODB.stream
    
    stream.Type = adTypeBinary
    stream.Open
    stream.LoadFromFile filePath
    
    ppPhoto.lifeNo = lifeNo
    ppPhoto.photo = encodeBase64(stream.Read)
    ppPhotoDao.Save ppPhoto

End Sub

Private Sub cmdSaveVisaPhoto_Click()
    Dim filePath As String
    Dim visa_cd As String
    With Me.lstHistory
        If .ListCount > 0 Then
            visa_cd = .List(.listIndex)
        Else
            Exit Sub
        End If
    End With
    
    filePath = Application.GetOpenFilename("Image File Format, *.jpg;*.jpeg;", , "Photo Select")
    
    If filePath = "False" Then
        Exit Sub
    End If
    
    '--//DB에 사진 저장
    saveVisaPhoto visa_cd, filePath
    
    '--//성공메시지
    MsgBox "사진이 성공적으로 저장되었습니다.", , banner
    
    '--//라벨에 사진삽입
    InsertPicToLabel Me.lblVisaPhoto, visa_cd, VISA_PHOTO
End Sub

Private Sub saveVisaPhoto(visaCode As String, filePath As String)

    Dim vPhoto As New VisaPhoto
    Dim vPhotoDao As New VisaPhotoDao
    Dim stream As New ADODB.stream
    
    stream.Type = adTypeBinary
    stream.Open
    stream.LoadFromFile filePath
    
    vPhoto.visaCode = visaCode
    vPhoto.photo = encodeBase64(stream.Read)
    vPhotoDao.Save vPhoto

End Sub

Private Sub lblVisaPhoto_Click()

End Sub

Private Sub lstHistory_Click()
    
    '--//컨트롤 설정
    If Me.lstHistory.listIndex <> -1 Then
        Me.txtStart.Enabled = True
        Me.txtEnd.Enabled = True
        Me.cboVisa.Enabled = True
        Me.txtMemo.Enabled = True
        Me.cmdEdit.Enabled = True
        Me.cmdAdd.Enabled = True
        Me.cmdDelete.Enabled = True
        Me.chkPresent.Visible = True
    Else
        Me.txtStart.Enabled = False
        Me.txtEnd.Enabled = False
        Me.cboVisa.Enabled = False
        Me.txtMemo.Enabled = False
        Me.cmdEdit.Enabled = False
        Me.cmdAdd.Enabled = False
        Me.cmdDelete.Enabled = False
        Me.chkPresent.Visible = False
    End If
    
    '--//리스트 클릭 시 시작일, 종료일, 내용 표시
    Dim visa_cd As String
    With Me.lstHistory
        If .ListCount > 0 And .listIndex <> -1 Then
            Me.cboVisa = .List(.listIndex, 4)
            Me.txtStart = .List(.listIndex, 2)
            Me.txtEnd = IIf(.List(.listIndex, 3) = "9999-12-31", "현재", .List(.listIndex, 3))
            Me.txtMemo = .List(.listIndex, 5)
            
            visa_cd = .List(.listIndex)
            
            Me.cmdSavePassportPhoto.Enabled = True
        Else
            Me.cmdSavePassportPhoto.Enabled = False
        End If
    End With
    
    '--//체크박스 값 조정
    If Me.txtEnd = "현재" Then
        Me.chkPresent.Value = True
    Else
        Me.chkPresent.Value = False
    End If
    
    '--//사증사진 추가
    InsertPicToLabel Me.lblVisaPhoto, visa_cd, VISA_PHOTO
    If Not Me.lblVisaPhoto.Picture = 0 Then
        Me.cmdSaveVisaPhoto.Caption = "사증수정"
    Else
        Me.cmdSaveVisaPhoto.Caption = "사증등록"
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
    
    '--//컨트롤 설정
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
        .ColumnWidths = "0,0,65,65,70,120" '비자이력코드, 생명번호, 시작일, 종료일, 비자종류, 메모
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
    
    '--//선지자 클릭하면 여권저장 버튼 활성화
    With Me.lstPStaff
        If .listIndex > -1 Then
            Me.cmdSavePassportPhoto.Enabled = True
        Else
            Me.cmdSavePassportPhoto.Enabled = False
        End If
    End With
    
    '--//이력 리스트박스가 비어있지 않으면 마지막 데이터 클릭
    If Me.lstHistory.ListCount > 0 Then
        Me.lstHistory.listIndex = Me.lstHistory.ListCount - 1
        
        Me.cmdSaveVisaPhoto.Enabled = True
    Else
        Me.cmdSaveVisaPhoto.Enabled = False
    End If
    
    '--//사진추가
    Dim strLifeNo As String
    With Me.lstPStaff
        strLifeNo = .List(.listIndex)
    End With
    InsertPicToLabel Me.lblPic, strLifeNo
    
    '--//여권사진 추가
    InsertPicToLabel Me.lblPassportPhoto, strLifeNo, PASSPORT_PHOTO
    If Not Me.lblPassportPhoto.Picture = 0 Then
        Me.cmdSavePassportPhoto.Caption = "여권수정"
    Else
        Me.cmdSavePassportPhoto.Caption = "여권등록"
    End If
     
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

Private Sub txtStart_Change()
    Call Date_Format(Me.txtStart)
End Sub

Private Sub UserForm_Initialize()
    If connIP = Empty Then setGlobalVariant '//전역변수 재설정
    If checkLogin = 0 Then f_login.Show '//로그인체크
    If checkLogin = 0 Then End '--//로그인 실패 시 프로시저 종료
    
    '--//기초설정
    Me.cmdClose.Cancel = True
    TB1 = "op_system.v0_pstaff_information_all" '--//선지자정보
    TB2 = "op_system.db_visa" '--//비자이력
    
    '--//권한에 따른 컨트롤 설정
    Call HideDeleteButtonByUserAuth
    
    '--//컨트롤 설정
    Me.lstHistory.Enabled = False
    Me.txtStart.Enabled = False
    Me.txtEnd.Enabled = False
    Me.cboVisa.Enabled = False
    Me.txtMemo.Enabled = False
    Me.cmdEdit.Enabled = False
    Me.cmdAdd.Enabled = False
    Me.cmdAdd.Visible = False
    Me.cmdCancel.Visible = False
    Me.cmdNew.Enabled = False
    Me.cmdDelete.Enabled = False
    Me.chkPresent.Visible = False
    
    '--//콤보박스 아이템 추가
    Me.cboVisa.Clear
'    Me.cboVisa.AddItem "관광비자"
'    Me.cboVisa.AddItem "사업비자"
'    Me.cboVisa.AddItem "학생비자"
'    Me.cboVisa.AddItem "취업비자"
'    Me.cboVisa.AddItem "선교비자"
'    Me.cboVisa.AddItem "자원봉사자비자"
'    Me.cboVisa.AddItem "결혼비자"
    strSql = "SELECT a.visa_nm FROM op_system.a_visa a;"
    Call makeListData(strSql, "op_system.a_visa")
    If cntRecord > 0 Then
        Me.cboVisa.List = LISTDATA
    End If
    Call sbClearVariant
    
    '--//리스트박스 설정
    With Me.lstPStaff
        .ColumnCount = 4
        .ColumnHeads = False
        .ColumnWidths = "0,120,70,50" '생명번호, 교회명, 한글이름(직분), 직책
        .TextAlign = fmTextAlignLeft
        .Font = "굴림"
    End With
    
    Me.txtChurchNM.SetFocus
    
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
        If Me.chkNative.Value Then
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
        Else
            '--//교회코드, 교회명
            strSql = "SELECT a.`생명번호`,a.`교회명`,a.`한글이름(직분)`,a.`직책` " & _
                       "FROM " & TB1 & " a " & _
                       "WHERE (a.`한글이름(직분)` LIKE '%" & Me.txtChurchNM & "%' OR a.`교회명` LIKE '%" & Me.txtChurchNM & "%'" & _
                       " OR a.`영문이름` LIKE '%" & Me.txtChurchNM & "%'" & _
                       " OR a.`생명번호` LIKE '%" & Me.txtChurchNM & "%')" & _
                       " AND a.`국적` = '대한민국'" & _
                       " AND a.`관리부서` = " & SText(USER_DEPT) & _
                    " UNION " & _
                    "SELECT b.`배우자생번`,b.`교회명`,b.`사모한글이름(직분)`,b.`사모직책` " & _
                       "FROM " & TB1 & " b " & _
                       "WHERE (b.`사모한글이름(직분)` LIKE '%" & Me.txtChurchNM & "%' OR b.`교회명` LIKE '%" & Me.txtChurchNM & "%'" & _
                       " OR b.`사모영문이름` LIKE '%" & Me.txtChurchNM & "%'" & _
                       " OR b.`배우자생번` LIKE '%" & Me.txtChurchNM & "%')" & _
                       " AND b.`사모국적` = '대한민국'" & _
                       " AND b.`관리부서` = " & SText(USER_DEPT) & ";"
        End If
    Case TB2
        strSql = "SELECT a.visa_cd,a.lifeno,a.start_dt,if(a.end_dt='9999-12-31','현재',a.end_dt),a.`visa`,a.`memo` " & _
                "FROM " & TB2 & " a " & _
                "WHERE a.lifeno = " & SText(Me.lstPStaff.List(Me.lstPStaff.listIndex)) & " ORDER BY a.`start_dt`;"
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
                    ", a.visa = " & SText(Me.cboVisa) & ", a.memo = " & SText(Me.txtMemo) & _
                    " WHERE a.visa_cd = " & SText(.List(.listIndex)) & ";"
        queryKey = .listIndex
        End With
    Case Else
    End Select
    makeUpdateSQL = strSql
End Function
Private Function makeInsertSQL(ByVal tableNM As String, argData As T_VISA) As String
    
    Select Case tableNM
    Case TB1
    Case TB2
        strSql = "INSERT INTO " & TB2 & " VALUES(DEFAULT," & _
                    SText(argData.lifeNo) & "," & _
                    SText(argData.START_DT) & "," & _
                    SText(argData.END_DT) & "," & _
                    SText(argData.Visa) & "," & _
                    SText(argData.memo) & ");"
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
            strSql = "DELETE FROM " & TB2 & " WHERE visa_cd = " & SText(.List(.listIndex)) & ";"
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
    Me.cboVisa = ""
    Me.txtStart.Value = ""
    Me.txtEnd.Value = "현재"
    Me.txtMemo = ""
End Sub
Private Function fnData_Validation()
'---------------------------------------
'유저폼 입력값에 대한 데이터 유효성 검사
'TRUE: 이상없음, FALSE: 잘못됨.
'---------------------------------------
    fnData_Validation = True '데이터가 유효하다는 가정 하에 시작
    
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
    
    If Me.cboVisa = "" Or Me.txtStart = "" Or Me.txtEnd = "" Then
        MsgBox "필수 입력값이 누락되었습니다. 다시 확인해주세요.", vbCritical, banner
        If Me.cboVisa = "" Then Set txtBox_Focus = Me.cboVisa: fnData_Validation = False: Exit Function
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

Private Sub INPUTMODE(ByVal argBoolean As Boolean)
    Me.cmdNew.Visible = Not argBoolean
    Me.cmdEdit.Visible = Not argBoolean
    Me.cmdDelete.Visible = Not argBoolean
    Me.cmdCancel.Visible = argBoolean
    Me.cmdCancel.Enabled = argBoolean
    Me.cmdAdd.Visible = argBoolean
    Me.cmdAdd.Enabled = argBoolean
    
    Me.txtStart.Enabled = argBoolean
    Me.txtEnd.Enabled = argBoolean
    Me.cboVisa.Enabled = argBoolean
    Me.txtMemo.Enabled = argBoolean
    Me.chkPresent.Visible = argBoolean
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





