VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Update_Sermon 
   Caption         =   "발표평가 관리마법사"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7140
   OleObjectBlob   =   "frm_Update_Sermon.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "frm_Update_Sermon"
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

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdEdit_Click()
    
    Dim result As T_RESULT
    Dim argData As T_SERMON
    
    '--//수정된 내용 있는지 체크
    With Me.lstPStaff
        If Me.txtScore_Avg = .List(.listIndex, 4) And Me.txtSubject_Count = .List(.listIndex, 5) Then
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
    With Me.lstPStaff
        strSql = "SELECT * FROM " & TB2 & " a WHERE a.lifeno = " & SText(.List(.listIndex)) & ";"
        Call makeListData(strSql, TB2)
    
        If cntRecord > 0 Then
            strSql = makeUpdateSQL(TB2)
        Else
            argData.lifeNo = .List(.listIndex)
            argData.SCORE_AVG = Me.txtScore_Avg
            argData.SUBJECT_COUNT = Me.txtSubject_Count
            strSql = makeInsertSQL(TB2, argData)
        End If
    End With
    
    result.strSql = strSql
    connectTaskDB
    result.affectedCount = executeSQL("cmdEdit_Click", TB2, strSql, Me.Name, "발표점수 업데이트")
    writeLog "cmdEdit_Click", TB2, strSql, 0, Me.Name, "발표점수 업데이트", result.affectedCount
    disconnectALL
    
    Call sbClearVariant
    
    '--//메세지박스
    MsgBox "업데이트 되었습니다.", , banner
    
    '--//리스트박스 초기화
    Call cmdSearch_Click
    Call lstPStaff_Click
'    Me.lstPStaff.ListIndex = queryKey
    
End Sub

Private Sub lstPStaff_Click()
    
    Dim filePath As String
    Dim FileName As String
    
    Me.cmdEdit.Enabled = True
    Me.txtScore_Avg.Enabled = True
    Me.txtSubject_Count.Enabled = True
    
    '--//텍스트박스 초기화
    Me.txtScore_Avg = ""
    Me.txtSubject_Count = ""
    
    '--//텍스트박스 내용추가
    With Me.lstPStaff
        Me.txtScore_Avg = .List(.listIndex, 4)
        Me.txtSubject_Count = .List(.listIndex, 5)
    End With
    
    '--//사진추가
    filePath = fnFindPicPath
    FileName = Me.lstPStaff.List(Me.lstPStaff.listIndex) & ".jpg"
    
'    If Not Len(Dir(FilePath & FileName)) > 0 Then
'        FileName = Me.lstPStaff.List(Me.lstPStaff.ListIndex) & ".png"
'    End If
    
On Error Resume Next
    Me.lblPic.Picture = LoadPicture(filePath & FileName)
    If err.Number <> 0 Then
        Me.lblPic.Picture = LoadPicture("")
    End If
On Error GoTo 0
    
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

Private Sub txtScore_Avg_Change()
    If InStr(Me.txtScore_Avg, ".") > 0 Then
        Me.txtScore_Avg = Left(Me.txtScore_Avg, InStr(Me.txtScore_Avg, ".") + 2)
    Else
        Me.txtScore_Avg = Left(Me.txtScore_Avg, 2)
    End If
End Sub

Private Sub UserForm_Initialize()
    If connIP = Empty Then setGlobalVariant '//전역변수 재설정
    If checkLogin = 0 Then f_login.Show '//로그인체크
    If checkLogin = 0 Then End '--//로그인 실패 시 프로시저 종료
    
    '--//기초설정
    Me.cmdClose.Cancel = True
    TB1 = "op_system.v0_pstaff_information_all" '--//선지자정보(전체검색)
    TB2 = "op_system.db_sermon" '--//발표점수
    TB3 = "op_system.v0_pstaff_information" '--//선지자정보
    
    '--//컨트롤 설정
    Me.txtScore_Avg.Enabled = False
    Me.txtSubject_Count.Enabled = False
    Me.cmdEdit.Enabled = False
    
    '--//리스트박스 설정
    With Me.lstPStaff
        .ColumnCount = 4
        .ColumnHeads = False
        .ColumnWidths = "0,120,70,50,0,0" '생명번호, 교회명, 한글이름(직분), 직책, 평균점수, 발표개수
        .TextAlign = fmTextAlignLeft
        .Font = "굴림"
    End With
    
    Me.txtChurchNM.SetFocus
    
End Sub
Private Sub cmdSearch_Click()
    
    If Me.chkAll.Value Then
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
        strSql = "SELECT a.`생명번호`,a.`교회명`,a.`한글이름(직분)`,a.`직책`,b.score_avg,b.subject_count " & _
                    "FROM " & TB1 & " a " & _
                    "LEFT JOIN op_system.db_sermon b on a.`생명번호` = b.lifeno " & _
                    "WHERE (a.`한글이름(직분)` LIKE '%" & Me.txtChurchNM & "%' OR a.`교회명` LIKE '%" & Me.txtChurchNM & "%'" & _
                    " OR a.`영문이름` LIKE '%" & Me.txtChurchNM & "%' OR a.`지교회명` LIKE '%" & Me.txtChurchNM & "%'" & _
                    " OR a.`영문교회명` LIKE '%" & Me.txtChurchNM & "%' OR a.`영문지교회명` LIKE '%" & Me.txtChurchNM & "%'" & _
                    " OR a.`생명번호` LIKE '%" & Me.txtChurchNM & "%')" & _
                    " AND a.`관리부서` = " & SText(USER_DEPT) & ";"
    Case TB2
    Case TB3
        '--//교회코드, 교회명
        strSql = "SELECT a.`생명번호`,a.`교회명`,a.`한글이름(직분)`,a.`직책`,b.score_avg,b.subject_count " & _
                    "FROM " & TB3 & " a " & _
                    "LEFT JOIN op_system.db_sermon b on a.`생명번호` = b.lifeno " & _
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
        With Me.lstPStaff
            strSql = "UPDATE " & TB2 & " a " & _
                    "SET a.score_avg = " & SText(Me.txtScore_Avg) & ", a.subject_count = " & SText(Me.txtSubject_Count) & _
                    " WHERE a.lifeno = " & SText(.List(.listIndex)) & ";"
        queryKey = .listIndex
        End With
    Case Else
    End Select
    makeUpdateSQL = strSql
End Function

Private Function makeInsertSQL(ByVal tableNM As String, argData As T_SERMON) As String
    
    Select Case tableNM
    Case TB1
    Case TB2
        strSql = "INSERT INTO " & TB2 & " VALUES(" & _
                    SText(argData.lifeNo) & "," & _
                    SText(argData.SCORE_AVG) & "," & _
                    SText(argData.SUBJECT_COUNT) & ");"
    Case Else
    End Select
    makeInsertSQL = strSql
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
    
    If Not IsNumeric(Me.txtScore_Avg) And Me.txtScore_Avg <> "" Then
        MsgBox "발표점수를 잘못 입력하였습니다. 다시 확인 해주세요.", vbCritical, banner
        Set txtBox_Focus = Me.txtScore_Avg: fnData_Validation = False: Exit Function
    End If
    
    If Me.txtScore_Avg = "" Then
        MsgBox "발표점수는 필수 입력값 입니다. 다시 확인 해주세요.", vbCritical, banner
        Set txtBox_Focus = Me.txtScore_Avg: fnData_Validation = False: Exit Function
    End If
    
    If Not IsNumeric(Me.txtSubject_Count) And Me.txtSubject_Count <> "" Then
        MsgBox "발표개수를 잘못 입력하였습니다. 다시 확인 해주세요.", vbCritical, banner
        Set txtBox_Focus = Me.txtSubject_Count: fnData_Validation = False: Exit Function
    End If
    
End Function


