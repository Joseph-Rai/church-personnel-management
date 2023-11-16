VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Update_BCLeader_1 
   Caption         =   "관리자 검색"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6090
   OleObjectBlob   =   "frm_Update_BCLeader_1.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "frm_Update_BCLeader_1"
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

Private Sub lstPStaff_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Un'HookListBoxScroll
End Sub

Private Sub lstPStaff_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Me.lstPStaff.ListCount Then
        'HookListBoxScroll Me, Me.lstPStaff
    End If
End Sub

Private Sub UserForm_Initialize()

    '--//기초설정
    Me.cmdClose.Cancel = True
    TB1 = "op_system.v0_pstaff_information_all" '--//선지자리스트(전체)
    TB2 = "op_system.v0_pstaff_information" '--//선지자리스트
    
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
    If Me.chkAll.Value Then
        strSql = makeSelectSQL(TB1)
        Call makeListData(strSql, TB1)
        If cntRecord > 0 Then
            Me.lstPStaff.List = LISTDATA
        Else
            Me.lstPStaff.Clear
        End If
        Call sbClearVariant
    Else
        strSql = makeSelectSQL(TB2)
        Call makeListData(strSql, TB2)
        If cntRecord > 0 Then
            Me.lstPStaff.List = LISTDATA
        Else
            Me.lstPStaff.Clear
        End If
        Call sbClearVariant
    End If
End Sub
Private Sub txtChurch_Change()
    Me.txtChurchNM.BackColor = RGB(255, 255, 255)
End Sub
Private Sub lstPStaff_Click()
    Me.cmdOk.Enabled = True
End Sub
Private Sub cmdClose_Click()
    Unload frm_Update_BCLeader_1
End Sub
Private Sub cmdOK_Click()
    
    '--//교회 선택여부 판단
    If Me.lstPStaff.listIndex = -1 Then
        MsgBox "교회를 선택하세요.", vbCritical, banner
        Exit Sub
    End If
    
    Select Case argShow
    Case 1
        '--//교회정보 입력
        With Me.lstPStaff
            frm_Update_BCLeader.txtManager = .List(.listIndex, 2)
            frm_Update_BCLeader.txtLifeNo = .List(.listIndex)
        End With
    Case 2
        '--//교회정보 입력
        With Me.lstPStaff
            frm_Update_FamilyInfo.txtLifeNo = .List(.listIndex)
            strSql = "SELECT * FROM op_system.v0_pstaff_information_all a WHERE a.`생명번호` = " & SText(.List(.listIndex)) & " OR a.`배우자생번` = " & SText(.List(.listIndex)) & ";"
            Call makeListData(strSql, "op_system.v_pstaff_detail")
            frm_Update_FamilyInfo.txtName_ko = IIf(Mid(.List(.listIndex), 12, 1) = 1, LISTDATA(0, 22), LISTDATA(0, 23))
            frm_Update_FamilyInfo.txtName_en = IIf(Mid(.List(.listIndex), 12, 1) = 1, LISTDATA(0, 7), LISTDATA(0, 16))
            frm_Update_FamilyInfo.txtChurch = LISTDATA(0, 0)
            frm_Update_FamilyInfo.cboTitle = IIf(Mid(.List(.listIndex), 12, 1) = 1, LISTDATA(0, 26), LISTDATA(0, 27))
            frm_Update_FamilyInfo.cboPosition = IIf(Mid(.List(.listIndex), 12, 1) = 1, LISTDATA(0, 8), LISTDATA(0, 17))
            frm_Update_FamilyInfo.txtEducation = IIf(Mid(.List(.listIndex), 12, 1) = 1, LISTDATA(0, 24), LISTDATA(0, 25))
            frm_Update_FamilyInfo.txtBirthday = IIf(Mid(.List(.listIndex), 12, 1) = 1, LISTDATA(0, 10), LISTDATA(0, 18))
            frm_Update_FamilyInfo.cboReligion = "본교성도"
        End With
        
        
        
    Case Else
    End Select
    
    Unload frm_Update_BCLeader_1
    
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
    
    '//리스팅할 레코드 수 검토
    If cntRecord = 0 Then
        MsgBox "반환할 DB 데이터가 없습니다.", vbInformation, banner
        Exit Sub
    End If
    
End Sub
'-----------------------------------------
'  조건별 Select SQL문 작성
'    - makeSelectSQL(검색어, 필터)
'    - DB에서 반환할 리스트 필드수정
'-----------------------------------------
Private Function makeSelectSQL(ByVal tableNM As String) As String
    
    Select Case argShow
    Case 1
        Select Case tableNM
        Case TB1
            '생명번호, 교회명, 한글이름(직분), 직책
            strSql = "SELECT a.`생명번호`,a.`교회명`,a.`한글이름(직분)`,a.`직책` " & _
                        "FROM " & TB1 & " a " & _
                        "WHERE (a.`한글이름(직분)` LIKE '%" & Me.txtChurchNM & "%' OR a.`교회명` LIKE '%" & Me.txtChurchNM & "%'" & _
                        " OR a.`영문이름` LIKE '%" & Me.txtChurchNM & "%' OR a.`지교회명` LIKE '%" & Me.txtChurchNM & "%'" & _
                        " OR a.`영문교회명` LIKE '%" & Me.txtChurchNM & "%' OR a.`영문지교회명` LIKE '%" & Me.txtChurchNM & "%'" & _
                        " OR a.`생명번호` LIKE '%" & Me.txtChurchNM & "%')" & _
                        " AND a.`관리부서` = " & SText(USER_DEPT) & ";"
        Case TB2
            '생명번호, 교회명, 한글이름(직분), 직책
            strSql = "SELECT a.`생명번호`,a.`교회명`,a.`한글이름(직분)`,a.`직책` " & _
                        "FROM " & TB2 & " a " & _
                        "WHERE (a.`한글이름(직분)` LIKE '%" & Me.txtChurchNM & "%' OR a.`교회명` LIKE '%" & Me.txtChurchNM & "%'" & _
                        " OR a.`영문이름` LIKE '%" & Me.txtChurchNM & "%' OR a.`지교회명` LIKE '%" & Me.txtChurchNM & "%'" & _
                        " OR a.`영문교회명` LIKE '%" & Me.txtChurchNM & "%' OR a.`영문지교회명` LIKE '%" & Me.txtChurchNM & "%'" & _
                        " OR a.`생명번호` LIKE '%" & Me.txtChurchNM & "%')" & _
                        " AND a.`관리부서` = " & SText(USER_DEPT) & ";"
        Case Else
        End Select
    Case 2
        '생명번호, 교회명, 한글이름(직분), 직책
        strSql = "SELECT a.`생명번호`,a.`교회명`,a.`한글이름(직분)`,a.`직책` " & _
                    "FROM " & TB1 & " a " & _
                    "WHERE (a.`한글이름(직분)` LIKE '%" & Me.txtChurchNM & "%' OR a.`교회명` LIKE '%" & Me.txtChurchNM & "%'" & _
                    " OR a.`영문이름` LIKE '%" & Me.txtChurchNM & "%' OR a.`지교회명` LIKE '%" & Me.txtChurchNM & "%'" & _
                    " OR a.`영문교회명` LIKE '%" & Me.txtChurchNM & "%' OR a.`영문지교회명` LIKE '%" & Me.txtChurchNM & "%'" & _
                    " OR a.`생명번호` LIKE '%" & Me.txtChurchNM & "%')" & _
                    " AND a.`관리부서` = " & SText(USER_DEPT) & _
                    " UNION " & _
                    "SELECT a.`배우자생번`,a.`교회명`,a.`사모한글이름(직분)`,a.`사모직책` " & _
                    "FROM " & TB1 & " a " & _
                    "WHERE (a.`사모한글이름(직분)` LIKE '%" & Me.txtChurchNM & "%' OR a.`교회명` LIKE '%" & Me.txtChurchNM & "%'" & _
                    " OR a.`사모영문이름` LIKE '%" & Me.txtChurchNM & "%' OR a.`지교회명` LIKE '%" & Me.txtChurchNM & "%'" & _
                    " OR a.`영문교회명` LIKE '%" & Me.txtChurchNM & "%' OR a.`영문지교회명` LIKE '%" & Me.txtChurchNM & "%'" & _
                    " OR a.`배우자생번` LIKE '%" & Me.txtChurchNM & "%')" & _
                    " AND a.`관리부서` = " & SText(USER_DEPT) & ";"
    Case Else
    End Select
    makeSelectSQL = strSql
End Function

Private Sub sbClearVariant()
    Erase LISTFIELD
    Erase LISTDATA
    cntRecord = Empty
    strSql = vbNullString
End Sub



