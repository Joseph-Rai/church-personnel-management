VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Update_Union_1 
   Caption         =   "연합회 목록관리 마법사"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   2865
   OleObjectBlob   =   "frm_Update_Union_1.frx":0000
End
Attribute VB_Name = "frm_Update_Union_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cntlst1Col As Integer '//리스트 컬럼 수
Dim TB1 As String '//폼에 연결된 DB 테이블
Const strOrderBy As String = "currency_un ASC" '//DB에서 sort_order 필드
Dim caseSave As Integer '//1: 추가, 2: 수정, 3: 삭제(SUSPEND), 4: 완전삭제
Dim queryKey As Integer '//리스트 위치 반환에 사용될 id
Dim LISTDATA() As String '--//DB에서 받아온 rs를 배열로 저장
Dim LISTFIELD() As String '--//DB에서 받아온 rs의 필드를 배열로 저장
Dim cntRecord As Integer '--//DB에서 받아온 레코드의 개수
Dim strSql As String '--//SQL 쿼리문
Dim UnionNM As String '--//수정할 연합회명

Private Sub cmdADD_Click()
    '--//연합회 추가
    Dim argData As T_UNION
    Dim result As T_RESULT
    
    '--//중복체크
    strSql = "SELECT * FROM " & TB1 & " a WHERE a.suspend = 0 AND a.union_nm = " & SText(Me.txtUnion) & " AND a.ovs_dept = " & SText(USER_DEPT) & ";"
    Call makeListData(strSql, TB1)
   
    If cntRecord > 0 Then
        MsgBox "중복된 연합회명이 존재합니다. 다시 확인해주세요.", vbCritical, banner
        queryKey = LISTDATA(0, 0)
        Call returnListPosition(Me, Me.lstUnion.Name, queryKey)
        Exit Sub
    End If
    Call sbClearVariant
    
    '--//작업에 따라 쿼리문 실행 및 로그기록
    strSql = "SELECT * FROM " & TB1 & " a WHERE a.union_nm = " & SText(Me.txtUnion) & " AND a.ovs_dept = " & SText(USER_DEPT) & ";"
    Call makeListData(strSql, TB1)
    If cntRecord > 0 Then
        strSql = makeUpdateSQL2(TB1, LISTDATA(0, 0))
    Else
        strSql = makeInsertSQL(TB1, argData)
    End If
    connectTaskDB
    result.strSql = strSql
    result.affectedCount = executeSQL("cmdADD_Clikc", TB1, strSql, Me.Name, "연합회 추가")
    writeLog "cmdADD_Click", TB1, strSql, 0, Me.Name, "연합회 추가", result.affectedCount
    disconnectALL
    
    '--//메세지박스
    MsgBox "추가 되었습니다.", , banner
    Call UserForm_Initialize '--//새로고침
    Me.lstUnion.listIndex = Me.lstUnion.ListCount - 1
    
    '--//버튼설정 원래대로
    Call HideDeleteButtonByUserAuth
End Sub

Private Sub cmdDelete_Click()
    
    Dim result As T_RESULT
    
    If MsgBox("선택한 데이터를 삭제하시겠습니까?", vbQuestion + vbYesNo, banner) = vbNo Then
        Exit Sub
    End If
    
    '--//선택한 연합회 정렬값 불러오기
    With Me.lstUnion
        strSql = "SELECT a.sort_order FROM " & TB1 & " a WHERE union_cd = " & SText(.List(.listIndex)) & ";"
    End With
    Call makeListData(strSql, TB1)
    
    '--//선택한 연합회 논리삭제
    strSql = makeDeleteSQL(TB1)
    connectTaskDB
    result.strSql = strSql
    result.affectedCount = executeSQL("cmdDelete_Click", TB1, strSql, Me.Name, "연합회 삭제")
    writeLog "cmdDelete_Click", TB1, strSql, 0, Me.Name, "연합회 삭제"
    disconnectALL
    
    '--//나머지 연합회 정렬순서 조정
    strSql = makeDeleteSQL2(TB1)
    connectTaskDB
    result.strSql = strSql
    result.affectedCount = executeSQL("cmdDelete_Click", TB1, strSql, Me.Name, "연합회 삭제")
    writeLog "cmdDelete_Click", TB1, strSql, 0, Me.Name, "연합회 삭제"
    disconnectALL
    
    '--//메세지박스
    MsgBox "해당 데이터가 삭제되었습니다.", , banner
    
    '--//리스트박스 새로고침
    Call UserForm_Initialize '--//새로고침
    Me.lstUnion.listIndex = Me.lstUnion.ListCount - 1
    
End Sub

Private Sub cmdEdit_Click()
    Dim result As T_RESULT
    
    '--//수정할 연합회명 받아오기
'    UnionNM = Application.InputBox("수정할 연합회명을 입력하세요.", banner)
'    If UnionNM = "" Then Exit Sub
    
    '--//중복체크
    With Me.lstUnion
        strSql = "SELECT * FROM " & TB1 & " a WHERE a.union_nm = " & SText(Me.txtUnion) & " AND a.ovs_dept = " & SText(USER_DEPT) & ";"
        Call makeListData(strSql, TB1)
    End With
    
    If Me.lstUnion.listIndex < 0 Then
        Exit Sub
    End If
    
    If cntRecord > 0 Then
        MsgBox "중복된 연합회명이 존재합니다. 다시 확인해주세요.", vbCritical, banner
        queryKey = LISTDATA(0, 0)
        Call returnListPosition(Me, Me.lstUnion.Name, queryKey)
        Exit Sub
    End If
    
    Call sbClearVariant
    
    '--//SQL문 생성, 실행, 로그기록
    strSql = makeUpdateSQL(TB1)
    result.strSql = strSql
    connectTaskDB
    result.affectedCount = executeSQL("cmdEdit_Click", TB1, strSql, Me.Name, "연합회명 수정")
    writeLog "cmdEdit_Click", TB1, strSql, 0, Me.Name, "연합회명 수정", result.affectedCount
    disconnectALL
    
    '--//메세지박스
    MsgBox "수정 되었습니다.", , banner
    
    '--//리스트박스 초기화
    Call UserForm_Initialize '--//새로고침
    Me.lstUnion.listIndex = Me.lstUnion.ListCount - 1
End Sub

Private Sub cmdMoveDown_Click()
    Dim result As T_RESULT
    Dim noMax As Integer
    
    '--//연합회 정렬순서 Max값 산출
    strSql = "SELECT MAX(a.sort_order) FROM op_system.a_union a WHERE a.ovs_dept = " & SText(USER_DEPT) & ";"
    Call makeListData(strSql, "op_system.a_union")
    noMax = LISTDATA(0, 0)
    
    With Me.lstUnion
        '--//선택한 연합회 sort_order 픽업
        strSql = "SELECT sort_order FROM " & TB1 & " WHERE ovs_dept = " & SText(USER_DEPT) & " AND union_cd = " & SText(.List(.listIndex)) & ";"
        Call makeListData(strSql, TB1)
        
        If LISTDATA(0, 0) < noMax Then '--//현재 Sort_Order가 Max값이 아니면
            '--//선택한 연합회 sort_order 1 증가
            strSql = "UPDATE " & TB1 & " a SET a.sort_order = a.sort_order + 1 WHERE a.union_cd = " & SText(.List(.listIndex)) & ";"
            result.strSql = strSql
            connectTaskDB
            result.affectedCount = executeSQL("cmdMoveDown_Click", TB1, strSql, Me.Name, "연합회 정렬순서 변경")
            writeLog "cmdMoveDown_Click", TB1, strSql, 0, Me.Name, "연합회 정렬순서 변경", result.affectedCount
            disconnectALL
            
            '--//직후 연합회 sort_order 1 감소
            strSql = "UPDATE " & TB1 & " a SET a.sort_order = a.sort_order - 1 WHERE a.ovs_dept = " & SText(USER_DEPT) & " AND a.sort_order = " & SText(LISTDATA(0, 0) + 1) & " AND a.union_cd <> " & SText(.List(.listIndex)) & ";"
            result.strSql = strSql
            connectTaskDB
            result.affectedCount = executeSQL("cmdMoveDown_Click", TB1, strSql, Me.Name, "연합회 정렬순서 변경")
            writeLog "cmdMoveDown_Click", TB1, strSql, 0, Me.Name, "연합회 정렬순서 변경", result.affectedCount
            disconnectALL
        End If
    End With
    
    '--//리스트박스 새로고침
    Call UserForm_Initialize '--//새로고침
'    If Me.lstUnion.ListIndex < Me.lstUnion.ListCount - 1 Then
'        Me.lstUnion.ListIndex = Me.lstUnion.ListIndex + 1
'    End If
    
End Sub

Private Sub cmdMoveUp_Click()
    
    Dim result As T_RESULT
    
    With Me.lstUnion
        '--//선택한 연합회 sort_order 픽업
        strSql = "SELECT sort_order FROM " & TB1 & " WHERE ovs_dept = " & SText(USER_DEPT) & " AND union_cd = " & SText(.List(.listIndex)) & ";"
        Call makeListData(strSql, TB1)
        
        If LISTDATA(0, 0) > 1 Then '--//현재 Sort_Order가 Min값이 아니면
            '--//선택한 연합회 sort_order 1 감소
            strSql = "UPDATE " & TB1 & " a SET a.sort_order = a.sort_order - 1 WHERE a.union_cd = " & SText(.List(.listIndex)) & ";"
            result.strSql = strSql
            connectTaskDB
            result.affectedCount = executeSQL("cmdMoveDown_Click", TB1, strSql, Me.Name, "연합회 정렬순서 변경")
            writeLog "cmdMoveDown_Click", TB1, strSql, 0, Me.Name, "연합회 정렬순서 변경", result.affectedCount
            disconnectALL
            
            '--//직전 연합회 sort_order 1 증가
            strSql = "UPDATE " & TB1 & " a SET a.sort_order = a.sort_order + 1 WHERE a.ovs_dept = " & SText(USER_DEPT) & " AND a.sort_order = " & SText(LISTDATA(0, 0) - 1) & " AND a.union_cd <> " & SText(.List(.listIndex)) & ";"
            result.strSql = strSql
            connectTaskDB
            result.affectedCount = executeSQL("cmdMoveDown_Click", TB1, strSql, Me.Name, "연합회 정렬순서 변경")
            writeLog "cmdMoveDown_Click", TB1, strSql, 0, Me.Name, "연합회 정렬순서 변경", result.affectedCount
            disconnectALL
        End If
    End With
    
    '--//리스트박스 새로고침
    Call UserForm_Initialize '--//새로고침
'    If Me.lstUnion.ListIndex > 0 Then
'        Me.lstUnion.ListIndex = Me.lstUnion.ListIndex - 1
'    End If
    
End Sub

Private Sub lstUnion_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Un'HookListBoxScroll
End Sub

Private Sub lstUnion_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Me.lstUnion.ListCount Then
        'HookListBoxScroll Me, Me.lstUnion
    End If
End Sub

Private Sub UserForm_Initialize()

    '--//기초설정
    Me.cmdClose.Cancel = True
    TB1 = "op_system.a_union" '--//교회리스트
    
    '--//컨트롤설정
    Me.cmdDelete.Enabled = False
    Me.cmdAdd.Enabled = False
    Me.txtUnion = ""
    
    '--//리스트박스 설정
    With Me.lstUnion
        .ColumnCount = 2
        .ColumnHeads = False
        .ColumnWidths = "0,120" '연합회코드, 연합회명
        .TextAlign = fmTextAlignLeft
        .Font = "굴림"
    End With
    
    '--//표시위치
    Me.Top = frm_Update_Union.Top
    Me.Left = frm_Update_Union.Left + frm_Update_Union.Width
    
    '--//연합회 목록추가
    strSql = makeSelectSQL(TB1)
    Call makeListData(strSql, TB1)
    If cntRecord > 0 Then
        Me.lstUnion.List = LISTDATA
    Else
        Me.lstUnion.Clear
    End If
    Call sbClearVariant
    
    Me.txtUnion.SetFocus
End Sub

Private Sub txtUnion_Change()
    Me.txtUnion.BackColor = RGB(255, 255, 255)
    Me.cmdAdd.Enabled = True
End Sub
Private Sub lstUnion_Click()
    Me.cmdDelete.Enabled = True
    Me.txtUnion = Me.lstUnion.List(Me.lstUnion.listIndex, 1)
End Sub
Private Sub cmdClose_Click()
    Unload frm_Update_Union_1
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
'        MsgBox "반환할 DB 데이터가 없습니다.", vbInformation, banner
        Exit Sub
    End If
    
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
        strSql = "SELECT * " & _
                    "FROM " & TB1 & " a WHERE a.suspend = 0 AND a.ovs_dept = " & SText(USER_DEPT) & " ORDER BY a.sort_order;"
    Case Else
    End Select
    makeSelectSQL = strSql
End Function
Private Function makeUpdateSQL(ByVal tableNM As String) As String
    
    Select Case tableNM
    Case TB1
        With Me.lstUnion
            strSql = "UPDATE " & TB1 & " a " & _
                    "SET a.union_nm = " & SText(Me.txtUnion) & ",a.suspend = 0" & _
                    " WHERE a.union_cd = " & SText(.List(.listIndex)) & " AND a.ovs_dept = " & SText(USER_DEPT) & ";"
        queryKey = .listIndex
        End With
    Case Else
    End Select
    makeUpdateSQL = strSql
End Function
Private Function makeUpdateSQL2(ByVal tableNM As String, ByVal UNION_CD As Long) As String
    
    Select Case tableNM
    Case TB1
        With Me.lstUnion
            strSql = "SELECT MAX(a.sort_order) FROM op_system.a_union a WHERE a.ovs_dept = " & SText(USER_DEPT) & ";"
            Call makeListData(strSql, "op_system.a_union")
            
            strSql = "UPDATE " & TB1 & " a " & _
                    "SET a.union_nm = " & SText(Me.txtUnion) & ",a.suspend = 0" & ",a.sort_order = " & SText(LISTDATA(0, 0) + 1) & _
                    " WHERE a.union_cd = " & SText(UNION_CD) & " AND a.ovs_dept = " & SText(USER_DEPT) & ";"
        queryKey = .listIndex
        End With
    Case Else
    End Select
    makeUpdateSQL2 = strSql
End Function
Private Function makeInsertSQL(ByVal tableNM As String, argData As T_UNION) As String
    
    Select Case tableNM
    Case TB1
        strSql = "SELECT MAX(a.sort_order) FROM op_system.a_union a WHERE a.ovs_dept = " & SText(USER_DEPT) & ";"
        Call makeListData(strSql, "op_system.a_union")
        
        If cntRecord > 0 Then
            strSql = "INSERT INTO " & TB1 & " VALUES(DEFAULT," & _
                        SText(Me.txtUnion) & ",0," & SText(USER_DEPT) & "," & SText(IIf(LISTDATA(0, 0) = "", 0, LISTDATA(0, 0)) + 1) & ");"
        Else
            strSql = "INSERT INTO " & TB1 & " VALUES(DEFAULT," & _
                        SText(Me.txtUnion) & ",0," & SText(USER_DEPT) & ",1);"
        End If
        queryKey = Me.lstUnion.ListCount - 1
    Case Else
    End Select
    makeInsertSQL = strSql
End Function
Private Function makeDeleteSQL(ByVal tableNM As String) As String
    
    Select Case tableNM
    Case TB1
        With Me.lstUnion
            strSql = "UPDATE " & TB1 & " SET suspend = 1,sort_order = 0 WHERE union_cd = " & SText(.List(.listIndex)) & ";"
        End With
    Case Else
    End Select
    makeDeleteSQL = strSql
End Function
Private Function makeDeleteSQL2(ByVal tableNM As String) As String
    
    Select Case tableNM
    Case TB1
        With Me.lstUnion
            strSql = "UPDATE " & TB1 & " SET sort_order = sort_order - 1 WHERE ovs_dept = " & SText(USER_DEPT) & " AND sort_order > " & SText(LISTDATA(0, 0)) & ";"
        End With
    Case Else
    End Select
    makeDeleteSQL2 = strSql
End Function

Private Sub sbClearVariant()
    Erase LISTFIELD
    Erase LISTDATA
    cntRecord = Empty
    strSql = vbNullString
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




