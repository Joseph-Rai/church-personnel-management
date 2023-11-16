VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Update_Appointment_1 
   Caption         =   "교회검색"
   ClientHeight    =   2880
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   5640
   OleObjectBlob   =   "frm_Update_Appointment_1.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "frm_Update_Appointment_1"
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

Private Sub lstChurch_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Un'HookListBoxScroll
End Sub

Private Sub lstChurch_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Me.lstChurch.ListCount Then
        'HookListBoxScroll Me, Me.lstChurch
    End If
End Sub

Private Sub UserForm_Initialize()

    '--//기초설정
    Me.cmdClose.Cancel = True
    TB1 = "op_system.db_churchlist_custom" '--//교회리스트
    
    '--//리스트박스 설정
    With Me.lstChurch
        .ColumnCount = 4
        .ColumnHeads = False
        .ColumnWidths = "0,120,0,0" '교회코드, 교회명, 교회구분, 관리교회명
        .Width = 265.5
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
    Else
        Me.lstChurch.Clear
    End If
    Call sbClearVariant
End Sub
Private Sub txtChurch_Change()
    Me.txtChurch.BackColor = RGB(255, 255, 255)
End Sub
Private Sub lstChurch_Click()
    Me.cmdOk.Enabled = True
End Sub
Private Sub cmdClose_Click()
    Unload frm_Update_Appointment_1
End Sub
Private Sub cmdOK_Click()
    
    '--//교회 선택여부 판단
    If Me.lstChurch.listIndex = -1 Then
        MsgBox "교회를 선택하세요.", vbCritical, "오류"
        Exit Sub
    End If
    
    '--//교회정보 입력
    Select Case argShow3
    Case 1
        Select Case argShow
            Case 1
                With Me.lstChurch
                    frm_Update_Appointment.txtChurchNow = .List(.listIndex, 1)
                    frm_Update_Appointment.txtChurchNow_sid = .List(.listIndex)
                End With
            Case 2
'                With Me.lstChurch
'                    frm_Update_PInformation.txtChurchNow = .list(.listIndex, 1)
'                    frm_Update_PInformation.txtChurchNow_sid = .list(.listIndex)
'                End With
        Case Else
        End Select
    Case 2
        With Me.lstChurch
            frm_Update_FamilyInfo.txtChurch = .List(.listIndex, 1)
            frm_Update_FamilyInfo.txtChurch_Sid = .List(.listIndex)
        End With
    Case 3
        With Me.lstChurch
            frm_Search_Appointment.txtTo = .List(.listIndex, 1)
            frm_Search_Appointment.txtTo_sid = .List(.listIndex)
        End With
    Case Else
    End Select
    
    Unload frm_Update_Appointment_1
    
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
    
    Select Case tableNM
    Case TB1
        '--//교회코드, 교회명
        Select Case argShow
        Case 1 '--//본교회만 검색
            If Me.chkAll.Value = False Then
                strSql = "SELECT a.church_sid,a.church_nm,a.church_gb,b.church_nm " & _
                            "FROM " & TB1 & " a LEFT JOIN op_system.db_churchlist_custom b ON a.main_church_cd = b.church_sid " & _
                            "WHERE a.ovs_dept = " & USER_DEPT & " AND (a.church_gb = 'MC' OR a.church_gb = 'HBC') AND (a.church_nm LIKE '%" & Me.txtChurch & "%' OR b.church_nm LIKE '%" & Me.txtChurch & "%') ORDER BY a.sort_order;"
            Else
                strSql = "SELECT a.church_sid,a.church_nm,a.church_gb,b.church_nm " & _
                            "FROM " & TB1 & " a LEFT JOIN op_system.db_churchlist_custom b ON a.main_church_cd = b.church_sid " & _
                            "WHERE (a.church_gb = 'MC' OR a.church_gb = 'HBC') AND (a.church_nm LIKE '%" & Me.txtChurch & "%' OR b.church_nm LIKE '%" & Me.txtChurch & "%') ORDER BY a.sort_order;"
            End If
        Case 2 '--//본교회 지교회 모두 검색
            If Me.chkAll.Value = False Then
                strSql = "SELECT a.church_sid,a.church_nm,a.church_gb,b.church_nm " & _
                            "FROM " & TB1 & " a LEFT JOIN op_system.db_churchlist_custom b ON a.main_church_cd = b.church_sid " & _
                            "WHERE a.suspend=0 AND a.ovs_dept = " & USER_DEPT & " AND a.suspend = 0 AND (a.church_nm LIKE '%" & Me.txtChurch & "%' OR b.church_nm LIKE '%" & Me.txtChurch & "%') ORDER BY a.sort_order;"
            Else
                strSql = "SELECT a.church_sid,a.church_nm,a.church_gb,b.church_nm " & _
                            "FROM " & TB1 & " a LEFT JOIN op_system.db_churchlist_custom b ON a.main_church_cd = b.church_sid " & _
                            "WHERE (a.church_nm LIKE '%" & Me.txtChurch & "%' OR b.church_nm LIKE '%" & Me.txtChurch & "%') ORDER BY a.sort_order;"
            End If
        Case 3
            If Me.chkAll.Value = False Then
                strSql = "SELECT a.church_sid,a.church_nm,a.church_gb,b.church_nm " & _
                            "FROM " & TB1 & " a LEFT JOIN op_system.db_churchlist_custom b ON a.main_church_cd = b.church_sid " & _
                            "WHERE a.ovs_dept = " & USER_DEPT & " AND a.suspend = 0 AND (a.church_gb = 'MC' OR a.church_gb = 'HBC') AND (a.church_nm LIKE '%" & Me.txtChurch & "%' OR b.church_nm LIKE '%" & Me.txtChurch & "%') ORDER BY a.sort_order;"
            Else
                strSql = "SELECT a.church_sid,a.church_nm,a.church_gb,b.church_nm " & _
                            "FROM " & TB1 & " a LEFT JOIN op_system.db_churchlist_custom b ON a.main_church_cd = b.church_sid " & _
                            "WHERE (a.church_gb = 'MC' OR a.church_gb = 'HBC') AND a.suspend = 0 AND (a.church_nm LIKE '%" & Me.txtChurch & "%' OR b.church_nm LIKE '%" & Me.txtChurch & "%') ORDER BY a.sort_order;"
            End If
        End Select
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

