VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Search_Country 
   Caption         =   "국가 검색 마법사"
   ClientHeight    =   3000
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   5535
   OleObjectBlob   =   "frm_Search_Country.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "frm_Search_Country"
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
Dim ws As Worksheet

Private Sub lstCountry_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Un'HookListBoxScroll
End Sub

Private Sub lstCountry_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Me.lstCountry.ListCount Then
        'HookListBoxScroll Me, Me.lstCountry
    End If
End Sub

Private Sub UserForm_Initialize()
    
    '--//기초설정
    Me.cmdClose.Cancel = True
    TB1 = "op_system.db_country" '--//국가리스트
    
    '--//리스트박스 설정
    With Me.lstCountry
        .ColumnCount = 4
        .ColumnHeads = False
        .ColumnWidths = "270" '국가명
        .Width = 265.5
        .TextAlign = fmTextAlignLeft
        .Font = "굴림"
    End With
    
    Me.txtCountry.SetFocus
    
End Sub
Private Sub cmdSearch_Click()
    strSql = makeSelectSQL(TB1)
    Call makeListData(strSql, TB1)
    If cntRecord > 0 Then
        Me.lstCountry.List = LISTDATA
    End If
    Call sbClearVariant
End Sub
Private Sub txtCountry_Change()
    Me.txtCountry.BackColor = RGB(255, 255, 255)
End Sub
Private Sub lstCountry_Click()
    Me.cmdOk.Enabled = True
End Sub
Private Sub cmdClose_Click()
    Unload Me
End Sub
Private Sub cmdOK_Click()
    
    '--//교회 선택여부 판단
    If Me.lstCountry.listIndex = -1 Then
        MsgBox "국가를 선택하세요.", vbCritical, banner
        Exit Sub
    End If
    
    '--//교회정보 입력
    Select Case argShow
        Case 1
            With Me.lstCountry
                frm_Update_Flight.txtDeparture = .List(.listIndex)
            End With
        Case 2
            With Me.lstCountry
                frm_Update_Flight.txtDestination = .List(.listIndex)
            End With
        Case 3
            With Me.lstCountry
                frm_Update_PInformation.txtNationality = .List(.listIndex)
            End With
        Case 4
            With Me.lstCountry
                frm_Update_PInformation.txtNationality_Spouse = .List(.listIndex)
            End With
    Case Else
    End Select
    
    Unload Me
    
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
        strSql = "SELECT * FROM " & TB1 & " a WHERE a.ctry_nm LIKE '%" & Replace(Me.txtCountry, "한국", "대한민국") & "%';"
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



