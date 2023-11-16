VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Search_History 
   Caption         =   "선지자 검색"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5655
   OleObjectBlob   =   "frm_Search_History.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "frm_Search_History"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cntlst1Col As Integer '//리스트 컬럼 수
Dim TB1 As String, TB2 As String, TB3 As String, TB4 As String, TB5 As String, TB6 As String '//폼에 연결된 DB 테이블
Const strOrderBy As String = "currency_un ASC" '//DB에서 sort_order 필드
Dim caseSave As Integer '//1: 추가, 2: 수정, 3: 삭제(SUSPEND), 4: 완전삭제
Dim queryKey As Integer '//리스트 위치 반환에 사용될 id
Dim LISTDATA() As String '--//DB에서 받아온 rs를 배열로 저장
Dim LISTFIELD() As String '--//DB에서 받아온 rs의 필드를 배열로 저장
Dim cntRecord As Integer '--//DB에서 받아온 레코드의 개수
Dim strSql As String '--//SQL 쿼리문
Dim ws As Worksheet


Private Sub lstPStaff_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Un'HookListBoxScroll
End Sub

Private Sub lstPStaff_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Me.lstPStaff.ListCount > 0 Then
        'HookListBoxScroll Me, Me.lstPStaff
    End If
End Sub

Private Sub UserForm_Initialize()
    
    If connIP = Empty Then setGlobalVariant '//전역변수 재설정
    If checkLogin = 0 Then f_login.Show '//로그인체크
    If checkLogin = 0 Then End '--//로그인 실패 시 프로시저 종료
    
    '--//시트설정
    Set ws = ActiveSheet
    
    '--//기초설정
    Me.cmdClose.Cancel = True
    TB1 = "op_system.v0_pstaff_information" '--//유저폼에서 선지자 검색을 위한
    TB2 = "op_system.v_transfer_history" '--//선지자연혁 리포트 발령이력을 위한
    TB3 = "op_system.v_familyinfo" '--//가족정보
    
    '--//리스트박스 설정
    With Me.lstPStaff
        .ColumnCount = 4
        .ColumnHeads = False
        .ColumnWidths = "0,120,70,50" '생명번호, 교회명, 한글이름(직분), 직책
        .TextAlign = fmTextAlignLeft
        .Font = "굴림"
    End With
    
    Me.cmdOk.Enabled = False
    Me.txtChurchNM.SetFocus

End Sub

Private Sub cmdSearch_Click()
    Me.lstPStaff.Clear
    strSql = makeSelectSQL(TB1)
    Call makeListData(strSql, TB1)
    If cntRecord > 0 Then
        Me.lstPStaff.List = LISTDATA
    End If
    Call sbClearVariant
End Sub

Private Sub txtChurchNM_Change()
    Me.txtChurchNM.BackColor = RGB(255, 255, 255)
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub
Private Sub lstPStaff_Click()
    Me.cmdOk.Enabled = True
End Sub
Private Sub cmdOK_Click()

    Dim i As Integer, j As Integer
    Dim filePath As String
    Dim FileName As String
    Dim rngTarget As Range
    
    '--//시트 활성화 및 잠금해제
    WB_ORIGIN.Activate
    ws.Activate
    Call shUnprotect(globalSheetPW)
    
    '--//교회선택여부 확인
    If Me.lstPStaff.listIndex = -1 Then
        MsgBox "목록에 선택된 값이 없습니다.", vbCritical, "선택오류"
        Exit Sub
    End If
    
    '--//기존 데이터 삭제
    Range("His_rngTarget").CurrentRegion.ClearContents
    Range("His_rngFamily").CurrentRegion.ClearContents
    
    '--//기본정보 및 발령이력 삽입
        '--//SQL문
        strSql = makeSelectSQL(TB2)
        
        '--//DB에서 자료 호출하여 레코드셋 반환
        connectTaskDB
        Call makeListData(strSql, TB2)
    '    callDBtoRS "cmdOK_Click", TB2, strSQL, Me.Name
        
        '--//반환된 ListData를 보고서 시트에 삽입
        Optimization
        Range("His_rngTarget").Resize(, UBound(LISTFIELD) + 1) = LISTFIELD
        If cntRecord > 0 Then
            Range("His_rngTarget").Offset(1).Resize(cntRecord, UBound(LISTFIELD) + 1) = LISTDATA
        End If
        Normal
        
        Call sbClearVariant
        disconnectALL
    
    '--//가족정보 삽입
        strSql = makeSelectSQL(TB3) '--//가족정보
        connectTaskDB
        Call makeListData(strSql, TB3)
    '    callDBtoRS "cmdOK_Click", TB2, strSQL, Me.Name
    
        '--//반환된 ListData를 보고서 시트에 삽입
        Optimization
        Range("His_rngFamily").Resize(, UBound(LISTFIELD) + 1) = LISTFIELD
        If cntRecord > 0 Then
            Range("His_rngFamily").Offset(1).Resize(cntRecord, UBound(LISTFIELD) + 1) = LISTDATA
        End If
        Normal
        
        Call sbClearVariant
        disconnectALL
    
    '--//문자열로 된 숫자, 숫자데이터로 변환
    On Error Resume Next
    Range("His_rngFamily").Offset(-1).Copy
    Range("His_rngFamily").Offset(1, Range("AL29") - 1).Resize(Range("AK26")).PasteSpecial Paste:=xlPasteValues, operation:=xlPasteSpecialOperationAdd
    Range("His_rngFamily").Offset(1, 1).Resize(Range("AK26")).PasteSpecial Paste:=xlPasteValues, operation:=xlPasteSpecialOperationAdd
    On Error GoTo 0
    
    Application.CutCopyMode = False
    
    '--//가족,건강,기타 영역조정
    On Error Resume Next
    ActiveSheet.Outline.ShowLevels RowLevels:=2
    Range(Range("His_Family"), Range("His_Family").Offset(10)).Rows.Ungroup
    On Error GoTo 0
    For i = 0 To 8
        If Range("His_Family").Offset(i + 2) = "" And Range("His_Family").Offset(i + 2, 4) = "" Then
            Range("His_Family").Offset(i + 2).Rows.Group
        End If
    Next
    On Error Resume Next
'    rngTarget.Rows.Group
    On Error GoTo 0
    ActiveSheet.Outline.ShowLevels RowLevels:=1
    
    Range("23:23").EntireRow.AutoFit '--//건강 행높이 조절
    Range("24:24").EntireRow.AutoFit '--//기타 행높이 조절
    
    '--//사진삽입
On Error Resume Next
    ActiveSheet.Pictures.Delete

    If Range("His_LifeNo") <> "" Then
        InsertPStaffPic Range("His_LifeNo"), Range("His_Pic_M")
    End If
    
    If Not (Range("His_LifeNo_Spouse") = "" Or Range("His_LifeNo_Spouse") = "0") Then
        InsertPStaffPic Range("His_LifeNo_Spouse"), Range("His_Pic_F")
    End If
    
    '--//마지막에 가짜사진 추가 후 삭제하여 뒤틀어짐 방지
    InsertPStaffPic "", Range("Z9")
    If ActiveSheet.Pictures.Count > 0 Then
        ActiveSheet.Pictures(ActiveSheet.Pictures.Count).Delete
    End If
On Error GoTo 0

    Sheets("선지자연혁").Range("C1").Select
    
    Call shProtect(globalSheetPW)
    
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
        '생명번호, 교회명, 한글이름(직분), 직책
        strSql = "SELECT a.`생명번호`,a.`교회명`,a.`한글이름(직분)`,a.`직책` " & _
                    "FROM " & TB1 & " a " & _
                    "WHERE (a.`한글이름(직분)` LIKE '%" & Me.txtChurchNM & "%' OR a.`교회명` LIKE '%" & Me.txtChurchNM & "%'" & _
                    " OR a.`영문이름` LIKE '%" & Me.txtChurchNM & "%' OR a.`지교회명` LIKE '%" & Me.txtChurchNM & "%'" & _
                    " OR a.`영문교회명` LIKE '%" & Me.txtChurchNM & "%' OR a.`영문지교회명` LIKE '%" & Me.txtChurchNM & "%')" & _
                    " AND a.`관리부서` = " & SText(USER_DEPT) & ";"

    Case TB2
        strSql = "SELECT * FROM op_system.v_transfer_history a WHERE a.`생명번호` = " & SText(Me.lstPStaff.List(Me.lstPStaff.listIndex)) & ";"
    Case TB3
        If Range("T4") = 0 Then
            
            strSql = "SELECT DISTINCT a.family_cd FROM " & TB3 & " a WHERE a.lifeno = " & SText(Range("J4").Value) & " AND a.relations NOT IN (""부"",""모"");"
            Call makeListData(strSql, TB3)
            
            If cntRecord = 1 Then
                Range("AC26") = LISTDATA(0, 0)
                strSql = "SELECT * FROM (SELECT a.family_id,a.family_cd,REPLACE(a.relations,'(별세)','') AS 'relations',a.lifeno,a.name_ko,a.name_en,a.church_sid,a.church_nm,a.title,a.POSITION,a.birthday,a.education,a.religion,a.recognition,a.memo,a.SUSPEND,CONCAT(a.title,if(a.title<>'' AND a.position<>'','/',''),a.position) AS 'status',RANK() OVER (PARTITION BY a.family_cd,relations ORDER BY a.birthday) AS 'rank' FROM op_system.v_familyinfo a WHERE a.family_cd = " & SText(LISTDATA(0, 0)) & " ORDER BY family_cd,FIELD(relations,'부','모','형제','자매'),birthday) a WHERE a.lifeno <> " & SText(Range("J4").Value) & ";"
            ElseIf cntRecord > 1 Then
                MsgBox "선지자 가족정보 데이터에 중복오류가 있습니다. 중복된 자료를 제거하세요.", vbCritical, banner
            End If
        Else
            strSql = "SELECT DISTINCT a.family_cd FROM " & TB3 & " a WHERE a.lifeno = " & SText(Range("J4").Value) & " AND a.relations NOT IN (""부"",""모"")"
            Call makeListData(strSql, TB3)
            
            If cntRecord = 0 Then
                strSql = "SELECT DISTINCT a.family_cd FROM " & TB3 & " a WHERE a.lifeno = " & SText(Range("T4").Value) & " AND a.relations NOT IN (""부"",""모"");"
                Call makeListData(strSql, TB3)
                
                If cntRecord = 1 Then
                    Range("AC26") = LISTDATA(0, 0)
                    strSql = "SELECT * FROM (SELECT a.family_id,a.family_cd,REPLACE(a.relations,'(별세)','') AS 'relations',a.lifeno,a.name_ko,a.name_en,a.church_sid,a.church_nm,a.title,a.POSITION,a.birthday,a.education,a.religion,a.recognition,a.memo,a.SUSPEND,CONCAT(a.title,if(a.title<>'' AND a.position<>'','/',''),a.position) AS 'status',RANK() OVER (PARTITION BY a.family_cd,relations ORDER BY a.birthday) AS 'rank' FROM op_system.v_familyinfo a WHERE a.family_cd = " & SText(LISTDATA(0, 0)) & " ORDER BY family_cd,FIELD(relations,'부','모','형제','자매'),birthday) a WHERE a.lifeno NOT IN (" & SText(Range("T4").Value) & ");"
                End If
            ElseIf cntRecord = 1 Then
                strSql = "SELECT DISTINCT a.family_cd FROM " & TB3 & " a WHERE a.lifeno = " & SText(Range("T4").Value) & " AND a.relations NOT IN (""부"",""모"");"
                Call makeListData(strSql, TB3)
                
                If cntRecord = 1 Then
                    strSql = "SELECT DISTINCT a.family_cd FROM " & TB3 & " a WHERE a.lifeno = " & SText(Range("J4").Value) & " AND a.relations NOT IN (""부"",""모"")" & _
                                " UNION SELECT DISTINCT a.family_cd FROM " & TB3 & " a WHERE a.lifeno = " & SText(Range("T4").Value) & " AND a.relations NOT IN (""부"",""모"");"
                    Call makeListData(strSql, TB3)
                    
                    Range("AC26") = LISTDATA(0, 0)
                    Range("AD26") = LISTDATA(1, 0)
                    strSql = "SELECT * FROM (SELECT a.family_id,a.family_cd,REPLACE(a.relations,'(별세)','') AS 'relations',a.lifeno,a.name_ko,a.name_en,a.church_sid,a.church_nm,a.title,a.POSITION,a.birthday,a.education,a.religion,a.recognition,a.memo,a.SUSPEND,CONCAT(a.title,if(a.title<>'' AND a.position<>'','/',''),a.position) AS 'status',RANK() OVER (PARTITION BY a.family_cd,relations ORDER BY a.birthday) AS 'rank' FROM op_system.v_familyinfo a WHERE a.family_cd = " & SText(LISTDATA(0, 0)) & _
                            " UNION SELECT a.family_id,a.family_cd,REPLACE(a.relations,'(별세)','') AS 'relations',a.lifeno,a.name_ko,a.name_en,a.church_sid,a.church_nm,a.title,a.POSITION,a.birthday,a.education,a.religion,a.recognition,a.memo,a.SUSPEND,CONCAT(a.title,if(a.title<>'' AND a.position<>'','/',''),a.position) AS 'status',RANK() OVER (PARTITION BY a.family_cd,relations ORDER BY a.birthday) AS 'rank' FROM op_system.v_familyinfo a WHERE a.family_cd = " & SText(LISTDATA(1, 0)) & " ORDER BY family_cd,FIELD(relations,'부','모','형제','자매'),birthday) a WHERE a.lifeno NOT IN (" & SText(Range("J4").Value) & "," & SText(Range("T4").Value) & ");"
                ElseIf cntRecord = 0 Then
                    Range("AC26") = LISTDATA(0, 0)
                    strSql = "SELECT * FROM (SELECT a.family_id,a.family_cd,REPLACE(a.relations,'(별세)','') AS 'relations',a.lifeno,a.name_ko,a.name_en,a.church_sid,a.church_nm,a.title,a.POSITION,a.birthday,a.education,a.religion,a.recognition,a.memo,a.SUSPEND,CONCAT(a.title,if(a.title<>'' AND a.position<>'','/',''),a.position) AS 'status',RANK() OVER (PARTITION BY a.family_cd,relations ORDER BY a.birthday) AS 'rank' FROM op_system.v_familyinfo a WHERE a.family_cd = " & SText(LISTDATA(0, 0)) & " ORDER BY family_cd,FIELD(relations,'부','모','형제','자매'),birthday) a WHERE a.lifeno NOT IN (" & SText(Range("J4").Value) & ");"
                End If
            ElseIf cntRecord > 2 Then
                MsgBox "선지자 혹은 사모 가족정보 데이터에 중복오류가 있습니다. 중복된 자료를 제거하세요.", vbCritical, banner
            End If
        End If
        
        strSql = strSql & ";"
    Case Else
        '//화폐id, 화폐약칭, 조회일, 원화환율, 달러화환율
        'strSQL = "SELECT a.currency_id, a.currency_un, a.refer_dt, a.fx_rate_krw, a.fx_rate_usd " & _
                      "FROM " & TB1 & " a WHERE a.user_id = " & user_id & ";"
    End Select
    makeSelectSQL = strSql
End Function
Private Sub sbClearVariant()
    Erase LISTFIELD
    Erase LISTDATA
    cntRecord = Empty
    strSql = vbNullString
End Sub

