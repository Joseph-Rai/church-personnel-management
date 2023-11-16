VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Search_by_TitlePosition 
   Caption         =   "직분직책별 검색 마법사"
   ClientHeight    =   8580.001
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4800
   OleObjectBlob   =   "frm_Search_by_TitlePosition.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "frm_Search_by_TitlePosition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cntlst1Col As Integer '//리스트 컬럼 수
Dim TB1 As String, TB2 As String '//폼에 연결된 DB 테이블
Const strOrderBy As String = "currency_un ASC" '//DB에서 sort_order 필드
Dim caseSave As Integer '//1: 추가, 2: 수정, 3: 삭제(SUSPEND), 4: 완전삭제
Dim queryKey As Integer '//리스트 위치 반환에 사용될 id
Dim LISTDATA() As String '--//DB에서 받아온 rs를 배열로 저장
Dim LISTFIELD() As String '--//DB에서 받아온 rs의 필드를 배열로 저장
Dim cntRecord As Integer '--//DB에서 받아온 레코드의 개수
Dim strSql As String '--//SQL 쿼리문
Dim ws As Worksheet

Private Sub cboCountry_Change()
    '--//cboUnion 아이템 조정
    If Me.cboCountry.listIndex <> -1 Then
        strSql = "SELECT DISTINCT a.`연합회` FROM op_system.v_search_titleposition a WHERE a.`선교국가`=" & SText(Me.cboCountry.List(Me.cboCountry.listIndex)) & " AND a.`관리부서` = " & SText(USER_DEPT) & ";"
    Else
        strSql = "SELECT DISTINCT a.`연합회` FROM op_system.v_search_titleposition a WHERE a.`관리부서` = " & SText(USER_DEPT) & ";"
    End If
    Call makeListData(strSql, TB1)
    
    If Me.cboUnion.listIndex <> -1 Then
        Me.cboUnion.Clear
    End If
    Me.cboUnion.List = LISTDATA
End Sub

Private Sub cboSort1_Change()
    
'    Dim i As Long
'
'    If Me.cboSort1.ListIndex <> -1 Then
'        Me.cboSort2.Enabled = True
'    Else
'        Me.cboSort2.Enabled = False
'    End If
    
    '--//Sync list items of sort combobox
    Call setListItemForSortComboBox
    
End Sub

Private Sub setListItemForSortComboBox()

    Dim listItems As Object
    Set listItems = CreateObject("System.Collections.ArrayList")

    If Me.MultiPage1.Value = 1 Then
        strSql = "SELECT a.`연합회`,a.`선교국가`,a.`교회명(전체)` AS '교회명',a.`교회명` AS `요약교회명`,a.`본교회 정렬순서` AS '본교회 전산순서',a.`지교회 정렬순서` AS '지교회 전산순서',a.`생명번호`,a.`사모한글이름(직분)`,a.`사모영문이름`,a.`사모직분`,a.`사모직책`,a.`배우자 생년월일`," & _
                    "a.`사모국적`,a.`전체1회`,a.`현당회발령일`,a.`(해외)최초발령일`,a.`교회구분` FROM op_system.v_search_titleposition a WHERE a.`관리부서` = " & SText(USER_DEPT) & ";"
        Call makeListData(strSql, TB1)
    Else
        strSql = "SELECT a.`연합회`,a.`선교국가`,a.`교회명(전체)` AS '교회명',a.`교회명` AS `요약교회명`,a.`본교회 정렬순서` AS '본교회 전산순서',a.`지교회 정렬순서` AS '지교회 전산순서',a.`생명번호`,a.`한글이름(직분)`,a.`영문이름`,a.`직분`,a.`직책`,a.`생년월일`," & _
                    "a.`국적`,a.`전체1회`,a.`현당회발령일`,a.`(해외)최초발령일`,a.`교회구분` FROM op_system.v_search_titleposition a WHERE a.`관리부서` = " & SText(USER_DEPT) & ";"
        Call makeListData(strSql, TB1)
    End If
    
    Dim tmp As Variant
    For Each tmp In LISTFIELD
        listItems.Add tmp
    Next
    
    Dim cboBox As MSForms.control
    For Each cboBox In Me.Frame_Sort.controls
        If TypeName(cboBox) = "ComboBox" Then
            If cboBox.Value <> "" Then
                listItems.Remove cboBox.Value
            End If
        End If
    Next
    
    For Each cboBox In Me.Frame_Sort.controls
        If TypeName(cboBox) = "ComboBox" Then
            cboBox.List = listItems.ToArray
        End If
    Next
    
End Sub

Private Sub cboSort2_Change()
    
    '--//Sync list items of sort combobox
    Call setListItemForSortComboBox
    
End Sub

Private Sub cboSort3_Change()

    '--//Sync list items of sort combobox
    Call setListItemForSortComboBox

End Sub

Private Sub cboSort4_Change()
        
    '--//Sync list items of sort combobox
    Call setListItemForSortComboBox
    
End Sub

Private Sub MultiPage1_Change()
    
    Dim i As Long
    Dim chkBox As MSForms.control
    
    '--//기존에 선택한 체크박스 초기화
    If Me.MultiPage1.Value = 1 Then
        For Each chkBox In Me.Frame_Position.controls
            If TypeName(chkBox) = "CheckBox" Then
                chkBox.Value = 0
            End If
        Next
        For Each chkBox In Me.Frame_Title.controls
            If TypeName(chkBox) = "CheckBox" Then
                chkBox.Value = 0
            End If
        Next
        For Each chkBox In Me.Frame_Sort.controls
            If TypeName(chkBox) = "ComboBox" Then
                chkBox.Value = ""
            End If
        Next
    ElseIf Me.MultiPage1.Value = 0 Then
        For Each chkBox In Me.Frame_Position_Spouse.controls
            If TypeName(chkBox) = "CheckBox" Then
                chkBox.Value = 0
            End If
        Next
        For Each chkBox In Me.Frame_Title_Spouse.controls
            If TypeName(chkBox) = "CheckBox" Then
                chkBox.Value = 0
            End If
        Next
        For Each chkBox In Me.Frame_Sort.controls
            If TypeName(chkBox) = "ComboBox" Then
                chkBox.Value = ""
            End If
        Next
    End If
    
    '--//사모기준에 따른 국적 조정
    If Me.MultiPage1.Value = 1 Then
        '--//cboNationality 아이템 추가
        strSql = "SELECT DISTINCT a.`사모국적` FROM op_system.v_search_titleposition a WHERE a.`사모국적` is not null AND a.`관리부서` = " & SText(USER_DEPT) & " ORDER BY a.`사모국적`"
        Call makeListData(strSql, TB1)
        Me.cboNationality.List = LISTDATA
        Me.lblNationality.Caption = "사모국적"
    ElseIf Me.MultiPage1.Value = 0 Then
        '--//cboNationality 아이템 추가
        strSql = "SELECT DISTINCT a.`국적` FROM op_system.v_search_titleposition a WHERE a.`관리부서` = " & SText(USER_DEPT) & " ORDER BY a.`국적`"
        Call makeListData(strSql, TB1)
        Me.cboNationality.List = LISTDATA
        Me.lblNationality.Caption = "국적"
    End If
    
    '--//Sync list items of sort combobox
    Call setListItemForSortComboBox
    
End Sub

Private Sub UserForm_Initialize()
    
    Dim i As Long
    
    If connIP = Empty Then setGlobalVariant '//전역변수 재설정
    If checkLogin = 0 Then f_login.Show '//로그인체크
    If checkLogin = 0 Then End '--//로그인 실패 시 프로시저 종료
    
    '--//시트설정
    Set ws = ActiveSheet
    
    '--//기초설정
    Me.cmdClose.Cancel = True
    TB1 = "op_system.v_search_titleposition" '--//직분직책별 리스트
    
    '--//컨트롤 설정
    Me.cmdClose.Cancel = True
'    Me.cboSort2.Enabled = False '--//정렬기준1이 선택되고 난 이후에 활성화
'    Me.cboSort3.Enabled = False '--//정렬기준2이 선택되고 난 이후에 활성화
'    Me.cboSort4.Enabled = False '--//정렬기준3이 선택되고 난 이후에 활성화
    
    
'--//콤보박스 아이템 추가
    '--//cboCountry 아이템 추가
    strSql = "SELECT DISTINCT a.`선교국가` FROM op_system.v_search_titleposition a WHERE a.`관리부서` = " & SText(USER_DEPT) & " ORDER BY a.`선교국가`"
    Call makeListData(strSql, TB1)
    Me.cboCountry.List = LISTDATA
    
    '--//변수 초기화
    Call sbClearVariant
    
    '--//cboUnion 아이템 추가
    strSql = "SELECT DISTINCT a.`연합회` FROM op_system.v_search_titleposition a WHERE a.관리부서 = " & SText(USER_DEPT) & " ORDER BY a.`연합회`"
    Call makeListData(strSql, TB1)
    Me.cboUnion.List = LISTDATA
    
    '--//cboNationality 아이템 추가
    strSql = "SELECT DISTINCT a.`국적` FROM op_system.v_search_titleposition a WHERE a.관리부서 = " & SText(USER_DEPT) & " ORDER BY a.`국적`"
    Call makeListData(strSql, TB1)
    Me.cboNationality.List = LISTDATA
    
    '--//변수 초기화
    Call sbClearVariant
    
    '--//cboSort1,2,3,4 아이템 추가
    Call setListItemForSortComboBox
    
    '--//Set Multipage 0
    Me.MultiPage1.Value = 0
    
    '--//변수 초기화
    Call sbClearVariant
    
End Sub

Private Sub cmdSearch_Click()
    
    Dim i As Long
    
    '--//유효성 검사
    If isSelected_Posision = False And isSelected_Title = False Then
        MsgBox "검색할 직책 혹은 직분을 적어도 하나 이상 선택해 주세요.", vbCritical, "검색조건 오류"
        Exit Sub
    End If
    
    '--//시트 활성화, 최적화, 잠금해제
    WB_ORIGIN.Activate
    ws.Activate
    Call Optimization
    Call shUnprotect(globalSheetPW)
    
    '--//서식 초기화
    Call sbInitialize_From
    
    '--//strSQL문 기반으로 listData 생성
    strSql = makeSelectSQL(TB1)
    Call makeListData(strSql, TB1)
    
    '--//출력된 Data 개수에 따른 리포트 작성
    Call sbMakeReport
    
    '--//필요한 변수값 저장
    Range("TitlePosition_rngCntRecord") = cntRecord
    Range("TitlePosition_rngCntField") = UBound(LISTFIELD) + 1
    Range("TitlePosition_rngSearchCode") = Me.MultiPage1.Value
    
    '--//반환된 ListData를 보고서 시트에 삽입
    Optimization
    If cntRecord > 0 Then
        Range("TitlePosition_rngTarget").Resize(cntRecord, UBound(LISTFIELD) + 1) = LISTFIELD
        Range("TitlePosition_rngTarget").Offset(1).Resize(cntRecord, UBound(LISTFIELD) + 1) = LISTDATA
    End If
    
    Normal
    
    '--//사진삽입
    Call sbInsertPic
    
    '--//조회기준일 등록
    Range("TitlePosition_Date") = Format(DateSerial(year(Date), month(Date) - 1, 1), "yyyy년 mm월")
    
    '--//정렬기준 등록
    Range("TitlePosition_rngSort").ClearContents
    If Me.cboSort1.Value <> "" Then
        Range("TitlePosition_rngSort") = "정렬기준: 1. " & Me.cboSort1.Value
    End If
    If Me.cboSort2.Value <> "" Then
        Range("TitlePosition_rngSort") = Range("TitlePosition_rngSort") & ", 2. " & Me.cboSort2.Value
    End If
    If Me.cboSort3.Value <> "" Then
        Range("TitlePosition_rngSort") = Range("TitlePosition_rngSort") & ", 3. " & Me.cboSort3.Value
    End If
    If Me.cboSort4.Value <> "" Then
        Range("TitlePosition_rngSort") = Range("TitlePosition_rngSort") & ", 4. " & Me.cboSort4.Value
    End If
    
    '--//계산 한 번 돌려주고
    Application.CalculateFullRebuild
    Range("D3").Select
    
    '--//인쇄영역 설정
    ActiveSheet.PageSetup.PrintArea = Range(Cells(1, "D"), Cells(Cells(Rows.Count, "A").End(xlUp).Row, "O")).Address
    
    '--//제목설정
    Call sbMakeTitle
    If InStr(Range("TitlePosition_rngTitle"), "직책조건") > 0 Then
        With Range("TitlePosition_rngTitle").Characters(Start:=WorksheetFunction.Find("직책조건", Range("TitlePosition_rngTitle")), Length:=Len(Range("TitlePosition_rngTitle")) - 16).Font
            .color = vbBlue
            .FontStyle = "굵은 기울임꼴"
            .Size = 12
        End With
    Else
        With Range("TitlePosition_rngTitle").Characters(Start:=WorksheetFunction.Find("직분조건", Range("TitlePosition_rngTitle")), Length:=Len(Range("TitlePosition_rngTitle")) - 16).Font
            .color = vbBlue
            .FontStyle = "굵은 기울임꼴"
            .Size = 12
        End With
    End If
    With Range("TitlePosition_rngTitle").Characters(Start:=InStrRev(Range("TitlePosition_rngTitle"), " "), Length:=Len(Range("TitlePosition_rngTitle")) - InStrRev(Range("TitlePosition_rngTitle"), " ")).Font
        .color = vbRed
    End With
    
    '--//변수 초기화
    Call sbClearVariant
    
    shProtect globalSheetPW
    Call Normal
End Sub

Private Function isSelected_Posision()

    Dim control As MSForms.control
    Dim result As Boolean
    
    result = False
    
    If Me.MultiPage1.Value = 0 Then
        For Each control In Me.Frame_Position.controls
            If TypeName(control) = "CheckBox" Then
                If control.Value = True Then
                    result = True
                    Exit For
                End If
            End If
        Next
    Else
        For Each control In Me.Frame_Position_Spouse.controls
            If TypeName(control) = "CheckBox" Then
                If control.Value = True Then
                    result = True
                    Exit For
                End If
            End If
        Next
    End If
    
    isSelected_Posision = result

End Function

Private Function isSelected_Title()

    Dim control As MSForms.control
    Dim result As Boolean
    
    result = False
    
    If MultiPage1.Value = 0 Then
        For Each control In Me.Frame_Title.controls
            If TypeName(control) = "CheckBox" Then
                If control.Value = True Then
                    result = True
                    Exit For
                End If
            End If
        Next
    Else
        For Each control In Me.Frame_Title_Spouse.controls
            If TypeName(control) = "CheckBox" Then
                If control.Value = True Then
                    result = True
                    Exit For
                End If
            End If
        Next
    End If
    
    isSelected_Title = result

End Function

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub makeListData(ByVal strSql As String, ByVal tableNM As String)
    '#################################################
    'DB에서 받아온 데이터를 listData 배열로 지정합니다.
    '#################################################
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

Private Function makeSelectSQL(ByVal tableNM As String) As String
    '########################################
    '선택된 조건에 따라 Select문을 작성합니다.
    'makeSelectSQL(테이블명)
    '########################################
    Dim conPosition As String
    Dim conTitle As String
    Dim conSub As String
    Dim conSort As String
    Dim WhereClause As String
    Dim OrderClause As String
    
    Select Case tableNM
    Case TB1
        '--//조건별 SelectSQL문 생성
        strSql = "SELECT * FROM (SELECT a.* FROM " & TB1 & " a UNION " & _
                    "SELECT b.`배우자생번`,b.`교회명`,b.`영문교회명`,b.`지교회명`,b.`영문지교회명`,b.`선교국가`,b.`사모한글이름(직분)`,b.`사모영문이름`,b.`사모직책`,b.`사모직책2`,b.`배우자 생년월일`,b.`사모국적`,b.`(해외)최초발령일`,b.`현당회발령일`,NULL,NULL,NULL,NULL,NULL,NULL,NULL,b.`사모직분`,NULL,NULL,b.`연합회`,b.`전체1회`,b.`학생1회`,b.`지교회전체1회`,b.`지교회학생1회`,b.`관리지교회`,b.`관리예배소`,b.`동역`,b.`지교회관리자`,b.`예배소관리자`,b.`예비생도`,NULL,b.`사모직책2시작일`,NULL,b.`전체1회(2달 전)`,b.`학생1회(2달 전)`,b.`지교회전체1회(2달 전)`,b.`지교회학생1회(2달 전)`,b.`관리부서`,b.`연합회 정렬순서`,b.`본교회 정렬순서`,b.`교회구분`,b.`지교회 정렬순서`,b.`교회명(전체)`, b.`지교회역할` FROM " & TB1 & " b WHERE b.`사모직책2` <> '') a "
        
        '--//목회자 직책에 따른
        If Me.chkOverseer.Value Then
            If conPosition = "" Then
                conPosition = "a.`직책`='당회장'"
            Else
                conPosition = conPosition & " OR a.`직책`='당회장'"
            End If
        End If
        
        If Me.chkOverseer_Temp.Value Then
            If conPosition = "" Then
                conPosition = "a.`직책`='당회장대리'"
            Else
                conPosition = conPosition & " OR a.`직책`='당회장대리'"
            End If
        End If
        
        If Me.chkAssistant.Value Then
            If conPosition = "" Then
                conPosition = "a.`직책`='동역'"
            Else
                conPosition = conPosition & " OR a.`직책`='동역'"
            End If
        End If
        
        If Me.chkTheological.Value Then
            If conPosition = "" Then
                conPosition = "a.`생도기수` is not null"
            Else
                conPosition = conPosition & " OR a.`생도기수` is not null"
            End If
        End If
        
        If Me.chkBCLeader.Value Then
            If conPosition = "" Then
                conPosition = "a.`직책`='지교회관리자'"
            Else
                conPosition = conPosition & " OR a.`직책`='지교회관리자'"
            End If
        End If
        
        If Me.chkPBCLeader.Value Then
            If conPosition = "" Then
                conPosition = "a.`직책`='예배소관리자'"
            Else
                conPosition = conPosition & " OR a.`직책`='예배소관리자'"
            End If
        End If
        
        If Me.chkBuildingManager.Value Then
            If conPosition = "" Then
                conPosition = "a.`직책2` LIKE '%건물관리%'"
            Else
                conPosition = conPosition & " OR a.`직책2` LIKE '%건물관리%'"
            End If
        End If
        
        If Me.chkTranslator.Value Then
            If conPosition = "" Then
                conPosition = "a.`직책2` LIKE '%번역%'"
            Else
                conPosition = conPosition & " OR a.`직책2` LIKE '%번역%'"
            End If
        End If
        
        If Me.chkGeneralAffair.Value Then
            If conPosition = "" Then
                conPosition = "a.`직책2` LIKE '%행정%'"
            Else
                conPosition = conPosition & " OR a.`직책2` LIKE '%행정%'"
            End If
        End If
        
        If Me.chkMission.Value Then
            If conPosition = "" Then
                conPosition = "a.`직책2` LIKE '%자비량%'"
            Else
                conPosition = conPosition & " OR a.`직책2` LIKE '%자비량%'"
            End If
        End If
        
        '--//사모 직책에 따른
        If Me.chkOverseerWife.Value Then
            If conPosition = "" Then
                conPosition = "a.`사모직책`='당사모'"
            Else
                conPosition = conPosition & " OR a.`사모직책`='당사모'"
            End If
        End If
        
        If Me.chkOverseerWife_Temp.Value Then
            If conPosition = "" Then
                conPosition = "a.`사모직책`='당대리사모'"
            Else
                conPosition = conPosition & " OR a.`사모직책`='당대리사모'"
            End If
        End If
        
        If Me.chkAssistantWife.Value Then
            If conPosition = "" Then
                conPosition = "a.`사모직책`='동사모'"
            Else
                conPosition = conPosition & " OR a.`사모직책`='동사모'"
            End If
        End If
        
        If Me.chkTheologicalWife.Value Then
            If conPosition = "" Then
                conPosition = "a.`사모직책`='생도사모'"
            Else
                conPosition = conPosition & " OR a.`사모직책`='생도사모'"
            End If
        End If
        
        If Me.chkBCLeaderWife.Value Then
            If conPosition = "" Then
                conPosition = "a.`사모직책`='지관자사모'"
            Else
                conPosition = conPosition & " OR a.`사모직책`='지관자사모'"
            End If
        End If
        
        If Me.chkPBCLeaderWife.Value Then
            If conPosition = "" Then
                conPosition = "a.`사모직책`='예관자사모'"
            Else
                conPosition = conPosition & " OR a.`사모직책`='예관자사모'"
            End If
        End If
        
        '--//목회자 직분에 따른
        If Me.chkPastor.Value Then
            If conTitle = "" Then
                conTitle = "a.`직분`='목사'"
            Else
                conTitle = conTitle & " OR a.`직분`='목사'"
            End If
        End If
        
        If Me.chkElder.Value Then
            If conTitle = "" Then
                conTitle = "a.`직분`='장로'"
            Else
                conTitle = conTitle & " OR a.`직분`='장로'"
            End If
        End If
        
        If Me.chkMissionary.Value Then
            If conTitle = "" Then
                conTitle = "a.`직분`='전도사'"
            Else
                conTitle = conTitle & " OR a.`직분`='전도사'"
            End If
        End If
        
        If Me.chkDeacon.Value Then
            If conTitle = "" Then
                conTitle = "a.`직분`='집사'"
            Else
                conTitle = conTitle & " OR a.`직분`='집사'"
            End If
        End If
        
        If Me.chkBrother.Value Then
            If conTitle = "" Then
                conTitle = "a.`직분` is null"
            Else
                conTitle = conTitle & " OR a.`직분` is null"
            End If
        End If
        
        '--//사모 직분에 따른
        If Me.chkSeniorDeaconess.Value Then
            If conTitle = "" Then
                conTitle = "a.`사모직분`='권사'"
            Else
                conTitle = conTitle & " OR a.`사모직분`='권사'"
            End If
        End If
        
        If Me.chkMissionaryF.Value Then
            If conTitle = "" Then
                conTitle = "a.`사모직분`='전도사'"
            Else
                conTitle = conTitle & " OR a.`사모직분`='전도사'"
            End If
        End If
        
        If Me.chkDeaconess.Value Then
            If conTitle = "" Then
                conTitle = "a.`사모직분`='집사'"
            Else
                conTitle = conTitle & " OR a.`사모직분`='집사'"
            End If
        End If
        
        If Me.chkSister.Value Then
            If conTitle = "" Then
                conTitle = "a.`사모직분` is null"
            Else
                conTitle = conTitle & " OR a.`사모직분` is null"
            End If
        End If
        
        '--//서브조건에 따른
        If Me.cboCountry.listIndex <> -1 Then
            If conSub = "" Then
                conSub = "a.`선교국가`=" & SText(Me.cboCountry.List(Me.cboCountry.listIndex))
            Else
                conSub = conSub & "AND a.`선교국가`=" & SText(Me.cboCountry.List(Me.cboCountry.listIndex))
            End If
        End If
        
        If Me.cboUnion.listIndex <> -1 Then
            If conSub = "" Then
                conSub = "a.`연합회`=" & SText(Me.cboUnion.List(Me.cboUnion.listIndex))
            Else
                conSub = conSub & "AND a.`연합회`=" & SText(Me.cboUnion.List(Me.cboUnion.listIndex))
            End If
        End If
        
        If Me.MultiPage1.Value = 0 Then '--//목회자 검색일 때
            If Me.cboNationality.listIndex <> -1 Then
                If conSub = "" Then
                    conSub = "a.`국적`=" & SText(Me.cboNationality.List(Me.cboNationality.listIndex))
                Else
                    conSub = conSub & "AND a.`국적`=" & SText(Me.cboNationality.List(Me.cboNationality.listIndex))
                End If
            End If
        ElseIf Me.MultiPage1.Value = 1 Then '--//사모 검색일 때
            If Me.cboNationality.listIndex <> -1 Then
                If conSub = "" Then
                    conSub = "a.`사모국적`=" & SText(Me.cboNationality.List(Me.cboNationality.listIndex))
                Else
                    conSub = conSub & "AND a.`사모국적`=" & SText(Me.cboNationality.List(Me.cboNationality.listIndex))
                End If
            End If
        Else
        End If
        
        '--//정렬기준에 따른
        conSort = makeOrderByClause(Me.cboSort1, conSort)
        conSort = makeOrderByClause(Me.cboSort2, conSort)
        conSort = makeOrderByClause(Me.cboSort3, conSort)
        conSort = makeOrderByClause(Me.cboSort4, conSort)
        
        '--//WHERE절 구문생성
        If conPosition <> "" Then
            If WhereClause = "" Then
                WhereClause = "WHERE (" & conPosition & ")"
            Else
                WhereClause = WhereClause & " AND (" & conPosition & ")"
            End If
        End If
        
        If conTitle <> "" Then
            If WhereClause = "" Then
                WhereClause = "WHERE (" & conTitle & ")"
            Else
                WhereClause = WhereClause & " AND (" & conTitle & ")"
            End If
        End If
        
        If conSub <> "" Then
            If WhereClause = "" Then
                WhereClause = "WHERE (" & conSub & ")"
            Else
                WhereClause = WhereClause & " AND (" & conSub & ")"
            End If
        End If
        
        If WhereClause = "" Then
            WhereClause = "WHERE" & " `관리부서` = " & SText(USER_DEPT)
        Else
            WhereClause = WhereClause & " AND `관리부서` = " & SText(USER_DEPT)
        End If
        
        '--//ORDER BY절 생성
        If conSort <> "" Then
            OrderClause = " ORDER BY " & conSort
        End If
        
        strSql = strSql & WhereClause & OrderClause & ";"
        
    Case Else
    End Select
    makeSelectSQL = strSql
End Function

Private Function makeOrderByClause(cboList As MSForms.comboBox, conSort As String)

    If cboList.Value <> "" Then
        Select Case cboList.Value
            Case "연합회":
                conSort = AppendText(conSort, "a.`" & cboList.Value & " 정렬순서`")
            Case "전체1회":
                conSort = AppendText(conSort, "a.`" & cboList.Value & "`" & " DESC, a.`전체1회(2달 전)` DESC")
            Case "요약교회명":
                conSort = AppendText(conSort, "a.`" & Replace(cboList.Value, "요약", "") & "`")
            Case "교회명":
                conSort = AppendText(conSort, "a.`교회명(전체)`")
            Case "본교회 전산순서"
                conSort = AppendText(conSort, "a.`" & Replace(cboList.Value, "전산", "정렬") & "`")
            Case "지교회 전산순서"
                conSort = AppendText(conSort, "a.`" & Replace(cboList.Value, "전산", "정렬") & "`")
            Case "직분", "사모직분":
                If Me.chkAssistantWife.Value Then
                    conSort = AppendText(conSort, "a." & "`사모직분` IS NULL ASC, FIELD(`사모직분`,'목사','장로','권사','전도사','집사','형제','자매','')")
                Else
                    conSort = AppendText(conSort, "a." & "`직분` IS NULL ASC, FIELD(`직분`,'목사','장로','권사','전도사','집사','형제','자매','')")
                End If
            Case "직책", "사모직책":
                If Me.chkAssistantWife.Value Then
                    conSort = AppendText(conSort, "a." & "`사모직책` IS NULL ASC, FIELD(`사모직책`,'당사모','당대리사모','동사모','지관자사모','예관자사모','생도사모'," & getPosition2Joining & ",'')")
                Else
                    conSort = AppendText(conSort, "a." & "`직책` IS NULL ASC, FIELD(`직책`,'당회장','당회장대리','동역','지교회관리자','예배소관리자','예비생도1단계','예비생도2단계','예비생도3단계','생도사모'," & getPosition2Joining & ",'')")
                End If
            Case "교회구분":
                conSort = AppendText(conSort, "FIELD(`교회구분`,'지교회','예배소','')")
            Case Else
                conSort = AppendText(conSort, "a.`" & cboList.Value & "`")
        End Select
    End If
    
    makeOrderByClause = conSort

End Function

Private Function AppendText(sourceText As String, targetText As String)

    Dim result As String
    
    If sourceText = "" Then
        result = targetText
    Else
        result = sourceText & ", " & targetText
    End If
    
    AppendText = result

End Function

Public Sub sbInsertPic()
    '###############################
    '지정된 위치에 사진을 삽입합니다.
    '###############################
    Dim lifeNo As String
    
    ActiveSheet.Pictures.Delete

    '--//사진 넣기 프로세스
On Error Resume Next
    Dim i As Long, j As Long
    For i = 4 To Cells(Rows.Count, "A").End(xlUp).Row Step 7: For j = 5 To 14 Step 2
        '--//변수설정
        lifeNo = Cells(i, j).Value
        
        '--//사진삽입
        If lifeNo <> "" Then
            InsertPStaffPic lifeNo, Cells(i, j).Resize(4)
        End If
    Next j: Next i
    
    '--//마지막에 가짜사진 하나 추가(사진틀어짐 버그 방지)
    InsertPStaffPic lifeNo, Cells(1, "D")
    
    '--//사진 틀어짐 방지를 위해 마지막 사진 삭제
    If ActiveSheet.Pictures.Count > 0 Then
        ActiveSheet.Pictures(ActiveSheet.Pictures.Count).Delete
    End If
    
On Error GoTo 0

End Sub
Private Sub sbClearVariant()
    '##########################
    '각종 변수를 초기화 합니다.
    '##########################
    Erase LISTFIELD
    Erase LISTDATA
    cntRecord = Empty
    strSql = vbNullString
End Sub
Private Sub sbSortData_By_Union()
    '######################################################################
    '정렬기준에 연합회가 포함되어 있을 경우 지정된 연합회 순으로 정렬합니다.
    '######################################################################
    
    Select Case USER_DEPT
    Case 10 '--//아시아3과
        ActiveWorkbook.Worksheets("직분직책별 현황").Sort.SortFields.Clear
        ActiveWorkbook.Worksheets("직분직책별 현황").Sort.SortFields.Add key:=Range("TitlePosition_rngTarget").Offset(1, 24).Resize(Range("TitlePosition_rngTarget").CurrentRegion.Rows.Count - 1) _
            , SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder:= _
            "카트만두,네팔동부,네팔중부,네팔서부" _
            , DataOption:=xlSortNormal
        With ActiveWorkbook.Worksheets("직분직책별 현황").Sort
            .SetRange Range("TitlePosition_rngTarget").CurrentRegion
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    Case Else
        ActiveWorkbook.Worksheets("직분직책별 현황").Sort.SortFields.Clear
        ActiveWorkbook.Worksheets("직분직책별 현황").Sort.SortFields.Add key:=Range("TitlePosition_rngTarget").Offset(1, 24).Resize(Range("TitlePosition_rngTarget").CurrentRegion.Rows.Count - 1) _
            , SortOn:=xlSortOnValues, Order:=xlAscending _
            , DataOption:=xlSortNormal
        With ActiveWorkbook.Worksheets("직분직책별 현황").Sort
            .SetRange Range("TitlePosition_rngTarget").CurrentRegion
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    End Select
End Sub

Public Sub sbInitialize_From()
    '############################
    '리포트 양식을 초기화 합니다.
    '############################
    Dim R As Long
    
    '--//기존 데이터 삭제
    Range("TitlePosition_rngTarget").CurrentRegion.Offset(1).ClearContents
    
    '--//마지막 행 찾기
    R = Cells(Rows.Count, "A").End(xlUp).Row
    Range("10:10").Resize(R).Delete Shift:=xlUp

End Sub

Private Sub sbMakeReport()
    '################################################
    '검색된 데이터 개수에 따른 리포트 양식 조정합니다.
    '################################################
    
    Dim rngTarget As Range
    Dim cntRow As Long
    
    '--//삽입에 필요한 행 개수
    cntRow = ((cntRecord - 1) \ 5) * 7
    
    '--//첫 줄을 제외한 리포트 범위
    If cntRow <> 0 Then
        Set rngTarget = Range("10:10").Resize(cntRow)
    
    
        '--//첫 줄 복사
        Range("3:9").Copy
    
        '--//rngTarget에 붙여넣기(1. 서식, 2. 수시)
        rngTarget.PasteSpecial Paste:=xlPasteFormats
        rngTarget.PasteSpecial Paste:=xlPasteFormulas
    End If
    
End Sub

Private Sub sbMakeTitle()

    '##################################
    '조건별로 리포트 제목을 생성합니다.
    '##################################
    
    Dim strTitle As String
    Dim conPosition As String
    Dim conTitle As String
    Dim conSub As String
    
    strSql = "SELECT a.dept_nm FROM op_system.db_ovs_dept a WHERE a.dept_id = " & SText(USER_DEPT)
    Call makeListData(strSql, "op_system.db_ovs_dept")
    
    strTitle = LISTDATA(0, 0) & " 조건별 목회자 명단" & vbNewLine
    
    '--//목회자 직책에 따른
    If Me.chkOverseer.Value Then
        If conPosition = "" Then
            conPosition = "당회장" & "(" & WorksheetFunction.CountIf(Range("AD:AD"), "당회장") & ")"
        Else
            conPosition = conPosition & ",당회장" & "(" & WorksheetFunction.CountIf(Range("AD:AD"), "당회장") & ")"
        End If
    End If
    
    If Me.chkOverseer_Temp.Value Then
        If conPosition = "" Then
            conPosition = "당대리" & "(" & WorksheetFunction.CountIf(Range("AD:AD"), "당회장대리") & ")"
        Else
            conPosition = conPosition & ",당대리" & "(" & WorksheetFunction.CountIf(Range("AD:AD"), "당회장대리") & ")"
        End If
    End If
    
    If Me.chkAssistant.Value Then
        If conPosition = "" Then
            conPosition = "동역" & "(" & WorksheetFunction.CountIf(Range("AD:AD"), "동역") & ")"
        Else
            conPosition = conPosition & ",동역" & "(" & WorksheetFunction.CountIf(Range("AD:AD"), "동역") & ")"
        End If
    End If
    
    If Me.chkTheological.Value Then
        If conPosition = "" Then
            conPosition = "예비생도" & "(" & WorksheetFunction.CountIf(Range("AP:AP"), "*예비생도*") & ")"
        Else
            conPosition = conPosition & ",예비생도" & "(" & WorksheetFunction.CountIf(Range("AP:AP"), "*예비생도*") & ")"
        End If
    End If
    
    If Me.chkBCLeader.Value Then
        If conPosition = "" Then
            conPosition = "지관자" & "(" & WorksheetFunction.CountIf(Range("AD:AD"), "지교회관리자") & ")"
        Else
            conPosition = conPosition & ",지관자" & "(" & WorksheetFunction.CountIf(Range("AD:AD"), "지교회관리자") & ")"
        End If
    End If
    
    If Me.chkPBCLeader.Value Then
        If conPosition = "" Then
            conPosition = "예관자" & "(" & WorksheetFunction.CountIf(Range("AD:AD"), "예배소관리자") & ")"
        Else
            conPosition = conPosition & ",예관자" & "(" & WorksheetFunction.CountIf(Range("AD:AD"), "예배소관리자") & ")"
        End If
    End If
    
    If Me.chkBuildingManager.Value Then
        If conPosition = "" Then
            conPosition = "건물관리" & "(" & WorksheetFunction.CountIf(Range("AE:AE"), "*건물관리*") & ")"
        Else
            conPosition = conPosition & ",건물관리" & "(" & WorksheetFunction.CountIf(Range("AE:AE"), "*건물관리*") & ")"
        End If
    End If
    
    If Me.chkTranslator.Value Then
        If conPosition = "" Then
            conPosition = "번역자" & "(" & WorksheetFunction.CountIf(Range("AE:AE"), "*번역자*") & ")"
        Else
            conPosition = conPosition & ",번역자" & "(" & WorksheetFunction.CountIf(Range("AE:AE"), "*번역자*") & ")"
        End If
    End If
    
    If Me.chkGeneralAffair.Value Then
        If conPosition = "" Then
            conPosition = "행정직원" & "(" & WorksheetFunction.CountIf(Range("AE:AE"), "*행정직원*") & ")"
        Else
            conPosition = conPosition & ",행정직원" & "(" & WorksheetFunction.CountIf(Range("AE:AE"), "*행정직원*") & ")"
        End If
    End If
    
    If Me.chkMission.Value Then
        If conPosition = "" Then
            conPosition = "자비량" & "(" & WorksheetFunction.CountIf(Range("AE:AE"), "*자비량*") & ")"
        Else
            conPosition = conPosition & ",자비량" & "(" & WorksheetFunction.CountIf(Range("AE:AE"), "*자비량*") & ")"
        End If
    End If
    
    '--//사모 직책에 따른
    If Me.chkOverseerWife.Value Then
        If conPosition = "" Then
            conPosition = "당사모" & "(" & WorksheetFunction.CountIf(Range("AM:AM"), "당사모") & ")"
        Else
            conPosition = conPosition & ",당사모" & "(" & WorksheetFunction.CountIf(Range("AM:AM"), "당사모") & ")"
        End If
    End If
    
    If Me.chkOverseerWife_Temp.Value Then
        If conPosition = "" Then
            conPosition = "당대리사모" & "(" & WorksheetFunction.CountIf(Range("AM:AM"), "당대리사모") & ")"
        Else
            conPosition = conPosition & ",당대리사모" & "(" & WorksheetFunction.CountIf(Range("AM:AM"), "당대리사모") & ")"
        End If
    End If
    
    If Me.chkAssistantWife.Value Then
        If conPosition = "" Then
            conPosition = "동사모" & "(" & WorksheetFunction.CountIf(Range("AM:AM"), "동사모") & ")"
        Else
            conPosition = conPosition & ",동사모" & "(" & WorksheetFunction.CountIf(Range("AM:AM"), "동사모") & ")"
        End If
    End If
    
    If Me.chkTheologicalWife.Value Then
        If conPosition = "" Then
            conPosition = "생도사모" & "(" & WorksheetFunction.CountIfs(Range("AM:AM"), "*생도사모*", Range("AJ:AJ"), "<>""""") & ")"
        Else
            conPosition = conPosition & ",생도사모" & "(" & WorksheetFunction.CountIfs(Range("AM:AM"), "*생도사모*", Range("AJ:AJ"), "<>""""") & ")"
        End If
    End If
    
    If Me.chkBCLeaderWife.Value Then
        If conPosition = "" Then
            conPosition = "지관자사모" & "(" & WorksheetFunction.CountIf(Range("AM:AM"), "지관자사모") & ")"
        Else
            conPosition = conPosition & ",지관자사모" & "(" & WorksheetFunction.CountIf(Range("AM:AM"), "지관자사모") & ")"
        End If
    End If
    
    If Me.chkPBCLeaderWife.Value Then
        If conPosition = "" Then
            conPosition = "예관자사모" & "(" & WorksheetFunction.CountIf(Range("AM:AM"), "예관자사모") & ")"
        Else
            conPosition = conPosition & ",예관자사모" & "(" & WorksheetFunction.CountIf(Range("AM:AM"), "예관자사모") & ")"
        End If
    End If
    
    '--//목회자 직분에 따른
    If Me.chkPastor.Value Then
        If conTitle = "" Then
            conTitle = "목사" & "(" & WorksheetFunction.CountIf(Range("AQ:AQ"), "목사") & ")"
        Else
            conTitle = conTitle & ",목사" & "(" & WorksheetFunction.CountIf(Range("AQ:AQ"), "목사") & ")"
        End If
    End If
    
    If Me.chkElder.Value Then
        If conTitle = "" Then
            conTitle = "장로" & "(" & WorksheetFunction.CountIf(Range("AQ:AQ"), "장로") & ")"
        Else
            conTitle = conTitle & ",장로" & "(" & WorksheetFunction.CountIf(Range("AQ:AQ"), "장로") & ")"
        End If
    End If
    
    If Me.chkMissionary.Value Then
        If conTitle = "" Then
            conTitle = "전도사" & "(" & WorksheetFunction.CountIf(Range("AQ:AQ"), "전도사") & ")"
        Else
            conTitle = conTitle & ",전도사" & "(" & WorksheetFunction.CountIf(Range("AQ:AQ"), "전도사") & ")"
        End If
    End If
    
    If Me.chkDeacon.Value Then
        If conTitle = "" Then
            conTitle = "집사" & "(" & WorksheetFunction.CountIf(Range("AQ:AQ"), "집사") & ")"
        Else
            conTitle = conTitle & ",집사" & "(" & WorksheetFunction.CountIf(Range("AQ:AQ"), "집사") & ")"
        End If
    End If
    
    If Me.chkBrother.Value Then
        If conTitle = "" Then
            conTitle = "형제" & "(" & WorksheetFunction.CountA(Range("V:V")) - WorksheetFunction.CountA(Range("AQ:AQ")) & ")"
        Else
            conTitle = conTitle & ",형제" & "(" & WorksheetFunction.CountA(Range("V:V")) - WorksheetFunction.CountA(Range("AQ:AQ")) & ")"
        End If
    End If
    
    '--//사모 직분에 따른
    If Me.chkSeniorDeaconess.Value Then
        If conTitle = "" Then
            conTitle = "권사" & "(" & WorksheetFunction.CountIf(Range("AR:AR"), "권사") & ")"
        Else
            conTitle = conTitle & ",권사" & "(" & WorksheetFunction.CountIf(Range("AR:AR"), "권사") & ")"
        End If
    End If
    
    If Me.chkMissionaryF.Value Then
        If conTitle = "" Then
            conTitle = "전도사" & "(" & WorksheetFunction.CountIf(Range("AR:AR"), "전도사") & ")"
        Else
            conTitle = conTitle & ",전도사" & "(" & WorksheetFunction.CountIf(Range("AR:AR"), "전도사") & ")"
        End If
    End If
    
    If Me.chkDeaconess.Value Then
        If conTitle = "" Then
            conTitle = "집사" & "(" & WorksheetFunction.CountIf(Range("AR:AR"), "집사") & ")"
        Else
            conTitle = conTitle & ",집사" & "(" & WorksheetFunction.CountIf(Range("AR:AR"), "집사") & ")"
        End If
    End If
    
    If Me.chkSister.Value Then
        If conTitle = "" Then
            conTitle = "자매" & "(" & WorksheetFunction.CountA(Range("V:V")) - WorksheetFunction.CountA(Range("AR:AR")) & ")"
        Else
            conTitle = conTitle & ",자매" & "(" & WorksheetFunction.CountA(Range("V:V")) - WorksheetFunction.CountA(Range("AR:AR")) & ")"
        End If
    End If
    
    '--//서브조건에 따른
    If Me.cboCountry.listIndex <> -1 Then
        If conSub = "" Then
            conSub = "국가: " & Me.cboCountry.List(Me.cboCountry.listIndex)
        Else
            conSub = conSub & " / 국가: " & Me.cboCountry.List(Me.cboCountry.listIndex)
        End If
    End If
    
    If Me.cboUnion.listIndex <> -1 Then
        If conSub = "" Then
            conSub = "연합회: " & Me.cboUnion.List(Me.cboUnion.listIndex)
        Else
            conSub = conSub & " / 연합회: " & Me.cboUnion.List(Me.cboUnion.listIndex)
        End If
    End If
    
    If Me.cboNationality.listIndex <> -1 Then
        If conSub = "" Then
            conSub = "국적: " & Me.cboNationality.List(Me.cboNationality.listIndex)
        Else
            conSub = conSub & " / 국적: " & Me.cboNationality.List(Me.cboNationality.listIndex)
        End If
    End If
    
    '--//조건 연결하여 제목생성
    If conPosition <> "" Then
        strTitle = strTitle & _
                    "직책조건: " & conPosition & vbNewLine
    End If
    
    If conTitle <> "" Then
        strTitle = strTitle & _
                    "직분조건: " & conTitle & vbNewLine
    End If
    
    If conSub <> "" Then
        strTitle = strTitle & _
                    conSub & vbNewLine
    End If
    
'    strTitle = Left(strTitle, Len(strTitle) - 1)
    strTitle = strTitle & "검색 총 인원: " & WorksheetFunction.CountA(Range("V:V")) - 1 & "명"
    Range("TitlePosition_rngTitle") = strTitle
    
End Sub

Private Function getPosition2Joining()

    Dim strQuery As String
    strQuery = "SELECT * FROM op_system.a_position2;"
    Call makeListData(strQuery, "op_system.a_position2")
        
    Dim result As String
    Dim i As Integer
    For i = 0 To cntRecord - 1
        If i < cntRecord - 1 Then
            result = result & "'" & LISTDATA(i, 0) & "', "
        Else
            result = result & "'" & LISTDATA(i, 0) & "'"
        End If
    Next
    
    getPosition2Joining = result

End Function
