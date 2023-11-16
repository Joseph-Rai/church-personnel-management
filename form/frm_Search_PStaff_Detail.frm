VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Search_PStaff_Detail 
   Caption         =   "선지자 상세정보 검색마법사"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5850
   OleObjectBlob   =   "frm_Search_PStaff_Detail.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "frm_Search_PStaff_Detail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cntlst1Col As Integer '//리스트 컬럼 수
Dim TB1 As String, TB2 As String, TB3 As String, TB4 As String, TB5 As String, TB6 As String, TB7 As String, TB8 As String, TB9 As String, TB10 As String, TB11 As String, TB12 As String, TB13 As String, TB14 As String '//폼에 연결된 DB 테이블
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
    
    TB3 = "op_system.v_pstaff_detail" '--//목회자 기본정보
    TB4 = "op_system.v_pstaff_detail_title" '--//직분이력
    TB5 = "op_system.v_pstaff_detail_transfer" '--//발령이력
    TB6 = "op_system.v_pstaff_detail_flight" '--//항공스케줄
    TB7 = "op_system.v_pstaff_detail_accomplishment" '--//복음성과
    TB8 = "op_system.v_familyinfo" '--//가족정보
    TB9 = "op_system.v_pstaff_detail_accomplishment_main" '--//복음성과(본교회)
    TB10 = "op_system.v_pstaff_detail_accomplishment_both" '--//복음성과(전체+본교회)
    TB11 = "op_system.v_pstaff_detail_transfer2" '--//발령이력2
    TB12 = "op_system.v_pstaff_detail_concise_transfer_history" '--//역대 교회목록
    TB13 = "op_system.v_pstaff_detail_concise_transfer_history_main" '--//역대 교회목록
    TB14 = "op_system.v_pstaff_detail_concise_transfer_history_both" '--//역대 교회목록
    
    '--//리스트박스 설정
    With Me.lstPStaff
        .ColumnCount = 5
        .ColumnHeads = False
        .ColumnWidths = "0,120,70,50,0" '생명번호, 교회명, 한글이름(직분), 직책,배우자생번
        .TextAlign = fmTextAlignLeft
        .Font = "굴림"
    End With
    
    Me.optAll.Value = True
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
    
    '--//시트 활성화, 최적화, 잠금해제
    WB_ORIGIN.Activate
    ws.Activate
    Call Optimization
    Call shUnprotect(globalSheetPW)
    
    '--//기존 데이터 삭제
    Range("PStaff_Detail_rngTarget").CurrentRegion.ClearContents
    Range("PStaff_Detail_Title").Offset(1).Resize(3, 6).ClearContents
    Range("PStaff_Detail_Transfer").Offset(1).Resize(10, 6).ClearContents
    Range("PStaff_Detail_Flight").Offset(1).Resize(5, 6).ClearContents
    Range("PStaff_Detail_rngAtten").CurrentRegion.ClearContents
    Range("PStaff_Detail_rngFamily").CurrentRegion.ClearContents
    
    '--//선지자 기본정보 삽입
        strSql = makeSelectSQL(TB3)
        connectTaskDB
        Call makeListData(strSql, TB3)
    '    callDBtoRS "cmdOK_Click", TB2, strSQL, Me.Name
        
        '--//반환된 ListData를 보고서 시트에 삽입
        Optimization
        Range("PStaff_Detail_rngTarget").Resize(, UBound(LISTFIELD) + 1) = LISTFIELD
        If cntRecord > 0 Then
            Range("PStaff_Detail_rngTarget").Offset(1).Resize(cntRecord, UBound(LISTFIELD) + 1) = LISTDATA
        End If
        Normal
        
        Call sbClearVariant
        disconnectALL
    
    '--//가족정보 삽입
        strSql = makeSelectSQL(TB8) '--//가족정보
        connectTaskDB
        Call makeListData(strSql, TB8)
    '    callDBtoRS "cmdOK_Click", TB2, strSQL, Me.Name
    
        '--//반환된 ListData를 보고서 시트에 삽입
        Optimization
        Range("PStaff_Detail_rngFamily").Resize(, UBound(LISTFIELD) + 1) = LISTFIELD
        If cntRecord > 0 Then
            Range("PStaff_Detail_rngFamily").Offset(1).Resize(cntRecord, UBound(LISTFIELD) + 1) = LISTDATA
        End If
        Normal
        
        Call sbClearVariant
        disconnectALL
    
    '--//직분 임명이력 삽입
        strSql = makeSelectSQL(TB4) '--//선지자 직분임명 이력
        connectTaskDB
        Call makeListData(strSql, TB4)
    '    callDBtoRS "cmdOK_Click", TB2, strSQL, Me.Name
        
        '--//반환된 ListData를 보고서 시트에 삽입
        Optimization
        Range("PStaff_Detail_Title").Resize(, UBound(LISTFIELD) + 1) = LISTFIELD
        If cntRecord > 0 Then
            Range("PStaff_Detail_Title").Offset(1).Resize(cntRecord, UBound(LISTFIELD) + 1) = LISTDATA
        End If
        Normal
        
        Call sbClearVariant
        disconnectALL
        
        strSql = makeSelectSQL2(TB4) '--//사모 직분임명 이력
        connectTaskDB
        Call makeListData(strSql, TB4)
    '    callDBtoRS "cmdOK_Click", TB2, strSQL, Me.Name
        
        '--//반환된 ListData를 보고서 시트에 삽입
        Optimization
        Range("PStaff_Detail_Title").Offset(, 3).Resize(, UBound(LISTFIELD) + 1) = LISTFIELD
        If cntRecord > 0 Then
            Range("PStaff_Detail_Title").Offset(1, 3).Resize(cntRecord, UBound(LISTFIELD) + 1) = LISTDATA
        End If
        Normal
        
        Call sbClearVariant
        disconnectALL
    
    '--//목회경력 삽입
        Dim objPastroalCareerDao As New PastoralCareerDao
        '--//동역시작일
        
        Range("PStaff_Detail_PHistory").Offset(1) = objPastroalCareerDao.GetMinDateForAssistantOverseer(Range("PStaff_Detail_LifeNo"))
        '--//당회장 시작일
        Range("PStaff_Detail_PHistory").Offset(1, 1) = objPastroalCareerDao.GetMinDateForOverseer(Range("PStaff_Detail_LifeNo"))
        '--//동역경력
        Range("PStaff_Detail_PHistory").Offset(3) = objPastroalCareerDao.GetAssistantOverseerCareer(Range("PStaff_Detail_LifeNo"))
        '--//당회장경력
        Range("PStaff_Detail_PHistory").Offset(3, 1) = objPastroalCareerDao.GetOverseerCareer(Range("PStaff_Detail_LifeNo"))
        '--//목회경력
        Range("PStaff_Detail_PHistory").Offset(3, 5) = objPastroalCareerDao.GetTotalPastoralCareer(Range("PStaff_Detail_LifeNo"))
    
    connectTaskDB
    '--//발령이력 삽입
        If Range("D4") <> 0 Then '--//직책이 있으면
            strSql = "SELECT * FROM (SELECT `발령일`,`직분/직책`,`교회구분`,`교회명`,'',`기간` FROM op_system.v_pstaff_detail_transfer WHERE `생명번호` = " & SText(Range("PStaff_Detail_LifeNo")) & " AND `직분/직책` IS NOT NULL LIMIT 10) a ORDER BY `발령일`;"
            Call makeListData(strSql, TB5)
        Else
            strSql = "SELECT * FROM (SELECT `발령일`,`직분/직책`,`교회구분`,`교회명`,'',`기간` FROM op_system.v_pstaff_detail_transfer WHERE `생명번호` = " & SText(Range("PStaff_Detail_LifeNo")) & " LIMIT 10) a ORDER BY `발령일`;"
            Call makeListData(strSql, TB5)
        End If
        
        '--//반환된 ListData를 보고서 시트에 삽입
        Optimization
        
        Range("PStaff_Detail_Transfer").Resize(, UBound(LISTFIELD) + 1) = LISTFIELD
        If cntRecord > 0 Then
            Range("PStaff_Detail_Transfer").Offset(1).Resize(cntRecord, UBound(LISTFIELD) + 1) = LISTDATA
        End If
        
        Normal
        
        Call sbClearVariant
        disconnectALL
    
    '--//출입국 이력 삽입
        strSql = makeSelectSQL(TB6) '--//선지자 출입국 이력
        connectTaskDB
        Call makeListData(strSql, TB6)
    '    callDBtoRS "cmdOK_Click", TB2, strSQL, Me.Name
        
        '--//반환된 ListData를 보고서 시트에 삽입
        Optimization
        Range("PStaff_Detail_Flight").Resize(, UBound(LISTFIELD) + 1) = LISTFIELD
        If cntRecord > 0 Then
            Range("PStaff_Detail_Flight").Offset(1).Resize(cntRecord, UBound(LISTFIELD) + 1) = LISTDATA
        End If
        Normal
        
        Call sbClearVariant
        disconnectALL
        
        strSql = makeSelectSQL2(TB6) '--//사모 출입국 이력
        connectTaskDB
        Call makeListData(strSql, TB6)
    '    callDBtoRS "cmdOK_Click", TB2, strSQL, Me.Name
        
        '--//반환된 ListData를 보고서 시트에 삽입
        Optimization
        Range("PStaff_Detail_Flight").Offset(, 3).Resize(, UBound(LISTFIELD) + 1) = LISTFIELD
        If cntRecord > 0 Then
            Range("PStaff_Detail_Flight").Offset(1, 3).Resize(cntRecord, UBound(LISTFIELD) + 1) = LISTDATA
        End If
        Normal
        
        Call sbClearVariant
        disconnectALL
    
    Application.CalculateFullRebuild
    
    Range("9:9").EntireRow.AutoFit '--//건강 행높이 자동맞춤
'    Range("15:15").EntireRow.AutoFit '--//가족사항 행높이 자동맞춤
    
    '--//다녀간 교회목록 삽입
    If Me.optAll.Value Then
        strSql = makeSelectSQL(TB12)
    Call makeListData(strSql, TB12)
    ElseIf Me.optMain.Value Then
        strSql = makeSelectSQL(TB13)
    Call makeListData(strSql, TB13)
    ElseIf Me.optBoth Then
        strSql = makeSelectSQL(TB14)
    Call makeListData(strSql, TB14)
    End If
    Range("PStaff_Detail_cntChurch").Offset(1).Resize(15, UBound(LISTFIELD) + 1).ClearContents
    If cntRecord > 0 Then
        Range("PStaff_Detail_cntChurch").Offset(1).Resize(cntRecord, UBound(LISTFIELD)) = LISTDATA
        Range("PStaff_Detail_cntChurch").Offset(1, UBound(LISTFIELD)).Resize(cntRecord).FormulaR1C1 = _
        "=SUMIFS(OFFSET(PStaff_Detail_rngAtten,,2,1000,1),OFFSET(PStaff_Detail_rngAtten,,,1000,1),RC20,OFFSET(PStaff_Detail_rngAtten,,1,1000,1),RC21)"
        Range("PStaff_Detail_cntChurch").Offset(-1).Copy
        Range("PStaff_Detail_cntChurch").Offset(1, 1).Resize(cntRecord, 5).PasteSpecial xlPasteValues, xlPasteSpecialOperationAdd
        Application.CutCopyMode = False
    End If
    
    '--//중복 제거
    Dim curChurchName As String
    Dim prevChurchName As String
    Dim curStartDate As Date
    Dim prevStartDate As Date
    For i = 1 To cntRecord
        curChurchName = Range("PStaff_Detail_cntChurch").Offset(i)
        prevChurchName = Range("PStaff_Detail_cntChurch").Offset(i + 1)
        If curChurchName = prevChurchName Then
            curStartDate = Range("PStaff_Detail_cntChurch").Offset(i, 1)
            prevStartDate = Range("PStaff_Detail_cntChurch").Offset(i + 1, 1)
            If prevStartDate >= curStartDate And prevStartDate <> 0 Then
                '--//2023.09.25 종료일 외에는 조정되면 안됨.
                '--//직분,직책이 조정되면 "담당시 직분직책"에 "현재" 직분직책이 표시됨
'                Range("PStaff_Detail_cntChurch").Offset(i, 1) = Range("PStaff_Detail_cntChurch").Offset(i + 1, 1) '//시작일 조정
'                Range("PStaff_Detail_cntChurch").Offset(i, 6) = Range("PStaff_Detail_cntChurch").Offset(i + 1, 6) '//직분 조정
'                Range("PStaff_Detail_cntChurch").Offset(i, 7) = Range("PStaff_Detail_cntChurch").Offset(i + 1, 7) '//직책 조정
'                Range("PStaff_Detail_cntChurch").Offset(i, 8) = Range("PStaff_Detail_cntChurch").Offset(i + 1, 8) '//교회구분 조정
                Range("PStaff_Detail_cntChurch").Offset(i, 2) = Range("PStaff_Detail_cntChurch").Offset(i + 1, 2) '//종료일 조정
                
                Range("PStaff_Detail_cntChurch").Offset(i, 5) = _
                    DateDiff("m", Range("PStaff_Detail_cntChurch").Offset(i, 1), Range("PStaff_Detail_cntChurch").Offset(i, 2)) '//기간 조정
'                Range("PStaff_Detail_cntChurch").Offset(i + 1).Resize(, UBound(listField) + 1).Delete Shift:=xlUp
                Range("PStaff_Detail_cntChurch").Offset(i + 2).Resize(100, UBound(LISTFIELD) + 1).Copy
                Range("PStaff_Detail_cntChurch").Offset(i + 1).Resize(100, UBound(LISTFIELD) + 1).PasteSpecial Paste:=xlPasteFormulas
                Application.CutCopyMode = False
                cntRecord = cntRecord - 1
            End If
        End If
    Next
    Range("PStaff_Detail_cntChurch").Offset(-3) = cntRecord
    Range("PStaff_Detail_cntChurch").Offset(-3, 2) = UBound(LISTFIELD)
    
    
    '--//복음성과 출석데이터 삽입
    If Me.optAll.Value Then
        strSql = makeSelectSQL(TB7)
        connectTaskDB
        Call makeListData(strSql, TB7)
    ElseIf Me.optMain.Value Then
        strSql = makeSelectSQL(TB9)
        connectTaskDB
        Call makeListData(strSql, TB9)
    ElseIf Me.optBoth Then
        strSql = makeSelectSQL(TB10)
        connectTaskDB
        Call makeListData(strSql, TB10)
    End If
    
    '--//반환된 ListData를 보고서 시트에 삽입
    Optimization
    Range("PStaff_Detail_rngAtten").Resize(, UBound(LISTFIELD) + 1) = LISTFIELD
    If cntRecord > 0 Then
        Range("PStaff_Detail_rngAtten").Offset(1).Resize(cntRecord, UBound(LISTFIELD) + 1) = LISTDATA
    End If
    Normal
    
    Call sbClearVariant
    disconnectALL
    
    On Error Resume Next
    Range("PStaff_Detail_rngAtten").Offset(-1).Copy
    Range("PStaff_Detail_rngAtten").CurrentRegion.Offset(1, 1).PasteSpecial Paste:=xlPasteValues, operation:=xlPasteSpecialOperationAdd
    Range("PStaff_Detail_rngFamily").Offset(1, Range("PStaff_Detail_rngFamily_Rank") - 1).Resize(Range("PStaff_Detail_rngFamily_cntData")).PasteSpecial Paste:=xlPasteValues, operation:=xlPasteSpecialOperationAdd
    Range("PStaff_Detail_rngFamily").Offset(1, 0).Resize(Range("PStaff_Detail_rngFamily_cntData"), 2).PasteSpecial Paste:=xlPasteValues, operation:=xlPasteSpecialOperationAdd
    On Error GoTo 0
    
    Application.CutCopyMode = False
    
    '--//출석 0으로 시작하면 다음 달로 이동
    Dim fieldCount As Integer
    i = 1
    fieldCount = Range("PStaff_Detail_cntChurch").Offset(-3, 2)
    Do While Range("PStaff_Detail_cntChurch").Offset(i) <> ""
        Do While Range("PStaff_Detail_cntChurch").Offset(i, fieldCount) = 0 And _
            Range("PStaff_Detail_cntChurch").Offset(i, 1) <= Range("PStaff_Detail_cntChurch").Offset(i, 2)
                Range("PStaff_Detail_cntChurch").Offset(i, 1) = WorksheetFunction.EDate(Range("PStaff_Detail_cntChurch").Offset(i, 1), 1)
        Loop
        i = i + 1
    Loop
    
    '--//차트정렬
    Application.CalculateFullRebuild
    Call sbArrangeChart_Atten
    
    '--//안쓰는 복음성과 차트 숨기기
    ActiveSheet.Outline.ShowLevels RowLevels:=2
    On Error Resume Next
    Range(Range("PStaff_Detail_Church1").Offset(-1), Range("PStaff_Detail_Church8").Offset(20)).Rows.Ungroup
    Range("PStaff_Detail_Nationality").Resize(3).EntireRow.Rows.Ungroup
    Range("PStaff_Detail_rngFamily").Resize(11).EntireRow.Rows.Ungroup
    Range("PStaff_Detail_PHistory").Resize(4).EntireRow.Rows.Ungroup
    Range("PStaff_Detail_Transfer").Resize(11).EntireRow.Rows.Ungroup
    On Error GoTo 0
    '--//가족영역 그루핑
    For i = 0 To 8
        If Range("PStaff_Detail_Family").Offset(i + 2) = "" And Range("PStaff_Detail_Family").Offset(i + 2, 3) = "" Then
            Range("PStaff_Detail_Family").Offset(i + 2).Rows.Group
        End If
    Next
    On Error Resume Next
'    rngTarget.Rows.Group
    On Error GoTo 0
    '--//목회경력영역 그루핑
    If Not (Range("PStaff_Detail_CurrentPosition") = "당회장" Or _
            Range("PStaff_Detail_CurrentPosition") = "당회장대리" Or _
            Range("PStaff_Detail_CurrentPosition") = "동역" Or _
            Range("PStaff_Detail_CurrentPosition") Like "*관리자*" Or _
            Range("PStaff_Detail_CurrentPosition") Like "*생도*") Then
            
        Range("PStaff_Detail_PHistory").Resize(4).EntireRow.Rows.Group
    End If
    
    '--//발령이력영역 그루핑(최소 4개 행은 보이도록 유지)
    For i = 5 To 10
        If Range("PStaff_Detail_Transfer").Offset(i) = "" Then
            Range("PStaff_Detail_Transfer").Offset(i).EntireRow.Rows.Group
        End If
    Next
    
    '--//비자영역 그루핑
    '--//선지자, 배우자의 선교국가와 국적이 모두 같은 경우에만 그루핑(비자가 필요 없는 경우)
    If Range("PStaff_Detail_GospelCountry") = Range("PStaff_Detail_Nationality").Offset(1) And _
        Range("PStaff_Detail_GospelCountry") = Range("PStaff_Detail_Nationality").Offset(1, 3) Then
        Range(Range("PStaff_Detail_Nationality"), Range("PStaff_Detail_Nationality").Offset(2)).EntireRow.Rows.Group
    End If
    '--//차트영역 그루핑
    Select Case Range("PStaff_Detail_cntChurch").Offset(-3)
    Case 0
        Range(Range("PStaff_Detail_Church1").Offset(-1), Range("PStaff_Detail_Church8").Offset(20)).EntireRow.Rows.Group
    Case 1
        Range(Range("PStaff_Detail_Church2"), Range("PStaff_Detail_Church8").Offset(20)).EntireRow.Rows.Group
    Case 2
        Range(Range("PStaff_Detail_Church3").Offset(-1), Range("PStaff_Detail_Church8").Offset(20)).EntireRow.Rows.Group
    Case 3
        Range(Range("PStaff_Detail_Church4"), Range("PStaff_Detail_Church8").Offset(20)).EntireRow.Rows.Group
    Case 4
        Range(Range("PStaff_Detail_Church5").Offset(-1), Range("PStaff_Detail_Church8").Offset(20)).EntireRow.Rows.Group
    Case 5
        Range(Range("PStaff_Detail_Church6"), Range("PStaff_Detail_Church8").Offset(20)).EntireRow.Rows.Group
    Case 6
        Range(Range("PStaff_Detail_Church7").Offset(-1), Range("PStaff_Detail_Church8").Offset(20)).EntireRow.Rows.Group
    Case 7
        Range(Range("PStaff_Detail_Church8"), Range("PStaff_Detail_Church8").Offset(20)).EntireRow.Rows.Group
    Case Else
    End Select
    ActiveSheet.Outline.ShowLevels RowLevels:=1
    
    '--//프린트영역 조정
    ActiveSheet.PageSetup.Zoom = 100
    Set ActiveSheet.HPageBreaks(1).Location = Range("PStaff_Detail_Church1").Offset(-1)
    
    '--//사진삽입
On Error Resume Next
    ActiveSheet.Pictures.Delete
    
    If Range("PStaff_Detail_LifeNo") <> "" Then
        InsertPStaffPic Range("PStaff_Detail_LifeNo"), Range("PStaff_Detail_Pic_M")
    End If

    If Not (Range("PStaff_Detail_LifeNo_Spouse") = "" Or Range("PStaff_Detail_LifeNo_Spouse") = "0") Then
        InsertPStaffPic Range("PStaff_Detail_LifeNo_Spouse"), Range("PStaff_Detail_Pic_F")
    End If
    
    InsertPStaffPic "", Range("J1")
    
    If ActiveSheet.Pictures.Count > 0 Then
        ActiveSheet.Pictures(ActiveSheet.Pictures.Count).Delete
    End If
Here:
On Error GoTo 0
    
    Sheets("선지자 상세정보").Range("A10").Select
    Sheets("선지자 상세정보").Range("A1").Select
    
    Call shProtect(globalSheetPW)
    Call Normal
    
    MsgBox "작업이 완료되었습니다."
    
End Sub

Private Function GetMinDateForAssistantOverseer(lifeNo As String) As String

    strSql = "" & _
        " SELECT MIN(p.Start_dt)" & _
        " FROM op_system.db_position p" & _
        " WHERE p.lifeno = " & SText(lifeNo) & _
        "     AND p.Position IN ('동역');"
    makeListData strSql, "op_system.db_position"
    
    Dim result As String
    
    If cntRecord > 0 Then
        result = LISTDATA(0, 0)
    Else
        result = ""
    End If
    
    GetMinDateForAssistantOverseer = result

End Function

Private Function GetMinDateForOverseer(lifeNo As String) As String

    strSql = "" & _
        " SELECT MIN(p.Start_dt)" & _
        " FROM op_system.db_position p" & _
        " WHERE p.lifeno = " & SText(lifeNo) & _
        "     AND p.Position IN ('당회장', '당회장대리');"
    makeListData strSql, "op_system.db_position"
    
    Dim result As String
    
    If cntRecord > 0 Then
        result = LISTDATA(0, 0)
    Else
        result = ""
    End If
    
    GetMinDateForOverseer = result

End Function

Private Function GetAssistantOverseerCareer(lifeNo As String) As String

    '--//동역으로 활동한 이력 추출
    strSql = "" & _
        " SELECT p.Start_dt, p.End_dt" & _
        " FROM op_system.db_position p" & _
        " WHERE p.lifeno = " & SText(lifeNo) & _
        "     AND p.Position IN ('동역');"
    makeListData strSql, "op_system.db_position"
    
    Dim result As String
    result = GetConvertedFormatPeriod
    
    GetAssistantOverseerCareer = result

End Function

Private Function GetOverseerCareer(lifeNo As String) As String

    '--//당회장으로 활동한 이력 추출
    strSql = "" & _
        " SELECT p.Start_dt, p.End_dt" & _
        " FROM op_system.db_position p" & _
        " WHERE p.lifeno = " & SText(lifeNo) & _
        "     AND p.Position IN ('당회장', '당회장대리');"
    makeListData strSql, "op_system.db_position"
    
    Dim result As String
    result = GetConvertedFormatPeriod
    
    GetOverseerCareer = result

End Function

Private Function GetTotalPastoralCareer(lifeNo As String) As String
    
    '--//목회자로 활동한 이력 추출
    strSql = "" & _
        " SELECT p.Start_dt, p.End_dt" & _
        " FROM op_system.db_position p" & _
        " WHERE p.lifeno = " & SText(lifeNo) & _
        "     AND p.Position IN ('당회장', '당회장대리', '동역');"
    makeListData strSql, "op_system.db_position"
    
    Dim result As String
    result = GetConvertedFormatPeriod
    
    GetTotalPastoralCareer = result

End Function

Private Function GetConvertedFormatPeriod() As String

    Dim year As Integer
    Dim month As Integer

    '--//탐색을 위해 minDate, maxDate 검출
    Dim minDate As Date: minDate = DateSerial(9999, 12, 31)
    Dim maxDate As Date: maxDate = DateSerial(1900, 1, 1)
    
    If cntRecord <= 0 Then
        GetConvertedFormatPeriod = ""
    End If
    
    Dim i As Integer
    For i = 0 To cntRecord - 1
        minDate = WorksheetFunction.Min(minDate, LISTDATA(i, 0))
        maxDate = WorksheetFunction.Max(maxDate, WorksheetFunction.Min(Date, LISTDATA(i, 1)))
    Next
    
    '--//한 달씩 건너뛰며 목회경력에 포함된 날짜라면 개월수 추가
    Dim tempDate As Date: tempDate = WorksheetFunction.EoMonth(minDate, 0)
    Do
        For i = 0 To cntRecord - 1
            Dim startDate As Date: startDate = LISTDATA(i, 0)
            Dim endDate As Date: endDate = LISTDATA(i, 1)
            If startDate <= tempDate And tempDate <= endDate Then
                month = month + 1
                If month >= 12 Then
                    month = 0
                    year = year + 1
                    Exit For
                End If
            End If
        Next
        
        If tempDate = DateSerial(9999, 12, 31) Then
            tempDate = DateSerial(9999, 12, 31)
        Else
            tempDate = WorksheetFunction.EDate(tempDate, 1)
        End If
        
        If tempDate > maxDate Then
            Exit Do
        End If
    Loop

    '--//year, month => Y년 M개월 형식으로 변환
    Dim result As String
    If year > 0 Then
        result = result & year & "년"
    End If
    
    If month > 0 Then
        If result = "" Then
            result = month & "개월"
        Else
            result = result & " " & month & "개월"
        End If
    End If
    
    GetConvertedFormatPeriod = result

End Function

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
        strSql = "SELECT a.`생명번호`,a.`교회명`,a.`한글이름(직분)`,a.`직책`,a.`배우자생번` " & _
                    "FROM " & TB1 & " a " & _
                    "WHERE (a.`한글이름(직분)` LIKE '%" & Me.txtChurchNM & "%' OR a.`교회명` LIKE '%" & Me.txtChurchNM & "%'" & _
                    " OR a.`영문이름` LIKE '%" & Me.txtChurchNM & "%' OR a.`지교회명` LIKE '%" & Me.txtChurchNM & "%'" & _
                    " OR a.`영문교회명` LIKE '%" & Me.txtChurchNM & "%' OR a.`영문지교회명` LIKE '%" & Me.txtChurchNM & "%' " & _
                    " OR a.`생명번호` LIKE '%" & Me.txtChurchNM & "%') " & " AND a.`관리부서` = " & SText(USER_DEPT) & ";"
    Case TB2
    Case TB3 '--//기본정보
        With Me.lstPStaff
            strSql = "SELECT * FROM " & TB3 & " a WHERE a.`생명번호` = " & SText(.List(.listIndex))
        End With
    Case TB4 '--//직분이력
        With Me.lstPStaff
            strSql = "SELECT * FROM (SELECT a.`교회명`,a.`임명일`,a.`직분` FROM " & TB4 & " a WHERE a.`생명번호` = " & SText(.List(.listIndex)) & " ORDER BY a.`임명일` DESC LIMIT 3) a ORDER BY a.`임명일`;"
        End With
    Case TB5 '--//발령이력
        With Me.lstPStaff
            strSql = "SELECT * FROM (SELECT a.`발령일`,a.`직분/직책`,a.`교회구분`,a.`교회명` FROM " & TB5 & " a WHERE a.`생명번호` = " & SText(.List(.listIndex)) & " ORDER BY a.`발령일` DESC LIMIT 10) a ORDER BY a.`발령일`, FIELD(a.`교회구분`,'MC','HBC','BC','PBC');"
        End With
    Case TB6 '--//출입국이력
        With Me.lstPStaff
            strSql = "SELECT * FROM (SELECT a.`방문일자`,a.`방문목적` FROM " & TB6 & " a WHERE a.`생명번호` = " & SText(.List(.listIndex)) & " ORDER BY a.`방문일자` DESC LIMIT 5) a ORDER BY a.`방문일자`;"
        End With
    Case TB7 '--//복음성과
        With Me.lstPStaff
            strSql = "SELECT a.`교회명` ,a.`날짜`,a.`전체1회`,a.`전체4회`,a.`학생1회`,a.`학생4회`,a.`반차`,a.`침례`,a.`전도인`,a.`구역장`,a.`지역장`,a.`직분`,a.`직책`,a.`관리시작일`,a.`관리종료일`,a.`교회구분` FROM " & TB7 & " a WHERE a.`생명번호` = " & SText(.List(.listIndex)) & ";"
        End With
    Case TB8 '--//가족정보
        
        If Range("F6") = 0 Then
            
            strSql = "SELECT DISTINCT a.family_cd FROM " & TB8 & " a WHERE a.lifeno = " & SText(Range("C6").Value) & " AND a.relations NOT IN (""부"",""모"");"
            Call makeListData(strSql, TB8)
            
            If cntRecord = 1 Then
                Range("PStaff_Detail_MaleFamilyCode") = LISTDATA(0, 0)
                strSql = "SELECT * FROM (SELECT a.family_id,a.family_cd,REPLACE(a.relations,'(별세)','') AS 'relations',a.lifeno,a.name_ko,a.name_en,a.church_sid,a.church_nm,a.title,a.POSITION,a.birthday,a.education,a.religion,a.recognition,a.memo,a.SUSPEND,CONCAT(a.title,if(a.title<>'' AND a.position<>'','/',''),a.position) AS 'status',RANK() OVER (PARTITION BY a.family_cd,relations ORDER BY a.birthday) AS 'rank' FROM op_system.v_familyinfo a WHERE a.family_cd = " & SText(LISTDATA(0, 0)) & " ORDER BY family_cd,FIELD(relations,'부','부(별세)','모','모(별세)','형제','자매'),birthday) a WHERE a.lifeno <> " & SText(Range("C6").Value) & ";"
            ElseIf cntRecord > 1 Then
                MsgBox "선지자 가족정보 데이터에 중복오류가 있습니다. 중복된 자료를 제거하세요.", vbCritical, banner
            End If
        Else
            strSql = "SELECT DISTINCT a.family_cd FROM " & TB8 & " a WHERE a.lifeno = " & SText(Range("C6").Value) & " AND a.relations NOT IN (""부"",""모"")"
            Call makeListData(strSql, TB8)
            
            If cntRecord = 0 Then
                strSql = "SELECT DISTINCT a.family_cd FROM " & TB8 & " a WHERE a.lifeno = " & SText(Range("F6").Value) & " AND a.relations NOT IN (""부"",""모"");"
                Call makeListData(strSql, TB8)
                
                If cntRecord = 1 Then
                    Range("PStaff_Detail_MaleFamilyCode") = LISTDATA(0, 0)
                    strSql = "SELECT * FROM (SELECT a.family_id,a.family_cd,REPLACE(a.relations,'(별세)','') AS 'relations',a.lifeno,a.name_ko,a.name_en,a.church_sid,a.church_nm,a.title,a.POSITION,a.birthday,a.education,a.religion,a.recognition,a.memo,a.SUSPEND,CONCAT(a.title,if(a.title<>'' AND a.position<>'','/',''),a.position) AS 'status',RANK() OVER (PARTITION BY a.family_cd,relations ORDER BY a.birthday) AS 'rank' FROM op_system.v_familyinfo a WHERE a.family_cd = " & SText(LISTDATA(0, 0)) & " ORDER BY family_cd,FIELD(relations,'부','부(별세)','모','모(별세)','형제','자매'),birthday) a WHERE a.lifeno NOT IN (" & SText(Range("F6").Value) & ");"
                End If
            ElseIf cntRecord = 1 Then
                strSql = "SELECT DISTINCT a.family_cd FROM " & TB8 & " a WHERE a.lifeno = " & SText(Range("F6").Value) & " AND a.relations NOT IN (""부"",""모"");"
                Call makeListData(strSql, TB8)
                
                If cntRecord = 1 Then
                    strSql = "SELECT DISTINCT a.family_cd FROM " & TB8 & " a WHERE a.lifeno = " & SText(Range("C6").Value) & " AND a.relations NOT IN (""부"",""모"")" & _
                                " UNION SELECT DISTINCT a.family_cd FROM " & TB8 & " a WHERE a.lifeno = " & SText(Range("F6").Value) & " AND a.relations NOT IN (""부"",""모"");"
                    Call makeListData(strSql, TB8)
                    
                    Range("PStaff_Detail_MaleFamilyCode") = LISTDATA(0, 0)
                    Range("PStaff_Detail_FemaleFamilyCode") = LISTDATA(1, 0)
                    strSql = "SELECT * FROM (SELECT a.family_id,a.family_cd,REPLACE(a.relations,'(별세)','') AS 'relations',a.lifeno,a.name_ko,a.name_en,a.church_sid,a.church_nm,a.title,a.POSITION,a.birthday,a.education,a.religion,a.recognition,a.memo,a.SUSPEND,CONCAT(a.title,if(a.title<>'' AND a.position<>'','/',''),a.position) AS 'status',RANK() OVER (PARTITION BY a.family_cd,relations ORDER BY a.birthday) AS 'rank' FROM op_system.v_familyinfo a WHERE a.family_cd = " & SText(LISTDATA(0, 0)) & _
                            " UNION SELECT a.family_id,a.family_cd,REPLACE(a.relations,'(별세)','') AS 'relations',a.lifeno,a.name_ko,a.name_en,a.church_sid,a.church_nm,a.title,a.POSITION,a.birthday,a.education,a.religion,a.recognition,a.memo,a.SUSPEND,CONCAT(a.title,if(a.title<>'' AND a.position<>'','/',''),a.position) AS 'status',RANK() OVER (PARTITION BY a.family_cd,relations ORDER BY a.birthday) AS 'rank' FROM op_system.v_familyinfo a WHERE a.family_cd = " & SText(LISTDATA(1, 0)) & " ORDER BY family_cd,FIELD(relations,'부','부(별세)','모','모(별세)','형제','자매'),birthday) a WHERE a.lifeno NOT IN (" & SText(Range("C6").Value) & "," & SText(Range("F6").Value) & ");"
                ElseIf cntRecord = 0 Then
                    Range("PStaff_Detail_MaleFamilyCode") = LISTDATA(0, 0)
                    strSql = "SELECT * FROM (SELECT a.family_id,a.family_cd,REPLACE(a.relations,'(별세)','') AS 'relations',a.lifeno,a.name_ko,a.name_en,a.church_sid,a.church_nm,a.title,a.POSITION,a.birthday,a.education,a.religion,a.recognition,a.memo,a.SUSPEND,CONCAT(a.title,if(a.title<>'' AND a.position<>'','/',''),a.position) AS 'status',RANK() OVER (PARTITION BY a.family_cd,relations ORDER BY a.birthday) AS 'rank' FROM op_system.v_familyinfo a WHERE a.family_cd = " & SText(LISTDATA(0, 0)) & " ORDER BY family_cd,FIELD(relations,'부','부(별세)','모','모(별세)','형제','자매'),birthday) a WHERE a.lifeno NOT IN (" & SText(Range("C6").Value) & ");"
                End If
            ElseIf cntRecord > 2 Then
                MsgBox "선지자 혹은 사모 가족정보 데이터에 중복오류가 있습니다. 중복된 자료를 제거하세요.", vbCritical, banner
            End If
        End If
        
        strSql = strSql & ";"
    Case TB9 '--//복음성과(본교회)
        With Me.lstPStaff
            strSql = " SELECT a.`교회명` ,a.`날짜`,a.`전체1회`,a.`전체4회`,a.`학생1회`,a.`학생4회`,a.`반차`,a.`침례`,a.`전도인`,a.`구역장`,a.`지역장`,a.`직분`,a.`직책`,a.`관리시작일`,a.`관리종료일`,a.`교회구분` FROM " & TB9 & " a WHERE a.`생명번호` = " & SText(.List(.listIndex)) & ";"
        End With
    Case TB10 '--//복음성과(전체+본교회)
        With Me.lstPStaff
            strSql = "SELECT a.`교회명` ,a.`날짜`,a.`전체1회`,a.`전체4회`,a.`학생1회`,a.`학생4회`,a.`반차`,a.`침례`,a.`전도인`,a.`구역장`,a.`지역장`,a.`직분`,a.`직책`,a.`관리시작일`,a.`관리종료일`,a.`교회구분` FROM " & TB7 & " a WHERE a.`생명번호` = " & SText(.List(.listIndex)) & " UNION " & _
                    " SELECT a.`교회명` ,a.`날짜`,a.`전체1회`,a.`전체4회`,a.`학생1회`,a.`학생4회`,a.`반차`,a.`침례`,a.`전도인`,a.`구역장`,a.`지역장`,a.`직분`,a.`직책`,a.`관리시작일`,a.`관리종료일`,a.`교회구분` FROM " & TB9 & " a WHERE a.`생명번호` = " & SText(.List(.listIndex)) & " ORDER BY `관리종료일`,`교회명`,`날짜`;"
        End With
    Case TB11 '--//발령이력
        With Me.lstPStaff
            strSql = "SELECT * FROM (SELECT a.`발령일`,a.`직분/직책`,a.`교회구분`,a.`교회명` FROM " & TB11 & " a WHERE a.`생명번호` = " & SText(.List(.listIndex)) & " ORDER BY a.`발령일` DESC LIMIT 10) a ORDER BY a.`발령일`, FIELD(a.`교회구분`,'MC','HBC','BC','PBC');"
        End With
    Case TB12 '--//역대 교회목록
        With Me.lstPStaff
            strSql = "SELECT * FROM " & TB12 & " a WHERE a.`생명번호` = " & SText(.List(.listIndex)) & ";"
        End With
    Case TB13
        With Me.lstPStaff
            strSql = "SELECT * FROM " & TB13 & " a WHERE a.`생명번호` = " & SText(.List(.listIndex)) & ";"
        End With
    Case TB14
        With Me.lstPStaff
            strSql = "SELECT * FROM " & TB14 & " a WHERE a.`생명번호` = " & SText(.List(.listIndex)) & ";"
        End With
    Case Else
        '//화폐id, 화폐약칭, 조회일, 원화환율, 달러화환율
        'strSQL = "SELECT a.currency_id, a.currency_un, a.refer_dt, a.fx_rate_krw, a.fx_rate_usd " & _
                      "FROM " & TB1 & " a WHERE a.user_id = " & user_id & ";"
    End Select
    makeSelectSQL = strSql
End Function
'-----------------------------------------
'  조건별 Select SQL문 작성
'    - makeSelectSQL(검색어, 필터)
'    - DB에서 반환할 리스트 필드수정
'-----------------------------------------
Private Function makeSelectSQL2(ByVal tableNM As String) As String
    Select Case tableNM
    Case TB4 '--//사모직분 이력
        With Me.lstPStaff
            strSql = "SELECT * FROM (SELECT a.`교회명`,a.`임명일`,a.`직분` FROM " & TB4 & " a WHERE a.`생명번호` = " & SText(.List(.listIndex, 4)) & " ORDER BY a.`임명일` DESC LIMIT 3) a ORDER BY a.`임명일`;"
        End With
    Case TB6 '--//사모 출입국 이력
        With Me.lstPStaff
            strSql = "SELECT * FROM (SELECT a.`방문일자`,a.`방문목적` FROM " & TB6 & " a WHERE a.`생명번호` = " & SText(.List(.listIndex, 4)) & " ORDER BY a.`방문일자` DESC LIMIT 5) a ORDER BY a.`방문일자`;"
        End With
    Case Else
        '//화폐id, 화폐약칭, 조회일, 원화환율, 달러화환율
        'strSQL = "SELECT a.currency_id, a.currency_un, a.refer_dt, a.fx_rate_krw, a.fx_rate_usd " & _
                      "FROM " & TB1 & " a WHERE a.user_id = " & user_id & ";"
    End Select
    makeSelectSQL2 = strSql
End Function

Private Sub sbClearVariant()
    Erase LISTFIELD
    Erase LISTDATA
    cntRecord = Empty
    strSql = vbNullString
End Sub

Private Sub sbArrangeChart_Atten()

    Dim noMax As Integer
    Dim noMin As Integer
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim Term As Long
    Dim rngTarget_above1_All As Range
    Dim rngTarget_above1_Stu As Range
    Dim rngTarget_4Rate As Range
    Dim rngTarget_TitheRate As Range
    Dim ChartNM As String
    Dim RangeNM As String
    Dim rngChartNM As String
    Dim cycle As Long
    Dim Rate_Above4() As Variant
    
    
On Error Resume Next
    
    For cycle = 1 To 8
        
        '--//차트이름 설정
        ChartNM = "Chart_Church" & cycle
        
        '--//범위이름 설정
        RangeNM = "Result_rngChurch" & cycle
        
        '--//범위 설정
        Set rngTarget_above1_All = Range(RangeNM & "_1All")
        Set rngTarget_above1_Stu = Range(RangeNM & "_1Stu")
        Set rngTarget_4Rate = Range(RangeNM & "_4Rate")
        Set rngTarget_TitheRate = Range(RangeNM & "_TitheRate")
        
        '--//차트의 최대값, 최소값 범위설정
        noMax = WorksheetFunction.Max(rngTarget_above1_All) '--//학생이상 1회출석 최대값
        noMin = WorksheetFunction.Min(rngTarget_above1_Stu) '--//학생이상 4회출석 최소값
        i = 1: j = 1
          
          
          '--//출석 그래프 차트에서
          With Sheets("선지자 상세정보").ChartObjects(ChartNM).Chart.Axes(xlValue)
            
            '--//식구 규모에 따라 스케일을 달리 합니다.
            Select Case noMax
                Case Is <= 100: Term = 10
                Case Is <= 500: Term = 50
                Case Is <= 1000: Term = 100
                Case Else: Term = 100
            End Select
            
            '--//범위의 최대값을 구합니다..
            Do
                If Term * i > noMax Then
                    .MaximumScale = Term * i
                    Exit Do
                End If
                i = i + 1
            Loop
            
            '--//범위의 최소값을 구합니다.
            Do
                If Term * j >= noMin * 0.3 Then
                    .MinimumScale = Term * (j - 1)
                    Exit Do
                End If
                j = j + 1
            Loop
            
            '--//범위의 최대값과 최소값의 차이가 4의 배수가 아니면 최대값 수정
            Do
                If (.MaximumScale - .MinimumScale) Mod 4 = 0 Then Exit Do
                i = i + 1
                .MaximumScale = Term * i
            Loop
            
            .MajorUnit = (.MaximumScale - .MinimumScale) / 4
            
          End With
          
          
          '--//4회비율 배열로 저장
'          ReDim Rate_Above4(0 To Range(RangeNM).Columns.Count - 1)
'          For k = 0 To UBound(Rate_Above4)
'            Rate_Above4(k) = rngTarget_above4_Stu.Cells(k) / rngTarget_above1_Stu.Cells(k)
'          Next
          
          '--//보조축 범위설정
          With Sheets("선지자 상세정보").ChartObjects(ChartNM).Chart.Axes(xlValue, xlSecondary)
'            .MaximumScale = Application.Max(WorksheetFunction.RoundUp(WorksheetFunction.Max(rngTarget_4Rate), 1), WorksheetFunction.RoundUp(WorksheetFunction.Max(rngTarget_TitheRate), 1))
'            .MinimumScale = Application.Min(WorksheetFunction.RoundDown(WorksheetFunction.Min(rngTarget_4Rate), 1), WorksheetFunction.RoundDown(WorksheetFunction.Min(rngTarget_TitheRate), 1))
            .MaximumScale = 3
            .MinimumScale = 0
          End With
          
          '--//차트 위치조정
'          rngChartNM = "PStaff_Detail_Church" & cycle
'          Sheets("선지자 상세정보").ChartObjects(ChartNM).Top = Range(rngChartNM).Offset(3).Top + 4
      
      Next
On Error GoTo 0

End Sub

