VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Search_Appointment 
   Caption         =   "발령대상자 검색"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7050
   OleObjectBlob   =   "frm_Search_Appointment.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "frm_Search_Appointment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    
    Dim objPStaffView As New PStaffInfoView
    Dim objPStaffViewDao As New PStaffInfoViewDao
    Dim objPastoralCareer As New PastoralCareer
    Dim objChurchFrom As New Church
    Dim objChurchTo As New Church
    Dim objChurchDao As New ChurchDao
    Dim objAttendance As New Attendance
    Dim objAttendanceDao As New AttendanceDao
    Dim objGeoDataFrom As New GeoData
    Dim objGeoDataTo As New GeoData
    Dim objGeoDataDao As New GeoDataDao
    Dim objTransfer As New Transfer
    Dim objTransferDao As New TransferDao
    
    If Me.txtTo_sid = "" Then
        MsgBox "전입교회를 선택해 주세요.", vbCritical, banner
        Exit Sub
    End If
    
'---------------------------------------------- 관련 객체 생성 ----------------------------------------------
    
    Me.lblStatus.Caption = "필요한 자료 조회 중..."
    Me.lblStatus.Visible = True
    Me.Repaint
    
    Dim strLifeNo As String
    With Me.lstPStaff
        strLifeNo = .List(.listIndex)
    End With
    Set objPStaffView = objPStaffViewDao.FindByLifeNo(strLifeNo)
    
    '--//현재교회 발령이력
    Set objTransfer = objTransferDao.GetLastTransfer(strLifeNo)
    
    '--//전출교회, 전입교회
    Set objChurchFrom = objChurchDao.FindByChurchCode(objPStaffView.churchCode)
    Set objChurchTo = objChurchDao.FindByChurchCode(Me.txtTo_sid)
    
    '--//전출교회 지역정보, 전입교회 지역정보
    Set objGeoDataFrom = objGeoDataDao.FindGeoDataById(objChurchFrom.GeoCode)
    Set objGeoDataTo = objGeoDataDao.FindGeoDataById(objChurchTo.GeoCode)
    
    '--//전출교회 시작시점과 현재시점 MC, MM, BC, PBC 개수
    Dim collectCntMcBcPbcStart As Collection
    Dim collectCntMcBcPbcNow As Collection
    Dim curStartDate As Date
    curStartDate = WorksheetFunction.EoMonth(objPStaffView.AppoCur, 0) + 1 '발령받아 온 다음달 출석이 출발인원
    Set collectCntMcBcPbcStart = objChurchDao.GetMcBcPbcCount(objChurchFrom, curStartDate)
    Set collectCntMcBcPbcNow = objChurchDao.GetMcBcPbcCount(objChurchFrom)
    
    '--//전출교회 시작시점과 현재시점 MC, MM, BC, PBC 출석요약
    Dim collectAttenSummaryStart As Collection
    Dim collectAttenSummaryNow As Collection
    Set collectAttenSummaryStart = objAttendanceDao.GetMcBcPbcAttenSummary(objChurchFrom, curStartDate)
    Set collectAttenSummaryNow = objAttendanceDao.GetMcBcPbcAttenSummary(objChurchFrom)
    
    '--//전입교회 시작시점과 현재시점 MC, MM, BC, PBC 개수
    Dim collectCntMcBcPbcNow2 As Collection
    Set collectCntMcBcPbcNow2 = objChurchDao.GetMcBcPbcCount(objChurchTo)
    
    '--//전입교회 시작시점과 현재시점 MC, MM, BC, PBC 출석요약
    Dim collectAttenSummaryNow2 As Collection
    Set collectAttenSummaryNow2 = objAttendanceDao.GetMcBcPbcAttenSummary(objChurchTo)
    
    '--//전출교회 현재 동역, 생도 현황
    '--//동역, 생도 수는 Attendance 객체로 파악할 수 없어 따로 계산
    Dim objPStaffInfo As New PStaffInfoView
    Dim objPStaffInfoDao As New PStaffInfoViewDao
    Dim tmpChurchList As Object
    Dim tmpChurch As New Church
    Dim intAssistantMC As Integer, intTheologicalStuMC As Integer
    Dim intAssistantBC As Integer, intTheologicalStuBC As Integer
    Dim intAssistantPBC As Integer, intTheologicalStuPBC As Integer
    
    Set tmpChurchList = objChurchDao.GetMcBcPbcList(objChurchFrom) '--//특정 교회의 본교회, 지교회, 예배소 리스트 추출
    Dim collectCountPastoralStaff As Collection
    Dim tmpPStaffInfoList As Object
    If tmpChurchList.Count > 0 Then
        For Each tmpChurch In tmpChurchList '--//각 교회들을 순환하며
            If tmpChurch.Gb <> "MM" Then
                Set tmpPStaffInfoList = objPStaffInfoDao.FindByChurchName(tmpChurch.Name, True) '--//거기 소속된 인원목록을 추출한 후
                Dim tmpPStaffInfo As New PStaffInfoView
                For Each tmpPStaffInfo In tmpPStaffInfoList
                    Dim tmpChurchGb As String
                    Select Case tmpPStaffInfo.position
                        Case "동역" '--//동역이면 교회구분에 따라 인원수 계수
                            tmpChurchGb = objChurchDao.FindByChurchName(tmpPStaffInfo.BranchNameKo).Gb
                            If tmpChurchGb = "MC" Then
                                intAssistantMC = intAssistantMC + 1 '--//본교회 동역수
                            ElseIf tmpChurchGb = "BC" Then
                                intAssistantBC = intAssistantBC + 1 '--//지교회 동역수
                            ElseIf tmpChurchGb = "PBC" Then
                                intAssistantPBC = intAssistantPBC + 1 '--//예배소 동역수
                            End If
                        Case "예비생도1단계", "예비생도2단계", "예비생도3단계" '--//생도면 교회구분에 따라 인원수 계수
                            tmpChurchGb = objChurchDao.FindByChurchName(tmpPStaffInfo.BranchNameKo).Gb
                            If tmpChurchGb = "MC" Then
                                intTheologicalStuMC = intTheologicalStuMC + 1 '--//본교회 생도수
                            ElseIf tmpChurchGb = "BC" Then
                                intTheologicalStuBC = intTheologicalStuBC + 1 '--//지교회 생도수
                            ElseIf tmpChurchGb = "PBC" Then
                                intTheologicalStuPBC = intTheologicalStuPBC + 1 '--//예배소 생도 수
                            End If
                    End Select
                Next
            End If
        Next
    End If
    
    '--//복음이력 최근 4개까지 추출
    Dim objPastoralCareerDao As New PastoralCareerDao
    Dim objPastoralCareerList As Object
    Set objPastoralCareerList = objPastoralCareerDao.GetPastoralCareers(objPStaffView.lifeNo, 4)
    
    '--//WorkerEmission 교회분가, 동역배출, 생도배출, 지역장 배출, 구역장 배출 인원계산
    Dim objWorkerEmit As WorkerEmission
    Dim objWorkerEmitDao As New WorkerEmissionDao
    Dim collectWorkerEmit As New Collection
    
    Dim tmpPastoralCareer As PastoralCareer
    For Each tmpPastoralCareer In objPastoralCareerList
        '--//일꾼배출 인원 산출로직
        Set objWorkerEmit = New WorkerEmission
        
        Dim tmpChurchCode As String
        Dim tmpStartDate As Date
        Dim tmpEndDate As Date
        tmpChurchCode = tmpPastoralCareer.churchCode
        tmpStartDate = tmpPastoralCareer.startDate
        tmpEndDate = WorksheetFunction.Min(tmpPastoralCareer.endDate, WorksheetFunction.EoMonth(Date, -1) + 1)
        objWorkerEmit.EmitAssistant = objWorkerEmitDao.GetEmitAssistant(tmpChurchCode, tmpStartDate, tmpEndDate)
        objWorkerEmit.EmitTheologicalStu = objWorkerEmitDao.GetEmitTheologicalStu(tmpChurchCode, tmpStartDate, tmpEndDate)
        objWorkerEmit.EmitGroupLeader = objWorkerEmitDao.GetEmitGroupLeader(tmpChurchCode, tmpStartDate, tmpEndDate)
        objWorkerEmit.EmitUnitLeader = objWorkerEmitDao.GetEmitUnitLeader(tmpChurchCode, tmpStartDate, tmpEndDate)
        
        collectWorkerEmit.Add objWorkerEmit, tmpChurchCode & tmpStartDate
        
        '--//교회분가 개수 산출로직
        Dim collectChurchBranchedOut As New Collection
        'key = tmpChurchCode & tmpStartDate
        collectChurchBranchedOut.Add objChurchDao.CountBranchedOut(tmpChurchCode, tmpStartDate, tmpEndDate), tmpChurchCode & tmpStartDate
    Next
    
    
'------------------------------------------------- 서식 초기화 -------------------------------------------------
    

    InitializeReportPage
    
    
'---------------------------------------------- 서식에 값 삽입하기 ----------------------------------------------
    
ActiveSheet.Unprotect globalSheetPW
    
    Me.lblStatus.Caption = "발령대상자 기본정보 삽입 중..."
    Me.Repaint
    
    '--//지역정보
    objGeoDataFrom.InsertToRange Range("A3_Appointment_GeoFrom_RawData")
    objGeoDataTo.InsertToRange Range("A3_Appointment_GeoTo_RawData")
    
    '--//선지자 기본정보
    Range("A3_Appointment_Church") = objPStaffView.ChurchNameKo
    Range("A3_Appointment_LifeNo") = objPStaffView.lifeNo
    InsertPStaffPic objPStaffView.lifeNo, Range("A3_Appointment_Pic_M")
    Range("A3_Appointment_Name") = objPStaffView.nameKo
    Range("A3_Appointment_TitlePosition") = objPStaffView.title & "/" & objPStaffView.position
    Range("A3_Appointment_BirthDayAndAge") = objPStaffView.Birthday & Chr(10) & "(" & CalculateOnlyAge(objPStaffView.Birthday) & ")"
    
    Dim objSermon As New SermonScore
    Dim objSermonDao As New SermonScoreDao
    
    Set objSermon = objSermonDao.FindByLifeNo(objPStaffView.lifeNo)
    Range("A3_Appointment_Sermon_Score") = objSermon.AvgScore
    
    
    '--//사모 기본정보
    Range("A3_Appointment_LifeNo_Spouse") = objPStaffView.lifeNoSpouse
    InsertPStaffPic objPStaffView.lifeNoSpouse, Range("A3_Appointment_Pic_F")
    Range("A3_Appointment_Name_Spouse") = objPStaffView.NameKoSpouse
    Range("A3_Appointment_TitlePosition_Spouse") = objPStaffView.TitleSpouse & "/" & objPStaffView.PositionSpouse
    Range("A3_Appointment_BirthdayAndAge_Spouse") = objPStaffView.BirthdaySpouse & Chr(10) & "(" & CalculateOnlyAge(objPStaffView.BirthdaySpouse) & ")"
    
    Set objSermon = objSermonDao.FindByLifeNo(objPStaffView.lifeNoSpouse)
    Range("A3_Appointment_Sermon_Score_Spouse") = objSermon.AvgScore
    
    Dim objPWife As New PastoralWife
    Dim objPWifeDao As New PastoralWifeDao
    Set objPWife = objPWifeDao.FindByLifeNo(objPStaffView.lifeNoSpouse)
    Range("A3_Appointment_Health_Spouse") = objPWife.Health
    
    Dim tmpAtten As New Attendance
    Dim tmpTitheRate As Double
    tmpAtten.Sum collectAttenSummaryNow.Item("MM")
    tmpAtten.Sum collectAttenSummaryNow.Item("BC")
    tmpAtten.Sum collectAttenSummaryNow.Item("PBC")
    If tmpAtten.OnceStu = 0 Then
        tmpTitheRate = 0
    Else
        tmpTitheRate = tmpAtten.TitheStu / tmpAtten.OnceStu
    End If
    '--//현 교회 상황
    Range("A3_Appointment_ChurchFrom_Status_Summary").FormulaR1C1 = _
        "=""★ " & objChurchFrom.Name & """&CHAR(10)&CHAR(10)&""" & _
        "※ 인구: ""&TEXT(OFFSET(A3_Appointment_GeoFrom_RawData,,A3_Appointment_Geo_Select*3+2),""#,##0"")" & "&CHAR(10)&CHAR(10)&""" & _
        "1. 본교회: " & collectAttenSummaryNow.Item("MM").OnceAll & "/" & collectAttenSummaryNow.Item("MM").OnceStu & """&CHAR(10)&""" & _
        "2. 지교회: " & collectAttenSummaryNow.Item("BC").OnceAll & "/" & collectAttenSummaryNow.Item("BC").OnceStu & """&CHAR(10)&""" & _
        "3. 예배소: " & collectAttenSummaryNow.Item("PBC").OnceAll & "/" & collectAttenSummaryNow.Item("PBC").OnceStu & """&CHAR(10)&""" & _
        "4. 전체: " & tmpAtten.OnceAll & "/" & tmpAtten.OnceStu & """&CHAR(10)&CHAR(10)&""" & _
        "○ 반차(학생↑)" & """&CHAR(10)&""" & tmpAtten.TitheStu & "명" & "(" & Format(tmpTitheRate, "#,##0%") & ")" & """&CHAR(10)&CHAR(10)&""" & _
        "○ 전도인: " & tmpAtten.Evangelist & "명"""
    
    Dim tmpAtten2 As New Attendance
    tmpAtten2.Sum collectAttenSummaryNow2.Item("MM")
    tmpAtten2.Sum collectAttenSummaryNow2.Item("BC")
    tmpAtten2.Sum collectAttenSummaryNow2.Item("PBC")
    If tmpAtten2.OnceStu = 0 Then
        tmpTitheRate = 0
    Else
        tmpTitheRate = tmpAtten2.TitheStu / tmpAtten2.OnceStu
    End If
    Range("A3_Appointment_ChurchTo_Status_Summary").FormulaR1C1 = _
        "=""★ " & objChurchTo.Name & """&CHAR(10)&CHAR(10)&""" & _
        "※ 인구: ""&TEXT(OFFSET(A3_Appointment_GeoTo_RawData,,A3_Appointment_Geo_Select*3+2),""#,##0"")" & "&CHAR(10)&CHAR(10)&""" & _
        "1. 본교회: " & collectAttenSummaryNow2.Item("MM").OnceAll & "/" & collectAttenSummaryNow2.Item("MM").OnceStu & """&CHAR(10)&""" & _
        "2. 지교회: " & collectAttenSummaryNow2.Item("BC").OnceAll & "/" & collectAttenSummaryNow2.Item("BC").OnceStu & """&CHAR(10)&""" & _
        "3. 예배소: " & collectAttenSummaryNow2.Item("PBC").OnceAll & "/" & collectAttenSummaryNow2.Item("PBC").OnceStu & """&CHAR(10)&""" & _
        "4. 전체: " & tmpAtten2.OnceAll & "/" & tmpAtten2.OnceStu & """&CHAR(10)&CHAR(10)&""" & _
        "○ 반차(학생↑)" & """&CHAR(10)&""" & tmpAtten2.TitheStu & "명" & "(" & Format(tmpTitheRate, "#,##0%") & ")" & """&CHAR(10)&CHAR(10)&""" & _
        "○ 전도인: " & tmpAtten2.Evangelist & "명"""
    
    
    Me.lblStatus.Caption = "현재교회 정보 삽입 중..."
    Me.Repaint
    
    
    '--//시작일자 및 종료일자
    Range("A3_Appointment_ChurchFrom_StartDate") = objPStaffView.AppoCur
    Range("A3_Appointment_ChurchFrom_EndDate") = IIf(objTransfer.endDate >= Date, "현재", objTransfer.endDate)
    
    '--//시작인원
    Set objAttendance = objAttendanceDao.FindByChurchIdAndDate(objChurchFrom.Id, curStartDate)
    Range("A3_Appointment_ChurchFrom_AttenStart") = "'" & objAttendance.OnceAll & "/" & objAttendance.OnceStu
    
    '--//현재인원
    Dim tmpAttenMaxDate As Date
    tmpAttenMaxDate = objAttendanceDao.GetMaxDate(objAttendanceDao.GetAllAttenByChurchId(objChurchFrom.Id))
    Set objAttendance = objAttendanceDao.FindByChurchIdAndDate(objChurchFrom.Id, tmpAttenMaxDate)
    Range("A3_Appointment_ChurchFrom_AttenEnd") = "'" & objAttendance.OnceAll & "/" & objAttendance.OnceStu
    
    '--//복음현황
    '-------MC,BC,PBC 개수-------
    Range("A3_Appointment_ChurchFrom_Atten_Summary_Standard").Offset(1) = 1
    Range("A3_Appointment_ChurchFrom_Atten_Summary_Standard").Offset(2) = "'" & _
        collectCntMcBcPbcStart.Item("BC") & "/" & collectCntMcBcPbcNow.Item("BC")
    Range("A3_Appointment_ChurchFrom_Atten_Summary_Standard").Offset(3) = "'" & _
        collectCntMcBcPbcStart.Item("PBC") & "/" & collectCntMcBcPbcNow.Item("PBC")
    
    '-------출발인원--------
    Range("A3_Appointment_ChurchFrom_Atten_Summary_Standard").Offset(1, 1) = _
        "'" & collectAttenSummaryStart.Item("MM").OnceAll & "/" & collectAttenSummaryStart.Item("MM").OnceStu
    Range("A3_Appointment_ChurchFrom_Atten_Summary_Standard").Offset(2, 1) = _
        "'" & collectAttenSummaryStart.Item("BC").OnceAll & "/" & collectAttenSummaryStart.Item("BC").OnceStu
    Range("A3_Appointment_ChurchFrom_Atten_Summary_Standard").Offset(3, 1) = _
        "'" & collectAttenSummaryStart.Item("PBC").OnceAll & "/" & collectAttenSummaryStart.Item("PBC").OnceStu
    
    '-------현재인원--------
    Range("A3_Appointment_ChurchFrom_Atten_Summary_Standard").Offset(1, 2) = _
        "'" & collectAttenSummaryNow.Item("MM").OnceAll & "/" & collectAttenSummaryNow.Item("MM").OnceStu
    Range("A3_Appointment_ChurchFrom_Atten_Summary_Standard").Offset(2, 2) = _
        "'" & collectAttenSummaryNow.Item("BC").OnceAll & "/" & collectAttenSummaryNow.Item("BC").OnceStu
    Range("A3_Appointment_ChurchFrom_Atten_Summary_Standard").Offset(3, 2) = _
        "'" & collectAttenSummaryNow.Item("PBC").OnceAll & "/" & collectAttenSummaryNow.Item("PBC").OnceStu
    
    '-------동역수--------
    Range("A3_Appointment_ChurchFrom_Atten_Summary_Standard").Offset(1, 3) = intAssistantMC
    Range("A3_Appointment_ChurchFrom_Atten_Summary_Standard").Offset(2, 3) = intAssistantBC
    Range("A3_Appointment_ChurchFrom_Atten_Summary_Standard").Offset(3, 3) = intAssistantPBC
    
    '-------예비생도수--------
    Range("A3_Appointment_ChurchFrom_Atten_Summary_Standard").Offset(1, 4) = intTheologicalStuMC
    Range("A3_Appointment_ChurchFrom_Atten_Summary_Standard").Offset(2, 4) = intTheologicalStuBC
    Range("A3_Appointment_ChurchFrom_Atten_Summary_Standard").Offset(3, 4) = intTheologicalStuPBC
    
    '-------지역장--------
    Range("A3_Appointment_ChurchFrom_Atten_Summary_Standard").Offset(1, 5) = collectAttenSummaryNow.Item("MM").GroupLeader
    Range("A3_Appointment_ChurchFrom_Atten_Summary_Standard").Offset(2, 5) = collectAttenSummaryNow.Item("BC").GroupLeader
    Range("A3_Appointment_ChurchFrom_Atten_Summary_Standard").Offset(3, 5) = collectAttenSummaryNow.Item("PBC").GroupLeader
    
    '-------구역장--------
    Range("A3_Appointment_ChurchFrom_Atten_Summary_Standard").Offset(1, 6) = collectAttenSummaryNow.Item("MM").UnitLeader
    Range("A3_Appointment_ChurchFrom_Atten_Summary_Standard").Offset(2, 6) = collectAttenSummaryNow.Item("BC").UnitLeader
    Range("A3_Appointment_ChurchFrom_Atten_Summary_Standard").Offset(3, 6) = collectAttenSummaryNow.Item("PBC").UnitLeader
    
    '-------반차(학생이상)--------
    Range("A3_Appointment_ChurchFrom_Atten_Summary_Standard").Offset(1, 7) = collectAttenSummaryNow.Item("MM").TitheStu
    Range("A3_Appointment_ChurchFrom_Atten_Summary_Standard").Offset(2, 7) = collectAttenSummaryNow.Item("BC").TitheStu
    Range("A3_Appointment_ChurchFrom_Atten_Summary_Standard").Offset(3, 7) = collectAttenSummaryNow.Item("PBC").TitheStu
    
    '-------전도인--------
    Range("A3_Appointment_ChurchFrom_Atten_Summary_Standard").Offset(1, 8) = collectAttenSummaryNow.Item("MM").Evangelist
    Range("A3_Appointment_ChurchFrom_Atten_Summary_Standard").Offset(2, 8) = collectAttenSummaryNow.Item("BC").Evangelist
    Range("A3_Appointment_ChurchFrom_Atten_Summary_Standard").Offset(3, 8) = collectAttenSummaryNow.Item("PBC").Evangelist
    
    
    Me.lblStatus.Caption = "복음기여자료 삽입 중..."
    Me.Repaint
    
    
    '--//복음기여
    Dim rowNum As Integer
    rowNum = 1
    For Each tmpPastoralCareer In objPastoralCareerList
        With tmpPastoralCareer
            
            Range("A3_Appointment_PastoralCareer_Standard").Offset(rowNum, 0) = Replace(.churchName, "[", Chr(10) & "[")
            Range("A3_Appointment_PastoralCareer_Standard").Offset(rowNum, 1) = _
                Format(.startDate, "yy.mm") & "~" & Chr(10) & IIf(.endDate = DateSerial(9999, 12, 31), "현재", Format(.endDate, "yy.mm"))
            Range("A3_Appointment_PastoralCareer_Standard").Offset(rowNum, 2) = .title
            Range("A3_Appointment_PastoralCareer_Standard").Offset(rowNum, 3) = .position
            Range("A3_Appointment_PastoralCareer_Standard").Offset(rowNum, 4) = .ChurchClass
            If .ChurchClass = "MC" And .position Like "*동역*" Then
                '--// 해당 기간 역대 당회장 목록 메모로 삽입
                Dim objOverseers As Object
                Dim objOverseerDao As New OverseerDao
                Set objOverseers = objOverseerDao.GetOverseersBetween(.churchCode, .startDate, .endDate)
                
                Dim memo As String
                Dim objoverseer As Overseer
                memo = "[거처간 당회장]" & vbCrLf
                For Each objoverseer In objOverseers
                    memo = memo & objoverseer.startDate & " ~ " & Replace(Format(WorksheetFunction.Min(.endDate, objoverseer.endDate), "yyyy-mm-dd"), "9999-12-31", "현재") & " " & objoverseer.nameKo & "(" & Left(objoverseer.title, 1) & ")" & vbCrLf
                Next
                Range("A3_Appointment_PastoralCareer_Standard").Offset(rowNum, 0).AddComment memo
                Range("A3_Appointment_PastoralCareer_Standard").Offset(rowNum, 0).Comment.Shape.TextFrame.AutoSize = True
                memo = vbNullString
            End If
            
            '--//본교회의 경우 당회장, 당대리만 출석 및 일꾼배출 입력
            '--//2023.09.01 안지혜 전도사님 요청으로 본교회 동역도 본교회 기여내용이 표시되도록 수정함
            '--//필요하지 않을 경우 삭제는 쉽지만, 필요할 경우 수동으로 내용을 삽입하기 어려워 수정반영하는 것으로 결정함.
'            If Not ((.ChurchClass = "MC" Or .ChurchClass = "HBC") And (Not .Position Like "*당*")) Then
                Set objAttendance = objAttendanceDao.FindByChurchIdAndDate(.churchCode, DateSerial(year(.startDate), month(.startDate) + 1, 1)) '발령받아 온 다음달 출석이 출발인원
                Range("A3_Appointment_PastoralCareer_Standard").Offset(rowNum, 5) = "'" & objAttendance.OnceAll & "/" & objAttendance.OnceStu
                If .endDate = DateSerial(9999, 12, 31) Then
                    Set tmpAtten = objAttendanceDao.GetLastAttendance(.churchCode)
                Else
                    Set tmpAtten = objAttendanceDao.FindByChurchIdAndDate(.churchCode, DateSerial(year(.endDate), month(.endDate), 1))
                    Do While tmpAtten.OnceAll = 0 And .startDate < .endDate
                        .endDate = WorksheetFunction.EDate(.endDate, -1)
                        Set tmpAtten = objAttendanceDao.FindByChurchIdAndDate(.churchCode, DateSerial(year(.endDate), month(.endDate), 1))
                    Loop
                End If
                Range("A3_Appointment_PastoralCareer_Standard").Offset(rowNum, 6) = "'" & tmpAtten.OnceAll & "/" & tmpAtten.OnceStu
                
                Dim key As String
                key = .churchCode & .startDate
                Range("A3_Appointment_PastoralCareer_Standard").Offset(rowNum, 9) = collectChurchBranchedOut.Item(key)
                Range("A3_Appointment_PastoralCareer_Standard").Offset(rowNum, 10) = collectWorkerEmit.Item(key).EmitAssistant
                Range("A3_Appointment_PastoralCareer_Standard").Offset(rowNum, 11) = collectWorkerEmit.Item(key).EmitTheologicalStu
                Range("A3_Appointment_PastoralCareer_Standard").Offset(rowNum, 12) = collectWorkerEmit.Item(key).EmitGroupLeader
                Range("A3_Appointment_PastoralCareer_Standard").Offset(rowNum, 13) = collectWorkerEmit.Item(key).EmitUnitLeader
'            End If
            
            '--//기간은 직책 상관없이 모두 입력필요함
            Range("A3_Appointment_PastoralCareer_Standard").Offset(rowNum, 16) = .Period
            
            rowNum = rowNum + 1
            
        End With
    Next
    
    
    Me.lblStatus.Caption = "3년 복음실적 삽입 중..."
    Me.Repaint
    
    
    '--//최근 3년 복음실적
    '-------기초정보-------
    Range("A3_Appointment_Last3Year_Summary_Standard").Offset(1) = year(WorksheetFunction.EDate(Date, -36)) & "~" & year(Date)
    Range("A3_Appointment_Last3Year_Atten_Standard").Offset(1, -1) = year(WorksheetFunction.EDate(Date, -36)) & "~" & year(WorksheetFunction.EDate(Date, -25))
    Range("A3_Appointment_Last3Year_Atten_Standard").Offset(2, -1) = year(WorksheetFunction.EDate(Date, -24)) & "~" & year(WorksheetFunction.EDate(Date, -13))
    Range("A3_Appointment_Last3Year_Atten_Standard").Offset(3, -1) = year(WorksheetFunction.EDate(Date, -12)) & "~" & year(Date)
    
    Dim i As Integer
    For i = 12 To 1 Step -1
        Range("A3_Appointment_Last3Year_Atten_Standard").Offset(, -1 * i + 13) = month(WorksheetFunction.EDate(Date, -i)) & "월"
    Next
    
    With objPStaffView
        Range("A3_Appointment_Last3Year_Summary_Standard").Offset(1, 1) = .nameKo
        Range("A3_Appointment_Last3Year_Summary_Standard").Offset(1, 2) = .title
        Range("A3_Appointment_Last3Year_Summary_Standard").Offset(1, 3) = .position
    End With
    
    '--//최근 3년 전부터 현재까지 순환
    Dim tmpAttenList As Object
    Set tmpAttenList = CreateObject("System.Collections.ArrayList")
    Dim accumulAttenList As Object
    Set accumulAttenList = CreateObject("System.Collections.ArrayList")
    Dim keySet As Object
    Set keySet = CreateObject("System.Collections.ArrayList")
    For i = 36 To 1 Step -1
        Dim tmpDate As Date
        tmpDate = WorksheetFunction.EoMonth(Date, -1 * i - 1) + 1 'i 개월 전 1일
        
        '--//복음기여(이력) 안에서 해당 날짜에 속한 교회 찾아서
        Set tmpChurch = Nothing
        Set objPastoralCareerList = objPastoralCareerDao.GetPastoralCareers(objPStaffView.lifeNo)
        For Each tmpPastoralCareer In objPastoralCareerList
            If tmpDate > tmpPastoralCareer.startDate And _
                tmpDate <= tmpPastoralCareer.endDate Then
                '--//찾아지면 해당 교회를 Pick
                Select Case Left(tmpPastoralCareer.churchCode, 2)
                    Case "PBC": '2순위
                        If tmpChurch.Gb <> "BC" Then
                            Set tmpChurch = objChurchDao.FindByChurchCode(tmpPastoralCareer.churchCode)
                        End If
                    Case "MC", "HBC": '3순위
                        If Not (tmpChurch.Gb = "BC" Or tmpChurch.Gb = "PBC") Then
                            Set tmpChurch = objChurchDao.FindByChurchCode(tmpPastoralCareer.churchCode)
                        End If
                    Case Else '나머지
                        'BC 또는 최초 1순위
                        Set tmpChurch = objChurchDao.FindByChurchCode(tmpPastoralCareer.churchCode)
                End Select
                
                '--//동역일 경우 순수 본교회 출석을 삽입하기 위해
                If Not tmpPastoralCareer.position Like "*당*" Then
                    tmpChurch.Id = Replace(tmpChurch.Id, "MC", "MM")
                End If
            End If
        Next
        
        '--//이력에 없으면 소속 교회 출석으로 입력
        '--//★★발령이력이 있기 때문에 이력이 없을 수가 없음
'        If tmpChurch Is Nothing Then
'            Dim tmpTransfer As New Transfer
'            Set tmpTransfer = objTransferDao.FindByLifeNoAndDate(objPStaffView.LifeNo, tmpDate)
'            Set tmpChurch = objChurchDao.FindByChurchCode(tmpTransfer.ChurchID)
'        End If

        '--//그 달의 출석을 ArrayList에 저장
        tmpAttenList.Add objAttendanceDao.FindByChurchIdAndDate(tmpChurch.Id, tmpDate)
        accumulAttenList.Add tmpAttenList.Item(tmpAttenList.Count - 1)
        
        '--//대표교회 이름 뽑기
        Dim cntRepresentChurchName As New Collection
        If tmpChurch.Name <> "" Then
            If keySet.Contains(tmpChurch.Name) = False Then
                keySet.Add tmpChurch.Name
                cntRepresentChurchName.Add 1, tmpChurch.Name
            Else
                Dim tmpCnt As Integer
                tmpCnt = cntRepresentChurchName.Item(tmpChurch.Name)
                cntRepresentChurchName.Remove tmpChurch.Name
                cntRepresentChurchName.Add tmpCnt + 1, tmpChurch.Name
            End If
        End If
        
        '--//1년 단위로 처리하기
        Dim strRepresentChurchName As String
        If i <> 36 And i Mod 12 = 1 Then
            Dim intOffsetRowValue As Integer
            intOffsetRowValue = ((-1 * (i - 1)) + 36) \ 12
            
            '--//대표교회 이름 뽑기
            Dim j As Integer
            strRepresentChurchName = ""
            For j = 0 To keySet.Count - 1
                If strRepresentChurchName = "" Then
                    strRepresentChurchName = keySet.Item(j)
                End If
                If cntRepresentChurchName.Item(strRepresentChurchName) < cntRepresentChurchName.Item(keySet.Item(j)) Then
                    strRepresentChurchName = keySet.Item(j)
                End If
            Next
            
            '--//변수 초기화
            Set cntRepresentChurchName = New Collection
            Set keySet = CreateObject("System.Collections.ArrayList")
            
            With Range("A3_Appointment_Last3Year_Atten_Standard")
                '--//요약교회명
                .Offset(intOffsetRowValue) = Right(strRepresentChurchName, Len(strRepresentChurchName) - InStrRev(strRepresentChurchName, " "))
                Dim k As Integer
                Set tmpChurch = objChurchDao.FindByChurchName(strRepresentChurchName)
                .Offset(intOffsetRowValue, 22) = tmpChurch.Id '--//해당 년도 대표교회 코드 따로 저장
                For k = 1 To 12
                    .Offset(intOffsetRowValue, k) = "'" & tmpAttenList.Item(k - 1).OnceAll & "/" & tmpAttenList.Item(k - 1).OnceStu
                    
                    If tmpChurch.Id <> Replace(tmpAttenList.Item(k - 1).ChurchID, "MM", "MC") Then
                        '--//해당 년도 대표 실적이 아닌 교회는 회색 글꼴
                        .Offset(intOffsetRowValue, k).Font.color = RGB(191, 191, 191)
                    Else
                        '--//해당 년도 대표 실적이 맞는 교회는 검은색 글꼴
                        .Offset(intOffsetRowValue, k).Font.color = RGB(0, 0, 0)
                    End If
                Next
            End With
            
            
            
            '--//분가현황 및 일꾼배출
            Dim tmpMinDate As Date
            Dim tmpMaxDate As Date
            Dim tmpCntBranchedOut As Integer
            Dim tmpWorkerEmit As WorkerEmission
            Dim tmpAttenPlusMinus As Attendance
            Set tmpAttenPlusMinus = New Attendance
            Set tmpWorkerEmit = New WorkerEmission
            For Each tmpPastoralCareer In objPastoralCareerList '복음기여 리스트를 돌면서
                '--//첫번째와 두번째의 경계가 1개월 차이남
                tmpMinDate = WorksheetFunction.Max(tmpPastoralCareer.startDate, WorksheetFunction.EDate(Date, -12 * (i \ 12 + 1)))
                tmpMaxDate = WorksheetFunction.Min(tmpPastoralCareer.endDate, WorksheetFunction.EDate(Date, -12 * (i \ 12) - 1))
                
                If tmpMinDate < tmpMaxDate Then
                    
                    Dim IsMinContained As Boolean: IsMinContained = False
                    Dim IsMaxContained As Boolean: IsMaxContained = False
                    Dim tmpAttenMin As Attendance
                    Dim tmpAttenMax As Attendance
                    Do While (tmpMinDate < tmpMaxDate)
                        If IsMinContained = False Then
                            Set tmpAttenMin = objAttendanceDao.FindByChurchIdAndDate( _
                                IIf(Not tmpPastoralCareer.position Like "*당*", Replace(tmpPastoralCareer.churchCode, "MC", "MM"), tmpPastoralCareer.churchCode), _
                                WorksheetFunction.EoMonth(tmpMinDate, -1) + 1)
                            
                            If tmpAttenMin.OnceAll <> 0 Then
                                For Each tmpAtten In tmpAttenList
                                    If tmpAtten.IsEqual(tmpAttenMin) Then
                                        IsMinContained = True
                                    End If
                                Next
                                If IsMinContained = True Then
                                    Exit Do
                                End If
                            End If
                            tmpMinDate = WorksheetFunction.EDate(tmpMinDate, 1)
                        End If
                    Loop
                    
                    Do While (tmpMinDate < tmpMaxDate)
                        If IsMaxContained = False Then
                            Set tmpAttenMax = objAttendanceDao.FindByChurchIdAndDate( _
                                IIf(Not tmpPastoralCareer.position Like "*당*", Replace(tmpPastoralCareer.churchCode, "MC", "MM"), tmpPastoralCareer.churchCode), _
                                WorksheetFunction.EoMonth(tmpMaxDate, -1) + 1)
                            
                            If tmpAttenMax.OnceAll <> 0 Then
                                For Each tmpAtten In tmpAttenList
                                    If tmpAtten.IsEqual(tmpAttenMax) Then
                                        IsMaxContained = True
                                    End If
                                Next
                                If IsMaxContained = True Then
                                    Exit Do
                                End If
                            End If
                            tmpMaxDate = WorksheetFunction.EDate(tmpMaxDate, -1)
                        End If
                    Loop
                    
                    If (IsMinContained And IsMaxContained) And _
                        (Not ((tmpPastoralCareer.ChurchClass = "MC" Or tmpPastoralCareer.ChurchClass = "HBC") And (Not tmpPastoralCareer.position Like "*당*"))) Then
                        '--//증감계산
                        '--//해당 년도 첫 월이면 그 직전 월 출석도 검색
                        '--//직전월 출석도 같은 church_sid의 출석이면 증감계산 시 전년도 마지막 월부터 계산
                        If month(tmpAttenMin.AttendanceDate) = month(WorksheetFunction.EDate(Date, -12)) Then
                            Dim tmpAttenPrev As Attendance
                            Set tmpAttenPrev = objAttendanceDao.FindByChurchIdAndDate(tmpAttenMin.ChurchID, WorksheetFunction.EDate(tmpAttenMin.AttendanceDate, -1))
                            For Each tmpAtten In accumulAttenList
                                If tmpAtten.IsEqual(tmpAttenPrev) Then
                                    Set tmpAttenMin = tmpAttenPrev
                                    Exit For
                                End If
                            Next
                        End If
                        tmpAttenMax.Subtract tmpAttenMin '--//증감인원 계산
                        tmpAttenPlusMinus.Sum tmpAttenMax '--//해당 연도 증감누적 계산
                        
                        '--//해당 년도 분가현황 추출
                        tmpCntBranchedOut = objChurchDao.CountBranchedOut(tmpPastoralCareer.churchCode, tmpMinDate, tmpMaxDate)
                        
                        '--//해당 년도 일꾼배출 추출
                        If tmpPastoralCareer.ChurchClass = "MC" Then
                            tmpWorkerEmit.EmitAssistant = tmpWorkerEmit.EmitAssistant + _
                                objWorkerEmitDao.GetEmitAssistant(tmpPastoralCareer.churchCode, tmpMinDate, tmpMaxDate)
                            tmpWorkerEmit.EmitTheologicalStu = tmpWorkerEmit.EmitTheologicalStu + _
                                objWorkerEmitDao.GetEmitTheologicalStu(tmpPastoralCareer.churchCode, tmpMinDate, tmpMaxDate)
                        End If
                        tmpWorkerEmit.EmitGroupLeader = tmpWorkerEmit.EmitGroupLeader + _
                            objWorkerEmitDao.GetEmitGroupLeader(tmpPastoralCareer.churchCode, tmpMinDate, tmpMaxDate)
                        tmpWorkerEmit.EmitUnitLeader = tmpWorkerEmit.EmitUnitLeader + _
                            objWorkerEmitDao.GetEmitUnitLeader(tmpPastoralCareer.churchCode, tmpMinDate, tmpMaxDate)
                    End If
                End If
            Next
            '--//서식 각 자리에 값 넣기
            With Range("A3_Appointment_Last3Year_Atten_Standard")
                .Offset(intOffsetRowValue, 13) = "'" & tmpAttenPlusMinus.OnceAll & "/" & tmpAttenPlusMinus.OnceStu
                .Offset(intOffsetRowValue, 14) = tmpCntBranchedOut
                .Offset(intOffsetRowValue, 15) = _
                    tmpWorkerEmit.EmitAssistant + tmpWorkerEmit.EmitTheologicalStu + tmpWorkerEmit.EmitGroupLeader + tmpWorkerEmit.EmitUnitLeader
            End With
            With Range("A3_Appointment_Last3Year_Summary_Standard")
                .Offset(1, 9) = .Offset(1, 9) + tmpWorkerEmit.EmitAssistant
                .Offset(1, 10) = .Offset(1, 10) + tmpWorkerEmit.EmitTheologicalStu
                .Offset(1, 11) = .Offset(1, 11) + tmpWorkerEmit.EmitGroupLeader
                .Offset(1, 12) = .Offset(1, 12) + tmpWorkerEmit.EmitUnitLeader
            End With
            
            '--//다음 연도 계산을 위해 리스트 초기화
            tmpAttenList.Clear
            
        End If
    Next
    
    Dim strGrandRepresenCode As String
    strGrandRepresenCode = Range("A3_Appointment_Last3Year_Grand_Representative").Offset(, 1)
    For Each tmpAtten In accumulAttenList
        If Replace(tmpAtten.ChurchID, "MM", "MC") = strGrandRepresenCode Then
            If Range("A3_Appointment_Last3Year_Summary_Standard").Offset(1, 5) = "" Then
                '--//시작일이 비어있는 경우에만 채움(최초이므로 시작일임)
                Range("A3_Appointment_Last3Year_Summary_Standard").Offset(1, 5) = "'" & tmpAtten.OnceAll & "/" & tmpAtten.OnceStu
            End If
            '--//순서대로 계속 넣다보면 언젠가 종료일자가 됨
            Range("A3_Appointment_Last3Year_Summary_Standard").Offset(1, 6) = "'" & tmpAtten.OnceAll & "/" & tmpAtten.OnceStu
        End If
    Next
    
    '--//추천자와 사모에 관한 소견
    Range("A3_Appointment_Comprehesive_Opinion") = _
        "※복음경력: " & objPastoralCareerDao.GetTotalPastoralCareer(objPStaffView.lifeNo) & Chr(10) & _
        "○성      품: " & Chr(10) & _
        "○말씀능력: " & Chr(10) & _
        "○교회운영: " & Chr(10) & _
        "○열      정: " & Chr(10) & _
        "○사모성품: "
    
    Dim intLength As Integer
    intLength = InStr(Range("A3_Appointment_Comprehesive_Opinion"), "○성      품") - 1
    With Range("A3_Appointment_Comprehesive_Opinion").Characters(Start:=1, Length:=intLength).Font
        .color = RGB(255, 0, 0)
        .FontStyle = "굵게"
    End With
    
    Dim intStart As Integer
    intStart = InStr(Range("A3_Appointment_Comprehesive_Opinion"), "○사모성품")
    intLength = Len(Range("A3_Appointment_Comprehesive_Opinion"))
    With Range("A3_Appointment_Comprehesive_Opinion").Characters(Start:=intStart, Length:=intLength - intStart).Font
        .color = RGB(255, 0, 0)
        .FontStyle = "굵게"
    End With
    
    Me.lblStatus.Caption = "작성완료"
    Me.Repaint
    
    Normal
    
ActiveSheet.Protect globalSheetPW
    
    MsgBox "완료되었습니다.", , banner
    
    Me.lblStatus.Visible = False
    
    ActiveSheet.Range("A1").Select
    
End Sub

Public Sub InitializeReportPage()

    ActiveSheet.Unprotect globalSheetPW

    Me.lblStatus.Caption = "서식 초기화 중..."
    Me.Repaint
    
    Optimization
    
    ActiveSheet.Pictures.Delete
    Range("A3_Appointment_LifeNo").ClearContents
    Range("A3_Appointment_LifeNo_Spouse").ClearContents
    Range("A3_Appointment_Name").Resize(10).ClearContents
    Range("A3_Appointment_ChurchFrom_Status_Summary").Resize(11, 4).ClearContents
    Range("A3_Appointment_Comprehesive_Opinion").Resize(4, 7).ClearContents
    With Range("A3_Appointment_Comprehesive_Opinion").Characters.Font
        .color = RGB(0, 0, 0)
        .FontStyle = "보통"
    End With
    Range("A3_Appointment_Representative_Title").Resize(3, 2).ClearContents
    Range("A3_Appointment_ChurchFrom_StartDate").Resize(, 2).ClearContents
    Range("A3_Appointment_ChurchFrom_AttenStart").Resize(, 2).ClearContents
    Range("A3_Appointment_ChurchFrom_PlusMinus_Reason").ClearContents
    Range("A3_Appointment_ChurchFrom_StartDate").Offset(2).Resize(2, 8).ClearContents
    Range("A3_Appointment_ChurchFrom_Atten_Summary_Standard").Offset(1).Resize(3, 11).ClearContents
    Range("A3_Appointment_PastoralCareer_Standard").Offset(1).Resize(4, 7).ClearContents
    Range("A3_Appointment_PastoralCareer_Standard").Offset(1, 9).Resize(4, 10).ClearContents
    On Error Resume Next
    Dim i As Long
    For i = 1 To 4
        Range("A3_Appointment_PastoralCareer_Standard").Offset(i).Comment.Delete
    Next
    On Error GoTo 0
    Range("A3_Appointment_Last3Year_Summary_Standard").Offset(1).Resize(, 7).ClearContents
    Range("A3_Appointment_Last3Year_Summary_Standard").Offset(1, 9).Resize(, 7).ClearContents
    Range("A3_Appointment_Last3Year_Atten_Standard").Offset(1, -1).Resize(3, 17).ClearContents
'    Range("A3_Appointment_Last3Year_Atten_Standard").Offset(1, 14).Resize(3, 2).ClearContents
    Range("AD21").Resize(2, 3).ClearContents
    Range("AJ19").Resize(3).ClearContents
    Range("A3_Appointment_GeoFrom_RawData").Resize(, 50).ClearContents
    Range("A3_Appointment_GeoTo_RawData").Resize(, 50).ClearContents
    
    
    ActiveSheet.Protect globalSheetPW

End Sub

Private Sub cmdSearch_Church_Click()
    argShow = 3
    argShow3 = 3
    frm_Update_Appointment_1.Show
End Sub

Private Sub cmdSearch_Click()
    Me.lstPStaff.Clear
    
    Dim objPStaffInfo As New PStaffInfoView
    Dim objPStaffInfoDao As New PStaffInfoViewDao
    Dim pStaffInfoList As Object
    '--//DB에서 목록을 받아옵니다.
    Set pStaffInfoList = objPStaffInfoDao.FindBySearchText(Me.txtSearchText, False)
    
    '--//받아온 목록이 없다면
    If pStaffInfoList.Count = 0 Then
        MsgBox "반환할 DB 데이터가 없습니다.", vbInformation, banner
        Exit Sub
    End If
    
    Dim tmpPStaffInfo As New PStaffInfoView
    With Me.lstPStaff
        '--//받아온 목록을 lstPStaff에 추가 합니다.
        For Each tmpPStaffInfo In pStaffInfoList
            If tmpPStaffInfo.position Like "*당*" Or tmpPStaffInfo.position Like "*동*" Then
                Me.lstPStaff.AddItem tmpPStaffInfo.lifeNo
                .List(.ListCount - 1, 1) = tmpPStaffInfo.ChurchNameKo
                .List(.ListCount - 1, 2) = tmpPStaffInfo.NameKoAndTitle
                .List(.ListCount - 1, 3) = tmpPStaffInfo.position
            End If
        Next
    End With
    Me.lstPStaff.Enabled = True
End Sub

'--//lstPStaff를 위한 마우스 스크롤
Private Sub lstPStaff_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Un'HookListBoxScroll
End Sub

'--//lstPStaff를 위한 마우스 스크롤
Private Sub lstPStaff_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Me.lstPStaff.ListCount > 0 Then
        'HookListBoxScroll Me, Me.lstPStaff
    End If
End Sub

Private Sub lstPStaff_Click()
    
    Dim objTransfer As New Transfer
    Dim objTransferDao As New TransferDao
    
    Dim strLifeNo As String
    With Me.lstPStaff
        Me.txtFrom = .List(.listIndex, 1)
        strLifeNo = .List(.listIndex)
    End With
    
    '--//사진추가
    InsertPicToLabel Me.lblPic, strLifeNo
    
    '--//컨트롤 설정
    Me.cmdSearch_Church.Enabled = True
    
End Sub

Private Sub UserForm_Initialize()
    
    If connIP = Empty Then setGlobalVariant '//전역변수 재설정
    If checkLogin = 0 Then f_login.Show '//로그인체크
    If checkLogin = 0 Then End '--//로그인 실패 시 프로시저 종료
    
    '--//기초설정
    Me.cmdClose.Cancel = True
    Me.cmdSearch_Church.Enabled = False
    Me.txtTo_sid.Enabled = False
    
    '--//리스트박스 설정
    With Me.lstPStaff
        .ColumnCount = 4
        .ColumnHeads = False
        .ColumnWidths = "0,120,70,50" '생명번호, 교회명, 한글이름(직분), 직책
        .TextAlign = fmTextAlignLeft
        .Font = "굴림"
    End With
    
    Me.txtFrom.Enabled = False
    Me.txtTo.Enabled = False
    Me.lblStatus.Visible = False
    
End Sub
