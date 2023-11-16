VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Search_Appointment 
   Caption         =   "�߷ɴ���� �˻�"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7050
   OleObjectBlob   =   "frm_Search_Appointment.frx":0000
   StartUpPosition =   1  '������ ���
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
        MsgBox "���Ա�ȸ�� ������ �ּ���.", vbCritical, banner
        Exit Sub
    End If
    
'---------------------------------------------- ���� ��ü ���� ----------------------------------------------
    
    Me.lblStatus.Caption = "�ʿ��� �ڷ� ��ȸ ��..."
    Me.lblStatus.Visible = True
    Me.Repaint
    
    Dim strLifeNo As String
    With Me.lstPStaff
        strLifeNo = .List(.listIndex)
    End With
    Set objPStaffView = objPStaffViewDao.FindByLifeNo(strLifeNo)
    
    '--//���米ȸ �߷��̷�
    Set objTransfer = objTransferDao.GetLastTransfer(strLifeNo)
    
    '--//���ⱳȸ, ���Ա�ȸ
    Set objChurchFrom = objChurchDao.FindByChurchCode(objPStaffView.churchCode)
    Set objChurchTo = objChurchDao.FindByChurchCode(Me.txtTo_sid)
    
    '--//���ⱳȸ ��������, ���Ա�ȸ ��������
    Set objGeoDataFrom = objGeoDataDao.FindGeoDataById(objChurchFrom.GeoCode)
    Set objGeoDataTo = objGeoDataDao.FindGeoDataById(objChurchTo.GeoCode)
    
    '--//���ⱳȸ ���۽����� ������� MC, MM, BC, PBC ����
    Dim collectCntMcBcPbcStart As Collection
    Dim collectCntMcBcPbcNow As Collection
    Dim curStartDate As Date
    curStartDate = WorksheetFunction.EoMonth(objPStaffView.AppoCur, 0) + 1 '�߷ɹ޾� �� ������ �⼮�� ����ο�
    Set collectCntMcBcPbcStart = objChurchDao.GetMcBcPbcCount(objChurchFrom, curStartDate)
    Set collectCntMcBcPbcNow = objChurchDao.GetMcBcPbcCount(objChurchFrom)
    
    '--//���ⱳȸ ���۽����� ������� MC, MM, BC, PBC �⼮���
    Dim collectAttenSummaryStart As Collection
    Dim collectAttenSummaryNow As Collection
    Set collectAttenSummaryStart = objAttendanceDao.GetMcBcPbcAttenSummary(objChurchFrom, curStartDate)
    Set collectAttenSummaryNow = objAttendanceDao.GetMcBcPbcAttenSummary(objChurchFrom)
    
    '--//���Ա�ȸ ���۽����� ������� MC, MM, BC, PBC ����
    Dim collectCntMcBcPbcNow2 As Collection
    Set collectCntMcBcPbcNow2 = objChurchDao.GetMcBcPbcCount(objChurchTo)
    
    '--//���Ա�ȸ ���۽����� ������� MC, MM, BC, PBC �⼮���
    Dim collectAttenSummaryNow2 As Collection
    Set collectAttenSummaryNow2 = objAttendanceDao.GetMcBcPbcAttenSummary(objChurchTo)
    
    '--//���ⱳȸ ���� ����, ���� ��Ȳ
    '--//����, ���� ���� Attendance ��ü�� �ľ��� �� ���� ���� ���
    Dim objPStaffInfo As New PStaffInfoView
    Dim objPStaffInfoDao As New PStaffInfoViewDao
    Dim tmpChurchList As Object
    Dim tmpChurch As New Church
    Dim intAssistantMC As Integer, intTheologicalStuMC As Integer
    Dim intAssistantBC As Integer, intTheologicalStuBC As Integer
    Dim intAssistantPBC As Integer, intTheologicalStuPBC As Integer
    
    Set tmpChurchList = objChurchDao.GetMcBcPbcList(objChurchFrom) '--//Ư�� ��ȸ�� ����ȸ, ����ȸ, ����� ����Ʈ ����
    Dim collectCountPastoralStaff As Collection
    Dim tmpPStaffInfoList As Object
    If tmpChurchList.Count > 0 Then
        For Each tmpChurch In tmpChurchList '--//�� ��ȸ���� ��ȯ�ϸ�
            If tmpChurch.Gb <> "MM" Then
                Set tmpPStaffInfoList = objPStaffInfoDao.FindByChurchName(tmpChurch.Name, True) '--//�ű� �Ҽӵ� �ο������ ������ ��
                Dim tmpPStaffInfo As New PStaffInfoView
                For Each tmpPStaffInfo In tmpPStaffInfoList
                    Dim tmpChurchGb As String
                    Select Case tmpPStaffInfo.position
                        Case "����" '--//�����̸� ��ȸ���п� ���� �ο��� ���
                            tmpChurchGb = objChurchDao.FindByChurchName(tmpPStaffInfo.BranchNameKo).Gb
                            If tmpChurchGb = "MC" Then
                                intAssistantMC = intAssistantMC + 1 '--//����ȸ ������
                            ElseIf tmpChurchGb = "BC" Then
                                intAssistantBC = intAssistantBC + 1 '--//����ȸ ������
                            ElseIf tmpChurchGb = "PBC" Then
                                intAssistantPBC = intAssistantPBC + 1 '--//����� ������
                            End If
                        Case "�������1�ܰ�", "�������2�ܰ�", "�������3�ܰ�" '--//������ ��ȸ���п� ���� �ο��� ���
                            tmpChurchGb = objChurchDao.FindByChurchName(tmpPStaffInfo.BranchNameKo).Gb
                            If tmpChurchGb = "MC" Then
                                intTheologicalStuMC = intTheologicalStuMC + 1 '--//����ȸ ������
                            ElseIf tmpChurchGb = "BC" Then
                                intTheologicalStuBC = intTheologicalStuBC + 1 '--//����ȸ ������
                            ElseIf tmpChurchGb = "PBC" Then
                                intTheologicalStuPBC = intTheologicalStuPBC + 1 '--//����� ���� ��
                            End If
                    End Select
                Next
            End If
        Next
    End If
    
    '--//�����̷� �ֱ� 4������ ����
    Dim objPastoralCareerDao As New PastoralCareerDao
    Dim objPastoralCareerList As Object
    Set objPastoralCareerList = objPastoralCareerDao.GetPastoralCareers(objPStaffView.lifeNo, 4)
    
    '--//WorkerEmission ��ȸ�а�, ��������, ��������, ������ ����, ������ ���� �ο����
    Dim objWorkerEmit As WorkerEmission
    Dim objWorkerEmitDao As New WorkerEmissionDao
    Dim collectWorkerEmit As New Collection
    
    Dim tmpPastoralCareer As PastoralCareer
    For Each tmpPastoralCareer In objPastoralCareerList
        '--//�ϲ۹��� �ο� �������
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
        
        '--//��ȸ�а� ���� �������
        Dim collectChurchBranchedOut As New Collection
        'key = tmpChurchCode & tmpStartDate
        collectChurchBranchedOut.Add objChurchDao.CountBranchedOut(tmpChurchCode, tmpStartDate, tmpEndDate), tmpChurchCode & tmpStartDate
    Next
    
    
'------------------------------------------------- ���� �ʱ�ȭ -------------------------------------------------
    

    InitializeReportPage
    
    
'---------------------------------------------- ���Ŀ� �� �����ϱ� ----------------------------------------------
    
ActiveSheet.Unprotect globalSheetPW
    
    Me.lblStatus.Caption = "�߷ɴ���� �⺻���� ���� ��..."
    Me.Repaint
    
    '--//��������
    objGeoDataFrom.InsertToRange Range("A3_Appointment_GeoFrom_RawData")
    objGeoDataTo.InsertToRange Range("A3_Appointment_GeoTo_RawData")
    
    '--//������ �⺻����
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
    
    
    '--//��� �⺻����
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
    '--//�� ��ȸ ��Ȳ
    Range("A3_Appointment_ChurchFrom_Status_Summary").FormulaR1C1 = _
        "=""�� " & objChurchFrom.Name & """&CHAR(10)&CHAR(10)&""" & _
        "�� �α�: ""&TEXT(OFFSET(A3_Appointment_GeoFrom_RawData,,A3_Appointment_Geo_Select*3+2),""#,##0"")" & "&CHAR(10)&CHAR(10)&""" & _
        "1. ����ȸ: " & collectAttenSummaryNow.Item("MM").OnceAll & "/" & collectAttenSummaryNow.Item("MM").OnceStu & """&CHAR(10)&""" & _
        "2. ����ȸ: " & collectAttenSummaryNow.Item("BC").OnceAll & "/" & collectAttenSummaryNow.Item("BC").OnceStu & """&CHAR(10)&""" & _
        "3. �����: " & collectAttenSummaryNow.Item("PBC").OnceAll & "/" & collectAttenSummaryNow.Item("PBC").OnceStu & """&CHAR(10)&""" & _
        "4. ��ü: " & tmpAtten.OnceAll & "/" & tmpAtten.OnceStu & """&CHAR(10)&CHAR(10)&""" & _
        "�� ����(�л���)" & """&CHAR(10)&""" & tmpAtten.TitheStu & "��" & "(" & Format(tmpTitheRate, "#,##0%") & ")" & """&CHAR(10)&CHAR(10)&""" & _
        "�� ������: " & tmpAtten.Evangelist & "��"""
    
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
        "=""�� " & objChurchTo.Name & """&CHAR(10)&CHAR(10)&""" & _
        "�� �α�: ""&TEXT(OFFSET(A3_Appointment_GeoTo_RawData,,A3_Appointment_Geo_Select*3+2),""#,##0"")" & "&CHAR(10)&CHAR(10)&""" & _
        "1. ����ȸ: " & collectAttenSummaryNow2.Item("MM").OnceAll & "/" & collectAttenSummaryNow2.Item("MM").OnceStu & """&CHAR(10)&""" & _
        "2. ����ȸ: " & collectAttenSummaryNow2.Item("BC").OnceAll & "/" & collectAttenSummaryNow2.Item("BC").OnceStu & """&CHAR(10)&""" & _
        "3. �����: " & collectAttenSummaryNow2.Item("PBC").OnceAll & "/" & collectAttenSummaryNow2.Item("PBC").OnceStu & """&CHAR(10)&""" & _
        "4. ��ü: " & tmpAtten2.OnceAll & "/" & tmpAtten2.OnceStu & """&CHAR(10)&CHAR(10)&""" & _
        "�� ����(�л���)" & """&CHAR(10)&""" & tmpAtten2.TitheStu & "��" & "(" & Format(tmpTitheRate, "#,##0%") & ")" & """&CHAR(10)&CHAR(10)&""" & _
        "�� ������: " & tmpAtten2.Evangelist & "��"""
    
    
    Me.lblStatus.Caption = "���米ȸ ���� ���� ��..."
    Me.Repaint
    
    
    '--//�������� �� ��������
    Range("A3_Appointment_ChurchFrom_StartDate") = objPStaffView.AppoCur
    Range("A3_Appointment_ChurchFrom_EndDate") = IIf(objTransfer.endDate >= Date, "����", objTransfer.endDate)
    
    '--//�����ο�
    Set objAttendance = objAttendanceDao.FindByChurchIdAndDate(objChurchFrom.Id, curStartDate)
    Range("A3_Appointment_ChurchFrom_AttenStart") = "'" & objAttendance.OnceAll & "/" & objAttendance.OnceStu
    
    '--//�����ο�
    Dim tmpAttenMaxDate As Date
    tmpAttenMaxDate = objAttendanceDao.GetMaxDate(objAttendanceDao.GetAllAttenByChurchId(objChurchFrom.Id))
    Set objAttendance = objAttendanceDao.FindByChurchIdAndDate(objChurchFrom.Id, tmpAttenMaxDate)
    Range("A3_Appointment_ChurchFrom_AttenEnd") = "'" & objAttendance.OnceAll & "/" & objAttendance.OnceStu
    
    '--//������Ȳ
    '-------MC,BC,PBC ����-------
    Range("A3_Appointment_ChurchFrom_Atten_Summary_Standard").Offset(1) = 1
    Range("A3_Appointment_ChurchFrom_Atten_Summary_Standard").Offset(2) = "'" & _
        collectCntMcBcPbcStart.Item("BC") & "/" & collectCntMcBcPbcNow.Item("BC")
    Range("A3_Appointment_ChurchFrom_Atten_Summary_Standard").Offset(3) = "'" & _
        collectCntMcBcPbcStart.Item("PBC") & "/" & collectCntMcBcPbcNow.Item("PBC")
    
    '-------����ο�--------
    Range("A3_Appointment_ChurchFrom_Atten_Summary_Standard").Offset(1, 1) = _
        "'" & collectAttenSummaryStart.Item("MM").OnceAll & "/" & collectAttenSummaryStart.Item("MM").OnceStu
    Range("A3_Appointment_ChurchFrom_Atten_Summary_Standard").Offset(2, 1) = _
        "'" & collectAttenSummaryStart.Item("BC").OnceAll & "/" & collectAttenSummaryStart.Item("BC").OnceStu
    Range("A3_Appointment_ChurchFrom_Atten_Summary_Standard").Offset(3, 1) = _
        "'" & collectAttenSummaryStart.Item("PBC").OnceAll & "/" & collectAttenSummaryStart.Item("PBC").OnceStu
    
    '-------�����ο�--------
    Range("A3_Appointment_ChurchFrom_Atten_Summary_Standard").Offset(1, 2) = _
        "'" & collectAttenSummaryNow.Item("MM").OnceAll & "/" & collectAttenSummaryNow.Item("MM").OnceStu
    Range("A3_Appointment_ChurchFrom_Atten_Summary_Standard").Offset(2, 2) = _
        "'" & collectAttenSummaryNow.Item("BC").OnceAll & "/" & collectAttenSummaryNow.Item("BC").OnceStu
    Range("A3_Appointment_ChurchFrom_Atten_Summary_Standard").Offset(3, 2) = _
        "'" & collectAttenSummaryNow.Item("PBC").OnceAll & "/" & collectAttenSummaryNow.Item("PBC").OnceStu
    
    '-------������--------
    Range("A3_Appointment_ChurchFrom_Atten_Summary_Standard").Offset(1, 3) = intAssistantMC
    Range("A3_Appointment_ChurchFrom_Atten_Summary_Standard").Offset(2, 3) = intAssistantBC
    Range("A3_Appointment_ChurchFrom_Atten_Summary_Standard").Offset(3, 3) = intAssistantPBC
    
    '-------���������--------
    Range("A3_Appointment_ChurchFrom_Atten_Summary_Standard").Offset(1, 4) = intTheologicalStuMC
    Range("A3_Appointment_ChurchFrom_Atten_Summary_Standard").Offset(2, 4) = intTheologicalStuBC
    Range("A3_Appointment_ChurchFrom_Atten_Summary_Standard").Offset(3, 4) = intTheologicalStuPBC
    
    '-------������--------
    Range("A3_Appointment_ChurchFrom_Atten_Summary_Standard").Offset(1, 5) = collectAttenSummaryNow.Item("MM").GroupLeader
    Range("A3_Appointment_ChurchFrom_Atten_Summary_Standard").Offset(2, 5) = collectAttenSummaryNow.Item("BC").GroupLeader
    Range("A3_Appointment_ChurchFrom_Atten_Summary_Standard").Offset(3, 5) = collectAttenSummaryNow.Item("PBC").GroupLeader
    
    '-------������--------
    Range("A3_Appointment_ChurchFrom_Atten_Summary_Standard").Offset(1, 6) = collectAttenSummaryNow.Item("MM").UnitLeader
    Range("A3_Appointment_ChurchFrom_Atten_Summary_Standard").Offset(2, 6) = collectAttenSummaryNow.Item("BC").UnitLeader
    Range("A3_Appointment_ChurchFrom_Atten_Summary_Standard").Offset(3, 6) = collectAttenSummaryNow.Item("PBC").UnitLeader
    
    '-------����(�л��̻�)--------
    Range("A3_Appointment_ChurchFrom_Atten_Summary_Standard").Offset(1, 7) = collectAttenSummaryNow.Item("MM").TitheStu
    Range("A3_Appointment_ChurchFrom_Atten_Summary_Standard").Offset(2, 7) = collectAttenSummaryNow.Item("BC").TitheStu
    Range("A3_Appointment_ChurchFrom_Atten_Summary_Standard").Offset(3, 7) = collectAttenSummaryNow.Item("PBC").TitheStu
    
    '-------������--------
    Range("A3_Appointment_ChurchFrom_Atten_Summary_Standard").Offset(1, 8) = collectAttenSummaryNow.Item("MM").Evangelist
    Range("A3_Appointment_ChurchFrom_Atten_Summary_Standard").Offset(2, 8) = collectAttenSummaryNow.Item("BC").Evangelist
    Range("A3_Appointment_ChurchFrom_Atten_Summary_Standard").Offset(3, 8) = collectAttenSummaryNow.Item("PBC").Evangelist
    
    
    Me.lblStatus.Caption = "�����⿩�ڷ� ���� ��..."
    Me.Repaint
    
    
    '--//�����⿩
    Dim rowNum As Integer
    rowNum = 1
    For Each tmpPastoralCareer In objPastoralCareerList
        With tmpPastoralCareer
            
            Range("A3_Appointment_PastoralCareer_Standard").Offset(rowNum, 0) = Replace(.churchName, "[", Chr(10) & "[")
            Range("A3_Appointment_PastoralCareer_Standard").Offset(rowNum, 1) = _
                Format(.startDate, "yy.mm") & "~" & Chr(10) & IIf(.endDate = DateSerial(9999, 12, 31), "����", Format(.endDate, "yy.mm"))
            Range("A3_Appointment_PastoralCareer_Standard").Offset(rowNum, 2) = .title
            Range("A3_Appointment_PastoralCareer_Standard").Offset(rowNum, 3) = .position
            Range("A3_Appointment_PastoralCareer_Standard").Offset(rowNum, 4) = .ChurchClass
            If .ChurchClass = "MC" And .position Like "*����*" Then
                '--// �ش� �Ⱓ ���� ��ȸ�� ��� �޸�� ����
                Dim objOverseers As Object
                Dim objOverseerDao As New OverseerDao
                Set objOverseers = objOverseerDao.GetOverseersBetween(.churchCode, .startDate, .endDate)
                
                Dim memo As String
                Dim objoverseer As Overseer
                memo = "[��ó�� ��ȸ��]" & vbCrLf
                For Each objoverseer In objOverseers
                    memo = memo & objoverseer.startDate & " ~ " & Replace(Format(WorksheetFunction.Min(.endDate, objoverseer.endDate), "yyyy-mm-dd"), "9999-12-31", "����") & " " & objoverseer.nameKo & "(" & Left(objoverseer.title, 1) & ")" & vbCrLf
                Next
                Range("A3_Appointment_PastoralCareer_Standard").Offset(rowNum, 0).AddComment memo
                Range("A3_Appointment_PastoralCareer_Standard").Offset(rowNum, 0).Comment.Shape.TextFrame.AutoSize = True
                memo = vbNullString
            End If
            
            '--//����ȸ�� ��� ��ȸ��, ��븮�� �⼮ �� �ϲ۹��� �Է�
            '--//2023.09.01 ������ ������� ��û���� ����ȸ ������ ����ȸ �⿩������ ǥ�õǵ��� ������
            '--//�ʿ����� ���� ��� ������ ������, �ʿ��� ��� �������� ������ �����ϱ� ����� �����ݿ��ϴ� ������ ������.
'            If Not ((.ChurchClass = "MC" Or .ChurchClass = "HBC") And (Not .Position Like "*��*")) Then
                Set objAttendance = objAttendanceDao.FindByChurchIdAndDate(.churchCode, DateSerial(year(.startDate), month(.startDate) + 1, 1)) '�߷ɹ޾� �� ������ �⼮�� ����ο�
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
            
            '--//�Ⱓ�� ��å ������� ��� �Է��ʿ���
            Range("A3_Appointment_PastoralCareer_Standard").Offset(rowNum, 16) = .Period
            
            rowNum = rowNum + 1
            
        End With
    Next
    
    
    Me.lblStatus.Caption = "3�� �������� ���� ��..."
    Me.Repaint
    
    
    '--//�ֱ� 3�� ��������
    '-------��������-------
    Range("A3_Appointment_Last3Year_Summary_Standard").Offset(1) = year(WorksheetFunction.EDate(Date, -36)) & "~" & year(Date)
    Range("A3_Appointment_Last3Year_Atten_Standard").Offset(1, -1) = year(WorksheetFunction.EDate(Date, -36)) & "~" & year(WorksheetFunction.EDate(Date, -25))
    Range("A3_Appointment_Last3Year_Atten_Standard").Offset(2, -1) = year(WorksheetFunction.EDate(Date, -24)) & "~" & year(WorksheetFunction.EDate(Date, -13))
    Range("A3_Appointment_Last3Year_Atten_Standard").Offset(3, -1) = year(WorksheetFunction.EDate(Date, -12)) & "~" & year(Date)
    
    Dim i As Integer
    For i = 12 To 1 Step -1
        Range("A3_Appointment_Last3Year_Atten_Standard").Offset(, -1 * i + 13) = month(WorksheetFunction.EDate(Date, -i)) & "��"
    Next
    
    With objPStaffView
        Range("A3_Appointment_Last3Year_Summary_Standard").Offset(1, 1) = .nameKo
        Range("A3_Appointment_Last3Year_Summary_Standard").Offset(1, 2) = .title
        Range("A3_Appointment_Last3Year_Summary_Standard").Offset(1, 3) = .position
    End With
    
    '--//�ֱ� 3�� ������ ������� ��ȯ
    Dim tmpAttenList As Object
    Set tmpAttenList = CreateObject("System.Collections.ArrayList")
    Dim accumulAttenList As Object
    Set accumulAttenList = CreateObject("System.Collections.ArrayList")
    Dim keySet As Object
    Set keySet = CreateObject("System.Collections.ArrayList")
    For i = 36 To 1 Step -1
        Dim tmpDate As Date
        tmpDate = WorksheetFunction.EoMonth(Date, -1 * i - 1) + 1 'i ���� �� 1��
        
        '--//�����⿩(�̷�) �ȿ��� �ش� ��¥�� ���� ��ȸ ã�Ƽ�
        Set tmpChurch = Nothing
        Set objPastoralCareerList = objPastoralCareerDao.GetPastoralCareers(objPStaffView.lifeNo)
        For Each tmpPastoralCareer In objPastoralCareerList
            If tmpDate > tmpPastoralCareer.startDate And _
                tmpDate <= tmpPastoralCareer.endDate Then
                '--//ã������ �ش� ��ȸ�� Pick
                Select Case Left(tmpPastoralCareer.churchCode, 2)
                    Case "PBC": '2����
                        If tmpChurch.Gb <> "BC" Then
                            Set tmpChurch = objChurchDao.FindByChurchCode(tmpPastoralCareer.churchCode)
                        End If
                    Case "MC", "HBC": '3����
                        If Not (tmpChurch.Gb = "BC" Or tmpChurch.Gb = "PBC") Then
                            Set tmpChurch = objChurchDao.FindByChurchCode(tmpPastoralCareer.churchCode)
                        End If
                    Case Else '������
                        'BC �Ǵ� ���� 1����
                        Set tmpChurch = objChurchDao.FindByChurchCode(tmpPastoralCareer.churchCode)
                End Select
                
                '--//������ ��� ���� ����ȸ �⼮�� �����ϱ� ����
                If Not tmpPastoralCareer.position Like "*��*" Then
                    tmpChurch.Id = Replace(tmpChurch.Id, "MC", "MM")
                End If
            End If
        Next
        
        '--//�̷¿� ������ �Ҽ� ��ȸ �⼮���� �Է�
        '--//�ڡڹ߷��̷��� �ֱ� ������ �̷��� ���� ���� ����
'        If tmpChurch Is Nothing Then
'            Dim tmpTransfer As New Transfer
'            Set tmpTransfer = objTransferDao.FindByLifeNoAndDate(objPStaffView.LifeNo, tmpDate)
'            Set tmpChurch = objChurchDao.FindByChurchCode(tmpTransfer.ChurchID)
'        End If

        '--//�� ���� �⼮�� ArrayList�� ����
        tmpAttenList.Add objAttendanceDao.FindByChurchIdAndDate(tmpChurch.Id, tmpDate)
        accumulAttenList.Add tmpAttenList.Item(tmpAttenList.Count - 1)
        
        '--//��ǥ��ȸ �̸� �̱�
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
        
        '--//1�� ������ ó���ϱ�
        Dim strRepresentChurchName As String
        If i <> 36 And i Mod 12 = 1 Then
            Dim intOffsetRowValue As Integer
            intOffsetRowValue = ((-1 * (i - 1)) + 36) \ 12
            
            '--//��ǥ��ȸ �̸� �̱�
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
            
            '--//���� �ʱ�ȭ
            Set cntRepresentChurchName = New Collection
            Set keySet = CreateObject("System.Collections.ArrayList")
            
            With Range("A3_Appointment_Last3Year_Atten_Standard")
                '--//��౳ȸ��
                .Offset(intOffsetRowValue) = Right(strRepresentChurchName, Len(strRepresentChurchName) - InStrRev(strRepresentChurchName, " "))
                Dim k As Integer
                Set tmpChurch = objChurchDao.FindByChurchName(strRepresentChurchName)
                .Offset(intOffsetRowValue, 22) = tmpChurch.Id '--//�ش� �⵵ ��ǥ��ȸ �ڵ� ���� ����
                For k = 1 To 12
                    .Offset(intOffsetRowValue, k) = "'" & tmpAttenList.Item(k - 1).OnceAll & "/" & tmpAttenList.Item(k - 1).OnceStu
                    
                    If tmpChurch.Id <> Replace(tmpAttenList.Item(k - 1).ChurchID, "MM", "MC") Then
                        '--//�ش� �⵵ ��ǥ ������ �ƴ� ��ȸ�� ȸ�� �۲�
                        .Offset(intOffsetRowValue, k).Font.color = RGB(191, 191, 191)
                    Else
                        '--//�ش� �⵵ ��ǥ ������ �´� ��ȸ�� ������ �۲�
                        .Offset(intOffsetRowValue, k).Font.color = RGB(0, 0, 0)
                    End If
                Next
            End With
            
            
            
            '--//�а���Ȳ �� �ϲ۹���
            Dim tmpMinDate As Date
            Dim tmpMaxDate As Date
            Dim tmpCntBranchedOut As Integer
            Dim tmpWorkerEmit As WorkerEmission
            Dim tmpAttenPlusMinus As Attendance
            Set tmpAttenPlusMinus = New Attendance
            Set tmpWorkerEmit = New WorkerEmission
            For Each tmpPastoralCareer In objPastoralCareerList '�����⿩ ����Ʈ�� ���鼭
                '--//ù��°�� �ι�°�� ��谡 1���� ���̳�
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
                                IIf(Not tmpPastoralCareer.position Like "*��*", Replace(tmpPastoralCareer.churchCode, "MC", "MM"), tmpPastoralCareer.churchCode), _
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
                                IIf(Not tmpPastoralCareer.position Like "*��*", Replace(tmpPastoralCareer.churchCode, "MC", "MM"), tmpPastoralCareer.churchCode), _
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
                        (Not ((tmpPastoralCareer.ChurchClass = "MC" Or tmpPastoralCareer.ChurchClass = "HBC") And (Not tmpPastoralCareer.position Like "*��*"))) Then
                        '--//�������
                        '--//�ش� �⵵ ù ���̸� �� ���� �� �⼮�� �˻�
                        '--//������ �⼮�� ���� church_sid�� �⼮�̸� ������� �� ���⵵ ������ ������ ���
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
                        tmpAttenMax.Subtract tmpAttenMin '--//�����ο� ���
                        tmpAttenPlusMinus.Sum tmpAttenMax '--//�ش� ���� �������� ���
                        
                        '--//�ش� �⵵ �а���Ȳ ����
                        tmpCntBranchedOut = objChurchDao.CountBranchedOut(tmpPastoralCareer.churchCode, tmpMinDate, tmpMaxDate)
                        
                        '--//�ش� �⵵ �ϲ۹��� ����
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
            '--//���� �� �ڸ��� �� �ֱ�
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
            
            '--//���� ���� ����� ���� ����Ʈ �ʱ�ȭ
            tmpAttenList.Clear
            
        End If
    Next
    
    Dim strGrandRepresenCode As String
    strGrandRepresenCode = Range("A3_Appointment_Last3Year_Grand_Representative").Offset(, 1)
    For Each tmpAtten In accumulAttenList
        If Replace(tmpAtten.ChurchID, "MM", "MC") = strGrandRepresenCode Then
            If Range("A3_Appointment_Last3Year_Summary_Standard").Offset(1, 5) = "" Then
                '--//�������� ����ִ� ��쿡�� ä��(�����̹Ƿ� ��������)
                Range("A3_Appointment_Last3Year_Summary_Standard").Offset(1, 5) = "'" & tmpAtten.OnceAll & "/" & tmpAtten.OnceStu
            End If
            '--//������� ��� �ִٺ��� ������ �������ڰ� ��
            Range("A3_Appointment_Last3Year_Summary_Standard").Offset(1, 6) = "'" & tmpAtten.OnceAll & "/" & tmpAtten.OnceStu
        End If
    Next
    
    '--//��õ�ڿ� ��� ���� �Ұ�
    Range("A3_Appointment_Comprehesive_Opinion") = _
        "�غ������: " & objPastoralCareerDao.GetTotalPastoralCareer(objPStaffView.lifeNo) & Chr(10) & _
        "�ۼ�      ǰ: " & Chr(10) & _
        "�۸����ɷ�: " & Chr(10) & _
        "�۱�ȸ�: " & Chr(10) & _
        "�ۿ�      ��: " & Chr(10) & _
        "�ۻ��ǰ: "
    
    Dim intLength As Integer
    intLength = InStr(Range("A3_Appointment_Comprehesive_Opinion"), "�ۼ�      ǰ") - 1
    With Range("A3_Appointment_Comprehesive_Opinion").Characters(Start:=1, Length:=intLength).Font
        .color = RGB(255, 0, 0)
        .FontStyle = "����"
    End With
    
    Dim intStart As Integer
    intStart = InStr(Range("A3_Appointment_Comprehesive_Opinion"), "�ۻ��ǰ")
    intLength = Len(Range("A3_Appointment_Comprehesive_Opinion"))
    With Range("A3_Appointment_Comprehesive_Opinion").Characters(Start:=intStart, Length:=intLength - intStart).Font
        .color = RGB(255, 0, 0)
        .FontStyle = "����"
    End With
    
    Me.lblStatus.Caption = "�ۼ��Ϸ�"
    Me.Repaint
    
    Normal
    
ActiveSheet.Protect globalSheetPW
    
    MsgBox "�Ϸ�Ǿ����ϴ�.", , banner
    
    Me.lblStatus.Visible = False
    
    ActiveSheet.Range("A1").Select
    
End Sub

Public Sub InitializeReportPage()

    ActiveSheet.Unprotect globalSheetPW

    Me.lblStatus.Caption = "���� �ʱ�ȭ ��..."
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
        .FontStyle = "����"
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
    '--//DB���� ����� �޾ƿɴϴ�.
    Set pStaffInfoList = objPStaffInfoDao.FindBySearchText(Me.txtSearchText, False)
    
    '--//�޾ƿ� ����� ���ٸ�
    If pStaffInfoList.Count = 0 Then
        MsgBox "��ȯ�� DB �����Ͱ� �����ϴ�.", vbInformation, banner
        Exit Sub
    End If
    
    Dim tmpPStaffInfo As New PStaffInfoView
    With Me.lstPStaff
        '--//�޾ƿ� ����� lstPStaff�� �߰� �մϴ�.
        For Each tmpPStaffInfo In pStaffInfoList
            If tmpPStaffInfo.position Like "*��*" Or tmpPStaffInfo.position Like "*��*" Then
                Me.lstPStaff.AddItem tmpPStaffInfo.lifeNo
                .List(.ListCount - 1, 1) = tmpPStaffInfo.ChurchNameKo
                .List(.ListCount - 1, 2) = tmpPStaffInfo.NameKoAndTitle
                .List(.ListCount - 1, 3) = tmpPStaffInfo.position
            End If
        Next
    End With
    Me.lstPStaff.Enabled = True
End Sub

'--//lstPStaff�� ���� ���콺 ��ũ��
Private Sub lstPStaff_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Un'HookListBoxScroll
End Sub

'--//lstPStaff�� ���� ���콺 ��ũ��
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
    
    '--//�����߰�
    InsertPicToLabel Me.lblPic, strLifeNo
    
    '--//��Ʈ�� ����
    Me.cmdSearch_Church.Enabled = True
    
End Sub

Private Sub UserForm_Initialize()
    
    If connIP = Empty Then setGlobalVariant '//�������� �缳��
    If checkLogin = 0 Then f_login.Show '//�α���üũ
    If checkLogin = 0 Then End '--//�α��� ���� �� ���ν��� ����
    
    '--//���ʼ���
    Me.cmdClose.Cancel = True
    Me.cmdSearch_Church.Enabled = False
    Me.txtTo_sid.Enabled = False
    
    '--//����Ʈ�ڽ� ����
    With Me.lstPStaff
        .ColumnCount = 4
        .ColumnHeads = False
        .ColumnWidths = "0,120,70,50" '�����ȣ, ��ȸ��, �ѱ��̸�(����), ��å
        .TextAlign = fmTextAlignLeft
        .Font = "����"
    End With
    
    Me.txtFrom.Enabled = False
    Me.txtTo.Enabled = False
    Me.lblStatus.Visible = False
    
End Sub
