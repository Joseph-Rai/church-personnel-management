VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorkerEmissionDao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function GetEmitAssistant(ByVal argChurchCode As String, ByVal argStartDate As Date, ByVal argEndDate As Date) As Integer

    Dim sql As String
    Dim tRecordSet As T_RECORD_SET
    
    sql = "" & _
        "SELECT * FROM " & TABLE_POSITION & " POSITION1 " & _
        "LEFT JOIN " & TABLE_PASTORALSTAFF & " pstaff " & _
        "    ON pstaff.lifeno=POSITION1.LifeNo " & _
        "LEFT JOIN " & TABLE_TRANSFER & " trans " & _
        "    ON pstaff.lifeno=trans.lifeno AND POSITION1.Start_dt BETWEEN trans.start_dt AND trans.end_dt " & _
        "LEFT JOIN " & TABLE_CHURCH_CUSTOM & " churchlist " & _
        "    ON churchlist.church_sid=trans.church_sid " & _
        "WHERE POSITION1.Position='����' " & _
        "    AND (date_add(POSITION1.Start_dt,INTERVAL -1 MONTH) BETWEEN " & SText(argStartDate) & " AND " & SText(argEndDate) & " AND churchlist.church_nm = " & SText(argChurchCode) & ")"
    
    tRecordSet = makeListData(sql, TABLE_POSITION)
    
    GetEmitAssistant = tRecordSet.CNT_RECORD
    
End Function

Public Function GetEmitTheologicalStu(argChurchCode As String, ByVal argStartDate As Date, ByVal argEndDate As Date) As Integer

    Dim sql As String
    Dim tRecordSet As T_RECORD_SET
    
    sql = "" & _
        "SELECT * FROM " & TABLE_THEOLOGICAL & " theo " & _
        "LEFT JOIN " & TABLE_PASTORALSTAFF & " pstaff " & _
        "    ON pstaff.lifeno=theo.LifeNo " & _
        "LEFT JOIN " & TABLE_CHURCH_CUSTOM & " churchlist " & _
        "    ON churchlist.church_sid=theo.church_sid " & _
        "WHERE churchlist.church_sid = " & SText(argChurchCode) & _
        "    AND theo.Start_dt BETWEEN " & SText(argStartDate) & " AND " & SText(argEndDate) & _
        "    AND theo.`Level` LIKE '%1�ܰ�%'"
    
    tRecordSet = makeListData(sql, TABLE_THEOLOGICAL)
    
    GetEmitTheologicalStu = tRecordSet.CNT_RECORD

End Function


'--//������ ������ argStartDate���� argEndDate���� �ش� ��ȸ�� �⼮�� ��� �����Ͽ�
'--//���� ������ ���� ���� ���� ������ �����մϴ�.
'--//������ ���ڰ� ������ ���� �ݿ����� �ʰ� ������ ���ڸ� ����մϴ�.
Public Function GetEmitGroupLeader(argChurchCode As String, ByVal argStartDate As Date, ByVal argEndDate As Date) As Integer

    Dim objAttenPrev As Attendance
    Dim objAttenNext As Attendance
    Dim objAttenDao As New AttendanceDao
    
    Dim cntEmitGroupLeader As Integer
    
    Dim searchDate As Date
    searchDate = DateSerial(year(argStartDate), month(argStartDate), 1)
    Do While searchDate < argEndDate
        Set objAttenPrev = objAttenDao.FindByChurchIdAndDate(argChurchCode, WorksheetFunction.EDate(searchDate, -1))
        Set objAttenNext = objAttenDao.FindByChurchIdAndDate(argChurchCode, searchDate)
        If objAttenNext.GroupLeader <> 0 And objAttenPrev.GroupLeader <> 0 Then '--//����� Ȥ�� ���� ������ ���� 0�� ��� ������� �ʰ� �н�
            cntEmitGroupLeader = cntEmitGroupLeader + WorksheetFunction.Max(0, objAttenNext.GroupLeader - objAttenPrev.GroupLeader)
        End If
        
        '--//���� �޷� �̵�
        searchDate = WorksheetFunction.EDate(searchDate, 1)
    Loop
    
    GetEmitGroupLeader = cntEmitGroupLeader

End Function

'--//������ ������ argStartDate���� argEndDate���� �ش� ��ȸ�� �⼮�� ��� �����Ͽ�
'--//���� ������ ���� ���� ���� ������ �����մϴ�.
'--//������ ���ڰ� ������ ���� �ݿ����� �ʰ� ������ ���ڸ� ����մϴ�.
Public Function GetEmitUnitLeader(argChurchCode As String, ByVal argStartDate As Date, ByVal argEndDate As Date) As Integer

    Dim objAttenPrev As Attendance
    Dim objAttenNext As Attendance
    Dim objAttenDao As New AttendanceDao
    
    Dim cntEmitUnitLeader As Integer
    
    Dim searchDate As Date
    searchDate = DateSerial(year(argStartDate), month(argStartDate), 1)
    Do While searchDate < argEndDate
        Set objAttenPrev = objAttenDao.FindByChurchIdAndDate(argChurchCode, WorksheetFunction.EDate(searchDate, -1))
        Set objAttenNext = objAttenDao.FindByChurchIdAndDate(argChurchCode, searchDate)
        If objAttenNext.UnitLeader <> 0 And objAttenPrev.UnitLeader <> 0 Then '--//����� Ȥ�� ���� ������ ���� 0�� ��� ������� �ʰ� �н�
            cntEmitUnitLeader = cntEmitUnitLeader + WorksheetFunction.Max(0, objAttenNext.UnitLeader - objAttenPrev.UnitLeader)
        End If
        
        '--//���� �޷� �̵�
        searchDate = WorksheetFunction.EDate(searchDate, 1)
    Loop
    
    GetEmitUnitLeader = cntEmitUnitLeader

End Function
