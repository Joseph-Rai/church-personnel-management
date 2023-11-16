VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AttendanceDao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'@param argMC: �ش� ��ȸ�� �⼮��� �� ���� �ֱ� �����͸� �����ɴϴ�.
Public Function GetLastAttendance(ByRef argChurchCode As String) As Attendance

    Dim result As New Attendance

    Dim attenList As Object
    Set attenList = GetAllAttenByChurchId(argChurchCode)
    
    Dim tmpAtten As Attendance
    For Each tmpAtten In attenList
        If result.AttendanceDate < tmpAtten.AttendanceDate Then
            Set result = tmpAtten
        End If
    Next
    
    Set GetLastAttendance = result

End Function

'@param argMC: �⼮ ����Ʈ�� �ְ� �� �ȿ��� MaxDate ���� ���� �մϴ�.
'@param searchDate: �˻��ϰ��� �ϴ� ���� �Է��մϴ�. �Է����� ���� �� ���� �� ���� ������ �⼮ �����͸� �ҷ��ɴϴ�.
Public Function GetMcBcPbcAttenSummary(ByRef argMC As Church, Optional ByRef searchDate As Date = "1900-01-01") As Collection

    Dim objChurchDao As New ChurchDao
    Dim colAttenSummary As New Collection

    '--//������� MC�� ������ MC,MM,BC,PBC ����Ʈ
    Dim lstCurMcBcPbc As Object
    Set lstCurMcBcPbc = objChurchDao.GetMcBcPbcList(argMC, searchDate)
    
    Dim objAttenSummaryMC As New Attendance
    Dim objAttenSummaryMM As New Attendance
    Dim objAttenSummaryBC As New Attendance
    Dim objAttenSummaryPBC As New Attendance
    
    Dim tmpChurch As Church
    For Each tmpChurch In lstCurMcBcPbc
        Dim tmpAtten As New Attendance
        Set tmpAtten = FindByChurchIdAndDate(tmpChurch.Id, searchDate, True)
        Select Case tmpChurch.Gb
        Case "MC", "HBC":
            objAttenSummaryMC.Sum tmpAtten
        Case "MM":
            objAttenSummaryMM.Sum tmpAtten
        Case "BC":
            objAttenSummaryBC.Sum tmpAtten
        Case "PBC":
            objAttenSummaryPBC.Sum tmpAtten
        End Select
    Next
    
    colAttenSummary.Add objAttenSummaryMC, "MC"
    colAttenSummary.Add objAttenSummaryMM, "MM"
    colAttenSummary.Add objAttenSummaryBC, "BC"
    colAttenSummary.Add objAttenSummaryPBC, "PBC"
    
    Set GetMcBcPbcAttenSummary = colAttenSummary

End Function

'@param lstAtten: �⼮ ����Ʈ�� �ְ� �� �ȿ��� MaxDate ���� ���� �մϴ�.
Public Function GetMaxDate(ByRef lstAtten As Object) As Date

    Dim tmpAtten As New Attendance
    Dim maxDate As Date
    maxDate = DateSerial(1900, 1, 1)
    For Each tmpAtten In lstAtten
        If maxDate < tmpAtten.AttendanceDate Then
            maxDate = tmpAtten.AttendanceDate
        End If
    Next
    
    GetMaxDate = maxDate

End Function

'@param argChurchId: ��ȸ�ڵ��� �������� DB���� �˻��մϴ�.
Public Function GetAllAttenByChurchId(ByVal argChurchId As String) As Object
    
    Dim sql As String
    Dim tRecordSet As T_RECORD_SET
    Dim attenList As Object
    Set attenList = CreateObject("System.Collections.ArrayList")
    
    sql = "SELECT * FROM " & TABLE_ATTENDANCE & _
                    " a WHERE a.church_sid = " & SText(argChurchId) & ";"
    
    tRecordSet = makeListData(sql, TABLE_ATTENDANCE)
    
    Dim tmpAtten As New Attendance
    Dim i As Long
    For i = 0 To tRecordSet.CNT_RECORD - 1
        If tRecordSet.CNT_RECORD > 0 Then
            tmpAtten.ParseFromRecordSet tRecordSet, i
        Else
            Set tmpAtten = Nothing
        End If
        attenList.Add tmpAtten
    Next
    
    Set GetAllAttenByChurchId = attenList
    
End Function

'@param argChurchId: ��ȸ�ڵ带 �������� DB���� �˻��մϴ�.
'@param argDate: �⼮��¥�� �������� DB���� �˻��մϴ�.
'@param blnIsRestricted: church_sid_custom ���� �������� �⼮�� �˻��ϴ��� ����
'                        DEFAULT: FALSE-church_sid_custom ���� �������� �˻� / TRUE-church_sid ���� �������� �˻�
'                                (�����̷� ����� ��� ��ȸ�� �⼮���� �˻�) / (�����̷°� �����ϰ� �ش� ��ȸ ���� �⼮���� �˻�)
Public Function FindByChurchIdAndDate(ByVal argChurchId As String, ByVal argDate As Date, Optional ByVal blnIsRestricted As Boolean = False) As Attendance
    
    Dim sql As String
    Dim tRecordSet As T_RECORD_SET
    
    
        
    '--//church_sid_custom �� ���
    sql = "SELECT a.church_sid_custom FROM " & TABLE_HISTORY_CHURCH_ESTABLISH & " a " & _
        " WHERE a.church_sid = " & SText(Replace(argChurchId, "MM", "MC"))
    tRecordSet = makeListData(sql, TABLE_HISTORY_CHURCH_ESTABLISH)
    
    
    '--//�⼮��ȸ
    Dim strChurchIdCustom As String
    If tRecordSet.CNT_RECORD > 0 Then
        strChurchIdCustom = tRecordSet.LISTDATA(0, 0)
        sql = "SELECT a.church_sid, a.attendance_dt,MAX(a.once_all),MAX(a.forth_all),MAX(a.once_stu), " & _
                "     MAX(a.forth_stu),MAX(a.tithe_all),MAX(a.tithe_stu),MAX(a.baptism_all),MAX(a.evangelist),MAX(a.gl),MAX(a.ul) " & _
                " FROM " & TABLE_ATTENDANCE & " a " & _
                " LEFT JOIN " & TABLE_HISTORY_CHURCH_ESTABLISH & " b " & _
                "     ON a.church_sid = b.church_sid " & _
                " WHERE b.church_sid_custom = " & SText(strChurchIdCustom) & _
                "     AND a.attendance_dt = " & SText(argDate) & _
                " GROUP BY a.attendance_dt ORDER BY a.attendance_dt;"
    End If
    
    '--//church_sid_custom�� �����ϰ� ���� ���� ���
    If blnIsRestricted Or tRecordSet.CNT_RECORD = 0 Then
        sql = "SELECT * FROM " & TABLE_ATTENDANCE & _
                " a WHERE a.church_sid = " & SText(argChurchId) & _
                " AND a.attendance_dt = " & SText(argDate) & ";"
    End If
    
    tRecordSet = makeListData(sql, TABLE_ATTENDANCE)
    
    Dim tmpAtten As New Attendance
    If tRecordSet.CNT_RECORD > 0 Then
        tmpAtten.ParseFromRecordSet tRecordSet
        tmpAtten.ChurchID = argChurchId
    Else
        Set tmpAtten = Nothing
    End If
    
    Set FindByChurchIdAndDate = tmpAtten
End Function

'@param argAtten: argAtten ��ü�� �޾� DB�� Insert �Ǵ� Update �մϴ�.
Public Sub SaveAll(ByRef attenList As Object)
    Dim tmpAtten As New Attendance
    Dim sql As String
    Dim procesureNM As String
    Dim jobNM As String
    Dim tableNM As String
    
    procesureNM = "AttendanceDao.SaveAll()"
    jobNM = "�⼮ �߰� �� ������Ʈ"
    tableNM = TABLE_ATTENDANCE
    
    Dim vAllAttenList As Object
    Set vAllAttenList = getAllAttenList
    
    Dim cnt As Long
    Dim map As New Scripting.Dictionary
    For Each tmpAtten In vAllAttenList
        map.Add tmpAtten.ChurchID & tmpAtten.AttendanceDate, tmpAtten
    Next
    
    For Each tmpAtten In attenList
        If isContain(map, tmpAtten) = True Then
            sql = sql & vbNewLine & getUpdateSql(tmpAtten)
        Else
            sql = sql & vbNewLine & getInsertSql(tmpAtten)
        End If
        
        cnt = cnt + 1
        
        If cnt Mod 1000 = 0 Then
            executeSqlWithLog sql, procesureNM, tableNM, jobNM
            cnt = 0
            sql = ""
        End If
    Next
    
    If cnt <> 0 Then
        executeSqlWithLog sql, procesureNM, tableNM, jobNM
    End If
    
End Sub

Public Function isContain(ByRef map As Scripting.Dictionary, ByRef argAtten As Attendance)

    Dim tmpAtten As Attendance
    isContain = False
    
    isContain = map.Exists(argAtten.ChurchID & argAtten.AttendanceDate)

End Function

'@param argAtten: argAtten ��ü�� �޾� DB�� Insert �Ǵ� Update �մϴ�.
Public Sub Save(ByRef argAtten As Attendance)
    Dim tmpAtten As New Attendance
    
    Set tmpAtten = FindByChurchIdAndDate(argAtten.ChurchID, argAtten.AttendanceDate, True)
    If tmpAtten.ChurchID = "" Then
        Insert argAtten
    Else
        Update argAtten
    End If
    
End Sub

Private Function getInsertSql(ByRef argAtten As Attendance)

    Dim sql As String
    Dim tableNM As String

    tableNM = TABLE_ATTENDANCE
    
    '--//INSERT ��
    sql = _
        "INSERT INTO " & tableNM & _
            " (church_sid, attendance_dt, once_all, forth_all, once_stu, forth_stu, tithe_all, tithe_stu, baptism_all, evangelist, gl, ul) " & _
        "VALUES(" & _
            SText(argAtten.ChurchID) & "," & SText(argAtten.AttendanceDate) & "," & SText(argAtten.OnceAll) & "," & SText(argAtten.ForthAll) & _
            "," & SText(argAtten.OnceStu) & "," & SText(argAtten.ForthStu) & "," & SText(argAtten.TitheAll) & "," & SText(argAtten.TitheStu) & _
            "," & SText(argAtten.BaptismAll) & "," & SText(argAtten.Evangelist) & "," & SText(argAtten.GroupLeader) & "," & SText(argAtten.UnitLeader) & ");"

    getInsertSql = sql

End Function

'@param: pPhoto: pPhoto��ü�� �޾� DB�� Insert �մϴ�.
Private Sub Insert(ByRef argAtten As Attendance)
    Dim sql As String
    Dim procesureNM As String
    Dim jobNM As String
    Dim tableNM As String
    
    procesureNM = "AttendanceDao.Insert()"
    jobNM = "�⼮ �߰�"
    tableNM = TABLE_ATTENDANCE
    
    '--//INSERT ��
    sql = getInsertSql(argAtten)
    
    executeSqlWithLog sql, procesureNM, tableNM, jobNM
    
End Sub

Private Function getUpdateSql(ByRef argAtten As Attendance)

    Dim sql As String
    Dim tableNM As String
    
    tableNM = TABLE_ATTENDANCE
    
    '--//UPDATE ��
    sql = _
        "UPDATE " & tableNM & " SET " & _
        "once_all = " & SText(argAtten.OnceAll) & ", forth_all = " & SText(argAtten.ForthAll) & _
        ", once_stu = " & SText(argAtten.OnceStu) & ", forth_stu = " & SText(argAtten.ForthStu) & _
        ", tithe_all = " & SText(argAtten.TitheAll) & ", tithe_stu = " & SText(argAtten.TitheStu) & ", baptism_all = " & SText(argAtten.BaptismAll) & ", evangelist = " & SText(argAtten.Evangelist) & _
        ", gl = " & SText(argAtten.GroupLeader) & ", ul = " & SText(argAtten.UnitLeader) & _
        " WHERE church_sid = " & SText(argAtten.ChurchID) & " AND attendance_dt = " & SText(argAtten.AttendanceDate) & ";"

    getUpdateSql = sql

End Function

'@param: pWife: pWife ��ü�� �޾� DB�� Update �մϴ�.
Private Sub Update(ByRef argAtten As Attendance)
    Dim sql As String
    Dim procesureNM As String
    Dim jobNM As String
    Dim tableNM As String
    
    procesureNM = "AttendanceDao.Update()"
    jobNM = "�⼮ ������Ʈ"
    tableNM = TABLE_ATTENDANCE
    
    '--//UPDATE ��
    sql = getUpdateSql(argAtten)
    
    executeSqlWithLog sql, procesureNM, tableNM, jobNM
    
End Sub

Public Sub Delete(ByRef argAtten As Attendance)

    Dim sql As String
    Dim procesureNM As String
    Dim jobNM As String
    Dim tableNM As String
    
    procesureNM = "AttendanceDao.Delete()"
    jobNM = "�⼮ ����"
    tableNM = TABLE_ATTENDANCE
    
    '--//UPDATE ��
    sql = _
        "DELETE a FROM " & tableNM & " a WHERE a.church_sid = " & SText(argAtten.ChurchID) & " AND c.attendance_dt = " & SText(argAtten.AttendanceDate) & ";"
    
    executeSqlWithLog sql, procesureNM, tableNM, jobNM

End Sub

Public Function getAllAttenList() As Object

    Dim attenList As Object
    Set attenList = CreateObject("System.Collections.Arraylist")
    
    Dim tRecordSet As T_RECORD_SET
    Dim sql As String
    
    sql = _
    "SELECT * FROM op_system.db_attendance a " & _
    "WHERE a.attendance_dt > ADDDATE(CURDATE(),INTERVAL -2 YEAR);"
    
    tRecordSet = makeListData(sql, TABLE_ATTENDANCE)
    
    Dim attenTemp As Attendance
    Dim i As Long
    If tRecordSet.CNT_RECORD > 0 Then
        For i = 0 To tRecordSet.CNT_RECORD - 1
            Set attenTemp = New Attendance
            attenTemp.ParseFromRecordSet tRecordSet, i
            attenList.Add attenTemp
        Next
    Else
        Set attenTemp = Nothing
    End If
    
    Set getAllAttenList = attenList

End Function