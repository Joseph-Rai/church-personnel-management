VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ChurchDao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'--//

'--//DB에 저장된 교회리스트를 가져옵니다.
Public Function GetChurchList() As Object

    Dim churchList As Object
    Dim sql As String
    Dim tRecordSet As T_RECORD_SET
    
    Set churchList = CreateObject("System.Collections.ArrayList")
    
    sql = _
        "SELECT c.ctry_nm FROM " & TABLE_CHURCH & " c;"
    
    tRecordSet = makeListData(sql, TABLE_CHURCH)
    
    Dim recordChurch As Variant
    If tRecordSet.CNT_RECORD > 0 Then
        For Each recordChurch In tRecordSet.LISTDATA
            Dim churchTemp As New Church
            churchTemp.ParseFromRecordSet recordChurch
            churchList.Add churchTemp
        Next
    End If
    
    Set GetChurchList = churchList

End Function

'--//매개변수 MC와 관련된 현재 본교회,지교회,예배소 목록을 가져옵니다.
'@param argMC: MC자신, MM(순수본교회) 및 MC가 관리하는 BC, PBC의 개수를 세어 ArrayList로 반환
Public Function GetMcBcPbcCount(ByRef argMC As Church, Optional ByRef searchDate As Date = "1900-01-01") As Object

    Dim collectionCountMcBcPbc As New Collection
    
    Dim lstMcBcPbc As Object
    Set lstMcBcPbc = GetMcBcPbcList(argMC, searchDate)
    
    Dim cntMC As Integer, cntMM As Integer, cntBC As Integer, cntPBC As Integer
    
    Dim tmpChurch As New Church
    For Each tmpChurch In lstMcBcPbc
        Select Case tmpChurch.Gb
        Case "MC"
            cntMC = cntMC + 1
        Case "MM"
            cntMM = cntMM + 1
        Case "BC"
            cntBC = cntBC + 1
        Case "PBC"
            cntPBC = cntPBC + 1
        End Select
    Next
    
    With collectionCountMcBcPbc
        .Add cntMC, "MC"
        .Add cntMM, "MM"
        .Add cntBC, "BC"
        .Add cntPBC, "PBC"
    End With
    
    Set GetMcBcPbcCount = collectionCountMcBcPbc
    
End Function

'--//매개변수 MC와 관련된 현재 본교회,지교회,예배소 목록을 가져옵니다.
'@param argMC: MC자신, MM(순수본교회) 및 MC가 관리하는 BC, PBC 목록을 가져옵니다.
Public Function GetMcBcPbcList(ByRef argMC As Church, Optional ByRef searchDate As Date = "1900-01-01") As Object

    Dim churchList As Object
    Dim sql As String
    Dim tRecordSet As T_RECORD_SET
    Dim objAtten As New Attendance
    Dim objAttenDao As New AttendanceDao
    
    Set churchList = CreateObject("System.Collections.ArrayList")
    
    If searchDate = "1900-01-01" Then
        Dim lstAtten As Object
        Set lstAtten = objAttenDao.GetAllAttenByChurchId(argMC.Id)
        searchDate = objAttenDao.GetMaxDate(lstAtten)
    End If
    sql = "CALL `Routine_churchlist_by_time`(" & SText(searchDate) & "," & SText(USER_DEPT) & ");"
    connectTaskDB
    executeSQL "GetMcBcPbcList", "Routine_churchlist_by_time", sql
    disconnectALL
        
    sql = _
        "SELECT * FROM " & TEMP_CHURCHLIST_BY_TIME & " c " & _
        " WHERE c.church_sid = " & SText(argMC.Id) & _
        "    OR c.main_church_cd = " & SText(argMC.Id) & ";"
    tRecordSet = makeListData(sql, TEMP_CHURCHLIST_BY_TIME)
    
    Dim recordChurch As Variant
    Dim i As Long
    If tRecordSet.CNT_RECORD > 0 Then
        For i = 0 To tRecordSet.CNT_RECORD - 1
            Dim churchTemp As Church
            Set churchTemp = New Church
            churchTemp.ParseFromRecordSet tRecordSet, i
            churchList.Add churchTemp
            If churchTemp.Gb = "MC" Then '--//순수본교회 리스트 추가
                Dim churchTemp2 As Church
                Set churchTemp2 = New Church
                churchTemp2.ParseFromRecordSet tRecordSet, i
                churchTemp2.Gb = "MM"
                churchTemp2.Id = Replace(churchTemp2.Id, "MC", "MM")
                churchList.Add churchTemp2
            End If
        Next
    End If
    
    Set GetMcBcPbcList = churchList

End Function

'@param ChurchCode: 부서명을 기준으로 DB에서 검색합니다.
Public Function FindByChurchCode(ByVal churchCode As String) As Church
    
    Dim sql As String
    Dim tRecordSet As T_RECORD_SET
    
    sql = "SELECT * FROM " & TABLE_CHURCH & _
                    " c WHERE c.church_sid = " & SText(churchCode) & ";"
    
    tRecordSet = makeListData(sql, TABLE_CHURCH)
    
    Dim churchTemp As New Church
    If tRecordSet.CNT_RECORD > 0 Then
        churchTemp.ParseFromRecordSet tRecordSet
    Else
        Set churchTemp = Nothing
    End If
    
    Set FindByChurchCode = churchTemp
End Function

'@param churchName: 부서명을 기준으로 DB에서 검색합니다.
Public Function FindByChurchName(ByVal churchName As String) As Church
    
    Dim sql As String
    Dim tRecordSet As T_RECORD_SET
    
    sql = "SELECT * FROM " & TABLE_CHURCH & _
                    " c WHERE c.church_nm = " & SText(churchName) & ";"
    
    tRecordSet = makeListData(sql, TABLE_CHURCH)
    
    Dim churchTemp As New Church
    If tRecordSet.CNT_RECORD > 0 Then
        churchTemp.ParseFromRecordSet tRecordSet
    Else
        Set churchTemp = Nothing
    End If
    
    Set FindByChurchName = churchTemp
End Function

'@param argChurchList: argChurch 객체를 받아 DB에 Insert 또는 Update 합니다.
Public Sub SaveAll(ByRef argChurchList As Object)
    Dim tmpChurch As New Church
    Dim sql As String
    Dim procesureNM As String
    Dim jobNM As String
    Dim tableNM As String
    
    procesureNM = "ChurchDao.SaveAll()"
    jobNM = "교회리스트 월 정기 업데이트"
    tableNM = TABLE_CHURCH
    
'    Dim vAllChurchList As Object
'    Set vAllChurchList = getAllChurchList
    
    Dim cnt As Long
'    Dim map As New Scripting.Dictionary
'    For Each tmpChurch In vAllChurchList
'        map.Add tmpChurch.Id, tmpChurch
'    Next
    
    sql = "TRUNCATE TABLE " & tableNM & ";" '--//기존에 있던 모든 데이터 삭제 후 삽입
    For Each tmpChurch In argChurchList
        sql = sql & vbNewLine & getInsertSql(tmpChurch)
'        If map.Exists(tmpChurch.Id) = True Then
'            sql = sql & vbNewLine & getUpdateSql(tmpChurch)
'        Else
'            sql = sql & vbNewLine & getInsertSql(tmpChurch)
'        End If
        
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

'@param argChurch: argChurch 객체를 받아 DB에 Insert 또는 Update 합니다.
Public Sub Save(ByRef argChurch As Church)
    Dim churchTemp As New Church
    
    Set churchTemp = FindByChurchName(argChurch.Name)
    If churchTemp.Id = "" Then
        Insert argChurch
    Else
        Update argChurch
    End If
    
End Sub

Private Function getInsertSql(ByRef argChurch As Church)

    Dim sql As String
    Dim tableNM As String

    tableNM = TABLE_CHURCH
    
    '--//INSERT 문
    sql = _
        "INSERT INTO " & tableNM & _
            " (church_sid, church_nm, church_gb, manager_cd, main_church_cd, start_dt, end_dt, ovs_dept, suspend, sort_order, geo_cd) " & _
        "VALUES(" & _
            SText(argChurch.Id) & "," & SText(argChurch.Name) & "," & SText(argChurch.Gb) & "," & SText(argChurch.Manager) & _
            "," & SText(argChurch.MainChurch) & "," & SText(argChurch.startDate) & "," & SText(argChurch.endDate) & _
            "," & SText(argChurch.OvsDeptName) & "," & SText(IIf(argChurch.Suspend, 1, 0)) & "," & SText(argChurch.SortOrder) & "," & SText(argChurch.GeoCode) & ");"

    getInsertSql = sql

End Function

'@param: pPhoto: pPhoto객체를 받아 DB에 Insert 합니다.
Private Sub Insert(ByRef argChurch As Church)
    Dim sql As String
    Dim procesureNM As String
    Dim jobNM As String
    Dim tableNM As String
    
    procesureNM = "ChurchDao.Insert()"
    jobNM = "교회 추가"
    tableNM = TABLE_CHURCH
    
    '--//INSERT 문
    sql = getInsertSql(argChurch)
    
    executeSqlWithLog sql, procesureNM, tableNM, jobNM
    
End Sub

Private Function getUpdateSql(ByRef argChurch As Church)

    Dim sql As String
    Dim tableNM As String
    
    tableNM = TABLE_CHURCH
    
    '--//UPDATE 문
    sql = _
        "UPDATE " & tableNM & " SET " & _
        "church_nm = " & SText(argChurch.Name) & ", church_gb = " & SText(argChurch.Gb) & ", manager_cd = " & SText(argChurch.Manager) & ", main_church_cd = " & SText(argChurch.MainChurch) & _
        ", start_dt = " & SText(argChurch.startDate) & ", end_dt = " & SText(argChurch.endDate) & ", ovs_dept = " & SText(argChurch.OvsDeptName) & ", suspend = " & SText(IIf(argChurch.Suspend, 1, 0)) & _
        ", sort_order = " & SText(argChurch.SortOrder) & ", geo_cd = " & SText(argChurch.GeoCode) & _
        " WHERE church_sid = " & SText(argChurch.Id) & ";"
    
    getUpdateSql = sql

End Function

'@param: pWife: pWife 객체를 받아 DB에 Update 합니다.
Private Sub Update(ByRef argChurch As Church)
    Dim sql As String
    Dim procesureNM As String
    Dim jobNM As String
    Dim tableNM As String
    
    procesureNM = "ChurchDao.Update()"
    jobNM = "교회 업데이트"
    tableNM = TABLE_CHURCH
    
    '--//UPDATE 문
    sql = getUpdateSql(argChurch)
    
    executeSqlWithLog sql, procesureNM, tableNM, jobNM
    
End Sub

Public Sub Delete(ByRef argChurch As Church)

    Dim sql As String
    Dim procesureNM As String
    Dim jobNM As String
    Dim tableNM As String
    
    procesureNM = "ChurchDao.Delete()"
    jobNM = "교회 삭제"
    tableNM = TABLE_CHURCH
    
    '--//UPDATE 문
    sql = _
        "DELETE c FROM " & tableNM & " c WHERE c.church_sid = " & SText(argChurch.Id) & ";"
    
    executeSqlWithLog sql, procesureNM, tableNM, jobNM

End Sub

Public Function CountBranchedOut(ByVal argChurchCode, ByVal argStartDate As Date, ByVal argEndDate As Date) As Integer

    Dim sql As String
    Dim tRecordSet As T_RECORD_SET
    
'    sql = "" & _
'        "SELECT * FROM " & TABLE_CHURCH & " churchlist " & _
'        "LEFT JOIN " & TABLE_CHURCH & " mainchurch " & _
'        "    ON churchlist.main_church_cd = mainchurch.church_sid " & _
'        "WHERE mainchurch.church_sid = " & SText(argChurchCode) & _
'        "    AND churchlist.start_dt BETWEEN " & SText(argStartDate) & " AND  " & SText(argEndDate)
    sql = "" & _
        "SELECT * FROM op_system.v_history_church hc " & _
        "WHERE hc.`교회연혁` LIKE '%분가%' " & _
        "    AND hc.`교회코드` = " & SText(argChurchCode) & _
        "    AND hc.`날짜` BETWEEN " & SText(argStartDate) & " AND " & SText(argEndDate)
    
    tRecordSet = makeListData(sql, TABLE_CHURCH)
    
    CountBranchedOut = tRecordSet.CNT_RECORD
    
End Function

Public Function getAllChurchList(Optional ByVal isReverseOrder As Boolean = False) As Object

    Dim churchList As Object
    Set churchList = CreateObject("System.Collections.Arraylist")
    
    Dim tRecordSet As T_RECORD_SET
    Dim sql As String
    
    sql = _
    "SELECT * FROM op_system.db_churchlist_custom a ORDER BY a.sort_order"
    
    If isReverseOrder = True Then sql = sql & " DESC"
    sql = sql & ";"
    
    tRecordSet = makeListData(sql, TABLE_CHURCH)
    
    Dim churchTemp As Church
    Dim i As Long
    If tRecordSet.CNT_RECORD > 0 Then
        For i = 0 To tRecordSet.CNT_RECORD - 1
            Set churchTemp = New Church
            churchTemp.ParseFromRecordSet tRecordSet, i
            churchList.Add churchTemp
        Next
    Else
        Set churchTemp = Nothing
    End If
    
    Set getAllChurchList = churchList

End Function

Public Function getAllChurchListReverseOrder() As Object
    
    Set getAllChurchListReverseOrder = getAllChurchList(True)

End Function



