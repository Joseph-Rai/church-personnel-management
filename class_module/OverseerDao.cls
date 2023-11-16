VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OverseerDao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'--//특정 기간에 해당 교회를 관리했던 당회장 목록을 반환
'@param argChurchCode: 검색하고자 하는 교회의 코드
'@param argStartDate: 검색하고자 하는 기간의 시작일
'@param argEndDate: 검색하고자 하는 기간의 종료일
Public Function GetOverseersBetween(ByVal argChurchCode, ByVal argStartDate As Date, ByVal argEndDate As Date) As Object

    Dim sql As String
    Dim tRecordSet As T_RECORD_SET
    
    sql = "" & _
        "SELECT " & _
        "    t.church_sid " & _
        "    ,c.church_nm " & _
        "    ,t.start_dt " & _
        "    ,t.end_dt " & _
        "    ,p.lifeno " & _
        "    ,p.name_ko " & _
        "    ,ti.title " & _
        "    ,pos.position " & _
        "FROM " & TABLE_TRANSFER & " t " & _
        "INNER JOIN " & TABLE_CHURCH & " c " & _
        "    ON t.church_sid = c.church_sid " & _
        "INNER JOIN " & TABLE_PASTORALSTAFF & " p " & _
        "    ON t.lifeno = p.lifeno " & _
        "INNER JOIN " & TABLE_TITLE & " ti " & _
        "    ON p.lifeno = ti.lifeno " & _
        "        AND t.start_dt BETWEEN ti.start_dt AND ti.end_dt " & _
        "INNER JOIN " & TABLE_POSITION & " pos " & _
        "    ON p.lifeno = pos.lifeno " & _
        "        AND  pos.position LIKE '%당%' " & _
        "        AND t.start_dt BETWEEN pos.start_dt AND pos.end_dt " & _
        "WHERE t.church_sid = " & SText(argChurchCode) & " " & _
        "    AND (" & SText(argStartDate) & " < t.end_dt AND " & SText(argEndDate) & " > t.start_dt) " & _
        "ORDER BY t.start_dt;"
    tRecordSet = makeListData(sql, TABLE_CHURCH)
    
    Dim objOverseers As Object
    Set objOverseers = CreateObject("System.Collections.ArrayList")
    
    Dim i As Long
    Dim objoverseer As Overseer
    For i = 0 To tRecordSet.CNT_RECORD - 1
        Set objoverseer = New Overseer
        objoverseer.ParseFromRecordSet tRecordSet, i
        objOverseers.Add objoverseer
    Next
    
    Set GetOverseersBetween = objOverseers
    
End Function

