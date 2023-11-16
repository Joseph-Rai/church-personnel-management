VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PastoralCareerDao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'@param argLifeNo: 생명번호를 기준으로 발령이력 목록을 검색합니다.
Public Function GetPastoralCareers(ByVal argLifeNo As String, Optional argCareerCount As Integer = 10) As Object

    Dim sql As String
    Dim tRecordSet As T_RECORD_SET
    
    sql = "SELECT * FROM (SELECT * FROM " & VIEW_PASTORAL_CAREER & " WHERE `생명번호` = " & SText(argLifeNo) & " AND `직분/직책` IS NOT NULL LIMIT " & argCareerCount & ") a ORDER BY `발령일`;"
    
    tRecordSet = makeListData(sql, VIEW_PASTORAL_CAREER)
    
    Dim objPastoralCareer As PastoralCareer
    Dim objPastoralCareerList As Object
    Set objPastoralCareerList = CreateObject("System.Collections.ArrayList")
    
    With tRecordSet
        If .CNT_RECORD > 0 Then
            Dim i As Long
            For i = 0 To .CNT_RECORD - 1
                Set objPastoralCareer = New PastoralCareer
                objPastoralCareer.ParseFromRecordSet tRecordSet, i
                objPastoralCareerList.Add objPastoralCareer
            Next
        Else
'            objPastoralCareerList = Nothing
        End If
    End With
    
    Set GetPastoralCareers = objPastoralCareerList

End Function

'@param LifeNo: 생명번호를 기준으로 동역 시작일을 반환합니다.
Public Function GetMinDateForAssistantOverseer(lifeNo As String) As String

    Dim sql As String
    Dim tRecordSet As T_RECORD_SET

    sql = "" & _
        " SELECT MIN(p.Start_dt)" & _
        " FROM " & TABLE_POSITION & " p" & _
        " WHERE p.lifeno = " & SText(lifeNo) & _
        "     AND p.Position IN ('동역');"
    tRecordSet = makeListData(sql, TABLE_POSITION)
    
    Dim result As String
    
    If tRecordSet.CNT_RECORD > 0 Then
        result = tRecordSet.LISTDATA(0, 0)
    Else
        result = ""
    End If
    
    GetMinDateForAssistantOverseer = result

End Function

'@param LifeNo: 생명번호를 기준으로 당회장 시작일을 반환합니다.
Public Function GetMinDateForOverseer(lifeNo As String) As String

    Dim sql As String
    Dim tRecordSet As T_RECORD_SET

    sql = "" & _
        " SELECT MIN(p.Start_dt)" & _
        " FROM " & TABLE_POSITION & " p" & _
        " WHERE p.lifeno = " & SText(lifeNo) & _
        "     AND p.Position IN ('당회장', '당회장대리');"
    tRecordSet = makeListData(sql, TABLE_POSITION)
    
    Dim result As String
    If tRecordSet.CNT_RECORD > 0 Then
        result = tRecordSet.LISTDATA(0, 0)
    Else
        result = ""
    End If
    
    GetMinDateForOverseer = result

End Function

'@param LifeNo: 생명번호를 기준으로 동역 경력을 반환합니다.
Public Function GetAssistantOverseerCareer(lifeNo As String) As String

    Dim sql As String
    Dim tRecordSet As T_RECORD_SET

    '--//동역으로 활동한 이력 추출
    sql = "" & _
        " SELECT p.Start_dt, p.End_dt" & _
        " FROM " & TABLE_POSITION & " p" & _
        " WHERE p.lifeno = " & SText(lifeNo) & _
        "     AND p.Position IN ('동역');"
    tRecordSet = makeListData(sql, TABLE_POSITION)
    
    Dim result As String
    result = GetConvertedFormatPeriod(tRecordSet)
    
    GetAssistantOverseerCareer = result

End Function

'@param LifeNo: 생명번호를 기준으로 당회장 이력을 반환합니다.
Public Function GetOverseerCareer(lifeNo As String) As String

    Dim sql As String
    Dim tRecordSet As T_RECORD_SET

    '--//당회장으로 활동한 이력 추출
    sql = "" & _
        " SELECT p.Start_dt, p.End_dt" & _
        " FROM " & TABLE_POSITION & " p" & _
        " WHERE p.lifeno = " & SText(lifeNo) & _
        "     AND p.Position IN ('당회장', '당회장대리');"
    tRecordSet = makeListData(sql, TABLE_POSITION)
    
    Dim result As String
    result = GetConvertedFormatPeriod(tRecordSet)
    
    GetOverseerCareer = result

End Function

'@param LifeNo: 생명번호를 기준으로 목회 경력을 반환합니다.
Public Function GetTotalPastoralCareer(lifeNo As String) As String
    
    Dim sql As String
    Dim tRecordSet As T_RECORD_SET
    
    '--//목회자로 활동한 이력 추출
    sql = "" & _
        " SELECT p.Start_dt, p.End_dt" & _
        " FROM " & TABLE_POSITION & " p" & _
        " WHERE p.lifeno = " & SText(lifeNo) & _
        "     AND p.Position IN ('당회장', '당회장대리', '동역');"
    tRecordSet = makeListData(sql, TABLE_POSITION)
    
    Dim result As String
    result = GetConvertedFormatPeriod(tRecordSet)
    
    GetTotalPastoralCareer = result

End Function

'@param LifeNo: 생명번호를 기준으로 동역 시작일 반환합니다.
Public Function GetConvertedFormatPeriod(argRecordSet As T_RECORD_SET) As String

    Dim year As Integer
    Dim month As Integer

    '--//탐색을 위해 minDate, maxDate 검출
    Dim minDate As Date: minDate = DateSerial(9999, 12, 31)
    Dim maxDate As Date: maxDate = DateSerial(1900, 1, 1)
    
    If argRecordSet.CNT_RECORD <= 0 Then
        GetConvertedFormatPeriod = ""
    End If
    
    Dim i As Integer
    For i = 0 To argRecordSet.CNT_RECORD - 1
        minDate = WorksheetFunction.Min(minDate, argRecordSet.LISTDATA(i, 0))
        maxDate = WorksheetFunction.Max(maxDate, WorksheetFunction.Min(Date, argRecordSet.LISTDATA(i, 1)))
    Next
    
    '--//한 달씩 건너뛰며 목회경력에 포함된 날짜라면 개월수 추가
    Dim tempDate As Date: tempDate = WorksheetFunction.EoMonth(minDate, 0)
    Do
        For i = 0 To argRecordSet.CNT_RECORD - 1
            Dim startDate As Date: startDate = argRecordSet.LISTDATA(i, 0)
            Dim endDate As Date: endDate = argRecordSet.LISTDATA(i, 1)
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

