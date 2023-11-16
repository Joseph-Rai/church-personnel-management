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

'@param argLifeNo: ������ȣ�� �������� �߷��̷� ����� �˻��մϴ�.
Public Function GetPastoralCareers(ByVal argLifeNo As String, Optional argCareerCount As Integer = 10) As Object

    Dim sql As String
    Dim tRecordSet As T_RECORD_SET
    
    sql = "SELECT * FROM (SELECT * FROM " & VIEW_PASTORAL_CAREER & " WHERE `������ȣ` = " & SText(argLifeNo) & " AND `����/��å` IS NOT NULL LIMIT " & argCareerCount & ") a ORDER BY `�߷���`;"
    
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

'@param LifeNo: ������ȣ�� �������� ���� �������� ��ȯ�մϴ�.
Public Function GetMinDateForAssistantOverseer(lifeNo As String) As String

    Dim sql As String
    Dim tRecordSet As T_RECORD_SET

    sql = "" & _
        " SELECT MIN(p.Start_dt)" & _
        " FROM " & TABLE_POSITION & " p" & _
        " WHERE p.lifeno = " & SText(lifeNo) & _
        "     AND p.Position IN ('����');"
    tRecordSet = makeListData(sql, TABLE_POSITION)
    
    Dim result As String
    
    If tRecordSet.CNT_RECORD > 0 Then
        result = tRecordSet.LISTDATA(0, 0)
    Else
        result = ""
    End If
    
    GetMinDateForAssistantOverseer = result

End Function

'@param LifeNo: ������ȣ�� �������� ��ȸ�� �������� ��ȯ�մϴ�.
Public Function GetMinDateForOverseer(lifeNo As String) As String

    Dim sql As String
    Dim tRecordSet As T_RECORD_SET

    sql = "" & _
        " SELECT MIN(p.Start_dt)" & _
        " FROM " & TABLE_POSITION & " p" & _
        " WHERE p.lifeno = " & SText(lifeNo) & _
        "     AND p.Position IN ('��ȸ��', '��ȸ��븮');"
    tRecordSet = makeListData(sql, TABLE_POSITION)
    
    Dim result As String
    If tRecordSet.CNT_RECORD > 0 Then
        result = tRecordSet.LISTDATA(0, 0)
    Else
        result = ""
    End If
    
    GetMinDateForOverseer = result

End Function

'@param LifeNo: ������ȣ�� �������� ���� ����� ��ȯ�մϴ�.
Public Function GetAssistantOverseerCareer(lifeNo As String) As String

    Dim sql As String
    Dim tRecordSet As T_RECORD_SET

    '--//�������� Ȱ���� �̷� ����
    sql = "" & _
        " SELECT p.Start_dt, p.End_dt" & _
        " FROM " & TABLE_POSITION & " p" & _
        " WHERE p.lifeno = " & SText(lifeNo) & _
        "     AND p.Position IN ('����');"
    tRecordSet = makeListData(sql, TABLE_POSITION)
    
    Dim result As String
    result = GetConvertedFormatPeriod(tRecordSet)
    
    GetAssistantOverseerCareer = result

End Function

'@param LifeNo: ������ȣ�� �������� ��ȸ�� �̷��� ��ȯ�մϴ�.
Public Function GetOverseerCareer(lifeNo As String) As String

    Dim sql As String
    Dim tRecordSet As T_RECORD_SET

    '--//��ȸ������ Ȱ���� �̷� ����
    sql = "" & _
        " SELECT p.Start_dt, p.End_dt" & _
        " FROM " & TABLE_POSITION & " p" & _
        " WHERE p.lifeno = " & SText(lifeNo) & _
        "     AND p.Position IN ('��ȸ��', '��ȸ��븮');"
    tRecordSet = makeListData(sql, TABLE_POSITION)
    
    Dim result As String
    result = GetConvertedFormatPeriod(tRecordSet)
    
    GetOverseerCareer = result

End Function

'@param LifeNo: ������ȣ�� �������� ��ȸ ����� ��ȯ�մϴ�.
Public Function GetTotalPastoralCareer(lifeNo As String) As String
    
    Dim sql As String
    Dim tRecordSet As T_RECORD_SET
    
    '--//��ȸ�ڷ� Ȱ���� �̷� ����
    sql = "" & _
        " SELECT p.Start_dt, p.End_dt" & _
        " FROM " & TABLE_POSITION & " p" & _
        " WHERE p.lifeno = " & SText(lifeNo) & _
        "     AND p.Position IN ('��ȸ��', '��ȸ��븮', '����');"
    tRecordSet = makeListData(sql, TABLE_POSITION)
    
    Dim result As String
    result = GetConvertedFormatPeriod(tRecordSet)
    
    GetTotalPastoralCareer = result

End Function

'@param LifeNo: ������ȣ�� �������� ���� ������ ��ȯ�մϴ�.
Public Function GetConvertedFormatPeriod(argRecordSet As T_RECORD_SET) As String

    Dim year As Integer
    Dim month As Integer

    '--//Ž���� ���� minDate, maxDate ����
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
    
    '--//�� �޾� �ǳʶٸ� ��ȸ��¿� ���Ե� ��¥��� ������ �߰�
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

    '--//year, month => Y�� M���� �������� ��ȯ
    Dim result As String
    If year > 0 Then
        result = result & year & "��"
    End If
    
    If month > 0 Then
        If result = "" Then
            result = month & "����"
        Else
            result = result & " " & month & "����"
        End If
    End If
    
    GetConvertedFormatPeriod = result

End Function
