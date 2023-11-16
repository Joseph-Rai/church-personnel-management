VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PStaffInfoViewDao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function FindByChurchName(ByVal argChurchName As String, Optional ByVal blnStrict As Boolean = False) As Object

    Dim result As Object
    Dim sql As String
    Dim tRecordSet As T_RECORD_SET
    Set result = CreateObject("System.Collections.ArrayList")
    
    If blnStrict Then
        sql = "SELECT * FROM " & VIEW_PSTAFF_INFORMATION & " p " & _
            " WHERE p.`지교회명` = " & SText(argChurchName) & ";"
    Else
        sql = "SELECT * FROM " & VIEW_PSTAFF_INFORMATION & " p " & _
            " WHERE p.`지교회명` LIKE '%" & argChurchName & "%';"
    End If
    
    tRecordSet = makeListData(sql, VIEW_PSTAFF_INFORMATION)
    
    Dim tmpPStaffView As PStaffInfoView
    If tRecordSet.CNT_RECORD > 0 Then
        Dim i As Long
        For i = 0 To tRecordSet.CNT_RECORD - 1
            Set tmpPStaffView = New PStaffInfoView
            tmpPStaffView.ParseFromRecordSet tRecordSet, i
            result.Add tmpPStaffView
        Next
    Else
'        Set tmpPStaffView = Nothing
    End If
    
    Set FindByChurchName = result

End Function

'@param lifeNo: 생명번호를 기준으로 DB에서 검색합니다.
Public Function FindBySearchText(ByVal searchText As String, ByVal blnChkAll As Boolean) As Object
    
    Dim sql As String
    Dim tRecordSet As T_RECORD_SET
    Dim pStaffInfoList As Object
    Set pStaffInfoList = CreateObject("System.Collections.ArrayList")
    
    Dim table As String
    
    If blnChkAll Then
        table = VIEW_PSTAFF_INFORMATION_ALL
    Else
        table = VIEW_PSTAFF_INFORMATION
    End If
    
    sql = "SELECT * FROM " & table & " p " & _
        " WHERE (p.`한글이름(직분)` LIKE '%" & searchText & "%'" & _
        "    OR p.`영문이름` LIKE '%" & searchText & "%'" & _
        "    OR p.`교회명` LIKE '%" & searchText & "%'" & _
        "    OR p.`영문교회명` LIKE '%" & searchText & "%'" & _
        "    OR p.`지교회명` LIKE '%" & searchText & "%'" & _
        "    OR p.`영문지교회명` LIKE '%" & searchText & "%'" & _
        "    OR p.`생명번호` LIKE '%" & searchText & "%'" & _
        "    OR p.`배우자생번` LIKE '%" & searchText & "%'" & _
        "    OR p.`사모한글이름(직분)` LIKE '%" & searchText & "%'" & _
        "    OR p.`사모영문이름` LIKE '%" & searchText & "%')" & _
        "    AND p.`관리부서` = " & SText(USER_DEPT)

    If Not blnChkAll Then
        sql = sql & "    AND p.`교회코드` IS NOT NULL"
    End If
    
    sql = sql & " ORDER BY `교회명`," & GetDefaultOrderByPosition
    
    tRecordSet = makeListData(sql, table)
    
    Dim pStaffInfo As PStaffInfoView
    If tRecordSet.CNT_RECORD > 0 Then
        Dim i As Long
        For i = 0 To tRecordSet.CNT_RECORD - 1
            Set pStaffInfo = New PStaffInfoView
            pStaffInfo.ParseFromRecordSet tRecordSet, i
            pStaffInfoList.Add pStaffInfo
        Next
    Else
        Set pStaffInfo = Nothing
    End If
    
    Set FindBySearchText = pStaffInfoList
    
End Function

'@param argLifeNo: 생명번호를 기준으로 DB에서 검색합니다.
Public Function FindByLifeNo(ByVal argLifeNo As String) As PStaffInfoView
    
    Dim sql As String
    Dim tRecordSet As T_RECORD_SET
    
    sql = "SELECT * FROM " & VIEW_PSTAFF_INFORMATION & " p " & _
        " WHERE p.`생명번호` LIKE '%" & argLifeNo & "%'" & _
        "    AND p.`교회코드` IS NOT NULL" & _
        " ORDER BY `교회명`," & GetDefaultOrderByPosition
    
    tRecordSet = makeListData(sql, VIEW_PSTAFF_INFORMATION)
    
    Dim tmpPStaffView As New PStaffInfoView
    If tRecordSet.CNT_RECORD > 0 Then
        tmpPStaffView.ParseFromRecordSet tRecordSet
    Else
        Set tmpPStaffView = Nothing
    End If
    
    Set FindByLifeNo = tmpPStaffView
    
End Function
