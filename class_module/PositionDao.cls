VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PositionDao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'@param argCode: 코드값을 기준으로 DB에서 검색합니다.
Public Function FindByCode(ByVal argCode As Integer) As position
    
    Dim sql As String
    Dim tRecordSet As T_RECORD_SET
    
    sql = _
        "SELECT * FROM " & TABLE_POSITION & " t" & _
        "    WHERE t.position_cd = " & SText(argCode) & ";"
    
    tRecordSet = makeListData(sql, TABLE_POSITION)
    
    Dim objPosition As New position
    If tRecordSet.CNT_RECORD > 0 Then
        objPosition.ParseFromRecordSet tRecordSet
    Else
        Set objPosition = Nothing
    End If
    
    Set FindByCode = objPosition
    
End Function

'@param argLifeNo: 생명번호를 기준으로 현재 직책을 DB에서 검색합니다.
Public Function FindPositionByLifeNoAndDate(ByVal argLifeNo As String, argDate As Date) As position

    Dim sql As String
    Dim tRecordSet As T_RECORD_SET
    
    sql = _
        "SELECT * FROM " & TABLE_POSITION & " t" & _
        "    WHERE t.lifeno = " & SText(argLifeNo) & _
        "        AND " & SText(Format(argDate, "yyyy-mm-dd")) & " BETWEEN t.start_dt AND t.end_dt;"
    
    tRecordSet = makeListData(sql, TABLE_POSITION)
    
    Dim objPosition As position
    If tRecordSet.CNT_RECORD > 0 Then
        Set objPosition = New position
        objPosition.ParseFromRecordSet tRecordSet
    End If
    
    Set FindPositionByLifeNoAndDate = objPosition

End Function

'@param argLifeNo: 생명번호를 기준으로 DB에서 검색합니다. 리스트를 반환합니다.
Public Function FindByLifeNo(ByVal argLifeNo As String) As Object
    
    Dim sql As String
    Dim tRecordSet As T_RECORD_SET
    Dim positionList As Object
    Set positionList = CreateObject("System.Collections.ArrayList")
    
    sql = _
        "SELECT * FROM " & TABLE_POSITION & " t" & _
        "    WHERE t.lifeno = " & SText(argLifeNo) & _
        "    ORDER BY t.start_dt;"
    
    tRecordSet = makeListData(sql, TABLE_POSITION)
    
    Dim objPosition As position
    If tRecordSet.CNT_RECORD > 0 Then
        Dim i As Long
        For i = 0 To tRecordSet.CNT_RECORD - 1
            Set objPosition = New position
            objPosition.ParseFromRecordSet tRecordSet, i
            positionList.Add objPosition
        Next
    Else
        Set positionList = Nothing
    End If
    
    Set FindByLifeNo = positionList
    
End Function

'@param argLifeNo: 생명번호를 기준으로 DB에서 검색합니다. 리스트를 반환합니다.
Public Function FindByPosition(ByRef argPosition As position) As position
    
'    Dim sql As String
'    Dim tRecordSet As T_RECORD_SET
'
'    sql = _
'        "SELECT * FROM " & TABLE_POSITION & " t" & _
'        "    WHERE t.lifeno = " & SText(argPosition.LifeNo) & _
'        "      AND t.start_dt = " & SText(argPosition.StartDate) & _
'        "      AND t.Position = " & SText(argPosition.Position) & ";"
'
'    tRecordSet = makeListData(sql, TABLE_POSITION)
'
'    Dim objPosition As Position
'    If tRecordSet.CNT_RECORD > 0 Then
'        objPosition.ParseFromRecordSet tRecordSet
'    Else
'        Set objPosition = Nothing
'    End If
    
    Set FindByPosition = FindByCode(argPosition.Code)
    
End Function

'@param argPosition: argPosition 객체를 받아 DB에 Insert 또는 Update 합니다.
Public Sub Save(ByRef argPosition As position)
    Dim tmpPosition As New position
    
    Set tmpPosition = FindByPosition(argPosition)
    If tmpPosition.Code = 0 Then
        Insert argPosition
    Else
        Update argPosition
    End If
    
End Sub

'@param: argPosition: argPosition 객체를 받아 DB에 Insert 합니다.
Private Sub Insert(ByRef argPosition As position)
    Dim sql As String
    Dim procesureNM As String
    Dim jobNM As String
    Dim tableNM As String
    
    procesureNM = "PositionDao.Insert()"
    jobNM = "직책이력 추가"
    tableNM = TABLE_POSITION
    
    '--//INSERT 문
    sql = _
        "INSERT INTO " & tableNM & _
            " (lifeno, start_dt, end_dt, position) " & _
        "VALUES(" & _
            SText(argPosition.lifeNo) & "," & SText(argPosition.startDate) & "," & _
            SText(argPosition.endDate) & "," & SText(argPosition.position) & _
        ");"
    
    executeSqlWithLog sql, procesureNM, tableNM, jobNM
    
End Sub

'@param: argPosition: argPosition 객체를 받아 DB에 Update 합니다.
Private Sub Update(ByRef argPosition As position)
    Dim sql As String
    Dim procesureNM As String
    Dim jobNM As String
    Dim tableNM As String
    
    procesureNM = "PositionDao.Update()"
    jobNM = "직책이력 업데이트"
    tableNM = TABLE_POSITION
    
    '--//UPDATE 문
    sql = _
        "UPDATE " & tableNM & " SET " & _
        "lifeno = " & SText(argPosition.lifeNo) & ", start_dt = " & SText(argPosition.startDate) & _
        ", end_dt = " & SText(argPosition.endDate) & ", position = " & SText(argPosition.position) & _
        " WHERE position_cd = " & SText(argPosition.Code) & ";"
    
    executeSqlWithLog sql, procesureNM, tableNM, jobNM
    
End Sub

'@param: argPosition: argPosition 객체를 받아 DB에서 Delete 합니다.
Public Sub Delete(ByRef argPosition As position)

    Dim sql As String
    Dim procesureNM As String
    Dim jobNM As String
    Dim tableNM As String
    
    procesureNM = "PositionDao.Delete()"
    jobNM = "직책이력 삭제"
    tableNM = TABLE_POSITION
    
    '--//UPDATE 문
    sql = _
        "DELETE w FROM " & tableNM & " w WHERE w.position_cd = " & SText(argPosition.Code) & ";"
    
    executeSqlWithLog sql, procesureNM, tableNM, jobNM

End Sub

'@param argPosition: argPosition 객체에 설정된 시작일과 종료일이 DB에 저장된 다른 발령이력과 겹치는지 점검
Public Function CheckTimeOverlapped(ByRef argPosition As position) As Boolean
    
    Dim sql As String
    Dim tRecordSet As T_RECORD_SET
    
    sql = "SELECT * FROM " & TABLE_POSITION & " a WHERE a.lifeno = " & SText(argPosition.lifeNo) & _
            " AND IF(a.start_dt > " & SText(argPosition.startDate) & ", a.start_dt, " & SText(argPosition.startDate) & ") <= " & _
            "IF(a.end_dt < " & SText(argPosition.endDate) & ", a.end_dt, " & SText(argPosition.endDate) & ")" & _
            " AND a.position_cd <> " & SText(argPosition.Code) & ";"
    
    tRecordSet = makeListData(sql, TABLE_POSITION)
    
    CheckTimeOverlapped = (tRecordSet.CNT_RECORD > 0)
    
End Function

'--//직책목록을 DB에서 받아 반환 합니다.
Public Function GetPositionList() As Object

    Dim sql As String
    Dim tRecordSet As T_RECORD_SET
    Dim positionList As Object
    Set positionList = CreateObject("System.Collections.ArrayList")
    
    sql = "SELECT * FROM " & A_POSITION & ";"
    tRecordSet = makeListData(sql, A_POSITION)
    
    With tRecordSet
        If .CNT_RECORD > 0 Then
            Dim strPosition As Variant
            For Each strPosition In .LISTDATA
                positionList.Add strPosition
            Next
        End If
    End With
    
    Set GetPositionList = positionList

End Function
