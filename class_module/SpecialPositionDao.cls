VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SpecialPositionDao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'@param argCode: 코드값을 기준으로 DB에서 검색합니다.
Public Function FindByCode(ByVal argCode As Integer) As SpecialPosition
    
    Dim sql As String
    Dim tRecordSet As T_RECORD_SET
    
    sql = _
        "SELECT * FROM " & TABLE_SPECIAL_POSITION & " t" & _
        "    WHERE t.position2_cd = " & SText(argCode) & ";"
    
    tRecordSet = makeListData(sql, TABLE_SPECIAL_POSITION)
    
    Dim objSpecialPosition As New SpecialPosition
    If tRecordSet.CNT_RECORD > 0 Then
        objSpecialPosition.ParseFromRecordSet tRecordSet
    Else
        Set objSpecialPosition = Nothing
    End If
    
    Set FindByCode = objSpecialPosition
    
End Function

'@param argLifeNo: 생명번호를 기준으로 DB에서 검색합니다. 리스트를 반환합니다.
Public Function FindByLifeNo(ByVal argLifeNo As String) As Object
    
    Dim sql As String
    Dim tRecordSet As T_RECORD_SET
    Dim specialPositionList As Object
    Set specialPositionList = CreateObject("System.Collections.ArrayList")
    
    sql = _
        "SELECT * FROM " & TABLE_SPECIAL_POSITION & " t" & _
        "    WHERE t.lifeno = " & SText(argLifeNo) & _
        "    ORDER BY t.start_dt;"
    
    tRecordSet = makeListData(sql, TABLE_SPECIAL_POSITION)
    
    Dim objSpecialPosition As SpecialPosition
    If tRecordSet.CNT_RECORD > 0 Then
        Dim i As Long
        For i = 0 To tRecordSet.CNT_RECORD - 1
            Set objSpecialPosition = New SpecialPosition
            objSpecialPosition.ParseFromRecordSet tRecordSet, i
            specialPositionList.Add objSpecialPosition
        Next
    Else
        Set specialPositionList = Nothing
    End If
    
    Set FindByLifeNo = specialPositionList
    
End Function

'@param argLifeNo: 생명번호를 기준으로 DB에서 검색합니다. 리스트를 반환합니다.
Public Function FindBySpecialPosition(ByRef argSpecialPosition As SpecialPosition) As SpecialPosition
    
'    Dim sql As String
'    Dim tRecordSet As T_RECORD_SET
'
'    sql = _
'        "SELECT * FROM " & TABLE_SPECIAL_POSITION & " t" & _
'        "    WHERE t.lifeno = " & SText(argSpecialPosition.LifeNo) & _
'        "      AND t.start_dt = " & SText(argSpecialPosition.StartDate) & _
'        "      AND t.position2 = " & SText(argSpecialPosition.SpecialPosition) & ";"
'
'    tRecordSet = makeListData(sql, TABLE_SPECIAL_POSITION)
'
'    Dim objSpecialPosition As SpecialPosition
'    If tRecordSet.CNT_RECORD > 0 Then
'        objSpecialPosition.ParseFromRecordSet tRecordSet
'    Else
'        Set objSpecialPosition = Nothing
'    End If
    
    Set FindBySpecialPosition = FindByCode(argSpecialPosition.Code)
    
End Function

'@param argSpecialPosition: argSpecialPosition 객체를 받아 DB에 Insert 또는 Update 합니다.
Public Sub Save(ByRef argSpecialPosition As SpecialPosition)
    Dim tmpSpecialPosition As New SpecialPosition
    
    Set tmpSpecialPosition = FindBySpecialPosition(argSpecialPosition)
    If tmpSpecialPosition.Code = 0 Then
        Insert argSpecialPosition
    Else
        Update argSpecialPosition
    End If
    
End Sub

'@param: argSpecialPosition: argSpecialPosition 객체를 받아 DB에 Insert 합니다.
Private Sub Insert(ByRef argSpecialPosition As SpecialPosition)
    Dim sql As String
    Dim procesureNM As String
    Dim jobNM As String
    Dim tableNM As String
    
    procesureNM = "SpecialPositionDao.Insert()"
    jobNM = "특수직책이력 추가"
    tableNM = TABLE_SPECIAL_POSITION
    
    '--//INSERT 문
    sql = _
        "INSERT INTO " & tableNM & _
            " (lifeno, start_dt, end_dt, position2) " & _
        "VALUES(" & _
            SText(argSpecialPosition.lifeNo) & "," & SText(argSpecialPosition.startDate) & "," & _
            SText(argSpecialPosition.endDate) & "," & SText(argSpecialPosition.SpecialPosition) & _
        ");"
    
    executeSqlWithLog sql, procesureNM, tableNM, jobNM
    
End Sub

'@param: argSpecialPosition: argSpecialPosition 객체를 받아 DB에 Update 합니다.
Private Sub Update(ByRef argSpecialPosition As SpecialPosition)
    Dim sql As String
    Dim procesureNM As String
    Dim jobNM As String
    Dim tableNM As String
    
    procesureNM = "SpecialPositionDao.Update()"
    jobNM = "특수직책이력 업데이트"
    tableNM = TABLE_SPECIAL_POSITION
    
    '--//UPDATE 문
    sql = _
        "UPDATE " & tableNM & " SET " & _
        "lifeno = " & SText(argSpecialPosition.lifeNo) & ", start_dt = " & SText(argSpecialPosition.startDate) & _
        ", end_dt = " & SText(argSpecialPosition.endDate) & ", position2 = " & SText(argSpecialPosition.SpecialPosition) & _
        " WHERE position2_cd = " & SText(argSpecialPosition.Code) & ";"
    
    executeSqlWithLog sql, procesureNM, tableNM, jobNM
    
End Sub

'@param: argSpecialPosition: argSpecialPosition 객체를 받아 DB에서 Delete 합니다.
Public Sub Delete(ByRef argSpecialPosition As SpecialPosition)

    Dim sql As String
    Dim procesureNM As String
    Dim jobNM As String
    Dim tableNM As String
    
    procesureNM = "SpecialPositionDao.Delete()"
    jobNM = "특수직책이력 삭제"
    tableNM = TABLE_SPECIAL_POSITION
    
    '--//UPDATE 문
    sql = _
        "DELETE w FROM " & tableNM & " w WHERE w.position2_cd = " & SText(argSpecialPosition.Code) & ";"
    
    executeSqlWithLog sql, procesureNM, tableNM, jobNM

End Sub

'@param argSpecialPosition: argSpecialPosition 객체에 설정된 시작일과 종료일이 DB에 저장된 다른 발령이력과 겹치는지 점검
Public Function CheckTimeOverlapped(ByRef argSpecialPosition As SpecialPosition) As Boolean
    
    Dim sql As String
    Dim tRecordSet As T_RECORD_SET
    
    sql = "SELECT * FROM " & TABLE_SPECIAL_POSITION & " a WHERE a.lifeno = " & SText(argSpecialPosition.lifeNo) & _
            " AND IF(a.start_dt > " & SText(argSpecialPosition.startDate) & ", a.start_dt, " & SText(argSpecialPosition.startDate) & ") <= " & _
            "IF(a.end_dt < " & SText(argSpecialPosition.endDate) & ", a.end_dt, " & SText(argSpecialPosition.endDate) & ")" & _
            " AND a.position2_cd <> " & SText(argSpecialPosition.Code) & ";"
    
    tRecordSet = makeListData(sql, TABLE_SPECIAL_POSITION)
    
    CheckTimeOverlapped = (tRecordSet.CNT_RECORD > 0)
    
End Function

'--//특수직책목록을 DB에서 받아 반환 합니다.
Public Function GetSpecialPositionList() As Object

    Dim sql As String
    Dim tRecordSet As T_RECORD_SET
    Dim specialPositionList As Object
    Set specialPositionList = CreateObject("System.Collections.ArrayList")
    
    sql = "SELECT * FROM " & A_SPECIAL_POSITION & ";"
    tRecordSet = makeListData(sql, A_SPECIAL_POSITION)
    
    With tRecordSet
        If .CNT_RECORD > 0 Then
            Dim strSpecialPosition As Variant
            For Each strSpecialPosition In .LISTDATA
                specialPositionList.Add strSpecialPosition
            Next
        End If
    End With
    
    Set GetSpecialPositionList = specialPositionList

End Function

