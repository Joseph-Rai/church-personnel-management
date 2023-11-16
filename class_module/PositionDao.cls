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

'@param argCode: �ڵ尪�� �������� DB���� �˻��մϴ�.
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

'@param argLifeNo: ������ȣ�� �������� ���� ��å�� DB���� �˻��մϴ�.
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

'@param argLifeNo: ������ȣ�� �������� DB���� �˻��մϴ�. ����Ʈ�� ��ȯ�մϴ�.
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

'@param argLifeNo: ������ȣ�� �������� DB���� �˻��մϴ�. ����Ʈ�� ��ȯ�մϴ�.
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

'@param argPosition: argPosition ��ü�� �޾� DB�� Insert �Ǵ� Update �մϴ�.
Public Sub Save(ByRef argPosition As position)
    Dim tmpPosition As New position
    
    Set tmpPosition = FindByPosition(argPosition)
    If tmpPosition.Code = 0 Then
        Insert argPosition
    Else
        Update argPosition
    End If
    
End Sub

'@param: argPosition: argPosition ��ü�� �޾� DB�� Insert �մϴ�.
Private Sub Insert(ByRef argPosition As position)
    Dim sql As String
    Dim procesureNM As String
    Dim jobNM As String
    Dim tableNM As String
    
    procesureNM = "PositionDao.Insert()"
    jobNM = "��å�̷� �߰�"
    tableNM = TABLE_POSITION
    
    '--//INSERT ��
    sql = _
        "INSERT INTO " & tableNM & _
            " (lifeno, start_dt, end_dt, position) " & _
        "VALUES(" & _
            SText(argPosition.lifeNo) & "," & SText(argPosition.startDate) & "," & _
            SText(argPosition.endDate) & "," & SText(argPosition.position) & _
        ");"
    
    executeSqlWithLog sql, procesureNM, tableNM, jobNM
    
End Sub

'@param: argPosition: argPosition ��ü�� �޾� DB�� Update �մϴ�.
Private Sub Update(ByRef argPosition As position)
    Dim sql As String
    Dim procesureNM As String
    Dim jobNM As String
    Dim tableNM As String
    
    procesureNM = "PositionDao.Update()"
    jobNM = "��å�̷� ������Ʈ"
    tableNM = TABLE_POSITION
    
    '--//UPDATE ��
    sql = _
        "UPDATE " & tableNM & " SET " & _
        "lifeno = " & SText(argPosition.lifeNo) & ", start_dt = " & SText(argPosition.startDate) & _
        ", end_dt = " & SText(argPosition.endDate) & ", position = " & SText(argPosition.position) & _
        " WHERE position_cd = " & SText(argPosition.Code) & ";"
    
    executeSqlWithLog sql, procesureNM, tableNM, jobNM
    
End Sub

'@param: argPosition: argPosition ��ü�� �޾� DB���� Delete �մϴ�.
Public Sub Delete(ByRef argPosition As position)

    Dim sql As String
    Dim procesureNM As String
    Dim jobNM As String
    Dim tableNM As String
    
    procesureNM = "PositionDao.Delete()"
    jobNM = "��å�̷� ����"
    tableNM = TABLE_POSITION
    
    '--//UPDATE ��
    sql = _
        "DELETE w FROM " & tableNM & " w WHERE w.position_cd = " & SText(argPosition.Code) & ";"
    
    executeSqlWithLog sql, procesureNM, tableNM, jobNM

End Sub

'@param argPosition: argPosition ��ü�� ������ �����ϰ� �������� DB�� ����� �ٸ� �߷��̷°� ��ġ���� ����
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

'--//��å����� DB���� �޾� ��ȯ �մϴ�.
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