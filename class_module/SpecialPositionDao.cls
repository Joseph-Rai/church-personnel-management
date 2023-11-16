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

'@param argCode: �ڵ尪�� �������� DB���� �˻��մϴ�.
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

'@param argLifeNo: ������ȣ�� �������� DB���� �˻��մϴ�. ����Ʈ�� ��ȯ�մϴ�.
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

'@param argLifeNo: ������ȣ�� �������� DB���� �˻��մϴ�. ����Ʈ�� ��ȯ�մϴ�.
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

'@param argSpecialPosition: argSpecialPosition ��ü�� �޾� DB�� Insert �Ǵ� Update �մϴ�.
Public Sub Save(ByRef argSpecialPosition As SpecialPosition)
    Dim tmpSpecialPosition As New SpecialPosition
    
    Set tmpSpecialPosition = FindBySpecialPosition(argSpecialPosition)
    If tmpSpecialPosition.Code = 0 Then
        Insert argSpecialPosition
    Else
        Update argSpecialPosition
    End If
    
End Sub

'@param: argSpecialPosition: argSpecialPosition ��ü�� �޾� DB�� Insert �մϴ�.
Private Sub Insert(ByRef argSpecialPosition As SpecialPosition)
    Dim sql As String
    Dim procesureNM As String
    Dim jobNM As String
    Dim tableNM As String
    
    procesureNM = "SpecialPositionDao.Insert()"
    jobNM = "Ư����å�̷� �߰�"
    tableNM = TABLE_SPECIAL_POSITION
    
    '--//INSERT ��
    sql = _
        "INSERT INTO " & tableNM & _
            " (lifeno, start_dt, end_dt, position2) " & _
        "VALUES(" & _
            SText(argSpecialPosition.lifeNo) & "," & SText(argSpecialPosition.startDate) & "," & _
            SText(argSpecialPosition.endDate) & "," & SText(argSpecialPosition.SpecialPosition) & _
        ");"
    
    executeSqlWithLog sql, procesureNM, tableNM, jobNM
    
End Sub

'@param: argSpecialPosition: argSpecialPosition ��ü�� �޾� DB�� Update �մϴ�.
Private Sub Update(ByRef argSpecialPosition As SpecialPosition)
    Dim sql As String
    Dim procesureNM As String
    Dim jobNM As String
    Dim tableNM As String
    
    procesureNM = "SpecialPositionDao.Update()"
    jobNM = "Ư����å�̷� ������Ʈ"
    tableNM = TABLE_SPECIAL_POSITION
    
    '--//UPDATE ��
    sql = _
        "UPDATE " & tableNM & " SET " & _
        "lifeno = " & SText(argSpecialPosition.lifeNo) & ", start_dt = " & SText(argSpecialPosition.startDate) & _
        ", end_dt = " & SText(argSpecialPosition.endDate) & ", position2 = " & SText(argSpecialPosition.SpecialPosition) & _
        " WHERE position2_cd = " & SText(argSpecialPosition.Code) & ";"
    
    executeSqlWithLog sql, procesureNM, tableNM, jobNM
    
End Sub

'@param: argSpecialPosition: argSpecialPosition ��ü�� �޾� DB���� Delete �մϴ�.
Public Sub Delete(ByRef argSpecialPosition As SpecialPosition)

    Dim sql As String
    Dim procesureNM As String
    Dim jobNM As String
    Dim tableNM As String
    
    procesureNM = "SpecialPositionDao.Delete()"
    jobNM = "Ư����å�̷� ����"
    tableNM = TABLE_SPECIAL_POSITION
    
    '--//UPDATE ��
    sql = _
        "DELETE w FROM " & tableNM & " w WHERE w.position2_cd = " & SText(argSpecialPosition.Code) & ";"
    
    executeSqlWithLog sql, procesureNM, tableNM, jobNM

End Sub

'@param argSpecialPosition: argSpecialPosition ��ü�� ������ �����ϰ� �������� DB�� ����� �ٸ� �߷��̷°� ��ġ���� ����
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

'--//Ư����å����� DB���� �޾� ��ȯ �մϴ�.
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
