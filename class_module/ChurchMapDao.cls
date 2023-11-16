VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ChurchMapDao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'@param argId: ��ȸ�ڵ带 �������� DB���� �˻��մϴ�.
Public Function FindByChurchId(ByVal argId As String) As ChurchMap
    
    Dim sql As String
    Dim tRecordSet As T_RECORD_SET
    
    sql = _
        "SELECT * FROM " & TABLE_CHURCH_MAP & " cm" & _
        "    WHERE cm.sid = " & SText(argId) & ";"
    
    tRecordSet = makeListData(sql, TABLE_CHURCH_MAP)
    
    Dim churchMapTemp As New ChurchMap
    If tRecordSet.CNT_RECORD > 0 Then
        churchMapTemp.ParseFromRecordSet tRecordSet
    Else
        Set churchMapTemp = Nothing
    End If
    
    Set FindByChurchId = churchMapTemp
    
End Function


'@param argChurchMap: argChurchMap ��ü�� �޾� DB�� Insert �Ǵ� Update �մϴ�.
Public Sub Save(ByRef argChurchMap As ChurchMap)
    Dim churchMapTemp As New ChurchMap
    
    Set churchMapTemp = FindByChurchId(argChurchMap.Id)
    If churchMapTemp.Id = "" Then
        Insert argChurchMap
    Else
        Update argChurchMap
    End If
    
End Sub

'@param: argChurchMap: argChurchMap ��ü�� �޾� DB�� Insert �մϴ�.
Private Sub Insert(ByRef argChurchMap As ChurchMap)
    Dim sql As String
    Dim procesureNM As String
    Dim jobNM As String
    Dim tableNM As String
    
    procesureNM = "ChurchMapDao.Insert()"
    jobNM = "��ȸ ���� �߰�"
    tableNM = TABLE_CHURCH_MAP
    
    '--//INSERT ��
    sql = _
        "INSERT INTO " & tableNM & _
            " (sid, map) " & _
        "VALUES(" & _
            SText(argChurchMap.Id) & "," & SText(argChurchMap.map) & ");"
    
    executeSqlWithLog sql, procesureNM, tableNM, jobNM
    
End Sub

'@param: argChurchMap: argChurchMap ��ü�� �޾� DB�� Update �մϴ�.
Private Sub Update(ByRef argChurchMap As ChurchMap)
    Dim sql As String
    Dim procesureNM As String
    Dim jobNM As String
    Dim tableNM As String
    
    procesureNM = "ChurchMapDao.Update()"
    jobNM = "��ȸ ���� ������Ʈ"
    tableNM = TABLE_CHURCH_MAP
    
    '--//UPDATE ��
    sql = _
        "UPDATE " & tableNM & " SET " & _
        "map = " & SText(argChurchMap.map) & _
        " WHERE sid = " & SText(argChurchMap.Id) & ";"
    
    executeSqlWithLog sql, procesureNM, tableNM, jobNM
    
End Sub

Public Sub Delete(ByRef argChurchMap As ChurchMap)

    Dim sql As String
    Dim procesureNM As String
    Dim jobNM As String
    Dim tableNM As String
    
    procesureNM = "ChurchMapDao.Delete()"
    jobNM = "��ȸ ���� ����"
    tableNM = TABLE_CHURCH_MAP
    
    '--//UPDATE ��
    sql = _
        "DELETE w FROM " & tableNM & " w WHERE w.sid = " & SText(argChurchMap.Id) & ";"
    
    executeSqlWithLog sql, procesureNM, tableNM, jobNM

End Sub

