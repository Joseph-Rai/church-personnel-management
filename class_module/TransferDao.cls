VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TransferDao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function GetLastTransfer(ByVal argLifeNo As String) As Transfer

    Dim transferList As Object
    Set transferList = FindByLifeNo(argLifeNo)
    
    Dim tmpTransfer As New Transfer
    Dim maxEndDate As Date
    For Each tmpTransfer In transferList
        If maxEndDate < tmpTransfer.endDate Then
            maxEndDate = tmpTransfer.endDate
            Set GetLastTransfer = tmpTransfer
        End If
    Next

End Function

'@param argCode: �ڵ尪�� �������� DB���� �˻��մϴ�.
Public Function FindByCode(ByVal argCode As Integer) As Transfer
    
    Dim sql As String
    Dim tRecordSet As T_RECORD_SET
    
    sql = _
        "SELECT t.*, c.church_nm FROM " & TABLE_TRANSFER & " t" & _
        "    LEFT JOIN " & TABLE_CHURCH & " c" & _
        "        ON t.church_sid = c.church_sid" & _
        "    WHERE t.transfer_cd = " & SText(argCode) & ";"
    
    tRecordSet = makeListData(sql, TABLE_TRANSFER)
    
    Dim objTrans As New Transfer
    If tRecordSet.CNT_RECORD > 0 Then
        objTrans.ParseFromRecordSet tRecordSet
    Else
        Set objTrans = Nothing
    End If
    
    Set FindByCode = objTrans
    
End Function

'@param argLifeNo: ������ȣ��
'@param argDate : ��¥�� �������� �߷��̷� ��ü 1���� ��ȯ�մϴ�.
Public Function FindByLifeNoAndDate(ByVal argLifeNo As String, ByVal argDate As Date) As Transfer
    
    Dim sql As String
    Dim tRecordSet As T_RECORD_SET
    Dim transferList As Object
    Set transferList = CreateObject("System.Collections.ArrayList")
    
    sql = _
        "SELECT t.*, c.church_nm FROM " & TABLE_TRANSFER & " t" & _
        "    LEFT JOIN " & TABLE_CHURCH & " c" & _
        "        ON t.church_sid = c.church_sid" & _
        "    WHERE t.lifeno = " & SText(argLifeNo) & _
        "       AND " & SText(argDate) & " BETWEEN t.start_dt and t.end_dt " & _
        "    ORDER BY t.start_dt;"
    
    tRecordSet = makeListData(sql, TABLE_TRANSFER)
    
    Dim objTrans As Transfer
    If tRecordSet.CNT_RECORD > 0 Then
        Dim i As Long
        For i = 0 To tRecordSet.CNT_RECORD - 1
            Set objTrans = New Transfer
            objTrans.ParseFromRecordSet tRecordSet, i
            transferList.Add objTrans
        Next
    Else
        Set transferList = Nothing
    End If
    
    Set FindByLifeNo = transferList
    
End Function

'@param argLifeNo: ������ȣ�� �������� DB���� �˻��մϴ�. ����Ʈ�� ��ȯ�մϴ�.
Public Function FindByLifeNo(ByVal argLifeNo As String) As Object
    
    Dim sql As String
    Dim tRecordSet As T_RECORD_SET
    Dim transferList As Object
    Set transferList = CreateObject("System.Collections.ArrayList")
    
    sql = _
        "SELECT t.*, c.church_nm FROM " & TABLE_TRANSFER & " t" & _
        "    LEFT JOIN " & TABLE_CHURCH & " c" & _
        "        ON t.church_sid = c.church_sid" & _
        "    WHERE t.lifeno = " & SText(argLifeNo) & _
        "    ORDER BY t.start_dt;"
    
    tRecordSet = makeListData(sql, TABLE_TRANSFER)
    
    Dim objTrans As Transfer
    If tRecordSet.CNT_RECORD > 0 Then
        Dim i As Long
        For i = 0 To tRecordSet.CNT_RECORD - 1
            Set objTrans = New Transfer
            objTrans.ParseFromRecordSet tRecordSet, i
            transferList.Add objTrans
        Next
    Else
        Set transferList = Nothing
    End If
    
    Set FindByLifeNo = transferList
    
End Function

'@param argLifeNo: ������ȣ�� �������� DB���� �˻��մϴ�. ����Ʈ�� ��ȯ�մϴ�.
Public Function FindByTrans(ByRef argTrans As Transfer) As Transfer
    
    Set FindByTrans = FindByCode(argTrans.Code)
    
End Function

'@param argTrans: argTrans ��ü�� �޾� DB�� Insert �Ǵ� Update �մϴ�.
Public Sub Save(ByRef argTrans As Transfer)
    Dim tmpTrans As New Transfer
    
    Set tmpTrans = FindByTrans(argTrans)
    If tmpTrans.Code = 0 Then
        Insert argTrans
    Else
        Update argTrans
    End If
    
End Sub

'@param: argTrans: argTrans ��ü�� �޾� DB�� Insert �մϴ�.
Private Sub Insert(ByRef argTrans As Transfer)
    Dim sql As String
    Dim procesureNM As String
    Dim jobNM As String
    Dim tableNM As String
    
    procesureNM = "TransferDao.Insert()"
    jobNM = "�߷��̷� �߰�"
    tableNM = TABLE_TRANSFER
    
    '--//INSERT ��
    sql = _
        "INSERT INTO " & tableNM & _
            " (lifeno, start_dt, end_dt, church_sid) " & _
        "VALUES(" & _
            SText(argTrans.lifeNo) & "," & SText(argTrans.startDate) & "," & _
            SText(argTrans.endDate) & "," & SText(argTrans.ChurchID) & _
        ");"
    
    executeSqlWithLog sql, procesureNM, tableNM, jobNM
    
End Sub

'@param: argTrans: argTrans ��ü�� �޾� DB�� Update �մϴ�.
Private Sub Update(ByRef argTrans As Transfer)
    Dim sql As String
    Dim procesureNM As String
    Dim jobNM As String
    Dim tableNM As String
    
    procesureNM = "TransferDao.Update()"
    jobNM = "�߷��̷� ������Ʈ"
    tableNM = TABLE_TRANSFER
    
    '--//UPDATE ��
    sql = _
        "UPDATE " & tableNM & " SET " & _
        "lifeno = " & SText(argTrans.lifeNo) & ", start_dt = " & SText(argTrans.startDate) & _
        ", end_dt = " & SText(argTrans.endDate) & ", church_sid = " & SText(argTrans.ChurchID) & _
        " WHERE transfer_cd = " & SText(argTrans.Code) & ";"
    
    executeSqlWithLog sql, procesureNM, tableNM, jobNM
    
End Sub

'@param: argTrans: argTrans ��ü�� �޾� DB���� Delete �մϴ�.
Public Sub Delete(ByRef argTrans As Transfer)

    Dim sql As String
    Dim procesureNM As String
    Dim jobNM As String
    Dim tableNM As String
    
    procesureNM = "TransferDao.Delete()"
    jobNM = "�߷��̷� ����"
    tableNM = TABLE_TRANSFER
    
    '--//UPDATE ��
    sql = _
        "DELETE w FROM " & tableNM & " w WHERE w.transfer_cd = " & SText(argTrans.Code) & ";"
    
    executeSqlWithLog sql, procesureNM, tableNM, jobNM

End Sub

'@param argTrans: argTrans ��ü�� ������ �����ϰ� �������� DB�� ����� �ٸ� �߷��̷°� ��ġ���� ����
Public Function CheckTimeOverlapped(ByRef argTrans As Transfer) As Boolean
    
    Dim sql As String
    Dim tRecordSet As T_RECORD_SET
    
    sql = "SELECT * FROM " & TABLE_TRANSFER & " a WHERE a.lifeno = " & SText(argTrans.lifeNo) & _
            " AND IF(a.start_dt > " & SText(argTrans.startDate) & ", a.start_dt, " & SText(argTrans.startDate) & ") <= " & _
            "IF(a.end_dt < " & SText(argTrans.endDate) & ", a.end_dt, " & SText(argTrans.endDate) & ")" & _
            " AND a.transfer_cd <> " & SText(argTrans.Code) & ";"
    
    tRecordSet = makeListData(sql, TABLE_TRANSFER)
    
    CheckTimeOverlapped = (tRecordSet.CNT_RECORD > 0)
    
End Function