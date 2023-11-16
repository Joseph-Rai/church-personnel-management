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

'@param argCode: 코드값을 기준으로 DB에서 검색합니다.
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

'@param argLifeNo: 생명번호와
'@param argDate : 날짜를 기준으로 발령이력 객체 1개를 반환합니다.
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

'@param argLifeNo: 생명번호를 기준으로 DB에서 검색합니다. 리스트를 반환합니다.
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

'@param argLifeNo: 생명번호를 기준으로 DB에서 검색합니다. 리스트를 반환합니다.
Public Function FindByTrans(ByRef argTrans As Transfer) As Transfer
    
    Set FindByTrans = FindByCode(argTrans.Code)
    
End Function

'@param argTrans: argTrans 객체를 받아 DB에 Insert 또는 Update 합니다.
Public Sub Save(ByRef argTrans As Transfer)
    Dim tmpTrans As New Transfer
    
    Set tmpTrans = FindByTrans(argTrans)
    If tmpTrans.Code = 0 Then
        Insert argTrans
    Else
        Update argTrans
    End If
    
End Sub

'@param: argTrans: argTrans 객체를 받아 DB에 Insert 합니다.
Private Sub Insert(ByRef argTrans As Transfer)
    Dim sql As String
    Dim procesureNM As String
    Dim jobNM As String
    Dim tableNM As String
    
    procesureNM = "TransferDao.Insert()"
    jobNM = "발령이력 추가"
    tableNM = TABLE_TRANSFER
    
    '--//INSERT 문
    sql = _
        "INSERT INTO " & tableNM & _
            " (lifeno, start_dt, end_dt, church_sid) " & _
        "VALUES(" & _
            SText(argTrans.lifeNo) & "," & SText(argTrans.startDate) & "," & _
            SText(argTrans.endDate) & "," & SText(argTrans.ChurchID) & _
        ");"
    
    executeSqlWithLog sql, procesureNM, tableNM, jobNM
    
End Sub

'@param: argTrans: argTrans 객체를 받아 DB에 Update 합니다.
Private Sub Update(ByRef argTrans As Transfer)
    Dim sql As String
    Dim procesureNM As String
    Dim jobNM As String
    Dim tableNM As String
    
    procesureNM = "TransferDao.Update()"
    jobNM = "발령이력 업데이트"
    tableNM = TABLE_TRANSFER
    
    '--//UPDATE 문
    sql = _
        "UPDATE " & tableNM & " SET " & _
        "lifeno = " & SText(argTrans.lifeNo) & ", start_dt = " & SText(argTrans.startDate) & _
        ", end_dt = " & SText(argTrans.endDate) & ", church_sid = " & SText(argTrans.ChurchID) & _
        " WHERE transfer_cd = " & SText(argTrans.Code) & ";"
    
    executeSqlWithLog sql, procesureNM, tableNM, jobNM
    
End Sub

'@param: argTrans: argTrans 객체를 받아 DB에서 Delete 합니다.
Public Sub Delete(ByRef argTrans As Transfer)

    Dim sql As String
    Dim procesureNM As String
    Dim jobNM As String
    Dim tableNM As String
    
    procesureNM = "TransferDao.Delete()"
    jobNM = "발령이력 삭제"
    tableNM = TABLE_TRANSFER
    
    '--//UPDATE 문
    sql = _
        "DELETE w FROM " & tableNM & " w WHERE w.transfer_cd = " & SText(argTrans.Code) & ";"
    
    executeSqlWithLog sql, procesureNM, tableNM, jobNM

End Sub

'@param argTrans: argTrans 객체에 설정된 시작일과 종료일이 DB에 저장된 다른 발령이력과 겹치는지 점검
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
