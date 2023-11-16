VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PStaffPhotoDao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'@param lifeNo: ������ȣ�� �������� DB���� �˻��մϴ�.
Public Function FindByLifeNo(ByVal lifeNo As String) As PStaffPhoto
    
    Dim sql As String
    Dim tRecordSet As T_RECORD_SET
    
    sql = _
        "SELECT * FROM " & TABLE_PSTAFF_PHOTO & " pp" & _
        "    WHERE pp.lifeno = " & SText(lifeNo) & ";"
    
    tRecordSet = makeListData(sql, TABLE_PSTAFF_PHOTO)
    
    Dim pPhoto As New PStaffPhoto
    If tRecordSet.CNT_RECORD > 0 Then
        pPhoto.ParseFromRecordSet tRecordSet
    Else
        Set pPhoto = Nothing
    End If
    
    Set FindByLifeNo = pPhoto
    
End Function


'@param pPhoto: pPhoto ��ü�� �޾� DB�� Insert �Ǵ� Update �մϴ�.
Public Sub Save(ByRef pPhoto As PStaffPhoto)
    Dim pPhotoTemp As New PStaffPhoto
    
    Set pPhotoTemp = FindByLifeNo(pPhoto.lifeNo)
    If pPhotoTemp.lifeNo = "" Then
        Insert pPhoto
    Else
        Update pPhoto
    End If
    
End Sub

'@param: pPhoto: pPhoto��ü�� �޾� DB�� Insert �մϴ�.
Private Sub Insert(ByRef pPhoto As PStaffPhoto)
    Dim sql As String
    Dim procesureNM As String
    Dim jobNM As String
    Dim tableNM As String
    
    procesureNM = "PStaffPhotoDao.Insert()"
    jobNM = "������(�����) ���� �߰�"
    tableNM = TABLE_PSTAFF_PHOTO
    
    '--//INSERT ��
    sql = _
        "INSERT INTO " & tableNM & _
            " (lifeno, photo) " & _
        "VALUES(" & _
            SText(pPhoto.lifeNo) & "," & SText(pPhoto.photo) & ");"
    
    executeSqlWithLog sql, procesureNM, tableNM, jobNM
    
End Sub

'@param: pWife: pWife ��ü�� �޾� DB�� Update �մϴ�.
Private Sub Update(ByRef pPhoto As PStaffPhoto)
    Dim sql As String
    Dim procesureNM As String
    Dim jobNM As String
    Dim tableNM As String
    
    procesureNM = "PastoralWifeDao.Update()"
    jobNM = "������(�����) ���� ������Ʈ"
    tableNM = TABLE_PSTAFF_PHOTO
    
    '--//UPDATE ��
    sql = _
        "UPDATE " & tableNM & " SET " & _
        "photo = " & SText(pPhoto.photo) & _
        " WHERE lifeno = " & SText(pPhoto.lifeNo) & ";"
    
    executeSqlWithLog sql, procesureNM, tableNM, jobNM
    
End Sub

Public Sub Delete(ByRef pWife As PastoralWife)

    Dim sql As String
    Dim procesureNM As String
    Dim jobNM As String
    Dim tableNM As String
    
    procesureNM = "PastoralWifeDao.Delete()"
    jobNM = "������(�����) ���� ����"
    tableNM = TABLE_PSTAFF_PHOTO
    
    '--//UPDATE ��
    sql = _
        "DELETE w FROM " & tableNM & " w WHERE w.lifeno = " & SText(pPhoto.lifeNo) & ";"
    
    executeSqlWithLog sql, procesureNM, tableNM, jobNM

End Sub