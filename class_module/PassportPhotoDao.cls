VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PassportPhotoDao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'@param lifeNo: ������ȣ�� �������� DB���� �˻��մϴ�.
Public Function FindByLifeNo(ByVal lifeNo As String) As PassportPhoto
    
    Dim sql As String
    Dim tRecordSet As T_RECORD_SET
    
    sql = _
        "SELECT * FROM " & TABLE_PASSPORT_PHOTO & " pp" & _
        "    WHERE pp.lifeno = " & SText(lifeNo) & ";"
    
    tRecordSet = makeListData(sql, TABLE_PASSPORT_PHOTO)
    
    Dim ppPhoto As New PassportPhoto
    If tRecordSet.CNT_RECORD > 0 Then
        ppPhoto.ParseFromRecordSet tRecordSet
    Else
        Set ppPhoto = Nothing
    End If
    
    Set FindByLifeNo = ppPhoto
    
End Function


'@param ppPhoto: ppPhoto ��ü�� �޾� DB�� Insert �Ǵ� Update �մϴ�.
Public Sub Save(ByRef ppPhoto As PassportPhoto)
    Dim ppPhotoTemp As New PassportPhoto
    
    Set ppPhotoTemp = FindByLifeNo(ppPhoto.lifeNo)
    If ppPhotoTemp.lifeNo = "" Then
        Insert ppPhoto
    Else
        Update ppPhoto
    End If
    
End Sub

'@param: ppPhoto: ppPhoto��ü�� �޾� DB�� Insert �մϴ�.
Private Sub Insert(ByRef ppPhoto As PassportPhoto)
    Dim sql As String
    Dim procesureNM As String
    Dim jobNM As String
    Dim tableNM As String
    
    procesureNM = "PassportPhotoDao.Insert()"
    jobNM = "���� ���� �߰�"
    tableNM = TABLE_PASSPORT_PHOTO
    
    '--//INSERT ��
    sql = _
        "INSERT INTO " & tableNM & _
            " (lifeno, photo) " & _
        "VALUES(" & _
            SText(ppPhoto.lifeNo) & "," & SText(ppPhoto.photo) & ");"
    
    executeSqlWithLog sql, procesureNM, tableNM, jobNM
    
End Sub

'@param: pWife: pWife ��ü�� �޾� DB�� Update �մϴ�.
Private Sub Update(ByRef ppPhoto As PassportPhoto)
    Dim sql As String
    Dim procesureNM As String
    Dim jobNM As String
    Dim tableNM As String
    
    procesureNM = "PastoralWifeDao.Update()"
    jobNM = "���� ���� ������Ʈ"
    tableNM = TABLE_PASSPORT_PHOTO
    
    '--//UPDATE ��
    sql = _
        "UPDATE " & tableNM & " SET " & _
        "photo = " & SText(ppPhoto.photo) & _
        " WHERE lifeno = " & SText(ppPhoto.lifeNo) & ";"
    
    executeSqlWithLog sql, procesureNM, tableNM, jobNM
    
End Sub

Public Sub Delete(ByRef pWife As PastoralWife)

    Dim sql As String
    Dim procesureNM As String
    Dim jobNM As String
    Dim tableNM As String
    
    procesureNM = "PastoralWifeDao.Delete()"
    jobNM = "���� ���� ����"
    tableNM = TABLE_PASSPORT_PHOTO
    
    '--//UPDATE ��
    sql = _
        "DELETE w FROM " & tableNM & " w WHERE w.lifeno = " & SText(ppPhoto.lifeNo) & ";"
    
    executeSqlWithLog sql, procesureNM, tableNM, jobNM

End Sub
