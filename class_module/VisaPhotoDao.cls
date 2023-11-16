VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VisaPhotoDao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'@param visacode: ������ȣ�� �������� DB���� �˻��մϴ�.
Public Function FindByVisaCode(ByVal visa_cd As String) As VisaPhoto
    
    Dim sql As String
    Dim tRecordSet As T_RECORD_SET
    
    sql = _
        "SELECT * FROM " & TABLE_VISA_PHOTO & " pp" & _
        "    WHERE pp.visa_cd = " & SText(visa_cd) & ";"
    
    tRecordSet = makeListData(sql, TABLE_VISA_PHOTO)
    
    Dim vPhoto As New VisaPhoto
    If tRecordSet.CNT_RECORD > 0 Then
        vPhoto.ParseFromRecordSet tRecordSet
    Else
        Set vPhoto = Nothing
    End If
    
    Set FindByVisaCode = vPhoto
    
End Function


'@param vPhoto: vPhoto ��ü�� �޾� DB�� Insert �Ǵ� Update �մϴ�.
Public Sub Save(ByRef vPhoto As VisaPhoto)
    Dim vPhotoTemp As New VisaPhoto
    
    Set vPhotoTemp = FindByVisaCode(vPhoto.visaCode)
    If vPhotoTemp.visaCode = "" Then
        Insert vPhoto
    Else
        Update vPhoto
    End If
    
End Sub

'@param: vPhoto: vPhoto��ü�� �޾� DB�� Insert �մϴ�.
Private Sub Insert(ByRef vPhoto As VisaPhoto)
    Dim sql As String
    Dim procesureNM As String
    Dim jobNM As String
    Dim tableNM As String
    
    procesureNM = "VisaPhotoDao.Insert()"
    jobNM = "���� ���� �߰�"
    tableNM = TABLE_VISA_PHOTO
    
    '--//INSERT ��
    sql = _
        "INSERT INTO " & tableNM & _
            " (visa_cd, photo) " & _
        "VALUES(" & _
            SText(vPhoto.visaCode) & "," & SText(vPhoto.photo) & ");"
    
    executeSqlWithLog sql, procesureNM, tableNM, jobNM
    
End Sub

'@param: pWife: pWife ��ü�� �޾� DB�� Update �մϴ�.
Private Sub Update(ByRef vPhoto As VisaPhoto)
    Dim sql As String
    Dim procesureNM As String
    Dim jobNM As String
    Dim tableNM As String
    
    procesureNM = "PastoralWifeDao.Update()"
    jobNM = "���� ���� ������Ʈ"
    tableNM = TABLE_VISA_PHOTO
    
    '--//UPDATE ��
    sql = _
        "UPDATE " & tableNM & " SET " & _
        "photo = " & SText(vPhoto.photo) & _
        " WHERE visacode = " & SText(vPhoto.visaCode) & ";"
    
    executeSqlWithLog sql, procesureNM, tableNM, jobNM
    
End Sub

Public Sub Delete(ByRef pWife As PastoralWife)

    Dim sql As String
    Dim procesureNM As String
    Dim jobNM As String
    Dim tableNM As String
    
    procesureNM = "PastoralWifeDao.Delete()"
    jobNM = "���� ���� ����"
    tableNM = TABLE_VISA_PHOTO
    
    '--//UPDATE ��
    sql = _
        "DELETE w FROM " & tableNM & " w WHERE w.visacode = " & SText(vPhoto.visaCode) & ";"
    
    executeSqlWithLog sql, procesureNM, tableNM, jobNM

End Sub
