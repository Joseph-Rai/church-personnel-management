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

'@param visacode: 생명번호를 기준으로 DB에서 검색합니다.
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


'@param vPhoto: vPhoto 객체를 받아 DB에 Insert 또는 Update 합니다.
Public Sub Save(ByRef vPhoto As VisaPhoto)
    Dim vPhotoTemp As New VisaPhoto
    
    Set vPhotoTemp = FindByVisaCode(vPhoto.visaCode)
    If vPhotoTemp.visaCode = "" Then
        Insert vPhoto
    Else
        Update vPhoto
    End If
    
End Sub

'@param: vPhoto: vPhoto객체를 받아 DB에 Insert 합니다.
Private Sub Insert(ByRef vPhoto As VisaPhoto)
    Dim sql As String
    Dim procesureNM As String
    Dim jobNM As String
    Dim tableNM As String
    
    procesureNM = "VisaPhotoDao.Insert()"
    jobNM = "사증 사진 추가"
    tableNM = TABLE_VISA_PHOTO
    
    '--//INSERT 문
    sql = _
        "INSERT INTO " & tableNM & _
            " (visa_cd, photo) " & _
        "VALUES(" & _
            SText(vPhoto.visaCode) & "," & SText(vPhoto.photo) & ");"
    
    executeSqlWithLog sql, procesureNM, tableNM, jobNM
    
End Sub

'@param: pWife: pWife 객체를 받아 DB에 Update 합니다.
Private Sub Update(ByRef vPhoto As VisaPhoto)
    Dim sql As String
    Dim procesureNM As String
    Dim jobNM As String
    Dim tableNM As String
    
    procesureNM = "PastoralWifeDao.Update()"
    jobNM = "사증 사진 업데이트"
    tableNM = TABLE_VISA_PHOTO
    
    '--//UPDATE 문
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
    jobNM = "사증 사진 삭제"
    tableNM = TABLE_VISA_PHOTO
    
    '--//UPDATE 문
    sql = _
        "DELETE w FROM " & tableNM & " w WHERE w.visacode = " & SText(vPhoto.visaCode) & ";"
    
    executeSqlWithLog sql, procesureNM, tableNM, jobNM

End Sub

