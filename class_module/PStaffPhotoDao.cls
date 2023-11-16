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

'@param lifeNo: 생명번호를 기준으로 DB에서 검색합니다.
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


'@param pPhoto: pPhoto 객체를 받아 DB에 Insert 또는 Update 합니다.
Public Sub Save(ByRef pPhoto As PStaffPhoto)
    Dim pPhotoTemp As New PStaffPhoto
    
    Set pPhotoTemp = FindByLifeNo(pPhoto.lifeNo)
    If pPhotoTemp.lifeNo = "" Then
        Insert pPhoto
    Else
        Update pPhoto
    End If
    
End Sub

'@param: pPhoto: pPhoto객체를 받아 DB에 Insert 합니다.
Private Sub Insert(ByRef pPhoto As PStaffPhoto)
    Dim sql As String
    Dim procesureNM As String
    Dim jobNM As String
    Dim tableNM As String
    
    procesureNM = "PStaffPhotoDao.Insert()"
    jobNM = "선지자(배우자) 사진 추가"
    tableNM = TABLE_PSTAFF_PHOTO
    
    '--//INSERT 문
    sql = _
        "INSERT INTO " & tableNM & _
            " (lifeno, photo) " & _
        "VALUES(" & _
            SText(pPhoto.lifeNo) & "," & SText(pPhoto.photo) & ");"
    
    executeSqlWithLog sql, procesureNM, tableNM, jobNM
    
End Sub

'@param: pWife: pWife 객체를 받아 DB에 Update 합니다.
Private Sub Update(ByRef pPhoto As PStaffPhoto)
    Dim sql As String
    Dim procesureNM As String
    Dim jobNM As String
    Dim tableNM As String
    
    procesureNM = "PastoralWifeDao.Update()"
    jobNM = "선지자(배우자) 사진 업데이트"
    tableNM = TABLE_PSTAFF_PHOTO
    
    '--//UPDATE 문
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
    jobNM = "선지자(배우자) 사진 삭제"
    tableNM = TABLE_PSTAFF_PHOTO
    
    '--//UPDATE 문
    sql = _
        "DELETE w FROM " & tableNM & " w WHERE w.lifeno = " & SText(pPhoto.lifeNo) & ";"
    
    executeSqlWithLog sql, procesureNM, tableNM, jobNM

End Sub
