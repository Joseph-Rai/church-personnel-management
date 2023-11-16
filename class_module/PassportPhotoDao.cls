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

'@param lifeNo: 생명번호를 기준으로 DB에서 검색합니다.
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


'@param ppPhoto: ppPhoto 객체를 받아 DB에 Insert 또는 Update 합니다.
Public Sub Save(ByRef ppPhoto As PassportPhoto)
    Dim ppPhotoTemp As New PassportPhoto
    
    Set ppPhotoTemp = FindByLifeNo(ppPhoto.lifeNo)
    If ppPhotoTemp.lifeNo = "" Then
        Insert ppPhoto
    Else
        Update ppPhoto
    End If
    
End Sub

'@param: ppPhoto: ppPhoto객체를 받아 DB에 Insert 합니다.
Private Sub Insert(ByRef ppPhoto As PassportPhoto)
    Dim sql As String
    Dim procesureNM As String
    Dim jobNM As String
    Dim tableNM As String
    
    procesureNM = "PassportPhotoDao.Insert()"
    jobNM = "여권 사진 추가"
    tableNM = TABLE_PASSPORT_PHOTO
    
    '--//INSERT 문
    sql = _
        "INSERT INTO " & tableNM & _
            " (lifeno, photo) " & _
        "VALUES(" & _
            SText(ppPhoto.lifeNo) & "," & SText(ppPhoto.photo) & ");"
    
    executeSqlWithLog sql, procesureNM, tableNM, jobNM
    
End Sub

'@param: pWife: pWife 객체를 받아 DB에 Update 합니다.
Private Sub Update(ByRef ppPhoto As PassportPhoto)
    Dim sql As String
    Dim procesureNM As String
    Dim jobNM As String
    Dim tableNM As String
    
    procesureNM = "PastoralWifeDao.Update()"
    jobNM = "여권 사진 업데이트"
    tableNM = TABLE_PASSPORT_PHOTO
    
    '--//UPDATE 문
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
    jobNM = "여권 사진 삭제"
    tableNM = TABLE_PASSPORT_PHOTO
    
    '--//UPDATE 문
    sql = _
        "DELETE w FROM " & tableNM & " w WHERE w.lifeno = " & SText(ppPhoto.lifeNo) & ";"
    
    executeSqlWithLog sql, procesureNM, tableNM, jobNM

End Sub

