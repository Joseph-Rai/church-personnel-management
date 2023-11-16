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

'@param argId: 교회코드를 기준으로 DB에서 검색합니다.
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


'@param argChurchMap: argChurchMap 객체를 받아 DB에 Insert 또는 Update 합니다.
Public Sub Save(ByRef argChurchMap As ChurchMap)
    Dim churchMapTemp As New ChurchMap
    
    Set churchMapTemp = FindByChurchId(argChurchMap.Id)
    If churchMapTemp.Id = "" Then
        Insert argChurchMap
    Else
        Update argChurchMap
    End If
    
End Sub

'@param: argChurchMap: argChurchMap 객체를 받아 DB에 Insert 합니다.
Private Sub Insert(ByRef argChurchMap As ChurchMap)
    Dim sql As String
    Dim procesureNM As String
    Dim jobNM As String
    Dim tableNM As String
    
    procesureNM = "ChurchMapDao.Insert()"
    jobNM = "교회 지도 추가"
    tableNM = TABLE_CHURCH_MAP
    
    '--//INSERT 문
    sql = _
        "INSERT INTO " & tableNM & _
            " (sid, map) " & _
        "VALUES(" & _
            SText(argChurchMap.Id) & "," & SText(argChurchMap.map) & ");"
    
    executeSqlWithLog sql, procesureNM, tableNM, jobNM
    
End Sub

'@param: argChurchMap: argChurchMap 객체를 받아 DB에 Update 합니다.
Private Sub Update(ByRef argChurchMap As ChurchMap)
    Dim sql As String
    Dim procesureNM As String
    Dim jobNM As String
    Dim tableNM As String
    
    procesureNM = "ChurchMapDao.Update()"
    jobNM = "교회 지도 업데이트"
    tableNM = TABLE_CHURCH_MAP
    
    '--//UPDATE 문
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
    jobNM = "교회 지도 삭제"
    tableNM = TABLE_CHURCH_MAP
    
    '--//UPDATE 문
    sql = _
        "DELETE w FROM " & tableNM & " w WHERE w.sid = " & SText(argChurchMap.Id) & ";"
    
    executeSqlWithLog sql, procesureNM, tableNM, jobNM

End Sub


