VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PastoralWifeDao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'@param lifeNo: 생명번호를 기준으로 DB에서 검색합니다.
Public Function FindByLifeNo(ByVal lifeNo As String) As PastoralWife
    
    Dim sql As String
    Dim tRecordSet As T_RECORD_SET
    
    sql = "SELECT * FROM " & TABLE_PASTORALWIFE & _
                    " a WHERE a.lifeno = " & SText(lifeNo) & ";"
    
    tRecordSet = makeListData(sql, TABLE_PASTORALWIFE)
    
    Dim pWife As New PastoralWife
    If tRecordSet.CNT_RECORD > 0 Then
        pWife.ParseFromRecordSet tRecordSet
    Else
        Set pWife = Nothing
    End If
    
    Set FindByLifeNo = pWife
    
End Function

'@param pStaff: pStaff 객체를 받아 생명번호를 기준으로 DB에서 검색합니다.
Public Function FindByWife(ByRef pWife As PastoralWife) As PastoralWife

    Set FindByWife = FindByLifeNo(pWife.lifeNo)

End Function

'@param lifeNo: 생명번호를 기준으로 DB에서 검색합니다.
Public Function FindByPStaff(ByVal pStaff As PastoralStaff) As PastoralWife
    
    Dim sql As String
    Dim tRecordSet As T_RECORD_SET
    
    sql = _
        "SELECT * FROM " & TABLE_PASTORALWIFE & " w WHERE w.lifeno_spouse = " & SText(pStaff.lifeNo) & ";"
    
    tRecordSet = makeListData(sql, TABLE_PASTORALWIFE)
    
    Dim pWife As New PastoralWife
    If tRecordSet.CNT_RECORD > 0 Then
        pWife.ParseFromRecordSet tRecordSet
    Else
        Set pWife = Nothing
    End If
    
    Set FindByPStaff = pWife
    
End Function

'@param pWife: pWife 객체를 받아 생명번호를 기준으로 DB에서 검색합니다.
Public Function FindByWifeAndSpouseLifeNo(ByRef pWife As PastoralWife) As PastoralWife

    Dim sql As String
    Dim tRecordSet As T_RECORD_SET
    
    sql = _
        "SELECT * FROM " & TABLE_PASTORALWIFE & _
        " a WHERE a.lifeno = " & SText(pWife.lifeNo) & _
        " AND a.lifeno_spouse <> ''"
    
    If pWife.lifeNoSpouse <> "" Then
        sql = sql & " AND a.lifeno_spouse <> " & SText(pWife.lifeNoSpouse) & ";"
    End If
    
    Set tRecordSet = makeListData(sql, TABLE_PASTORALWIFE)
    
    Dim pWifeResult As New PastoralWife
    If tRecordSet.CNT_RECORD > 0 Then
        pWifeResult.ParseFromRecordSet tRecordSet
    Else
        Set pWifeResult = Nothing
    End If
    
    Set FindByWifeAndSpouseLifeNo = pWifeResult

End Function

'@param pWife: pWife 객체를 받아 DB에 Insert 또는 Update 합니다.
Public Sub Save(ByRef pWife As PastoralWife)
    Dim pWifeTemp As New PastoralWife
    
    Set pWifeTemp = FindByWife(pWife)
    If pWifeTemp.lifeNo = "" Then
        Insert pWife
    Else
        Update pWife
    End If
    
End Sub

'@param: pWife: pWife 객체를 받아 DB에 Insert 합니다.
Private Sub Insert(ByRef pWife As PastoralWife)
    Dim sql As String
    Dim procesureNM As String
    Dim jobNM As String
    Dim tableNM As String
    
    procesureNM = "PastoralWifeDao.Insert()"
    jobNM = "사모 정보 추가"
    tableNM = TABLE_PASTORALWIFE
    
    '--//INSERT 문
    sql = _
        "INSERT INTO " & tableNM & _
            " (lifeno, nationality, name_ko, name_en, birthday, phone, " & _
            "home, family, health, other, lifeno_spouse, education, suspend, ovs_dept) " & _
        "VALUES(" & _
            SText(pWife.lifeNo) & "," & SText(pWife.Nationality) & "," & SText(pWife.nameKo) & "," & SText(pWife.NameEn) & "," & SText(pWife.Birthday) & "," & SText(pWife.Phone) & "," & _
            SText(pWife.Home) & "," & SText(pWife.Family) & "," & SText(pWife.Health) & "," & SText(pWife.Other) & "," & SText(pWife.lifeNoSpouse) & "," & SText(pWife.Education) & "," & SText(IIf(pWife.Suspend, 1, 0)) & "," & SText(pWife.OvsDept) & ");"
    
    executeSqlWithLog sql, procesureNM, tableNM, jobNM
    
End Sub

'@param: pWife: pWife 객체를 받아 DB에 Update 합니다.
Private Sub Update(ByRef pWife As PastoralWife)
    Dim sql As String
    Dim procesureNM As String
    Dim jobNM As String
    Dim tableNM As String
    
    procesureNM = "PastoralWifeDao.Update()"
    jobNM = "사모 정보 업데이트"
    tableNM = TABLE_PASTORALWIFE
    
    '--//UPDATE 문
    sql = _
        "UPDATE " & tableNM & " SET " & _
        "name_ko = " & SText(pWife.nameKo) & ",name_en = " & SText(pWife.NameEn) & ",birthday = " & IIf(pWife.Birthday = DateSerial(1900, 1, 1), "NULL", SText(pWife.Birthday)) & _
        ",phone = " & SText(pWife.Phone) & ",nationality = " & SText(pWife.Nationality) & ",home = " & SText(pWife.Home) & ",family = " & SText(pWife.Family) & _
        ",health = " & SText(pWife.Health) & ",lifeno_spouse = " & SText(pWife.lifeNoSpouse) & ",other = " & SText(pWife.Other) & ",suspend = " & SText(IIf(pWife.Suspend, 1, 0)) & _
        ",education = " & SText(pWife.Education) & ",ovs_dept = " & SText(USER_DEPT) & " WHERE lifeno = " & SText(pWife.lifeNo) & ";"
    
    executeSqlWithLog sql, procesureNM, tableNM, jobNM
    
End Sub

Public Sub Delete(ByRef pWife As PastoralWife)

    Dim sql As String
    Dim procesureNM As String
    Dim jobNM As String
    Dim tableNM As String
    
    procesureNM = "PastoralWifeDao.Delete()"
    jobNM = "사모 정보 삭제"
    tableNM = TABLE_PASTORALWIFE
    
    '--//UPDATE 문
    sql = _
        "DELETE w FROM " & tableNM & " w WHERE w.lifeno = " & SText(pWife.lifeNo) & ";"
    
    executeSqlWithLog sql, procesureNM, tableNM, jobNM

End Sub

