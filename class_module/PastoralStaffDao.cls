VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PastoralStaffDao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'@param lifeNo: 생명번호를 기준으로 DB에서 검색합니다.
Public Function FindByLifeNo(ByVal lifeNo As String) As PastoralStaff
    
    Dim sql As String
    Dim tRecordSet As T_RECORD_SET
    
    sql = "SELECT * FROM " & TABLE_PASTORALSTAFF & _
                    " a WHERE a.lifeno = " & SText(lifeNo) & ";"
    
    tRecordSet = makeListData(sql, TABLE_PASTORALSTAFF)
    
    Dim pStaff As New PastoralStaff
    If tRecordSet.CNT_RECORD > 0 Then
        pStaff.ParseFromRecordSet tRecordSet
    Else
        Set pStaff = Nothing
    End If
    
    Set FindByLifeNo = pStaff
    
End Function

'@param pStaff: pStaff 객체를 받아 생명번호를 기준으로 DB에서 검색합니다.
Public Function FindByStaff(ByRef pStaff As PastoralStaff) As PastoralStaff

    Set FindByStaff = FindByLifeNo(pStaff.lifeNo)
    
End Function

'@param pStaff: pStaff 객체를 받아 DB에 Insert 또는 Update 합니다.
Public Sub Save(ByRef pStaff As PastoralStaff)
    Dim pStaffTemp As New PastoralStaff
    
    Set pStaffTemp = FindByStaff(pStaff)
    If pStaffTemp.lifeNo = "" Then
        Insert pStaff
    Else
        Update pStaff
    End If
    
End Sub

'@param: pStaff: pStaff 객체를 받아 DB에 Insert 합니다.
Private Sub Insert(ByRef pStaff As PastoralStaff)
    Dim sql As String
    Dim procesureNM As String
    Dim jobNM As String
    Dim tableNM As String
    
    procesureNM = "PastoralStaffDao.Insert()"
    jobNM = "선지자 정보 추가"
    tableNM = TABLE_PASTORALSTAFF
    
    '--//INSERT 문
    sql = _
        "INSERT INTO " & tableNM & " (" & _
            "lifeno, name_ko, name_en, nationality, birthday, phone, " & _
            "lifeno_child1, name_ko_child1, name_en_child1, birthday_child1, phone_child1, " & _
            "lifeno_child2, name_ko_child2, name_en_child2, birthday_child2, phone_child2, " & _
            "lifeno_child3, name_ko_child3, name_en_child3, birthday_child3, phone_child3, " & _
            "home, family, health, other, baptism, ordination_prayer, appo_ovs, wedding_dt, theological_order, education, salary, suspend, ovs_dept)" & _
        " VALUES (" & _
            SText(pStaff.lifeNo) & "," & SText(pStaff.nameKo) & "," & SText(pStaff.NameEn) & "," & SText(pStaff.Nationality) & "," & SText(pStaff.Birthday) & "," & SText(pStaff.Phone) & "," & _
            SText(pStaff.LifeNoChild1) & "," & SText(pStaff.NameKoChild1) & "," & SText(pStaff.NameEnChild1) & "," & IIf(pStaff.BirthdayChild1 = "1900-01-01", "NULL", SText(pStaff.BirthdayChild1)) & "," & SText(pStaff.PhoneChild1) & "," & _
            SText(pStaff.LifeNoChild2) & "," & SText(pStaff.NameKoChild2) & "," & SText(pStaff.NameEnChild2) & "," & IIf(pStaff.BirthdayChild2 = "1900-01-01", "NULL", SText(pStaff.BirthdayChild2)) & "," & SText(pStaff.PhoneChild2) & "," & _
            SText(pStaff.LifeNoChild3) & "," & SText(pStaff.NameKoChild3) & "," & SText(pStaff.NameEnChild3) & "," & IIf(pStaff.BirthdayChild3 = "1900-01-01", "NULL", SText(pStaff.BirthdayChild3)) & "," & SText(pStaff.PhoneChild3) & "," & _
            SText(pStaff.Home) & "," & SText(pStaff.Family) & "," & SText(pStaff.Health) & "," & SText(pStaff.Other) & "," & _
            SText(pStaff.Baptism) & "," & IIf(pStaff.OrdinationPrayer = "1900-01-01", "NULL", SText(pStaff.OrdinationPrayer)) & "," & IIf(pStaff.AppoOvs = "1900-01-01", "NULL", SText(pStaff.AppoOvs)) & "," & _
            IIf(pStaff.WeddingDt = "1900-01-01", "NULL", SText(pStaff.WeddingDt)) & "," & SText(pStaff.TheologicalOrder) & "," & SText(pStaff.Education) & "," & SText(pStaff.Salary) & "," & _
            SText(IIf(pStaff.Suspend, 1, 0)) & "," & SText(pStaff.OvsDept) & _
        ");"
    
    executeSqlWithLog sql, procesureNM, tableNM, jobNM
    
End Sub

'@param: pStaff: pStaff 객체를 받아 DB에 Update 합니다.
Private Sub Update(ByRef pStaff As PastoralStaff)
    Dim sql As String
    Dim procesureNM As String
    Dim jobNM As String
    Dim tableNM As String
    
    procesureNM = "PastoralStaffDao.Update()"
    jobNM = "선지자 정보 업데이트"
    tableNM = TABLE_PASTORALSTAFF
    
    '--//UPDATE 문
    sql = _
        "UPDATE " & tableNM & " SET " & _
            "name_ko = " & SText(pStaff.nameKo) & ",name_en = " & SText(pStaff.NameEn) & ",nationality = " & SText(pStaff.Nationality) & ",birthday = " & IIf(pStaff.Birthday = DateSerial(1900, 1, 1), "NULL", SText(pStaff.Birthday)) & ",phone = " & SText(pStaff.Phone) & _
            ",home = " & SText(pStaff.Home) & ",family = " & SText(pStaff.Family) & ",health = " & SText(pStaff.Health) & ",other = " & SText(pStaff.Other) & _
            ",lifeno_child1 = " & SText(pStaff.LifeNoChild1) & ",name_ko_child1 = " & SText(pStaff.NameKoChild1) & ",name_en_child1 = " & SText(pStaff.NameEnChild1) & ",birthday_child1 = " & IIf(pStaff.BirthdayChild1 = DateSerial(1900, 1, 1), "NULL", SText(pStaff.BirthdayChild1)) & ",phone_child1 = " & SText(pStaff.PhoneChild1) & _
            ",lifeno_child2 = " & SText(pStaff.LifeNoChild2) & ",name_ko_child2 = " & SText(pStaff.NameKoChild2) & ",name_en_child2 = " & SText(pStaff.NameEnChild2) & ",birthday_child2 = " & IIf(pStaff.BirthdayChild2 = DateSerial(1900, 1, 1), "NULL", SText(pStaff.BirthdayChild2)) & ",phone_child2 = " & SText(pStaff.PhoneChild2) & _
            ",lifeno_child3 = " & SText(pStaff.LifeNoChild3) & ",name_ko_child3 = " & SText(pStaff.NameKoChild3) & ",name_en_child3 = " & SText(pStaff.NameEnChild3) & ",birthday_child3 = " & IIf(pStaff.BirthdayChild3 = DateSerial(1900, 1, 1), "NULL", SText(pStaff.BirthdayChild3)) & ",phone_child3 = " & SText(pStaff.PhoneChild3) & _
            ",baptism = " & SText(pStaff.Baptism) & ",appo_ovs = " & IIf(pStaff.AppoOvs = DateSerial(1900, 1, 1), "NULL", SText(pStaff.AppoOvs)) & ",wedding_dt = " & IIf(pStaff.WeddingDt = DateSerial(1900, 1, 1), "NULL", SText(pStaff.WeddingDt)) & _
            ",education = " & SText(pStaff.Education) & ",salary = " & SText(pStaff.Salary) & ",ordination_prayer = " & IIf(pStaff.OrdinationPrayer = DateSerial(1900, 1, 1), "NULL", SText(pStaff.OrdinationPrayer)) & ",theological_order = " & IIf(pStaff.TheologicalOrder = 0, "NULL", SText(pStaff.TheologicalOrder)) & _
            ",suspend = " & SText(IIf(pStaff.Suspend, 1, 0)) & ",ovs_dept = " & SText(pStaff.OvsDept) & _
            " WHERE lifeno = " & SText(pStaff.lifeNo) & ";"
    
    executeSqlWithLog sql, procesureNM, tableNM, jobNM
    
End Sub

Public Sub Delete(ByRef pStaff As PastoralStaff)

    Dim sql As String
    Dim procesureNM As String
    Dim jobNM As String
    Dim tableNM As String
    
    procesureNM = "PastoralStaffDao.Delete()"
    jobNM = "선지자 정보 삭제"
    tableNM = TABLE_PASTORALSTAFF
    
    '--//DELETE 문
    sql = _
        "DELETE p FROM " & tableNM & " p WHERE p.lifeno = " & SText(pStaff.lifeNo) & ";"
    
    executeSqlWithLog sql, procesureNM, tableNM, jobNM

End Sub

Public Sub DeleteChild1(ByRef pStaff As PastoralStaff)

    pStaff.LifeNoChild1 = ""
    pStaff.NameKoChild1 = ""
    pStaff.NameEnChild1 = ""
    pStaff.BirthdayChild1 = DateSerial(1900, 1, 1)
    pStaff.PhoneChild1 = ""
    
    Update pStaff

End Sub

Public Sub DeleteChild2(ByRef pStaff As PastoralStaff)

    pStaff.LifeNoChild2 = ""
    pStaff.NameKoChild2 = ""
    pStaff.NameEnChild2 = ""
    pStaff.BirthdayChild2 = DateSerial(1900, 1, 1)
    pStaff.PhoneChild2 = ""
    
    Update pStaff

End Sub

Public Sub DeleteChild3(ByRef pStaff As PastoralStaff)

    pStaff.LifeNoChild3 = ""
    pStaff.NameKoChild3 = ""
    pStaff.NameEnChild3 = ""
    pStaff.BirthdayChild3 = DateSerial(1900, 1, 1)
    pStaff.PhoneChild3 = ""
    
    Update pStaff

End Sub
