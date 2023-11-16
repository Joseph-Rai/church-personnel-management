VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TitleDao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'@param argCode: 코드값을 기준으로 DB에서 검색합니다.
Public Function FindByCode(ByVal argCode As Integer) As title
    
    Dim sql As String
    Dim tRecordSet As T_RECORD_SET
    
    sql = _
        "SELECT * FROM " & TABLE_TITLE & " t" & _
        "    WHERE t.title_cd = " & SText(argCode) & ";"
    
    tRecordSet = makeListData(sql, TABLE_TITLE)
    
    Dim objTitle As New title
    If tRecordSet.CNT_RECORD > 0 Then
        objTitle.ParseFromRecordSet tRecordSet
    Else
        Set objTitle = Nothing
    End If
    
    Set FindByCode = objTitle
    
End Function

'@param argLifeNo: 생명번호를 기준으로 DB에서 검색합니다. 리스트를 반환합니다.
Public Function FindByLifeNo(ByVal argLifeNo As String) As Object
    
    Dim sql As String
    Dim tRecordSet As T_RECORD_SET
    Dim titleList As Object
    Set titleList = CreateObject("System.Collections.ArrayList")
    
    sql = _
        "SELECT * FROM " & TABLE_TITLE & " t" & _
        "    WHERE t.lifeno = " & SText(argLifeNo) & _
        "    ORDER BY t.start_dt;"
    
    tRecordSet = makeListData(sql, TABLE_TITLE)
    
    Dim objTitle As title
    If tRecordSet.CNT_RECORD > 0 Then
        Dim i As Long
        For i = 0 To tRecordSet.CNT_RECORD - 1
            Set objTitle = New title
            objTitle.ParseFromRecordSet tRecordSet, i
            titleList.Add objTitle
        Next
    Else
        Set titleList = Nothing
    End If
    
    Set FindByLifeNo = titleList
    
End Function

'@param argLifeNo: 생명번호를 기준으로 DB에서 검색합니다. 리스트를 반환합니다.
Public Function FindByTitle(ByRef argTitle As title) As title
    
    Set FindByTitle = FindByCode(argTitle.Code)
    
End Function

'@param argTitle: argTitle 객체를 받아 DB에 Insert 또는 Update 합니다.
Public Sub Save(ByRef argTitle As title)
    Dim tmpTitle As New title
    
    Set tmpTitle = FindByTitle(argTitle)
    
    If tmpTitle.Code = 0 Then
        Insert argTitle
    Else
        Update argTitle
    End If
    
End Sub

'@param: argTitle: argTitle 객체를 받아 DB에 Insert 합니다.
Private Sub Insert(ByRef argTitle As title)
    Dim sql As String
    Dim procesureNM As String
    Dim jobNM As String
    Dim tableNM As String
    
    procesureNM = "TitleDao.Insert()"
    jobNM = "직분이력 추가"
    tableNM = TABLE_TITLE
    
    '--//INSERT 문
    sql = _
        "INSERT INTO " & tableNM & _
            " (lifeno, start_dt, end_dt, Title, title_ordinary_date) " & _
        "VALUES(" & _
            SText(argTitle.lifeNo) & "," & SText(argTitle.startDate) & "," & _
            SText(argTitle.endDate) & "," & SText(argTitle.title) & "," & _
            SText(argTitle.TitleOrdinaryDate) & _
        ");"
    
    executeSqlWithLog sql, procesureNM, tableNM, jobNM
    
End Sub

'@param: argTitle: argTitle 객체를 받아 DB에 Update 합니다.
Private Sub Update(ByRef argTitle As title)
    Dim sql As String
    Dim procesureNM As String
    Dim jobNM As String
    Dim tableNM As String
    
    procesureNM = "TitleDao.Update()"
    jobNM = "직분이력 업데이트"
    tableNM = TABLE_TITLE
    
    '--//UPDATE 문
    sql = _
        "UPDATE " & tableNM & " SET " & _
        "lifeno = " & SText(argTitle.lifeNo) & ", start_dt = " & SText(argTitle.startDate) & _
        ", end_dt = " & SText(argTitle.endDate) & ", Title = " & SText(argTitle.title) & ", title_ordinary_date = " & SText(argTitle.TitleOrdinaryDate) & _
        " WHERE title_cd = " & SText(argTitle.Code) & ";"
    
    executeSqlWithLog sql, procesureNM, tableNM, jobNM
    
End Sub

'@param: argTitle: argTitle 객체를 받아 DB에서 Delete 합니다.
Public Sub Delete(ByRef argTitle As title)

    Dim sql As String
    Dim procesureNM As String
    Dim jobNM As String
    Dim tableNM As String
    
    procesureNM = "TitleDao.Delete()"
    jobNM = "직분이력 삭제"
    tableNM = TABLE_TITLE
    
    '--//UPDATE 문
    sql = _
        "DELETE w FROM " & tableNM & " w WHERE w.title_cd = " & SText(argTitle.Code) & ";"
    
    executeSqlWithLog sql, procesureNM, tableNM, jobNM

End Sub

'@param argTitle: argTitle 객체에 설정된 시작일과 종료일이 DB에 저장된 다른 발령이력과 겹치는지 점검
Public Function CheckTimeOverlapped(ByRef argTitle As title) As Boolean
    
    Dim sql As String
    Dim tRecordSet As T_RECORD_SET
    
    sql = "SELECT * FROM " & TABLE_TITLE & " a WHERE a.lifeno = " & SText(argTitle.lifeNo) & _
            " AND IF(a.start_dt > " & SText(argTitle.startDate) & ", a.start_dt, " & SText(argTitle.startDate) & ") <= " & _
            "IF(a.end_dt < " & SText(argTitle.endDate) & ", a.end_dt, " & SText(argTitle.endDate) & ")" & _
            " AND a.title_cd <> " & SText(argTitle.Code) & ";"
    
    tRecordSet = makeListData(sql, TABLE_TITLE)
    
    CheckTimeOverlapped = (tRecordSet.CNT_RECORD > 0)
    
End Function

'--//직분목록을 DB에서 받아 반환 합니다.
Public Function GetTitleList() As Object

    Dim sql As String
    Dim tRecordSet As T_RECORD_SET
    Dim titleList As Object
    Set titleList = CreateObject("System.Collections.ArrayList")
    
    sql = "SELECT * FROM " & A_TITLE & ";"
    tRecordSet = makeListData(sql, A_TITLE)
    
    With tRecordSet
        If .CNT_RECORD > 0 Then
            Dim strTitle As Variant
            For Each strTitle In .LISTDATA
                titleList.Add strTitle
            Next
        End If
    End With
    
    Set GetTitleList = titleList

End Function
