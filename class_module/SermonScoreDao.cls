VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SermonScoreDao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'@param lifeNo: 생명번호를 기준으로 DB에서 검색합니다.
Public Function FindByLifeNo(ByVal lifeNo As String) As SermonScore
    
    Dim sql As String
    Dim tRecordSet As T_RECORD_SET
    
    sql = _
        "SELECT * FROM " & TABLE_SERMON & " s" & _
        "    WHERE s.lifeno = " & SText(lifeNo) & ";"
    
    tRecordSet = makeListData(sql, TABLE_SERMON)
    
    Dim objSermon As New SermonScore
    If tRecordSet.CNT_RECORD > 0 Then
        objSermon.ParseFromRecordSet tRecordSet
    Else
'        Set objSermon = Nothing
    End If
    
    Set FindByLifeNo = objSermon
    
End Function

