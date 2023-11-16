VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CountryDao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'--//DB에 저장된 국가리스트를 가져옵니다.
Public Function GetCountryList() As Object

    Dim CountryList As Object
    Dim sql As String
    Dim tRecordSet As T_RECORD_SET
    
    Set CountryList = CreateObject("System.Collections.ArrayList")
    
    sql = _
        "SELECT c.ctry_nm FROM " & TABLE_COUNTRY & " c;"
    
    tRecordSet = makeListData(sql, TABLE_COUNTRY)
    
    Dim recordCountry As Variant
    If tRecordSet.CNT_RECORD > 0 Then
        For Each recordCountry In tRecordSet.LISTDATA
            CountryList.Add recordCountry
        Next
    End If
    
    Set GetCountryList = CountryList

End Function

