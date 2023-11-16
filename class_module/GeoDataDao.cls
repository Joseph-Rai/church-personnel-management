VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GeoDataDao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function FindGeoDataById(argId As Long) As GeoData
    Dim sql As String
    Dim tRecordSet As T_RECORD_SET
    
    sql = "SELECT * FROM " & TABLE_GEODATA & _
                    " a WHERE a.geo_cd = " & SText(argId) & ";"
    
    tRecordSet = makeListData(sql, TABLE_GEODATA)
    
    Dim tmpGeoData As New GeoData
    If tRecordSet.CNT_RECORD > 0 Then
        tmpGeoData.ParseFromRecordSet tRecordSet
    Else
        Set tmpGeoData = Nothing
    End If
    
    Set FindGeoDataById = tmpGeoData
End Function
