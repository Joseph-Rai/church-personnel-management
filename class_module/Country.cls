VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Country"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'##################################
'TABLE: DB_COUNTRYø° ¥Î«— CLASS
'##################################

Dim vCountryNameKo As String
Dim vCountryNameEn As String
Dim vPopulation As Long

Public Property Get CountryNameKo() As String
    CountryNameKo = vCountryNameKo
End Property

Public Property Let CountryNameKo(argCountryNameKo As String)
    vCountryNameKo = argCountryNameKo
End Property

Public Property Get CountryNameEn() As String
    CountryNameEn = vCountryNameEn
End Property

Public Property Let CountryNameEn(argCountryNameEn As String)
    vCountryNameEn = argCountryNameEn
End Property

Public Property Get Population() As String
    Population = vPopulation
End Property

Public Property Let Population(argPopulation As String)
    vPopulation = argPopulation
End Property

Public Sub ParseFromRecordSet(ByRef tRecordSet As T_RECORD_SET, Optional index As Long = 0)
    vCountryNameKo = tRecordSet.LISTDATA(index, 0)
    vCountryNameEn = tRecordSet.LISTDATA(index, 1)
    vPopulation = tRecordSet.LISTDATA(index, 2)
End Sub

