VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GeoData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim vId As Long
Dim vCountry As String
Dim vCountryPopulation As Long
Dim vAdmin1ID As Long
Dim vAdmin1Name As String
Dim vAdmin1Population As Long
Dim vAdmin2ID As Long
Dim vAdmin2Name As String
Dim vAdmin2Population As Long
Dim vAdmin3ID As Long
Dim vAdmin3Name As String
Dim vAdmin3Population As Long
Dim vAdmin4ID As Long
Dim vAdmin4Name As String
Dim vAdmin4Population As Long
Dim vGospelContinent As Long
Dim vGospelCountry As Long
Dim vGospelRegion As Long
Dim vGospelCity As Long
Dim vIsCenter As String
Dim vLatitude As Double
Dim vLongitude As Double
Dim vMissionContinent As String
Dim vMissionDepartment As String
Dim vDepartment As String
Dim vDivision As String

Public Property Get Id() As Long
    Id = vId
End Property

Public Property Let Id(argId As Long)
    vId = argId
End Property

Public Property Get Country() As String
    Country = vCountry
End Property

Public Property Let Country(argCountry As String)
    vCountry = argCountry
End Property

Public Property Get CountryPopulation() As Long
    CountryPopulation = vCountryPopulation
End Property

Public Property Let CountryPopulation(argCountryPopulation As Long)
    vCountryPopulation = argCountryPopulation
End Property

Public Property Get Admin1ID() As Long
    Admin1ID = vAdmin1ID
End Property

Public Property Let Admin1ID(argAdmin1ID As Long)
    vAdmin1ID = argAdmin1ID
End Property

Public Property Get Admin1Name() As String
    Admin1Name = vAdmin1Name
End Property

Public Property Let Admin1Name(argAdmin1Name As String)
    vAdmin1Name = argAdmin1Name
End Property

Public Property Get Admin1Population() As Long
    Admin1Population = vAdmin1Population
End Property

Public Property Let Admin1Population(argAdmin1Population As Long)
    vAdmin1Population = argAdmin1Population
End Property

Public Property Get Admin2ID() As Long
    Admin2ID = vAdmin2ID
End Property

Public Property Let Admin2ID(argAdmin2ID As Long)
    vAdmin2ID = argAdmin2ID
End Property

Public Property Get Admin2Name() As String
    Admin2Name = vAdmin2Name
End Property

Public Property Let Admin2Name(argAdmin2Name As String)
    vAdmin2Name = argAdmin2Name
End Property

Public Property Get Admin2Population() As Long
    Admin2Population = vAdmin2Population
End Property

Public Property Let Admin2Population(argAdmin2Population As Long)
    vAdmin2Population = argAdmin2Population
End Property

Public Property Get Admin3ID() As Long
    Admin3ID = vAdmin3ID
End Property

Public Property Let Admin3ID(argAdmin3ID As Long)
    vAdmin3ID = argAdmin3ID
End Property

Public Property Get Admin3Name() As String
    Admin3Name = vAdmin3Name
End Property

Public Property Let Admin3Name(argAdmin3Name As String)
    vAdmin3Name = argAdmin3Name
End Property

Public Property Get Admin3Population() As Long
    Admin3Population = vAdmin3Population
End Property

Public Property Let Admin3Population(argAdmin3Population As Long)
    vAdmin3Population = argAdmin3Population
End Property

Public Property Get Admin4ID() As Long
    Admin4ID = vAdmin4ID
End Property

Public Property Let Admin4ID(argAdmin4ID As Long)
    vAdmin4ID = argAdmin4ID
End Property

Public Property Get Admin4Name() As String
    Admin4Name = vAdmin4Name
End Property

Public Property Let Admin4Name(argAdmin4Name As String)
    vAdmin4Name = argAdmin4Name
End Property

Public Property Get Admin4Population() As Long
    Admin4Population = vAdmin4Population
End Property

Public Property Let Admin4Population(argAdmin4Population As Long)
    vAdmin4Population = argAdmin4Population
End Property

Public Property Get GospelContinent() As Long
    GospelContinent = vGospelContinent
End Property

Public Property Let GospelContinent(argGospelContinent As Long)
    vGospelContinent = argGospelContinent
End Property

Public Property Get GospelCountry() As Long
    GospelCountry = vGospelCountry
End Property

Public Property Let GospelCountry(argGospelCountry As Long)
    vGospelCountry = argGospelCountry
End Property

Public Property Get GospelRegion() As Long
    GospelRegion = vGospelRegion
End Property

Public Property Let GospelRegion(argGospelRegion As Long)
    vGospelRegion = argGospelRegion
End Property

Public Property Get GospelCity() As Long
    GospelCity = vGospelCity
End Property

Public Property Let GospelCity(argGospelCity As Long)
    vGospelCity = argGospelCity
End Property

Public Property Get IsCenter() As String
    IsCenter = vIsCenter
End Property

Public Property Let IsCenter(argIsCenter As String)
    vIsCenter = argIsCenter
End Property

Public Property Get Latitude() As Double
    Latitude = vLatitude
End Property

Public Property Let Latitude(argLatitude As Double)
    vLatitude = argLatitude
End Property

Public Property Get Longitude() As Double
    Longitude = vLongitude
End Property

Public Property Let Longitude(argLongitude As Double)
    vLongitude = argLongitude
End Property

Public Property Get MissionContinent() As String
    MissionContinent = vMissionContinent
End Property

Public Property Let MissionContinent(argMissionContinent As String)
    vMissionContinent = argMissionContinent
End Property

Public Property Get MissionDepartment() As String
    MissionDepartment = vMissionDepartment
End Property

Public Property Let MissionDepartment(argMissionDepartment As String)
    vMissionDepartment = argMissionDepartment
End Property

Public Property Get Department() As String
    Department = vDepartment
End Property

Public Property Let Department(argDepartment As String)
    vDepartment = argDepartment
End Property

Public Property Get Division() As String
    Division = vDivision
End Property

Public Property Let Division(argDivision As String)
    vDivision = argDivision
End Property

Public Sub ParseFromRecordSet(ByRef tRecordSet As T_RECORD_SET, Optional index As Long = 0)
    With tRecordSet
        vId = .LISTDATA(index, 0)
        vCountry = .LISTDATA(index, 1)
        If .LISTDATA(index, 2) <> "" Then
            vCountryPopulation = .LISTDATA(index, 2)
        Else
            vCountryPopulation = 0
        End If
        If .LISTDATA(index, 3) <> "" Then
            vAdmin1ID = .LISTDATA(index, 3)
        Else
            vAdmin1ID = 0
        End If
        vAdmin1Name = .LISTDATA(index, 4)
        If .LISTDATA(index, 5) <> "" Then
            vAdmin1Population = .LISTDATA(index, 5)
        Else
            vAdmin1Population = 0
        End If
        If .LISTDATA(index, 6) <> "" Then
            vAdmin2ID = .LISTDATA(index, 6)
        Else
            vAdmin2ID = 0
        End If
        vAdmin2Name = .LISTDATA(index, 7)
        If .LISTDATA(index, 8) <> "" Then
            vAdmin2Population = .LISTDATA(index, 8)
        Else
            vAdmin2Population = 0
        End If
        If .LISTDATA(index, 9) <> "" Then
            vAdmin3ID = .LISTDATA(index, 9)
        Else
            vAdmin3ID = 0
        End If
        vAdmin3Name = .LISTDATA(index, 10)
        If .LISTDATA(index, 11) <> "" Then
            vAdmin3Population = .LISTDATA(index, 11)
        Else
            vAdmin3Population = 0
        End If
        If .LISTDATA(index, 12) <> "" Then
            vAdmin4ID = .LISTDATA(index, 12)
        Else
            vAdmin4ID = 0
        End If
        vAdmin4Name = .LISTDATA(index, 13)
        If .LISTDATA(index, 14) <> "" Then
            vAdmin4Population = .LISTDATA(index, 14)
        Else
            vAdmin4Population = 0
        End If
        If .LISTDATA(index, 15) <> "" Then
            vGospelContinent = .LISTDATA(index, 15)
        Else
            vGospelContinent = 0
        End If
        If .LISTDATA(index, 16) <> "" Then
            vGospelCountry = .LISTDATA(index, 16)
        Else
            vGospelCountry = 0
        End If
        If .LISTDATA(index, 17) <> "" Then
            vGospelRegion = .LISTDATA(index, 17)
        Else
            vGospelRegion = 0
        End If
        If .LISTDATA(index, 18) <> "" Then
            vGospelCity = .LISTDATA(index, 18)
        Else
            vGospelCity = 0
        End If
        vIsCenter = .LISTDATA(index, 19)
        If .LISTDATA(index, 20) <> "" Then
            vLatitude = .LISTDATA(index, 20)
        Else
            vLatitude = 0
        End If
        If .LISTDATA(index, 21) <> "" Then
            vLongitude = .LISTDATA(index, 21)
        Else
            vLongitude = 0
        End If
        vMissionContinent = .LISTDATA(index, 22)
        vMissionDepartment = .LISTDATA(index, 23)
        vDepartment = .LISTDATA(index, 24)
        vDivision = .LISTDATA(index, 25)
    End With
End Sub

Public Sub InsertToRange(ByRef rngT As Range)

    rngT.Offset(0, 0) = vId
    rngT.Offset(0, 1) = vCountry
    rngT.Offset(0, 2) = vCountryPopulation
    rngT.Offset(0, 3) = vAdmin1ID
    rngT.Offset(0, 4) = vAdmin1Name
    rngT.Offset(0, 5) = vAdmin1Population
    rngT.Offset(0, 6) = vAdmin2ID
    rngT.Offset(0, 7) = vAdmin2Name
    rngT.Offset(0, 8) = vAdmin2Population
    rngT.Offset(0, 9) = vAdmin3ID
    rngT.Offset(0, 10) = vAdmin3Name
    rngT.Offset(0, 11) = vAdmin3Population
    rngT.Offset(0, 12) = vAdmin4ID
    rngT.Offset(0, 13) = vAdmin4Name
    rngT.Offset(0, 14) = vAdmin4Population
    rngT.Offset(0, 15) = vGospelContinent
    rngT.Offset(0, 16) = vGospelCountry
    rngT.Offset(0, 17) = vGospelRegion
    rngT.Offset(0, 18) = vGospelCity
    rngT.Offset(0, 19) = vIsCenter
    rngT.Offset(0, 20) = vLatitude
    rngT.Offset(0, 21) = vLongitude
    rngT.Offset(0, 22) = vMissionContinent
    rngT.Offset(0, 23) = vMissionDepartment
    rngT.Offset(0, 24) = vDepartment
    rngT.Offset(0, 25) = vDivision

End Sub
