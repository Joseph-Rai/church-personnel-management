VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PastoralCareer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim vLifeNo As String
Dim vStartDate As Date
Dim vTitlePosition As String
Dim vChurchClass As String
Dim vChurchName As String
Dim vChurchCode As String
Dim vPeriod As String
Dim vHistoryClass As String
Dim vEndDate As Date
Dim vTitle As String
Dim vPosition As String

Public Property Get lifeNo() As String
    lifeNo = vLifeNo
End Property

Public Property Let lifeNo(argLifeNo As String)
    vLifeNo = argLifeNo
End Property

Public Property Get startDate() As Date
    startDate = vStartDate
End Property

Public Property Let startDate(argStartDate As Date)
    vStartDate = argStartDate
End Property

Public Property Get TitlePosition() As String
    TitlePosition = vTitlePosition
End Property

Public Property Let TitlePosition(argTitlePosition As String)
    vTitlePosition = argTitlePosition
End Property

Public Property Get ChurchClass() As String
    ChurchClass = vChurchClass
End Property

Public Property Let ChurchClass(argChurchClass As String)
    vChurchClass = argChurchClass
End Property

Public Property Get churchName() As String
    churchName = vChurchName
End Property

Public Property Let churchName(argChurchName As String)
    vChurchName = argChurchName
End Property

Public Property Get churchCode() As String
    churchCode = vChurchCode
End Property

Public Property Let churchCode(argChurchCode As String)
    vChurchCode = argChurchCode
End Property

Public Property Get Period() As String
    Period = vPeriod
End Property

Public Property Let Period(argPeriod As String)
    vPeriod = argPeriod
End Property

Public Property Get HistoryClass() As String
    HistoryClass = vHistoryClass
End Property

Public Property Let HistoryClass(argHistoryClass As String)
    vHistoryClass = argHistoryClass
End Property

Public Property Get endDate() As Date
    endDate = vEndDate
End Property

Public Property Let endDate(argEndDate As Date)
    vEndDate = argEndDate
End Property

Public Property Get title() As String
    title = vTitle
End Property

Public Property Let title(argTitle As String)
    vTitle = argTitle
End Property

Public Property Get position() As String
    position = vPosition
End Property

Public Property Let position(argPosition As String)
    vPosition = argPosition
End Property

'@param tRecordSet: DB에서 받은 RecordSet 타입을 변수로 받아 객체 속성에 값을 채웁니다.
Public Sub ParseFromRecordSet(ByRef tRecordSet As T_RECORD_SET, Optional index As Long = 0)
    
    With tRecordSet
        vLifeNo = .LISTDATA(index, 0)
        If .LISTDATA(index, 1) <> "" Then
            vStartDate = .LISTDATA(index, 1)
        End If
        vTitlePosition = .LISTDATA(index, 2)
        vChurchClass = .LISTDATA(index, 3)
        vChurchName = .LISTDATA(index, 4)
        vChurchCode = .LISTDATA(index, 5)
        vPeriod = .LISTDATA(index, 6)
        vHistoryClass = .LISTDATA(index, 7)
        If .LISTDATA(index, 8) <> "" Then
            vEndDate = .LISTDATA(index, 8)
        End If
        vTitle = .LISTDATA(index, 9)
        vPosition = .LISTDATA(index, 10)
    End With
End Sub
