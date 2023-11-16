VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Attendance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'##################################
'TABLE: DB_ATTENDANCE에 대한 CLASS
'##################################

Dim vChurchID As String
Dim vAttendanceDate As Date
Dim vOnceAll As Long
Dim vForthAll As Long
Dim vOnceStu As Long
Dim vForthStu As Long
Dim vTitheAll As Long
Dim vTitheStu As Long
Dim vBaptismAll As Long
Dim vEvangelist As Long
Dim vGroupLeader As Long
Dim vUnitLeader As Long

Public Property Get ChurchID() As String
    ChurchID = vChurchID
End Property

Public Property Let ChurchID(argChurchId As String)
    vChurchID = argChurchId
End Property

Public Property Get AttendanceDate() As Date
    AttendanceDate = vAttendanceDate
End Property

Public Property Let AttendanceDate(argAttendanceDate As Date)
    vAttendanceDate = argAttendanceDate
End Property

Public Property Get OnceAll() As Long
    OnceAll = vOnceAll
End Property

Public Property Let OnceAll(argOnceAll As Long)
    vOnceAll = argOnceAll
End Property

Public Property Get ForthAll() As Long
    ForthAll = vForthAll
End Property

Public Property Let ForthAll(argForthAll As Long)
    vForthAll = argForthAll
End Property

Public Property Get OnceStu() As Long
    OnceStu = vOnceStu
End Property

Public Property Let OnceStu(argOnceStu As Long)
    vOnceStu = argOnceStu
End Property

Public Property Get ForthStu() As Long
    ForthStu = vForthStu
End Property

Public Property Let ForthStu(argForthStu As Long)
    vForthStu = argForthStu
End Property

Public Property Get TitheAll() As Long
    TitheAll = vTitheAll
End Property

Public Property Let TitheAll(argTitheAll As Long)
    vTitheAll = argTitheAll
End Property

Public Property Get TitheStu() As Long
    TitheStu = vTitheStu
End Property

Public Property Let TitheStu(argTitheStu As Long)
    vTitheStu = argTitheStu
End Property

Public Property Get BaptismAll() As Long
    BaptismAll = vBaptismAll
End Property

Public Property Let BaptismAll(argBaptismAll As Long)
    vBaptismAll = argBaptismAll
End Property

Public Property Get Evangelist() As Long
    Evangelist = vEvangelist
End Property

Public Property Let Evangelist(argEvangelist As Long)
    vEvangelist = argEvangelist
End Property

Public Property Get GroupLeader() As Long
    GroupLeader = vGroupLeader
End Property

Public Property Let GroupLeader(argGroupLeader As Long)
    vGroupLeader = argGroupLeader
End Property

Public Property Get UnitLeader() As Long
    UnitLeader = vUnitLeader
End Property

Public Property Let UnitLeader(argUnitLeader As Long)
    vUnitLeader = argUnitLeader
End Property

'@param tRecordSet: DB에서 받은 RecordSet 타입을 변수로 받아 객체 속성에 값을 채웁니다.
Public Sub ParseFromRecordSet(ByRef tRecordSet As T_RECORD_SET, Optional index As Long = 0)
    With tRecordSet
        vChurchID = .LISTDATA(index, 0)
        If .LISTDATA(index, 1) <> "" Then
            vAttendanceDate = .LISTDATA(index, 1)
        Else
            vAttendanceDate = DateSerial(1900, 1, 1)
        End If
        vOnceAll = .LISTDATA(index, 2)
        vForthAll = .LISTDATA(index, 3)
        vOnceStu = .LISTDATA(index, 4)
        vForthStu = .LISTDATA(index, 5)
        vTitheAll = .LISTDATA(index, 6)
        vTitheStu = .LISTDATA(index, 7)
        vBaptismAll = .LISTDATA(index, 8)
        vEvangelist = .LISTDATA(index, 9)
        vGroupLeader = .LISTDATA(index, 10)
        vUnitLeader = .LISTDATA(index, 11)
    End With
End Sub

Public Sub ParseFromRange(ByRef targetRange As Range)
    
    vChurchID = targetRange
    vAttendanceDate = targetRange.Offset(, 1)
    vOnceAll = targetRange.Offset(, 2)
    vForthAll = targetRange.Offset(, 3)
    vOnceStu = targetRange.Offset(, 4)
    vForthStu = targetRange.Offset(, 5)
    vTitheAll = targetRange.Offset(, 6)
    vTitheStu = targetRange.Offset(, 7)
    vBaptismAll = targetRange.Offset(, 8)
    vEvangelist = targetRange.Offset(, 9)
    vGroupLeader = targetRange.Offset(, 10)
    vUnitLeader = targetRange.Offset(, 11)
    
End Sub

Public Sub Sum(ByRef argAtten As Attendance)

    With argAtten
        vOnceAll = vOnceAll + .OnceAll
        vForthAll = vForthAll + .ForthAll
        vOnceStu = vOnceStu + .OnceStu
        vForthStu = vForthStu + .ForthStu
        vTitheAll = vTitheAll + .TitheAll
        vTitheStu = vTitheStu + .TitheStu
        vBaptismAll = vBaptismAll + .BaptismAll
        vEvangelist = vEvangelist + .Evangelist
        vGroupLeader = vGroupLeader + .GroupLeader
        vUnitLeader = vUnitLeader + .UnitLeader
    End With

End Sub

Public Sub Subtract(ByRef argAtten As Attendance)

    With argAtten
        vOnceAll = vOnceAll - .OnceAll
        vForthAll = vForthAll - .ForthAll
        vOnceStu = vOnceStu - .OnceStu
        vForthStu = vForthStu - .ForthStu
        vTitheAll = vTitheAll - .TitheAll
        vTitheStu = vTitheStu - .TitheStu
        vBaptismAll = vBaptismAll - .BaptismAll
        vEvangelist = vEvangelist - .Evangelist
        vGroupLeader = vGroupLeader - .GroupLeader
        vUnitLeader = vUnitLeader - .UnitLeader
    End With

End Sub

'@param argAtten: argAtten 객체를 변수로 받아 교회코드와 날짜를 기준으로 동일한지 여부를 점검합니다.
'@param blnStrictly: 변수 값이 True는 강한 검증을 말하며, 교회코드와 날짜를 제외한 나머지 정보까지 모두 같을 때 True를 반환합니다.
'                    선택사항이며 기본값은 False 입니다.
Public Function IsEqual(ByRef argAtten As Attendance, Optional blnStrictly As Boolean = False) As Boolean
    
    If Me.ChurchID <> argAtten.ChurchID Then IsEqual = False: Exit Function
    If Me.AttendanceDate <> argAtten.AttendanceDate Then IsEqual = False: Exit Function
    
    If blnStrictly And argAtten.ChurchID <> "" Then
        If Me.OnceAll <> argAtten.OnceAll Then IsEqual = False: Exit Function
        If Me.ForthAll <> argAtten.ForthAll Then IsEqual = False: Exit Function
        If Me.OnceStu <> argAtten.OnceStu Then IsEqual = False: Exit Function
        If Me.ForthStu <> argAtten.ForthStu Then IsEqual = False: Exit Function
        If Me.TitheAll <> argAtten.TitheAll Then IsEqual = False: Exit Function
        If Me.TitheStu <> argAtten.TitheStu Then IsEqual = False: Exit Function
        If Me.BaptismAll <> argAtten.BaptismAll Then IsEqual = False: Exit Function
        If Me.Evangelist <> argAtten.Evangelist Then IsEqual = False: Exit Function
        If Me.GroupLeader <> argAtten.GroupLeader Then IsEqual = False: Exit Function
        If Me.UnitLeader <> argAtten.UnitLeader Then IsEqual = False: Exit Function
    End If
    IsEqual = True
End Function
