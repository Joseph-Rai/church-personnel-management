VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Church"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'##################################
'TABLE: DB_CHURCH에 대한 CLASS
'##################################

Dim vId As String
Dim vName As String
Dim vGb As String
Dim vManager As String
Dim vMainChurch As String
Dim vStartDate As Date
Dim vEndDate As Date
Dim vOvsDeptName As String
Dim vSuspend As Boolean
Dim vSortOrder As Long
Dim vGeoCode As Long
Dim vCountryId As Integer

Public Property Get Id() As String
    Id = vId
End Property

Public Property Let Id(argId As String)
    vId = argId
End Property

Public Property Get Name() As String
    Name = vName
End Property

Public Property Let Name(argName As String)
    vName = argName
End Property

Public Property Get Gb() As String
    Gb = vGb
End Property

Public Property Let Gb(argGb As String)
    vGb = argGb
End Property

Public Property Get Manager() As String
    Manager = vManager
End Property

Public Property Let Manager(argManager As String)
    vManager = argManager
End Property

Public Property Get MainChurch() As String
    MainChurch = vMainChurch
End Property

Public Property Let MainChurch(argMainChurch As String)
    vMainChurch = argMainChurch
End Property

Public Property Get startDate() As Date
    startDate = vStartDate
End Property

Public Property Let startDate(argStartDate As Date)
    vStartDate = argStartDate
End Property

Public Property Get endDate() As Date
    endDate = vEndDate
End Property

Public Property Let endDate(argEndDate As Date)
    vEndDate = argEndDate
End Property

Public Property Get OvsDeptName() As String
    OvsDeptName = vOvsDeptName
End Property

Public Property Let OvsDeptName(argOvdDeptName As String)
    vOvdDeptName = argOvdDeptName
End Property

Public Property Get Suspend() As Boolean
    If vSuspend = 0 Then
        Suspend = False
    Else
        Suspend = True
    End If
End Property

Public Property Let Suspend(argSuspend As Boolean)
    If argSuspend Then
        vSuspend = 1
    Else
        vSuspend = 0
    End If
End Property

Public Property Get SortOrder() As Long
    SortOrder = vSortOrder
End Property

Public Property Let SortOrder(argSortOrder As Long)
    vSortOrder = argSortOrder
End Property

Public Property Get GeoCode() As Long
    GeoCode = vGeoCode
End Property

Public Property Let GeoCode(argGeoCode As Long)
    vGeoCode = argGeoCode
End Property

Public Property Get CountryId() As Integer
    CountryId = vCountryId
End Property

Public Property Let CountryId(argCountryId As Integer)
    vCountryId = argCountryId
End Property

'@param tRecordSet: DB에서 받은 RecordSet 타입을 변수로 받아 객체 속성에 값을 채웁니다.
Public Sub ParseFromRecordSet(ByRef tRecordSet As T_RECORD_SET, Optional index As Long = 0)
    With tRecordSet
        vId = .LISTDATA(index, 0)
        vName = .LISTDATA(index, 1)
        vGb = .LISTDATA(index, 2)
        vManager = .LISTDATA(index, 3)
        vMainChurch = .LISTDATA(index, 4)
        If .LISTDATA(index, 5) <> "" Then
            vStartDate = .LISTDATA(index, 5)
        Else
            vStartDate = DateSerial(1900, 1, 1)
        End If
        If .LISTDATA(index, 6) <> "" Then
            vEndDate = .LISTDATA(index, 6)
        Else
            vEndDate = DateSerial(1900, 1, 1)
        End If
        vOvsDeptName = .LISTDATA(index, 7)
        vSuspend = .LISTDATA(index, 8)
        vSortOrder = .LISTDATA(index, 9)
        If IsNumeric(.LISTDATA(index, 10)) Then
            vGeoCode = .LISTDATA(index, 10)
        End If
        If UBound(.LISTFIELD) > 10 Then
            vCountryId = .LISTDATA(index, 11)
        End If
    End With
End Sub

'@param targetRange: church 객체 정보를 담고 있는 Worksheet의 한 행을 인수로 전달 합니다.
Public Sub ParseFromRange(ByRef targetRange As Range)
    
    vId = targetRange
    vName = targetRange.Offset(, 1)
    vGb = targetRange.Offset(, 2)
    vManager = targetRange.Offset(, 3)
    vMainChurch = targetRange.Offset(, 4)
    vStartDate = targetRange.Offset(, 5)
    vEndDate = targetRange.Offset(, 6)
    vOvsDeptName = targetRange.Offset(, 7)
    vSuspend = targetRange.Offset(, 8)
    vSortOrder = targetRange.Offset(, 9)
    vGeoCode = targetRange.Offset(, 10)
    
End Sub


'@param argChurch: argChurch 객체를 변수로 받아 교회코드 기준으로 동일한지 여부를 점검합니다.
'@param blnStrictly: 변수 값 True는 강한 검증을 말하며, 교회코드를 제외한 나머지 정보까지 모두 같을 때 True를 반환합니다.
'                    선택사항이며 기본값은 False 입니다.
Public Function IsEqual(ByRef argChurch As Church, Optional blnStrictly As Boolean = False) As Boolean
            
    If Me.Id <> argChurch.Id Then IsEqual = False: Exit Function
    
    If blnStrictly Then
        If Me.Name <> argChurch.Name Then IsEqual = False: Exit Function
        If Me.Gb <> argChurch.Gb Then IsEqual = False: Exit Function
        If Me.Manager <> argChurch.Manager Then IsEqual = False: Exit Function
        If Me.MainChurch <> argChurch.MainChurch Then IsEqual = False: Exit Function
        If Me.startDate <> argChurch.startDate Then IsEqual = False: Exit Function
        If Me.endDate <> argChurch.endDate Then IsEqual = False: Exit Function
        If Me.OvsDeptName <> argChurch.OvsDeptName Then IsEqual = False: Exit Function
        If Me.Suspend <> argChurch.Suspend Then IsEqual = False: Exit Function
        If Me.SortOrder <> argChurch.SortOrder Then IsEqual = False: Exit Function
        If Me.GeoCode <> argChurch.GeoCode Then IsEqual = False: Exit Function
        If Me.CountryId <> argChurch.CountryId Then IsEqual = False: Exit Function
    End If
    IsEqual = True
End Function
