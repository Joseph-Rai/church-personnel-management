VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OvsDepartment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'##################################
'TABLE: DB_OVS_DEPTø° ¥Î«— CLASS
'##################################

Dim vId As Integer
Dim vLevel1 As String
Dim vLevel2 As String
Dim vLevel3 As String
Dim vDeptName As String
Dim vPhoneCard As String
Dim vPicPath As String
Dim vSortOrder As Integer
Dim vSuspend As Integer

Public Property Get Id() As String
    Id = vId
End Property

Public Property Let Id(argId As String)
    vId = argId
End Property

Public Property Get Level1() As String
    Level1 = vLevel1
End Property

Public Property Let Level1(argLevel1 As String)
    vLevel1 = argLevel1
End Property

Public Property Get Level2() As String
    Level2 = vLevel2
End Property

Public Property Let Level2(argLevel2 As String)
    vLevel2 = argLevel2
End Property

Public Property Get Level3() As String
    Level3 = vLevel3
End Property

Public Property Let Level3(argLevel3 As String)
    vLevel3 = argLevel3
End Property

Public Property Get DeptName() As String
    DeptName = vDeptName
End Property

Public Property Let DeptName(argDeptName As String)
    vDeptName = argDeptName
End Property

Public Property Get PhoneCard() As String
    PhoneCard = vPhoneCard
End Property

Public Property Let PhoneCard(argPhoneCard As String)
    vPhoneCard = argPhoneCard
End Property

Public Property Get PicPath() As String
    PicPath = vPicPath
End Property

Public Property Let PicPath(argPicPath As String)
    vPicPath = argPicPath
End Property

Public Property Get SortOrder() As String
    SortOrder = vSortOrder
End Property

Public Property Let SortOrder(argSortOrder As String)
    vSortOrder = argSortOrder
End Property

Public Property Get Suspend() As String
    Suspend = vSuspend
End Property

Public Property Let Suspend(argSuspend As String)
    vSuspend = argSuspend
End Property

Public Sub ParseFromRecordSet(ByRef tRecordSet As T_RECORD_SET, Optional index As Long = 0)
    vId = tRecordSet.LISTDATA(index, 0)
    vLevel1 = tRecordSet.LISTDATA(index, 1)
    vLevel2 = tRecordSet.LISTDATA(index, 2)
    vLevel3 = tRecordSet.LISTDATA(index, 3)
    DeptName = tRecordSet.LISTDATA(index, 4)
    PhoneCard = tRecordSet.LISTDATA(index, 5)
    PicPath = tRecordSet.LISTDATA(index, 6)
    SortOrder = tRecordSet.LISTDATA(index, 7)
    Suspend = tRecordSet.LISTDATA(index, 8)
End Sub
