VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "User"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'################################
'TABLE: COMMON.USERSø° ¥Î«— CLASS
'################################

Dim vId As Integer
Dim vName As String
Dim vGb As String
Dim vPassword As String
Dim vInitialize As Integer
Dim vIP As String
Dim vDept As String
Dim vSuspend As Integer

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

Public Property Get Password() As String
    Password = vPassword
End Property

Public Property Let Password(argPassword As String)
    vPassword = argPassword
End Property

Public Property Get Initialize() As String
    Initialize = vInitialize
End Property

Public Property Let Initialize(argInitialize As String)
    vInitialize = argInitialize
End Property

Public Property Get IP() As String
    IP = vIP
End Property

Public Property Let IP(argIP As String)
    vIP = argIP
End Property

Public Property Get Dept() As String
    Dept = vDept
End Property

Public Property Let Dept(argDept As String)
    vDept = argDept
End Property

Public Property Get Suspend() As String
    Suspend = vSuspend
End Property

Public Property Let Suspend(argSuspend As String)
    vSuspend = argSuspend
End Property

Public Sub ParseFromForm(ByRef form As frm_Update_User)
    vName = form.txtUsername
    vDept = form.cboDepartment
End Sub

Public Sub ParseFromRecordSet(ByRef tRecordSet As T_RECORD_SET)
    With tRecordSet
        vId = .LISTDATA(0, 0)
        vName = .LISTDATA(0, 1)
        vGb = .LISTDATA(0, 2)
        vPassword = .LISTDATA(0, 3)
        vInitialize = .LISTDATA(0, 4)
        vIP = .LISTDATA(0, 5)
        vDept = .LISTDATA(0, 6)
        argIP = .LISTDATA(0, 8)
        argDB = .LISTDATA(0, 9)
        argUN = .LISTDATA(0, 10)
        argPW = .LISTDATA(0, 11)
        vSuspend = .LISTDATA(0, 7)
    End With
End Sub
