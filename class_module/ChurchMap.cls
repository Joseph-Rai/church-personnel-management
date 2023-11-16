VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ChurchMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'##################################
'TABLE: DB_CHURCH_MAP에 대한 CLASS
'##################################

Dim vId As String
Dim vMap As String

Public Property Get Id() As String
    Id = vId
End Property

Public Property Let Id(argId As String)
    vId = argId
End Property

Public Property Get map() As String
    map = vMap
End Property

Public Property Let map(argMap As String)
    vMap = argMap
End Property

'@param tRecordSet: DB에서 받은 RecordSet 타입을 변수로 받아 객체 속성에 값을 채웁니다.
Public Sub ParseFromRecordSet(ByRef tRecordSet As T_RECORD_SET, Optional index As Long = 0)
    With tRecordSet
        vId = .LISTDATA(index, 0)
        vMap = .LISTDATA(index, 1)
    End With
End Sub
