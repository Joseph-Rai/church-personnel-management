VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SermonScore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim vLifeNo As String
Dim vAvgScore As Double
Dim vCntSubject As Integer

Public Property Get lifeNo() As String
    lifeNo = vLifeNo
End Property

Public Property Let lifeNo(argLifeNo As String)
    vLifeNo = argLifeNo
End Property

Public Property Get AvgScore() As Double
    AvgScore = vAvgScore
End Property

Public Property Let AvgScore(argAvgScore As Double)
    vAvgScore = argAvgScore
End Property

Public Property Get CntSubject() As Integer
    CntSubject = vCntSubject
End Property

Public Property Let CntSubject(argCntSubject As Integer)
    vCntSubject = argCntSubject
End Property

'@param tRecordSet: DB에서 받은 RecordSet 타입을 변수로 받아 객체 속성에 값을 채웁니다.
Public Sub ParseFromRecordSet(ByRef tRecordSet As T_RECORD_SET, Optional index As Long = 0)
    With tRecordSet
        vLifeNo = .LISTDATA(index, 0)
        vAvgScore = .LISTDATA(index, 1)
        vCntSubject = .LISTDATA(index, 2)
    End With
End Sub
