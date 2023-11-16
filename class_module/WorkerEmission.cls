VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorkerEmission"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'#####################################
'동역자, 예비생도, 지역장, 구역장 배출
'#####################################

Dim vEmitAssistant As Integer
Dim vEmitTheologicalStu As Integer
Dim vEmitGroupLeader As Integer
Dim vEmitUnitLeader As Integer

Public Property Get EmitAssistant() As Long
    EmitAssistant = vEmitAssistant
End Property

Public Property Let EmitAssistant(argEmitAssistant As Long)
    vEmitAssistant = argEmitAssistant
End Property

Public Property Get EmitTheologicalStu() As Long
    EmitTheologicalStu = vEmitTheologicalStu
End Property

Public Property Let EmitTheologicalStu(argEmitTheologicalStu As Long)
    vEmitTheologicalStu = argEmitTheologicalStu
End Property

Public Property Get EmitGroupLeader() As Long
    EmitGroupLeader = vEmitGroupLeader
End Property

Public Property Let EmitGroupLeader(argEmitGroupLeader As Long)
    vEmitGroupLeader = argEmitGroupLeader
End Property

Public Property Get EmitUnitLeader() As Long
    EmitUnitLeader = vEmitUnitLeader
End Property

Public Property Let EmitUnitLeader(argEmitUnitLeader As Long)
    vEmitUnitLeader = argEmitUnitLeader
End Property
