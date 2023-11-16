VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SpecialPosition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'################################
'TABLE: DB_POSITION2에 대한 CLASS
'################################

Dim vCode As Long
Dim vLifeNo As String
Dim vStartDate As Date
Dim vEndDate As Date
Dim vSpecialPosition As String

Public Property Get Code() As Long
    Code = vCode
End Property

Public Property Let Code(argCode As Long)
    vCode = argCode
End Property

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

Public Property Get endDate() As Date
    endDate = vEndDate
End Property

Public Property Let endDate(argEndDate As Date)
    vEndDate = argEndDate
End Property

Public Property Get SpecialPosition() As String
    SpecialPosition = vSpecialPosition
End Property

Public Property Let SpecialPosition(argSpecialPosition As String)
    vSpecialPosition = argSpecialPosition
End Property

'@param form: 발령직분직책 폼을 변수로 받아 객체 속성에 값을 채웁니다.
Public Sub ParseFromForm(ByRef form As frm_Update_Appointment)
'    Call frm_Update_PInformation.FillWithDefaultValue
    With form.lstHistory
        If .listIndex >= 0 Then
            If form.cmdEdit.Visible Then
                vCode = .List(.listIndex, 0)
            End If
        End If
    End With
    
    With form.lstPStaff
        vLifeNo = .List(.listIndex)
    End With
    
    vStartDate = form.txtStart
    
    If form.txtEnd = "현재" Then
        vEndDate = DateSerial(9999, 12, 31)
    Else
        vEndDate = form.txtEnd
    End If
    
    vSpecialPosition = form.txtChurchNow
End Sub

'@param tRecordSet: DB에서 받은 RecordSet 타입을 변수로 받아 객체 속성에 값을 채웁니다.
Public Sub ParseFromRecordSet(ByRef tRecordSet As T_RECORD_SET, Optional index As Long = 0)
    With tRecordSet
        vCode = .LISTDATA(index, 0)
        vLifeNo = .LISTDATA(index, 1)
        
        If .LISTDATA(index, 2) <> "" Then
            vStartDate = .LISTDATA(index, 2)
        Else
            vStartDate = DateSerial(1900, 1, 1)
        End If
        
        If .LISTDATA(index, 3) <> "" Then
            vEndDate = .LISTDATA(index, 3)
        Else
            vEndDate = DateSerial(1900, 1, 1)
        End If
        
        vSpecialPosition = .LISTDATA(index, 4)
    End With
End Sub

'@param argSpecialPosition: argSpecialPosition 객체를 변수로 받아 동일객체 여부를 점검합니다.
'                 생명번호, 시작일, 교회코드가 모두 동일하면 같은 객체로 간주합니다.
'@param blnStrictly: 변수 값이 True는 강한 검증을 말하며, 종료일까지 같아야 같은 객체로 간주합니다.
'                    선택사항이며 기본값은 False 입니다.
Public Function IsEqual(ByRef argSpecialPosition As SpecialPosition, Optional blnStrictly As Boolean = False) As Boolean
            
    If Me.lifeNo <> argSpecialPosition.lifeNo Then IsEqual = False: Exit Function
    If Me.startDate <> argSpecialPosition.startDate Then IsEqual = False: Exit Function
    If Me.SpecialPosition <> argSpecialPosition.SpecialPosition Then IsEqual = False: Exit Function
    
    If blnStrictly And argSpecialPosition.lifeNo <> "" Then
        If Me.endDate <> argSpecialPosition.endDate Then IsEqual = False: Exit Function
    End If
    IsEqual = True
End Function


