VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Visa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'############################
'TABLE: DB_VISA에 대한 CLASS
'############################

Dim vCode As Long
Dim vLifeNo As String
Dim vStartDate As Date
Dim vEndDate As Date
Dim vVisa As String
Dim vMemo As String

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

Public Property Get Visa() As String
    Visa = vVisa
End Property

Public Property Let Visa(argVisa As String)
    vVisa = argVisa
End Property

Public Property Get memo() As String
    memo = vMemo
End Property

Public Property Let memo(argMemo As String)
    vMemo = argMemo
End Property

'@param form: 비자 폼을 변수로 받아 객체 속성에 값을 채웁니다.
Public Sub ParseFromForm(ByRef form As frm_Update_Visa)
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
    
    vVisa = form.cboVisa
    vMemo = form.txtMemo
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
        
        vVisa = .LISTDATA(index, 4)
        vMemo = .LISTDATA(index, 5)
    End With
End Sub

'@param argVisa: argVisa 객체를 변수로 받아 동일객체 여부를 점검합니다.
'                 생명번호, 시작일, 교회코드가 모두 동일하면 같은 객체로 간주합니다.
'@param blnStrictly: 변수 값이 True는 강한 검증을 말하며, 종료일, 안수일까지 같아야 같은 객체로 간주합니다.
'                    선택사항이며 기본값은 False 입니다.
Public Function IsEqual(ByRef argVisa As Visa, Optional blnStrictly As Boolean = False) As Boolean
            
    If Me.lifeNo <> argVisa.lifeNo Then IsEqual = False: Exit Function
    If Me.startDate <> argVisa.startDate Then IsEqual = False: Exit Function
    If Me.Visa <> argVisa.Visa Then IsEqual = False: Exit Function
    
    If blnStrictly And argVisa.lifeNo <> "" Then
        If Me.endDate <> argVisa.endDate Then IsEqual = False: Exit Function
        If Me.memo <> argVisa.memo Then IsEqual = False: Exit Function
    End If

    IsEqual = True
End Function

