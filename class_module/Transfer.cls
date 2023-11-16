VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Transfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'###############################
'TABLE: DB_TRANSFER�� ���� CLASS
'###############################

Dim vCode As Long
Dim vLifeNo As String
Dim vStartDate As Date
Dim vEndDate As Date
Dim vChurchID As String
Dim vChurchName As String

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

Public Property Get ChurchID() As String
    ChurchID = vChurchID
End Property

Public Property Let ChurchID(argChurchId As String)
    vChurchID = argChurchId
End Property

Public Property Get churchName() As String
    churchName = vChurchName
End Property

Public Property Let churchName(argChurchName As String)
    vChurchName = argChurchName
End Property

'@param form: �߷�������å ���� ������ �޾� ��ü �Ӽ��� ���� ä��ϴ�.
Public Sub ParseFromForm(ByRef form As frm_Update_Appointment)
    
    With form.lstHistory
        If .listIndex >= 0 Then
            If form.cmdEdit.Visible Then
                vCode = .List(.listIndex)
            End If
        End If
    End With
    
    With form.lstPStaff
        vLifeNo = .List(.listIndex)
    End With
    
    vStartDate = form.txtStart
    
    If form.txtEnd = "����" Then
        vEndDate = DateSerial(9999, 12, 31)
    Else
        vEndDate = form.txtEnd
    End If
    vChurchID = form.txtChurchNow_sid
    vChurchName = form.txtChurchNow
    
End Sub

'@param tRecordSet: DB���� ���� RecordSet Ÿ���� ������ �޾� ��ü �Ӽ��� ���� ä��ϴ�.
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
        
        vChurchID = .LISTDATA(index, 4)
        vChurchName = .LISTDATA(index, 5)
    End With
    
End Sub

'@param argTrans: argTrans ��ü�� ������ �޾� ���ϰ�ü ���θ� �����մϴ�.
'                 ������ȣ, ������, ��ȸ�ڵ尡 ��� �����ϸ� ���� ��ü�� �����մϴ�.
'@param blnStrictly: ���� ���� True�� ���� ������ ���ϸ�, �����ϱ��� ���ƾ� ���� ��ü�� �����մϴ�.
'                    ���û����̸� �⺻���� False �Դϴ�.
Public Function IsEqual(ByRef argTrans As Transfer, Optional blnStrictly As Boolean = False) As Boolean
            
    If Me.lifeNo <> argTrans.lifeNo Then IsEqual = False: Exit Function
    If Me.startDate <> argTrans.startDate Then IsEqual = False: Exit Function
    If Me.ChurchID <> argTrans.ChurchID Then IsEqual = False: Exit Function
    
    If blnStrictly And argTrans.lifeNo <> "" Then
        If Me.endDate <> argTrans.endDate Then IsEqual = False: Exit Function
    End If
    IsEqual = True
    
End Function