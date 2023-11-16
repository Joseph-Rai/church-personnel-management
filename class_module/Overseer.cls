VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Overseer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'--//�� Ŭ������ A3�λ�߷� ��Ʈ���� ������ ����ȸ�� �־��� ��ÿ� �޾��� ��ȸ��� ����� �޸�� ǥ���ϱ� ���� ���� ��ü �Դϴ�.

Public churchSid As String
Public churchName As String
Public startDate As Date
Public endDate As Date
Public lifeNo As String
Public nameKo As String
Public title As String
Public position As String

'@param tRecordSet: DB���� ���� RecordSet Ÿ���� ������ �޾� ��ü �Ӽ��� ���� ä��ϴ�.
Public Sub ParseFromRecordSet(ByRef tRecordSet As T_RECORD_SET, Optional index As Long = 0)
    With tRecordSet
        churchSid = .LISTDATA(index, 0)
        churchName = .LISTDATA(index, 1)
        If .LISTDATA(index, 2) <> "" Then
            startDate = .LISTDATA(index, 2)
        Else
            startDate = DateSerial(1900, 1, 1)
        End If
        If .LISTDATA(index, 3) <> "" Then
            endDate = .LISTDATA(index, 3)
        Else
            endDate = DateSerial(1900, 1, 1)
        End If
        lifeNo = .LISTDATA(index, 4)
        nameKo = .LISTDATA(index, 5)
        title = .LISTDATA(index, 6)
        position = .LISTDATA(index, 7)
    End With
End Sub