VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PassportPhoto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'##########################################
'TABLE: DB_PASSPORT_PHOTO�� ���� CLASS
'##########################################

Dim vLifeNo As String
Dim vPhoto As String

Public Property Get lifeNo() As String
    lifeNo = vLifeNo
End Property

Public Property Let lifeNo(argLifeNo As String)
    vLifeNo = argLifeNo
End Property

Public Property Get photo() As String
    photo = vPhoto
End Property

Public Property Let photo(argPhoto As String)
    vPhoto = argPhoto
End Property

'@param tRecordSet: DB���� ���� RecordSet Ÿ���� ������ �޾� ��ü �Ӽ��� ���� ä��ϴ�.
Public Sub ParseFromRecordSet(ByRef tRecordSet As T_RECORD_SET, Optional index As Long = 0)
    With tRecordSet
        vLifeNo = .LISTDATA(index, 0)
        vPhoto = .LISTDATA(index, 1)
    End With
End Sub