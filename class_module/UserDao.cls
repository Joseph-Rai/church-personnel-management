VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UserDao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'@param argUser: ������ ������ �ִ� ���Ѹ���� �����ɴϴ�.
Public Function GetUserAuthorities() As Object

    Dim authList As Object
    Dim sql As String
    Dim tRecordSet As T_RECORD_SET
    
    Set authList = CreateObject("System.Collections.ArrayList")
    
    sql = _
        "SELECT " & _
        "    AuthList.authority " & _
        "FROM " & TABLE_AUTH_TABLE & " auth " & _
        "LEFT JOIN " & TABLE_AUTHORITY & " authList " & _
        "    ON auth.authority_id = authList.id " & _
        "WHERE auth.user_id = " & SText(USER_ID) & ";"
    
    tRecordSet = makeListData(sql, TABLE_AUTH_TABLE)
    
    Dim strAuth As Variant
    If tRecordSet.CNT_RECORD > 0 Then
        For Each strAuth In tRecordSet.LISTDATA
            authList.Add strAuth
        Next
    End If
    
    Set GetUserAuthorities = authList

End Function


'@param id: ���� ���̵� �������� DB���� �˻��մϴ�.
Public Function FindById(ByRef Id As String) As User

    Dim sql As String
    Dim tRecordSet As T_RECORD_SET
    
    sql = "SELECT * FROM " & TABLE_USER & _
                    " a WHERE a.user_id = " & SText(Id) & ";"
    
    tRecordSet = makeListData(sql, TABLE_USER)
    
    Dim objUser As New User
    If tRecordSet.CNT_RECORD > 0 Then
        objUser.ParseFromRecordSet tRecordSet
    Else
        Set objUser = Nothing
    End If
    
    Set FindById = objUser
    
End Function

'@param user: user�� ��ü�� �޾� ���̵� �������� DB���� �˻��մϴ�.
Public Function FindByUser(ByRef argUser As User) As User
    FindByUser = FindById(argUser.Id)
End Function

'@param name: �������� �������� DB���� �˻��մϴ�.
Public Function FindByName(ByVal Name As String) As User

    Dim sql As String
    Dim tRecordSet As T_RECORD_SET
    
    sql = "SELECT * FROM " & TABLE_USER & _
                    " a WHERE a.name = " & SText(Name) & ";"
    
    tRecordSet = makeListData(sql, TABLE_USER)
    
    Dim objUser As New User
    If tRecordSet.CNT_RECORD > 0 Then
        objUser.ParseFromRecordSet tRecordSet
    Else
        Set objUser = Nothing
    End If
    
    Set FindByName = objUser

End Function
