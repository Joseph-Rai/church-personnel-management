Attribute VB_Name = "fn_GetPosition2Joining"
Option Explicit

Public Function getPosition2Joining()

    Dim strQuery As String
    Dim tResultSet As T_RECORD_SET
    
    strQuery = "SELECT * FROM op_system.a_position2;"
    tResultSet = makeListData(strQuery, "op_system.a_position2")
        
    Dim result As String
    Dim i As Integer
    With tResultSet
        For i = 0 To .CNT_RECORD - 1
            If i < .CNT_RECORD - 1 Then
                result = result & "'" & .LISTDATA(i, 0) & "', "
            Else
                result = result & "'" & .LISTDATA(i, 0) & "'"
            End If
        Next
    End With
    
    getPosition2Joining = result

End Function
