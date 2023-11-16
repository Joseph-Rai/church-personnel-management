Attribute VB_Name = "fn_makeListData"
Option Explicit

Public Function makeListData(ByVal strSql As String, ByVal tableNM As String) As T_RECORD_SET

    Dim tRecordSet As T_RECORD_SET
    Dim LISTDATA() As String '--//DB에서 받아온 rs를 배열로 저장
    Dim LISTFIELD() As String '--//DB에서 받아온 rs의 필드를 배열로 저장
    Dim cntRecord As Long '--//DB에서 받아온 레코드의 개수
    Dim i As Long, j As Long
    
    '//DB에서 자료 호출하여 레코드셋에 반환
    connectTaskDB
    callDBtoRS "makeListData", tableNM, strSql
    
    '//레코드셋의 데이터를 listData 배열에 반환
    If Not rs.EOF Then
        ReDim LISTDATA(0 To rs.RecordCount - 1, 0 To rs.Fields.Count - 1) '//DB에서 반환할 배열의 크기 지정: 레코드셋의 레코드 수, 필드 수
        rs.MoveFirst
        For i = 0 To rs.RecordCount - 1
            For j = 0 To rs.Fields.Count - 1
                If IsNull(rs.Fields(j).Value) = True Then
                    LISTDATA(i, j) = ""
                Else
                    LISTDATA(i, j) = rs.Fields(j).Value
                End If
            Next j
            rs.MoveNext
        Next i
    End If
    
    '--//필드명 배열 채우기
    ReDim LISTFIELD(0 To rs.Fields.Count - 1)
    For i = 0 To rs.Fields.Count - 1
        LISTFIELD(i) = rs.Fields(i).Name
    Next i
    
    cntRecord = rs.RecordCount
    
    disconnectALL
    
    tRecordSet.LISTDATA = LISTDATA
    tRecordSet.LISTFIELD = LISTFIELD
    tRecordSet.CNT_RECORD = cntRecord
    
    makeListData = tRecordSet
    
End Function
