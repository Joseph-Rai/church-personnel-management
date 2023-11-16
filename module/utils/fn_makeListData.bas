Attribute VB_Name = "fn_makeListData"
Option Explicit

Public Function makeListData(ByVal strSql As String, ByVal tableNM As String) As T_RECORD_SET

    Dim tRecordSet As T_RECORD_SET
    Dim LISTDATA() As String '--//DB���� �޾ƿ� rs�� �迭�� ����
    Dim LISTFIELD() As String '--//DB���� �޾ƿ� rs�� �ʵ带 �迭�� ����
    Dim cntRecord As Long '--//DB���� �޾ƿ� ���ڵ��� ����
    Dim i As Long, j As Long
    
    '//DB���� �ڷ� ȣ���Ͽ� ���ڵ�¿� ��ȯ
    connectTaskDB
    callDBtoRS "makeListData", tableNM, strSql
    
    '//���ڵ���� �����͸� listData �迭�� ��ȯ
    If Not rs.EOF Then
        ReDim LISTDATA(0 To rs.RecordCount - 1, 0 To rs.Fields.Count - 1) '//DB���� ��ȯ�� �迭�� ũ�� ����: ���ڵ���� ���ڵ� ��, �ʵ� ��
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
    
    '--//�ʵ�� �迭 ä���
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
