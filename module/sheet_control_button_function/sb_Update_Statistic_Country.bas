Attribute VB_Name = "sb_Update_Statistic_Country"
Option Explicit
Dim rngStart As Range '--//����Ʈ ���ۼ�
Dim rngEnd As Range '--//����Ʈ ���Ἷ
Dim LISTDATA() As String '--//DB���� �޾ƿ� rs�� �迭�� ����
Dim LISTFIELD() As String '--//DB���� �޾ƿ� rs�� �ʵ带 �迭�� ����
Dim cntRecord As Integer '--//DB���� �޾ƿ� ���ڵ��� ����
Dim strSql As String '--//SQL ������
Const TB1 As String = "op_system.v_statistic_by_country"
Sub Sheet_Init()

    Dim R As Long
    Dim i As Long

    If connIP = Empty Then setGlobalVariant '//�������� �缳��
    If checkLogin = 0 Then f_login.Show '//�α���üũ
    
    '--//��������
    Set rngStart = Range("A3")
    Set rngEnd = Range("A:A").Find("�հ�", lookat:=xlWhole)
    
    '--//���� ������ ��� ����
    If rngEnd.Row - rngStart.Row > 2 Then
        Range(Cells(rngStart.Row + 2, "A"), Cells(rngEnd.Row - 1, "A")).EntireRow.Delete
    End If
    
    '--//������ �ҷ�����
    strSql = "SELECT * FROM " & TB1 & ";"
    Call makeListData(strSql, TB1)
    
    '--//������ ����ŭ ���߰�
    rngStart.Offset(2).Resize(cntRecord - 1).EntireRow.Insert Shift:=xlDown
    
    '--//�ҷ��� ������ ����
    rngStart.Offset(1, 1).Resize(cntRecord, UBound(LISTFIELD) + 1) = LISTDATA
    
    '--//���ĺ���
    rngStart.Offset(1).EntireRow.Copy
    rngStart.Offset(1).Resize(cntRecord).EntireRow.PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
    
    '--//A�� ��ȣ ä���
    rngStart.Offset(1).Resize(cntRecord).Formula = "=ROW()-3"
    
    '--//�ؽ�Ʈ ������ ���� ������������ �ٲٱ�
    Range("A2").Copy
    rngStart.Offset(1, 2).Resize(cntRecord, UBound(LISTFIELD) + 1).PasteSpecial Paste:=xlPasteValues, operation:=xlPasteSpecialOperationAdd
    Application.CutCopyMode = False
    
    '--//0�� ������ �ʰ� ó��
    rngStart.Offset(1, 2).Resize(cntRecord, UBound(LISTFIELD) + 1).Replace 0, "", lookat:=xlWhole
    
    '--//�հ� �� ����ó��
    Set rngEnd = Range("A:A").Find("�հ�", lookat:=xlWhole)
    rngEnd.Offset(, 2).Resize(, UBound(LISTFIELD) + 1).FormulaR1C1 = "=SUM(R[" & rngStart.Row - rngEnd.Row + 1 & "]C:R[-1]C)"
    
    '--//rngStart ���� �� ����
    rngStart.Select
    
End Sub
Private Sub makeListData(ByVal strSql As String, ByVal tableNM As String)

    Dim i As Integer, j As Integer
    
    '//DB���� �ڷ� ȣ���Ͽ� ���ڵ�¿� ��ȯ
    connectTaskDB
    callDBtoRS "makeListData", tableNM, strSql, "sb_Update_Statistic_Country"
    
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
    cntRecord = rs.RecordCount '--//���ڵ� �� ����
    disconnectALL
    
End Sub
