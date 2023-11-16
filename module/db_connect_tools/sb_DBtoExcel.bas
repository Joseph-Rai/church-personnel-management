Attribute VB_Name = "sb_DBtoExcel"
Option Explicit
'-----------------------------------------------
'  DB �ڷḦ ������ ��ȯ
'    - excel_export(���Ͽ��¿���)
'    - ��ũ���� ���� ����ȭ�鿡 ����
'    - ���Ͽ����� true�� ���� �� ����
'------------------------------------------------
Sub excel_export() 'Optional FileOpen As Boolean = False)
    
    Dim tableNM As String, dbNM As String
    Dim strSql As String
    Dim i As Integer
    Dim fileNM As String
    Dim fileSNM As String
    Dim FileOpen As Boolean
    FileOpen = False
    
    '//db�� ����
    tableNM = "op_system.db_churchlist_custom" '//db��.���̺�� - �����ڡ�
    dbNM = "op_system" '//�����ڡ�

    
    '//DB����
    connectTaskDB
    
    '//Select��
    strSql = "SELECT " & _
                    "CONCAT(a.church_nm, IFNULL(b.church_nm, REPLACE(a.church_nm,' ����ȸ', ''))) AS 'id' " & _
                    ",a.* " & _
                    ",IFNULL(b.church_nm, REPLACE(a.church_nm,' ����ȸ', '')) AS 'main_church_name' " & _
                "FROM op_system.db_churchlist_custom a " & _
                "LEFT JOIN op_system.db_churchlist b " & _
                    "ON a.main_church_cd = b.church_sid " & _
                "ORDER BY a.sort_order DESC"
    
    
    '//SQL�� �����ϰ� ��ȸ�� �ڷḦ ���ڵ�¿� ����
    Call callDBtoRS("excel_export", tableNM, strSql)
    If rs.EOF = True Then
        MsgBox "��ȸ ���ǿ� �´� �ڷᰡ �����ϴ�.", vbInformation, banner
        disconnectALL
        Exit Sub
    End If
    
    '//������ �ڷ� ��������
    Optimization
    Workbooks.Add
    For i = i To rs.Fields.Count - 1
        Cells(1, 1).Offset(0, i).Value = rs.Fields(i).Name
    Next i
    Cells(2, 1).CopyFromRecordset rs
    Cells(1.1).CurrentRegion.Columns.AutoFit
    fileNM = GetDesktopPath() & tableNM & "(" & Format(Date, "yyyymmdd") & "_" & Format(Time, "hhmm") & ")" & ".xlsx"
    Application.DisplayAlerts = False
        ActiveWorkbook.SaveAs FileName:=fileNM
    Application.DisplayAlerts = True
    If FileOpen = False Then
        ActiveWorkbook.Close
    End If
    Normal
    
    '//�������, ������
    fileSNM = Right(fileNM, Len(fileNM) - InStrRev(fileNM, "\"))
    MsgBox "����ȭ�鿡 ������ �����Ǿ����ϴ�." & vbNewLine & vbNewLine & _
        " - �����̸�: " & fileSNM, vbInformation, banner
    
    disconnectALL
End Sub

Sub Excel_Export_For_Atten_Update() 'Optional FileOpen As Boolean = False)
    
    Dim tableNM As String, dbNM As String
    Dim strSql As String
    Dim i As Integer
    Dim fileNM As String
    Dim fileSNM As String
    Dim FileOpen As Boolean
    FileOpen = False
    
    '//db�� ����
    tableNM = "op_system.db_churchlist_custom" '//db��.���̺�� - �����ڡ�
    dbNM = "op_system" '//�����ڡ�

    
    '//DB����
    connectTaskDB
    
    '//Select��
    strSql = "SELECT " & _
                    "CONCAT(a.church_nm, IFNULL(b.church_nm, REPLACE(a.church_nm,' ����ȸ', ''))) AS 'id' " & _
                    ",a.* " & _
                    ",IFNULL(b.church_nm, REPLACE(a.church_nm,' ����ȸ', '')) AS 'main_church_name' " & _
                "FROM op_system.db_churchlist_custom a " & _
                "LEFT JOIN op_system.db_churchlist b " & _
                    "ON a.main_church_cd = b.church_sid " & _
                "ORDER BY a.sort_order DESC"
    
    
    '//SQL�� �����ϰ� ��ȸ�� �ڷḦ ���ڵ�¿� ����
    Call callDBtoRS("excel_export", tableNM, strSql)
    If rs.EOF = True Then
        MsgBox "��ȸ ���ǿ� �´� �ڷᰡ �����ϴ�.", vbInformation, banner
        disconnectALL
        Exit Sub
    End If
    
    '//������ �ڷ� ��������
    Optimization
    Workbooks.Add
    For i = i To rs.Fields.Count - 1
        Cells(1, 1).Offset(0, i).Value = rs.Fields(i).Name
    Next i
    Cells(2, 1).CopyFromRecordset rs
    Cells(1.1).CurrentRegion.Columns.AutoFit
    fileNM = GetDesktopPath() & tableNM & "(" & Format(Date, "yyyymmdd") & "_" & Format(Time, "hhmm") & ")" & ".xlsx"
    Application.DisplayAlerts = False
        ActiveWorkbook.SaveAs FileName:=fileNM
    Application.DisplayAlerts = True
    If FileOpen = False Then
        ActiveWorkbook.Close
    End If
    Normal
    
    '//�������, ������
    fileSNM = Right(fileNM, Len(fileNM) - InStrRev(fileNM, "\"))
    MsgBox "����ȭ�鿡 ������ �����Ǿ����ϴ�." & vbNewLine & vbNewLine & _
        " - �����̸�: " & fileSNM, vbInformation, banner
    
    disconnectALL
End Sub

'-----------------------------------------------
'  DB����
'    - connectDB(���� IP, ��Ű��, ID, PW)
'-----------------------------------------------
Sub connectDB(argIP As String, argDB As String, argId As String, argPW As String)
    Set conn = New ADODB.Connection
    conn.ConnectionString = "Driver={" & ODBCDriver & "};Server=" & argIP & ";Port=3306;Database=" & argDB & ";User=" & argId & ";Password=" & argPW & ";Option=2;"
    conn.Open
End Sub

