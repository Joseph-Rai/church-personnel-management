Attribute VB_Name = "a_Common"
Option Explicit

Public Const programv As String = "V20201230" '���α׷� ���� �����ڡ�
Public Const banner As String = "�ؿܱ� ���հ��� ���α׷� " & programv '�ڡ�
Public Const ODBCDriver As String = "MariaDB ODBC 3.1 Driver" 'Client PC�� ��ġ�� ODBC Driver�ڡ�
Public Const IPAddress As String = "172.17.109.45" 'DB IP Address�ڡ�
Public Const commonDB As String = "common" 'Common DB��ڡ�
Public Const commonID As String = "common" 'Common DB UN�ڡ�
Public Const commonPW As String = "12345" 'Common DB ��й�ȣ�ڡ�
Public Const globalSheetPW As String = "12345" '��Ʈ��й�ȣ
Public Const updateDay As Integer = 7 '�ſ� ������Ʈ�ϴ� ��¥

Public Const SCHEMA_OPSYSTEM = "op_system." '��Ű����: op_system
Public Const SCHEME_COMMON = "common." '��Ű����: common

Public conn As ADODB.Connection 'ADO Connection ��ü ����
Public rs As New ADODB.RecordSet 'ADO Recordset ��ü ����
Public connIP As String, connDB As String, connUN As String, connPW As String 'Task DB ���� ����
Public USER_ID As Integer '������ڵ�
Public USER_GB As String '����ڱ���(SA, AM, MG, WP)
Public USER_NM As String '������̸�
Public USER_DEPT As String '����ںμ�
Public checkLogin As Integer '�α��� ���� 0: �α��� ����, 1 = �α���
Public today As Date '�޺��ڽ��� ��¥ �ɼ��� ����� ��� ���� ��¥�� ��ȸ
Public TASK_CODE As Integer '--//1. ��ȸ�߷�, 2. �����Ӹ�, 3. ��å�Ӹ�
Public SEARCH_CODE As Integer '--//1. ������ ���, 2. ��ȸ���, 3. ��ȸ�����
Public argShow As Integer '--//1. frm_Update_Appointment, 2. frm_Update_PInformation, 3. frm_Search_Appointment
                          '--//1. BCLeader, 2. FamilyInfo
Public argShow2 As Integer '--//1. ������ ���� ������Ʈ, 2. ����� ���� ������Ʈ
Public argShow3 As Integer '--//1. frm_Update_Appointment, 2. frm_Update_FamilyInfo, 3. frm_Search_Appointment
Public PICPATH_PSTAFF As String '--//�μ��� �������� ���
Public WB_ORIGIN As Workbook '--//�α��� �� ���繮�� �޾ƿ���

Public Const COUNT_PAGE_WIDTH_CELLS = 15 '--//��ȸ�� ��������Ȳ �������ʺ�
Public Const COUNT_PAGE_HEIGHT_CELLS = 37 '--//��ȸ�� ��������Ȳ ����������


'-------------------------------
'  Common DB����
'    - �α��� üũ
'    - TaskDB���� ���� ��ȯ
'-------------------------------
Sub connectCommonDB()
    connectDB IPAddress, commonDB, commonID, commonPW
End Sub

'-------------------
'  Task DB����
'-------------------
Sub connectTaskDB()
    connectDB connIP, connDB, connUN, connPW
End Sub

'-----------------------------------------------
'  DB���� ���ν���
'    - connectDB(���� IP, ��Ű��, ID, PW)
'-----------------------------------------------
Sub connectDB(argIP As String, argDB As String, argId As String, argPW As String)
    
    If argIP = "" Then
        MsgBox "�α��� �Ǿ� ���� �ʽ��ϴ�.", vbCritical + vbOKOnly, banner
        End
    End If
    
    Set conn = New ADODB.Connection
    conn.ConnectionString = "Driver={" & ODBCDriver & "};Server=" & argIP & ";Port=3306;Database=" & argDB & ";User=" & argId & ";Password=" & argPW & ";Option=2;"
    conn.Open
End Sub

'---------------------------------------------------------------------
'  ���ڵ�� ���� �� ������ ��ȯ
'    - calDBtoRS(���ν�����, ���̺��, SQL��, ���̸�, ���̸�)
'    - �����߻� �� ���� �ڵ鸵 �� �α� ���
'    - �����߻� ���ϸ� �� ���� ���ν������� �α� ���(�ʿ� ��)
'---------------------------------------------------------------------
Sub callDBtoRS(procedureNM As String, tableNM As String, SQLScript As String, Optional formNM As String = "NULL", Optional jobNM As String = "������ ��ȸ")
On Error GoTo ErrHandler
    Set rs = New ADODB.RecordSet
    rs.CursorLocation = adUseClient
    rs.Open Source:=SQLScript, ActiveConnection:=conn, CursorType:=adOpenForwardOnly, LockType:=adLockReadOnly, Options:=adCmdText
    Exit Sub
ErrHandler:
    ErrHandle procedureNM, tableNM, SQLScript, formNM, jobNM
    writeLog procedureNM, tableNM, Replace(SQLScript, ";", ""), 1, formNM, jobNM '//�����ڵ� 1
End Sub

'-------------------------------------------------------------------------------------
'  SQL���� �����ϰ� �αױ��
'    - executeSqlWithLog(SQL��, ���ν�����, ���̺��, ���̸�(�ɼ�))
'    - �۾����� �� �αױ��
'    - �����߻� �� executeSQL ���� �� �αױ��
'-------------------------------------------------------------------------------------
Public Sub executeSqlWithLog(sql As String, procedureNM As String, tableNM As String, Optional jobNM As String = "NULL")
    Dim result As T_RESULT
    
    '--//DB�� ����
    connectTaskDB
    
    result.strSql = sql
    result.affectedCount = executeSQL(procedureNM, tableNM, sql, , jobNM)
    writeLog procedureNM, tableNM, Replace(sql, ";", ""), 0, , jobNM, result.affectedCount
    
    '--//DB���� ����
    disconnectALL
End Sub

'-------------------------------------------------------------------------------------
'  SQL���� �����ϰ� ������ ������ ���� ���ڵ� ���� ��ȯ
'    - executeSQL(���ν�����, ���̺��, SQL��, ���̸�(�ɼ�), ���̸�(�ɼ�))
'    - SQL�� ���� ��� ���� ���θ� �˱� ���� ���� ���� ���ڵ� �� ����
'    - �����߻� �� ���� �ڵ鸵 �� �α� ���
'    - �����߻� ���ϸ� �� ���� ���ν������� �α� ���
'-------------------------------------------------------------------------------------
Public Function executeSQL(procedureNM As String, tableNM As String, SQLScript As String, Optional formNM As String = "NULL", Optional jobNM As String = "��Ÿ") As Long
On Error GoTo ErrHandler
    Dim affectedCount As Long
    Dim sql As Variant
    
    For Each sql In Split(SQLScript, ";")
        If Len(sql) > 0 Then
            Debug.Print sql
            conn.Execute CommandText:=sql, recordsaffected:=affectedCount
        End If
    Next
    
    executeSQL = affectedCount
    Exit Function
ErrHandler:
    ErrHandle procedureNM, tableNM, SQLScript, formNM, jobNM
    writeLog procedureNM, tableNM, Replace(SQLScript, ";", ""), 1, formNM, jobNM '//�����ڵ� 1
End Function

'--------------------------
'  DB �� RS ���� ����
'--------------------------
Sub disconnectRS()
    On Error Resume Next
        rs.Close
        Set rs = Nothing
    On Error GoTo 0
End Sub
Sub disconnectDB()
    On Error Resume Next
        conn.Close
        Set conn = Nothing
    On Error GoTo 0
End Sub
Sub disconnectALL()
    On Error Resume Next
        rs.Close
        Set rs = Nothing
        conn.Close
        Set conn = Nothing
    On Error GoTo 0
End Sub

'------------------------------------------------
'  SQL ���ϸ�Ī �˻��� ó��('%�˻���%')
'------------------------------------------------
Public Function PText(argString As Variant) As String
    If argString = "" Or Len(argString) = 0 Then
        PText = "'%%'"
    Else
        PText = "'%" & Trim(Replace(Replace(argString, "%", "\%"), "'", "''")) & "%'"
    End If
End Function

'---------------------------------------------
'  SQL ��Į���Ī �˻��� ó��('�˻���')
'---------------------------------------------
Public Function SText(argString As Variant) As String
    If argString = "" Or Len(argString) = 0 Then
        SText = "''"
    Else
        SText = "'" & Trim(Replace(Replace(argString, "%", "\%"), "'", "''")) & "'"
    End If
End Function

'--------------------
'  ��ũ�� ����ȭ
'--------------------
Sub Optimization()
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
'        .Calculation = xlCalculationManual
    End With
End Sub

'-------------------------
'  ��ũ�� ����ȭ ����
'-------------------------
Sub Normal()
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
'        .Calculation = xlCalculationAutomatic
    End With
End Sub

'----------------
'  ��üȭ��On
'----------------
Sub FullscreenOn()
    ActiveWindow.DisplayHeadings = False
    ActiveWindow.DisplayWorkbookTabs = False
    Application.DisplayFullScreen = True
    Application.DisplayFormulaBar = False
End Sub

'----------------
'  ��üȭ��Off
'----------------
Sub FullscreenOff()
    ActiveWindow.DisplayHeadings = True
    ActiveWindow.DisplayWorkbookTabs = True
    Application.DisplayFullScreen = False
    Application.DisplayFormulaBar = True
End Sub

'---------------------
'  ����ȭ�� �����
'---------------------
Sub HideExcel()
    Application.Visible = False
End Sub

'---------------------
'  ����ȭ�� ���̱�
'---------------------
Sub ShowExcel()
    Application.Visible = True
End Sub
Public Function GetDesktopPath(Optional BackSlash As Boolean = True)
    Dim oWSHShell As Object
    
    Set oWSHShell = CreateObject("WScript.Shell")
    If BackSlash = True Then
        GetDesktopPath = oWSHShell.SpecialFolders("Desktop") & "\"
    Else
        GetDesktopPath = oWSHShell.SpecialFolders("Desktop")
    End If
    
    Set oWSHShell = Nothing
End Function
