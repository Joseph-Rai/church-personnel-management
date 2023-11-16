VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Search_Statistic 
   Caption         =   "������ ��赥���� ��ȸ"
   ClientHeight    =   1470
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3210
   OleObjectBlob   =   "frm_Search_Statistic.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "frm_Search_Statistic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cntlst1Col As Integer '//����Ʈ �÷� ��
Dim TB1 As String, TB2 As String, TB3 As String, TB4 As String '//���� ����� DB ���̺�
Const strOrderBy As String = "currency_un ASC" '//DB���� sort_order �ʵ�
Dim caseSave As Integer '//1: �߰�, 2: ����, 3: ����(SUSPEND), 4: ��������
Dim queryKey As Integer '//����Ʈ ��ġ ��ȯ�� ���� id
Dim LISTDATA() As String '--//DB���� �޾ƿ� rs�� �迭�� ����
Dim LISTFIELD() As String '--//DB���� �޾ƿ� rs�� �ʵ带 �迭�� ����
Dim cntRecord As Integer '--//DB���� �޾ƿ� ���ڵ��� ����
Dim strSql As String '--//SQL ������
Dim txtBox_Focus As MSForms.control
Dim rngA As Range, rngB As Range, rngC As Range
Dim ws As Worksheet

Private Sub cboYear_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Un'HookListBoxScroll
End Sub

Private Sub cboYear_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Me.cboYear.ListCount > 0 Then
        'HookListBoxScroll Me, Me.cboYear
    End If
End Sub
Private Sub cboMonth_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Un'HookListBoxScroll
End Sub

Private Sub cboMonth_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Me.cboMonth.ListCount > 0 Then
        'HookListBoxScroll Me, Me.cboMonth
    End If
End Sub

Private Sub UserForm_Initialize()
    
    Dim i As Long
    
    If connIP = Empty Then setGlobalVariant '//�������� �缳��
    If checkLogin = 0 Then f_login.Show '//�α���üũ
    If checkLogin = 0 Then End '--//�α��� ���� �� ���ν��� ����
    
    '--//��Ʈ����
    Select Case SEARCH_CODE
    Case 1
        Set ws = WB_ORIGIN.Sheets("������ ���") '--//������ ��� ��Ʈ
    Case 2
        Set ws = WB_ORIGIN.Sheets("��ȸ���") '--//��ȸ��� ��Ʈ
    Case 3
        Set ws = WB_ORIGIN.Sheets("��ȸ�����") '--//��ȸ����� ��Ʈ
    Case 4
        Set ws = WB_ORIGIN.Sheets("��ȸ����") '--//��ȸ���� ��Ʈ
    Case Else
    End Select
    
    '--//���ʼ���
    Me.cmdClose.Cancel = True
    TB1 = "op_system.temp_statistic_by_country" '--//���������
    TB2 = "op_system.temp_statistic_by_church" '--//��ȸ�����
    TB3 = "op_system.temp_statistic_by_pstaff" '--//��ȸ�� ��ȸ�����
    TB4 = "op_system.temp_statistic_by_church_all" '--//��ȸ������
    
    '--//�޺��ڽ� �� �߰�
    For i = 1 To 12
        With Me.cboMonth
            .AddItem i
        End With
    Next
    
    '--//�޺��ڽ� �⵵ �߰�
    For i = year(Date) To year(Date) - 10 Step -1
        With Me.cboYear
            .AddItem i
        End With
    Next
    
    '--//�޺��ڽ� �ֽŵ����� ��¥����
    Me.cboYear = IIf(Day(Date) < 10, year(DateAdd("m", -2, Date)), year(DateAdd("m", -1, Date)))
    Me.cboMonth = IIf(Day(Date) < 10, month(DateAdd("m", -2, Date)), month(DateAdd("m", -1, Date)))
    
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer, j As Integer
    Dim result As T_RESULT
    
    '--//��Ʈ Ȱ��ȭ �� �������
    WB_ORIGIN.Activate
    ws.Activate
    Call shUnprotect(globalSheetPW)
    
    Select Case SEARCH_CODE
    Case 1
        WB_ORIGIN.Activate
        Sheets("������ ���").Activate
        Optimization
        '--//���� ������ ����
        Call initializeRepport
        
        '--//temp_churchlist_by_time ���̺� ������Ʈ
        strSql = "CALL `Routine_churchlist_by_time`(" & SText(DateSerial(Me.cboYear, Me.cboMonth, 1)) & "," & SText(USER_DEPT) & ");"
        connectTaskDB
        result.strSql = strSql
        result.affectedCount = executeSQL("cmdADD_Clikc", "temp_pstaff_by_time", strSql, Me.Name, "�ӽ���ȸ ���̺� ����")
    '    writeLog "cmdADD_Click", "temp_pstaff_by_time", strSQL, 0, Me.Name, "�ӽ���ȸ ���̺� ����", result.affectedCount
        disconnectALL
        
        '--//temp_pstaff_by_time ���̺� ������Ʈ
        strSql = "CALL `Routine_pstaff_by_time`(" & SText(DateSerial(Me.cboYear, Me.cboMonth, 1)) & "," & SText(USER_DEPT) & ");"
        connectTaskDB
        result.strSql = strSql
        result.affectedCount = executeSQL("cmdADD_Clikc", "temp_pstaff_by_time", strSql, Me.Name, "�ӽ���ȸ ���̺� ����")
    '    writeLog "cmdADD_Click", "temp_pstaff_by_time", strSQL, 0, Me.Name, "�ӽ���ȸ ���̺� ����", result.affectedCount
        disconnectALL
        
        '--//temp_statistic_by_country ���̺� ������Ʈ
        strSql = "CALL `Routine_statistic_by_Country`(" & SText(DateSerial(Me.cboYear, Me.cboMonth, 1)) & ", " & SText(USER_DEPT) & ");"
        connectTaskDB
        result.strSql = strSql
        result.affectedCount = executeSQL("cmdADD_Clikc", "temp_statistic_by_country", strSql, Me.Name, "�ӽ���ȸ ���̺� ����")
    '    writeLog "cmdADD_Click", "temp_pstaff_by_time", strSQL, 0, Me.Name, "�ӽ���ȸ ���̺� ����", result.affectedCount
        disconnectALL
        
        '--//SQL��
        strSql = makeSelectSQL(TB1)
        
        '--//DB���� �ڷ� ȣ���Ͽ� ���ڵ�� ��ȯ
        Call makeListData(strSql, TB1)
        
        '--//����Ʈ ���˼���
        Call makeReport
        
        '--//��ȯ�� ListData�� ���� ��Ʈ�� ����
        If cntRecord > 0 Then
            Range("Stat_Country_Start").Resize(, UBound(LISTFIELD) + 1) = LISTFIELD
            Range("Stat_Country_Start").Offset(1).Resize(cntRecord, UBound(LISTFIELD) + 1) = LISTDATA
        End If
        Range("A2").Copy
        Range(Cells(4, "C"), Cells(4 + cntRecord - 1, "AG")).PasteSpecial Paste:=xlPasteValues, operation:=xlPasteSpecialOperationAdd
        Application.CutCopyMode = False
        
        '--//���� �ʱ�ȭ
        Call sbClearVariant
        
        '--//��ȸ������ ���
        Range("Stat_Country_Date") = DateSerial(Me.cboYear, Me.cboMonth, 1)
        
        strSql = "SELECT a.dept_nm FROM op_system.db_ovs_dept a WHERE a.dept_id = " & SText(USER_DEPT)
        Call makeListData(strSql, "op_system.db_ovs_dept")
        
        Sheets("������ ���").Range("A1").FormulaR1C1 = _
                            "=""" & LISTDATA(0, 0) & " ������ �⼮��Ȳ �� ��ȸ�� ���ǥ [""&TEXT(Stat_Country_Date,""yyyy�� mm��"")&""]"""
        Call sbClearVariant
        
        Range("A2").Select
        Normal
    Case 2
        WB_ORIGIN.Activate
        Sheets("��ȸ���").Activate
        Optimization
        '--//���� ������ ����
        Call initializeRepport
        
        '--//temp_churchlist_by_time ���̺� ������Ʈ
        strSql = "CALL `Routine_churchlist_by_time`(" & SText(DateSerial(Me.cboYear, Me.cboMonth, 1)) & "," & SText(USER_DEPT) & ");"
        connectTaskDB
        result.strSql = strSql
        result.affectedCount = executeSQL("cmdADD_Clikc", "temp_pstaff_by_time", strSql, Me.Name, "�ӽ���ȸ ���̺� ����")
    '    writeLog "cmdADD_Click", "temp_pstaff_by_time", strSQL, 0, Me.Name, "�ӽ���ȸ ���̺� ����", result.affectedCount
        disconnectALL
        
        '--//temp_pstaff_by_time ���̺� ������Ʈ
        strSql = "CALL `Routine_pstaff_by_time`(" & SText(Format(WorksheetFunction.EoMonth(DateSerial(Me.cboYear, Me.cboMonth, 1), 0), "yyyy-mm-dd")) & ", " & SText(USER_DEPT) & ");"
        connectTaskDB
        result.strSql = strSql
        result.affectedCount = executeSQL("cmdADD_Clikc", "temp_pstaff_by_time", strSql, Me.Name, "�ӽ���ȸ ���̺� ����")
    '    writeLog "cmdADD_Click", "temp_pstaff_by_time", strSQL, 0, Me.Name, "�ӽ���ȸ ���̺� ����", result.affectedCount
        disconnectALL
        
        '--//temp_statistic_by_church ���̺� ������Ʈ
        strSql = "CALL `Routine_statistic_by_Church`(" & SText(DateSerial(Me.cboYear, Me.cboMonth, 1)) & ");"
        connectTaskDB
        result.strSql = strSql
        result.affectedCount = executeSQL("cmdADD_Clikc", "temp_statistic_by_country", strSql, Me.Name, "�ӽ���ȸ ���̺� ����")
    '    writeLog "cmdADD_Click", "temp_pstaff_by_time", strSQL, 0, Me.Name, "�ӽ���ȸ ���̺� ����", result.affectedCount
        disconnectALL
        
        '--//SQL��
        strSql = makeSelectSQL(TB2)
        
        '--//DB���� �ڷ� ȣ���Ͽ� ���ڵ�� ��ȯ
        Call makeListData(strSql, TB2)
        
        '--//����Ʈ ���˼���
        Call makeReport
        
        '--//��ȯ�� ListData�� ���� ��Ʈ�� ����
        If cntRecord > 0 Then
            Range("Stat_Church_Start").Resize(, UBound(LISTFIELD) + 1) = LISTFIELD
            Range("Stat_Church_Start").Offset(1).Resize(cntRecord, UBound(LISTFIELD) + 1) = LISTDATA
        End If
        Range("A2").Copy
        Range(Cells(4, "F"), Cells(4 + cntRecord - 1, "AF")).PasteSpecial Paste:=xlPasteValues, operation:=xlPasteSpecialOperationAdd
        Application.CutCopyMode = False
        
        '--//����ȸ ������ ����
'        Select Case USER_DEPT
'        Case 10
'            ActiveWorkbook.Worksheets("��ȸ���").Sort.SortFields.Clear
'            ActiveWorkbook.Worksheets("��ȸ���").Sort.SortFields.Add Key:=Range("Stat_Church_Start").Offset(1, 1).Resize(cntRecord), _
'                SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder:= _
'                "�ߵ�,īƮ����,���ȵ���,�����ߺ�,���ȼ���,���ε�,�ε����̳�", DataOption:=xlSortNormal
'            With ActiveWorkbook.Worksheets("��ȸ���").Sort
'                .SetRange Range("Stat_Church_Start").Offset(1, -1).Resize(cntRecord, UBound(listField) + 2)
'                .Header = xlGuess
'                .MatchCase = False
'                .Orientation = xlTopToBottom
'                .SortMethod = xlPinYin
'                .Apply
'            End With
'        Case Else
'        End Select
        
        '--//���� �ʱ�ȭ
        Call sbClearVariant
        
        '--//��ȸ������ ���
        Range("Stat_Church_Date") = DateSerial(Me.cboYear, Me.cboMonth, 1)
        
        strSql = "SELECT a.dept_nm FROM op_system.db_ovs_dept a WHERE a.dept_id = " & SText(USER_DEPT)
        Call makeListData(strSql, "op_system.db_ovs_dept")
        
        Sheets("��ȸ���").Range("A1").FormulaR1C1 = _
                        "=""" & LISTDATA(0, 0) & " ��ȸ�� �⼮��Ȳ �� ��ȸ�� ���ǥ  [""&TEXT(Stat_Church_Date,""yyyy�� mm��"")&"" ����]"""
        Call sbClearVariant
        
        Range("A2").Select
        Normal
    Case 3
        WB_ORIGIN.Activate
        Sheets("��ȸ�����").Activate
        Optimization
        '--//���� ������ ����
        Call initializeRepport
        
        '--//temp_churchlist_by_time ���̺� ������Ʈ
        strSql = "CALL `Routine_churchlist_by_time`(" & SText(DateSerial(Me.cboYear, Me.cboMonth, 1)) & "," & SText(USER_DEPT) & ");"
        connectTaskDB
        result.strSql = strSql
        result.affectedCount = executeSQL("cmdADD_Clikc", "temp_pstaff_by_time", strSql, Me.Name, "�ӽ���ȸ ���̺� ����")
    '    writeLog "cmdADD_Click", "temp_pstaff_by_time", strSQL, 0, Me.Name, "�ӽ���ȸ ���̺� ����", result.affectedCount
        disconnectALL
        
        '--//temp_pstaff_by_time ���̺� ������Ʈ
        strSql = "CALL `Routine_pstaff_by_time`(" & SText(Format(DateSerial(Me.cboYear, Me.cboMonth, 1), "yyyy-mm-dd")) & ", " & SText(USER_DEPT) & ");"
        connectTaskDB
        result.strSql = strSql
        result.affectedCount = executeSQL("cmdADD_Clikc", "temp_pstaff_by_time", strSql, Me.Name, "�ӽ���ȸ ���̺� ����")
    '    writeLog "cmdADD_Click", "temp_pstaff_by_time", strSQL, 0, Me.Name, "�ӽ���ȸ ���̺� ����", result.affectedCount
        disconnectALL
        
        '--//temp_statistic_by_pstaff ���̺� ������Ʈ
        strSql = "CALL `Routine_statistic_by_pstaff`(" & SText(DateSerial(Me.cboYear, Me.cboMonth, 1)) & ");"
        connectTaskDB
        result.strSql = strSql
        result.affectedCount = executeSQL("cmdADD_Clikc", "temp_statistic_by_country", strSql, Me.Name, "�ӽ���ȸ ���̺� ����")
    '    writeLog "cmdADD_Click", "temp_pstaff_by_time", strSQL, 0, Me.Name, "�ӽ���ȸ ���̺� ����", result.affectedCount
        disconnectALL
        
        '--//SQL��
        strSql = makeSelectSQL(TB3)
        
        '--//DB���� �ڷ� ȣ���Ͽ� ���ڵ�� ��ȯ
        Call makeListData(strSql, TB3)
        
        '--//����Ʈ ���˼���
        Call makeReport
        
        '--//��ȯ�� ListData�� ���� ��Ʈ�� ����
        If cntRecord > 0 Then
            Range("Stat_PStaff_Start").Resize(, UBound(LISTFIELD) + 1) = LISTFIELD
            Range("Stat_PStaff_Start").Offset(1).Resize(cntRecord, UBound(LISTFIELD) + 1) = LISTDATA
        End If
        Range("A2").Copy
        Range(Cells(5, "F"), Cells(5 + cntRecord - 1, "S")).PasteSpecial Paste:=xlPasteValues, operation:=xlPasteSpecialOperationAdd
        Application.CutCopyMode = False
        
        '--//����ȸ ������ ����
'        Select Case USER_DEPT
'        Case 10
'            ActiveWorkbook.Worksheets("��ȸ�����").Sort.SortFields.Clear
'            ActiveWorkbook.Worksheets("��ȸ�����").Sort.SortFields.Add Key:=Range("Stat_PStaff_Start").Offset(1, 1).Resize(cntRecord), _
'                SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder:= _
'                "�ߵ�,īƮ����,���ȵ���,�����ߺ�,���ȼ���,���ε�,�ε����̳�", DataOption:=xlSortNormal
'            With ActiveWorkbook.Worksheets("��ȸ�����").Sort
'                .SetRange Range("Stat_PStaff_Start").Offset(1, -1).Resize(cntRecord, UBound(listField) + 2)
'                .Header = xlGuess
'                .MatchCase = False
'                .Orientation = xlTopToBottom
'                .SortMethod = xlPinYin
'                .Apply
'            End With
'        Case Else
'        End Select
        
        '--//���� �ʱ�ȭ
        Call sbClearVariant
        
        '--//��ȸ������ ���
        Range("Stat_PStaff_Date") = DateSerial(Me.cboYear, Me.cboMonth, 1)
        
        strSql = "SELECT a.dept_nm FROM op_system.db_ovs_dept a WHERE a.dept_id = " & SText(USER_DEPT)
        Call makeListData(strSql, "op_system.db_ovs_dept")
        
        Sheets("��ȸ�����").Range("A1").FormulaR1C1 = _
                        "=""" & LISTDATA(0, 0) & " ��ȸ�� ��ȸ�� ���ǥ [""&TEXT(Stat_PStaff_Date,""yyyy�� mm��"")&"" ����]"""
        Call sbClearVariant
        
        Range("A2").Select
        Normal
    Case 4
        WB_ORIGIN.Activate
        Sheets("��ȸ����").Activate
        Optimization
        '--//���� ������ ����
        Call initializeRepport
        
        '--//temp_churchlist_by_time ���̺� ������Ʈ
        strSql = "CALL `Routine_churchlist_by_time`(" & SText(Format(Application.EoMonth(DateSerial(Me.cboYear, Me.cboMonth, 1), 0), "yyyy-mm-dd")) & "," & SText(USER_DEPT) & ");"
        connectTaskDB
        result.strSql = strSql
        result.affectedCount = executeSQL("cmdADD_Clikc", "temp_pstaff_by_time", strSql, Me.Name, "�ӽ���ȸ ���̺� ����")
    '    writeLog "cmdADD_Click", "temp_pstaff_by_time", strSQL, 0, Me.Name, "�ӽ���ȸ ���̺� ����", result.affectedCount
        disconnectALL
        
        '--//temp_statistic_by_church ���̺� ������Ʈ
        strSql = "CALL `Routine_statistic_by_Church_all`(" & SText(DateSerial(Me.cboYear, Me.cboMonth, 1)) & ");"
        connectTaskDB
        result.strSql = strSql
        result.affectedCount = executeSQL("cmdADD_Clikc", "temp_statistic_by_country", strSql, Me.Name, "�ӽ���ȸ ���̺� ����")
    '    writeLog "cmdADD_Click", "temp_pstaff_by_time", strSQL, 0, Me.Name, "�ӽ���ȸ ���̺� ����", result.affectedCount
        disconnectALL
        
        '--//SQL��
        strSql = makeSelectSQL(TB4)
        
        '--//DB���� �ڷ� ȣ���Ͽ� ���ڵ�� ��ȯ
        Call makeListData(strSql, TB4)
        
        '--//����Ʈ ���˼���
        Call makeReport
        
        '--//��ȯ�� ListData�� ���� ��Ʈ�� ����
        If cntRecord > 0 Then
            Range("Stat_ChurchDetail_Start").Resize(, UBound(LISTFIELD) + 1) = LISTFIELD
            Range("Stat_ChurchDetail_Start").Offset(1).Resize(cntRecord, UBound(LISTFIELD) + 1) = LISTDATA
        End If
        Range("A2").Copy
        Range(Cells(4, "D"), Cells(4 + cntRecord - 1, "T")).PasteSpecial Paste:=xlPasteValues, operation:=xlPasteSpecialOperationAdd
        Application.CutCopyMode = False
        
        '--//��ȸ��������ȸ MM ������ ����
        Dim indexR As Integer
        For indexR = 3 To Cells(Rows.Count, "B").End(xlUp).Row
            If Cells(indexR, "C") = "MM" And Cells(indexR - 1, "C") = "HBC" Then
                Cells(indexR, "E") = Cells(indexR - 1, "E")
                Cells(indexR, "F") = Cells(indexR - 1, "F")
            End If
        Next

        
        '--//���� �ʱ�ȭ
        Call sbClearVariant
        
        '--//��ȸ������ ���
        Range("Stat_ChurchDetail_Date") = DateSerial(Me.cboYear, Me.cboMonth, 1)
        
        strSql = "SELECT a.dept_nm FROM op_system.db_ovs_dept a WHERE a.dept_id = " & SText(USER_DEPT)
        Call makeListData(strSql, "op_system.db_ovs_dept")
        
        Sheets("��ȸ����").Range("A1").FormulaR1C1 = _
                        "=""" & LISTDATA(0, 0) & " ��ȸ�� �⼮��Ȳ�� [""&TEXT(Stat_ChurchDetail_Date,""yyyy�� mm��"")&"" ����]"""
        Call sbClearVariant
        
        Range("A2").Select
        Normal
    
    Case Else
    End Select
    
    '--//��Ʈ��ȣ
    Call shProtect(globalSheetPW)
    
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub makeListData(ByVal strSql As String, ByVal tableNM As String)

    Dim i As Integer, j As Integer
    
    '//DB���� �ڷ� ȣ���Ͽ� ���ڵ�¿� ��ȯ
    connectTaskDB
    callDBtoRS "makeListData", tableNM, strSql, Me.Name
    
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
    
End Sub
'-----------------------------------------
'  ���Ǻ� Select SQL�� �ۼ�
'    - makeSelectSQL(�˻���, ����)
'    - DB���� ��ȯ�� ����Ʈ �ʵ����
'-----------------------------------------
Private Function makeSelectSQL(ByVal tableNM As String) As String
    
    Dim strOrderByClause As String
    Dim strTemp As Variant
    
    Select Case tableNM
    Case TB1
        strSql = "SELECT * FROM " & TB1 & " a WHERE a.`�����μ�` = " & SText(USER_DEPT) & ";"
    Case TB2
        '--//�ش� �μ��� ����ȸ ��� ����
        strSql = "SELECT a.union_nm FROM " & "op_system.a_union" & " a WHERE a.ovs_dept = " & SText(USER_DEPT) & " ORDER BY a.sort_order;"
        Call makeListData(strSql, "op_system.a_union")
        
        If cntRecord > 0 Then '--//��ϵ� ����ȸ�� ������
            '--//ORDER BY ��������
            For Each strTemp In LISTDATA
                If strOrderByClause = "" Then
                    strOrderByClause = SText(strTemp)
                Else
                    strOrderByClause = strOrderByClause & "," & SText(strTemp)
                End If
            Next
            strOrderByClause = "FIELD(`����ȸ`," & strOrderByClause & ")"
            
            '--//���� strSQL�� ����
            strSql = "SELECT * FROM " & TB2 & " a WHERE a.`�����μ�` = " & SText(USER_DEPT) & " ORDER BY " & strOrderByClause & ",`���ļ���`;"
        Else
            '--//���� strSQL�� ����
            strSql = "SELECT * FROM " & TB2 & " a WHERE a.`�����μ�` = " & SText(USER_DEPT) & " ORDER BY `���ļ���`;"
        End If
    Case TB3
        '--//�ش� �μ��� ����ȸ ��� ����
        strSql = "SELECT a.union_nm FROM " & "op_system.a_union" & " a WHERE a.ovs_dept = " & SText(USER_DEPT) & " ORDER BY a.sort_order;"
        Call makeListData(strSql, "op_system.a_union")
        
        If cntRecord > 0 Then '--//��ϵ� ����ȸ�� ������
            '--//ORDER BY ��������
            For Each strTemp In LISTDATA
                If strOrderByClause = "" Then
                    strOrderByClause = SText(strTemp)
                Else
                    strOrderByClause = strOrderByClause & "," & SText(strTemp)
                End If
            Next
            strOrderByClause = "FIELD(`����ȸ`," & strOrderByClause & ")"
            
            '--//���� strSQL�� ����
            strSql = "SELECT * FROM " & TB3 & " a WHERE a.`�����μ�` = " & SText(USER_DEPT) & " ORDER BY " & strOrderByClause & ",`���ļ���`;"
        Else
            '--//���� strSQL�� ����
            strSql = "SELECT * FROM " & TB3 & " a WHERE a.`�����μ�` = " & SText(USER_DEPT) & " ORDER BY `���ļ���`;"
        End If
    Case TB4
        '--//���� strSQL�� ����
        strSql = "SELECT '�μ���ü',NULL '��ȸ����',NULL '������',NULL '��å',NULL '������',SUM(a.`��ü1ȸ`) '��ü1ȸ',SUM(a.`��ü4ȸ`) '��ü4ȸ',SUM(a.`�л�1ȸ`) '�л�1ȸ',SUM(a.`�л�4ȸ`) '�л�4ȸ',SUM(a.`�л�����`) '�л�����',SUM(a.`��üħ��`) '��üħ��',SUM(a.`������`) '������',SUM(a.`������`) '������',SUM(a.`������`) '������',NULL '�����μ�',NULL '����ȸ',NULL '����ȸ��',NULL '���ļ���','��ȸ����' FROM (SELECT * FROM " & TB4 & " a WHERE a.`�����μ�` = " & SText(USER_DEPT) & " AND a.`��ȸ����` IN ('MC','HBC') ORDER BY `���ļ���`) a" & _
                " UNION SELECT d.* FROM (SELECT b.`����ȸ`,NULL '��ȸ����',NULL '������',NULL '��å',NULL '������',SUM(b.`��ü1ȸ`),SUM(b.`��ü4ȸ`),SUM(b.`�л�1ȸ`),SUM(b.`�л�4ȸ`),SUM(b.`�л�����`),SUM(b.`��üħ��`),SUM(b.`������`),SUM(b.`������`),SUM(b.`������`),NULL '�����μ�',NULL '����ȸ��',NULL '����ȸ��',NULL '���ļ���','��ȸ����' FROM (SELECT * FROM " & TB4 & " a WHERE a.`�����μ�` = " & SText(USER_DEPT) & " AND a.`��ȸ����` IN ('MC','HBC') ORDER BY `���ļ���`) b LEFT JOIN op_system.a_union union_order ON union_order.union_nm=b.`����ȸ` GROUP BY `����ȸ` ORDER BY union_order.sort_order) d" & _
                " UNION SELECT c.* FROM (SELECT * FROM " & TB4 & " a WHERE a.`�����μ�` = " & SText(USER_DEPT) & " ORDER BY `���ļ���`) c;"
    Case Else
    End Select
    makeSelectSQL = strSql
End Function

Private Sub sbClearVariant()
    Erase LISTFIELD
    Erase LISTDATA
    cntRecord = Empty
    strSql = vbNullString
End Sub
Private Function fnData_Validation()
'---------------------------------------
'������ �Է°��� ���� ������ ��ȿ�� �˻�
'TRUE: �̻����, FALSE: �߸���.
'---------------------------------------
    fnData_Validation = True '�����Ͱ� ��ȿ�ϴٴ� ���� �Ͽ� ����
    
    
    If Not IsNumeric(Me.cboYear.Value) Then
        fnData_Validation = False
        MsgBox "�⵵ �Է� ���� ��ȿ���� �ʽ��ϴ�. �ٽ� Ȯ�� ���ּ���.", vbCritical, banner
        Set txtBox_Focus = Me.cboYear:  Exit Function
    End If
    If Me.cboYear < 1900 Or Me.cboYear > 2100 Then
        fnData_Validation = False
        MsgBox "�⵵ �Է� ���� ��ȿ���� �ʽ��ϴ�. �ٽ� Ȯ�� ���ּ���.", vbCritical, banner
        Set txtBox_Focus = Me.cboYear: Exit Function
    End If
    If Not IsNumeric(Me.cboMonth.Value) Then
        fnData_Validation = False
        MsgBox "�� �Է� ���� ��ȿ���� �ʽ��ϴ�. �ٽ� Ȯ�� ���ּ���.", vbCritical, banner
        Set txtBox_Focus = Me.cboMonth: Exit Function
    End If
    If Me.cboMonth > 12 Or Me.cboMonth < 1 Then
        fnData_Validation = False
        MsgBox "�� �Է� ���� ��ȿ���� �ʽ��ϴ�. �ٽ� Ȯ�� ���ּ���.", vbCritical, banner
        Set txtBox_Focus = Me.cboMonth: Exit Function
    End If
End Function

Private Sub initializeRepport()
    Select Case SEARCH_CODE
    Case 1
        With Sheets("������ ���")
            
            '[��������]
            Set rngA = Range("Stat_Country_Start")
            Set rngB = .Columns("A").Find("�հ�", lookat:=xlWhole)
            
            '[�Է³����ʱ�ȭ]
            rngA.Offset(1).Resize(rngB.Row - rngA.Row - 1, 32).ClearContents
            
            '[��⿵�� ����]
            rngB.Offset(1).Resize(Rows.Count - rngB.Row - 1).EntireRow.Delete Shift:=xlUp
            
            '[������]
            Set rngA = Nothing
            Set rngB = Nothing
        End With
    Case 2
        With Sheets("��ȸ���")
            
            '[��������]
            Set rngA = Range("Stat_Church_Start")
            Set rngB = .Columns("A").Find("�հ�", lookat:=xlWhole)
            
            '[�Է³����ʱ�ȭ]
            rngA.Offset(1).Resize(rngB.Row - rngA.Row - 1, 32).ClearContents
            
            '[��⿵�� ����]
            rngB.Offset(1).Resize(Rows.Count - rngB.Row - 1).EntireRow.Delete Shift:=xlUp
            
            '[������]
            Set rngA = Nothing
            Set rngB = Nothing
        End With
    Case 3
        With Sheets("��ȸ�����")
            
            '[��������]
            Set rngA = Range("Stat_PStaff_Start")
            Set rngB = .Columns("A").Find("�հ�", lookat:=xlWhole)
            
            '[�Է³����ʱ�ȭ]
            rngA.Offset(1).Resize(rngB.Row - rngA.Row - 1, 18).ClearContents
            
            '[��⿵�� ����]
            rngB.Offset(1).Resize(Rows.Count - rngB.Row - 1).EntireRow.Delete Shift:=xlUp
            
            '[������]
            Set rngA = Nothing
            Set rngB = Nothing
        End With
    Case 4
        With Sheets("��ȸ����")
            
            '[��������]
            Set rngA = Range("Stat_ChurchDetail_Start")
            Set rngB = .Columns("A").Find("�հ�", lookat:=xlWhole)
            
            '[�Է³����ʱ�ȭ]
            rngA.Offset(1).Resize(rngB.Row - rngA.Row - 1, 18).ClearContents
            
            '[��⿵�� ����]
            rngB.Offset(1).Resize(Rows.Count - rngB.Row - 1).EntireRow.Delete Shift:=xlUp
            
            '[������]
            Set rngA = Nothing
            Set rngB = Nothing
        End With
    Case Else
    End Select
End Sub

Private Sub makeReport()
    Dim i As Integer
    Dim iRow As Integer, jRow As Integer
    Dim cntColumn As Integer
    
    Select Case SEARCH_CODE
    Case 1
        With Sheets("������ ���")
            '//��������
            Set rngA = Range("Stat_Country_Start")
            Set rngB = .Columns("A").Find("�հ�", lookat:=xlWhole)
            cntColumn = 33
            
            '//������1 ����Ʈ �ۼ�
            '[�������]
            i = cntRecord
            
            '[���� ���� ����]
            iRow = rngB.Row - rngA.Row - 1 '���� ����Ʈ ����
            jRow = i - iRow '�ʰ� ����Ʈ ����
            
            If jRow > 0 Then '�����Ͱ� ������ �������� ���� ���
'                .Rows(rngB.Row - 1 & ":" & rngB.Row - 1 + jRow - 1).Insert Shift:=xlDown
                .Rows(rngB.Row).Resize(cntRecord - iRow).Insert Shift:=xlDown
                rngA.Offset(1).Resize(1, cntColumn).Copy .Range(rngA.Offset(1), rngA.Offset(2 + i - 2))
            ElseIf jRow < 0 And i <> 0 Then '�����Ͱ� ������ �������� ���� ���
                .Rows(rngB.Row - 2 & ":" & rngB.Row - 1 + jRow).Delete Shift:=xlUp
            ElseIf jRow < 0 And i = 0 And iRow > 1 Then '��ȸ �����Ͱ� ���� ���
                .Rows(rngB.Row - 2 & ":" & rngB.Row + jRow).Delete Shift:=xlUp
            End If
            
            '//��� ���� ����
            i = Cells(Rows.Count, "A").End(xlUp).Row
            Cells(Rows.Count, "A").End(xlUp).Offset(1).Resize(Rows.Count - i).EntireRow.Delete Shift:=xlUp
            With Range(rngA.Offset(1), rngB.Offset(-1)).Resize(, cntColumn).Borders(xlInsideHorizontal)
                .LineStyle = xlContinuous
                .ColorIndex = xlAutomatic
                .TintAndShade = 0
                .Weight = xlHairline
            End With
            With Range(rngA.Offset(1), rngB.Offset(-1)).Resize(, cntColumn).Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            
            '//�Լ�����
            Set rngB = .Columns("A").Find("�հ�", lookat:=xlWhole)
            Range(rngA.Offset(1, -1), rngB.Offset(-1)).Formula = "=row()-3"
            rngB.Offset(, 2).Resize(, cntColumn - 2).FormulaR1C1 = "=SUM(R[-" & cntRecord & "]C:R[-1]C)"
            
            '//������
            Set rngA = Nothing
            Set rngB = Nothing
        End With
    Case 2
        With Sheets("��ȸ���")
            '//��������
            Set rngA = Range("Stat_Church_Start")
            Set rngB = .Columns("A").Find("�հ�", lookat:=xlWhole)
            cntColumn = 32
            
            '//������1 ����Ʈ �ۼ�
            '[�������]
            i = cntRecord
            
            '[���� ���� ����]
            iRow = rngB.Row - rngA.Row - 1 '���� ����Ʈ ����
            jRow = i - iRow '�ʰ� ����Ʈ ����
            
            If jRow > 0 Then '�����Ͱ� ������ �������� ���� ���
'                .Rows(rngB.Row - 1 & ":" & rngB.Row - 1 + jRow - 1).Insert Shift:=xlDown
                .Rows(rngB.Row).Resize(cntRecord - iRow).Insert Shift:=xlDown
                rngA.Offset(1).Resize(1, cntColumn).Copy .Range(rngA.Offset(1), rngA.Offset(2 + i - 2))
            ElseIf jRow < 0 And i <> 0 Then '�����Ͱ� ������ �������� ���� ���
                .Rows(rngB.Row - 2 & ":" & rngB.Row - 1 + jRow).Delete Shift:=xlUp
            ElseIf jRow < 0 And i = 0 And iRow > 1 Then '��ȸ �����Ͱ� ���� ���
                .Rows(rngB.Row - 2 & ":" & rngB.Row + jRow).Delete Shift:=xlUp
            End If
            
            '//��� ���� ����
            i = Cells(Rows.Count, "A").End(xlUp).Row
            Cells(Rows.Count, "A").End(xlUp).Offset(1).Resize(Rows.Count - i).EntireRow.Delete Shift:=xlUp
            With Range(rngA.Offset(1), rngB.Offset(-1)).Resize(, cntColumn).Borders(xlInsideHorizontal)
                .LineStyle = xlContinuous
                .ColorIndex = xlAutomatic
                .TintAndShade = 0
                .Weight = xlHairline
            End With
            With Range(rngA.Offset(1), rngB.Offset(-1)).Resize(, cntColumn).Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            
            '//�Լ�����
            Set rngB = .Columns("A").Find("�հ�", lookat:=xlWhole)
            Range(rngA.Offset(1, -1), rngB.Offset(-1)).Formula = "=row()-3"
            rngB.Offset(, 5).Resize(, cntColumn - 5).FormulaR1C1 = "=SUM(R[-" & cntRecord & "]C:R[-1]C)"
            
            '//������
            Set rngA = Nothing
            Set rngB = Nothing
        End With
    Case 3
        With Sheets("��ȸ�����")
            '//��������
            Set rngA = Range("Stat_PStaff_Start")
            Set rngB = .Columns("A").Find("�հ�", lookat:=xlWhole)
            cntColumn = 19
            
            '//������1 ����Ʈ �ۼ�
            '[�������]
            i = cntRecord
            
            '[���� ���� ����]
            iRow = rngB.Row - rngA.Row - 1 '���� ����Ʈ ����
            jRow = i - iRow '�ʰ� ����Ʈ ����
            
            If jRow > 0 Then '�����Ͱ� ������ �������� ���� ���
'                .Rows(rngB.Row - 1 & ":" & rngB.Row - 1 + jRow - 1).Insert Shift:=xlDown
                .Rows(rngB.Row).Resize(cntRecord - iRow).Insert Shift:=xlDown
                rngA.Offset(1).Resize(1, cntColumn).Copy .Range(rngA.Offset(1), rngA.Offset(2 + i - 2))
            ElseIf jRow < 0 And i <> 0 Then '�����Ͱ� ������ �������� ���� ���
                .Rows(rngB.Row - 2 & ":" & rngB.Row - 1 + jRow).Delete Shift:=xlUp
            ElseIf jRow < 0 And i = 0 And iRow > 1 Then '��ȸ �����Ͱ� ���� ���
                .Rows(rngB.Row - 2 & ":" & rngB.Row + jRow).Delete Shift:=xlUp
            End If
            
            '//��� ���� ����
            i = Cells(Rows.Count, "A").End(xlUp).Row
            Cells(Rows.Count, "A").End(xlUp).Offset(1).Resize(Rows.Count - i).EntireRow.Delete Shift:=xlUp
            With Range(rngA.Offset(1), rngB.Offset(-1)).Resize(, cntColumn).Borders(xlInsideHorizontal)
                .LineStyle = xlContinuous
                .ColorIndex = xlAutomatic
                .TintAndShade = 0
                .Weight = xlHairline
            End With
            With Range(rngA.Offset(1), rngB.Offset(-1)).Resize(, cntColumn).Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            
            '//�Լ�����
            Set rngB = .Columns("A").Find("�հ�", lookat:=xlWhole)
            Range(rngA.Offset(1, -1), rngB.Offset(-1)).Formula = "=row()-4"
            rngB.Offset(, 5).Resize(, cntColumn - 5).FormulaR1C1 = "=SUM(R[-" & cntRecord & "]C:R[-1]C)"
            
            '//������
            Set rngA = Nothing
            Set rngB = Nothing
        End With
    Case 4
        With Sheets("��ȸ����")
            '//��������
            Set rngA = Range("Stat_ChurchDetail_Start")
            Set rngB = .Columns("A").Find("�հ�", lookat:=xlWhole)
            cntColumn = 19
            
            '//������1 ����Ʈ �ۼ�
            '[�������]
            i = cntRecord
            
            '[���� ���� ����]
            iRow = rngB.Row - rngA.Row - 1 '���� ����Ʈ ����
            jRow = i - iRow '�ʰ� ����Ʈ ����
            
            If jRow > 0 Then '�����Ͱ� ������ �������� ���� ���
'                .Rows(rngB.Row - 1 & ":" & rngB.Row - 1 + jRow - 1).Insert Shift:=xlDown
                .Rows(rngB.Row).Resize(cntRecord - iRow).Insert Shift:=xlDown
                rngA.Offset(1).Resize(1, cntColumn).Copy .Range(rngA.Offset(1), rngA.Offset(2 + i - 2))
            ElseIf jRow < 0 And i <> 0 Then '�����Ͱ� ������ �������� ���� ���
                .Rows(rngB.Row - 2 & ":" & rngB.Row - 1 + jRow).Delete Shift:=xlUp
            ElseIf jRow < 0 And i = 0 And iRow > 1 Then '��ȸ �����Ͱ� ���� ���
                .Rows(rngB.Row - 2 & ":" & rngB.Row + jRow).Delete Shift:=xlUp
            End If
            
            '//��� ���� ����
            i = Cells(Rows.Count, "A").End(xlUp).Row
            Cells(Rows.Count, "A").End(xlUp).Offset(1).Resize(Rows.Count - i).EntireRow.Delete Shift:=xlUp
            With Range(rngA.Offset(1), rngB.Offset(-1)).Resize(, cntColumn).Borders(xlInsideHorizontal)
                .LineStyle = xlContinuous
                .ColorIndex = xlAutomatic
                .TintAndShade = 0
                .Weight = xlHairline
            End With
            With Range(rngA.Offset(1), rngB.Offset(-1)).Resize(, cntColumn).Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            
            '//�Լ�����
            Set rngB = .Columns("A").Find("�հ�", lookat:=xlWhole)
            Range(rngA.Offset(1, -1), rngB.Offset(-1)).Formula = "=row()-3"
            rngB.Offset(, 6).Resize(, cntColumn - 10).FormulaR1C1 = "=SUMIF(R4C3:R" & cntRecord + 3 & "C3,""*MC*"",R[-" & cntRecord & "]C:R[-1]C)+SUMIF(R4C3:R" & cntRecord + 3 & "C3,""*HBC*"",R[-" & cntRecord & "]C:R[-1]C)"
            
            '//������
            Set rngA = Nothing
            Set rngB = Nothing
        End With
    Case Else
    End Select
End Sub



