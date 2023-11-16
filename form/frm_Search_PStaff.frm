VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Search_PStaff 
   Caption         =   "��ȸ �˻�"
   ClientHeight    =   8280.001
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5250
   OleObjectBlob   =   "frm_Search_PStaff.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "frm_Search_PStaff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cntlst1Col As Integer '//����Ʈ �÷� ��
Dim TB1 As String, TB2 As String '//���� ����� DB ���̺�
Const strOrderBy As String = "currency_un ASC" '//DB���� sort_order �ʵ�
Dim caseSave As Integer '//1: �߰�, 2: ����, 3: ����(SUSPEND), 4: ��������
Dim queryKey As Integer '//����Ʈ ��ġ ��ȯ�� ���� id
Dim LISTDATA() As String '--//DB���� �޾ƿ� rs�� �迭�� ����
Dim LISTFIELD() As String '--//DB���� �޾ƿ� rs�� �ʵ带 �迭�� ����
Dim cntRecord As Integer '--//DB���� �޾ƿ� ���ڵ��� ����
Dim strSql As String '--//SQL ������
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

Private Sub chkAll_Change()
    If Me.chkAll.Value = True Then
        Me.chkManager.Enabled = False
        Me.chkManager.Value = True
        Me.chkOther.Enabled = False
        Me.chkOther.Value = True
        Me.chkPastoral.Enabled = False
        Me.chkPastoral.Value = True
        Me.chkTheological.Enabled = False
        Me.chkTheological.Value = True
        Me.chkAll.SetFocus
    Else
        Me.chkManager.Enabled = True
        Me.chkManager.Value = False
        Me.chkOther.Enabled = True
        Me.chkOther.Value = False
        Me.chkPastoral.Enabled = True
        Me.chkPastoral.Value = False
        Me.chkTheological.Enabled = True
        Me.chkTheological.Value = False
    End If
End Sub

Private Sub chkDate_Click()
    If Me.chkDate.Value = True Then
        Me.cboYear.Enabled = True
        Me.cboMonth.Enabled = True
        Me.cboYear = year(Range("PStaff_rngDate"))
        Me.cboMonth = month(Range("PStaff_rngDate"))
    Else
        Me.cboYear.Enabled = False
        Me.cboMonth.Enabled = False
        Me.cboYear = year(Date)
        Me.cboMonth = month(Date)
    End If
End Sub

Private Sub chkManager_Click()
    If Me.chkPastoral.Value = True And Me.chkOther.Value = True And Me.chkTheological.Value = True And Me.chkManager.Value = True Then
        Me.chkAll.Value = True
    End If
End Sub

Private Sub chkOther_Click()
    If Me.chkManager.Value = True And Me.chkPastoral.Value = True And Me.chkTheological.Value = True And Me.chkOther.Value = True Then
        Me.chkAll.Value = True
    End If
End Sub

Private Sub chkPastoral_Click()
    If Me.chkManager.Value = True And Me.chkOther.Value = True And Me.chkTheological.Value = True And Me.chkPastoral.Value = True Then
        Me.chkAll.Value = True
    End If
End Sub

Private Sub chkTheological_Click()
    If Me.chkManager.Value = True And Me.chkOther.Value = True And Me.chkPastoral.Value = True And Me.chkTheological.Value = True Then
        Me.chkAll.Value = True
    End If
End Sub

Private Sub cmdPrint_Click()
    Call sbPrint_PStaff
End Sub

Private Sub cmdPrintPDF_Click()
    Dim filePath As String
    
    filePath = fnPrintAsPDF
    MsgBox "�۾��� �Ϸ� �Ǿ����ϴ�." & vbNewLine & vbNewLine & _
            "����������: " & filePath, , banner
    
End Sub

Private Sub cmdPrintPDFAllList_Click()
    Dim i As Integer
    Dim filePath As String
    Dim FileName As String
    
    Call Optimization
    
    If Me.lstChurch.ListCount <= 0 Then
        MsgBox "�˻��� ��ȸ�� �����ϴ�." & vbNewLine & _
                "���� �������� �� ��ȸ����� ��ȸ�ϼ���.", vbCritical, banner
        GoTo Here
    End If
    
    '--//�ð��� ���� �ҿ�� �� �����Ƿ� ��� �޼��� ����
    If MsgBox("�ش� ����� �˻��� ��ȸ��� ��ü�� �� ���� ��ȸ�ϸ鼭" & vbNewLine & _
                "PDF�� ����ϹǷ� ����� �ð��� �ҿ�� �� �ֽ��ϴ�." & vbNewLine & vbNewLine & _
                "��� ���� �Ͻðڽ��ϱ�?", vbYesNo + vbInformation, banner) = vbNo Then
        GoTo Here
    End If
    
    filePath = GetDesktopPath '--//���� ��� ����
    filePath = filePath & "ExportByPDF"
    filePath = FileSequence(filePath) & Application.PathSeparator
    
    With Me.lstChurch
        For i = 0 To .ListCount - 1
            Me.lstChurch.listIndex = i
            Call cmdOK_Click
            filePath = fnPrintAsPDF(filePath)
        Next
    End With
    
    MsgBox "�۾��� �Ϸ� �Ǿ����ϴ�." & vbNewLine & vbNewLine & _
            "����������: " & filePath, , banner
Here:
    Call Normal
    
End Sub

Private Sub lstChurch_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Un'HookListBoxScroll
End Sub

Private Sub lstChurch_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Me.lstChurch.ListCount > 0 Then
        'HookListBoxScroll Me, Me.lstChurch
    End If
End Sub

Private Sub UserForm_Initialize()
    
    Dim i As Long
    
    If connIP = Empty Then setGlobalVariant '//�������� �缳��
    If checkLogin = 0 Then f_login.Show '//�α���üũ
    If checkLogin = 0 Then End '--//�α��� ���� �� ���ν��� ����
    
    '--//��Ʈ����
    Set ws = ActiveSheet
    
    '--//���ʼ���
    Me.cmdClose.Cancel = True
    TB1 = "op_system.db_churchlist_custom" '--//��ȸ����Ʈ
    TB2 = "op_system.temp_pstaff_by_time" '--//��������Ȳ
    
    '--//��Ʈ�� ����
    Me.cmdClose.Cancel = True
    Me.lblStatus.Visible = False
    Me.chkManager.Enabled = False
    Me.chkManager.Value = True
    Me.chkOther.Enabled = False
    Me.chkOther.Value = True
    Me.chkPastoral.Enabled = False
    Me.chkPastoral.Value = True
    Me.chkTheological.Enabled = False
    Me.chkTheological.Value = True
    Me.chkAll.Value = True
'    Me.cmdOK.Visible = False
    Me.cboYear.Enabled = False
    Me.cboMonth.Enabled = False
    Me.optSort1.Value = True
    
    '--//�޺��ڽ� ���� ����� �����߰�
    Me.cboYear = year(Date)
    Me.cboMonth = month(Date)
    
    '--//�޺��ڽ� ������ �߰�
    With Me.cboYear
        For i = year(Date) To year(Date) - 10 Step -1
            .AddItem i
        Next
    End With
    
    With Me.cboMonth
        For i = 1 To 12
            .AddItem i
        Next
    End With
    
    '--//����Ʈ�ڽ� ����
    With Me.lstChurch
        .ColumnCount = 2
        .ColumnHeads = False
        .ColumnWidths = "0,120" '��ȸ�ڵ�, ��ȸ��
        .Width = 241.45
        .TextAlign = fmTextAlignLeft
        .Font = "����"
    End With
    
    Me.txtChurch.SetFocus
End Sub

Private Sub cmdSearch_Click()
    strSql = makeSelectSQL(TB1)
    Call makeListData(strSql, TB1)
    If cntRecord > 0 Then
        Me.lstChurch.List = LISTDATA
    End If
    Call sbClearVariant
End Sub

Private Sub txtChurch_Change()
    Me.txtChurch.BackColor = RGB(255, 255, 255)
End Sub

Private Sub lstChurch_Click()
    cmdOk.Enabled = True
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub
Private Sub cmdOK_Click()
    Dim i As Integer, j As Integer
    Dim result As T_RESULT
    Dim page As Long
    Dim t As Single
    
    t = Timer
    
    '--//��ȸ ���ÿ��� �Ǵ�
    If Me.lstChurch.listIndex = -1 Then
        MsgBox "��ȸ�� �����ϼ���.", vbCritical, "����"
        Exit Sub
    End If
    
    '--//��Ʈ Ȱ��ȭ �� �������
    WB_ORIGIN.Activate
    ws.Activate
    Call shUnprotect(globalSheetPW)
    
'Debug.Print "��ƮȰ��ȭ �� �������: " & Format(Timer - t, "#0.00")
    
Application.Calculation = xlCalculationManual '�ڵ���� ���̱�
    
    '--//���� ������ ����
    Range("PStaff_rngTarget").CurrentRegion.ClearContents
    
    '--//temp_pstaff_by_time ���̺� ������Ʈ
    strSql = "CALL `Routine_pstaff_by_time`(" & SText(Format(WorksheetFunction.EoMonth(DateSerial(Me.cboYear, Me.cboMonth, 1), 0), "yyyy-mm-dd")) & ", " & SText(USER_DEPT) & ");"
    connectTaskDB
    result.strSql = strSql
    result.affectedCount = executeSQL("cmdADD_Click", "temp_pstaff_by_time", strSql, Me.Name, "�ӽ���ȸ ���̺� ����")
    disconnectALL

    '--//SQL��
    strSql = makeSelectSQL2
    
    '--//DB���� �ڷ� ȣ���Ͽ� ���ڵ�� ��ȯ
    Call makeListData(strSql, TB2)
    
    '--//��ȯ�� ListData�� ���� ��Ʈ�� ����
    Optimization
    If cntRecord > 0 Then
        Range("PStaff_rngTarget").Resize(cntRecord, UBound(LISTFIELD) + 1) = LISTFIELD
        Range("PStaff_rngTarget").Offset(1).Resize(cntRecord, UBound(LISTFIELD) + 1) = LISTDATA
    End If
    
    Normal
    
'Debug.Print "DB���� �ڷ� ��������: " & Format(Timer - t, "#0.00")
    
'    Application.Wait (Now + TimeValue("0:00:02")) '--//������ ���� ������ ��� ���
    Optimization
    '--//���� �ʱ�ȭ
    page = Int(WorksheetFunction.Quotient(cntRecord - 1, 9)) + 1
    Call sbClearVariant
    
    '--//���ڵ� ���� ���� ������ �����
    Call sbClearPic '���� �ʱ�ȭ
    Call sbMakePage(page) '--�ʿ��� ������ŭ ������ ����
    
Application.Calculation = xlCalculationAutomatic '�ڵ���� �츮��
'Application.CalculateFullRebuild
'Debug.Print "������ ����: " & Format(Timer - t, "#0.00")
    
    '--//��������
    Sheets("��ȸ�� ��������Ȳ").Range("B1").Select
    Me.lblStatus.Visible = True
    Me.Repaint
    Call sbInsertPic
    Me.lblStatus.Visible = False
    Me.Repaint
'Application.CalculateFullRebuild
'Debug.Print "��������: " + Format(Timer - t, "#0.00")
    
    '--//��ȸ������ ���
    Range("PStaff_rngDate") = DateSerial(Me.cboYear, Me.cboMonth, 1)
    
    '--//�ο���� '0��/0��' ������
    ActiveSheet.Outline.ShowLevels RowLevels:=0, ColumnLevels:=2
    On Error Resume Next
    Range(Range("PStaff_Stat_cntByPosition").Offset(, 0), Range("PStaff_Stat_cntByPosition").Offset(, 20)).Columns.Ungroup
    On Error GoTo 0
    If Not IsError(Range("PStaff_Stat_cntByPosition").Offset(, 1)) Then
        If Range("PStaff_Stat_cntByPosition").Offset(, 1) = "0��/0��" Then
            Range(Range("PStaff_Stat_cntByPosition").Offset(, 0), Range("PStaff_Stat_cntByPosition").Offset(, 1)).Columns.Group
        End If
    Else
        Range(Range("PStaff_Stat_cntByPosition").Offset(, 0), Range("PStaff_Stat_cntByPosition").Offset(, 1)).Columns.Group
    End If
    If Not IsError(Range("PStaff_Stat_cntByPosition").Offset(, 3)) Then
        If Range("PStaff_Stat_cntByPosition").Offset(, 3) = "0��/0��" Then
            Range(Range("PStaff_Stat_cntByPosition").Offset(, 2), Range("PStaff_Stat_cntByPosition").Offset(, 3)).Columns.Group
        End If
    Else
        Range(Range("PStaff_Stat_cntByPosition").Offset(, 2), Range("PStaff_Stat_cntByPosition").Offset(, 3)).Columns.Group
    End If
    If Not IsError(Range("PStaff_Stat_cntByPosition").Offset(, 5)) Then
        If Range("PStaff_Stat_cntByPosition").Offset(, 5) = "0��/0��" Then
            Range(Range("PStaff_Stat_cntByPosition").Offset(, 4), Range("PStaff_Stat_cntByPosition").Offset(, 5)).Columns.Group
        End If
    Else
        Range(Range("PStaff_Stat_cntByPosition").Offset(, 4), Range("PStaff_Stat_cntByPosition").Offset(, 5)).Columns.Group
    End If
    If Not IsError(Range("PStaff_Stat_cntByPosition").Offset(, 7)) Then
        If Range("PStaff_Stat_cntByPosition").Offset(, 7) = "0��/0��" Then
            Range(Range("PStaff_Stat_cntByPosition").Offset(, 6), Range("PStaff_Stat_cntByPosition").Offset(, 7)).Columns.Group
        End If
    Else
        Range(Range("PStaff_Stat_cntByPosition").Offset(, 6), Range("PStaff_Stat_cntByPosition").Offset(, 7)).Columns.Group
    End If
    ActiveSheet.Outline.ShowLevels RowLevels:=0, ColumnLevels:=1
    
    '--//�μ�ȸ�� �ʱ�ȭ
    Range("PStaff_rngPrint").ClearContents
'Debug.Print "�ο�����۾�: " + Format(Timer - t, "#0.00")
    Normal
    shProtect globalSheetPW
    
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
    
    '//�������� ���ڵ� �� ����
    If cntRecord = 0 Then
        MsgBox "��ȯ�� DB �����Ͱ� �����ϴ�.", vbInformation, banner
        Exit Sub
    End If
    
End Sub

'-----------------------------------------
'  ���Ǻ� Select SQL�� �ۼ�
'    - makeSelectSQL(�˻���, ����)
'    - DB���� ��ȯ�� ����Ʈ �ʵ����
'-----------------------------------------
Private Function makeSelectSQL(ByVal tableNM As String) As String
    Select Case tableNM
    Case TB1
        '--//��ȸ�ڵ�, ��ȸ��
        If Me.chkAllChurch Then
            strSql = "SELECT a.church_sid,a.church_nm " & _
                        "FROM " & TB1 & " a " & _
                        "WHERE a.ovs_dept = " & USER_DEPT & " AND a.church_gb IN ('MC','HBC') AND a.church_nm LIKE '%" & Me.txtChurch & "%' ORDER BY a.sort_order;"
        Else
            strSql = "SELECT a.church_sid,a.church_nm " & _
                        "FROM " & TB1 & " a " & _
                        "WHERE a.ovs_dept = " & USER_DEPT & " AND a.suspend = 0 AND a.church_gb IN ('MC','HBC') AND a.church_nm LIKE '%" & Me.txtChurch & "%' ORDER BY a.sort_order;"
        End If
    Case Else
        '//ȭ��id, ȭ���Ī, ��ȸ��, ��ȭȯ��, �޷�ȭȯ��
        'strSQL = "SELECT a.currency_id, a.currency_un, a.refer_dt, a.fx_rate_krw, a.fx_rate_usd " & _
                      "FROM " & TB1 & " a WHERE a.user_id = " & user_id & ";"
    End Select
    makeSelectSQL = strSql
End Function
'-----------------------------------------
'  ���Ǻ� Select SQL�� �ۼ�
'    - makeSelectSQL(�˻���, ����)
'    - DB���� ��ȯ�� ����Ʈ �ʵ����
'-----------------------------------------
Private Function makeSelectSQL2() As String
    
    '���Ǻ� sql�� ����
    strSql = makeSelectSqlByCondition()
    
    'Order By�� ����
    strSql = addOrderByClause(strSql)
    
    makeSelectSQL2 = strSql
End Function
Private Function addOrderByClause(query As String)

    If Me.optSort1.Value Then
        query = query & " Order By a.`��å` IS NULL ASC,FIELD(a.`��å`,'��ȸ��','����','��ȸ��븮','��븮���','����','�����','����ȸ������','����ȸ�����ڻ��','����Ұ�����','�������3�ܰ�','�������2�ܰ�','�������1�ܰ�','����Ұ����ڻ��', " & getPosition2Joining & ", '��������','����������','(��)��������','�屸����','���屸����','(��)�屸����','û������','��û������','(��)û������','û������','��û������','(��)û������','��������','����������','(��)��������','�б�����','���б�����','(��)�б�����','��å����',NULL),a.`����` IS NULL ASC,FIELD(a.`����`," & getTitleJoining & ",NULL),a.`���ʹ߷���`,a.`�������`;"
    End If
    
    If Me.optSort2.Value Then
        query = query & " Order By a.`��å` IS NULL ASC,FIELD(a.`��å`,'��ȸ��','����','��ȸ��븮','��븮���','����','�����','����ȸ������','����ȸ�����ڻ��','����Ұ�����','�������3�ܰ�','�������2�ܰ�','�������1�ܰ�','����Ұ����ڻ��', " & getPosition2Joining & ", '��������','����������','(��)��������','�屸����','���屸����','(��)�屸����','û������','��û������','(��)û������','û������','��û������','(��)û������','��������','����������','(��)��������','�б�����','���б�����','(��)�б�����','��å����',NULL),a.`����ü1ȸ` IS NULL ASC,a.`����ü1ȸ` DESC,a.`���ʹ߷���`;"
    End If
    
    If Me.optSort3.Value Then
        query = query & " Order By a.`����ü1ȸ` IS NULL DESC,a.`��å` IS NULL ASC,FIELD(a.`��å`,'��ȸ��','����','��ȸ��븮','��븮���','����','�����','����ȸ������','����ȸ�����ڻ��','����Ұ�����','�������3�ܰ�','�������2�ܰ�','�������1�ܰ�','����Ұ����ڻ��', " & getPosition2Joining & ", '��������','����������','(��)��������','�屸����','���屸����','(��)�屸����','û������','��û������','(��)û������','û������','��û������','(��)û������','��������','����������','(��)��������','�б�����','���б�����','(��)�б�����','��å����',NULL),a.`����` IS NULL ASC,FIELD(a.`����`, " & getTitleJoining & ",NULL),a.`����ü1ȸ` DESC,a.`���ʹ߷���`,a.`�������`;"
    End If
    
    If Me.optSort4.Value Then
        query = query & " Order By a.`����ü1ȸ` IS NULL DESC,a.`��å` IS NULL ASC,FIELD(a.`��å`,'��ȸ��','����','��ȸ��븮','��븮���','����','�����','����ȸ������','����ȸ�����ڻ��','����Ұ�����','�������3�ܰ�','�������2�ܰ�','�������1�ܰ�','����Ұ����ڻ��', " & getPosition2Joining & ", '��������','����������','(��)��������','�屸����','���屸����','(��)�屸����','û������','��û������','(��)û������','û������','��û������','(��)û������','��������','����������','(��)��������','�б�����','���б�����','(��)�б�����','��å����',NULL),a.`����ü1ȸ` DESC,a.`���ʹ߷���`,a.`�������`;"
    End If
    
    addOrderByClause = query

End Function
Private Function makeSelectSqlByCondition()

    Dim chk1st As Boolean
    Dim result As String
    
    chk1st = True
    result = "SELECT * FROM " & TB2 & " a WHERE a.`��ȸ��` = " & SText(Me.lstChurch.List(Me.lstChurch.listIndex, 1))
    
    If Me.chkAll.Value = False Then
        result = result & " AND ("
        If Me.chkPastoral.Value = True Then
            result = result & "a.`��å` LIKE '%��%' OR a.`��å` LIKE '%��%'"
            chk1st = False
        End If
        
        If Me.chkTheological.Value = True Then
            If chk1st = False Then
                result = result & " OR a.`�������` LIKE '%����%'"
            Else
                result = result & " a.`�������` LIKE '%����%'"
            End If
            chk1st = False
        End If
        
        If Me.chkManager.Value = True Then
            If chk1st = False Then
                result = result & " OR a.`��å` LIKE '%������%'"
            Else
                result = result & " a.`��å` LIKE '%������%'"
            End If
            chk1st = False
        End If
        
        If Me.chkOther.Value = True Then
            If chk1st = False Then
                result = result & " OR a.`��å2` IS NOT NULL"
            Else
                result = result & " a.`��å2` IS NOT NULL"
            End If
            chk1st = False
        End If
        
        result = result & ")"
        result = Replace(result, " AND ()", "") '��Ÿ�� ���� ���� �� ������ ���� ����
        
    End If
    
    If Me.optNoLeaderExclude.Value Then
        result = result & " AND a.`�����ȣ` IS NOT NULL"
    End If
    '--//�÷��� ����
    Range("PStaff_Stat_flagNoLeader") = Me.optNoLeaderInclude.Value
    
    makeSelectSqlByCondition = result

End Function
Public Sub sbInsertPic()

    Dim lifeNo As String
    Dim pageHeight As Integer: pageHeight = COUNT_PAGE_HEIGHT_CELLS
    Dim pageWidth As Integer: pageWidth = COUNT_PAGE_WIDTH_CELLS
    Dim lineHeight As Integer: lineHeight = (COUNT_PAGE_HEIGHT_CELLS - 1) / 3 '--//������ 1 ���� 3���� ����
    Dim targetRange As Range: Set targetRange = Range("C8")
    
    '--//���� �ʱ�ȭ
    Call sbClearPic

    '--//���� �ֱ� ���μ���
On Error Resume Next
    Dim i As Long, j As Long
    Dim tmpRange As Range
    For j = targetRange.Row To targetRange.Offset(lineHeight * 2).Row Step lineHeight
        For i = targetRange.Column To Range("A1").Value '--//Range("A1"): ������ �� ����ȣ ����
            '--//��������
            lifeNo = Cells(j, 1).Offset(, i - 1).Value
            Set tmpRange = Range(Cells(j, 1).Offset(, i - 1), Cells(j, 1).Offset(, i))
    
            '--//��������
            If Not (lifeNo = "" Or lifeNo = "0") Then
                InsertPStaffPic lifeNo, tmpRange
            End If
        Next i
    Next j
    
    '--//���� Ʋ���� ������ ���� ������ ���� ����
    InsertPStaffPic "", Range("A1")
    
    If ActiveSheet.Pictures.Count > 0 Then
        If Not ActiveSheet.Pictures(ActiveSheet.Pictures.Count).Name Like "*Stat_Shp*" Then
            ActiveSheet.Pictures(ActiveSheet.Pictures.Count).Delete
        End If
    End If
    
On Error GoTo 0

End Sub
Private Sub sbClearPic()
    Dim PicImage As Shape
    
    For Each PicImage In ActiveSheet.Shapes
        If PicImage.Name Like "*Stat_Shp*" Or PicImage.Name Like "*���簢��*" Or PicImage.Name Like "*Option*" Or PicImage.Name Like "*Rectangle*" Then
        Else
            PicImage.Delete
        End If
    Next
End Sub
Private Sub sbClearPicInRange(targetRange As Range)
    Dim PicImage As Shape
    
    For Each PicImage In ActiveSheet.Shapes
        If Not Intersect(PicImage.TopLeftCell, targetRange) Is Nothing Then
            PicImage.Delete
        End If
    Next
End Sub

Private Sub sbMakePage(page As Long)

    Dim targetCol As Long
    Dim i As Long
    Dim curPage As Long
    Dim PageStandard As Range
    Dim targetRange As Range
    
    '--// �������
    '--// COUNT_PAGE_WIDTH_CELLS
    '--// COUNT_PAGE_HEIGHT_CELLS
    
    '--//�ʿ��������� 0�̸� ���ν��� ����
    If page = 0 Then Exit Sub
    
    '--//ù������ ���ؼ� ����
    Set PageStandard = Range("C1")
    
    '--//�μ⿵�� �����ϱ� -> �����̸� ǥ�õ� ��� �������� 2���� ������ ���� �ذ����
On Error Resume Next
    ActiveSheet.HPageBreaks(1).DragOff Direction:=xlDown, RegionIndex:=1
On Error GoTo 0

    '--//���� ������ �� ��������
    curPage = Application.ExecuteExcel4Macro("Get.Document(50)")
    
    '--//�ʿ��� ������ ���� ���߱� ���μ���
    If curPage > page Then '���� �������� �ʿ� ���������� ���� ��
        '���� ������ ���� ���μ���
        With Range(PageStandard.Offset(, COUNT_PAGE_WIDTH_CELLS * (page)), PageStandard.Offset(, COUNT_PAGE_WIDTH_CELLS * (curPage) - 1))
            .EntireColumn.Delete Shift:=xlLeft
        End With
    ElseIf curPage < page Then '���� �������� �ʿ� ���������� ���� ��
        '�ű� ������ �߰� ���μ���
        Set targetRange = PageStandard.Offset(, 15 * curPage).Resize(, 15 * (page - curPage)).EntireColumn
        targetRange.Insert Shift:=xlRight
        PageStandard.Resize(COUNT_PAGE_HEIGHT_CELLS, COUNT_PAGE_WIDTH_CELLS).Copy
        For i = curPage To page - 1
            With PageStandard.Offset(, 15 * i).Resize(COUNT_PAGE_HEIGHT_CELLS, COUNT_PAGE_WIDTH_CELLS)
                .PasteSpecial Paste:=xlPasteFormats
                .PasteSpecial Paste:=xlPasteColumnWidths
                .PasteSpecial Paste:=xlPasteFormulas
            End With
        Next
        
        '����� �� �����ϱ�
        Set targetRange = PageStandard.Resize(COUNT_PAGE_HEIGHT_CELLS, COUNT_PAGE_WIDTH_CELLS)
        Call sbClearPicInRange(targetRange.Offset(, 15).Resize(2, 15 * (page - 1)))
    Dim rngUnion1 As Range
    Dim rngUnion2 As Range
    Dim rngUnion3 As Range
        For i = 1 To page - 1
            If rngUnion1 Is Nothing Then
                Set rngUnion1 = targetRange.Offset(, 15).Resize(1, 3).Offset(1, 15 * i - 4)
            Else
                Set rngUnion1 = UNION(rngUnion1, targetRange.Offset(, 15).Resize(1, 3).Offset(1, 15 * i - 4))
            End If
            
            If rngUnion2 Is Nothing Then
                Set rngUnion2 = targetRange.Offset(, 15).Resize(1, 2).Offset(, 15 * i - 3)
            Else
                Set rngUnion2 = UNION(rngUnion2, targetRange.Offset(, 15).Resize(1, 2).Offset(, 15 * i - 3))
            End If
            
            If rngUnion3 Is Nothing Then
                Set rngUnion3 = targetRange.Offset(, 15).Resize(2, 1).Offset(2, 15 * (i - 1))
            Else
                Set rngUnion3 = UNION(rngUnion3, targetRange.Offset(, 15).Resize(2, 1).Offset(2, 15 * (i - 1)))
            End If
            
        Next
        targetRange.Offset(, 15).Resize(1, 3).Offset(1, 6).Copy
        rngUnion1.PasteSpecial Paste:=xlPasteAll
        rngUnion2.ClearContents
        rngUnion3.FormatConditions.Delete
        
        '--//����Ʈ���� ����
        Set targetRange = PageStandard.Resize(COUNT_PAGE_HEIGHT_CELLS, COUNT_PAGE_WIDTH_CELLS)
        ActiveSheet.PageSetup.PrintArea = targetRange.Resize(, COUNT_PAGE_WIDTH_CELLS * page).Address
        For i = 1 To page - 1
            Set ActiveSheet.VPageBreaks(i).Location = PageStandard.Offset(, COUNT_PAGE_WIDTH_CELLS * i)
        Next
        
    Else
        '�ƹ��͵� ���ϱ�
    End If
    
End Sub

Private Sub sbClearVariant()
    Erase LISTFIELD
    Erase LISTDATA
    cntRecord = Empty
    strSql = vbNullString
End Sub
Private Sub sbSortData_PStaff()

    ActiveWorkbook.Worksheets("��ȸ�� ��������Ȳ").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("��ȸ�� ��������Ȳ").Sort.SortFields.Add key:=Range("PStaff_rngTarget").Offset(1, 8).Resize(Range("PStaff_rngTarget").CurrentRegion.Rows.Count - 1) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder:= _
        "��ȸ��,����,��ȸ��븮,��븮���,����,�����,����ȸ������,����ȸ�����ڻ��,����Ұ�����,�������3�ܰ�,�������2�ܰ�,�������1�ܰ�,����Ұ����ڻ��,��������,(��)��������,�屸����,(��)�屸����,û������,(��)û������,û������,(��)û������,��������,(��)��������,�б�����,(��)�б�����" _
        , DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("��ȸ�� ��������Ȳ").Sort.SortFields.Add key:=Range("PStaff_rngTarget").Offset(1, 12).Resize(Range("PStaff_rngTarget").CurrentRegion.Rows.Count - 1) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("��ȸ�� ��������Ȳ").Sort
        .SetRange Range("PStaff_rngTarget").CurrentRegion
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

Private Function fnPrintAsPDF(Optional filePath As String)

    Dim FileName As String
    
    '--//��Ʈ Ȱ��ȭ �� �������
    WB_ORIGIN.Activate
    ws.Activate
    
    '--//PDF�� ��������
    If filePath = "" Then
        filePath = GetDesktopPath '--//���� ��� ����
        filePath = filePath & "ExportByPDF"
        filePath = FileSequence(filePath) & Application.PathSeparator
    End If
    FileName = Range("E1") & ".pdf" '--//���ϸ��� ��ȸ�̸�
    If (Len(Dir(filePath, vbDirectory)) <= 0) Then
        MkDir (filePath)
    End If
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:=filePath & Me.lstChurch.listIndex + 1 & ". " & FileName
    
    fnPrintAsPDF = filePath
    
End Function
Private Function getPosition2Joining()

    Dim strQuery As String
    strQuery = "SELECT * FROM op_system.a_position2;"
    Call makeListData(strQuery, "op_system.a_position2")
        
    Dim result As String
    Dim i As Integer
    For i = 0 To cntRecord - 1
        If i < cntRecord - 1 Then
            result = result & "'" & LISTDATA(i, 0) & "', "
        Else
            result = result & "'" & LISTDATA(i, 0) & "'"
        End If
    Next
    
    getPosition2Joining = result

End Function
Private Function getTitleJoining()

    Dim strQuery As String
    strQuery = "SELECT * FROM op_system.a_title;"
    Call makeListData(strQuery, "op_system.a_title")
        
    Dim result As String
    Dim i As Integer
    For i = 0 To cntRecord - 1
        If i < cntRecord - 1 Then
            result = result & "'" & LISTDATA(i, 0) & "', "
        Else
            result = result & "'" & LISTDATA(i, 0) & "'"
        End If
    Next
    
    getTitleJoining = result

End Function
