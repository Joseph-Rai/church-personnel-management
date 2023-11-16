VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Search_Attendance 
   Caption         =   "��ȸ �˻�"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7155
   OleObjectBlob   =   "frm_Search_Attendance.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "frm_Search_Attendance"
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
Dim ws As Worksheet

Private Sub chkAllSelect_Change()
    If Me.chkAllSelect Then
        Me.Height = 220
        setValueSelectCheckBox (True)
    Else
        Me.Height = 276
        setValueSelectCheckBox (False)
    End If
    
    validateAnySelectionCheckBox
    
End Sub

Private Sub chkBC_Change()
    validateAnySelectionCheckBox
End Sub

Private Sub chkMC_Change()
    validateAnySelectionCheckBox
End Sub


Private Sub chkMM_Change()
    validateAnySelectionCheckBox
End Sub

Private Sub chkPBC_Change()
    validateAnySelectionCheckBox
End Sub

Private Sub cmdPrintPDFAllList_Click()
    Dim i As Integer
    Dim filePath As String
    Dim FileName As String
    Dim selectList As Object
    
    Set selectList = getSelectList
    
    Call Optimization
    
    If Me.lstChurch.ListCount <= 0 Then
        MsgBox "�˻��� ��ȸ�� �����ϴ�." & vbNewLine & _
                "���� �������� �� ��ȸ����� ��ȸ�ϼ���.", vbCritical, banner
        GoTo Here
    End If
    
    If Me.chkAllSelect = False And Not isAllSelectedCheckBox Then
        
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
            Dim arrList As Variant
            arrList = selectList.ToArray
            If IsInArray(.List(.listIndex, 2), arrList, True, rtnValue) <> -1 Then
                Call cmdOK_Click
                filePath = fnPrintAsPDF(filePath)
            End If
        Next
    End With
    
    MsgBox "�۾��� �Ϸ� �Ǿ����ϴ�." & vbNewLine & vbNewLine & _
            "����������: " & filePath, , banner
Here:
    Call Normal
End Sub

Private Sub cmdPrintPDFCurrentPage_Click()
    Dim filePath As String
    
    filePath = fnPrintAsPDF
    MsgBox "�۾��� �Ϸ� �Ǿ����ϴ�." & vbNewLine & vbNewLine & _
            "����������: " & filePath, , banner
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
    
    Dim intYear As Integer
    Dim intMonth As Integer
    
    If connIP = Empty Then setGlobalVariant '//�������� �缳��
    If checkLogin = 0 Then f_login.Show '//�α���üũ
    If checkLogin = 0 Then End '--//�α��� ���� �� ���ν��� ����
    
    '--//��Ʈ����
    Set ws = ActiveSheet
    
    '--//���ʼ���
    Me.cmdClose.Cancel = True
    TB1 = "op_system.db_churchlist_custom" '--//��ȸ����Ʈ
    TB2 = "op_system.temp_pstaff_by_time" '--//������ ����
    TB3 = "op_system.v_history_church" '--//��ȸ����
    TB4 = "op_system.db_attendance" '--//�⼮��Ȳ
    
    '--//����Ʈ�ڽ� ����
    With Me.lstChurch
        .ColumnCount = 4
        .ColumnHeads = False
        .ColumnWidths = "0,150,50,150,0" '��ȸ�ڵ�, ��ȸ��, ��ȸ����, ������ȸ��, Ŀ�����ڵ�
        .Width = 265.5
        .TextAlign = fmTextAlignLeft
        .Font = "����"
    End With
    
    '--//�޺��ڽ� ������ �߰�
    For intYear = year(Date) To 2005 Step -1 '--//�⵵ ä���
        Me.cboYear.AddItem intYear
    Next
    For intMonth = 12 To 1 Step -1 '--//�� ä���
        Me.cboMonth.AddItem intMonth
    Next
    
    '--//�޺��ڽ� ���� ������ ��¥�� �����߰�
    If Range("Atten_rngDate") = "" Then
        Me.cboYear = year(Date)
        Me.cboMonth = month(Date)
    Else
        If Range("Atten_MaxDate") <> WorksheetFunction.EDate(DateSerial(year(Date), month(Date), 1), -1) Then
            Me.cboYear = year(WorksheetFunction.EDate(Date, -1))
            Me.cboMonth = month(WorksheetFunction.EDate(Date, -1))
        Else
            Me.cboYear = year(Range("Atten_rngDate"))
            Me.cboMonth = month(Range("Atten_rngDate"))
        End If
    End If
    
    '--//üũ�ڽ� �ʱ�ȭ
    Me.chkAllSelect = True
    
    '--//���� �ʱ�ȭ
    Me.Height = 220
    
    Me.txtChurch.SetFocus
End Sub

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
    Me.cmdOk.Enabled = True
End Sub
Private Sub cmdClose_Click()
    Unload Me
End Sub
Private Sub cmdPrint_Click()
    ActiveWindow.SelectedSheets.PrintOut FROM:=1, to:=1, Copies:=1
End Sub
Private Sub cmdOK_Click()
    
    Dim filePath As String
    Dim FileName As String
    Dim result As T_RESULT
    Dim shp As Shape
    
    '--//��Ʈ Ȱ��ȭ �� �������
    WB_ORIGIN.Activate
    ws.Activate
    Call shUnprotect(globalSheetPW)
    
    '--//��ȸ ���ÿ��� �Ǵ�
    If Me.lstChurch.listIndex = -1 Then
        MsgBox "��ȸ�� �����ϼ���.", vbCritical, "����"
        Exit Sub
    End If
    
    '--//��¥ �Է¿���
    If Me.cboYear <> "" Or Me.cboMonth <> "" Then
        If Me.cboYear = "" And Me.cboMonth <> "" Then
            MsgBox "�˻� ���� �Է��� �ּ���.", vbCritical, banner
            Exit Sub
        End If
        
        If Me.cboYear <> "" And Me.cboMonth = "" Then
            MsgBox "�˻� ���� �Է��� �ּ���.", vbCritical, banner
            Exit Sub
        End If
        
        Range("Atten_rngDate") = DateSerial(Me.cboYear, Me.cboMonth, 1)
    End If
    
    '--//��ȸ���� �Է�
    With Me.lstChurch
        Range("E2:O2").ClearContents '--//���� �� �ʱ�ȭ
        Range("P2:R2").ClearContents
        Range("P3:R3").ClearContents
        
        '--//��ȸ�� �Է�
        If .List(.listIndex, 2) = "MC" Then
            Range("E2") = .List(.listIndex, 1) & " ��ü"
        Else
            Range("E2") = .List(.listIndex, 1)
        End If
        
        '--//��ȸ���� �Է�
        Range("P2:R2").ClearContents
        If .List(.listIndex, 2) = "BC" Then
            Range("P2") = "����ȸ"
        ElseIf .List(.listIndex, 2) = "PBC" Then
            Range("P2") = "�����"
        End If
        
        '--//������ȸ �Է�
        Range("P3:R3").ClearContents
        If Range("P2") <> "" Then
            Range("P3") = .List(.listIndex, 3)
        End If
    End With
    
    Call Optimization
    
    '--//��ȸ�� ���� ����
    Range("Atten_rngTarget").CurrentRegion.ClearContents '--//���� ������ ����
    
    strSql = "CALL `Routine_pstaff_by_time`(" & SText(Format(WorksheetFunction.EoMonth(DateSerial(Me.cboYear, Me.cboMonth, 1), 0), "yyyy-mm-dd")) & ", " & SText(USER_DEPT) & ");"
    connectTaskDB
    result.strSql = strSql
    result.affectedCount = executeSQL("cmdOK_Click", "temp_pstaff_by_time", strSql, Me.Name, "�ӽ���ȸ ���̺� ����")
'    writeLog "cmdADD_Click", "temp_pstaff_by_time", strSQL, 0, Me.Name, "�ӽ���ȸ ���̺� ����", result.affectedCount
    disconnectALL
    
    strSql = makeSelectSQL(TB2) '--//SQL��
    Call makeListData(strSql, TB2) '--//ListData �����
    
    '--//��ȯ�� ListData�� ���� ��Ʈ�� ����
    If cntRecord > 0 Then
        Range("Atten_rngTarget").Resize(1, UBound(LISTFIELD) + 1) = LISTFIELD
        Range("Atten_rngTarget").Offset(1).Resize(cntRecord, UBound(LISTFIELD) + 1) = LISTDATA
    End If
    Call sbClearVariant
    
    '--//��ȸ���� ����
'    Range("Atten_rngHistory").Resize(10, 2) = vbNullString
    Range("Atten_rngHistory_Data").CurrentRegion.ClearContents
    strSql = makeSelectSQL(TB3)
    Call makeListData(strSql, TB3)
    If cntRecord > 0 Then
        Range("Atten_rngHistory_Data").Resize(1, UBound(LISTFIELD) + 1) = LISTFIELD
        Range("Atten_rngHistory_Data").Offset(1).Resize(cntRecord, UBound(LISTFIELD) + 1) = LISTDATA
        Range("Atten_rngHistory_cntRecord") = cntRecord
        Range("Atten_rngHistory_Index") = IIf(cntRecord - 9 < 1, 1, cntRecord - 9)
        
        For Each shp In ActiveSheet.Shapes
            If shp.Name Like "*Move*" Then
                If cntRecord > 10 Then
                    shp.Visible = msoTrue
                Else
                    shp.Visible = msoFalse
                End If
            End If
        Next
    End If
    
    Call sbClearVariant
    
    '--//�⼮������ ����
    Range("Atten_rngAttendance_Data").CurrentRegion.ClearContents
    strSql = makeSelectSQL(TB4)
    Call makeListData(strSql, TB4)
    Range("Atten_rngAttendance_Data").Resize(, 10) = LISTFIELD
    If cntRecord > 0 Then
        Range("Atten_rngAttendance_Data").Offset(1).Resize(cntRecord, 10) = LISTDATA
        Range("Atten_cntRecord") = cntRecord
        Range("A1").Copy
        Range("Atten_rngAttendance_Data").Offset(1).Resize(cntRecord, 10).PasteSpecial Paste:=xlPasteValues, operation:=xlPasteSpecialOperationAdd
    End If
    Application.CutCopyMode = False
    Call sbClearVariant
    
    If Range("Atten_rngDate") > Range("Atten_MaxDate") And IsDate(Range("Atten_MaxDate")) Then
        Range("Atten_rngDate") = Range("Atten_MaxDate")
    End If
    
'------------------------------------��������--------------------------------------------
On Error Resume Next
    '--//���� ���� ����
    ActiveSheet.Pictures.Delete
    
    InsertPStaffPic Range("Atten_LifeNo"), Range("Atten_Pic_M") '--//������ ��������
    If Not (Range("Atten_LifeNo_Spouse") = "" Or Range("Atten_LifeNo_Spouse") = 0) Then
        InsertPStaffPic Range("Atten_LifeNo_Spouse"), Range("Atten_Pic_F") '--//����� ��������
    End If
    InsertChurchMap Range("Atten_ChurchCode"), Range("Atten_Church_Map") '--//��ȸ���� ����
    
    '--//�������� ��¥���� �ϳ� �߰� �� �����Ͽ� ��Ʋ���� ����
    InsertPStaffPic "", Range("T17")
    If ActiveSheet.Pictures.Count > 0 Then
        ActiveSheet.Pictures(ActiveSheet.Pictures.Count).Delete
    End If
On Error GoTo 0

    '--//��Ʈ����
    sbArrangeChart_Atten
    
    Normal
    
    Range("A2").Select
    
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
'        MsgBox "��ȯ�� DB �����Ͱ� �����ϴ�.", vbInformation, banner
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
        If Me.chkAll Then
            strSql = "SELECT a.church_sid,a.church_nm,a.church_gb,b.church_nm,c.church_sid_custom " & _
                        "FROM " & TB1 & " a LEFT JOIN op_system.db_churchlist_custom b ON a.main_church_cd = b.church_sid " & _
                        "LEFT JOIN op_system.db_history_church_establish c ON c.church_sid=REPLACE(a.church_sid,'MM','MC') " & _
                        "WHERE a.ovs_dept = " & USER_DEPT & " AND (a.church_nm LIKE '%" & Me.txtChurch & "%' OR b.church_nm LIKE '%" & Me.txtChurch & "%') ORDER BY a.sort_order;"
        Else
            strSql = "SELECT a.church_sid,a.church_nm,a.church_gb,b.church_nm,c.church_sid_custom " & _
                        "FROM " & TB1 & " a LEFT JOIN op_system.db_churchlist_custom b ON a.main_church_cd = b.church_sid " & _
                        "LEFT JOIN op_system.db_history_church_establish c ON c.church_sid=REPLACE(a.church_sid,'MM','MC') " & _
                        "WHERE a.ovs_dept = " & USER_DEPT & " AND a.suspend = 0 AND (a.church_nm LIKE '%" & Me.txtChurch & "%' OR b.church_nm LIKE '%" & Me.txtChurch & "%') ORDER BY a.sort_order;"
        End If
    Case TB2
        With Me.lstChurch
'            strSQL = "SELECT b.`�����ȣ`,b.`�ѱ��̸�(����)`,b.`����ڻ���`,b.`����ѱ��̸�(����)`,b.`���ʹ߷���`,b.`����ȸ�߷���` FROM " & TB2 & " b WHERE b.`�����μ�` = " & SText(USER_DEPT) & " AND b.`����ȸ��` = " & Replace(SText(.List(.ListIndex, 1)), " ����ȸ", "") & IIf(InStr(.List(.ListIndex, 2), "M") > 0, " AND b.`��å` LIKE '%��%'", "") & ";"
            strSql = "SELECT b.* FROM " & TB2 & " b WHERE b.`�����μ�` = " & SText(USER_DEPT) & " AND b.`����ȸ��` = " & Replace(SText(.List(.listIndex, 1)), " ����ȸ", "") & IIf(InStr(.List(.listIndex, 2), "M") > 0, " AND b.`��å` LIKE '%��%'", "") & ";"
        End With
    Case TB3
'        strSQL = "SELECT DATE_FORMAT(b.`��¥`,'%y��%c��') '��¥',b.`��ȸ����` FROM (SELECT a.`��¥`,a.`��ȸ����` FROM " & TB3 & " a WHERE a.`Ŀ�����ڵ�` = " & Replace(SText(Me.lstChurch.List(Me.lstChurch.ListIndex, 4)), "MM", "MC") & " ORDER BY a.`��¥` DESC LIMIT 10) b ORDER BY b.`��¥`;"
        strSql = "SELECT DATE_FORMAT(b.`��¥`,'%y��%c��') '��¥',b.`��ȸ����` FROM (SELECT a.`��¥`,a.`��ȸ����` FROM " & TB3 & " a WHERE a.`��¥` <= " & SText(Format(WorksheetFunction.EoMonth(DateSerial(Me.cboYear, Me.cboMonth, 1), 0), "yyyy-mm-dd")) & " AND a.`Ŀ�����ڵ�` = " & Replace(SText(Me.lstChurch.List(Me.lstChurch.listIndex, 4)), "MM", "MC") & " ORDER BY a.`��¥`) b ORDER BY b.`��¥`;"
    Case TB4
        If Left(Me.lstChurch.List(Me.lstChurch.listIndex), 2) = "MM" Then
            '--//���� ����ȸ �⼮������
'            strSQL = "SELECT a.attendance_dt,a.once_all,a.forth_all,a.once_stu,a.forth_stu,a.tithe_stu,a.baptism_all,a.evangelist,a.ul,a.gl FROM " & TB4 & " a WHERE a.church_sid = " & SText(Me.lstChurch.List(Me.lstChurch.ListIndex)) & _
'                            " AND attendance_dt BETWEEN ADDDATE(" & SText(Range("Atten_rngDate")) & ", INTERVAL -1 year) AND " & SText(Range("Atten_rngDate")) & " ORDER BY a.attendance_dt;"
            strSql = "SELECT a.attendance_dt,a.once_all,a.forth_all,a.once_stu,a.forth_stu,a.tithe_stu,a.baptism_all,a.evangelist,a.ul,a.gl FROM " & TB4 & " a WHERE a.church_sid = " & SText(Me.lstChurch.List(Me.lstChurch.listIndex)) & _
                            " ORDER BY a.attendance_dt;"
            Call makeListData(strSql, TB4)
            
            If Not cntRecord > 0 Then '--//���� ����ȸ �⼮�� ���� ��ȸ���� ��ü �⼮������ ����
                strSql = "SELECT a.church_sid_custom FROM op_system.db_history_church_establish a WHERE a.church_sid = " & SText(Replace(Me.lstChurch.List(Me.lstChurch.listIndex), "MM", "MC"))
                Call makeListData(strSql, "op_system.db_history_church_establish")
                
                If cntRecord > 0 Then
'                strSQL = "SELECT a.attendance_dt,a.once_all,a.forth_all,a.once_stu,a.forth_stu,a.tithe_stu,a.baptism_all,a.evangelist,a.ul,a.gl FROM " & TB4 & " a LEFT JOIN op_system.db_history_church_establish b ON a.church_sid = b.church_sid WHERE b.church_sid_custom = " & SText(listData(0, 0)) & " AND attendance_dt BETWEEN ADDDATE(" & _
'                            SText(Range("Atten_rngDate")) & ", INTERVAL -1 year) AND " & SText(Range("Atten_rngDate")) & " GROUP BY a.attendance_dt ORDER BY a.attendance_dt;"
                strSql = "SELECT a.attendance_dt,MAX(a.once_all),MAX(a.forth_all),MAX(a.once_stu),MAX(a.forth_stu),MAX(a.tithe_stu),MAX(a.baptism_all),MAX(a.evangelist),MAX(a.ul),MAX(a.gl) FROM " & TB4 & " a LEFT JOIN op_system.db_history_church_establish b ON a.church_sid = b.church_sid WHERE b.church_sid_custom = " & SText(LISTDATA(0, 0)) & _
                            "GROUP BY a.attendance_dt ORDER BY a.attendance_dt;"
                End If
            End If
        Else
            '--//�� �� �⼮ ������
            strSql = "SELECT a.church_sid_custom FROM op_system.db_history_church_establish a WHERE a.church_sid = " & SText(Replace(Me.lstChurch.List(Me.lstChurch.listIndex), "MM", "MC"))
            Call makeListData(strSql, "op_system.db_history_church_establish")
            
            If cntRecord > 0 Then
'                strSQL = "SELECT a.attendance_dt,a.once_all,a.forth_all,a.once_stu,a.forth_stu,a.tithe_stu,a.baptism_all,a.evangelist,a.ul,a.gl FROM " & TB4 & " a LEFT JOIN op_system.db_history_church_establish b ON a.church_sid = b.church_sid WHERE b.church_sid_custom = " & SText(listData(0, 0)) & " AND attendance_dt BETWEEN ADDDATE(" & _
'                            SText(Range("Atten_rngDate")) & ", INTERVAL -1 year) AND " & SText(Range("Atten_rngDate")) & " GROUP BY a.attendance_dt ORDER BY a.attendance_dt;"
                strSql = "SELECT a.attendance_dt,MAX(a.once_all),MAX(a.forth_all),MAX(a.once_stu),MAX(a.forth_stu),MAX(a.tithe_stu),MAX(a.baptism_all),MAX(a.evangelist),MAX(a.ul),MAX(a.gl) FROM " & TB4 & " a LEFT JOIN op_system.db_history_church_establish b ON a.church_sid = b.church_sid WHERE b.church_sid_custom = " & SText(LISTDATA(0, 0)) & _
                            "GROUP BY a.attendance_dt ORDER BY a.attendance_dt;"
            End If
        End If
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

'Private Sub sbArrangeChart_Atten()
'
'Dim noMax As Integer
'    Dim noMin As Integer
'    Dim i As Long
'    Dim j As Long
'    Dim Term As Long
'
'On Error Resume Next
'    noMax = WorksheetFunction.Max(Range("F17:R17")) '--//�л��̻� 1ȸ�⼮ �ִ밪
'    noMin = WorksheetFunction.Min(Range("F19:R19")) '--//�л��̻� 4ȸ�⼮ �ּҰ�
'    i = 1: j = 1
'
'
'
'      '--//�⼮ �׷��� ��Ʈ����
'      With Sheets("��ȸ�� �⼮��Ȳ").ChartObjects(1).Chart.Axes(xlValue)
'        '--//�ı� �Ը� ���� �������� �޸� �մϴ�.
'        Select Case noMax
'            Case Is <= 100: Term = 10
'            Case Is <= 500: Term = 50
'            Case Is <= 1000: Term = 100
'            Case Else: Term = 100
'        End Select
'
'        '--//������ �ִ밪�� ���մϴ�..
'        Do
'            If Term * i > noMax Then
'                .MaximumScale = Term * i
'                Exit Do
'            End If
'            i = i + 1
'        Loop
'
'        '--//������ �ּҰ��� ���մϴ�.
'        Do
'            If Term * j >= noMin * 0.9 Then
'                .MinimumScale = Term * (j - 1)
'                Exit Do
'            End If
'            j = j + 1
'        Loop
'
'        '--//������ �ִ밪�� �ּҰ��� ���̰� 4�� ����� �ƴϸ� �ִ밪 ����
'        Do
'            If (.MaximumScale - .MinimumScale) Mod 4 = 0 Then Exit Do
'            i = i + 1
'            .MaximumScale = Term * i
'        Loop
'
'        .MajorUnit = (.MaximumScale - .MinimumScale) / 4
'
'      End With
'
'      With Sheets("��ȸ�� �⼮��Ȳ").ChartObjects(1).Chart.Axes(xlValue, xlSecondary)
'        .MaximumScale = Application.Max(WorksheetFunction.RoundUp(WorksheetFunction.Max(Range("F26:R26")), 1), WorksheetFunction.RoundUp(WorksheetFunction.Max(Range("F27:R27")), 1))
'        .MinimumScale = Application.Min(WorksheetFunction.RoundDown(WorksheetFunction.Min(Range("F26:R26")), 1), WorksheetFunction.RoundDown(WorksheetFunction.Min(Range("F27:R27")), 1))
'      End With
'
'On Error GoTo 0
'
'End Sub

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
    FileName = Me.lstChurch.listIndex + 1 & ". " & Range("E2") & ".pdf" '--//���ϸ��� ��ȸ�̸�
    If (Len(Dir(filePath, vbDirectory)) <= 0) Then
        MkDir (filePath)
    End If
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:=filePath & FileName
    
    fnPrintAsPDF = filePath
    
End Function

Private Function getSelectList() As Object
    Dim selectList As Object
    Set selectList = CreateObject("System.Collections.ArrayList")
    
    If Me.chkMC Then
        selectList.Add "MC"
    End If
    
    If Me.chkMM Then
        selectList.Add "MM"
    End If
    
    If Me.chkBC Then
        selectList.Add "BC"
    End If
    
    If Me.chkPBC Then
        selectList.Add "PBC"
    End If
    
    Set getSelectList = selectList
    
End Function

Private Sub setValueSelectCheckBox(arg As Boolean)

    Me.chkMC = arg
    Me.chkMM = arg
    Me.chkBC = arg
    Me.chkPBC = arg

End Sub

Private Function isAllSelectedCheckBox() As Boolean

    Dim result As Boolean

    result = False
    If Me.chkMC And Me.chkMM And Me.chkBC And Me.chkPBC Then
        result = True
    End If
    
    isAllSelectedCheckBox = result

End Function

Private Function isNothingSelectedCheckBox() As Boolean

    Dim result As Boolean

    result = False
    If Not (Me.chkMC Or Me.chkMM Or Me.chkBC Or Me.chkPBC Or Me.chkAllSelect) Then
        result = True
    End If
    
    isNothingSelectedCheckBox = result

End Function

Private Sub validateAnySelectionCheckBox()

    If isNothingSelectedCheckBox Then
        Me.cmdPrintPDFAllList.Enabled = False
    Else
        Me.cmdPrintPDFAllList.Enabled = True
    End If

End Sub
