VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Update_Attendance 
   Caption         =   "�⼮������ ����������"
   ClientHeight    =   6960
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   8310.001
   OleObjectBlob   =   "frm_Update_Attendance.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "frm_Update_Attendance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cntlst1Col As Integer '//����Ʈ �÷� ��
Dim TB1 As String, TB2 As String, TB3 As String, TB4 As String '//���� ����� DB ���̺�
Const strOrderBy As String = "currency_un ASC" '//DB���� sort_order �ʵ�
Dim caseSave As Integer '//1: �߰�, 2: ����, 3: ����(SUSPEND), 4: ��������
Dim queryKey As Variant '//����Ʈ ��ġ ��ȯ�� ���� id
Dim LISTDATA() As String '--//DB���� �޾ƿ� rs�� �迭�� ����
Dim LISTFIELD() As String '--//DB���� �޾ƿ� rs�� �ʵ带 �迭�� ����
Dim cntRecord As Integer '--//DB���� �޾ƿ� ���ڵ��� ����
Dim strSql As String '--//SQL ������
Dim txtBox_Focus As MSForms.control

Private Sub lstAttendance_Click()
    
    '--//��Ʈ�� ����
    If Me.lstAttendance.listIndex <> -1 Then
        Me.cmdAdd.Enabled = True
        Me.cmdDelete.Enabled = True
        Me.cmdEdit.Enabled = True
        Me.txtOnce.Enabled = True
        Me.txtForth.Enabled = True
        Me.txtOnce_Stu.Enabled = True
        Me.txtForth_Stu.Enabled = True
        Me.txtTithe_All.Enabled = True
        Me.txtTithe_Stu.Enabled = True
        Me.txtBaptism.Enabled = True
        Me.txtEvangelist.Enabled = True
        Me.txtGL.Enabled = True
        Me.txtUL.Enabled = True
        Me.cboYear.Enabled = True
        Me.cboMonth.Enabled = False
    Else
        Call sbtxtBox_Init
        Me.cmdAdd.Enabled = False
        Me.cmdDelete.Enabled = False
        Me.cmdEdit.Enabled = False
        Me.txtOnce.Enabled = False
        Me.txtForth.Enabled = False
        Me.txtOnce_Stu.Enabled = False
        Me.txtForth_Stu.Enabled = False
        Me.txtTithe_All.Enabled = False
        Me.txtTithe_Stu.Enabled = False
        Me.txtBaptism.Enabled = False
        Me.txtEvangelist.Enabled = False
        Me.txtGL.Enabled = False
        Me.txtUL.Enabled = False
        Me.cboYear.Enabled = False
        Me.cboMonth.Enabled = False
    End If
    
    '--//�������
    If Me.lstAttendance.listIndex <> -1 Then
        With Me.lstAttendance
            Me.txtOnce = .List(.listIndex, 2)
            Me.txtOnce_Stu = .List(.listIndex, 3)
            Me.txtForth = .List(.listIndex, 4)
            Me.txtForth_Stu = .List(.listIndex, 5)
            Me.txtTithe_All = .List(.listIndex, 6)
            Me.txtTithe_Stu = .List(.listIndex, 7)
            Me.txtBaptism = .List(.listIndex, 8)
            Me.txtEvangelist = .List(.listIndex, 9)
            Me.txtGL = .List(.listIndex, 10)
            Me.txtUL = .List(.listIndex, 11)
            Me.cboYear = Left(.List(.listIndex, 1), 4)
            Me.cboMonth = Right(.List(.listIndex, 1), 2)
        End With
    End If
End Sub

Private Sub lstAttendance_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Un'HookListBoxScroll
End Sub

Private Sub lstAttendance_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Me.lstAttendance.ListCount > 0 Then
        'HookListBoxScroll Me, Me.lstAttendance
    End If
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
    
    '--//���ʼ���
    Me.cmdClose.Cancel = True
    TB1 = "op_system.db_churchlist_custom" '--//��ȸ����Ʈ
    TB2 = "op_system.db_attendance" '--//�⼮��Ȳ
    
    '--//���ѿ� ���� ��Ʈ�� ����
    Call HideDeleteButtonByUserAuth
    
    '--//��Ʈ�� ����
    Me.cmdAdd.Enabled = False
    Me.cmdAdd.Visible = False
    Me.cmdCancel.Visible = False
    Me.cmdNew.Enabled = False
    Me.cmdDelete.Enabled = False
    Me.cmdEdit.Enabled = False
    Me.txtOnce.Enabled = False
    Me.txtForth.Enabled = False
    Me.txtOnce_Stu.Enabled = False
    Me.txtForth_Stu.Enabled = False
    Me.txtTithe_All.Enabled = False
    Me.txtTithe_Stu.Enabled = False
    Me.txtBaptism.Enabled = False
    Me.txtEvangelist.Enabled = False
    Me.txtGL.Enabled = False
    Me.txtUL.Enabled = False
    Me.cboYear.Enabled = False
    Me.cboMonth.Enabled = False
    
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
    
    '--//����Ʈ�ڽ� ����
    With Me.lstChurch
        .ColumnCount = 4
        .ColumnHeads = False
        .ColumnWidths = "0,150,50,150" '��ȸ�ڵ�, ��ȸ��, ��ȸ����, ������ȸ��
'        .Width = 401
        .TextAlign = fmTextAlignLeft
        .Font = "����"
    End With
    
    With Me.lstAttendance
        .ColumnCount = 12
        .ColumnHeads = False
        .ColumnWidths = "0,45,30,30,30,40,40,30,30,35,30,100" '��ȸ�ڵ�, ��¥,1ȸ(��),1ȸ(�С�),4ȸ(��),4ȸ(�С�),��ü����,�л�����,ħ��,������,������,������
        .Width = 401
        .TextAlign = fmTextAlignLeft
        .Font = "����"
    End With
    
    Me.txtChurch.SetFocus
    Call WaitFor(0.005)
End Sub
Private Sub cmdSearch_Click()
    strSql = makeSelectSQL(TB1)
    Call makeListData(strSql, TB1)
    If cntRecord > 0 Then
        Me.lstChurch.List = LISTDATA
    Else
        Me.lstChurch.Clear
    End If
    Call sbClearVariant
End Sub
Private Sub txtChurch_Change()
    Me.txtChurch.BackColor = RGB(255, 255, 255)
End Sub
Private Sub cmdCancel_Click()
    'Call cmdbtn_visible
    Call INPUTMODE(False)
    Call sbtxtBox_Init
    Call HideDeleteButtonByUserAuth
    Call lstAttendance_Click
'    Me.cboYear.Enabled = False
'    Me.cboMonth.Enabled = False
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    
    Dim result As T_RESULT
    
    If MsgBox("������ �����͸� �����Ͻðڽ��ϱ�?", vbQuestion + vbYesNo, banner) = vbNo Then
        Exit Sub
    End If
    
    strSql = makeDeleteSQL(TB2)
    connectTaskDB
    result.strSql = strSql
    result.affectedCount = executeSQL("cmdDelete_Click", TB2, strSql, Me.Name, "�⼮�̷� ����")
    writeLog "cmdDelete_Click", TB2, strSql, 0, Me.Name, "�⼮�̷� ����"
    disconnectALL
    
    '--//�޼����ڽ�
    MsgBox "�ش� �����Ͱ� �����Ǿ����ϴ�.", , banner
    
    '--//����Ʈ�ڽ� ���ΰ�ħ
    Call lstChurch_Click
    Me.lstAttendance.listIndex = -1
    
End Sub

Private Sub cmdEdit_Click()
    
    Dim result As T_RESULT
    
    '--//������ ���� �ִ��� üũ
    With Me.lstAttendance
        If Me.txtOnce = .List(.listIndex, 2) And Me.txtOnce_Stu = .List(.listIndex, 3) And Me.txtForth = .List(.listIndex, 4) And _
            Me.txtForth_Stu = .List(.listIndex, 5) And Me.txtTithe_All = .List(.listIndex, 6) And Me.txtTithe_Stu = .List(.listIndex, 7) And Me.txtBaptism = .List(.listIndex, 8) And _
            Me.txtEvangelist = .List(.listIndex, 9) And Me.txtGL = .List(.listIndex, 10) And Me.txtUL = .List(.listIndex, 11) Then
            Exit Sub
        End If
    End With
    
    '--//������ ��ȿ�� �˻�
    If fnData_Validation = False Then
On Error Resume Next
        txtBox_Focus.SetFocus
        txtBox_Focus.SelStart = 0
        txtBox_Focus.SelLength = Len(txtBox_Focus)
On Error GoTo 0
        Exit Sub
    End If
    
    '--//SQL�� ����, ����, �αױ��

    strSql = makeUpdateSQL(TB4)
    result.strSql = strSql
    connectTaskDB
    result.affectedCount = executeSQL("cmdEdit_Click", TB4, strSql, Me.Name, "�⼮�̷� ������Ʈ")
    writeLog "cmdEdit_Click", TB4, strSql, 0, Me.Name, "�⼮�̷� ������Ʈ", result.affectedCount
    disconnectALL

    '--//�޼����ڽ�
    MsgBox "������Ʈ �Ǿ����ϴ�.", , banner
    
    '--//����Ʈ�ڽ� �ʱ�ȭ
    Call lstChurch_Click
    Me.lstAttendance.listIndex = queryKey
    
End Sub

Private Sub cmdNew_Click()
    'Call cmdbtn_visible
    Call INPUTMODE(True)
    Call HideDeleteButtonByUserAuth
    Me.lstAttendance.listIndex = Me.lstAttendance.ListCount - 1
    Call sbtxtBox_Init
    Me.cboYear = year(Date)
    Me.cboMonth = month(WorksheetFunction.EDate(Date, -1))
'    Me.cboYear.Enabled = True
'    Me.cboMonth.Enabled = True
End Sub

Private Sub cmdADD_Click()
    
    Dim argData As T_ATTENDANCE
    Dim result As T_RESULT
    
    '--//�ߺ�üũ
    With Me.lstChurch
        strSql = "SELECT * FROM " & TB2 & " a WHERE a.church_sid = " & SText(.List(.listIndex)) & _
                " AND a.attendance_dt = " & SText(DateSerial(Me.cboYear, Me.cboMonth, 1)) & ";"
        Call makeListData(strSql, TB2)
    End With
    
    If cntRecord > 0 Then
        MsgBox "�ߺ��� �Ⱓ�� ������ �� �����ϴ�. �Է� ���� �ٽ� Ȯ�����ּ���.", vbCritical, banner
        queryKey = Format(LISTDATA(0, 1), "yyyy-mm")
        Call returnListPosition2(Me, Me.lstAttendance.Name, queryKey)
        Exit Sub
    End If
    Call sbClearVariant
    
    '--//������ ��ȿ�� �˻�
    If fnData_Validation = False Then
On Error Resume Next
        txtBox_Focus.SetFocus
        txtBox_Focus.SelStart = 0
        txtBox_Focus.SelLength = Len(txtBox_Focus)
On Error GoTo 0
        Exit Sub
    End If
    
    '--//�۾��� ���� ����ü�� �� �߰�
    argData.church_sid = Me.lstChurch.List(Me.lstChurch.listIndex)
    argData.ATTENDANCE_DT = DateSerial(Me.cboYear, Me.cboMonth, 1)
    argData.ONCE_ALL = IIf(Me.txtOnce = "", 0, Me.txtOnce)
    argData.ONCE_STU = IIf(Me.txtOnce_Stu = "", 0, Me.txtOnce_Stu)
    argData.FORTH_ALL = IIf(Me.txtForth = "", 0, Me.txtForth)
    argData.FORTH_STU = IIf(Me.txtForth_Stu = "", 0, Me.txtForth_Stu)
    argData.TITHE_ALL = IIf(Me.txtTithe_All = "", 0, Me.txtTithe_All)
    argData.TITHE_STU = IIf(Me.txtTithe_Stu = "", 0, Me.txtTithe_Stu)
    argData.BAPTISM_ALL = IIf(Me.txtBaptism = "", 0, Me.txtBaptism)
    argData.Evangelist = IIf(Me.txtEvangelist = "", 0, Me.txtEvangelist)
    argData.GL = IIf(Me.txtGL = "", 0, Me.txtGL)
    argData.UL = IIf(Me.txtUL = "", 0, Me.txtUL)
    
    If WorksheetFunction.Sum(argData.BAPTISM_ALL, argData.Evangelist, argData.FORTH_ALL, argData.FORTH_STU, argData.GL, argData.ONCE_ALL, argData.ONCE_STU, argData.TITHE_STU, argData.UL) = 0 Then
        MsgBox "���� �Է��� �ּ���.", vbCritical, banner
        Exit Sub
    End If
    
    '--//������ ���� �� �αױ��
    strSql = makeInsertSQL(TB2, argData)
    connectTaskDB
    result.strSql = strSql
    result.affectedCount = executeSQL("cmdADD_Clikc", TB2, strSql, Me.Name, "�⼮�̷� �߰�")
    writeLog "cmdADD_Click", TB2, strSql, 0, Me.Name, "�⼮�̷� �߰�", result.affectedCount
    disconnectALL

    '--//�޼����ڽ�
    MsgBox "�߰� �Ǿ����ϴ�.", , banner
    Call lstChurch_Click
    Me.lstAttendance.listIndex = Me.lstAttendance.ListCount - 1
    
    '--//��ư���� �������
    'Call cmdbtn_visible
    Call INPUTMODE(False)
    Call HideDeleteButtonByUserAuth
    Me.cboYear.Enabled = False
    Me.cboMonth.Enabled = False
    
End Sub
Private Sub lstChurch_Click()
    
    '--//��ȸ ���ÿ��� �Ǵ�
    If Me.lstChurch.listIndex = -1 Then
        MsgBox "��ȸ�� �����ϼ���.", vbCritical, "����"
        Exit Sub
    End If
    
    '--//�⼮������ �߰�
    strSql = makeSelectSQL(TB2)
    Call makeListData(strSql, TB2)
On Error Resume Next
    Me.lstAttendance.List = LISTDATA
    If err.Number <> 0 Then
        Me.lstAttendance.Clear
    End If
On Error GoTo 0
    Call sbClearVariant
    
    '--//������ ������ ����
    With Me.lstAttendance
        .listIndex = .ListCount - 1
    End With
    
    '--//��Ʈ�� ����
    If Me.lstChurch.listIndex <> -1 Then
        Me.cmdNew.Enabled = True
    Else
        Me.cmdNew.Enabled = False
    End If
    
'    If Me.lstAttendance.listIndex <> -1 Then
'        Me.cmdADD.Enabled = True
'        Me.cmdDelete.Enabled = True
'        Me.cmdEdit.Enabled = True
'        Me.txtOnce.Enabled = True
'        Me.txtForth.Enabled = True
'        Me.txtOnce_Stu.Enabled = True
'        Me.txtForth_Stu.Enabled = True
'        Me.txtTithe_All.Enabled = True
'        Me.txtTithe_Stu.Enabled = True
'        Me.txtBaptism.Enabled = True
'        Me.txtEvangelist.Enabled = True
'        Me.txtGL.Enabled = True
'        Me.txtUL.Enabled = True
'        Me.cboYear.Enabled = True
'        Me.cboMonth.Enabled = True
'    Else
'        Call sbtxtBox_Init
'        Me.cmdADD.Enabled = False
'        Me.cmdDelete.Enabled = False
'        Me.cmdEdit.Enabled = False
'        Me.txtOnce.Enabled = False
'        Me.txtForth.Enabled = False
'        Me.txtOnce_Stu.Enabled = False
'        Me.txtForth_Stu.Enabled = False
'        Me.txtTithe_All.Enabled = False
'        Me.txtTithe_Stu.Enabled = False
'        Me.txtBaptism.Enabled = False
'        Me.txtEvangelist.Enabled = False
'        Me.txtGL.Enabled = False
'        Me.txtUL.Enabled = False
'        Me.cboYear.Enabled = False
'        Me.cboMonth.Enabled = False
'    End If
    
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
    
    Select Case tableNM
    Case TB1
        '--//��ȸ�ڵ�, ��ȸ��
        If Me.chkAll Then
            strSql = "SELECT a.church_sid,a.church_nm,a.church_gb,b.church_nm " & _
                        "FROM " & TB1 & " a LEFT JOIN op_system.db_churchlist_custom b ON a.main_church_cd = b.church_sid " & _
                        "WHERE a.ovs_dept = " & USER_DEPT & " AND (a.church_nm LIKE '%" & Me.txtChurch & "%' OR b.church_nm LIKE '%" & Me.txtChurch & "%') ORDER BY a.sort_order;"
        Else
            strSql = "SELECT a.church_sid,a.church_nm,a.church_gb,b.church_nm " & _
                        "FROM " & TB1 & " a LEFT JOIN op_system.db_churchlist_custom b ON a.main_church_cd = b.church_sid " & _
                        "WHERE a.ovs_dept = " & USER_DEPT & " AND a.suspend = 0 AND (a.church_nm LIKE '%" & Me.txtChurch & "%' OR b.church_nm LIKE '%" & Me.txtChurch & "%') ORDER BY a.sort_order;"
        End If
    Case TB2
        strSql = "SELECT a.church_sid,DATE_FORMAT(a.attendance_dt,'%Y-%m'),a.once_all,a.once_stu,a.forth_all,a.forth_stu,a.tithe_all,a.tithe_stu,a.baptism_all,a.evangelist,a.gl,a.ul FROM " & TB2 & " a WHERE church_sid = " & SText(Me.lstChurch.List(Me.lstChurch.listIndex)) & ";"
    Case Else
    End Select
    makeSelectSQL = strSql
End Function
Private Function makeUpdateSQL(ByVal tableNM As String) As String
    
    With Me.lstAttendance
        strSql = "UPDATE " & TB2 & " a " & _
                "SET a.once_all=" & SText(Me.txtOnce) & ",a.forth_all=" & SText(Me.txtForth) & ",a.once_stu=" & SText(Me.txtOnce_Stu) & _
                ",a.forth_stu=" & SText(Me.txtForth_Stu) & ",a.tithe_all=" & SText(Me.txtTithe_All) & " ,a.tithe_stu=" & SText(Me.txtTithe_Stu) & ",a.baptism_all=" & SText(Me.txtBaptism) & ",a.evangelist=" & SText(Me.txtEvangelist) & ",a.ul=" & SText(Me.txtUL) & ",a.gl=" & SText(Me.txtGL) & _
                " WHERE a.church_sid=" & SText(.List(.listIndex)) & " AND a.attendance_dt=" & SText(.List(.listIndex, 1) & "-01") & ";"
        queryKey = .listIndex
    End With
        
    makeUpdateSQL = strSql
End Function
Private Function makeInsertSQL(ByVal tableNM As String, _
                                argData As T_ATTENDANCE) As String
    With Me.lstAttendance
        strSql = "INSERT INTO " & TB2 & " VALUES(" & _
                    SText(argData.church_sid) & "," & _
                    SText(argData.ATTENDANCE_DT) & "," & _
                    SText(argData.ONCE_ALL) & "," & _
                    SText(argData.FORTH_ALL) & "," & _
                    SText(argData.ONCE_STU) & "," & _
                    SText(argData.FORTH_STU) & "," & _
                    SText(argData.TITHE_ALL) & "," & _
                    SText(argData.TITHE_STU) & "," & _
                    SText(argData.BAPTISM_ALL) & "," & _
                    SText(argData.Evangelist) & "," & _
                    SText(argData.GL) & "," & _
                    SText(argData.UL) & ");"
    End With
    queryKey = Me.lstAttendance.ListCount - 1
    makeInsertSQL = strSql
End Function
Private Function makeDeleteSQL(ByVal tableNM As String) As String
    With Me.lstAttendance
        strSql = "DELETE FROM " & TB2 & " WHERE church_sid = " & SText(.List(.listIndex)) & " AND attendance_dt = " & SText(.List(.listIndex, 1) & "-01") & ";"
    End With
    makeDeleteSQL = strSql
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
    If Not IsNumeric(Me.txtOnce.Value) Then
        fnData_Validation = False
        MsgBox "1ȸ �⼮(��ü) �Է� ���� ��ȿ���� �ʽ��ϴ�. �ٽ� Ȯ�� ���ּ���.", vbCritical, banner
        Set txtBox_Focus = Me.txtOnce: Exit Function
    End If
    If Not IsNumeric(Me.txtOnce_Stu.Value) Then
        fnData_Validation = False
        MsgBox "1ȸ �⼮(�л��̻�) �Է� ���� ��ȿ���� �ʽ��ϴ�. �ٽ� Ȯ�� ���ּ���.", vbCritical, banner
        Set txtBox_Focus = Me.txtOnce_Stu: Exit Function
    End If
    If Not IsNumeric(Me.txtForth.Value) Then
        fnData_Validation = False
        MsgBox "4ȸ �⼮(��ü) �Է� ���� ��ȿ���� �ʽ��ϴ�. �ٽ� Ȯ�� ���ּ���.", vbCritical, banner
        Set txtBox_Focus = Me.txtForth: Exit Function
    End If
    If Not IsNumeric(Me.txtForth_Stu.Value) Then
        fnData_Validation = False
        MsgBox "4ȸ �⼮(�л��̻�) �Է� ���� ��ȿ���� �ʽ��ϴ�. �ٽ� Ȯ�� ���ּ���.", vbCritical, banner
        Set txtBox_Focus = Me.txtForth_Stu: Exit Function
    End If
    If Not IsNumeric(Me.txtTithe_All.Value) Then
        fnData_Validation = False
        MsgBox "��ü���� �Է� ���� ��ȿ���� �ʽ��ϴ�. �ٽ� Ȯ�� ���ּ���.", vbCritical, banner
        Set txtBox_Focus = Me.txtTithe_All: Exit Function
    End If
    If Not IsNumeric(Me.txtTithe_Stu.Value) Then
        fnData_Validation = False
        MsgBox "�л��̻� ���� �Է� ���� ��ȿ���� �ʽ��ϴ�. �ٽ� Ȯ�� ���ּ���.", vbCritical, banner
        Set txtBox_Focus = Me.txtTithe_Stu: Exit Function
    End If
    If Not IsNumeric(Me.txtBaptism.Value) Then
        fnData_Validation = False
        MsgBox "ħ�� �Է� ���� ��ȿ���� �ʽ��ϴ�. �ٽ� Ȯ�� ���ּ���.", vbCritical, banner
        Set txtBox_Focus = Me.txtBaptism: Exit Function
    End If
    If Not IsNumeric(Me.txtEvangelist.Value) Then
        fnData_Validation = False
        MsgBox "���������� �Է� ���� ��ȿ���� �ʽ��ϴ�. �ٽ� Ȯ�� ���ּ���.", vbCritical, banner
        Set txtBox_Focus = Me.txtEvangelist: Exit Function
    End If
    If Not IsNumeric(Me.txtGL.Value) Then
        fnData_Validation = False
        MsgBox "������ �Է� ���� ��ȿ���� �ʽ��ϴ�. �ٽ� Ȯ�� ���ּ���.", vbCritical, banner
        Set txtBox_Focus = Me.txtGL: Exit Function
    End If
    If Not IsNumeric(Me.txtUL.Value) Then
        fnData_Validation = False
        MsgBox "������ �Է� ���� ��ȿ���� �ʽ��ϴ�. �ٽ� Ȯ�� ���ּ���.", vbCritical, banner
        Set txtBox_Focus = Me.txtUL: Exit Function
    End If
End Function
Sub sbtxtBox_Init()
    Me.txtOnce = ""
    Me.txtForth = ""
    Me.txtOnce_Stu = ""
    Me.txtForth_Stu = ""
    Me.txtTithe_All = ""
    Me.txtTithe_Stu = ""
    Me.txtBaptism = ""
    Me.txtEvangelist = ""
    Me.txtGL = ""
    Me.txtUL = ""
    Me.cboYear = ""
    Me.cboMonth = ""
End Sub
Private Sub HideDeleteButtonByUserAuth()
    Call GetUserAuthorities
    
    If cntRecord < 1 Then
        Exit Sub
    End If
    
    If IsInArray("DELETE_ITEM", LISTDATA) = -1 Then
        Me.cmdDelete.Visible = False
    End If
End Sub

Private Sub GetUserAuthorities()

    Dim sql As String
    
    sql = "SELECT b.authority FROM op_system.a_auth_table a" & _
          " LEFT JOIN op_system.a_authority b " & _
          "     ON a.authority_id = b.id" & _
          " WHERE a.user_id = " & USER_ID & ";"
    Call makeListData(sql, "op_system.a_auth_table")
    
End Sub

Private Sub cmdbtn_visible()
    Me.cmdNew.Visible = Not Me.cmdNew.Visible
    Me.cmdEdit.Visible = Not Me.cmdEdit.Visible
    Me.cmdDelete.Visible = Not Me.cmdDelete.Visible
    Me.cmdCancel.Visible = Not Me.cmdCancel.Visible
    Me.cmdAdd.Visible = Not Me.cmdAdd.Visible
End Sub
Private Sub INPUTMODE(ByVal argBoolean As Boolean)
    '--//��ư Ȱ��ȭ/��Ȱ��ȭ
    Me.cmdNew.Visible = Not argBoolean
    Me.cmdEdit.Visible = Not argBoolean
    Me.cmdDelete.Visible = Not argBoolean
    Me.cmdCancel.Visible = argBoolean
    Me.cmdCancel.Enabled = argBoolean
    Me.cmdAdd.Visible = argBoolean
    Me.cmdAdd.Enabled = argBoolean
    
    '--//��Ʈ�� Ȱ��ȭ/��Ȱ��ȭ
    Me.txtChurch.Enabled = Not argBoolean
    Me.cmdSearch.Enabled = Not argBoolean
    Me.lstChurch.Enabled = Not argBoolean
    Me.lstAttendance.Enabled = Not argBoolean
    Me.chkAll.Enabled = Not argBoolean
    
    Me.txtOnce.Enabled = argBoolean
    Me.txtOnce_Stu.Enabled = argBoolean
    Me.txtForth.Enabled = argBoolean
    Me.txtForth_Stu.Enabled = argBoolean
    Me.txtTithe_All.Enabled = argBoolean
    Me.txtTithe_Stu.Enabled = argBoolean
    Me.txtBaptism.Enabled = argBoolean
    Me.txtEvangelist.Enabled = argBoolean
    Me.txtGL.Enabled = argBoolean
    Me.txtUL.Enabled = argBoolean
    Me.cboYear.Enabled = argBoolean
    Me.cboMonth.Enabled = argBoolean
    Me.cmdAdd.Enabled = argBoolean
    
End Sub
Sub WaitFor(NumOfSeconds As Single)

    Dim SngSec As Single

    SngSec = Timer + NumOfSeconds

Do While Timer < SngSec

        DoEvents

   Loop

End Sub
