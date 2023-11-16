VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Update_BCLeader 
   Caption         =   "������ �̷°��� ������"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8280.001
   OleObjectBlob   =   "frm_Update_BCLeader.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "frm_Update_BCLeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cntlst1Col As Integer '//����Ʈ �÷� ��
Dim TB1 As String, TB2 As String, TB3 As String, TB4 As String, TB5 As String, TB6 As String, TB7 As String '//���� ����� DB ���̺�
Const strOrderBy As String = "currency_un ASC" '//DB���� sort_order �ʵ�
Dim caseSave As Integer '//1: �߰�, 2: ����, 3: ����(SUSPEND), 4: ��������
Dim queryKey As Integer '//����Ʈ ��ġ ��ȯ�� ���� id
Dim LISTDATA() As String '--//DB���� �޾ƿ� rs�� �迭�� ����
Dim LISTFIELD() As String '--//DB���� �޾ƿ� rs�� �ʵ带 �迭�� ����
Dim cntRecord As Integer '--//DB���� �޾ƿ� ���ڵ��� ����
Dim strSql As String '--//SQL ������
Dim txtBox_Focus As MSForms.textBox

Private Sub chkDirect_Click()
    Call Direct_Mode(Me.chkDirect.Value)
End Sub

Private Sub chkPresent_Change()
    Select Case Me.chkPresent.Value
        Case True
            Me.txtEnd.BackColor = &HE0E0E0
            Me.txtEnd.Value = "����"
            Me.txtEnd.Enabled = False
        Case False
            Me.txtEnd.Enabled = True
            Me.txtEnd.BackColor = RGB(255, 255, 255)
            If Me.lstHistory.listIndex = -1 Then
                Me.txtEnd = Date - 1
            Else
                If Me.txtEnd = "����" Then
                    Me.txtEnd.Value = Date - 1
                End If
            End If
    Case Else
    End Select
End Sub

Private Sub cmdCancel_Click()
    'Call cmdbtn_visible
    Call sbtxtBox_Init
    Call INPUTMODE(False)
    Call HideDeleteButtonByUserAuth
    Call lstHistory_Click
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    
    Dim result As T_RESULT
    
    '--//����Ʈ�ڽ��� �������� �ʾ����� ���ν��� ����
    If Me.lstHistory.listIndex = -1 Then
        MsgBox "������ �����͸� ������ �ּ���.", vbCritical, banner
        Exit Sub
    End If
    
    '--//�������� ��Ȯ��
    If MsgBox("������ �����͸� �����Ͻðڽ��ϱ�?", vbQuestion + vbYesNo, banner) = vbNo Then
        Exit Sub
    End If
    
    strSql = makeDeleteSQL(TB2)
    connectTaskDB
    result.strSql = strSql
    result.affectedCount = executeSQL("cmdDelete_Click", TB2, strSql, Me.Name, "�������̷� ����")
    writeLog "cmdDelete_Click", TB2, strSql, 0, Me.Name, "�������̷� ����"
    disconnectALL
    
    '--//�޼����ڽ�
    MsgBox "�ش� �����Ͱ� �����Ǿ����ϴ�.", , banner
    
    '--//����Ʈ�ڽ� ���ΰ�ħ
    Call lstChurch_Click
    Me.lstHistory.listIndex = -1
    
End Sub

Private Sub cmdEdit_Click()
    
    Dim result As T_RESULT
    
    '--//����Ʈ�ڽ��� ���õǾ� ���� ������ ���ν��� ����
    If Me.lstHistory.listIndex = -1 Then
        MsgBox "������ �����͸� ������ �ּ���.", vbCritical, banner
        Exit Sub
    End If
    
    '--//������ ���� �ִ��� üũ
    With Me.lstHistory
        If Me.txtStart = .List(.listIndex, 2) And Me.txtEnd = .List(.listIndex, 3) And Me.txtLifeNo = .List(.listIndex, 4) And Me.cboResponsibility = .List(.listIndex, 7) Then
            Exit Sub
        End If
    End With
    
    '--//�ߺ�üũ
    With Me.lstHistory
        If Me.cboResponsibility = "������" Then
'            strSQL = "SELECT * FROM " & TB2 & " a WHERE a.church_sid = " & SText(Me.lstChurch.List(Me.lstChurch.ListIndex)) & _
'                    " AND a.responsibility = '������'" & _
'                    " AND ((a.start_dt BETWEEN " & SText(Me.txtStart) & " AND " & SText(IIf(Me.txtEnd = "����", DateSerial(9999, 12, 31), Me.txtEnd)) & _
'                    ") OR (a.end_dt BETWEEN " & SText(Me.txtStart) & " AND " & SText(IIf(Me.txtEnd = "����", DateSerial(9999, 12, 31), Me.txtEnd)) & _
'                    ") OR (a.start_dt <= " & SText(Me.txtStart) & " AND a.end_dt >= " & SText(IIf(Me.txtEnd = "����", DateSerial(9999, 12, 31), Me.txtEnd)) & _
'                    ")) AND a.bcleader_cd <> " & SText(.List(.ListIndex)) & ";"
            strSql = "SELECT * FROM " & TB2 & " a WHERE a.church_sid = " & SText(Me.lstChurch.List(Me.lstChurch.listIndex)) & _
                    " AND a.responsibility = '������'" & _
                    " AND a.bcleader_cd <> " & SText(.List(.listIndex, 0)) & _
                    " AND IF(a.start_dt > " & SText(Me.txtStart) & ", a.start_dt, " & SText(Me.txtStart) & ") <= " & _
                    " IF(a.end_dt < " & SText(IIf(Me.txtEnd = "����", DateSerial(9999, 12, 31), Me.txtEnd)) & ", a.end_dt, " & _
                        SText(IIf(Me.txtEnd = "����", DateSerial(9999, 12, 31), Me.txtEnd)) & ");"
            Call makeListData(strSql, TB2)
        Else
            cntRecord = 0
        End If
    End With
    
    If cntRecord > 0 Then
        MsgBox "�ߺ��� �Ⱓ�� ������ �� �����ϴ�. �Է� ���� �ٽ� Ȯ�����ּ���.", vbCritical, banner
        queryKey = LISTDATA(0, 0)
        Call returnListPosition(Me, Me.lstHistory.Name, queryKey)
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
    
    '--//SQL�� ����, ����, �αױ��
    strSql = makeUpdateSQL(TB2)
    result.strSql = strSql
    connectTaskDB
    result.affectedCount = executeSQL("cmdEdit_Click", TB2, strSql, Me.Name, "�������̷� ������Ʈ")
    writeLog "cmdEdit_Click", TB2, strSql, 0, Me.Name, "�������̷� ������Ʈ", result.affectedCount
    disconnectALL
    
    '--//�޼����ڽ�
    MsgBox "������Ʈ �Ǿ����ϴ�.", , banner
    
    '--//����Ʈ�ڽ� �ʱ�ȭ
    Call lstChurch_Click
    Me.lstHistory.listIndex = queryKey
    
End Sub

Private Sub cmdNew_Click()
    
    '--//Ŀ�ǵ� ��ư Ȱ��ȭ�� ���� �̷� ����Ʈ�ڽ� Ŭ��
    If lstHistory.ListCount = 0 Then
        Call lstHistory_Click
    End If
    
    'Call cmdbtn_visible
    Me.lstHistory.listIndex = Me.lstHistory.ListCount - 1
    Call sbtxtBox_Init
    Call INPUTMODE(True)
    Me.txtEnd = "����"
    If Me.txtEnd = "����" Then
        Me.chkPresent.Value = True
    End If
    Call HideDeleteButtonByUserAuth
End Sub

Private Sub cmdADD_Click()
    
    Dim argData As T_BC_LEADER
    Dim result As T_RESULT
    
    '--//�ߺ�üũ
    With Me.lstHistory
        If Me.cboResponsibility = "������" Then
            strSql = "SELECT * FROM " & TB2 & " a WHERE a.church_sid = " & SText(Me.lstChurch.List(Me.lstChurch.listIndex)) & _
                    " AND a.responsibility = '������'" & _
                    " AND IF(a.start_dt > " & SText(Me.txtStart) & ", a.start_dt, " & SText(Me.txtStart) & ") <= " & _
                    " IF(a.end_dt < " & SText(IIf(Me.txtEnd = "����", DateSerial(9999, 12, 31), Me.txtEnd)) & ", a.end_dt, " & _
                        SText(IIf(Me.txtEnd = "����", DateSerial(9999, 12, 31), Me.txtEnd)) & ");"
            Call makeListData(strSql, TB2)
        Else
            cntRecord = 0
        End If
        
    End With
    
    If cntRecord > 0 And Me.lstHistory.ListCount > 0 Then
        MsgBox "�ߺ��� �Ⱓ�� ������ �� �����ϴ�. �Է� ���� �ٽ� Ȯ�����ּ���.", vbCritical, banner
        queryKey = LISTDATA(0, 0)
        Call returnListPosition(Me, Me.lstHistory.Name, queryKey)
        'Call cmdbtn_visible
        Call INPUTMODE(False)
        Call HideDeleteButtonByUserAuth
        Exit Sub
    End If
    Call sbClearVariant
    
    '--//������ ��ȿ�� �˻�
    If fnData_Validation = False Then
On Error Resume Next
        txtBox_Focus.SetFocus
        txtBox_Focus.SelStart = 0
        txtBox_Focus.Setlength = Len(txtBox_Focus)
On Error GoTo 0
        Exit Sub
    End If
    
    '--//����ü�� �� �߰�
    argData.church_sid = Me.lstChurch.List(Me.lstChurch.listIndex)
    argData.START_DT = Me.txtStart
    argData.END_DT = IIf(Me.txtEnd = "����", DateSerial(9999, 12, 31), Me.txtEnd)
    argData.lifeNo = Me.txtLifeNo
    argData.RESPONSIBILITY = Me.cboResponsibility
    
    '--//������ ���� �� �αױ��
    strSql = makeInsertSQL(TB2, argData)
    connectTaskDB
    result.strSql = strSql
    result.affectedCount = executeSQL("cmdADD_Clikc", TB2, strSql, Me.Name, "�������̷� �߰�")
    writeLog "cmdADD_Click", TB2, strSql, 0, Me.Name, "�������̷� �߰�", result.affectedCount
    disconnectALL
    
    '--//�޼����ڽ�
    MsgBox "�߰� �Ǿ����ϴ�.", , banner
    Call lstChurch_Click
    Me.lstHistory.listIndex = Me.lstHistory.ListCount - 1
    
    '--//��ư���� �������
    If Me.chkDirect.Value = True Then
        Me.chkDirect.Value = False
        Call Direct_Mode(Me.chkDirect.Value)
    End If
    'Call cmdbtn_visible
    Call INPUTMODE(False)
    Call HideDeleteButtonByUserAuth
    
End Sub
Private Sub cmdSearch_manager_Click()
    argShow = 1
    frm_Update_BCLeader_1.Show
End Sub
Private Sub lstHistory_Click()
    
    '--//��Ʈ�� ����
    If Me.lstHistory.listIndex <> -1 Then
        Me.txtStart.Enabled = True
        Me.txtEnd.Enabled = True
        Me.cmdEdit.Enabled = True
        Me.cmdAdd.Enabled = True
        Me.cmdDelete.Enabled = True
        Me.chkPresent.Visible = True
        Me.chkDirect.Visible = True
        Me.cmdSearch_Manager.Enabled = True
        Me.cboResponsibility.Enabled = True
    Else
        Me.txtStart.Enabled = False
        Me.txtEnd.Enabled = False
        Me.cmdEdit.Enabled = False
        Me.cmdAdd.Enabled = False
        Me.cmdDelete.Enabled = False
        Me.chkPresent.Visible = False
        Me.chkDirect.Visible = False
        Me.cmdSearch_Manager.Enabled = False
        Me.cboResponsibility.Enabled = False
    End If
    
    '--//����Ʈ Ŭ�� �� ������, ������, ���� ǥ��
    If Me.lstHistory.listIndex <> -1 Then
        With Me.lstHistory
            Me.txtStart = .List(.listIndex, 2)
            Me.txtEnd = .List(.listIndex, 3)
            Me.txtLifeNo = .List(.listIndex, 4)
            Me.txtManager = .List(.listIndex, 5)
            Me.cboResponsibility = .List(.listIndex, 7)
        End With
    End If
    
    If Me.txtEnd = "����" Then
        Me.chkPresent.Value = True
        Me.txtEnd.Enabled = False
    Else
        Me.chkPresent.Value = False
    End If
    
End Sub

Private Sub lstHistory_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Un'HookListBoxScroll
End Sub

Private Sub lstHistory_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Me.lstHistory.ListCount > 0 Then
        'HookListBoxScroll Me, Me.lstHistory
    End If
End Sub

Private Sub lstChurch_Click()
    
    Dim filePath As String
    Dim FileName As String
    
    '--//��Ʈ�� ����
    If Me.lstChurch.listIndex <> -1 Then
        Me.lstHistory.Enabled = True
        Me.cmdNew.Enabled = True
    Else
        Me.lstHistory.Enabled = False
        Me.cmdNew.Enabled = False
    End If
    
    If Me.lstHistory.listIndex <> -1 Then
        Me.txtStart.Enabled = True
        Me.txtEnd.Enabled = True
        Me.cmdEdit.Enabled = True
        Me.cmdAdd.Enabled = True
        Me.cmdDelete.Enabled = True
        Me.chkPresent.Visible = True
        Me.chkDirect.Visible = True
        Me.cmdSearch_Manager.Enabled = True
        Me.cboResponsibility.Enabled = True
    Else
        Me.txtStart.Enabled = False
        Me.txtEnd.Enabled = False
        Me.cmdEdit.Enabled = False
        Me.cmdAdd.Enabled = False
        Me.cmdDelete.Enabled = False
        Me.chkPresent.Visible = False
        Me.chkDirect.Visible = False
        Me.cmdSearch_Manager.Enabled = False
        Me.cboResponsibility.Enabled = False
    End If
    
    '--//Ŭ�� �� �ؽ�Ʈ�ڽ� �ʱ�ȭ
    Call sbtxtBox_Init
    
    '--//�̷� ��ϻ��� ����
    With Me.lstHistory
        .ColumnCount = 7
        .ColumnHeads = False
        .ColumnWidths = "0,0,80,93,0,100,200" '�������ڵ�, ��ȸ�ڵ�, ������, ������, �����ȣ, �̸�, �Ҽӱ�ȸ
'        .Width = 345
        .TextAlign = fmTextAlignLeft
        .Font = "����"
    End With
    
    '--//�̷¸�� ������ ä���
    Call makeSelectSQL(TB2)
    Call makeListData(strSql, TB2)
    If cntRecord > 0 Then
        Me.lstHistory.List = LISTDATA
    Else
        Me.lstHistory.Clear
        Me.txtStart = ""
        Me.txtEnd = ""
        Me.txtManager = ""
        Me.txtLifeNo = ""
    End If
    Call sbClearVariant
    
End Sub

Private Sub lstChurch_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Un'HookListBoxScroll
End Sub

Private Sub lstChurch_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Me.lstChurch.ListCount > 0 Then
        'HookListBoxScroll Me, Me.lstChurch
    End If
End Sub

Private Sub txtChurch_Change()
    Me.txtChurch.BackColor = RGB(255, 255, 255)
End Sub

Private Sub txtEnd_Change()
    Call Date_Format(Me.txtEnd)
End Sub

Private Sub txtStart_Change()
    Call Date_Format(Me.txtStart)
End Sub

Private Sub UserForm_Initialize()
    If connIP = Empty Then setGlobalVariant '//�������� �缳��
    If checkLogin = 0 Then f_login.Show '//�α���üũ
    If checkLogin = 0 Then End '--//�α��� ���� �� ���ν��� ����
    
    '--//���ʼ���
    Me.cmdClose.Cancel = True
    TB1 = "op_system.db_churchlist_custom" '--//��ȸ����Ʈ
    TB2 = "op_system.db_branchleader" '--//������ �̷�
    
    '--//���ѿ� ���� ��Ʈ�� ����
    Call HideDeleteButtonByUserAuth
    
    '--//��Ʈ�� ����
    Me.lstChurch.Enabled = False
    Me.lstHistory.Enabled = False
    Me.txtStart.Enabled = False
    Me.txtEnd.Enabled = False
    Me.txtManager.Enabled = False
    Me.cmdEdit.Enabled = False
    Me.cmdAdd.Enabled = False
    Me.cmdAdd.Visible = False
    Me.cmdCancel.Visible = False
    Me.cmdNew.Enabled = False
    Me.txtLifeNo.Enabled = False
    Me.cmdSearch_Manager.Enabled = False
    Me.cmdDelete.Enabled = False
    Me.chkPresent.Visible = False
    Me.chkDirect.Visible = False
    
    '--//�޺��ڽ� ����
    Me.cboResponsibility.Clear
    Me.cboResponsibility.AddItem "������"
    Me.cboResponsibility.AddItem "�ܼ��Ҽ�"
    
    '--//����Ʈ�ڽ� ����
    With Me.lstChurch
        .ColumnCount = 4
        .ColumnHeads = False
        .ColumnWidths = "0,150,50,200" '��ȸ�ڵ�, ��ȸ��, ��ȸ����, ������ȸ��
'        .Width = 330
        .TextAlign = fmTextAlignLeft
        .Font = "����"
    End With
    
End Sub
Private Sub cmdSearch_Click()
    
    Me.lstHistory.Clear
    
    strSql = makeSelectSQL(TB1)
    Call makeListData(strSql, TB1)
    
    If cntRecord = 0 Then
        MsgBox "��ȯ�� DB �����Ͱ� �����ϴ�.", vbInformation, banner
        Call sbClearVariant
        Exit Sub
    End If
    
    Me.lstChurch.List = LISTDATA
    Call sbClearVariant
    Me.lstChurch.Enabled = True
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
    cntRecord = rs.RecordCount '--//���ڵ� �� ����
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
        If Me.chkOld.Value = False Then
            strSql = "SELECT a.church_sid,a.church_nm,a.church_gb,b.church_nm " & _
                        "FROM " & TB1 & " a LEFT JOIN op_system.db_churchlist_custom b ON a.main_church_cd = b.church_sid " & _
                        "WHERE a.ovs_dept = " & USER_DEPT & " AND a.suspend = 0 AND (a.church_gb NOT LIKE '%M%' AND a.church_gb NOT LIKE '%H%') AND (a.church_nm LIKE '%" & Me.txtChurch & "%' OR b.church_nm LIKE '%" & Me.txtChurch & "%') ORDER BY a.sort_order;"
        Else
            strSql = "SELECT a.church_sid,a.church_nm,a.church_gb,b.church_nm " & _
                        "FROM " & TB1 & " a LEFT JOIN op_system.db_churchlist_custom b ON a.main_church_cd = b.church_sid " & _
                        "WHERE a.ovs_dept = " & USER_DEPT & " AND (a.church_gb NOT LIKE '%M%' AND a.church_gb NOT LIKE '%H%') AND (a.church_nm LIKE '%" & Me.txtChurch & "%' OR b.church_nm LIKE '%" & Me.txtChurch & "%') ORDER BY a.sort_order;"
        End If
    Case TB2
        strSql = "SELECT a.bcleader_cd,a.church_sid,a.start_dt,if(a.end_dt='9999-12-31','����',a.end_dt),a.lifeno,concat(If(isnull(b.name_ko),a.lifeno,b.name_ko),ifnull(concat('(',left(c.Title,1),')'),'')),e.church_nm,a.responsibility " & _
                "FROM " & TB2 & " a LEFT JOIN op_system.db_pastoralstaff b " & _
                "ON a.lifeno = b.lifeno " & _
                "LEFT JOIN op_system.db_title c ON a.lifeno = c.lifeno AND (CURRENT_DATE BETWEEN c.Start_dt AND c.End_dt) " & _
                "LEFT JOIN op_system.db_transfer d ON a.lifeno = d.lifeno AND CURDATE() BETWEEN d.start_dt AND d.end_dt LEFT JOIN op_system.db_churchlist e ON d.church_sid = e.church_sid " & _
                "WHERE a.church_sid = " & SText(Me.lstChurch.List(Me.lstChurch.listIndex)) & _
                " ORDER BY a.start_dt;"
    Case Else
    End Select
    makeSelectSQL = strSql
End Function
Private Function makeUpdateSQL(ByVal tableNM As String) As String
    
    Select Case tableNM
    Case TB1
    Case TB2
        With Me.lstHistory
            strSql = "UPDATE " & TB2 & " a " & _
                    "SET a.start_dt = " & SText(Me.txtStart) & ", a.end_dt = " & SText(IIf(Me.txtEnd = "����", DateSerial(9999, 12, 31), Me.txtEnd)) & _
                    ",a.lifeno = " & SText(Me.txtLifeNo) & ",a.responsibility = " & SText(Me.cboResponsibility) & " WHERE a.bcleader_cd = " & SText(.List(.listIndex)) & ";"
        queryKey = .listIndex
        End With
    Case Else
    End Select
    makeUpdateSQL = strSql
End Function
Private Function makeInsertSQL(ByVal tableNM As String, argData As T_BC_LEADER) As String
    
    Select Case tableNM
    Case TB1
    Case TB2
        strSql = "INSERT INTO " & TB2 & " VALUES(DEFAULT," & _
                    SText(argData.church_sid) & "," & _
                    SText(argData.START_DT) & "," & _
                    SText(argData.END_DT) & "," & _
                    SText(argData.lifeNo) & "," & _
                    SText(argData.RESPONSIBILITY) & ");"
        queryKey = Me.lstHistory.ListCount - 1
    Case Else
    End Select
    makeInsertSQL = strSql
End Function
Private Function makeDeleteSQL(ByVal tableNM As String) As String
    
    Select Case tableNM
    Case TB1
    Case TB2
        With Me.lstHistory
            strSql = "DELETE FROM " & TB2 & " WHERE bcleader_cd = " & SText(.List(.listIndex)) & ";"
        End With
    Case Else
    End Select
    makeDeleteSQL = strSql
End Function
Private Sub sbClearVariant()
    Erase LISTFIELD
    Erase LISTDATA
    cntRecord = Empty
    strSql = vbNullString
End Sub
Sub sbtxtBox_Init()
    Me.txtStart.Value = ""
    Me.chkPresent.Value = False
    Me.txtEnd.Value = ""
    Me.txtManager.Value = ""
    Me.txtLifeNo.Value = ""
    Me.cboResponsibility.Value = ""
End Sub
Private Function fnData_Validation()
'---------------------------------------
'������ �Է°��� ���� ������ ��ȿ�� �˻�
'TRUE: �̻����, FALSE: �߸���.
'---------------------------------------
    fnData_Validation = True '�����Ͱ� ��ȿ�ϴٴ� ���� �Ͽ� ����
    
    '--//�����ȣ ������ �Է¿��� Ȯ��
    If Me.txtLifeNo = "" Then
        If Me.chkDirect = False Then
            MsgBox "�����ڸ� �Է��� �ּ���.", vbCritical, banner
            Exit Function
        Else
            MsgBox "�����ȣ�� �Է��� �ּ���.", vbCritical, banner
            Exit Function
        End If
    End If
    
    '--//�����ȣ ����üũ
    If Me.txtLifeNo <> "" Then
        If Not IsNumeric(fnExtract(Me.txtLifeNo)) Then
            fnData_Validation = False
            MsgBox "������ �����ȣ�� �߸��Ǿ����ϴ�. �ٽ� Ȯ���� �ּ���.", vbCritical, banner
            Set txtBox_Focus = Me.txtLifeNo
            Exit Function
        ElseIf Mid(Me.txtLifeNo, 4, 1) <> "-" Or Mid(Me.txtLifeNo, 11, 1) <> "-" Then
            fnData_Validation = False
            MsgBox "������ �����ȣ�� �߸��Ǿ����ϴ�. �ٽ� Ȯ���� �ּ���.", vbCritical, banner
            Set txtBox_Focus = Me.txtLifeNo
            Exit Function
        End If
    End If
    
    '--//��¥ ����üũ
    If Not IsDate(Me.txtStart) Then
        MsgBox "�ùٸ� ��¥ ���°� �ƴմϴ�. �������� �ٽ� Ȯ�� ���ּ���.", vbCritical, banner
        Set txtBox_Focus = Me.txtStart: fnData_Validation = False: Exit Function
    End If
    If Not IsDate(Me.txtEnd) And Me.txtEnd <> "����" Then
        MsgBox "�ùٸ� ��¥ ���°� �ƴմϴ�. �������� �ٽ� Ȯ�� ���ּ���.", vbCritical, banner
        Set txtBox_Focus = Me.txtEnd: fnData_Validation = False: Exit Function
    End If
    
    '--//��¥ ��ȿ�� ���� ����
    If Me.txtEnd <> "����" Then
        If CDate(Me.txtEnd) <= CDate(Me.txtStart) Then
            MsgBox "�������� �����Ϻ��� �۰ų� ���� �� �����ϴ�.", vbCritical, banner
            fnData_Validation = False: Exit Function
        End If
    End If
    
    '--//�޺��ڽ� �� ����
    If Not (Me.cboResponsibility = "������" Or Me.cboResponsibility = "�ܼ��Ҽ�") Then
        MsgBox "������ Ȥ�� �ܼ��Ҽ� �߿��� ������ �ּ���.", vbCritical, banner
        Set txtBox_Focus = Me.cboResponsibility: fnData_Validation = False: Exit Function
    End If
    
    '--//������, ������ ��å�ڸ� �����ڷ� ��� ����
    Dim lifeNo As String
    Dim objPosition As position
    Dim objPositionDao As New PositionDao
    Dim availablePositionList As Object
    
    If Me.cboResponsibility = "������" And Me.chkDirect.Value = False Then
        Set availablePositionList = CreateObject("System.Collections.ArrayList")
        availablePositionList.Add "����"
        availablePositionList.Add "����ȸ������"
        availablePositionList.Add "����Ұ�����"
        
        lifeNo = Me.txtLifeNo
        Set objPosition = objPositionDao.FindPositionByLifeNoAndDate(lifeNo, Now)
        
        If objPosition Is Nothing Then Set objPosition = New position
        If Not availablePositionList.Contains(objPosition.position) Then
            MsgBox "����, ����ȸ������, ����Ұ����� ��å�ڸ� �����ڷ� ����� �� �ֽ��ϴ�." & vbNewLine & _
                    "�����ȣ�� ���� �Է��ϰų� ���� �� ������ ��å�� ���� ������ּ���."
            fnData_Validation = False: Exit Function
        End If
    End If
    
End Function
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
    Me.cmdSearch_Manager.Enabled = argBoolean
    
    '--//��Ʈ�� Ȱ��ȭ/��Ȱ��ȭ
    Me.txtChurch.Enabled = Not argBoolean
    Me.cmdSearch.Enabled = Not argBoolean
    Me.lstChurch.Enabled = Not argBoolean
    Me.lstHistory.Enabled = Not argBoolean
    Me.chkOld.Enabled = Not argBoolean
    
    Me.txtStart.Enabled = argBoolean
    Me.txtEnd.Enabled = argBoolean
    Me.chkPresent.Visible = argBoolean
    Me.chkDirect.Visible = argBoolean
    Me.cboResponsibility.Enabled = argBoolean
    
    If argBoolean = True Then
        Me.cboResponsibility.listIndex = 0
    End If
End Sub

Private Sub Direct_Mode(ByVal argBoolean As Boolean)
    Me.txtManager.Visible = Not argBoolean
    Me.txtLifeNo.Enabled = argBoolean
    Me.cmdSearch_Manager.Visible = Not argBoolean
    If argBoolean Then
        Me.lblKind2 = "�����ȣ"
    Else
        Me.lblKind2 = "�̸�"
    End If
End Sub

Private Sub Date_Format(textBox As MSForms.textBox)
    Dim strDate As String
    
    If Len(Replace(textBox, "-", "")) <= 3 Then
        strDate = Replace(textBox, "-", "")
        strDate = strDate
    End If
    
    If Len(Replace(textBox, "-", "")) >= 4 And Len(Replace(textBox, "-", "")) <= 6 Then
        strDate = Replace(textBox, "-", "")
        strDate = Left(strDate, 4) & "-" & Mid(strDate, 5, Len(strDate))
    End If
    
    If Len(Replace(textBox, "-", "")) > 6 Then
        strDate = Replace(textBox, "-", "")
        strDate = Left(strDate, 4) & "-" & Mid(strDate, 5, 2) & "-" & Mid(strDate, 7, 2)
    End If
    
    textBox = strDate
End Sub
