VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Update_Church_Esta 
   Caption         =   "��ȸ���� �̷°��� ������"
   ClientHeight    =   9765.001
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   19440
   OleObjectBlob   =   "frm_Update_Church_Esta.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "frm_Update_Church_Esta"
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
Dim txtBox_Focus As MSForms.control '--//��Ŀ���� �����ϴ� ��Ʈ�� ����

Private Sub chkPresent_Click()
    Select Case Me.chkPresent.Value
        Case True '--//������ ����� �ٲٰ� ����ó��
            Me.txtEnd.BackColor = &HE0E0E0
            Me.txtEnd.Value = "����"
            Me.txtEnd.Enabled = False
        Case False '--//������ ����ó�� ���󺹱� �� ���ó�¥ ����
            Me.txtEnd.Enabled = True
            Me.txtEnd.BackColor = RGB(255, 255, 255)
            If Me.lstNomatch.listIndex = -1 Then
                Me.txtEnd = "" '--//���ġ ����Ʈ�ڽ��� ���õǾ� ���� ������
            Else
                If Me.txtEnd = "����" Then
                    Me.txtEnd.Value = Date - 1 '--//���ġ ����Ʈ�ڽ��� ���õǾ� ������
                End If
            End If
    Case Else
    End Select
End Sub

Private Sub chkPresent_Nomatch_Click()
    Select Case Me.chkPresent_Nomatch.Value
        Case True '--//������ ����� �ٲٰ� ����ó��
            Me.txtEnd_Nomatch.BackColor = &HE0E0E0
            Me.txtEnd_Nomatch.Value = "����"
            Me.txtEnd_Nomatch.Enabled = False
        Case False '--//������ ����ó�� ���󺹱� �� ���ó�¥ ����
            Me.txtEnd_Nomatch.Enabled = True
            Me.txtEnd_Nomatch.BackColor = RGB(255, 255, 255)
            If Me.lstNomatch.listIndex = -1 Then
                Me.txtEnd_Nomatch = "" '--//���ġ ����Ʈ�ڽ��� ���õǾ� ���� ������
            Else
                With Me.lstNomatch
                    If .List(.listIndex, 5) = "" Then
                        Me.txtEnd_Nomatch.Value = Date - 1 '--//���ġ ����Ʈ�ڽ��� ���õǾ� ������
                    Else
                        Me.txtEnd_Nomatch.Value = .List(.listIndex, 5)
                    End If
                End With
            End If
    Case Else
    End Select
End Sub

Private Sub cmdADD_Click()
    
    Dim argData As T_CHURCH_ESTA
    Dim result As T_RESULT
    
    '--//���ġ ��ȸ����Ʈ�ڽ� �̼��� �� ���ν��� ����
    If Me.lstNomatch.listIndex = -1 Then
        MsgBox "��Ī �ϰ��� �ϴ� ��ȸ�� ������ �ּ���."
        If Me.lstNomatch.ListCount > 0 Then
            Me.lstNomatch.listIndex = 0
        End If
        Exit Sub
    End If
    
    '--//��ȸ��� ����Ʈ�ڽ� �̼��� �� ���ν��� ����
    If Me.lstChurch.listIndex = -1 Then
        MsgBox "� ��ȸ�� ��Ī���� ������ �ּ���."
        If Me.lstChurch.ListCount > 0 Then
            Me.lstChurch.listIndex = 0
        End If
        Exit Sub
    End If
    
    '--//�ߺ��˻�
    strSql = "SELECT * FROM " & TB3 & " a WHERE a.church_sid_custom = " & SText(Me.lstChurch.List(Me.lstChurch.listIndex)) & _
            " AND IF(a.start_dt > " & SText(Me.txtStart_Nomatch) & ", a.start_dt, " & SText(Me.txtStart_Nomatch) & ") <= " & _
            "IF(a.end_dt < " & SText(IIf(Me.txtEnd_Nomatch = "����", DateSerial(9999, 12, 31), Me.txtEnd_Nomatch)) & ", a.end_dt, " & _
                SText(IIf(Me.txtEnd_Nomatch = "����", DateSerial(9999, 12, 31), Me.txtEnd_Nomatch)) & ");"
    Call makeListData(strSql, TB2)

    If cntRecord > 0 Then
        MsgBox "�ߺ��� �Ⱓ�� ������ �� �����ϴ�. ������ Ȥ�� �������� �ٽ� Ȯ���� �ּ���.", vbCritical, banner
        queryKey = LISTDATA(0, 0)
        Call returnListPosition(Me, Me.lstHistory.Name, queryKey)
        Exit Sub
    End If
    Call sbClearVariant
    
    '--//������ ��ȿ�� �˻�
    TASK_CODE = 2
    If fnData_Validation = False Then
On Error Resume Next '--//��Ʈ���� ��Ȱ��ȭ �Ǿ� ������ ���� �Ʒ� ���� ����
        txtBox_Focus.SetFocus
        txtBox_Focus.SelStart = 0
        txtBox_Focus.SelLength = Len(txtBox_Focus)
On Error GoTo 0
        Exit Sub
    End If
    
    '--//������ ������ ���� ����ü ����
    With Me.lstChurch
        argData.CHURCH_SID_CUSTOM = .List(.listIndex)
        argData.START_DT = Me.txtStart_Nomatch
        argData.END_DT = IIf(Me.txtEnd_Nomatch = "����", DateSerial(9999, 12, 31), Me.txtEnd_Nomatch)
        argData.church_sid = Me.lstNomatch.List(Me.lstNomatch.listIndex)
    End With
    
    '--//���õ� ��ȸ ��ȸ�����̷�DB�� �߰�
    connectTaskDB
    strSql = makeInsertSQL(TB3, argData)
    result.affectedCount = executeSQL("cmdAdd_Click", TB3, strSql, Me.Name, "��ȸ �����̷� �߰�")
    writeLog "cmdADD_Click", TB3, strSql, 0, Me.Name, "��ȸ �����̷� �߰�", result.affectedCount
    disconnectALL
    Call sbClearVariant
    
    '--//�޼��� �ڽ�
    MsgBox "��ȸ �����̷��� �߰� �Ǿ����ϴ�.", , banner
    
    '--//��Ī�̷� ����Ʈ�ڽ� ���ΰ�ħ
    Call lstChurch_Click
    
    '--//����Ʈ�ڽ� ������ �� ����
    With Me.lstHistory
        .listIndex = .ListCount - 1
    End With
    
    '--//���Ī ����Ʈ�ڽ� ���ΰ�ħ
    Call cmdSearch_Nomatch_Click
    
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    
    Dim result As T_RESULT
    
    '--//��ȸ �����̷� ����Ʈ�ڽ� ���ÿ��� Ȯ��
    If Me.lstHistory.listIndex = -1 Then
        MsgBox "�����̷��� �űԻ��� �ϰ��� �ϴ� ��ȸ�� ������ �ּ���."
        If Me.lstHistory.ListCount > 0 Then
            Me.lstHistory.listIndex = 0
        End If
        Exit Sub
    End If
    
    '--//���� ������ ������ ��Ȯ��
    If MsgBox("���� �����Ͻðڽ��ϱ�?", vbQuestion + vbYesNo, banner) = vbNo Then
        Exit Sub
    End If
    
    '--//���� ������ ����
    connectTaskDB
    strSql = makeDeleteSQL(TB3)
    result.affectedCount = executeSQL("cmdDelete_Click", TB3, strSql, Me.Name, "��ȸ �����̷� ����")
    writeLog "cmdDelete_Click", TB3, strSql, 0, Me.Name, "��ȸ �����̷� ����", result.affectedCount
    disconnectALL
    Call sbClearVariant
    
    '--//�޼����ڽ�
    MsgBox "�����Ͻ� �����Ͱ� ���� �Ǿ����ϴ�.", , banner
    
    '--//���Ī ����Ʈ�ڽ� ���ΰ�ħ
    With Me.lstNomatch
        queryKey = .listIndex
    End With
    Call cmdSearch_Nomatch_Click
    
    '--//���Ī ����Ʈ�ڽ� ���� ���õǾ� �ִ� �� ����
    Me.lstNomatch.listIndex = queryKey
    
    If Me.lstHistory.ListCount = 1 Then '--//���������� ���
        Me.lstHistory.Clear
        Me.lstChurch.Clear
        Me.txtChurch = ""
        Me.txtStart = ""
        Me.txtEnd = ""
    Else '--//���� ������ �ƴ� ���
        '--//�����̷� ����Ʈ�ڽ� ���ΰ�ħ
        Call lstChurch_Click
    End If

End Sub

Private Sub cmdEdit_Click()
    
    Dim result As T_RESULT
    
    '--//��ȸ �����̷� ����Ʈ�ڽ� ���ÿ��� Ȯ��
    If Me.lstHistory.listIndex = -1 Then
        MsgBox "���� �ϰ��� �ϴ� ��ȸ�� ������ �ּ���."
        If Me.lstHistory.ListCount > 0 Then
            Me.lstHistory.listIndex = 0
        End If
        Exit Sub
    End If
    
    '--//�ߺ��˻�
    With Me.lstHistory
        strSql = "SELECT * FROM " & TB3 & " a WHERE a.church_sid_custom = " & SText(Me.lstChurch.List(Me.lstChurch.listIndex)) & _
                " AND IF(a.start_dt > " & SText(Me.txtStart) & ", a.start_dt, " & SText(Me.txtStart) & ") <= " & _
                "IF(a.end_dt < " & SText(IIf(Me.txtEnd = "����", DateSerial(9999, 12, 31), Me.txtEnd)) & ", a.end_dt, " & _
                    SText(IIf(Me.txtEnd = "����", DateSerial(9999, 12, 31), Me.txtEnd)) & ")" & _
                " AND a.church_sid <> " & SText(.List(.listIndex, 4)) & ";"
    End With
    Call makeListData(strSql, TB3)

    If cntRecord > 0 Then
        MsgBox "�ߺ��� �Ⱓ�� ������ �� �����ϴ�. ������ Ȥ�� �������� �ٽ� Ȯ���� �ּ���.", vbCritical, banner
        queryKey = LISTDATA(0, 0)
        Call returnListPosition(Me, Me.lstHistory.Name, queryKey)
        Exit Sub
    End If
    Call sbClearVariant
    
    '--//������ ��ȿ�� �˻�
    TASK_CODE = 3
    If fnData_Validation = False Then
On Error Resume Next '--//��Ʈ���� ��Ȱ��ȭ �Ǿ� ������ ���� �Ʒ� ���� ����
        txtBox_Focus.SetFocus
        txtBox_Focus.SelStart = 0
        txtBox_Focus.SelLength = Len(txtBox_Focus)
On Error GoTo 0
        Exit Sub
    End If
    
    '--//���� ������ ����
    connectTaskDB
    strSql = makeUpdateSQL(TB3)
    result.affectedCount = executeSQL("cmdEdit_Click", TB3, strSql, Me.Name, "��ȸ �����̷� ����")
    writeLog "cmdEdit_Click", TB3, strSql, 0, Me.Name, "��ȸ �����̷� ����", result.affectedCount
    disconnectALL
    Call sbClearVariant
    
    '--//�޼����ڽ�
    MsgBox "������ �����Ǿ����ϴ�.", , banner
    
    '--//��Ī�̷� ����Ʈ�ڽ� ���ΰ�ħ
    Call lstChurch_Click
    
End Sub

Private Sub cmdNew_Click()
    
    Dim result As T_RESULT
    Dim argData As T_CHURCH_ESTA
    
    '--//���ġ ��ȸ����Ʈ�ڽ� �̼��� �� ���ν��� ����
    If Me.lstNomatch.listIndex = -1 Then
        MsgBox "�����̷��� �űԻ��� �ϰ��� �ϴ� ��ȸ�� ������ �ּ���."
        If Me.lstNomatch.ListCount > 0 Then
            Me.lstNomatch.listIndex = 0
        End If
        Exit Sub
    End If
    
    '--//������ ��ȿ�� �˻�
    TASK_CODE = 1
    If fnData_Validation = False Then
On Error Resume Next '--//��Ʈ���� ��Ȱ��ȭ �Ǿ� ������ ���� �Ʒ� ���� ����
        txtBox_Focus.SetFocus
        txtBox_Focus.SelStart = 0
        txtBox_Focus.SelLength = Len(txtBox_Focus)
On Error GoTo 0
        Exit Sub
    End If
    
    '--//������ ������ ���� ����ü ����
    strSql = "SELECT max(a.`��ȸĿ�����ڵ�`) FROM op_system.v_churchlist_final a;"
    Call makeListData(strSql, TB3)
    With Me.lstChurch
        argData.CHURCH_SID_CUSTOM = LISTDATA(0, 0) + 1
        argData.START_DT = Me.txtStart_Nomatch
        argData.END_DT = IIf(Me.txtEnd_Nomatch = "����", DateSerial(9999, 12, 31), Me.txtEnd_Nomatch)
        argData.church_sid = Me.lstNomatch.List(Me.lstNomatch.listIndex)
    End With
    sbClearVariant
    
    '--//���õ� ��ȸ ��ȸ�����̷�DB�� ����(Ŀ�����ڵ� �űԻ���)
    connectTaskDB
    strSql = makeInsertSQL(TB3, argData)
    result.affectedCount = executeSQL("cmdNew_Click", TB3, strSql, Me.Name, "��ȸ �����̷� �߰�")
    writeLog "cmdADD_Click", TB3, strSql, 0, Me.Name, "��ȸ �����̷� �߰�", result.affectedCount
    disconnectALL
    Call sbClearVariant
    
    '--//�޼����ڽ�
    MsgBox "��ȸ ���� �̷��� �ű� �߰� �Ǿ����ϴ�.", , banner
    
    '--//txtChurchNM�� �ű� ������ ��ȸ�� ����
    With Me.lstNomatch
        Me.txtChurch = .List(.listIndex, 1)
    End With
    
    '--//��Ī ��ȸ����Ʈ ���ΰ�ħ
    Call cmdSearch_Click
    
    '--//��Ī�̷� ����Ʈ�ڽ� ���ΰ�ħ
    Me.lstChurch.listIndex = 0
    Call lstChurch_Click
    
    '--//����Ʈ�ڽ� ������ �� ����
    With Me.lstHistory
        .listIndex = .ListCount - 1
    End With
    
    '--//���Ī ����Ʈ�ڽ� ���ΰ�ħ
    With Me.lstNomatch
        queryKey = .listIndex
    End With
    Call cmdSearch_Nomatch_Click
    If queryKey >= Me.lstNomatch.ListCount Then
        Me.lstNomatch.listIndex = Me.lstNomatch.ListCount - 1
    Else
        Me.lstNomatch.listIndex = queryKey
    End If
    
End Sub

Private Sub cmdSearch_Click()
    '--//��Ī ��ȸ����Ʈ ����Ʈ�ڽ� ���ΰ�ħ
    connectTaskDB
    strSql = makeSelectSQL(TB2)
    Call makeListData(strSql, TB2)
    If cntRecord > 0 Then
        Me.lstChurch.List = LISTDATA
    End If
    disconnectALL
    sbClearVariant
End Sub

Private Sub cmdSearch_Nomatch_Click()
    '--//���ġ ��ȸ����Ʈ ����Ʈ�ڽ� ���ΰ�ħ
    connectTaskDB
    strSql = makeSelectSQL(TB1)
    Call makeListData(strSql, TB1)
    If cntRecord > 0 Then
        Me.lstNomatch.List = LISTDATA
    Else
        Me.lstNomatch.Clear
    End If
    disconnectALL
    sbClearVariant
End Sub

Private Sub lstChurch_Click()
    '--//�űԹ�ư Ȱ��ȭ
    If Me.lstNomatch.listIndex <> -1 Then
        Me.cmdNew.Enabled = True
        Me.cmdAdd.Enabled = True
    End If
    
    '--//��Ī�̷� ����Ʈ�ڽ� ���ΰ�ħ
    If Me.lstChurch.listIndex <> -1 Then
        connectTaskDB
        strSql = makeSelectSQL(TB3)
        Call makeListData(strSql, TB3)
        If cntRecord > 0 Then
            Me.lstHistory.List = LISTDATA
        End If
        disconnectALL
        sbClearVariant
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

Private Sub lstHistory_Click()
    '--//������, ������ �ؽ�Ʈ�ڽ� Ȱ��ȭ
    Me.txtStart.Enabled = True
    Me.txtStart.BackColor = vbWhite
    Me.txtEnd.Enabled = True
    Me.txtEnd.BackColor = vbWhite
    
    '--//������ư Ȱ��ȭ
    Me.cmdEdit.Enabled = True
    
    '--//�߰�,���� ��ư Ȱ��ȭ
    Me.cmdAdd.Enabled = True
    Me.cmdDelete.Enabled = True
    
    '--//������, �����Ͽ� ����ä���
    If Me.lstHistory.ListCount > 0 Then
        With Me.lstHistory
            Me.txtStart = .List(.listIndex, 2)
            Me.txtEnd = .List(.listIndex, 3)
        End With
    End If
    
    '--//�������� 9999-12-31�̸�?
    If Me.txtEnd = "����" Then
        Me.chkPresent.Value = True
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

Private Sub lstNomatch_Click()
    
    '--//�ű�, �߰���ư Ȱ��ȭ
    If Me.lstChurch.listIndex <> -1 Then
        Me.cmdAdd.Enabled = True
    End If
    Me.cmdNew.Enabled = True
    
    '--//������, �����Ͽ� ����ä���
    If Me.lstNomatch.ListCount > 0 Then
        With Me.lstNomatch
            Me.txtStart_Nomatch = .List(.listIndex, 4)
            Me.txtEnd_Nomatch = .List(.listIndex, 5)
            If Me.txtEnd_Nomatch = "����" Then
                Me.chkPresent_Nomatch.Value = True
                Me.txtEnd_Nomatch.Enabled = False
            Else
                Me.chkPresent_Nomatch.Value = False
            End If
        End With
    End If
    
    '--//������, ������ �ؽ�Ʈ�ڽ� Ȱ��ȭ
    Me.txtStart_Nomatch.Enabled = True
    Me.txtStart_Nomatch.BackColor = vbWhite
    If Me.txtEnd_Nomatch <> "����" Then
        Me.txtEnd_Nomatch.Enabled = True
        Me.txtEnd_Nomatch.BackColor = vbWhite
    End If
    
End Sub

Private Sub lstNomatch_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Un'HookListBoxScroll
End Sub

Private Sub lstNomatch_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Me.lstNomatch.ListCount > 0 Then
        'HookListBoxScroll Me, Me.lstNomatch
    End If
End Sub

Private Sub txtChurch_Change()
    Me.txtChurch.BackColor = RGB(255, 255, 255)
End Sub

Private Sub txtChurch_Nomatch_Change()
    Me.txtChurch_Nomatch.BackColor = RGB(255, 255, 255)
End Sub

Private Sub txtEnd_Change()
    Call Date_Format(Me.txtEnd)
End Sub

Private Sub txtEnd_Nomatch_Change()
    Call Date_Format(Me.txtEnd_Nomatch)
End Sub

Private Sub txtStart_Change()
    Call Date_Format(Me.txtStart)
End Sub

Private Sub txtStart_Nomatch_Change()
    Call Date_Format(Me.txtStart_Nomatch)
End Sub

Private Sub UserForm_Initialize()
    If connIP = Empty Then setGlobalVariant '//�������� �缳��
    If checkLogin = 0 Then f_login.Show '//�α���üũ
    If checkLogin = 0 Then End '--//�α��� ���� �� ���ν��� ����
    
    '--//���ʼ���
    Me.cmdClose.Cancel = True
    TB1 = "op_system.v_churchlist_nomatch" '--//���ġ ��ȸ����Ʈ
    TB2 = "op_system.v_churchlist_final" '--//��Ī�Ϸ� ��ȸ����Ʈ
    TB3 = "op_system.db_history_church_establish" '--//��Ī�̷�
    
    '--//��Ʈ�� ����
    Me.cmdNew.Enabled = False
    Me.cmdAdd.Enabled = False
    Me.cmdDelete.Enabled = False
    Me.txtStart.Enabled = False
    Me.txtStart.BackColor = &HE0E0E0
    Me.txtEnd.Enabled = False
    Me.txtEnd.BackColor = &HE0E0E0
    Me.txtStart_Nomatch.Enabled = False
    Me.txtStart_Nomatch.BackColor = &HE0E0E0
    Me.txtEnd_Nomatch.Enabled = False
    Me.txtEnd_Nomatch.BackColor = &HE0E0E0
    
    '--//����Ʈ�ڽ� ����
    With Me.lstNomatch
        .ColumnCount = 6
        .ColumnHeads = False
        .ColumnWidths = "0,150,50,150,70,70" '��ȸ�ڵ�, ��ȸ��, ��ȸ����, ������ȸ��, ������, ������
        .Width = 531
        .TextAlign = fmTextAlignLeft
        .Font = "����"
    End With
    With Me.lstChurch
        .ColumnCount = 7
        .ColumnHeads = False
        .ColumnWidths = "0,0,120,40,0,100,70" 'Ŀ���ұ�ȸ�ڵ�.��ȸ�ڵ�,��ȸ��(ko),��ȸ����,����ȸ�ڵ�,����ȸ��,������
        .Width = 344
        .TextAlign = fmTextAlignLeft
        .Font = "����"
    End With
    With Me.lstHistory
        .ColumnCount = 11
        .ColumnHeads = False
        .ColumnWidths = "0,0,70,70,0,130,20" 'DBKey��, Ŀ���ұ�ȸ�ڵ�,������,������,��ȸ�ڵ�,��ȸ��
        .Width = 344
        .TextAlign = fmTextAlignLeft
        .Font = "����"
    End With
    
    '--//���ġ ��ȸ����Ʈ ���ΰ�ħ
    Call cmdSearch_Nomatch_Click
    
    '--//���ġ ��ȸ�˻� ��Ŀ��
    Me.txtChurch.SetFocus

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
        strSql = "SELECT * FROM " & TB1 & " a WHERE a.ovs_dept = " & SText(USER_DEPT) & " AND (a.church_nm LIKE '%" & Me.txtChurch_Nomatch & "%' OR a.main_church LIKE '%" & Me.txtChurch_Nomatch & "%');"
    Case TB2
        strSql = "SELECT a.`��ȸĿ�����ڵ�`,`��ȸ�ڵ�`,a.`��ȸ��(ko)`,a.`��ȸ����`,a.`����ȸ�ڵ�`,a.`����ȸ��`,b.`start_dt` " & _
                    "FROM " & TB2 & " a " & _
                    "LEFT JOIN " & TB3 & " b ON a.`��ȸ�ڵ�` = b.church_sid " & _
                    "WHERE a.`�����μ�` = " & SText(USER_DEPT) & " AND (a.`��ȸ��(ko)` LIKE '%" & Me.txtChurch & "%' OR a.`��ȸ��(en)` LIKE '%" & Me.txtChurch & "%' OR a.`����ȸ��` = '%" & Me.txtChurch & "%');"
    Case TB3
        With Me.lstChurch
            strSql = "SELECT a.church_esta_cd,a.church_sid_custom,a.start_dt,replace(a.end_dt,'9999-12-31','����'),a.church_sid,b.church_nm,b.church_gb FROM " & TB3 & _
                        " a LEFT JOIN op_system.db_churchlist b ON a.church_sid = b.church_sid WHERE a.church_sid_custom = " & SText(.List(.listIndex)) & "ORDER BY a.start_dt;"
        End With
    Case Else
    End Select
    makeSelectSQL = strSql
End Function
Private Function makeUpdateSQL(ByVal tableNM As String) As String
    
    With Me.lstHistory
        strSql = "UPDATE " & TB3 & " a " & _
                    "SET a.start_dt = " & SText(Me.txtStart) & ", a.end_dt = " & SText(IIf(Me.txtEnd = "����", DateSerial(9999, 12, 31), Me.txtEnd)) & _
                    " WHERE a.church_esta_cd = " & SText(.List(.listIndex)) & ";"
    End With
        
    makeUpdateSQL = strSql
End Function
Private Function makeInsertSQL(ByVal tableNM As String, _
                                argData As T_CHURCH_ESTA) As String
    strSql = "INSERT INTO " & TB3 & " VALUES(DEFAULT," & _
                SText(argData.CHURCH_SID_CUSTOM) & "," & _
                SText(argData.START_DT) & "," & _
                SText(argData.END_DT) & "," & _
                SText(argData.church_sid) & ");"
    makeInsertSQL = strSql
End Function
Private Function makeDeleteSQL(ByVal tableNM As String) As String
    With Me.lstHistory
        strSql = "DELETE FROM " & TB3 & " WHERE church_esta_cd = " & SText(.List(.listIndex)) & ";"
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
    
    Select Case TASK_CODE
        Case 1, 2 '--//�ű�,�߰��� ��
            If Not IsDate(Me.txtStart_Nomatch) Then
                MsgBox "��¥ ������ �߸� �Ǿ����ϴ�. �������� �ٽ� Ȯ���� �ּ���.", vbCritical, banner
                Set txtBox_Focus = Me.txtStart_Nomatch: fnData_Validation = False: Exit Function
            End If
            If Not IsDate(Me.txtEnd_Nomatch) And Me.txtEnd_Nomatch <> "����" Then
                MsgBox "��¥ ������ �߸� �Ǿ����ϴ�. ���������� �ٽ� Ȯ���� �ּ���.", vbCritical, banner
                Set txtBox_Focus = Me.txtEnd_Nomatch: fnData_Validation = False: Exit Function
            End If
        
        Case 3  '--//�����̷� ������ ���� ��
            If Not IsDate(Me.txtStart) Then
                MsgBox "��¥ ������ �߸� �Ǿ����ϴ�. �������� �ٽ� Ȯ���� �ּ���.", vbCritical, banner
                Set txtBox_Focus = Me.txtStart: fnData_Validation = False: Exit Function
            End If
            If Not IsDate(Me.txtEnd) And Me.txtEnd <> "����" Then
                MsgBox "��¥ ������ �߸� �Ǿ����ϴ�. ���������� �ٽ� Ȯ���� �ּ���.", vbCritical, banner
                Set txtBox_Focus = Me.txtEnd: fnData_Validation = False: Exit Function
            End If
    End Select
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
