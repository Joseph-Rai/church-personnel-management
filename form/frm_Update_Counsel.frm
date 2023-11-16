VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Update_Counsel 
   Caption         =   "��� ��������"
   ClientHeight    =   9825.001
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14670
   OleObjectBlob   =   "frm_Update_Counsel.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "frm_Update_Counsel"
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
Dim txtBox_Focus As MSForms.control

Private Sub cboSearchArg_Category_Change()
    If lstPStaff.listIndex <> -1 Then
        Call lstPStaff_Click
    End If
End Sub

Private Sub cboSearchArg_Duration_Change()
    If lstPStaff.listIndex <> -1 Then
        Call lstPStaff_Click
    End If
End Sub

Private Sub cboSearchArg_Status_Change()
    If lstPStaff.listIndex <> -1 Then
        Call lstPStaff_Click
    End If
End Sub

Private Sub cmdCancel_Click()
    'Call cmdbtn_visible
    Call sbtxtBox_Init
    Call INPUTMODE(False)
    Call HideDeleteButtonByUserAuth
    Call lstCounsel_Click
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    
    Dim result As T_RESULT
    
    '--//����Ʈ�ڽ��� �������� �ʾ����� ���ν��� ����
    If Me.lstCounsel.listIndex = -1 Then
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
    result.affectedCount = executeSQL("cmdDelete_Click", TB2, strSql, Me.Name, "����̷� ����")
    writeLog "cmdDelete_Click", TB2, strSql, 0, Me.Name, "����̷� ����"
    disconnectALL
    
    '--//�޼����ڽ�
    MsgBox "�ش� �����Ͱ� �����Ǿ����ϴ�.", , banner
    
    '--//����Ʈ�ڽ� ���ΰ�ħ
    Call lstPStaff_Click
    Me.lstCounsel.listIndex = -1
    Call lstCounsel_Click
    
End Sub

Private Sub cmdEdit_Click()
    
    Dim result As T_RESULT
    
    '--//����Ʈ�ڽ��� ���õǾ� ���� ������ ���ν��� ����
    If Me.lstCounsel.listIndex = -1 Then
        MsgBox "������ �����͸� ������ �ּ���.", vbCritical, banner
        Exit Sub
    End If
    
    '--//������ ���� �ִ��� üũ
    With Me.lstCounsel
        If Me.txtInputDate = .List(.listIndex, 2) And Me.cboCategory = .List(.listIndex, 3) And Me.txtTitle = .List(.listIndex, 4) And Me.txtContent = .List(.listIndex, 5) And _
            Me.txtResult = .List(.listIndex, 6) And Me.txtRemark = .List(.listIndex, 7) And Me.cboStatus = .List(.listIndex, 8) Then
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
    strSql = makeUpdateSQL(TB2)
    result.strSql = strSql
    connectTaskDB
    result.affectedCount = executeSQL("cmdEdit_Click", TB2, strSql, Me.Name, "����̷� ������Ʈ")
    writeLog "cmdEdit_Click", TB2, strSql, 0, Me.Name, "����̷� ������Ʈ", result.affectedCount
    disconnectALL
    
    '--//�޼����ڽ�
    MsgBox "������Ʈ �Ǿ����ϴ�.", , banner
    
    '--//����Ʈ�ڽ� �ʱ�ȭ
    Call lstPStaff_Click
    If Me.lstCounsel.ListCount > 0 Then
        Me.lstCounsel.listIndex = queryKey
    End If
    
End Sub

Private Sub cmdNew_Click()
    
    '--//Ŀ�ǵ� ��ư Ȱ��ȭ�� ���� �̷� ����Ʈ�ڽ� Ŭ��
    If lstCounsel.ListCount = 0 Then
        Call lstCounsel_Click
    End If
    
    'Call cmdbtn_visible
    Me.lstCounsel.listIndex = Me.lstCounsel.ListCount - 1
    Call sbtxtBox_Init
    Call INPUTMODE(True)
    Call HideDeleteButtonByUserAuth
End Sub

Private Sub cmdADD_Click()
    
    Dim argData As T_COUNSEL
    Dim result As T_RESULT
    
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
    argData.LIFE_NO = Me.lstPStaff.List(Me.lstPStaff.listIndex, 6)
    argData.COUNSEL_DT = Me.txtInputDate
    argData.CATEGORY = Me.cboCategory
    argData.title = Me.txtTitle
    argData.CONTENT = Me.txtContent
    argData.result = Me.txtResult
    argData.REMARK = Me.txtRemark
    argData.STATUS = Me.cboStatus
    
    '--//������ ���� �� �αױ��
    strSql = makeInsertSQL(TB2, argData)
    connectTaskDB
    result.strSql = strSql
    result.affectedCount = executeSQL("cmdADD_Clikc", TB2, strSql, Me.Name, "����̷� �߰�")
    writeLog "cmdADD_Click", TB2, strSql, 0, Me.Name, "����̷� �߰�", result.affectedCount
    disconnectALL
    
    '--//��ư���� �������
    Call INPUTMODE(False)
    Call HideDeleteButtonByUserAuth
    
    '--//�޼����ڽ�
    MsgBox "�߰� �Ǿ����ϴ�.", , banner
    Call lstPStaff_Click
    Me.lstCounsel.listIndex = Me.lstCounsel.ListCount - 1
    
End Sub
Private Sub lstCounsel_Click()
    
    '--//��Ʈ�� ����
    Call controlSettingByClickListCouncel
    
    '--//����Ʈ Ŭ�� �� ������, ������, ���� ǥ��
    If Me.lstCounsel.listIndex <> -1 Then
        With Me.lstCounsel
            Me.txtTitle = .List(.listIndex, 4)
            Me.txtContent = .List(.listIndex, 5)
            Me.txtResult = .List(.listIndex, 6)
            Me.txtRemark = .List(.listIndex, 7)
            Me.txtInputDate = .List(.listIndex, 2)
            Me.cboCategory = .List(.listIndex, 3)
            Me.cboStatus = .List(.listIndex, 8)
        End With
    End If
    
End Sub
Private Sub controlSettingByClickListCouncel()

    If Me.lstCounsel.listIndex <> -1 Then
        Me.txtInputDate.Enabled = True
        Me.cmdEdit.Enabled = True
        Me.cmdAdd.Enabled = True
        Me.cmdDelete.Enabled = True
        Me.txtTitle.Enabled = True
        Me.txtContent.Enabled = True
        Me.txtResult.Enabled = True
        Me.txtRemark.Enabled = True
        Me.txtInputDate.Enabled = True
        Me.cboCategory.Enabled = True
        Me.cboStatus.Enabled = True
    Else
        Me.txtInputDate.Enabled = False
        Me.cmdEdit.Enabled = False
        Me.cmdAdd.Enabled = False
        Me.cmdDelete.Enabled = False
        Me.txtTitle.Enabled = False
        Me.txtContent.Enabled = False
        Me.txtResult.Enabled = False
        Me.txtRemark.Enabled = False
        Me.txtInputDate.Enabled = False
        Me.cboCategory.Enabled = False
        Me.cboStatus.Enabled = False
    End If

End Sub

Private Sub lstCounsel_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Un'HookListBoxScroll
End Sub

Private Sub lstCounsel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Me.lstCounsel.ListCount > 0 Then
        'HookListBoxScroll Me, Me.lstCounsel
    End If
End Sub

Private Sub lstPStaff_Click()
    
    Dim filePath As String
    Dim FileName As String
    
    '--//��Ʈ�� ����
    If Me.lstPStaff.listIndex <> -1 Then
        Me.lstCounsel.Enabled = True
        Me.cmdNew.Enabled = True
    Else
        Me.lstCounsel.Enabled = False
        Me.cmdNew.Enabled = False
    End If
    Call controlSettingByClickListCouncel
    
    '--//Ŭ�� �� �ؽ�Ʈ�ڽ� �ʱ�ȭ
    Call sbtxtBox_Init
    
    '--//�̷� ��ϻ��� ����
    With Me.lstCounsel
        .ColumnCount = 9
        .ColumnHeads = False
        .ColumnWidths = "0,0,80,80,120,0,0,0,80" '����ڵ�, �����ȣ, �����, ī�װ�, ����, ����, ���, ���, ����
        .Width = 351.75
        .Height = 91.95
        .TextAlign = fmTextAlignLeft
        .Font = "����"
    End With
    
    '--//�����߰�
    Dim strLifeNo As String
    With Me.lstPStaff
        strLifeNo = .List(.listIndex, 6)
    End With
    InsertPicToLabel Me.lblPic, strLifeNo
    
    '--//�̷¸�� ������ ä���
    Call makeSelectSQL(TB2)
    Call makeListData(strSql, TB2)
    If cntRecord > 0 Then
        Me.lstCounsel.List = LISTDATA
    Else
        Call truncateControlContent
    End If
    Call sbClearVariant
    
End Sub

Private Sub truncateControlContent()

    Me.lstCounsel.Clear
    Me.txtTitle = ""
    Me.txtContent = ""
    Me.txtResult = ""
    Me.txtRemark = ""
    Me.txtInputDate = ""
    Me.cboCategory = ""
    Me.cboStatus = ""

End Sub

Private Sub lstPStaff_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Un'HookListBoxScroll
End Sub

Private Sub lstPStaff_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Me.lstPStaff.ListCount > 0 Then
        'HookListBoxScroll Me, Me.lstPStaff
    End If
End Sub

Private Sub cboSearchArg_Duration_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Un'HookListBoxScroll
End Sub

Private Sub cboSearchArg_Duration_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Me.cboSearchArg_Duration.ListCount > 0 Then
        'HookListBoxScroll Me, Me.cboSearchArg_Duration
    End If
End Sub

Private Sub cboSearchArg_Category_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Un'HookListBoxScroll
End Sub

Private Sub cboSearchArg_Category_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Me.cboSearchArg_Category.ListCount > 0 Then
        'HookListBoxScroll Me, Me.cboSearchArg_Category
    End If
End Sub

Private Sub cboSearchArg_Status_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Un'HookListBoxScroll
End Sub

Private Sub cboSearchArg_Status_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Me.cboSearchArg_Status.ListCount > 0 Then
        'HookListBoxScroll Me, Me.cboSearchArg_Status
    End If
End Sub

Private Sub txtChurch_Change()
    Me.txtChurch.BackColor = RGB(255, 255, 255)
End Sub

Private Sub txtInputDate_Change()
    Call Date_Format(Me.txtInputDate)
End Sub

Private Sub UserForm_Initialize()
    If connIP = Empty Then setGlobalVariant '//�������� �缳��
    If checkLogin = 0 Then f_login.Show '//�α���üũ
    If checkLogin = 0 Then End '--//�α��� ���� �� ���ν��� ����
    
    '--//���ʼ���
    Me.cmdClose.Cancel = True
    TB1 = "op_system.v0_pstaff_information" '--//��ȸ�ڸ���Ʈ
    TB2 = "op_system.db_counsel" '--//����̷�
    TB3 = "op_system.v0_pstaff_information_all" '--//��ȸ�ڸ���Ʈ ��ü
    
    '--//���ѿ� ���� ��Ʈ�� ����
    Call HideDeleteButtonByUserAuth
    
    '--//��Ʈ�� ����
    Call controlSettingByClickListCouncel
    Me.cmdCancel.Visible = False
    Me.cmdAdd.Visible = False
    Me.cmdNew.Enabled = False
    
    '--//�޺��ڽ� ����
    Me.cboStatus.Clear
    Me.cboStatus.AddItem "����"
    Me.cboStatus.AddItem "�Ϸ�"
    Me.cboStatus.AddItem "���"
    
    Me.cboCategory.Clear
    Me.cboSearchArg_Category.Clear
    Call makeListData("select * from op_system.a_counsel_category;", "op_system.a_counsel_category")
    If cntRecord > 0 Then
        Me.cboCategory.List = LISTDATA
        Me.cboSearchArg_Category.List = LISTDATA
        Me.cboSearchArg_Category.AddItem "��ü", 0
    End If
    Call sbClearVariant
    
    Me.cboSearchArg_Duration.Clear
    Me.cboSearchArg_Duration.AddItem "��ü�Ⱓ"
    Me.cboSearchArg_Duration.AddItem "�ֱ� 1��"
    Me.cboSearchArg_Duration.AddItem "�ֱ� 1����"
    Me.cboSearchArg_Duration.AddItem "�ֱ� 3����"
    
    Me.cboSearchArg_Status.Clear
    Me.cboSearchArg_Status.AddItem "��ü"
    Me.cboSearchArg_Status.AddItem "����"
    Me.cboSearchArg_Status.AddItem "�Ϸ�"
    Me.cboSearchArg_Status.AddItem "���"
    
    Me.cboSearchArg_Duration.listIndex = 0
    Me.cboSearchArg_Category.listIndex = 0
    Me.cboSearchArg_Status.listIndex = 0
    
    '--//����Ʈ�ڽ� ����
    With Me.lstPStaff
        .ColumnCount = 10
        .ColumnHeads = False
        .ColumnWidths = "0,120,0,0,0,0,0,80,0,60" '��ȸ�ڵ�, ��ȸ��, ������ȸ��, ����ȸ��, ��������ȸ��, ��������, �����ȣ, �ѱ��̸�(����), �����̸�, ��å
'        .Width = 330
        .TextAlign = fmTextAlignLeft
        .Font = "����"
    End With
    
End Sub
Private Sub cmdSearch_Click()
    
    Me.lstCounsel.Clear
    
    If Me.chkAll Then
        strSql = makeSelectSQL(TB3)
        Call makeListData(strSql, TB3)
    Else
        strSql = makeSelectSQL(TB1)
        Call makeListData(strSql, TB1)
    End If
    
    If cntRecord = 0 Then
        MsgBox "��ȯ�� DB �����Ͱ� �����ϴ�.", vbInformation, banner
        Call sbClearVariant
        Exit Sub
    End If
    
    Me.lstPStaff.List = LISTDATA
    Call sbClearVariant
    Me.lstPStaff.Enabled = True
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
        strSql = "SELECT `��ȸ�ڵ�`,`��ȸ��`,`������ȸ��`,`����ȸ��`,`��������ȸ��`,`��������`,`�����ȣ`,`�ѱ��̸�(����)`,`�����̸�`,`��å`,`�����μ�`" & _
                " FROM " & TB1 & _
                " WHERE (`�����μ�` = " & USER_DEPT & ") AND (`��ȸ��` is not null)" & _
                " AND (`��ȸ��` LIKE '%" & Me.txtName & "%' OR `����ȸ��` LIKE '%" & Me.txtName & "%'" & _
                " OR `������ȸ��` LIKE '%" & Me.txtName & "%' OR `��������ȸ��` LIKE '%" & Me.txtName & "%'" & _
                " OR `�����ȣ` LIKE '%" & Me.txtName & "%' OR `�ѱ��̸�(����)` LIKE '%" & Me.txtName & "%'" & _
                " OR `�����̸�` LIKE '%" & Me.txtName & "%')" & _
                " UNION" & _
                " SELECT `��ȸ�ڵ�`,`��ȸ��`,`������ȸ��`,`����ȸ��`,`��������ȸ��`,`��������`,`����ڻ���`,`����ѱ��̸�(����)`,`��𿵹��̸�`,`�����å`,`�����μ�`" & _
                " FROM " & TB1 & _
                " WHERE (`�����μ�` = " & USER_DEPT & ") AND (`��ȸ��` is not null)" & _
                " AND (`��ȸ��` LIKE '%" & Me.txtName & "%' OR `����ȸ��` LIKE '%" & Me.txtName & "%'" & _
                " OR `������ȸ��` LIKE '%" & Me.txtName & "%' OR `��������ȸ��` LIKE '%" & Me.txtName & "%'" & _
                " OR `����ڻ���` LIKE '%" & Me.txtName & "%' OR `����ѱ��̸�(����)` LIKE '%" & Me.txtName & "%'" & _
                " OR `��𿵹��̸�` LIKE '%" & Me.txtName & "%')" & _
                " ORDER BY `��å` IS NULL ASC, FIELD(`��å`,'��ȸ��','��ȸ��븮','����','��븮���','����','�����','����ȸ������','�����ڻ��','����Ұ�����','�����ڻ��','�������1�ܰ�','�������2�ܰ�','�������3�ܰ�','�������'," & getPosition2Joining & ",'');"
    Case TB2
        
        Select Case Me.cboSearchArg_Duration.listIndex
        Case 0:
            strSql = "SELECT *" & _
                " FROM " & TB2 & _
                " WHERE life_no = " & SText(Me.lstPStaff.List(Me.lstPStaff.listIndex, 6)) & _
                " AND (category LIKE '%" & Replace(Me.cboSearchArg_Category, "��ü", "") & "%'" & _
                " AND status LIKE '%" & Replace(Me.cboSearchArg_Status, "��ü", "") & "%')" & _
                " ORDER BY counsel_dt DESC;"
        Case 1:
            strSql = "SELECT *" & _
                " FROM " & TB2 & _
                " WHERE life_no = " & SText(Me.lstPStaff.List(Me.lstPStaff.listIndex, 6)) & _
                " AND counsel_dt BETWEEN ADDDATE(CURDATE(), INTERVAL -1 WEEK) AND CURDATE()" & _
                " AND (category LIKE '%" & Replace(Me.cboSearchArg_Category, "��ü", "") & "%'" & _
                " AND status LIKE '%" & Replace(Me.cboSearchArg_Status, "��ü", "") & "%')" & _
                " ORDER BY counsel_dt DESC;"
        Case 2:
            strSql = "SELECT *" & _
                " FROM " & TB2 & _
                " WHERE life_no = " & SText(Me.lstPStaff.List(Me.lstPStaff.listIndex, 6)) & _
                " AND counsel_dt BETWEEN ADDDATE(CURDATE(), INTERVAL -1 MONTH) AND CURDATE()" & _
                " AND (category LIKE '%" & Replace(Me.cboSearchArg_Category, "��ü", "") & "%'" & _
                " AND status LIKE '%" & Replace(Me.cboSearchArg_Status, "��ü", "") & "%')" & _
                " ORDER BY counsel_dt DESC;"
        Case 3:
            strSql = "SELECT *" & _
                " FROM " & TB2 & _
                " WHERE life_no = " & SText(Me.lstPStaff.List(Me.lstPStaff.listIndex, 6)) & _
                " AND counsel_dt BETWEEN ADDDATE(CURDATE(), INTERVAL -3 MONTH) AND CURDATE()" & _
                " AND (category LIKE '%" & Replace(Me.cboSearchArg_Category, "��ü", "") & "%'" & _
                " AND status LIKE '%" & Replace(Me.cboSearchArg_Status, "��ü", "") & "%')" & _
                " ORDER BY counsel_dt DESC;"
        End Select
        
    Case TB3
        strSql = "SELECT `��ȸ�ڵ�`,`��ȸ��`,`������ȸ��`,`����ȸ��`,`��������ȸ��`,`��������`,`�����ȣ`,`�ѱ��̸�(����)`,`�����̸�`,`��å`,`�����μ�`" & _
                " FROM " & TB3 & _
                " WHERE (`�����μ�` = " & USER_DEPT & ") AND (`��ȸ��` is not null)" & _
                " AND (`��ȸ��` LIKE '%" & Me.txtName & "%' OR `����ȸ��` LIKE '%" & Me.txtName & "%'" & _
                " OR `������ȸ��` LIKE '%" & Me.txtName & "%' OR `��������ȸ��` LIKE '%" & Me.txtName & "%'" & _
                " OR `�����ȣ` LIKE '%" & Me.txtName & "%' OR `�ѱ��̸�(����)` LIKE '%" & Me.txtName & "%'" & _
                " OR `�����̸�` LIKE '%" & Me.txtName & "%')" & _
                " UNION" & _
                " SELECT `��ȸ�ڵ�`,`��ȸ��`,`������ȸ��`,`����ȸ��`,`��������ȸ��`,`��������`,`����ڻ���`,`����ѱ��̸�(����)`,`��𿵹��̸�`,`�����å`,`�����μ�`" & _
                " FROM " & TB3 & _
                " WHERE (`�����μ�` = " & USER_DEPT & ") AND (`��ȸ��` is not null)" & _
                " AND (`��ȸ��` LIKE '%" & Me.txtName & "%' OR `����ȸ��` LIKE '%" & Me.txtName & "%'" & _
                " OR `������ȸ��` LIKE '%" & Me.txtName & "%' OR `��������ȸ��` LIKE '%" & Me.txtName & "%'" & _
                " OR `����ڻ���` LIKE '%" & Me.txtName & "%' OR `����ѱ��̸�(����)` LIKE '%" & Me.txtName & "%'" & _
                " OR `��𿵹��̸�` LIKE '%" & Me.txtName & "%')" & _
                " ORDER BY `��å` IS NULL ASC, FIELD(`��å`,'��ȸ��','��ȸ��븮','����','��븮���','����','�����','����ȸ������','�����ڻ��','����Ұ�����','�����ڻ��','�������1�ܰ�','�������2�ܰ�','�������3�ܰ�','�������'," & getPosition2Joining & ",'');"
    Case Else
    End Select
    makeSelectSQL = strSql
End Function
Private Function makeUpdateSQL(ByVal tableNM As String) As String
    
    Select Case tableNM
    Case TB1
    Case TB2
        With Me.lstCounsel
            strSql = "UPDATE " & TB2 & " a " & _
                    "SET a.counsel_dt = " & SText(Me.txtInputDate) & _
                    ",a.category = " & SText(Me.cboCategory) & ",a.title = " & SText(Me.txtTitle) & _
                    ",a.content = " & SText(Me.txtContent) & ",a.result = " & SText(Me.txtResult) & _
                    ",a.remark = " & SText(Me.txtRemark) & ",a.status = " & SText(Me.cboStatus) & _
                    " WHERE a.counsel_id = " & SText(.List(.listIndex)) & ";"
        queryKey = .listIndex
        End With
    Case Else
    End Select
    makeUpdateSQL = strSql
End Function
Private Function makeInsertSQL(ByVal tableNM As String, argData As T_COUNSEL) As String
    
    Select Case tableNM
    Case TB1
    Case TB2
        strSql = "INSERT INTO " & TB2 & " VALUES(DEFAULT," & _
                    SText(argData.LIFE_NO) & "," & _
                    SText(argData.COUNSEL_DT) & "," & _
                    SText(argData.CATEGORY) & "," & _
                    SText(argData.title) & "," & _
                    SText(argData.CONTENT) & "," & _
                    SText(argData.result) & "," & _
                    SText(argData.REMARK) & "," & _
                    SText(argData.STATUS) & ");"
        queryKey = Me.lstCounsel.ListCount - 1
    Case Else
    End Select
    makeInsertSQL = strSql
End Function
Private Function makeDeleteSQL(ByVal tableNM As String) As String
    
    Select Case tableNM
    Case TB1
    Case TB2
        With Me.lstCounsel
            strSql = "DELETE FROM " & TB2 & " WHERE counsel_id = " & SText(.List(.listIndex)) & ";"
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
    Me.txtInputDate.Value = ""
    Me.txtTitle.Value = ""
    Me.txtContent.Value = ""
    Me.txtResult.Value = ""
    Me.txtRemark.Value = ""
    Me.cboCategory.Value = ""
    Me.cboStatus.Value = ""
End Sub
Private Function fnData_Validation()
'---------------------------------------
'������ �Է°��� ���� ������ ��ȿ�� �˻�
'TRUE: �̻����, FALSE: �߸���.
'---------------------------------------
    fnData_Validation = True '�����Ͱ� ��ȿ�ϴٴ� ���� �Ͽ� ����
    
    '--//��¥ ����üũ
    If Not IsDate(Me.txtInputDate) Then
        MsgBox "�ùٸ� ��¥ ���°� �ƴմϴ�. ������� �ٽ� Ȯ�� ���ּ���.", vbCritical, banner
        Set txtBox_Focus = Me.txtInputDate: fnData_Validation = False: Exit Function
    End If
    
    '--//�޺��ڽ� �� ����
    If Not (Me.cboStatus = "����" Or Me.cboStatus = "�Ϸ�" Or Me.cboStatus = "���") Then
        MsgBox "���� ���� �ùٸ��� �ʽ��ϴ�. �ٽ� �� �� Ȯ���� �ּ���.", vbCritical, banner
        Set txtBox_Focus = Me.cboStatus: fnData_Validation = False: Exit Function
    End If
    
    strSql = "SELECT * FROM op_system.a_counsel_category"
    makeListData strSql, "op_system.a_counsel_category"
    If IsInArray(Me.cboCategory, LISTDATA, True, rtnValue) = -1 Then
        MsgBox "��� ī�װ��� �ùٸ��� �ʽ��ϴ�. �ٽ� �� �� Ȯ���� �ּ���.", vbCritical, banner
        Set txtBox_Focus = Me.cboCategory: fnData_Validation = False: Exit Function
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
    
    '--//��Ʈ�� Ȱ��ȭ/��Ȱ��ȭ
    Me.lstPStaff.Enabled = Not argBoolean
    Me.lstCounsel.Enabled = Not argBoolean
    Me.cboSearchArg_Duration.Enabled = Not argBoolean
    Me.cboSearchArg_Category.Enabled = Not argBoolean
    Me.cboSearchArg_Status.Enabled = Not argBoolean
    Me.txtName.Enabled = Not argBoolean
    Me.chkAll.Enabled = Not argBoolean
    Me.cmdSearch.Enabled = Not argBoolean
    
    Me.txtTitle.Enabled = argBoolean
    Me.txtContent.Enabled = argBoolean
    Me.txtResult.Enabled = argBoolean
    Me.txtRemark.Enabled = argBoolean
    Me.txtInputDate.Enabled = argBoolean
    Me.cboCategory.Enabled = argBoolean
    Me.cboStatus.Enabled = argBoolean
    
    '--//�⺻�� ����
    If argBoolean = True Then
        Me.cboStatus.listIndex = 0
        Me.txtInputDate = Date
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
