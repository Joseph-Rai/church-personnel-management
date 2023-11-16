VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Update_FamilyInfo 
   Caption         =   "�������� ����������"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13260
   OleObjectBlob   =   "frm_Update_FamilyInfo.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "frm_Update_FamilyInfo"
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
Dim DUPLICATION As Boolean '--//������� �� �ߺ��� ������ �����ڵ常 ������Ʈ

Private Sub cboPosition_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Un'HookListBoxScroll
End Sub

Private Sub cboPosition_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Me.cboPosition.ListCount > 0 Then
        'HookListBoxScroll Me, Me.cboPosition
    End If
End Sub

Private Sub cboRelations_Change()
    
    '--//���� �������� �޺��ڽ��� ���� '��' Ȥ�� '����'�� �ƴϸ�
    If IsInArray(Me.cboRelations.Value, Array("��", "����"), , rtnSequence) = -1 Then
        '--//������������ ����
        strSql = "SELECT * FROM op_system.a_title a WHERE a.title NOT IN ('���', '���')"
    Else
        '--//������������ ����
        strSql = "SELECT * FROM op_system.a_title a WHERE a.title NOT IN ('�ǻ�')"
    End If
    
    Call makeListData(strSql, "op_system.a_title")
    
    Me.cboTitle.Clear
    Me.cboTitle.List = LISTDATA
    
    Call sbClearVariant
    
End Sub

Private Sub cboReligion_Change()
        
    Dim CtrlBox As MSForms.control
    
    If Me.cboReligion = "��������" Then
        For Each CtrlBox In Me.controls
            If InStr(CtrlBox.Name, "Title") > 0 Or InStr(CtrlBox.Name, "Position") > 0 Or InStr(CtrlBox.Name, "Church") > 0 Then
                CtrlBox.Visible = True
                On Error Resume Next
                Me.controls("lbl" & Right(CtrlBox.Name, Len(CtrlBox.Name) - 3)).Visible = True
                Me.controls("cmd" & Right(CtrlBox.Name, Len(CtrlBox.Name) - 3)).Visible = True
                On Error GoTo 0
            End If
        Next
    Else
        For Each CtrlBox In Me.controls
            If InStr(CtrlBox.Name, "Title") > 0 Or InStr(CtrlBox.Name, "Position") > 0 Or InStr(CtrlBox.Name, "Church") > 0 Then
                CtrlBox.Visible = False
                On Error Resume Next
                Me.controls("lbl" & Right(CtrlBox.Name, Len(CtrlBox.Name) - 3)).Visible = False
                Me.controls("cmd" & Right(CtrlBox.Name, Len(CtrlBox.Name) - 3)).Visible = False
                On Error GoTo 0
            End If
        Next
    End If
End Sub

Private Sub chkActivate_Click()
    Call Search_Mode(Me.chkActivate.Value)
End Sub

Private Sub cmdCancel_Click()
    Call Input_Mode(False)
    Call HideDeleteButtonByUserAuth
    Call lstFamily_Click
End Sub

Private Sub cmdChurch_Click()
    argShow = 1 '--//����ȸ�� �˻�
    argShow3 = 2 '--//Ȯ�� ��ư ���� �� frm_Update_FamilyInfo�� �ڷ� ����
    frm_Update_Appointment_1.Show
End Sub

Private Sub cmdClose_Click()
    
    Dim result As T_RESULT
    
    '--//������ ���� ���� �����ڵ� �ֱ�
    Select Case argShow2
    Case 1
        With Me.lstFamily
            frm_Update_PInformation.txtFamily = .List(IIf(.listIndex = -1, .ListCount - 1, .listIndex), 1)
        
            strSql = "SELECT a.family FROM op_system.db_pastoralstaff a WHERE a.lifeno = " & SText(frm_Update_PInformation.txtLifeNo) & ";"
            Call makeListData(strSql, "op_system.db_pastoralstaff")
            
            If LISTDATA(0, 0) <> .List(IIf(.listIndex = -1, .ListCount - 1, .listIndex), 1) Then
                strSql = "UPDATE op_system.db_pastoralstaff a SET a.family = " & SText(.List(.listIndex, 1)) & " WHERE a.lifeno = " & SText(frm_Update_PInformation.txtLifeNo) & ";"
                connectTaskDB
                result.strSql = strSql
                result.affectedCount = executeSQL("cmdClose_Click", "op_system.db_pastoralstaff", strSql, Me.Name, "�����ڵ� ������Ʈ")
                writeLog "cmdClose_Click", "op_system.db_pastoralstaff", strSql, 0, Me.Name, "�����ڵ� ������Ʈ"
                disconnectALL
            End If
        End With
    Case 2
        With Me.lstFamily
            frm_Update_PInformation.txtFamily_Spouse = .List(IIf(.listIndex = -1, .ListCount - 1, .listIndex), 1)
        
            strSql = "SELECT a.family FROM op_system.db_pastoralwife a WHERE a.lifeno = " & SText(frm_Update_PInformation.txtLifeNo_Spouse) & ";"
            Call makeListData(strSql, "op_system.db_pastoralwife")
            
            If LISTDATA(0, 0) <> .List(IIf(.listIndex = -1, .ListCount - 1, .listIndex), 1) Then
                strSql = "UPDATE op_system.db_pastoralwife a SET a.family = " & SText(.List(.listIndex, 1)) & " WHERE a.lifeno = " & SText(frm_Update_PInformation.txtLifeNo_Spouse) & ";"
                connectTaskDB
                result.strSql = strSql
                result.affectedCount = executeSQL("cmdClose_Click", "op_system.db_pastoralwife", strSql, Me.Name, "�����ڵ� ������Ʈ")
                writeLog "cmdClose_Click", "op_system.db_pastoralwife", strSql, 0, Me.Name, "�����ڵ� ������Ʈ"
                disconnectALL
            End If
        End With
    Case Else
    End Select
    
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    
    Dim result As T_RESULT
    
    '--//����Ʈ�ڽ��� �������� �ʾ����� ���ν��� ����
    If Me.lstFamily.listIndex = -1 Then
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
    result.affectedCount = executeSQL("cmdDelete_Click", TB2, strSql, Me.Name, "���������� ����")
    writeLog "cmdDelete_Click", TB2, strSql, 0, Me.Name, "���������� ����"
    disconnectALL
    
    '--//�޼����ڽ�
    MsgBox "�ش� �����Ͱ� �����Ǿ����ϴ�.", , banner
    
    '--//����Ʈ�ڽ� ���ΰ�ħ
    Call sbtxtBox_Init
    Call UserForm_Initialize
    
End Sub

Private Sub cmdEdit_Click()
    
    Dim result As T_RESULT
    
    DUPLICATION = False '--//�ߺ��� ���ٴ� ���� �Ͽ�
    
    '--//����Ʈ�ڽ��� ���õǾ� ���� ������ ���ν��� ����
    If Me.lstFamily.listIndex = -1 Then
        MsgBox "������ �����͸� ������ �ּ���.", vbCritical, banner
        Exit Sub
    End If
    
    '--//������ ���� �ִ��� üũ
    '--//listindex: 0:family_id, 1:family_cd, 2:��������, 3:�����ȣ 4:�ѱ��̸�, 5:�����̸�,6:��ȸ�ڵ�, 7:�Ҽӱ�ȸ, 8:����, 9:��å, 10:�������, 11:�����з�, 12:����, 13:�����ν�, 14:�޸�, 15:��������
    With Me.lstFamily
        If Me.cboRelations = Replace(.List(.listIndex, 2), "(����)", "") And Me.txtName_ko = .List(.listIndex, 4) And Me.txtName_en = .List(.listIndex, 5) And Me.txtChurch_Sid = .List(.listIndex, 6) And Me.cboTitle = .List(.listIndex, 8) And _
            Me.cboPosition = .List(.listIndex, 9) And Me.txtBirthday = .List(.listIndex, 10) And Me.txtEducation = .List(.listIndex, 11) And Me.cboReligion = .List(.listIndex, 12) And _
            Me.cboRecognition = .List(.listIndex, 13) And Me.txtMemo = .List(.listIndex, 14) And Int(Me.chkDecedent) * -1 = Int(.List(.listIndex, 15)) Then
            Exit Sub
        End If
    End With
    
    '--//�ߺ�üũ: ��ģ�� ��ģ�� �� �и� �Է� ����
    If Me.cboRelations = "��" Or Me.cboRelations = "��" Then
    '--//listindex: 0:family_id, 1:family_cd, 2:��������, 3:�����ȣ 4:�ѱ��̸�, 5:�����̸�,6:��ȸ�ڵ�, 7:�Ҽӱ�ȸ, 8:����, 9:��å, 10:�������, 11:�����з�, 12:����, 13:�����ν�, 14:�޸�, 15:��������
        With Me.lstFamily
            strSql = "SELECT * FROM " & TB2 & " a WHERE a.family_cd = " & SText(.List(.listIndex, 1)) & " AND a.relations = " & SText(Me.cboRelations) & ";"
            Call makeListData(strSql, TB2)
        End With
        
        If cntRecord > 0 Then
            If LISTDATA(0, 0) <> Me.lstFamily.List(Me.lstFamily.listIndex) Then '--//id���� �ٸ��� �����߻�
                MsgBox Me.cboRelations & "ģ�� �ߺ��� �� �����ϴ�. �ٽ� Ȯ�����ּ���.", vbCritical, banner
                Me.cboRelations.SetFocus
                Me.cboRelations.SelStart = 0
                Me.cboRelations.SelLength = Len(Me.cboRelations)
                Exit Sub
            End If
        End If
        Call sbClearVariant
    Else
        If Me.chkActivate.Value = True And Me.txtLifeNo <> "" Then
            With Me.lstFamily
                strSql = "SELECT * FROM " & TB2 & " a WHERE a.family_cd <> " & SText(.List(.listIndex, 1)) & " AND a.lifeno = " & SText(Me.txtLifeNo) & ";"
                Call makeListData(strSql, TB2)
            End With
            
            If cntRecord > 0 Then
                DUPLICATION = True '--//�ߺ�
            End If
        End If
    End If
    
    '--//������ ��ȿ�� �˻�
    If fnData_Validation = False Then
On Error Resume Next
        txtBox_Focus.SetFocus '--//������ ������ ��Ŀ��
        txtBox_Focus.SelStart = 0
        txtBox_Focus.SelLength = Len(txtBox_Focus)
On Error GoTo 0
        Exit Sub
    End If
    
    '--//SQL�� ����, ����, �αױ��
    If DUPLICATION Then
        strSql = makeUpdateSQL2(TB2)
        result.strSql = strSql
        connectTaskDB
        result.affectedCount = executeSQL("cmdEdit_Click", TB2, strSql, Me.Name, "���������� ������Ʈ")
        writeLog "cmdEdit_Click", TB2, strSql, 0, Me.Name, "���������� ������Ʈ", result.affectedCount
        disconnectALL
    Else
        strSql = makeUpdateSQL(TB2)
        result.strSql = strSql
        connectTaskDB
        result.affectedCount = executeSQL("cmdEdit_Click", TB2, strSql, Me.Name, "���������� ������Ʈ")
        writeLog "cmdEdit_Click", TB2, strSql, 0, Me.Name, "���������� ������Ʈ", result.affectedCount
        disconnectALL
    End If
    
    '--//�޼����ڽ�
    MsgBox "������Ʈ �Ǿ����ϴ�.", , banner
    
    '--//����Ʈ�ڽ� �ʱ�ȭ
    If DUPLICATION Then
        strSql = "SELECT a.family_id FROM op_system.db_familyinfo a WHERE a.lifeno = " & SText(Me.txtLifeNo)
        Call makeListData(strSql, TB2)
        queryKey = LISTDATA(0, 0)
        Call sbClearVariant
        
        Call UserForm_Initialize
        Call returnListPosition(Me, Me.lstFamily.Name, queryKey)
    Else
        Call UserForm_Initialize
        Me.lstFamily.listIndex = queryKey
    End If
    
End Sub

Private Sub cmdNew_Click()
    
    '--//Ŀ�ǵ� ��ư Ȱ��ȭ�� ���� �̷� ����Ʈ�ڽ� Ŭ��
    If lstFamily.ListCount = 0 Then
        Call lstFamily_Click
    End If
    
    Me.lstFamily.listIndex = Me.lstFamily.ListCount - 1
    Call Input_Mode(True)
    Call HideDeleteButtonByUserAuth
'    Call sbtxtBox_Init
    Call Search_Mode(False)
    
    Me.cboRelations.SetFocus
End Sub

Private Sub cmdADD_Click()
    
    Dim argData As T_FAMILY
    Dim result As T_RESULT
    Dim i As Integer
    
    DUPLICATION = False
    
    '--//�ߺ�üũ: ��ģ�� ��ģ�� �� �и� �Է� ����
    If Me.cboRelations = "��" Or Me.cboRelations = "��" Then
    '--//listindex: 0:family_id, 1:family_cd, 2:��������, 3:�����ȣ 4:�ѱ��̸�, 5:�����̸�,6:��ȸ�ڵ�, 7:�Ҽӱ�ȸ, 8:����, 9:��å, 10:�������, 11:�����з�, 12:����, 13:�����ν�, 14:�޸�, 15:��������
        With Me.lstFamily
            strSql = "SELECT * FROM " & TB2 & " a WHERE a.family_cd = " & SText(.List(.listIndex, 1)) & " AND a.relations = " & SText(Me.cboRelations) & ";"
            Call makeListData(strSql, TB2)
        End With
        
        If cntRecord > 0 Then
            MsgBox Me.cboRelations & "ģ�� �ߺ��� �� �����ϴ�. �ٽ� Ȯ�����ּ���.", vbCritical, banner
            Me.cboRelations.SetFocus
            Me.cboRelations.SelStart = 0
            Me.cboRelations.SelLength = Len(Me.cboRelations)
            Exit Sub
        End If
        Call sbClearVariant
    Else
        If Me.chkActivate.Value = True And Me.txtLifeNo <> "" Then
            With Me.lstFamily
                strSql = "SELECT * FROM " & TB2 & " a WHERE a.family_cd <> " & SText(.List(.listIndex, 1)) & " AND a.lifeno = " & SText(Me.txtLifeNo) & ";"
                Call makeListData(strSql, TB2)
            End With
            
            If cntRecord > 0 Then
                DUPLICATION = True '--//�ߺ�
            End If
        End If
    End If
    
    '--//������ ��ȿ�� �˻�
    If fnData_Validation = False Then
On Error Resume Next
        txtBox_Focus.SetFocus '--//������ ������ ��Ŀ��
        txtBox_Focus.SelStart = 0
        txtBox_Focus.SelLength = Len(txtBox_Focus)
On Error GoTo 0
        Exit Sub
    End If
    
    If DUPLICATION Then
        '--//�ش� �����ȣ�� ���� �����ڵ� ������Ʈ
        strSql = makeUpdateSQL2(TB2)
        result.strSql = strSql
        connectTaskDB
        result.affectedCount = executeSQL("cmdEdit_Click", TB2, strSql, Me.Name, "���������� ������Ʈ")
        writeLog "cmdEdit_Click", TB2, strSql, 0, Me.Name, "���������� ������Ʈ", result.affectedCount
        disconnectALL
    Else
        '--//����ü�� �� �߰�
        With Me.lstFamily
            argData.FAMILY_CD = .List(.listIndex, 1)
        End With
        If Me.txtLifeNo <> "" Then
            argData.lifeNo = Me.txtLifeNo
        End If
        argData.RELATIONS = Me.cboRelations.Value
        argData.NAME_KO = Me.txtName_ko
        argData.name_en = Me.txtName_en
        argData.church_sid = Me.txtChurch_Sid
        argData.title = Me.cboTitle
        argData.position = Me.cboPosition
        argData.Birthday = IIf(Me.txtBirthday = "", "1900-01-01", Me.txtBirthday)
        argData.Education = Me.txtEducation
        argData.RELIGION = Me.cboReligion.Value
        argData.memo = Me.txtMemo
        argData.RECOGNITION = Me.cboRecognition.Value
        argData.Suspend = Me.chkDecedent.Value
        
        '--//������ ���� �� �αױ��
        strSql = makeInsertSQL(TB2, argData)
        connectTaskDB
        result.strSql = strSql
        result.affectedCount = executeSQL("cmdADD_Clikc", TB2, strSql, Me.Name, "���������� �߰�")
        writeLog "cmdADD_Click", TB2, strSql, 0, Me.Name, "���������� �߰�", result.affectedCount
        disconnectALL
    End If
    
    '--//�޼����ڽ�
    MsgBox "�߰� �Ǿ����ϴ�.", , banner
    
    If DUPLICATION Then
        strSql = "SELECT a.family_id FROM op_system.db_familyinfo a WHERE a.lifeno = " & SText(Me.txtLifeNo)
        Call makeListData(strSql, TB2)
        queryKey = LISTDATA(0, 0)
        Call sbClearVariant
        
        Call UserForm_Initialize
        Call returnListPosition(Me, Me.lstFamily.Name, queryKey)
    Else
        Call UserForm_Initialize
        With Me.lstFamily
            strSql = "SELECT a.family_id FROM op_system.db_familyinfo a WHERE a.family_cd = " & SText(.List(.listIndex, 1))
            Call makeListData(strSql, TB2)
            For i = 0 To UBound(LISTDATA)
                queryKey = WorksheetFunction.Max(queryKey, LISTDATA(i, 0))
            Next
            Call sbClearVariant
        End With
        Call returnListPosition(Me, Me.lstFamily.Name, queryKey)
    End If
    
    '--//��ư���� �������
    Call Input_Mode(False)
    Me.chkActivate.Value = False
'    If Me.chkActivate.Value = True Then
'        Me.chkActivate.Value = False
'        Call Search_Mode(Me.chkActivate.Value)
'    End If
'    Call cmdbtn_visible
'    Call HideDeleteButtonByUserAuth
    Call lstFamily_Click
    
End Sub
Private Sub cmdSearch_Click()
    argShow = 2
    frm_Update_BCLeader_1.Show
    
    '--//�����ȣ�� ������� ������
    If Me.txtLifeNo <> "" Then
        Me.cboRelations.Enabled = True
        Me.cboRelations.BackColor = RGB(255, 255, 255)
        Me.chkDecedent.Enabled = True
        Me.chkDecedent.BackColor = RGB(255, 255, 255)
    End If
End Sub

Private Sub Frame1_Click()

End Sub

Private Sub lstFamily_Click()
    
    Dim CtrlBox As MSForms.control
    
    '--//��Ʈ�� ����
    Call sbtxtBox_Init
    Me.cmdSearch.Enabled = True
    Me.cmdEdit.Enabled = True
    Me.cmdDelete.Enabled = True
    
    '--//����Ʈ Ŭ�� �� �ؽ�Ʈ�ڽ�, �޺��ڽ��� �����߰�
    If Me.lstFamily.listIndex <> -1 Then
        With Me.lstFamily
            Me.txtLifeNo = .List(.listIndex, 3)
            If InStr(.List(.listIndex, 2), "����") > 0 Then
                Me.cboRelations = Left(.List(.listIndex, 2), InStr(.List(.listIndex, 2), "(") - 1)
            Else
                Me.cboRelations = .List(.listIndex, 2)
            End If
            Me.txtName_ko = .List(.listIndex, 4)
            Me.txtName_en = .List(.listIndex, 5)
            Me.txtChurch_Sid = .List(.listIndex, 6)
            Me.txtChurch = .List(.listIndex, 16)
            Me.cboTitle = .List(.listIndex, 8)
            Me.cboPosition = .List(.listIndex, 9)
            Me.txtBirthday = .List(.listIndex, 10)
            Me.txtEducation = .List(.listIndex, 11)
            Me.cboReligion = .List(.listIndex, 12)
            Me.cboRecognition = .List(.listIndex, 13)
            Me.txtMemo = .List(.listIndex, 14)
            Me.chkDecedent.Value = CBool(.List(.listIndex, 15))
        End With
    End If
    
    '--//txtChurch ���뿡 ���� ���� ����
    If Me.txtChurch = "" Then
        Me.txtChurch.BackColor = &HC0FFFF
    Else
        Me.txtChurch.BackColor = &HE0E0E0
    End If
    
    '--//������� ���ο� ���� ��Ʈ�� �ڽ� ����
    '1. �����ȣ�� ��ϵ� ������ �����ȣ�� ��������
    '2. �����ȣ ���� ���� �Էµ� ������ �����ȣ �����Ұ�
    '3. �����̸� �ƹ��͵� �����Ұ�
    With Me.lstFamily
        If .List(.listIndex, 3) <> "" Then
            If .List(.listIndex, 2) = "����" Then
                Call Search_Mode(True, False)
                Me.txtLifeNo.Enabled = False
                Me.txtLifeNo.BackColor = &HE0E0E0
                Me.cmdSearch.Enabled = False
                Me.cmdDelete.Enabled = False
'                Me.cboRecognition.Enabled = False
'                Me.cboRecognition.BackColor = &HE0E0E0
            Else
                Call Search_Mode(True, False)
                Me.cmdSearch.Enabled = True
                Me.cmdDelete.Enabled = True
            End If
        Else
            Call Search_Mode(False, False)
        End If
    End With
    
End Sub

Private Sub lstfamily_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Un'HookListBoxScroll
End Sub

Private Sub lstfamily_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Me.lstFamily.ListCount > 0 Then
        'HookListBoxScroll Me, Me.lstFamily
    End If
End Sub

Private Sub txtBirthday_Change()
    Call Date_Format(Me.txtBirthday)
End Sub

Private Sub UserForm_Initialize()
    
    Dim CtrlBox As MSForms.control
    Dim result As T_RESULT
    
    If connIP = Empty Then setGlobalVariant '//�������� �缳��
    If checkLogin = 0 Then f_login.Show '//�α���üũ
    If checkLogin = 0 Then End '--//�α��� ���� �� ���ν��� ����
    
    '--//���ʼ���
    Me.cmdClose.Cancel = True
    TB1 = "op_system.v_familyinfo" '--//�������� ��
    TB2 = "op_system.db_familyinfo" '--//�������� ���̺�
    
    '--//���ѿ� ���� ��Ʈ�� ����
    Call HideDeleteButtonByUserAuth
    
    '--//��Ʈ�� ����
    Me.cmdEdit.Enabled = False
    Me.cmdAdd.Visible = False
    Me.cmdCancel.Visible = False
    Me.cmdSearch.Enabled = False
    Me.cmdDelete.Enabled = False
    Me.chkActivate.Visible = False
    Me.lblInfo1.Visible = False
    For Each CtrlBox In Me.controls
        If TypeName(CtrlBox) = "TextBox" Or TypeName(CtrlBox) = "ComboBox" Or TypeName(CtrlBox) = "CheckBox" Then
                CtrlBox.Enabled = False
                CtrlBox.BackColor = &HE0E0E0
        End If
    Next
'    Me.cboPosition.List = Array("��ȸ��", "����", "��ȸ��븮", "��븮���", "����", "�ӽõ���", "�����", "����", "�������", "�������", "����ȸ������", "���������", "����Ұ�����", "�����ڻ��", "��������", "(��)��������", "�屸����", "(��)�屸����", "û������", "(��)û������", "û������", "(��)û������", "��������", "(��)��������", "�б�����", "(��)�б�����", "����������", "���屸����", "��û������", "��û������", "����������", "���б�����")
    strSql = "SELECT * FROM op_system.a_position;"
    Call makeListData(strSql, "a_position")
    Me.cboPosition.List = LISTDATA
    Me.cboPosition.AddItem "�������"
    Me.cboPosition.AddItem "����"
    Me.cboPosition.AddItem "��븮���"
    Me.cboPosition.AddItem "�����"
    Me.cboPosition.AddItem "�������"
    Me.cboPosition.AddItem "�����ڻ��"
    Me.cboPosition.AddItem "�����ڻ��"
    sbClearVariant
    
    '--//����Ʈ�ڽ� ä���
    Call Make_FirstRecord '--//ù ������ ������ ����
    
    Select Case argShow2 '--//�����ȣ �������� family_cd ��������
    Case 1
        strSql = "SELECT a.family_cd FROM " & TB1 & " a WHERE a.lifeno = " & SText(frm_Update_PInformation.txtLifeNo) & " AND a.relations NOT IN ('��','��');"
    Case 2
        strSql = "SELECT a.family_cd FROM " & TB1 & " a WHERE a.lifeno = " & SText(frm_Update_PInformation.txtLifeNo_Spouse) & " AND a.relations NOT IN ('��','��');"
    Case Else
    End Select
    Call makeListData(strSql, TB1)
    
    Call makeSelectSQL2(TB1) '--//family_cd�������� ���������� �ҷ�����
    Call makeListData(strSql, TB1)
    If cntRecord > 0 Then
        Me.lstFamily.List = LISTDATA
    End If
    Call sbClearVariant
    
    '--//����Ʈ�ڽ� ����
    With Me.lstFamily
        If .listIndex = -1 Or Me.lstFamily.Width < 500 Then '--//������ ó�� ������ ������ �ǽ�
            .ColumnCount = 17
            .ColumnHeads = False
            .ColumnWidths = "0,0,50,0,70,80,0,100,40,50,70,0,70,70,0,0,0" '������id,�����ڵ�,��������,�����ȣ,�ѱ��̸�,�����̸�,��ȸ�ڵ�,�Ҽӱ�ȸ,����,��å,�������,�����з�,����,�����ν�,�޸�,��������,��ȸǮ����
            .Width = 624.45
            .TextAlign = fmTextAlignLeft
            .Font = "����"
        
            '--//�޺��ڽ� ä���
            Me.cboRecognition.Clear
            Me.cboRelations.Clear
            Me.cboReligion.Clear
            Me.cboRecognition.List = Array("��ȣ", "����", "����")
            Me.cboRelations.List = Array("��", "��", "����", "�ڸ�")
            Me.cboReligion.List = Array("��������", "�⵶��", "õ�ֱ�", "���α�", "�ұ�", "�̽���", "���ŷ�", "��Ÿ")
        End If
    End With
    
    '--//�ڽ��� ����Ű���� ����Ʈ�ε��� ����
    Select Case argShow2
    Case 1
        strSql = "SELECT a.family_id FROM op_system.db_familyinfo a WHERE a.lifeno = " & SText(frm_Update_PInformation.txtLifeNo) & " AND a.relations NOT IN ('��','��');"
    Case 2
        strSql = "SELECT a.family_id FROM op_system.db_familyinfo a WHERE a.lifeno = " & SText(frm_Update_PInformation.txtLifeNo_Spouse) & " AND a.relations NOT IN ('��','��');"
    Case Else
    End Select
    
    Call makeListData(strSql, "op_system.db_familyinfo")
    If Me.lstFamily.listIndex = -1 Then
        queryKey = LISTDATA(0, 0)
        Call returnListPosition(Me, "lstFamily", queryKey)
    End If
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
        strSql = "SELECT * FROM " & TB1 & " a WHERE a.lifeno = " & SText(Int(LISTDATA(0, 0))) & " ORDER BY a.birthday;"
    Case TB2
    
    Case Else
    End Select
    makeSelectSQL = strSql
End Function
Private Function makeSelectSQL2(ByVal tableNM As String) As String
    
    Select Case tableNM
    Case TB1
        Select Case argShow2
        Case 1
            strSql = "SELECT a.family_id,a.family_cd,IF(a.lifeno= " & SText(frm_Update_PInformation.txtLifeNo) & ",'����',a.relations),a.lifeno,a.name_ko,a.name_en,a.church_sid,a.church_nm,a.title,a.position,a.birthday,a.education,a.religion,a.recognition,a.memo,a.suspend, a.churchFullName FROM " & TB1 & " a WHERE a.family_cd = " & SText(Int(LISTDATA(0, 0))) & " ORDER BY a.birthday;"
        Case 2
            strSql = "SELECT a.family_id,a.family_cd,IF(a.lifeno= " & SText(frm_Update_PInformation.txtLifeNo_Spouse) & ",'����',a.relations),a.lifeno,a.name_ko,a.name_en,a.church_sid,a.church_nm,a.title,a.position,a.birthday,a.education,a.religion,a.recognition,a.memo,a.suspend, a.churchFullName FROM " & TB1 & " a WHERE a.family_cd = " & SText(Int(LISTDATA(0, 0))) & " ORDER BY a.birthday;"
        Case Else
        End Select
    Case TB2
    
    Case Else
    End Select
    makeSelectSQL2 = strSql
End Function
Private Function makeUpdateSQL(ByVal tableNM As String) As String
    
    Select Case tableNM
    Case TB1
    Case TB2
        With Me.lstFamily
            '--//�����ȣ�� �ִ� ��������̸�
            If .List(.listIndex, 3) <> "" Then
                strSql = "UPDATE " & TB2 & " a " & _
                        "SET a.recognition = " & SText(Me.cboRecognition.Value) & ",a.suspend = " & SText(Int(Me.chkDecedent) * -1) & ",a.memo = " & SText(Me.txtMemo) & _
                        " WHERE a.family_id = " & SText(.List(.listIndex)) & ";"
            Else
                strSql = "UPDATE " & TB2 & " a " & _
                        "SET a.relations = " & SText(Me.cboRelations.Value) & ", a.lifeno = " & SText(Me.txtLifeNo) & ", a.name_ko = " & SText(Me.txtName_ko) & ", a.name_en = " & SText(Me.txtName_en) & _
                        ",a.title = " & SText(Me.cboTitle) & ",a.position = " & SText(Me.cboPosition) & ",a.birthday = " & IIf(Me.txtBirthday = "", "NULL", SText(Me.txtBirthday)) & ",a.education = " & SText(Me.txtEducation) & _
                        ",a.religion = " & SText(Me.cboReligion.Value) & ",a.recognition = " & SText(Me.cboRecognition.Value) & ",a.suspend = " & SText(Int(Me.chkDecedent) * -1) & ",a.church_sid = " & SText(Me.txtChurch_Sid) & ",a.memo = " & SText(Me.txtMemo) & _
                        " WHERE a.family_id = " & SText(.List(.listIndex)) & ";"
            End If
        queryKey = .listIndex
        End With
    Case Else
    End Select
    makeUpdateSQL = strSql
End Function
Private Function makeUpdateSQL2(ByVal tableNM As String) As String
    
    Select Case tableNM
    Case TB1
    Case TB2
        With Me.lstFamily
            '--//�����ȣ�� �ִ� ��������̸�
            strSql = "UPDATE " & TB2 & " a " & _
                    "SET a.family_cd = " & .List(.listIndex, 1) & " ,a.recognition = " & SText(Me.cboRecognition.Value) & ",a.suspend = " & SText(Int(Me.chkDecedent) * -1) & ",a.memo = " & SText(Me.txtMemo) & _
                    " WHERE a.lifeno = " & SText(Me.txtLifeNo) & ";"
'        queryKey = .ListIndex
        End With
    Case Else
    End Select
    makeUpdateSQL2 = strSql
End Function
Private Function makeInsertSQL(ByVal tableNM As String, argData As T_FAMILY) As String
    
    Select Case tableNM
    Case TB1
    Case TB2
        If Me.txtLifeNo = "" Then
            strSql = "INSERT INTO " & TB2 & " VALUES(DEFAULT," & _
                        SText(argData.FAMILY_CD) & "," & _
                        SText(argData.RELATIONS) & "," & _
                        SText(argData.lifeNo) & "," & _
                        SText(argData.NAME_KO) & "," & _
                        SText(argData.name_en) & "," & _
                        SText(argData.church_sid) & "," & _
                        SText(argData.title) & "," & _
                        SText(argData.position) & "," & _
                        IIf(argData.Birthday = "1900-01-01", "NULL", SText(argData.Birthday)) & "," & _
                        SText(argData.Education) & "," & _
                        SText(argData.RELIGION) & "," & _
                        SText(argData.RECOGNITION) & "," & _
                        SText(argData.memo) & "," & _
                        SText(Int(argData.Suspend) * -1) & ");"
        Else
            strSql = "INSERT INTO " & TB2 & "(family_id,family_cd,relations,lifeno,recognition,suspend) VALUES(DEFAULT," & _
                        SText(argData.FAMILY_CD) & "," & _
                        SText(argData.RELATIONS) & "," & _
                        SText(argData.lifeNo) & "," & _
                        SText(argData.RECOGNITION) & "," & _
                        SText(Int(argData.Suspend) * -1) & ");"

        End If
'        queryKey = Me.lstFamily.ListCount - 1
    Case Else
    End Select
    makeInsertSQL = strSql
End Function
Private Function makeDeleteSQL(ByVal tableNM As String) As String
    
    Select Case tableNM
    Case TB1
    Case TB2
        With Me.lstFamily
            strSql = "DELETE FROM " & TB2 & " WHERE family_id = " & SText(.List(.listIndex)) & ";"
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

Private Function fnData_Validation()
'---------------------------------------
'������ �Է°��� ���� ������ ��ȿ�� �˻�
'TRUE: �̻����, FALSE: �߸���.
'---------------------------------------
    Dim CtrlBox As MSForms.control
    
    fnData_Validation = True '�����Ͱ� ��ȿ�ϴٴ� ���� �Ͽ� ����
    
    '--//�����ȣ ������ �Է¿��� Ȯ��
    If Me.chkActivate.Value = True Then
        If Me.txtLifeNo = "" Then
            If Me.chkActivate = True Then
                MsgBox "�����ȣ�� �Է��� �ּ���.", vbCritical, banner
                fnData_Validation = False: Exit Function
            End If
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
    If Not IsDate(Me.txtBirthday) And Me.txtBirthday <> "" Then
        MsgBox "�ùٸ� ��¥ ���°� �ƴմϴ�. �������� �ٽ� Ȯ�� ���ּ���.", vbCritical, banner
        Set txtBox_Focus = Me.txtBirthday: fnData_Validation = False: Exit Function
    End If
    
    '--//�޺��ڽ� ��ȿ�� �˻�(����Ʈ�� ���� �� ���� �� ����)
    For Each CtrlBox In Me.controls
        If TypeName(CtrlBox) = "ComboBox" Then
            If IsInArray(CtrlBox.Value, CtrlBox.List, , rtnSequence) = -1 And CtrlBox <> "" And Me.chkActivate = False Then
                If Not (CtrlBox.Name = "cboRelations" And CtrlBox.Value = "����") Then
                    If CtrlBox.Name Like "*Religion*" Then
                        MsgBox "������ �߸� �����ϼ̽��ϴ�. �ٽ� Ȯ���� �ּ���.", vbCritical, banner
                        Set txtBox_Focus = CtrlBox: fnData_Validation = False: Exit Function
                    End If
                    If CtrlBox.Name Like "*Title*" Then
                        MsgBox "������ �߸� �����ϼ̽��ϴ�. �ٽ� Ȯ���� �ּ���.", vbCritical, banner
                        Set txtBox_Focus = CtrlBox: fnData_Validation = False: Exit Function
                    End If
                    If CtrlBox.Name Like "*Position*" Then
                        MsgBox "��å�� �߸� �����ϼ̽��ϴ�. �ٽ� Ȯ���� �ּ���.", vbCritical, banner
                        Set txtBox_Focus = CtrlBox: fnData_Validation = False: Exit Function
                    End If
                End If
            End If
        End If
    Next
    
    '--//�̸� ��ȿ�� �˻�
    If fnExtract(Me.txtName_ko, "E") <> "" Then
        Set txtBox_Focus = Me.txtName_ko: fnData_Validation = False: Exit Function
    End If
    If fnExtract(Me.txtName_en, "H") <> "" Then
        Set txtBox_Focus = Me.txtName_en: fnData_Validation = False: Exit Function
    End If
    
    '--//�ʼ��� üũ
    If Me.cboRelations = "" Then
        MsgBox "�������踦 �Է��� �ּ���.", vbCritical, banner
        Set txtBox_Focus = Me.cboRelations: fnData_Validation = False: Exit Function
    End If
    
End Function

Sub sbtxtBox_Init()
    
    Dim CtrlBox As MSForms.control
    
    For Each CtrlBox In Me.controls
        If TypeName(CtrlBox) = "TextBox" Or TypeName(CtrlBox) = "ComboBox" Or TypeName(CtrlBox) = "CheckBox" Then
            If CtrlBox.Name <> Me.txtChurch.Name Then '--//txtChurch�� ����
                CtrlBox.Enabled = True
            End If
            If TypeName(CtrlBox) <> "CheckBox" Then
                CtrlBox.Value = ""
            End If
            CtrlBox.BackColor = RGB(255, 255, 255)
        End If
    Next
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

Private Sub Search_Mode(ByVal argBoolean As Boolean, Optional blnClear As Boolean = True)
    
    Dim CtrlBox As MSForms.control
    
    For Each CtrlBox In Me.Frame1.controls
        If TypeName(CtrlBox) = "TextBox" Or TypeName(CtrlBox) = "ComboBox" Or TypeName(CtrlBox) = "CheckBox" Then
            If CtrlBox.Name <> Me.cboRecognition.Name Then '--//cboRecognition�� ����
                
                If CtrlBox.Name <> Me.txtChurch.Name Then '--//txtChurch�� ����
                    CtrlBox.Enabled = Not argBoolean
                    CtrlBox.BackColor = IIf(argBoolean, &HE0E0E0, RGB(255, 255, 255))
                End If
                
                If blnClear Then
                    If TypeName(CtrlBox) = "CheckBox" Then
                        CtrlBox.Value = 0
                    Else
                        CtrlBox.Value = ""
                    End If
                End If
            End If
        End If
    Next
    Me.cmdSearch.Enabled = argBoolean
    Me.txtLifeNo.Enabled = False
    Me.txtLifeNo.BackColor = IIf(argBoolean, &HC0FFFF, &HE0E0E0)
End Sub
Private Sub Input_Mode(ByVal argBoolean As Boolean)
    Call sbtxtBox_Init
    Me.cmdNew.Visible = Not argBoolean
    Me.cmdEdit.Visible = Not argBoolean
    Me.cmdDelete.Visible = Not argBoolean
    Me.cmdClose.Visible = Not argBoolean
    Me.cmdCancel.Visible = argBoolean
    Me.cmdCancel.Enabled = argBoolean
    Me.cmdAdd.Visible = argBoolean
    Me.cmdAdd.Enabled = argBoolean
    
    Me.chkActivate.Value = Not argBoolean
    Me.chkActivate.Visible = argBoolean
    Me.chkActivate.Enabled = argBoolean
    
    Me.lblInfo1.Visible = argBoolean
    
    Me.lstFamily.Enabled = Not argBoolean
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

Private Sub Make_FirstRecord()

    Dim result As T_RESULT

    Select Case argShow2
    Case 1
        strSql = "SELECT a.family_cd FROM " & TB1 & " a WHERE a.lifeno = " & SText(frm_Update_PInformation.txtLifeNo) & " AND a.relations NOT IN ('��','��');"
    Case 2
        strSql = "SELECT a.family_cd FROM " & TB1 & " a WHERE a.lifeno = " & SText(frm_Update_PInformation.txtLifeNo_Spouse) & " AND a.relations NOT IN ('��','��');"
    Case Else
    End Select
            Call makeListData(strSql, TB1)
    If cntRecord = 0 Then '--//�����ڵ尡 ������ �ű��߰�
        strSql = "SELECT MAX(a.family_cd) FROM " & TB1 & " a;"
        Call makeListData(strSql, TB2)
        Select Case argShow2
        Case 1
            strSql = "INSERT INTO " & TB2 & "(family_id,family_cd,lifeno,relations) VALUES (DEFAULT," & SText(Int(IIf(LISTDATA(0, 0) = "", 0, LISTDATA(0, 0))) + 1) & "," & SText(frm_Update_PInformation.txtLifeNo) & _
                        "," & SText(IIf(Mid(frm_Update_PInformation.txtLifeNo, 12, 1) = 1, "����", "�ڸ�")) & ");"
        Case 2
            strSql = "INSERT INTO " & TB2 & "(family_id,family_cd,lifeno,relations) VALUES (DEFAULT," & SText(Int(IIf(LISTDATA(0, 0) = "", 0, LISTDATA(0, 0))) + 1) & "," & SText(frm_Update_PInformation.txtLifeNo_Spouse) & _
                        "," & SText(IIf(Mid(frm_Update_PInformation.txtLifeNo_Spouse, 12, 1) = 1, "����", "�ڸ�")) & ");"
        Case Else
        End Select
        result.strSql = strSql
        connectTaskDB
        result.affectedCount = executeSQL("cmdEdit_Click", TB2, strSql, Me.Name, "���������� �߰�")
        writeLog "cmdEdit_Click", TB2, strSql, 0, Me.Name, "���������� �߰�", result.affectedCount
        disconnectALL
    End If
    sbClearVariant

End Sub
