VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Update_User 
   Caption         =   "����� ���� ������"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5325
   OleObjectBlob   =   "frm_Update_User.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "frm_Update_User"
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
Dim txtBox_Focus As MSForms.control

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdEdit_Click()
    
    Dim result As T_RESULT
    Dim argData As T_USERS
    
    '--//������ ���� �ִ��� üũ
    With Me.lstUser
        If Me.txtUsername = .List(.listIndex, 1) And Me.cboDepartment = .List(.listIndex, 6) Then
            Exit Sub
        End If
    
        If Me.cboDepartment <> .List(.listIndex, 6) Then
            Call GetUserAuthorities
            If IsInArray("DEPT_NUM_CHANGE", LISTDATA) = -1 Then
                MsgBox "�μ� ���� ������ �����ϴ�.", vbCritical, "���ѿ���"
                Me.cboDepartment.text = .List(.listIndex, 7)
                Exit Sub
            End If
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
    Dim listIndex As Integer
    With Me.lstUser
        strSql = "SELECT * FROM " & TB1 & " a WHERE a.user_id = " & SText(.List(.listIndex)) & ";"
        Call makeListData(strSql, TB1)
    
        If cntRecord > 0 Then
            strSql = makeUpdateSQL(TB1)
        End If
    End With
    
    result.strSql = strSql
    connectTaskDB
    result.affectedCount = executeSQL("cmdEdit_Click", TB1, strSql, Me.Name, "����� ������Ʈ")
    writeLog "cmdEdit_Click", TB1, strSql, 0, Me.Name, "����� ������Ʈ", result.affectedCount
    disconnectALL
    
    Call sbClearVariant
    
    '--//�޼����ڽ�
    MsgBox "������Ʈ �Ǿ����ϴ�.", , banner
    
    '--//����Ʈ�ڽ� �ʱ�ȭ
    Call cmdSearch_Click
    Call lstUser_Click
    Call setGlobalVariant
    
End Sub

Private Sub lstUser_Click()
    
    Me.cmdEdit.Enabled = True
    Me.txtUsername.Enabled = True
    Me.cboDepartment.Enabled = True
    
    '--//�ؽ�Ʈ�ڽ� �ʱ�ȭ
    Me.txtUsername = ""
    Me.cboDepartment = ""
    
    '--//��Ʈ�� ����
    If Me.lstUser.listIndex >= 0 Then
        Me.cmdDelete.Enabled = True
    Else
        Me.cmdDelete.Enabled = False
    End If
    
    '--//�ؽ�Ʈ�ڽ� �����߰�
    Dim i As Integer
    With Me.lstUser
        If .listIndex < 0 Then
            .listIndex = .ListCount - 1
        End If
        Me.txtUsername = .List(.listIndex, 1)
        For i = 0 To Me.cboDepartment.ListCount - 1
            If Me.cboDepartment.List(i, 1) = .List(.listIndex, 7) Then
                Me.cboDepartment.listIndex = i
            End If
        Next
    End With
    
End Sub

Private Sub lstUser_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Un'HookListBoxScroll
End Sub

Private Sub lstUser_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    'HookListBoxScroll Me, Me.lstUser
End Sub

Private Sub txtSearchName_Change()
    Me.txtSearchName.BackColor = RGB(255, 255, 255)
End Sub

Private Sub UserForm_Initialize()
    If connIP = Empty Then setGlobalVariant '//�������� �缳��
    If checkLogin = 0 Then f_login.Show '//�α���üũ
    If checkLogin = 0 Then End '--//�α��� ���� �� ���ν��� ����
    
    '--//���ʼ���
    Me.cmdClose.Cancel = True
    TB1 = "common.users" '--//����� ����
    TB2 = "op_system.db_ovs_dept" '--//��ȸ�μ�
    
    '--//���ѿ� ���� ��Ʈ�� ����
    Call HideDeleteButtonByUserAuth
    
    '--//��Ʈ�� ����
    Me.txtUsername.Enabled = False
    Me.cboDepartment.Enabled = False
    Me.cmdEdit.Enabled = False
    Me.cmdAdd.Enabled = False
    Me.cmdAdd.Visible = False
    Me.cmdCancel.Visible = False
    Me.cmdDelete.Enabled = False
    Me.cmdNew.Enabled = False
    
    '--//���ѿ� ���� ����
    Call GetUserAuthorities
    If IsInArray("USER_EDIT", LISTDATA) <> -1 Then
        Me.cmdNew.Enabled = True
    End If
    
    '--//����Ʈ�ڽ� ����
    With Me.lstUser
        .ColumnCount = 8
        .ColumnHeads = False
        .ColumnWidths = "0,120,0,0,0,0,0,250" '����id, ������, ��������, ��й�ȣ, �ʱ�ȭ, IP�ּ�, �μ�id, �μ���
        .TextAlign = fmTextAlignLeft
        .Font = "����"
    End With
    
    '--//�޺��ڽ� �ʱ�ȭ
    With Me.cboDepartment
        .ColumnCount = 2
        .ColumnHeads = False
        .ColumnWidths = "0, 100" '�μ�id, �μ���
    End With
    
    strSql = makeSelectSQL(TB2)
    Call makeListData(strSql, TB2)
    
    Me.cboDepartment.List = LISTDATA
    
    Call cmdSearch_Click
    Me.txtSearchName.SetFocus
    
End Sub
Private Sub cmdSearch_Click()
    
    strSql = makeSelectSQL(TB1)
    Call makeListData(strSql, TB1)
    
    If cntRecord = 0 Then
        MsgBox "��ȯ�� DB �����Ͱ� �����ϴ�.", vbInformation, banner
        Call sbClearVariant
        Exit Sub
    End If
    
    Me.lstUser.List = LISTDATA
    Call sbClearVariant
    Me.lstUser.Enabled = True
    
End Sub
Private Sub cmdCancel_Click()
    Call INPUTMODE(False)
    Call HideDeleteButtonByUserAuth
    Call sbtxtBox_Init
    Call lstUser_Click
End Sub

Private Sub cmdDelete_Click()
    
    Dim result As T_RESULT
    
    If MsgBox("������ �����͸� �����Ͻðڽ��ϱ�?", vbQuestion + vbYesNo, banner) = vbNo Then
        Exit Sub
    End If
    
    strSql = makeDeleteSQL(TB1)
    connectTaskDB
    result.strSql = strSql
    result.affectedCount = executeSQL("cmdDelete_Click", TB1, strSql, Me.Name, "����� ����")
    writeLog "cmdDelete_Click", TB1, strSql, 0, Me.Name, "����� ����"
    disconnectALL
    
    '--//�޼����ڽ�
    MsgBox "�ش� �����Ͱ� �����Ǿ����ϴ�.", , banner
    
    '--//����Ʈ�ڽ� ���ΰ�ħ
    Call cmdSearch_Click
    Call lstUser_Click
    Me.lstUser.listIndex = -1
    
End Sub

Private Sub cmdNew_Click()
    Call cmdSearch_Click
    Call lstUser_Click
    Call INPUTMODE(True)
    Call HideDeleteButtonByUserAuth
    Me.lstUser.listIndex = Me.lstUser.ListCount - 1
    Call sbtxtBox_Init
End Sub

Private Sub cmdADD_Click()
    
    Dim argData As T_USERS
    Dim result As T_RESULT
    
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
    With Me.lstUser
        argData.USER_NM = Me.txtUsername
        argData.USER_DEPT = Me.cboDepartment.List(Me.cboDepartment.listIndex, 0)
    End With
    
    '--//�۾��� ���� ������ ���� �� �αױ��
    strSql = makeInsertSQL(TB1, argData)
    connectTaskDB
    result.strSql = strSql
    result.affectedCount = executeSQL("cmdADD_Clikc", TB1, strSql, Me.Name, "����� �߰�")
    writeLog "cmdADD_Click", TB1, strSql, 0, Me.Name, "����� �߰�", result.affectedCount
    disconnectALL
    
    '--//�޼����ڽ�
    MsgBox "�߰� �Ǿ����ϴ�.", , banner
    Call lstUser_Click
    Me.lstUser.listIndex = Me.lstUser.ListCount - 1
    
    '--//��ư���� �������
    Call INPUTMODE(False)
    Call HideDeleteButtonByUserAuth
    Me.cmdCancel.Visible = False
    
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
        '--//����� ��� �ҷ�����
        Call GetUserAuthorities
                 
        If cntRecord > 0 And IsInArray("USER_EDIT", LISTDATA) <> -1 Then
            strSql = "SELECT u.user_id, u.user_nm, u.user_gb, u.user_pw, u.pw_initialize, u.user_ip, u.user_dept, d.dept_nm" & _
                      " FROM " & TB1 & " u" & _
                      " LEFT JOIN " & TB2 & " d" & _
                      "    ON u.user_dept = d.dept_id" & _
                      " WHERE u.user_nm LIKE '%" & Me.txtSearchName & "%';"
        ElseIf IsInArray("SECTION_CHIEF", LISTDATA) <> -1 Then
            strSql = "SELECT u.user_id, u.user_nm, u.user_gb, u.user_pw, u.pw_initialize, u.user_ip, u.user_dept, d.dept_nm" & _
                      " FROM " & TB1 & " u" & _
                      " LEFT JOIN " & TB2 & " d" & _
                      "    ON u.user_dept = d.dept_id" & _
                      " WHERE u.user_nm LIKE '%" & Me.txtSearchName & "%'" & _
                      "     AND u.user_dept = " & USER_DEPT & ";"
        Else
            strSql = "SELECT u.user_id, u.user_nm, u.user_gb, u.user_pw, u.pw_initialize, u.user_ip, u.user_dept, d.dept_nm" & _
                      " FROM " & TB1 & " u" & _
                      " LEFT JOIN " & TB2 & " d" & _
                      "    ON u.user_dept = d.dept_id" & _
                      " WHERE u.user_nm LIKE '%" & USER_NM & "%';"
        End If
    Case TB2
        '--//�μ� ��� �ҷ�����
        strSql = "SELECT d.dept_id, d.dept_nm FROM " & TB2 & " d WHERE d.dept_lv1 = '�ؿܼ�����';"
    Case Else
    End Select
    makeSelectSQL = strSql
End Function

Private Function makeUpdateSQL(ByVal tableNM As String) As String
    
    Select Case tableNM
    Case TB1
        strSql = "UPDATE " & TB1 & " a" & _
                 " SET a.user_nm = " & SText(Me.txtUsername) & ", a.user_dept = " & Me.cboDepartment.List(Me.cboDepartment.listIndex, 0) & _
                 " WHERE a.user_id = " & Me.lstUser.List(Me.lstUser.listIndex, 0) & ";"
    Case Else
    End Select
    makeUpdateSQL = strSql
End Function

Private Function makeInsertSQL(ByVal tableNM As String, argData As T_USERS) As String
    
    Select Case tableNM
    Case TB1
        strSql = "INSERT INTO " & TB1 & " (user_nm, user_dept) VALUES(" & _
                 SText(argData.USER_NM) & ", " & _
                 SText(argData.USER_DEPT) & ");"
    Case Else
    End Select
    makeInsertSQL = strSql
End Function

Private Function makeDeleteSQL(ByVal tableNM As String) As String
    
    Select Case tableNM
    Case TB1
        With Me.lstUser
            strSql = "DELETE a.* FROM " & TB1 & " a WHERE a.user_id = " & SText(.List(.listIndex)) & ";"
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

Private Sub HideDeleteButtonByUserAuth()
    Call GetUserAuthorities
    
    If cntRecord < 1 Then
        Exit Sub
    End If
    
    If IsInArray("DELETE_ITEM", LISTDATA) = -1 Then
        Me.cmdDelete.Visible = False
    End If
End Sub

Sub sbtxtBox_Init()
    Me.txtUsername = ""
    Me.cboDepartment.listIndex = -1
End Sub

Private Sub INPUTMODE(ByVal argBoolean As Boolean)
    Call sbtxtBox_Init
    Me.cmdNew.Visible = Not argBoolean
    Me.cmdEdit.Visible = Not argBoolean
    Me.cmdDelete.Visible = Not argBoolean
    Me.cmdClose.Visible = Not argBoolean
    Me.cmdCancel.Visible = argBoolean
    Me.cmdCancel.Enabled = argBoolean
    Me.cmdAdd.Visible = argBoolean
    Me.cmdAdd.Enabled = argBoolean
    
    Me.lstUser.Enabled = Not argBoolean
    Me.txtSearchName.Enabled = Not argBoolean
    Me.cmdSearch.Enabled = Not argBoolean
    
    Me.txtUsername.Enabled = argBoolean
    Me.cboDepartment.Enabled = argBoolean
    Me.cmdAdd.Enabled = argBoolean
End Sub

Private Function fnData_Validation()
'---------------------------------------
'������ �Է°��� ���� ������ ��ȿ�� �˻�
'TRUE: �̻����, FALSE: �߸���.
'---------------------------------------
    fnData_Validation = True '�����Ͱ� ��ȿ�ϴٴ� ���� �Ͽ� ����
    
    If IsInArray(Me.cboDepartment.Value, Me.cboDepartment.List) = -1 Then
        MsgBox "�μ� ������ �߸� �Ǿ����ϴ�.. �ٽ� Ȯ�� ���ּ���.", vbCritical, banner
        Set txtBox_Focus = Me.cboDepartment: fnData_Validation = False: Exit Function
    End If
    
End Function

Private Sub GetUserAuthorities()

    Dim sql As String
    
    sql = "SELECT b.authority FROM op_system.a_auth_table a" & _
          " LEFT JOIN op_system.a_authority b " & _
          "     ON a.authority_id = b.id" & _
          " WHERE a.user_id = " & USER_ID & ";"
    Call makeListData(sql, "op_system.a_auth_table")
    
End Sub


