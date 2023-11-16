VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Update_User_Authority 
   Caption         =   "����� ���Ѱ��� ������"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5775
   OleObjectBlob   =   "frm_Update_User_Authority.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "frm_Update_User_Authority"
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

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub lstUser_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Un'HookListBoxScroll
End Sub

Private Sub lstUser_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    'HookListBoxScroll Me, Me.lstUser
End Sub

Private Sub lstUser_Click()
    
    '--//��Ʈ�� ����
    Me.cmdAdd.Enabled = True
    Me.cmdDelete.Enabled = True
    
    Call refreshAuthorityTable
    
End Sub

Private Sub refreshAuthorityTable()

    strSql = makeSelectSQL(TB2)
    Call makeListData(strSql, TB2)
    If cntRecord > 0 Then
        Me.lstGivingAuthority.List = LISTDATA
    Else
        Me.lstGivingAuthority.Clear
    End If
    Me.lstGivingAuthority.listIndex = -1
    sbClearVariant
    
    strSql = makeSelectSQL(TB3)
    Call makeListData(strSql, TB3)
    If cntRecord > 0 Then
        Me.lstGivenAuthority.List = LISTDATA
    Else
        Me.lstGivenAuthority.Clear
    End If
    Me.lstGivenAuthority.listIndex = -1
    sbClearVariant
    
End Sub

Private Sub UserForm_Initialize()
    If connIP = Empty Then setGlobalVariant '//�������� �缳��
    If checkLogin = 0 Then f_login.Show '//�α���üũ
    If checkLogin = 0 Then End '--//�α��� ���� �� ���ν��� ����
    
    '--//���ʼ���
    TB1 = "common.users" '--//����� ����
    TB2 = "op_system.a_authority" '--//���Ѹ��
    TB3 = "op_system.a_auth_table" '--//����ں� �ο��� ���Ѹ��
    TB4 = "op_system.db_ovs_dept" '--//��ȸ�μ�
    
    '--//��Ʈ�� ����
    Me.cmdAdd.Enabled = False
    Me.cmdDelete.Enabled = False
    
    '--//����Ʈ�ڽ� ����
    With Me.lstUser
        .ColumnCount = 8
        .ColumnHeads = False
        .ColumnWidths = "0,120,0,0,0,0,0,250" '����id, ������, ��������, ��й�ȣ, �ʱ�ȭ, IP�ּ�, �μ�id, �μ���
        .TextAlign = fmTextAlignLeft
        .Font = "����"
    End With
    
    With Me.lstGivingAuthority
        .ColumnCount = 2
        .ColumnHeads = False
        .ColumnWidths = "0, 100" '����id, �����̸�
        .TextAlign = fmTextAlignLeft
        .Font = "����"
    End With
    
    With Me.lstGivenAuthority
        .ColumnCount = 4
        .ColumnHeads = False
        .ColumnWidths = "0, 0, 0, 100" '���id, �����id, ����id, �����̸�
        .TextAlign = fmTextAlignLeft
        .Font = "����"
    End With
    
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
    
    If cntRecord > 0 Then
        Me.lstUser.List = LISTDATA
    Else
        Me.lstUser.Clear
    End If
    Call sbClearVariant
    Me.lstUser.Enabled = True
    
End Sub

Private Sub cmdDelete_Click()
    
    Dim result As T_RESULT
    
    strSql = makeDeleteSQL(TB3)
    connectTaskDB
    result.strSql = strSql
    result.affectedCount = executeSQL("cmdDelete_Click", TB3, strSql, Me.Name, "����� ���� ����")
    writeLog "cmdDelete_Click", TB3, strSql, 0, Me.Name, "����� ���� ����"
    disconnectALL
    
    '--//����Ʈ�ڽ� ���ΰ�ħ
    Call refreshAuthorityTable
    
End Sub

Private Sub cmdADD_Click()
    
    Dim argData As T_AUTHORITY
    Dim result As T_RESULT
    
    '--//�۾��� ���� ����ü�� �� �߰�
    With Me.lstUser
        argData.USER_ID = .List(.listIndex)
    End With
    With Me.lstGivingAuthority
        argData.AUTHORITY_ID = .List(.listIndex)
    End With
    
    '--//�۾��� ���� ������ ���� �� �αױ��
    strSql = makeInsertSQL(TB3, argData)
    connectTaskDB
    result.strSql = strSql
    result.affectedCount = executeSQL("cmdADD_Clikc", TB3, strSql, Me.Name, "����� ���� �߰�")
    writeLog "cmdADD_Click", TB3, strSql, 0, Me.Name, "����� ���� �߰�", result.affectedCount
    disconnectALL
    
    '--//����Ʈ�ڽ� ���ΰ�ħ
    Call refreshAuthorityTable
    
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
                  " LEFT JOIN " & TB4 & " d" & _
                  "    ON u.user_dept = d.dept_id" & _
                  " WHERE u.user_nm LIKE '%" & Me.txtSearchName & "%';"
        ElseIf IsInArray("SECTION_CHIEF", LISTDATA) <> -1 Then
            strSql = "SELECT u.user_id, u.user_nm, u.user_gb, u.user_pw, u.pw_initialize, u.user_ip, u.user_dept, d.dept_nm" & _
                      " FROM " & TB1 & " u" & _
                      " LEFT JOIN " & TB4 & " d" & _
                      "    ON u.user_dept = d.dept_id" & _
                      " WHERE u.user_nm LIKE '%" & Me.txtSearchName & "%'" & _
                      "     AND u.user_dept = " & USER_DEPT & ";"
        End If
    Case TB2
        '--//Giving Authority ��� ��������
        Call GetUserAuthorities
        
        If IsInArray("USER_EDIT", LISTDATA) <> -1 Then
            With Me.lstUser
            strSql = "SELECT * FROM " & TB2 & " a" & _
                     " WHERE a.id" & _
                     "     NOT IN (SELECT b.authority_id FROM op_system.a_auth_table b WHERE b.user_id = " & .List(.listIndex) & ")"
            End With
        ElseIf IsInArray("SECTION_CHIEF", LISTDATA) <> -1 Then
            With Me.lstUser
            strSql = "SELECT * FROM " & TB2 & " a" & _
                     " WHERE a.id" & _
                     "     NOT IN (SELECT b.authority_id FROM op_system.a_auth_table b WHERE b.user_id = " & .List(.listIndex) & ")" & _
                     "     AND a.authority NOT IN ('USER_EDIT','SECTION_CHIEF', 'DEPT_NUM_CHANGE')"
            End With
        End If
    Case TB3
        '--//Given Authority ��� ��������
        Call GetUserAuthorities
        
        If IsInArray("USER_EDIT", LISTDATA) <> -1 Then
            With Me.lstUser
            strSql = "SELECT a.*, b.authority FROM " & TB3 & " a" & _
                     " LEFT JOIN " & TB2 & " b " & _
                     "     ON a.authority_id = b.id" & _
                     " WHERE a.user_id = " & .List(.listIndex) & ";"
            End With
        ElseIf IsInArray("SECTION_CHIEF", LISTDATA) <> -1 Then
            With Me.lstUser
            strSql = "SELECT a.*, b.authority FROM " & TB3 & " a" & _
                     " LEFT JOIN " & TB2 & " b " & _
                     "     ON a.authority_id = b.id" & _
                     " WHERE a.user_id = " & .List(.listIndex) & _
                     "     AND b.authority NOT IN ('USER_EDIT','SECTION_CHIEF', 'DEPT_NUM_CHANGE')"
            End With
        End If
    Case Else
    End Select
    makeSelectSQL = strSql
End Function

Private Function makeInsertSQL(ByVal tableNM As String, argData As T_AUTHORITY) As String
    
    Select Case tableNM
    Case TB1
    Case TB2
    Case TB3
        strSql = "INSERT INTO " & TB3 & " (user_id, authority_id) VALUES(" & _
                 argData.USER_ID & ", " & _
                 argData.AUTHORITY_ID & ");"
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
    Case TB2
    Case TB3
        With Me.lstGivenAuthority
            strSql = "DELETE a.* FROM " & TB3 & " a WHERE a.id = " & .List(.listIndex) & ";"
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

Private Sub GetUserAuthorities()

    Dim sql As String
    
    sql = "SELECT b.authority FROM op_system.a_auth_table a" & _
          " LEFT JOIN op_system.a_authority b " & _
          "     ON a.authority_id = b.id" & _
          " WHERE a.user_id = " & USER_ID & ";"
    Call makeListData(sql, "op_system.a_auth_table")
    
End Sub
