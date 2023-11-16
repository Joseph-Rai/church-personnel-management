VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Update_Union_1 
   Caption         =   "����ȸ ��ϰ��� ������"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   2865
   OleObjectBlob   =   "frm_Update_Union_1.frx":0000
End
Attribute VB_Name = "frm_Update_Union_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cntlst1Col As Integer '//����Ʈ �÷� ��
Dim TB1 As String '//���� ����� DB ���̺�
Const strOrderBy As String = "currency_un ASC" '//DB���� sort_order �ʵ�
Dim caseSave As Integer '//1: �߰�, 2: ����, 3: ����(SUSPEND), 4: ��������
Dim queryKey As Integer '//����Ʈ ��ġ ��ȯ�� ���� id
Dim LISTDATA() As String '--//DB���� �޾ƿ� rs�� �迭�� ����
Dim LISTFIELD() As String '--//DB���� �޾ƿ� rs�� �ʵ带 �迭�� ����
Dim cntRecord As Integer '--//DB���� �޾ƿ� ���ڵ��� ����
Dim strSql As String '--//SQL ������
Dim UnionNM As String '--//������ ����ȸ��

Private Sub cmdADD_Click()
    '--//����ȸ �߰�
    Dim argData As T_UNION
    Dim result As T_RESULT
    
    '--//�ߺ�üũ
    strSql = "SELECT * FROM " & TB1 & " a WHERE a.suspend = 0 AND a.union_nm = " & SText(Me.txtUnion) & " AND a.ovs_dept = " & SText(USER_DEPT) & ";"
    Call makeListData(strSql, TB1)
   
    If cntRecord > 0 Then
        MsgBox "�ߺ��� ����ȸ���� �����մϴ�. �ٽ� Ȯ�����ּ���.", vbCritical, banner
        queryKey = LISTDATA(0, 0)
        Call returnListPosition(Me, Me.lstUnion.Name, queryKey)
        Exit Sub
    End If
    Call sbClearVariant
    
    '--//�۾��� ���� ������ ���� �� �αױ��
    strSql = "SELECT * FROM " & TB1 & " a WHERE a.union_nm = " & SText(Me.txtUnion) & " AND a.ovs_dept = " & SText(USER_DEPT) & ";"
    Call makeListData(strSql, TB1)
    If cntRecord > 0 Then
        strSql = makeUpdateSQL2(TB1, LISTDATA(0, 0))
    Else
        strSql = makeInsertSQL(TB1, argData)
    End If
    connectTaskDB
    result.strSql = strSql
    result.affectedCount = executeSQL("cmdADD_Clikc", TB1, strSql, Me.Name, "����ȸ �߰�")
    writeLog "cmdADD_Click", TB1, strSql, 0, Me.Name, "����ȸ �߰�", result.affectedCount
    disconnectALL
    
    '--//�޼����ڽ�
    MsgBox "�߰� �Ǿ����ϴ�.", , banner
    Call UserForm_Initialize '--//���ΰ�ħ
    Me.lstUnion.listIndex = Me.lstUnion.ListCount - 1
    
    '--//��ư���� �������
    Call HideDeleteButtonByUserAuth
End Sub

Private Sub cmdDelete_Click()
    
    Dim result As T_RESULT
    
    If MsgBox("������ �����͸� �����Ͻðڽ��ϱ�?", vbQuestion + vbYesNo, banner) = vbNo Then
        Exit Sub
    End If
    
    '--//������ ����ȸ ���İ� �ҷ�����
    With Me.lstUnion
        strSql = "SELECT a.sort_order FROM " & TB1 & " a WHERE union_cd = " & SText(.List(.listIndex)) & ";"
    End With
    Call makeListData(strSql, TB1)
    
    '--//������ ����ȸ ������
    strSql = makeDeleteSQL(TB1)
    connectTaskDB
    result.strSql = strSql
    result.affectedCount = executeSQL("cmdDelete_Click", TB1, strSql, Me.Name, "����ȸ ����")
    writeLog "cmdDelete_Click", TB1, strSql, 0, Me.Name, "����ȸ ����"
    disconnectALL
    
    '--//������ ����ȸ ���ļ��� ����
    strSql = makeDeleteSQL2(TB1)
    connectTaskDB
    result.strSql = strSql
    result.affectedCount = executeSQL("cmdDelete_Click", TB1, strSql, Me.Name, "����ȸ ����")
    writeLog "cmdDelete_Click", TB1, strSql, 0, Me.Name, "����ȸ ����"
    disconnectALL
    
    '--//�޼����ڽ�
    MsgBox "�ش� �����Ͱ� �����Ǿ����ϴ�.", , banner
    
    '--//����Ʈ�ڽ� ���ΰ�ħ
    Call UserForm_Initialize '--//���ΰ�ħ
    Me.lstUnion.listIndex = Me.lstUnion.ListCount - 1
    
End Sub

Private Sub cmdEdit_Click()
    Dim result As T_RESULT
    
    '--//������ ����ȸ�� �޾ƿ���
'    UnionNM = Application.InputBox("������ ����ȸ���� �Է��ϼ���.", banner)
'    If UnionNM = "" Then Exit Sub
    
    '--//�ߺ�üũ
    With Me.lstUnion
        strSql = "SELECT * FROM " & TB1 & " a WHERE a.union_nm = " & SText(Me.txtUnion) & " AND a.ovs_dept = " & SText(USER_DEPT) & ";"
        Call makeListData(strSql, TB1)
    End With
    
    If Me.lstUnion.listIndex < 0 Then
        Exit Sub
    End If
    
    If cntRecord > 0 Then
        MsgBox "�ߺ��� ����ȸ���� �����մϴ�. �ٽ� Ȯ�����ּ���.", vbCritical, banner
        queryKey = LISTDATA(0, 0)
        Call returnListPosition(Me, Me.lstUnion.Name, queryKey)
        Exit Sub
    End If
    
    Call sbClearVariant
    
    '--//SQL�� ����, ����, �αױ��
    strSql = makeUpdateSQL(TB1)
    result.strSql = strSql
    connectTaskDB
    result.affectedCount = executeSQL("cmdEdit_Click", TB1, strSql, Me.Name, "����ȸ�� ����")
    writeLog "cmdEdit_Click", TB1, strSql, 0, Me.Name, "����ȸ�� ����", result.affectedCount
    disconnectALL
    
    '--//�޼����ڽ�
    MsgBox "���� �Ǿ����ϴ�.", , banner
    
    '--//����Ʈ�ڽ� �ʱ�ȭ
    Call UserForm_Initialize '--//���ΰ�ħ
    Me.lstUnion.listIndex = Me.lstUnion.ListCount - 1
End Sub

Private Sub cmdMoveDown_Click()
    Dim result As T_RESULT
    Dim noMax As Integer
    
    '--//����ȸ ���ļ��� Max�� ����
    strSql = "SELECT MAX(a.sort_order) FROM op_system.a_union a WHERE a.ovs_dept = " & SText(USER_DEPT) & ";"
    Call makeListData(strSql, "op_system.a_union")
    noMax = LISTDATA(0, 0)
    
    With Me.lstUnion
        '--//������ ����ȸ sort_order �Ⱦ�
        strSql = "SELECT sort_order FROM " & TB1 & " WHERE ovs_dept = " & SText(USER_DEPT) & " AND union_cd = " & SText(.List(.listIndex)) & ";"
        Call makeListData(strSql, TB1)
        
        If LISTDATA(0, 0) < noMax Then '--//���� Sort_Order�� Max���� �ƴϸ�
            '--//������ ����ȸ sort_order 1 ����
            strSql = "UPDATE " & TB1 & " a SET a.sort_order = a.sort_order + 1 WHERE a.union_cd = " & SText(.List(.listIndex)) & ";"
            result.strSql = strSql
            connectTaskDB
            result.affectedCount = executeSQL("cmdMoveDown_Click", TB1, strSql, Me.Name, "����ȸ ���ļ��� ����")
            writeLog "cmdMoveDown_Click", TB1, strSql, 0, Me.Name, "����ȸ ���ļ��� ����", result.affectedCount
            disconnectALL
            
            '--//���� ����ȸ sort_order 1 ����
            strSql = "UPDATE " & TB1 & " a SET a.sort_order = a.sort_order - 1 WHERE a.ovs_dept = " & SText(USER_DEPT) & " AND a.sort_order = " & SText(LISTDATA(0, 0) + 1) & " AND a.union_cd <> " & SText(.List(.listIndex)) & ";"
            result.strSql = strSql
            connectTaskDB
            result.affectedCount = executeSQL("cmdMoveDown_Click", TB1, strSql, Me.Name, "����ȸ ���ļ��� ����")
            writeLog "cmdMoveDown_Click", TB1, strSql, 0, Me.Name, "����ȸ ���ļ��� ����", result.affectedCount
            disconnectALL
        End If
    End With
    
    '--//����Ʈ�ڽ� ���ΰ�ħ
    Call UserForm_Initialize '--//���ΰ�ħ
'    If Me.lstUnion.ListIndex < Me.lstUnion.ListCount - 1 Then
'        Me.lstUnion.ListIndex = Me.lstUnion.ListIndex + 1
'    End If
    
End Sub

Private Sub cmdMoveUp_Click()
    
    Dim result As T_RESULT
    
    With Me.lstUnion
        '--//������ ����ȸ sort_order �Ⱦ�
        strSql = "SELECT sort_order FROM " & TB1 & " WHERE ovs_dept = " & SText(USER_DEPT) & " AND union_cd = " & SText(.List(.listIndex)) & ";"
        Call makeListData(strSql, TB1)
        
        If LISTDATA(0, 0) > 1 Then '--//���� Sort_Order�� Min���� �ƴϸ�
            '--//������ ����ȸ sort_order 1 ����
            strSql = "UPDATE " & TB1 & " a SET a.sort_order = a.sort_order - 1 WHERE a.union_cd = " & SText(.List(.listIndex)) & ";"
            result.strSql = strSql
            connectTaskDB
            result.affectedCount = executeSQL("cmdMoveDown_Click", TB1, strSql, Me.Name, "����ȸ ���ļ��� ����")
            writeLog "cmdMoveDown_Click", TB1, strSql, 0, Me.Name, "����ȸ ���ļ��� ����", result.affectedCount
            disconnectALL
            
            '--//���� ����ȸ sort_order 1 ����
            strSql = "UPDATE " & TB1 & " a SET a.sort_order = a.sort_order + 1 WHERE a.ovs_dept = " & SText(USER_DEPT) & " AND a.sort_order = " & SText(LISTDATA(0, 0) - 1) & " AND a.union_cd <> " & SText(.List(.listIndex)) & ";"
            result.strSql = strSql
            connectTaskDB
            result.affectedCount = executeSQL("cmdMoveDown_Click", TB1, strSql, Me.Name, "����ȸ ���ļ��� ����")
            writeLog "cmdMoveDown_Click", TB1, strSql, 0, Me.Name, "����ȸ ���ļ��� ����", result.affectedCount
            disconnectALL
        End If
    End With
    
    '--//����Ʈ�ڽ� ���ΰ�ħ
    Call UserForm_Initialize '--//���ΰ�ħ
'    If Me.lstUnion.ListIndex > 0 Then
'        Me.lstUnion.ListIndex = Me.lstUnion.ListIndex - 1
'    End If
    
End Sub

Private Sub lstUnion_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Un'HookListBoxScroll
End Sub

Private Sub lstUnion_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Me.lstUnion.ListCount Then
        'HookListBoxScroll Me, Me.lstUnion
    End If
End Sub

Private Sub UserForm_Initialize()

    '--//���ʼ���
    Me.cmdClose.Cancel = True
    TB1 = "op_system.a_union" '--//��ȸ����Ʈ
    
    '--//��Ʈ�Ѽ���
    Me.cmdDelete.Enabled = False
    Me.cmdAdd.Enabled = False
    Me.txtUnion = ""
    
    '--//����Ʈ�ڽ� ����
    With Me.lstUnion
        .ColumnCount = 2
        .ColumnHeads = False
        .ColumnWidths = "0,120" '����ȸ�ڵ�, ����ȸ��
        .TextAlign = fmTextAlignLeft
        .Font = "����"
    End With
    
    '--//ǥ����ġ
    Me.Top = frm_Update_Union.Top
    Me.Left = frm_Update_Union.Left + frm_Update_Union.Width
    
    '--//����ȸ ����߰�
    strSql = makeSelectSQL(TB1)
    Call makeListData(strSql, TB1)
    If cntRecord > 0 Then
        Me.lstUnion.List = LISTDATA
    Else
        Me.lstUnion.Clear
    End If
    Call sbClearVariant
    
    Me.txtUnion.SetFocus
End Sub

Private Sub txtUnion_Change()
    Me.txtUnion.BackColor = RGB(255, 255, 255)
    Me.cmdAdd.Enabled = True
End Sub
Private Sub lstUnion_Click()
    Me.cmdDelete.Enabled = True
    Me.txtUnion = Me.lstUnion.List(Me.lstUnion.listIndex, 1)
End Sub
Private Sub cmdClose_Click()
    Unload frm_Update_Union_1
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
        strSql = "SELECT * " & _
                    "FROM " & TB1 & " a WHERE a.suspend = 0 AND a.ovs_dept = " & SText(USER_DEPT) & " ORDER BY a.sort_order;"
    Case Else
    End Select
    makeSelectSQL = strSql
End Function
Private Function makeUpdateSQL(ByVal tableNM As String) As String
    
    Select Case tableNM
    Case TB1
        With Me.lstUnion
            strSql = "UPDATE " & TB1 & " a " & _
                    "SET a.union_nm = " & SText(Me.txtUnion) & ",a.suspend = 0" & _
                    " WHERE a.union_cd = " & SText(.List(.listIndex)) & " AND a.ovs_dept = " & SText(USER_DEPT) & ";"
        queryKey = .listIndex
        End With
    Case Else
    End Select
    makeUpdateSQL = strSql
End Function
Private Function makeUpdateSQL2(ByVal tableNM As String, ByVal UNION_CD As Long) As String
    
    Select Case tableNM
    Case TB1
        With Me.lstUnion
            strSql = "SELECT MAX(a.sort_order) FROM op_system.a_union a WHERE a.ovs_dept = " & SText(USER_DEPT) & ";"
            Call makeListData(strSql, "op_system.a_union")
            
            strSql = "UPDATE " & TB1 & " a " & _
                    "SET a.union_nm = " & SText(Me.txtUnion) & ",a.suspend = 0" & ",a.sort_order = " & SText(LISTDATA(0, 0) + 1) & _
                    " WHERE a.union_cd = " & SText(UNION_CD) & " AND a.ovs_dept = " & SText(USER_DEPT) & ";"
        queryKey = .listIndex
        End With
    Case Else
    End Select
    makeUpdateSQL2 = strSql
End Function
Private Function makeInsertSQL(ByVal tableNM As String, argData As T_UNION) As String
    
    Select Case tableNM
    Case TB1
        strSql = "SELECT MAX(a.sort_order) FROM op_system.a_union a WHERE a.ovs_dept = " & SText(USER_DEPT) & ";"
        Call makeListData(strSql, "op_system.a_union")
        
        If cntRecord > 0 Then
            strSql = "INSERT INTO " & TB1 & " VALUES(DEFAULT," & _
                        SText(Me.txtUnion) & ",0," & SText(USER_DEPT) & "," & SText(IIf(LISTDATA(0, 0) = "", 0, LISTDATA(0, 0)) + 1) & ");"
        Else
            strSql = "INSERT INTO " & TB1 & " VALUES(DEFAULT," & _
                        SText(Me.txtUnion) & ",0," & SText(USER_DEPT) & ",1);"
        End If
        queryKey = Me.lstUnion.ListCount - 1
    Case Else
    End Select
    makeInsertSQL = strSql
End Function
Private Function makeDeleteSQL(ByVal tableNM As String) As String
    
    Select Case tableNM
    Case TB1
        With Me.lstUnion
            strSql = "UPDATE " & TB1 & " SET suspend = 1,sort_order = 0 WHERE union_cd = " & SText(.List(.listIndex)) & ";"
        End With
    Case Else
    End Select
    makeDeleteSQL = strSql
End Function
Private Function makeDeleteSQL2(ByVal tableNM As String) As String
    
    Select Case tableNM
    Case TB1
        With Me.lstUnion
            strSql = "UPDATE " & TB1 & " SET sort_order = sort_order - 1 WHERE ovs_dept = " & SText(USER_DEPT) & " AND sort_order > " & SText(LISTDATA(0, 0)) & ";"
        End With
    Case Else
    End Select
    makeDeleteSQL2 = strSql
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

Private Sub GetUserAuthorities()

    Dim sql As String
    
    sql = "SELECT b.authority FROM op_system.a_auth_table a" & _
          " LEFT JOIN op_system.a_authority b " & _
          "     ON a.authority_id = b.id" & _
          " WHERE a.user_id = " & USER_ID & ";"
    Call makeListData(sql, "op_system.a_auth_table")
    
End Sub




