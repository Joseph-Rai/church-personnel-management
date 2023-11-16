VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Update_History 
   Caption         =   "��ȸ�̷� ����������"
   ClientHeight    =   6750
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   7380
   OleObjectBlob   =   "frm_Update_History.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "frm_Update_History"
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
Dim txtBox_Focus As MSForms.textBox

Private Sub lstChurch_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Un'HookListBoxScroll
End Sub

Private Sub lstChurch_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Me.lstChurch.ListCount > 0 Then
        'HookListBoxScroll Me, Me.lstChurch
    End If
End Sub

Private Sub lstHistory_Click()
    
    '--//��Ʈ�Ѽ���
    If Me.lstHistory.listIndex <> -1 Then
        Me.cmdAdd.Enabled = True
        Me.cmdDelete.Enabled = True
        Me.cmdEdit.Enabled = True
        Me.txtDate.Enabled = True
        Me.txtHistory.Enabled = True
    Else
        Me.cmdAdd.Enabled = False
        Me.cmdDelete.Enabled = False
        Me.cmdEdit.Enabled = False
        Me.txtDate.Enabled = False
        Me.txtHistory.Enabled = False
    End If
    
    '--//��Ʈ�� ���� ä���
    If Me.lstHistory.listIndex <> -1 Then
        With Me.lstHistory
            Me.txtDate = .List(.listIndex, 2)
            Me.txtHistory = .List(.listIndex, 3)
        End With
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

Private Sub txtDate_Change()
    Call Date_Format(Me.txtDate)
End Sub

Private Sub UserForm_Initialize()
    
    If connIP = Empty Then setGlobalVariant '//�������� �缳��
    If checkLogin = 0 Then f_login.Show '//�α���üũ
    If checkLogin = 0 Then End '--//�α��� ���� �� ���ν��� ����
    
    '--//���ʼ���
    Me.cmdClose.Cancel = True
    TB1 = "op_system.db_churchlist_custom" '--//��ȸ����Ʈ
    TB2 = "op_system.db_history_church" '--//��ȸ��Ȳ
    
    '--//���ѿ� ���� ��Ʈ�� ����
    Call HideDeleteButtonByUserAuth
    
    '--//��Ʈ�� ����
    Me.cmdAdd.Enabled = False
    Me.cmdAdd.Visible = False
    Me.cmdCancel.Visible = False
    Me.cmdNew.Enabled = False
    Me.cmdDelete.Enabled = False
    Me.cmdEdit.Enabled = False
    Me.txtDate.Enabled = False
    Me.txtHistory.Enabled = False
    
    '--//����Ʈ�ڽ� ����
    With Me.lstChurch
        .ColumnCount = 4
        .ColumnHeads = False
        .ColumnWidths = "0,150,50,150" '��ȸ�ڵ�, ��ȸ��, ��ȸ����, ������ȸ��
'        .Width = 330
        .TextAlign = fmTextAlignLeft
        .Font = "����"
    End With
    
    With Me.lstHistory
        .ColumnCount = 4
        .ColumnHeads = False
        .ColumnWidths = "0,0,60,280" '�̷��ڵ�, ��ȸ�ڵ�, ��¥, ��ȸ�̷�
'        .Width = 330
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
Private Sub cmdCancel_Click()
    'Call cmdbtn_visible
    Call INPUTMODE(False)
    Call HideDeleteButtonByUserAuth
    Call lstHistory_Click
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
    result.affectedCount = executeSQL("cmdDelete_Click", TB2, strSql, Me.Name, "��ȸ�̷� ����")
    writeLog "cmdDelete_Click", TB2, strSql, 0, Me.Name, "��ȸ�̷� ����"
    disconnectALL
    
    '--//�޼����ڽ�
    MsgBox "�ش� �����Ͱ� �����Ǿ����ϴ�.", , banner
    
    '--//����Ʈ�ڽ� ���ΰ�ħ
    Call lstChurch_Click
    Me.lstHistory.listIndex = -1
    
End Sub

Private Sub cmdEdit_Click()
    
    Dim result As T_RESULT
    
    '--//������ ���� �ִ��� üũ
    With Me.lstHistory
        If .listIndex > -1 Then
            If Me.txtDate = .List(.listIndex, 2) And Me.txtHistory = .List(.listIndex, 3) Then
                Exit Sub
            End If
        Else
            Exit Sub '--//����Ʈ�� ���õ��� �ʾ����� ���ν��� ����
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
    result.affectedCount = executeSQL("cmdEdit_Click", TB2, strSql, Me.Name, "��ȸ�̷� ������Ʈ")
    writeLog "cmdEdit_Click", TB2, strSql, 0, Me.Name, "��ȸ�̷� ������Ʈ", result.affectedCount
    disconnectALL

    
    '--//�޼����ڽ�
    MsgBox "������Ʈ �Ǿ����ϴ�.", , banner
    
    '--//����Ʈ�ڽ� �ʱ�ȭ
    Call lstChurch_Click
    Me.lstHistory.listIndex = queryKey
    
End Sub

Private Sub cmdNew_Click()
    'Call cmdbtn_visible
    Call INPUTMODE(True)
    Call HideDeleteButtonByUserAuth
    Me.lstHistory.listIndex = Me.lstHistory.ListCount - 1
    Call sbtxtBox_Init
End Sub

Private Sub cmdADD_Click()
    
    Dim argData As T_HISTORY_CHURCH
    Dim result As T_RESULT
    
    '--//�ߺ�üũ
    With Me.lstChurch
        strSql = "SELECT * FROM " & TB2 & " a WHERE a.church_sid = " & SText(.List(.listIndex)) & _
                    "AND a.his_dt = " & SText(Me.txtDate) & " AND a.history = " & SText(Me.txtHistory) & ";"
        Call makeListData(strSql, TB2)
    End With
    
    If cntRecord > 0 Then
        MsgBox "�ߺ��� ������ �����մϴ�. �Է� ���� �ٽ� Ȯ�����ּ���.", vbCritical, banner
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
    
    '--//�۾��� ���� ����ü�� �� �߰�
    argData.church_sid = Me.lstChurch.List(Me.lstChurch.listIndex)
    argData.HIS_DT = IIf(Me.txtDate = "", "1900-01-01", Me.txtDate)
    argData.HISTORY = Me.txtHistory
    
    If Me.txtDate = "" And Me.txtHistory = "" Then
        MsgBox "���� �Է��� �ּ���.", vbCritical, banner
        Exit Sub
    End If
    
    '--//������ ���� �� �αױ��
    strSql = makeInsertSQL(TB2, argData)
    connectTaskDB
    result.strSql = strSql
    result.affectedCount = executeSQL("cmdADD_Clikc", TB2, strSql, Me.Name, "��ȸ�̷� �߰�")
    writeLog "cmdADD_Click", TB2, strSql, 0, Me.Name, "��ȸ�̷� �߰�", result.affectedCount
    disconnectALL

    '--//�޼����ڽ�
    MsgBox "�߰� �Ǿ����ϴ�.", , banner
    Call lstChurch_Click
    Me.lstHistory.listIndex = Me.lstHistory.ListCount - 1
    
    '--//��ư���� �������
    'Call cmdbtn_visible
    Call INPUTMODE(False)
    Call HideDeleteButtonByUserAuth
    
End Sub
Private Sub lstChurch_Click()
    
    '--//��ȸ ���ÿ��� �Ǵ�
    If Me.lstChurch.listIndex = -1 Then
        MsgBox "��ȸ�� �����ϼ���.", vbCritical, "����"
        Exit Sub
    End If
    
    '--//��Ʈ�� ����
    If Me.lstChurch.listIndex <> -1 Then
        Me.cmdNew.Enabled = True
        Me.lstHistory.Enabled = True
    Else
        Me.cmdNew.Enabled = False
        Me.lstHistory.Enabled = False
    End If
    
    '--//�ؽ�Ʈ�ڽ� �ʱ�ȭ
    Call sbtxtBox_Init
    
    '--//��ȸ�̷µ����� �߰�
    Erase LISTDATA
    strSql = makeSelectSQL(TB2)
    Call makeListData(strSql, TB2)
On Error Resume Next '--//�̷� �����Ͱ� ���� �������� ���� ��� ����
    Me.lstHistory.List = LISTDATA
    If err.Number <> 0 Then
        Me.lstHistory.Clear
    End If
On Error GoTo 0
    Call sbClearVariant
    
    With Me.lstHistory
        .listIndex = .ListCount - 1
    End With
    
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
                        "WHERE a.ovs_dept = " & USER_DEPT & " AND a.church_gb <> 'MM' AND (a.church_nm LIKE '%" & Me.txtChurch & "%' OR b.church_nm LIKE '%" & Me.txtChurch & "%') ORDER BY a.sort_order;"
        Else
            strSql = "SELECT a.church_sid,a.church_nm,a.church_gb,b.church_nm " & _
                        "FROM " & TB1 & " a LEFT JOIN op_system.db_churchlist_custom b ON a.main_church_cd = b.church_sid " & _
                        "WHERE a.ovs_dept = " & USER_DEPT & " AND a.suspend = 0 AND a.church_gb <> 'MM' AND (a.church_nm LIKE '%" & Me.txtChurch & "%' OR b.church_nm LIKE '%" & Me.txtChurch & "%') ORDER BY a.sort_order;"
        End If
    Case TB2
        strSql = "SELECT * FROM " & TB2 & " a WHERE church_sid = " & SText(Me.lstChurch.List(Me.lstChurch.listIndex)) & ";"
    Case Else
    End Select
    makeSelectSQL = strSql
End Function
Private Function makeUpdateSQL(ByVal tableNM As String) As String
    
    With Me.lstHistory
        strSql = "UPDATE " & TB2 & " a " & _
                "SET a.his_dt = " & IIf(Me.txtDate = "", "NULL", SText(Me.txtDate)) & ",a.history = " & SText(Me.txtHistory) & _
                " WHERE a.his_cd=" & SText(.List(.listIndex)) & ";"
        queryKey = .listIndex
    End With
        
    makeUpdateSQL = strSql
End Function
Private Function makeInsertSQL(ByVal tableNM As String, _
                                argData As T_HISTORY_CHURCH) As String
    With Me.lstHistory
        strSql = "INSERT INTO " & TB2 & " VALUES(DEFAULT," & _
                    SText(argData.church_sid) & "," & _
                    IIf(argData.HIS_DT = "1900-01-01", "NULL", SText(argData.HIS_DT)) & "," & _
                    SText(argData.HISTORY) & ");"
    End With
    queryKey = Me.lstHistory.ListCount - 1
    makeInsertSQL = strSql
End Function
Private Function makeDeleteSQL(ByVal tableNM As String) As String
    With Me.lstHistory
        strSql = "DELETE FROM " & TB2 & " WHERE his_cd = " & SText(.List(.listIndex)) & ";"
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
    
    If Not IsDate(Me.txtDate) Then
        MsgBox "�ùٸ� ��¥ ���°� �ƴմϴ�. �ٽ� Ȯ�� ���ּ���.", vbCritical, banner
        Set txtBox_Focus = Me.txtDate: fnData_Validation = False: Exit Function
    End If
End Function
Sub sbtxtBox_Init()
    Me.txtDate = ""
    Me.txtHistory = ""
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
    Me.lstHistory.Enabled = Not argBoolean
    Me.chkAll.Enabled = Not argBoolean
    
    Me.txtDate.Enabled = argBoolean
    Me.txtHistory.Enabled = argBoolean
    
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


