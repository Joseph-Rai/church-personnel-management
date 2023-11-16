VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Update_Flight 
   Caption         =   "���Ա��̷� ����������"
   ClientHeight    =   8175
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8130
   OleObjectBlob   =   "frm_Update_Flight.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "frm_Update_Flight"
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
    result.affectedCount = executeSQL("cmdDelete_Click", TB2, strSql, Me.Name, "���Ա� �̷� ����")
    writeLog "cmdDelete_Click", TB2, strSql, 0, Me.Name, "���Ա� �̷� ����"
    disconnectALL
    
    '--//�޼����ڽ�
    MsgBox "�ش� �����Ͱ� �����Ǿ����ϴ�.", , banner
    
    '--//����Ʈ�ڽ� ���ΰ�ħ
    Call lstPStaff_Click
    Me.lstHistory.listIndex = -1
    
End Sub

Private Sub cmdEdit_Click()
    
    Dim result As T_RESULT
    
    '--//������ ���� �ִ��� üũ
    With Me.lstHistory
        If Me.txtDate = .List(.listIndex, 2) And Me.txtDeparture = .List(.listIndex, 3) And Me.txtDestination = .List(.listIndex, 4) And Me.txtPurpose = .List(.listIndex, 5) Then
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
    
'    '--//�ߺ�üũ
'    With Me.lstHistory
'        strSQL = "SELECT * FROM " & TB2 & " a WHERE a.lifeno = " & SText(.List(.ListIndex, 1)) & _
'                " AND a.flight_dt = " & SText(Me.txtDate) & ";"
'        Call makeListData(strSQL, TB2)
'    End With
'
'    If cntRecord > 0 Then
'        If MsgBox("������ ��¥�� �װ��������� �̹� ���� �մϴ�. ���� ���� �Ͻðڽ��ϱ�?", vbQuestion + vbYesNo, banner) = vbNo Then
'            queryKey = listData(0, 0)
'            Call returnListPosition(Me, Me.lstHistory.Name, queryKey)
'            Exit Sub
'        End If
'    End If
    
    Call sbClearVariant
    
    '--//SQL�� ����, ����, �αױ��
    strSql = makeUpdateSQL(TB2)
    result.strSql = strSql
    connectTaskDB
    result.affectedCount = executeSQL("cmdEdit_Click", TB2, strSql, Me.Name, "���Ա� �̷� ������Ʈ")
    writeLog "cmdEdit_Click", TB2, strSql, 0, Me.Name, "���Ա� �̷� ������Ʈ", result.affectedCount
    disconnectALL

    '--//�޼����ڽ�
    MsgBox "������Ʈ �Ǿ����ϴ�.", , banner
    
    '--//����Ʈ�ڽ� �ʱ�ȭ
    Call lstPStaff_Click
    Me.lstHistory.listIndex = queryKey
    
End Sub

Private Sub cmdNew_Click()
    
    '--//�̷� ����Ʈ �ڽ��� �������� �ʾ����� �ؽ�Ʈ �ڽ����� Ȱ��ȭ ���� �����Ƿ� Ŭ��ó�� �ϱ�
    If lstHistory.ListCount = 0 Then
        Call lstHistory_Click
    End If
    
    'Call cmdbtn_visible
    Call INPUTMODE(True)
    Call HideDeleteButtonByUserAuth
    Me.lstHistory.listIndex = Me.lstHistory.ListCount - 1
    Call sbtxtBox_Init
End Sub

Private Sub cmdADD_Click()
    
    Dim argData As T_FLIGHT_SCHEDULE
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
    
    '--//�ߺ�üũ
    With Me.lstPStaff
        strSql = "SELECT * FROM " & TB2 & " a WHERE a.lifeno = " & SText(.List(.listIndex)) & _
                " AND a.flight_dt = " & SText(Me.txtDate) & ";"
        Call makeListData(strSql, TB2)
    End With
    
    If cntRecord > 0 Then
        If MsgBox("������ ��¥�� �װ��������� �̹� ���� �մϴ�. ���� ���� �Ͻðڽ��ϱ�?", vbQuestion + vbYesNo, banner) = vbNo Then
            queryKey = LISTDATA(0, 0)
            Call returnListPosition(Me, Me.lstHistory.Name, queryKey)
            Exit Sub
        End If
    End If
    Call sbClearVariant
    
    '--//�۾��� ���� ����ü�� �� �߰�
    With Me.lstPStaff
        argData.lifeNo = .List(.listIndex)
        argData.FLIGHT_DT = Me.txtDate
        argData.DEPARTURE = Replace(Me.txtDeparture, "�ѱ�", "���ѹα�")
        argData.Destination = Replace(Me.txtDestination, "�ѱ�", "���ѹα�")
        argData.VISIT_PURPOSE = Me.txtPurpose
    End With

    
    '--//�۾��� ���� ������ ���� �� �αױ��
    strSql = makeInsertSQL(TB2, argData)
    connectTaskDB
    result.strSql = strSql
    result.affectedCount = executeSQL("cmdADD_Clikc", TB2, strSql, Me.Name, "���Ա� �̷� �߰�")
    writeLog "cmdADD_Click", TB2, strSql, 0, Me.Name, "���Ա� �̷� �߰�", result.affectedCount
    disconnectALL
    
    '--//�޼����ڽ�
    MsgBox "�߰� �Ǿ����ϴ�.", , banner
    Call lstPStaff_Click
    Me.lstHistory.listIndex = Me.lstHistory.ListCount - 1
    
    '--//��ư���� �������
    'Call cmdbtn_visible
    Call INPUTMODE(False)
    Call HideDeleteButtonByUserAuth
    
End Sub

Private Sub lstHistory_Click()
    
    '--//��Ʈ�� ����
    If Me.lstHistory.listIndex <> -1 Then
        Me.txtDate.Enabled = True
        Me.txtDeparture.Enabled = True
        Me.txtDestination.Enabled = True
        Me.txtPurpose.Enabled = True
        Me.cmdEdit.Enabled = True
        Me.cmdAdd.Enabled = True
        Me.cmdDelete.Enabled = True
    Else
        Me.txtDate.Enabled = False
        Me.txtDeparture.Enabled = False
        Me.txtDestination.Enabled = False
        Me.txtPurpose.Enabled = False
        Me.cmdEdit.Enabled = False
        Me.cmdAdd.Enabled = False
        Me.cmdDelete.Enabled = False
    End If
    
    '--//����Ʈ Ŭ�� �� ������, ������, ���� ǥ��
    If Me.lstHistory.listIndex <> -1 Then
        With Me.lstHistory
            Me.txtDate = .List(.listIndex, 2)
            Me.txtDeparture = .List(.listIndex, 3)
            Me.txtDestination = .List(.listIndex, 4)
            Me.txtPurpose = .List(.listIndex, 5)
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

Private Sub lstPStaff_Click()
    
    Dim filePath As String
    Dim FileName As String
    
    If Me.lstPStaff.listIndex <> -1 Then
        Me.lstHistory.Enabled = True
        Me.cmdNew.Enabled = True
    Else
        Me.lstHistory.Enabled = False
        Me.cmdNew.Enabled = False
    End If
    
    '--//�̷� ��ϻ��� ����
    With Me.lstHistory
        .ColumnCount = 6
        .ColumnHeads = False
        .ColumnWidths = "0,0,60,70,70,180" '���Ա��ڵ�, �����ȣ, ��¥, �����, ������, ���Ա� ����
        .Width = 380
        .TextAlign = fmTextAlignLeft
        .Font = "����"
    End With
    
    '--//���Ա� �̷� ����Ʈ�ڽ� ������
    Call makeSelectSQL(TB2)
    Call makeListData(strSql, TB2)
    If cntRecord > 0 Then
        Me.lstHistory.List = LISTDATA
    Else
        Me.lstHistory.Clear
    End If
    Call sbClearVariant
    
    '--//�ؽ�Ʈ�ڽ� �ʱ�ȭ
    Call sbtxtBox_Init
    Me.txtDate = ""
    
    '--//�����߰�
    Dim strLifeNo As String
    With Me.lstPStaff
        strLifeNo = .List(.listIndex)
    End With
    InsertPicToLabel Me.lblPic, strLifeNo
    
End Sub

Private Sub lstPStaff_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Un'HookListBoxScroll
End Sub

Private Sub lstPStaff_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Me.lstPStaff.ListCount > 0 Then
        'HookListBoxScroll Me, Me.lstPStaff
    End If
End Sub

Private Sub txtChurchNM_Change()
    Me.txtChurchNM.BackColor = RGB(255, 255, 255)
End Sub

Private Sub txtDate_Change()
    Call Date_Format(Me.txtDate)
End Sub

Private Sub txtDeparture_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    argShow = 1
    frm_Search_Country.Show
End Sub

Private Sub txtDeparture_Enter()
    If Me.txtDeparture = "" Then
        argShow = 1
        frm_Search_Country.Show
    End If
End Sub

Private Sub txtDestination_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    argShow = 2
    frm_Search_Country.Show
End Sub

Private Sub txtDestination_Enter()
    If Me.txtDestination = "" Then
        argShow = 2
        frm_Search_Country.Show
    End If
End Sub

Private Sub UserForm_Initialize()
    If connIP = Empty Then setGlobalVariant '//�������� �缳��
    If checkLogin = 0 Then f_login.Show '//�α���üũ
    If checkLogin = 0 Then End '--//�α��� ���� �� ���ν��� ����
    
    '--//���ʼ���
    Me.cmdClose.Cancel = True
    TB1 = "op_system.v0_pstaff_information_all" '--//������ ����Ʈ(��ü)
    TB2 = "op_system.db_flight_schedule" '--//���Ա� �̷� ���̺�
    TB3 = "op_system.v0_pstaff_information" '--//������ ����Ʈ
    
    '--//���ѿ� ���� ��Ʈ�� ����
    Call HideDeleteButtonByUserAuth
    
    '--//��Ʈ�� ����
    Me.lstPStaff.Enabled = False
    Me.lstHistory.Enabled = False
    Me.txtDate.Enabled = False
    Me.txtDeparture.Enabled = False
    Me.txtDestination.Enabled = False
    Me.txtPurpose.Enabled = False
    Me.cmdEdit.Enabled = False
    Me.cmdAdd.Enabled = False
    Me.cmdAdd.Visible = False
    Me.cmdCancel.Visible = False
    Me.cmdNew.Enabled = False
    Me.cmdDelete.Enabled = False
    
    '--//����Ʈ�ڽ� ����
    With Me.lstPStaff
        .ColumnCount = 4
        .ColumnHeads = False
        .ColumnWidths = "0,120,70,50" '�����ȣ, ��ȸ��, �ѱ��̸�(����), ��å
        .TextAlign = fmTextAlignLeft
        .Font = "����"
    End With
    
End Sub
Private Sub cmdSearch_Click()
    
    If Me.chkAll.Value Then
        Me.lstHistory.Clear '--//������ ���Ա� �̷� ����Ʈ�ڽ� �ʱ�ȭ
        
        strSql = makeSelectSQL(TB1)
        Call makeListData(strSql, TB1)
        
        '--//��ȯ�� �����Ͱ� ������ �޼��� �ڽ� �� ����
        If cntRecord = 0 Then
            MsgBox "��ȯ�� DB �����Ͱ� �����ϴ�.", vbInformation, banner
            Call sbClearVariant
            Exit Sub
        End If
        
        '--//������ ����Ʈ�� �˻��� ��� ����
        Me.lstPStaff.List = LISTDATA
        Call sbClearVariant '--//���� �ʱ�ȭ
        Me.lstPStaff.Enabled = True
    Else
        Me.lstHistory.Clear '--//������ ���Ա� �̷� ����Ʈ�ڽ� �ʱ�ȭ
        
        strSql = makeSelectSQL(TB3)
        Call makeListData(strSql, TB3)
        
        '--//��ȯ�� �����Ͱ� ������ �޼��� �ڽ� �� ����
        If cntRecord = 0 Then
            MsgBox "��ȯ�� DB �����Ͱ� �����ϴ�.", vbInformation, banner
            Call sbClearVariant
            Exit Sub
        End If
        
        '--//������ ����Ʈ�� �˻��� ��� ����
        Me.lstPStaff.List = LISTDATA
        Call sbClearVariant '--//���� �ʱ�ȭ
        Me.lstPStaff.Enabled = True
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
        '--//��ȸ�ڵ�, ��ȸ��
        strSql = "SELECT a.`�����ȣ`,a.`��ȸ��`,a.`�ѱ��̸�(����)`,a.`��å` " & _
                    "FROM " & TB1 & " a " & _
                    "WHERE (a.`�ѱ��̸�(����)` LIKE '%" & Me.txtChurchNM & "%' OR a.`��ȸ��` LIKE '%" & Me.txtChurchNM & "%'" & _
                    " OR a.`�����̸�` LIKE '%" & Me.txtChurchNM & "%'" & _
                    " OR a.`�����ȣ` LIKE '%" & Me.txtChurchNM & "%')" & _
                    " AND a.`�����μ�` = " & SText(USER_DEPT) & _
                 " UNION " & _
                 "SELECT b.`����ڻ���`,b.`��ȸ��`,b.`����ѱ��̸�(����)`,b.`�����å` " & _
                    "FROM " & TB1 & " b " & _
                    "WHERE (b.`����ѱ��̸�(����)` LIKE '%" & Me.txtChurchNM & "%' OR b.`��ȸ��` LIKE '%" & Me.txtChurchNM & "%'" & _
                    " OR b.`��𿵹��̸�` LIKE '%" & Me.txtChurchNM & "%'" & _
                    " OR b.`����ڻ���` LIKE '%" & Me.txtChurchNM & "%')" & _
                    " AND b.`�����μ�` = " & SText(USER_DEPT) & ";"
    Case TB2
        strSql = "SELECT * FROM " & TB2 & " a " & _
                "WHERE a.lifeno = " & SText(Me.lstPStaff.List(Me.lstPStaff.listIndex)) & _
                "ORDER BY a.flight_dt;"
    Case TB3
        '--//��ȸ�ڵ�, ��ȸ��
        strSql = "SELECT a.`�����ȣ`,a.`��ȸ��`,a.`�ѱ��̸�(����)`,a.`��å` " & _
                    "FROM " & TB3 & " a " & _
                    "WHERE (a.`�ѱ��̸�(����)` LIKE '%" & Me.txtChurchNM & "%' OR a.`��ȸ��` LIKE '%" & Me.txtChurchNM & "%'" & _
                    " OR a.`�����̸�` LIKE '%" & Me.txtChurchNM & "%'" & _
                    " OR a.`�����ȣ` LIKE '%" & Me.txtChurchNM & "%')" & _
                    " AND a.`�����μ�` = " & SText(USER_DEPT) & _
                 " UNION " & _
                 "SELECT b.`����ڻ���`,b.`��ȸ��`,b.`����ѱ��̸�(����)`,b.`�����å` " & _
                    "FROM " & TB3 & " b " & _
                    "WHERE (b.`����ѱ��̸�(����)` LIKE '%" & Me.txtChurchNM & "%' OR b.`��ȸ��` LIKE '%" & Me.txtChurchNM & "%'" & _
                    " OR b.`��𿵹��̸�` LIKE '%" & Me.txtChurchNM & "%'" & _
                    " OR b.`����ڻ���` LIKE '%" & Me.txtChurchNM & "%')" & _
                    " AND b.`�����μ�` = " & SText(USER_DEPT) & ";"
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
                    "SET a.flight_dt = " & SText(Me.txtDate) & ", a.departure = " & SText(Me.txtDeparture) & ",a.destination = " & SText(Me.txtDestination) & ",a.visit_purpose = " & SText(Me.txtPurpose) & _
                    " WHERE a.flight_cd = " & SText(.List(.listIndex)) & ";"
        queryKey = .listIndex
        End With
    Case Else
    End Select
    makeUpdateSQL = strSql
End Function
Private Function makeInsertSQL(ByVal tableNM As String, argData As T_FLIGHT_SCHEDULE) As String
    
    Select Case tableNM
    Case TB1
    Case TB2
        strSql = "INSERT INTO " & TB2 & " VALUES(DEFAULT," & _
                    SText(argData.lifeNo) & "," & _
                    SText(argData.FLIGHT_DT) & "," & _
                    SText(argData.DEPARTURE) & "," & _
                    SText(argData.Destination) & "," & _
                    SText(argData.VISIT_PURPOSE) & ");"
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
            strSql = "DELETE FROM " & TB2 & " WHERE flight_cd = " & SText(.List(.listIndex)) & ";"
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
    Me.txtDate.Value = Date
    Me.txtDeparture.Value = ""
    Me.txtDestination.Value = ""
    Me.txtPurpose.Value = ""
End Sub
Private Function fnData_Validation()
'---------------------------------------
'������ �Է°��� ���� ������ ��ȿ�� �˻�
'TRUE: �̻����, FALSE: �߸���.
'---------------------------------------
    fnData_Validation = True '�����Ͱ� ��ȿ�ϴٴ� ���� �Ͽ� ����
    
    strSql = "SELECT ctry_nm FROM op_system.db_country"
    Call makeListData(strSql, "op_system.db_country")
    
    If IsInArray(Me.txtDeparture, LISTDATA) = -1 Then
        MsgBox "������� �߸� �Է��Ͽ����ϴ�. �ٽ� Ȯ�� ���ּ���.", vbCritical, banner
        Set txtBox_Focus = Me.txtDeparture: fnData_Validation = False: Exit Function
    End If
    
    If IsInArray(Me.txtDestination, LISTDATA) = -1 Then
        MsgBox "�������� �߸� �Է��Ͽ����ϴ�. �ٽ� Ȯ�� ���ּ���.", vbCritical, banner
        Set txtBox_Focus = Me.txtDestination: fnData_Validation = False: Exit Function
    End If
    
    If Not IsDate(Me.txtDate) Then
        MsgBox "�ùٸ� ��¥ ���°� �ƴմϴ�. ��¥�� �ٽ� Ȯ�� ���ּ���.", vbCritical, banner
        Set txtBox_Focus = Me.txtDate: fnData_Validation = False: Exit Function
    End If
    
    If Me.txtPurpose = "" Then
        MsgBox "���Ա� ������ ����� �ּ���.", vbCritical, banner
        Set txtBox_Focus = Me.txtPurpose: fnData_Validation = False: Exit Function
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
    Me.txtChurchNM.Enabled = Not argBoolean
    Me.cmdSearch.Enabled = Not argBoolean
    Me.lstPStaff.Enabled = Not argBoolean
    Me.lstHistory.Enabled = Not argBoolean
    Me.chkAll.Enabled = Not argBoolean
    
    Me.txtDate.Enabled = argBoolean
    Me.txtDeparture.Enabled = argBoolean
    Me.txtDestination.Enabled = argBoolean
    Me.txtPurpose.Enabled = argBoolean
    
    
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

