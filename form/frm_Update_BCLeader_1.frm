VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Update_BCLeader_1 
   Caption         =   "������ �˻�"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6090
   OleObjectBlob   =   "frm_Update_BCLeader_1.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "frm_Update_BCLeader_1"
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

Private Sub lstPStaff_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Un'HookListBoxScroll
End Sub

Private Sub lstPStaff_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Me.lstPStaff.ListCount Then
        'HookListBoxScroll Me, Me.lstPStaff
    End If
End Sub

Private Sub UserForm_Initialize()

    '--//���ʼ���
    Me.cmdClose.Cancel = True
    TB1 = "op_system.v0_pstaff_information_all" '--//�����ڸ���Ʈ(��ü)
    TB2 = "op_system.v0_pstaff_information" '--//�����ڸ���Ʈ
    
    '--//����Ʈ�ڽ� ����
    With Me.lstPStaff
        .ColumnCount = 4
        .ColumnHeads = False
        .ColumnWidths = "0,120,70,50" '�����ȣ, ��ȸ��, �ѱ��̸�(����), ��å
        .TextAlign = fmTextAlignLeft
        .Font = "����"
    End With
    
    Me.txtChurchNM.SetFocus
End Sub
Private Sub cmdSearch_Click()
    If Me.chkAll.Value Then
        strSql = makeSelectSQL(TB1)
        Call makeListData(strSql, TB1)
        If cntRecord > 0 Then
            Me.lstPStaff.List = LISTDATA
        Else
            Me.lstPStaff.Clear
        End If
        Call sbClearVariant
    Else
        strSql = makeSelectSQL(TB2)
        Call makeListData(strSql, TB2)
        If cntRecord > 0 Then
            Me.lstPStaff.List = LISTDATA
        Else
            Me.lstPStaff.Clear
        End If
        Call sbClearVariant
    End If
End Sub
Private Sub txtChurch_Change()
    Me.txtChurchNM.BackColor = RGB(255, 255, 255)
End Sub
Private Sub lstPStaff_Click()
    Me.cmdOk.Enabled = True
End Sub
Private Sub cmdClose_Click()
    Unload frm_Update_BCLeader_1
End Sub
Private Sub cmdOK_Click()
    
    '--//��ȸ ���ÿ��� �Ǵ�
    If Me.lstPStaff.listIndex = -1 Then
        MsgBox "��ȸ�� �����ϼ���.", vbCritical, banner
        Exit Sub
    End If
    
    Select Case argShow
    Case 1
        '--//��ȸ���� �Է�
        With Me.lstPStaff
            frm_Update_BCLeader.txtManager = .List(.listIndex, 2)
            frm_Update_BCLeader.txtLifeNo = .List(.listIndex)
        End With
    Case 2
        '--//��ȸ���� �Է�
        With Me.lstPStaff
            frm_Update_FamilyInfo.txtLifeNo = .List(.listIndex)
            strSql = "SELECT * FROM op_system.v0_pstaff_information_all a WHERE a.`�����ȣ` = " & SText(.List(.listIndex)) & " OR a.`����ڻ���` = " & SText(.List(.listIndex)) & ";"
            Call makeListData(strSql, "op_system.v_pstaff_detail")
            frm_Update_FamilyInfo.txtName_ko = IIf(Mid(.List(.listIndex), 12, 1) = 1, LISTDATA(0, 22), LISTDATA(0, 23))
            frm_Update_FamilyInfo.txtName_en = IIf(Mid(.List(.listIndex), 12, 1) = 1, LISTDATA(0, 7), LISTDATA(0, 16))
            frm_Update_FamilyInfo.txtChurch = LISTDATA(0, 0)
            frm_Update_FamilyInfo.cboTitle = IIf(Mid(.List(.listIndex), 12, 1) = 1, LISTDATA(0, 26), LISTDATA(0, 27))
            frm_Update_FamilyInfo.cboPosition = IIf(Mid(.List(.listIndex), 12, 1) = 1, LISTDATA(0, 8), LISTDATA(0, 17))
            frm_Update_FamilyInfo.txtEducation = IIf(Mid(.List(.listIndex), 12, 1) = 1, LISTDATA(0, 24), LISTDATA(0, 25))
            frm_Update_FamilyInfo.txtBirthday = IIf(Mid(.List(.listIndex), 12, 1) = 1, LISTDATA(0, 10), LISTDATA(0, 18))
            frm_Update_FamilyInfo.cboReligion = "��������"
        End With
        
        
        
    Case Else
    End Select
    
    Unload frm_Update_BCLeader_1
    
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
        MsgBox "��ȯ�� DB �����Ͱ� �����ϴ�.", vbInformation, banner
        Exit Sub
    End If
    
End Sub
'-----------------------------------------
'  ���Ǻ� Select SQL�� �ۼ�
'    - makeSelectSQL(�˻���, ����)
'    - DB���� ��ȯ�� ����Ʈ �ʵ����
'-----------------------------------------
Private Function makeSelectSQL(ByVal tableNM As String) As String
    
    Select Case argShow
    Case 1
        Select Case tableNM
        Case TB1
            '�����ȣ, ��ȸ��, �ѱ��̸�(����), ��å
            strSql = "SELECT a.`�����ȣ`,a.`��ȸ��`,a.`�ѱ��̸�(����)`,a.`��å` " & _
                        "FROM " & TB1 & " a " & _
                        "WHERE (a.`�ѱ��̸�(����)` LIKE '%" & Me.txtChurchNM & "%' OR a.`��ȸ��` LIKE '%" & Me.txtChurchNM & "%'" & _
                        " OR a.`�����̸�` LIKE '%" & Me.txtChurchNM & "%' OR a.`����ȸ��` LIKE '%" & Me.txtChurchNM & "%'" & _
                        " OR a.`������ȸ��` LIKE '%" & Me.txtChurchNM & "%' OR a.`��������ȸ��` LIKE '%" & Me.txtChurchNM & "%'" & _
                        " OR a.`�����ȣ` LIKE '%" & Me.txtChurchNM & "%')" & _
                        " AND a.`�����μ�` = " & SText(USER_DEPT) & ";"
        Case TB2
            '�����ȣ, ��ȸ��, �ѱ��̸�(����), ��å
            strSql = "SELECT a.`�����ȣ`,a.`��ȸ��`,a.`�ѱ��̸�(����)`,a.`��å` " & _
                        "FROM " & TB2 & " a " & _
                        "WHERE (a.`�ѱ��̸�(����)` LIKE '%" & Me.txtChurchNM & "%' OR a.`��ȸ��` LIKE '%" & Me.txtChurchNM & "%'" & _
                        " OR a.`�����̸�` LIKE '%" & Me.txtChurchNM & "%' OR a.`����ȸ��` LIKE '%" & Me.txtChurchNM & "%'" & _
                        " OR a.`������ȸ��` LIKE '%" & Me.txtChurchNM & "%' OR a.`��������ȸ��` LIKE '%" & Me.txtChurchNM & "%'" & _
                        " OR a.`�����ȣ` LIKE '%" & Me.txtChurchNM & "%')" & _
                        " AND a.`�����μ�` = " & SText(USER_DEPT) & ";"
        Case Else
        End Select
    Case 2
        '�����ȣ, ��ȸ��, �ѱ��̸�(����), ��å
        strSql = "SELECT a.`�����ȣ`,a.`��ȸ��`,a.`�ѱ��̸�(����)`,a.`��å` " & _
                    "FROM " & TB1 & " a " & _
                    "WHERE (a.`�ѱ��̸�(����)` LIKE '%" & Me.txtChurchNM & "%' OR a.`��ȸ��` LIKE '%" & Me.txtChurchNM & "%'" & _
                    " OR a.`�����̸�` LIKE '%" & Me.txtChurchNM & "%' OR a.`����ȸ��` LIKE '%" & Me.txtChurchNM & "%'" & _
                    " OR a.`������ȸ��` LIKE '%" & Me.txtChurchNM & "%' OR a.`��������ȸ��` LIKE '%" & Me.txtChurchNM & "%'" & _
                    " OR a.`�����ȣ` LIKE '%" & Me.txtChurchNM & "%')" & _
                    " AND a.`�����μ�` = " & SText(USER_DEPT) & _
                    " UNION " & _
                    "SELECT a.`����ڻ���`,a.`��ȸ��`,a.`����ѱ��̸�(����)`,a.`�����å` " & _
                    "FROM " & TB1 & " a " & _
                    "WHERE (a.`����ѱ��̸�(����)` LIKE '%" & Me.txtChurchNM & "%' OR a.`��ȸ��` LIKE '%" & Me.txtChurchNM & "%'" & _
                    " OR a.`��𿵹��̸�` LIKE '%" & Me.txtChurchNM & "%' OR a.`����ȸ��` LIKE '%" & Me.txtChurchNM & "%'" & _
                    " OR a.`������ȸ��` LIKE '%" & Me.txtChurchNM & "%' OR a.`��������ȸ��` LIKE '%" & Me.txtChurchNM & "%'" & _
                    " OR a.`����ڻ���` LIKE '%" & Me.txtChurchNM & "%')" & _
                    " AND a.`�����μ�` = " & SText(USER_DEPT) & ";"
    Case Else
    End Select
    makeSelectSQL = strSql
End Function

Private Sub sbClearVariant()
    Erase LISTFIELD
    Erase LISTDATA
    cntRecord = Empty
    strSql = vbNullString
End Sub



