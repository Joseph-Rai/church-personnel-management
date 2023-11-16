VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Update_Appointment_1 
   Caption         =   "��ȸ�˻�"
   ClientHeight    =   2880
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   5640
   OleObjectBlob   =   "frm_Update_Appointment_1.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "frm_Update_Appointment_1"
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

Private Sub lstChurch_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Un'HookListBoxScroll
End Sub

Private Sub lstChurch_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Me.lstChurch.ListCount Then
        'HookListBoxScroll Me, Me.lstChurch
    End If
End Sub

Private Sub UserForm_Initialize()

    '--//���ʼ���
    Me.cmdClose.Cancel = True
    TB1 = "op_system.db_churchlist_custom" '--//��ȸ����Ʈ
    
    '--//����Ʈ�ڽ� ����
    With Me.lstChurch
        .ColumnCount = 4
        .ColumnHeads = False
        .ColumnWidths = "0,120,0,0" '��ȸ�ڵ�, ��ȸ��, ��ȸ����, ������ȸ��
        .Width = 265.5
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
    Else
        Me.lstChurch.Clear
    End If
    Call sbClearVariant
End Sub
Private Sub txtChurch_Change()
    Me.txtChurch.BackColor = RGB(255, 255, 255)
End Sub
Private Sub lstChurch_Click()
    Me.cmdOk.Enabled = True
End Sub
Private Sub cmdClose_Click()
    Unload frm_Update_Appointment_1
End Sub
Private Sub cmdOK_Click()
    
    '--//��ȸ ���ÿ��� �Ǵ�
    If Me.lstChurch.listIndex = -1 Then
        MsgBox "��ȸ�� �����ϼ���.", vbCritical, "����"
        Exit Sub
    End If
    
    '--//��ȸ���� �Է�
    Select Case argShow3
    Case 1
        Select Case argShow
            Case 1
                With Me.lstChurch
                    frm_Update_Appointment.txtChurchNow = .List(.listIndex, 1)
                    frm_Update_Appointment.txtChurchNow_sid = .List(.listIndex)
                End With
            Case 2
'                With Me.lstChurch
'                    frm_Update_PInformation.txtChurchNow = .list(.listIndex, 1)
'                    frm_Update_PInformation.txtChurchNow_sid = .list(.listIndex)
'                End With
        Case Else
        End Select
    Case 2
        With Me.lstChurch
            frm_Update_FamilyInfo.txtChurch = .List(.listIndex, 1)
            frm_Update_FamilyInfo.txtChurch_Sid = .List(.listIndex)
        End With
    Case 3
        With Me.lstChurch
            frm_Search_Appointment.txtTo = .List(.listIndex, 1)
            frm_Search_Appointment.txtTo_sid = .List(.listIndex)
        End With
    Case Else
    End Select
    
    Unload frm_Update_Appointment_1
    
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
    
    Select Case tableNM
    Case TB1
        '--//��ȸ�ڵ�, ��ȸ��
        Select Case argShow
        Case 1 '--//����ȸ�� �˻�
            If Me.chkAll.Value = False Then
                strSql = "SELECT a.church_sid,a.church_nm,a.church_gb,b.church_nm " & _
                            "FROM " & TB1 & " a LEFT JOIN op_system.db_churchlist_custom b ON a.main_church_cd = b.church_sid " & _
                            "WHERE a.ovs_dept = " & USER_DEPT & " AND (a.church_gb = 'MC' OR a.church_gb = 'HBC') AND (a.church_nm LIKE '%" & Me.txtChurch & "%' OR b.church_nm LIKE '%" & Me.txtChurch & "%') ORDER BY a.sort_order;"
            Else
                strSql = "SELECT a.church_sid,a.church_nm,a.church_gb,b.church_nm " & _
                            "FROM " & TB1 & " a LEFT JOIN op_system.db_churchlist_custom b ON a.main_church_cd = b.church_sid " & _
                            "WHERE (a.church_gb = 'MC' OR a.church_gb = 'HBC') AND (a.church_nm LIKE '%" & Me.txtChurch & "%' OR b.church_nm LIKE '%" & Me.txtChurch & "%') ORDER BY a.sort_order;"
            End If
        Case 2 '--//����ȸ ����ȸ ��� �˻�
            If Me.chkAll.Value = False Then
                strSql = "SELECT a.church_sid,a.church_nm,a.church_gb,b.church_nm " & _
                            "FROM " & TB1 & " a LEFT JOIN op_system.db_churchlist_custom b ON a.main_church_cd = b.church_sid " & _
                            "WHERE a.suspend=0 AND a.ovs_dept = " & USER_DEPT & " AND a.suspend = 0 AND (a.church_nm LIKE '%" & Me.txtChurch & "%' OR b.church_nm LIKE '%" & Me.txtChurch & "%') ORDER BY a.sort_order;"
            Else
                strSql = "SELECT a.church_sid,a.church_nm,a.church_gb,b.church_nm " & _
                            "FROM " & TB1 & " a LEFT JOIN op_system.db_churchlist_custom b ON a.main_church_cd = b.church_sid " & _
                            "WHERE (a.church_nm LIKE '%" & Me.txtChurch & "%' OR b.church_nm LIKE '%" & Me.txtChurch & "%') ORDER BY a.sort_order;"
            End If
        Case 3
            If Me.chkAll.Value = False Then
                strSql = "SELECT a.church_sid,a.church_nm,a.church_gb,b.church_nm " & _
                            "FROM " & TB1 & " a LEFT JOIN op_system.db_churchlist_custom b ON a.main_church_cd = b.church_sid " & _
                            "WHERE a.ovs_dept = " & USER_DEPT & " AND a.suspend = 0 AND (a.church_gb = 'MC' OR a.church_gb = 'HBC') AND (a.church_nm LIKE '%" & Me.txtChurch & "%' OR b.church_nm LIKE '%" & Me.txtChurch & "%') ORDER BY a.sort_order;"
            Else
                strSql = "SELECT a.church_sid,a.church_nm,a.church_gb,b.church_nm " & _
                            "FROM " & TB1 & " a LEFT JOIN op_system.db_churchlist_custom b ON a.main_church_cd = b.church_sid " & _
                            "WHERE (a.church_gb = 'MC' OR a.church_gb = 'HBC') AND a.suspend = 0 AND (a.church_nm LIKE '%" & Me.txtChurch & "%' OR b.church_nm LIKE '%" & Me.txtChurch & "%') ORDER BY a.sort_order;"
            End If
        End Select
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

