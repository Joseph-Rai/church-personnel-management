VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Search_Country 
   Caption         =   "���� �˻� ������"
   ClientHeight    =   3000
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   5535
   OleObjectBlob   =   "frm_Search_Country.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "frm_Search_Country"
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
Dim ws As Worksheet

Private Sub lstCountry_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Un'HookListBoxScroll
End Sub

Private Sub lstCountry_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Me.lstCountry.ListCount Then
        'HookListBoxScroll Me, Me.lstCountry
    End If
End Sub

Private Sub UserForm_Initialize()
    
    '--//���ʼ���
    Me.cmdClose.Cancel = True
    TB1 = "op_system.db_country" '--//��������Ʈ
    
    '--//����Ʈ�ڽ� ����
    With Me.lstCountry
        .ColumnCount = 4
        .ColumnHeads = False
        .ColumnWidths = "270" '������
        .Width = 265.5
        .TextAlign = fmTextAlignLeft
        .Font = "����"
    End With
    
    Me.txtCountry.SetFocus
    
End Sub
Private Sub cmdSearch_Click()
    strSql = makeSelectSQL(TB1)
    Call makeListData(strSql, TB1)
    If cntRecord > 0 Then
        Me.lstCountry.List = LISTDATA
    End If
    Call sbClearVariant
End Sub
Private Sub txtCountry_Change()
    Me.txtCountry.BackColor = RGB(255, 255, 255)
End Sub
Private Sub lstCountry_Click()
    Me.cmdOk.Enabled = True
End Sub
Private Sub cmdClose_Click()
    Unload Me
End Sub
Private Sub cmdOK_Click()
    
    '--//��ȸ ���ÿ��� �Ǵ�
    If Me.lstCountry.listIndex = -1 Then
        MsgBox "������ �����ϼ���.", vbCritical, banner
        Exit Sub
    End If
    
    '--//��ȸ���� �Է�
    Select Case argShow
        Case 1
            With Me.lstCountry
                frm_Update_Flight.txtDeparture = .List(.listIndex)
            End With
        Case 2
            With Me.lstCountry
                frm_Update_Flight.txtDestination = .List(.listIndex)
            End With
        Case 3
            With Me.lstCountry
                frm_Update_PInformation.txtNationality = .List(.listIndex)
            End With
        Case 4
            With Me.lstCountry
                frm_Update_PInformation.txtNationality_Spouse = .List(.listIndex)
            End With
    Case Else
    End Select
    
    Unload Me
    
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
        strSql = "SELECT * FROM " & TB1 & " a WHERE a.ctry_nm LIKE '%" & Replace(Me.txtCountry, "�ѱ�", "���ѹα�") & "%';"
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



