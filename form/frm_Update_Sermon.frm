VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Update_Sermon 
   Caption         =   "��ǥ�� ����������"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7140
   OleObjectBlob   =   "frm_Update_Sermon.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "frm_Update_Sermon"
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

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdEdit_Click()
    
    Dim result As T_RESULT
    Dim argData As T_SERMON
    
    '--//������ ���� �ִ��� üũ
    With Me.lstPStaff
        If Me.txtScore_Avg = .List(.listIndex, 4) And Me.txtSubject_Count = .List(.listIndex, 5) Then
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
    With Me.lstPStaff
        strSql = "SELECT * FROM " & TB2 & " a WHERE a.lifeno = " & SText(.List(.listIndex)) & ";"
        Call makeListData(strSql, TB2)
    
        If cntRecord > 0 Then
            strSql = makeUpdateSQL(TB2)
        Else
            argData.lifeNo = .List(.listIndex)
            argData.SCORE_AVG = Me.txtScore_Avg
            argData.SUBJECT_COUNT = Me.txtSubject_Count
            strSql = makeInsertSQL(TB2, argData)
        End If
    End With
    
    result.strSql = strSql
    connectTaskDB
    result.affectedCount = executeSQL("cmdEdit_Click", TB2, strSql, Me.Name, "��ǥ���� ������Ʈ")
    writeLog "cmdEdit_Click", TB2, strSql, 0, Me.Name, "��ǥ���� ������Ʈ", result.affectedCount
    disconnectALL
    
    Call sbClearVariant
    
    '--//�޼����ڽ�
    MsgBox "������Ʈ �Ǿ����ϴ�.", , banner
    
    '--//����Ʈ�ڽ� �ʱ�ȭ
    Call cmdSearch_Click
    Call lstPStaff_Click
'    Me.lstPStaff.ListIndex = queryKey
    
End Sub

Private Sub lstPStaff_Click()
    
    Dim filePath As String
    Dim FileName As String
    
    Me.cmdEdit.Enabled = True
    Me.txtScore_Avg.Enabled = True
    Me.txtSubject_Count.Enabled = True
    
    '--//�ؽ�Ʈ�ڽ� �ʱ�ȭ
    Me.txtScore_Avg = ""
    Me.txtSubject_Count = ""
    
    '--//�ؽ�Ʈ�ڽ� �����߰�
    With Me.lstPStaff
        Me.txtScore_Avg = .List(.listIndex, 4)
        Me.txtSubject_Count = .List(.listIndex, 5)
    End With
    
    '--//�����߰�
    filePath = fnFindPicPath
    FileName = Me.lstPStaff.List(Me.lstPStaff.listIndex) & ".jpg"
    
'    If Not Len(Dir(FilePath & FileName)) > 0 Then
'        FileName = Me.lstPStaff.List(Me.lstPStaff.ListIndex) & ".png"
'    End If
    
On Error Resume Next
    Me.lblPic.Picture = LoadPicture(filePath & FileName)
    If err.Number <> 0 Then
        Me.lblPic.Picture = LoadPicture("")
    End If
On Error GoTo 0
    
End Sub

Private Sub lstPStaff_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Un'HookListBoxScroll
End Sub

Private Sub lstPStaff_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    'HookListBoxScroll Me, Me.lstPStaff
End Sub

Private Sub txtChurchNM_Change()
    Me.txtChurchNM.BackColor = RGB(255, 255, 255)
End Sub

Private Sub txtScore_Avg_Change()
    If InStr(Me.txtScore_Avg, ".") > 0 Then
        Me.txtScore_Avg = Left(Me.txtScore_Avg, InStr(Me.txtScore_Avg, ".") + 2)
    Else
        Me.txtScore_Avg = Left(Me.txtScore_Avg, 2)
    End If
End Sub

Private Sub UserForm_Initialize()
    If connIP = Empty Then setGlobalVariant '//�������� �缳��
    If checkLogin = 0 Then f_login.Show '//�α���üũ
    If checkLogin = 0 Then End '--//�α��� ���� �� ���ν��� ����
    
    '--//���ʼ���
    Me.cmdClose.Cancel = True
    TB1 = "op_system.v0_pstaff_information_all" '--//����������(��ü�˻�)
    TB2 = "op_system.db_sermon" '--//��ǥ����
    TB3 = "op_system.v0_pstaff_information" '--//����������
    
    '--//��Ʈ�� ����
    Me.txtScore_Avg.Enabled = False
    Me.txtSubject_Count.Enabled = False
    Me.cmdEdit.Enabled = False
    
    '--//����Ʈ�ڽ� ����
    With Me.lstPStaff
        .ColumnCount = 4
        .ColumnHeads = False
        .ColumnWidths = "0,120,70,50,0,0" '�����ȣ, ��ȸ��, �ѱ��̸�(����), ��å, �������, ��ǥ����
        .TextAlign = fmTextAlignLeft
        .Font = "����"
    End With
    
    Me.txtChurchNM.SetFocus
    
End Sub
Private Sub cmdSearch_Click()
    
    If Me.chkAll.Value Then
        strSql = makeSelectSQL(TB1)
        Call makeListData(strSql, TB1)
        
        If cntRecord = 0 Then
            MsgBox "��ȯ�� DB �����Ͱ� �����ϴ�.", vbInformation, banner
            Call sbClearVariant
            Exit Sub
        End If
        
        Me.lstPStaff.List = LISTDATA
        Call sbClearVariant
        Me.lstPStaff.Enabled = True
    Else
        strSql = makeSelectSQL(TB3)
        Call makeListData(strSql, TB3)
        
        If cntRecord = 0 Then
            MsgBox "��ȯ�� DB �����Ͱ� �����ϴ�.", vbInformation, banner
            Call sbClearVariant
            Exit Sub
        End If
        
        Me.lstPStaff.List = LISTDATA
        Call sbClearVariant
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
        strSql = "SELECT a.`�����ȣ`,a.`��ȸ��`,a.`�ѱ��̸�(����)`,a.`��å`,b.score_avg,b.subject_count " & _
                    "FROM " & TB1 & " a " & _
                    "LEFT JOIN op_system.db_sermon b on a.`�����ȣ` = b.lifeno " & _
                    "WHERE (a.`�ѱ��̸�(����)` LIKE '%" & Me.txtChurchNM & "%' OR a.`��ȸ��` LIKE '%" & Me.txtChurchNM & "%'" & _
                    " OR a.`�����̸�` LIKE '%" & Me.txtChurchNM & "%' OR a.`����ȸ��` LIKE '%" & Me.txtChurchNM & "%'" & _
                    " OR a.`������ȸ��` LIKE '%" & Me.txtChurchNM & "%' OR a.`��������ȸ��` LIKE '%" & Me.txtChurchNM & "%'" & _
                    " OR a.`�����ȣ` LIKE '%" & Me.txtChurchNM & "%')" & _
                    " AND a.`�����μ�` = " & SText(USER_DEPT) & ";"
    Case TB2
    Case TB3
        '--//��ȸ�ڵ�, ��ȸ��
        strSql = "SELECT a.`�����ȣ`,a.`��ȸ��`,a.`�ѱ��̸�(����)`,a.`��å`,b.score_avg,b.subject_count " & _
                    "FROM " & TB3 & " a " & _
                    "LEFT JOIN op_system.db_sermon b on a.`�����ȣ` = b.lifeno " & _
                    "WHERE (a.`�ѱ��̸�(����)` LIKE '%" & Me.txtChurchNM & "%' OR a.`��ȸ��` LIKE '%" & Me.txtChurchNM & "%'" & _
                    " OR a.`�����̸�` LIKE '%" & Me.txtChurchNM & "%' OR a.`����ȸ��` LIKE '%" & Me.txtChurchNM & "%'" & _
                    " OR a.`������ȸ��` LIKE '%" & Me.txtChurchNM & "%' OR a.`��������ȸ��` LIKE '%" & Me.txtChurchNM & "%'" & _
                    " OR a.`�����ȣ` LIKE '%" & Me.txtChurchNM & "%')" & _
                    " AND a.`�����μ�` = " & SText(USER_DEPT) & ";"
    Case Else
    End Select
    makeSelectSQL = strSql
End Function

Private Function makeUpdateSQL(ByVal tableNM As String) As String
    
    Select Case tableNM
    Case TB1
    Case TB2
        With Me.lstPStaff
            strSql = "UPDATE " & TB2 & " a " & _
                    "SET a.score_avg = " & SText(Me.txtScore_Avg) & ", a.subject_count = " & SText(Me.txtSubject_Count) & _
                    " WHERE a.lifeno = " & SText(.List(.listIndex)) & ";"
        queryKey = .listIndex
        End With
    Case Else
    End Select
    makeUpdateSQL = strSql
End Function

Private Function makeInsertSQL(ByVal tableNM As String, argData As T_SERMON) As String
    
    Select Case tableNM
    Case TB1
    Case TB2
        strSql = "INSERT INTO " & TB2 & " VALUES(" & _
                    SText(argData.lifeNo) & "," & _
                    SText(argData.SCORE_AVG) & "," & _
                    SText(argData.SUBJECT_COUNT) & ");"
    Case Else
    End Select
    makeInsertSQL = strSql
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
    
    If Not IsNumeric(Me.txtScore_Avg) And Me.txtScore_Avg <> "" Then
        MsgBox "��ǥ������ �߸� �Է��Ͽ����ϴ�. �ٽ� Ȯ�� ���ּ���.", vbCritical, banner
        Set txtBox_Focus = Me.txtScore_Avg: fnData_Validation = False: Exit Function
    End If
    
    If Me.txtScore_Avg = "" Then
        MsgBox "��ǥ������ �ʼ� �Է°� �Դϴ�. �ٽ� Ȯ�� ���ּ���.", vbCritical, banner
        Set txtBox_Focus = Me.txtScore_Avg: fnData_Validation = False: Exit Function
    End If
    
    If Not IsNumeric(Me.txtSubject_Count) And Me.txtSubject_Count <> "" Then
        MsgBox "��ǥ������ �߸� �Է��Ͽ����ϴ�. �ٽ� Ȯ�� ���ּ���.", vbCritical, banner
        Set txtBox_Focus = Me.txtSubject_Count: fnData_Validation = False: Exit Function
    End If
    
End Function


