VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Update_Visa 
   Caption         =   "�������� ����������"
   ClientHeight    =   8295.001
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14625
   OleObjectBlob   =   "frm_Update_Visa.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "frm_Update_Visa"
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

Private Sub chkPresent_Change()
    Select Case Me.chkPresent.Value
        Case True
            Me.txtEnd.BackColor = &HE0E0E0
            Me.txtEnd.Value = "����"
            Me.txtEnd.Enabled = False
        Case False
            Me.txtEnd.Enabled = True
            Me.txtEnd.BackColor = RGB(255, 255, 255)
            If Me.lstHistory.listIndex = -1 Then
                Me.txtEnd = ""
            Else
                If Me.txtEnd = "����" Then
                    Me.txtEnd.Value = Date - 1
                End If
            End If
    Case Else
    End Select
End Sub

Private Sub cmdCancel_Click()
    Call INPUTMODE(False)
    Call HideDeleteButtonByUserAuth
    Call sbtxtBox_Init
    Me.txtEnd = ""
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
    result.affectedCount = executeSQL("cmdDelete_Click", TB2, strSql, Me.Name, "�����̷� ����")
    writeLog "cmdDelete_Click", TB2, strSql, 0, Me.Name, "�����̷� ����"
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
        If Me.cboVisa = .List(.listIndex, 4) And Me.txtStart = .List(.listIndex, 2) And Me.txtEnd = .List(.listIndex, 3) And Me.txtMemo = .List(.listIndex, 5) Then
            Exit Sub
        End If
    End With
    
    '--//�ߺ�üũ
    With Me.lstHistory
        strSql = "SELECT * FROM " & TB2 & " a WHERE a.lifeno = " & SText(.List(.listIndex, 1)) & _
                " AND ((a.start_dt BETWEEN " & SText(Me.txtStart) & " AND " & SText(IIf(Me.txtEnd = "����", DateSerial(9999, 12, 31), Me.txtEnd)) & _
                ") OR (a.end_dt BETWEEN " & SText(Me.txtStart) & " AND " & SText(IIf(Me.txtEnd = "����", DateSerial(9999, 12, 31), Me.txtEnd)) & _
                ") OR (a.start_dt <= " & SText(Me.txtStart) & " AND a.end_dt >= " & SText(IIf(Me.txtEnd = "����", DateSerial(9999, 12, 31), Me.txtEnd)) & _
                ")) AND a.visa_cd <> " & SText(.List(.listIndex)) & ";"
        Call makeListData(strSql, TB2)
    End With
    
    If cntRecord > 0 Then
        MsgBox "�ߺ��� �Ⱓ�� ������ �� �����ϴ�. �Է� ���� �ٽ� Ȯ�����ּ���.", vbCritical, banner
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
    
    '--//SQL�� ����, ����, �αױ��

    strSql = makeUpdateSQL(TB2)
    result.strSql = strSql
    connectTaskDB
    result.affectedCount = executeSQL("cmdEdit_Click", TB2, strSql, Me.Name, "�����̷� ������Ʈ")
    writeLog "cmdEdit_Click", TB2, strSql, 0, Me.Name, "�����̷� ������Ʈ", result.affectedCount
    disconnectALL
    
    '--//�޼����ڽ�
    MsgBox "������Ʈ �Ǿ����ϴ�.", , banner
    
    '--//����Ʈ�ڽ� �ʱ�ȭ
    Call lstPStaff_Click
    Call lstHistory_Click
    Me.lstHistory.listIndex = queryKey
    
End Sub

Private Sub cmdNew_Click()
    Call lstHistory_Click
    Call INPUTMODE(True)
    Call HideDeleteButtonByUserAuth
    Me.lstHistory.listIndex = Me.lstHistory.ListCount - 1
    Call sbtxtBox_Init
    Me.chkPresent.Value = True
End Sub

Private Sub cmdADD_Click()
    
    Dim argData As T_VISA
    Dim result As T_RESULT
    
    '--//�ߺ�üũ
    
    strSql = "SELECT * FROM " & TB2 & " a WHERE a.lifeno = " & SText(Me.lstPStaff.List(Me.lstPStaff.listIndex)) & _
            " AND ((a.start_dt BETWEEN " & SText(Me.txtStart) & " AND " & SText(IIf(Me.txtEnd = "����", DateSerial(9999, 12, 31), Me.txtEnd)) & _
            ") OR (a.end_dt BETWEEN " & SText(Me.txtStart) & " AND " & SText(IIf(Me.txtEnd = "����", DateSerial(9999, 12, 31), Me.txtEnd)) & _
            ") OR (a.start_dt <= " & SText(Me.txtStart) & " AND a.end_dt >= " & SText(IIf(Me.txtEnd = "����", DateSerial(9999, 12, 31), Me.txtEnd)) & "));"
    Call makeListData(strSql, TB2)
   
    
    If cntRecord > 0 Then
        MsgBox "�ߺ��� �Ⱓ�� ������ �� �����ϴ�. �Է� ���� �ٽ� Ȯ�����ּ���.", vbCritical, banner
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
    With Me.lstHistory
        argData.lifeNo = Me.lstPStaff.List(Me.lstPStaff.listIndex)
        argData.START_DT = Me.txtStart
        argData.END_DT = IIf(Me.txtEnd = "����", DateSerial(9999, 12, 31), Me.txtEnd)
        argData.Visa = Me.cboVisa
        argData.memo = Me.txtMemo
    End With
    
    '--//�۾��� ���� ������ ���� �� �αױ��
    strSql = makeInsertSQL(TB2, argData)
    connectTaskDB
    result.strSql = strSql
    result.affectedCount = executeSQL("cmdADD_Clikc", TB2, strSql, Me.Name, "�����̷� �߰�")
    writeLog "cmdADD_Click", TB2, strSql, 0, Me.Name, "�����̷� �߰�", result.affectedCount
    disconnectALL
    
    '--//�޼����ڽ�
    MsgBox "�߰� �Ǿ����ϴ�.", , banner
    Call lstPStaff_Click
    Me.lstHistory.listIndex = Me.lstHistory.ListCount - 1
    
    '--//��ư���� �������
    Call INPUTMODE(False)
    Call HideDeleteButtonByUserAuth
    
End Sub

Private Sub cmdSavePassportPhoto_Click()
    Dim filePath As String
    Dim lifeNo As String
    With Me.lstPStaff
        lifeNo = .List(.listIndex)
    End With
    
    filePath = Application.GetOpenFilename("Image File Format, *.jpg;*.jpeg;", , "Photo Select")
    
    If filePath = "False" Then
        Exit Sub
    End If
    
    '--//DB�� ���� ����
    savePassportPhoto lifeNo, filePath
    
    '--//�����޽���
    MsgBox "������ ���������� ����Ǿ����ϴ�.", , banner
    
    '--//�󺧿� ��������
    InsertPicToLabel Me.lblPassportPhoto, lifeNo, PASSPORT_PHOTO
End Sub

Private Sub savePassportPhoto(lifeNo As String, filePath As String)

    Dim ppPhoto As New PassportPhoto
    Dim ppPhotoDao As New PassportPhotoDao
    Dim stream As New ADODB.stream
    
    stream.Type = adTypeBinary
    stream.Open
    stream.LoadFromFile filePath
    
    ppPhoto.lifeNo = lifeNo
    ppPhoto.photo = encodeBase64(stream.Read)
    ppPhotoDao.Save ppPhoto

End Sub

Private Sub cmdSaveVisaPhoto_Click()
    Dim filePath As String
    Dim visa_cd As String
    With Me.lstHistory
        If .ListCount > 0 Then
            visa_cd = .List(.listIndex)
        Else
            Exit Sub
        End If
    End With
    
    filePath = Application.GetOpenFilename("Image File Format, *.jpg;*.jpeg;", , "Photo Select")
    
    If filePath = "False" Then
        Exit Sub
    End If
    
    '--//DB�� ���� ����
    saveVisaPhoto visa_cd, filePath
    
    '--//�����޽���
    MsgBox "������ ���������� ����Ǿ����ϴ�.", , banner
    
    '--//�󺧿� ��������
    InsertPicToLabel Me.lblVisaPhoto, visa_cd, VISA_PHOTO
End Sub

Private Sub saveVisaPhoto(visaCode As String, filePath As String)

    Dim vPhoto As New VisaPhoto
    Dim vPhotoDao As New VisaPhotoDao
    Dim stream As New ADODB.stream
    
    stream.Type = adTypeBinary
    stream.Open
    stream.LoadFromFile filePath
    
    vPhoto.visaCode = visaCode
    vPhoto.photo = encodeBase64(stream.Read)
    vPhotoDao.Save vPhoto

End Sub

Private Sub lblVisaPhoto_Click()

End Sub

Private Sub lstHistory_Click()
    
    '--//��Ʈ�� ����
    If Me.lstHistory.listIndex <> -1 Then
        Me.txtStart.Enabled = True
        Me.txtEnd.Enabled = True
        Me.cboVisa.Enabled = True
        Me.txtMemo.Enabled = True
        Me.cmdEdit.Enabled = True
        Me.cmdAdd.Enabled = True
        Me.cmdDelete.Enabled = True
        Me.chkPresent.Visible = True
    Else
        Me.txtStart.Enabled = False
        Me.txtEnd.Enabled = False
        Me.cboVisa.Enabled = False
        Me.txtMemo.Enabled = False
        Me.cmdEdit.Enabled = False
        Me.cmdAdd.Enabled = False
        Me.cmdDelete.Enabled = False
        Me.chkPresent.Visible = False
    End If
    
    '--//����Ʈ Ŭ�� �� ������, ������, ���� ǥ��
    Dim visa_cd As String
    With Me.lstHistory
        If .ListCount > 0 And .listIndex <> -1 Then
            Me.cboVisa = .List(.listIndex, 4)
            Me.txtStart = .List(.listIndex, 2)
            Me.txtEnd = IIf(.List(.listIndex, 3) = "9999-12-31", "����", .List(.listIndex, 3))
            Me.txtMemo = .List(.listIndex, 5)
            
            visa_cd = .List(.listIndex)
            
            Me.cmdSavePassportPhoto.Enabled = True
        Else
            Me.cmdSavePassportPhoto.Enabled = False
        End If
    End With
    
    '--//üũ�ڽ� �� ����
    If Me.txtEnd = "����" Then
        Me.chkPresent.Value = True
    Else
        Me.chkPresent.Value = False
    End If
    
    '--//�������� �߰�
    InsertPicToLabel Me.lblVisaPhoto, visa_cd, VISA_PHOTO
    If Not Me.lblVisaPhoto.Picture = 0 Then
        Me.cmdSaveVisaPhoto.Caption = "��������"
    Else
        Me.cmdSaveVisaPhoto.Caption = "�������"
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
    
    Call UserForm_Initialize
    
    '--//��Ʈ�� ����
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
        .ColumnWidths = "0,0,65,65,70,120" '�����̷��ڵ�, �����ȣ, ������, ������, ��������, �޸�
        .Width = 337
        .TextAlign = fmTextAlignLeft
        .Font = "����"
    End With
    
    '--//�̷¸�� ������ ä���
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
    Me.txtEnd.Value = ""
    Me.chkPresent.Value = False
    Me.chkPresent.Visible = False
    
    '--//������ Ŭ���ϸ� �������� ��ư Ȱ��ȭ
    With Me.lstPStaff
        If .listIndex > -1 Then
            Me.cmdSavePassportPhoto.Enabled = True
        Else
            Me.cmdSavePassportPhoto.Enabled = False
        End If
    End With
    
    '--//�̷� ����Ʈ�ڽ��� ������� ������ ������ ������ Ŭ��
    If Me.lstHistory.ListCount > 0 Then
        Me.lstHistory.listIndex = Me.lstHistory.ListCount - 1
        
        Me.cmdSaveVisaPhoto.Enabled = True
    Else
        Me.cmdSaveVisaPhoto.Enabled = False
    End If
    
    '--//�����߰�
    Dim strLifeNo As String
    With Me.lstPStaff
        strLifeNo = .List(.listIndex)
    End With
    InsertPicToLabel Me.lblPic, strLifeNo
    
    '--//���ǻ��� �߰�
    InsertPicToLabel Me.lblPassportPhoto, strLifeNo, PASSPORT_PHOTO
    If Not Me.lblPassportPhoto.Picture = 0 Then
        Me.cmdSavePassportPhoto.Caption = "���Ǽ���"
    Else
        Me.cmdSavePassportPhoto.Caption = "���ǵ��"
    End If
     
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

Private Sub txtEnd_Change()
    Call Date_Format(Me.txtEnd)
End Sub

Private Sub txtStart_Change()
    Call Date_Format(Me.txtStart)
End Sub

Private Sub UserForm_Initialize()
    If connIP = Empty Then setGlobalVariant '//�������� �缳��
    If checkLogin = 0 Then f_login.Show '//�α���üũ
    If checkLogin = 0 Then End '--//�α��� ���� �� ���ν��� ����
    
    '--//���ʼ���
    Me.cmdClose.Cancel = True
    TB1 = "op_system.v0_pstaff_information_all" '--//����������
    TB2 = "op_system.db_visa" '--//�����̷�
    
    '--//���ѿ� ���� ��Ʈ�� ����
    Call HideDeleteButtonByUserAuth
    
    '--//��Ʈ�� ����
    Me.lstHistory.Enabled = False
    Me.txtStart.Enabled = False
    Me.txtEnd.Enabled = False
    Me.cboVisa.Enabled = False
    Me.txtMemo.Enabled = False
    Me.cmdEdit.Enabled = False
    Me.cmdAdd.Enabled = False
    Me.cmdAdd.Visible = False
    Me.cmdCancel.Visible = False
    Me.cmdNew.Enabled = False
    Me.cmdDelete.Enabled = False
    Me.chkPresent.Visible = False
    
    '--//�޺��ڽ� ������ �߰�
    Me.cboVisa.Clear
'    Me.cboVisa.AddItem "��������"
'    Me.cboVisa.AddItem "�������"
'    Me.cboVisa.AddItem "�л�����"
'    Me.cboVisa.AddItem "�������"
'    Me.cboVisa.AddItem "��������"
'    Me.cboVisa.AddItem "�ڿ������ں���"
'    Me.cboVisa.AddItem "��ȥ����"
    strSql = "SELECT a.visa_nm FROM op_system.a_visa a;"
    Call makeListData(strSql, "op_system.a_visa")
    If cntRecord > 0 Then
        Me.cboVisa.List = LISTDATA
    End If
    Call sbClearVariant
    
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
    
    Me.lstHistory.Clear
    
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
        If Me.chkNative.Value Then
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
        Else
            '--//��ȸ�ڵ�, ��ȸ��
            strSql = "SELECT a.`�����ȣ`,a.`��ȸ��`,a.`�ѱ��̸�(����)`,a.`��å` " & _
                       "FROM " & TB1 & " a " & _
                       "WHERE (a.`�ѱ��̸�(����)` LIKE '%" & Me.txtChurchNM & "%' OR a.`��ȸ��` LIKE '%" & Me.txtChurchNM & "%'" & _
                       " OR a.`�����̸�` LIKE '%" & Me.txtChurchNM & "%'" & _
                       " OR a.`�����ȣ` LIKE '%" & Me.txtChurchNM & "%')" & _
                       " AND a.`����` = '���ѹα�'" & _
                       " AND a.`�����μ�` = " & SText(USER_DEPT) & _
                    " UNION " & _
                    "SELECT b.`����ڻ���`,b.`��ȸ��`,b.`����ѱ��̸�(����)`,b.`�����å` " & _
                       "FROM " & TB1 & " b " & _
                       "WHERE (b.`����ѱ��̸�(����)` LIKE '%" & Me.txtChurchNM & "%' OR b.`��ȸ��` LIKE '%" & Me.txtChurchNM & "%'" & _
                       " OR b.`��𿵹��̸�` LIKE '%" & Me.txtChurchNM & "%'" & _
                       " OR b.`����ڻ���` LIKE '%" & Me.txtChurchNM & "%')" & _
                       " AND b.`�����` = '���ѹα�'" & _
                       " AND b.`�����μ�` = " & SText(USER_DEPT) & ";"
        End If
    Case TB2
        strSql = "SELECT a.visa_cd,a.lifeno,a.start_dt,if(a.end_dt='9999-12-31','����',a.end_dt),a.`visa`,a.`memo` " & _
                "FROM " & TB2 & " a " & _
                "WHERE a.lifeno = " & SText(Me.lstPStaff.List(Me.lstPStaff.listIndex)) & " ORDER BY a.`start_dt`;"
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
                    "SET a.start_dt = " & SText(Me.txtStart) & _
                    ", a.end_dt = " & SText(IIf(Me.txtEnd = "����", DateSerial(9999, 12, 31), Me.txtEnd)) & _
                    ", a.visa = " & SText(Me.cboVisa) & ", a.memo = " & SText(Me.txtMemo) & _
                    " WHERE a.visa_cd = " & SText(.List(.listIndex)) & ";"
        queryKey = .listIndex
        End With
    Case Else
    End Select
    makeUpdateSQL = strSql
End Function
Private Function makeInsertSQL(ByVal tableNM As String, argData As T_VISA) As String
    
    Select Case tableNM
    Case TB1
    Case TB2
        strSql = "INSERT INTO " & TB2 & " VALUES(DEFAULT," & _
                    SText(argData.lifeNo) & "," & _
                    SText(argData.START_DT) & "," & _
                    SText(argData.END_DT) & "," & _
                    SText(argData.Visa) & "," & _
                    SText(argData.memo) & ");"
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
            strSql = "DELETE FROM " & TB2 & " WHERE visa_cd = " & SText(.List(.listIndex)) & ";"
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
    Me.cboVisa = ""
    Me.txtStart.Value = ""
    Me.txtEnd.Value = "����"
    Me.txtMemo = ""
End Sub
Private Function fnData_Validation()
'---------------------------------------
'������ �Է°��� ���� ������ ��ȿ�� �˻�
'TRUE: �̻����, FALSE: �߸���.
'---------------------------------------
    fnData_Validation = True '�����Ͱ� ��ȿ�ϴٴ� ���� �Ͽ� ����
    
    If Not IsDate(Me.txtStart) Then
        MsgBox "��¥ ������ �߸� �Ǿ����ϴ�. �������� �ٽ� Ȯ���� �ּ���.", vbCritical, banner
        Set txtBox_Focus = Me.txtStart: fnData_Validation = False: Exit Function
    End If
    
    If Not IsDate(Me.txtEnd) And Me.txtEnd <> "����" Then
        MsgBox "��¥ ������ �߸� �Ǿ����ϴ�. �������� �ٽ� Ȯ���� �ּ���.", vbCritical, banner
        Set txtBox_Focus = Me.txtEnd: fnData_Validation = False: Exit Function
    End If
    
    If Me.txtEnd <> "����" Then
        If CDate(Me.txtEnd) <= CDate(Me.txtStart) Then
            MsgBox "�������� �����Ϻ��� �۰ų� ���� �� �����ϴ�.", vbCritical, banner
            Set txtBox_Focus = Me.txtStart: fnData_Validation = False: Exit Function
        End If
    End If
    
    If Me.cboVisa = "" Or Me.txtStart = "" Or Me.txtEnd = "" Then
        MsgBox "�ʼ� �Է°��� �����Ǿ����ϴ�. �ٽ� Ȯ�����ּ���.", vbCritical, banner
        If Me.cboVisa = "" Then Set txtBox_Focus = Me.cboVisa: fnData_Validation = False: Exit Function
        If Me.txtStart = "" Then Set txtBox_Focus = Me.txtStart: fnData_Validation = False: Exit Function
        If Me.txtEnd = "" Then Set txtBox_Focus = Me.txtEnd: fnData_Validation = False: Exit Function
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

Private Sub INPUTMODE(ByVal argBoolean As Boolean)
    Me.cmdNew.Visible = Not argBoolean
    Me.cmdEdit.Visible = Not argBoolean
    Me.cmdDelete.Visible = Not argBoolean
    Me.cmdCancel.Visible = argBoolean
    Me.cmdCancel.Enabled = argBoolean
    Me.cmdAdd.Visible = argBoolean
    Me.cmdAdd.Enabled = argBoolean
    
    Me.txtStart.Enabled = argBoolean
    Me.txtEnd.Enabled = argBoolean
    Me.cboVisa.Enabled = argBoolean
    Me.txtMemo.Enabled = argBoolean
    Me.chkPresent.Visible = argBoolean
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





