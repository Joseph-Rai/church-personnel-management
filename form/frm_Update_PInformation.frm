VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Update_PInformation 
   Caption         =   "�λ��ڷ� ����������"
   ClientHeight    =   10530
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   16365
   OleObjectBlob   =   "frm_Update_PInformation.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "frm_Update_PInformation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim queryKey As Integer '//����Ʈ ��ġ ��ȯ�� ���� id
Dim txtBox_Focus As MSForms.textBox
Dim errorMessageAppeared As Boolean '--//����Լ����� �����޽����� ȣ�� ���� ����ŭ ��µǴ� ���� ����

Public Enum INPUTMODE
    INPUT_FOR_ALL
    INPUT_FOR_PART
End Enum

Private Sub chkChild1_Click()
    
    Dim control As MSForms.control

    If Me.chkChild1 = True Then
        If Me.txtLifeNo_Child1 = "" Then
            '--//Child1 �����ȣ�� ������ Input Mode
            ActivateInputMode Me.frameChild1.controls, True, INPUT_FOR_PART
        Else
            '--//Child1 �����ȣ�� ������ Frame �� Controls �ܼ� Ȱ��ȭ
            ActivateFrame Me.frameChild1, True
        End If
    Else
        '--//Child1 �����ȣ�� ������ Frame �� Controls ��Ȱ��ȭ
        ActivateFrame Me.frameChild1, False
    End If
    
End Sub

Private Sub chkChild2_Click()

    Dim control As MSForms.control
    
    If Me.chkChild2 = True Then
        If Me.txtLifeNo_Child2 = "" Then
            '--//Child2 �����ȣ�� ������ Input Mode
            ActivateInputMode Me.frameChild2.controls, True, INPUT_FOR_PART
        Else
            '--//Child2 �����ȣ�� ������ Frame �� Controls �ܼ� Ȱ��ȭ
            ActivateFrame Me.frameChild2, True
        End If
    Else
        '--//Child2 �����ȣ�� ������ Frame �� Controls ��Ȱ��ȭ
        ActivateFrame Me.frameChild2, False
    End If
    
End Sub

Private Sub chkChild3_Click()
    
    Dim control As MSForms.control
    
    If Me.chkChild3 = True Then
        If Me.txtLifeNo_Child3 = "" Then
            '--//Child3 �����ȣ�� ������ Input Mode
            ActivateInputMode Me.frameChild3.controls, True, INPUT_FOR_PART
        Else
            '--//Child3 �����ȣ�� ������ Frame �� Controls �ܼ� Ȱ��ȭ
            ActivateFrame Me.frameChild3, True
        End If
    Else
        '--//Child3 �����ȣ�� ������ Frame �� Controls ��Ȱ��ȭ
        ActivateFrame Me.frameChild3, False
    End If
    
End Sub

Private Sub chkSpouse_Click()
    
    Dim control As MSForms.control
    
    If Me.chkSpouse = True Then
        If Me.txtLifeNo_Spouse = "" Then
            '--//Spouse �����ȣ�� ������ Input Mode
            ActivateInputMode Me.frameSpouse.controls, True, INPUT_FOR_PART
        Else
            '--//Spouse �����ȣ�� ������ Frame �� Controls �ܼ� Ȱ��ȭ
            ActivateFrame Me.frameSpouse, True
            Me.cmdTransferSpouse.Enabled = True
        End If
    Else
        '--//Spouse �����ȣ�� ������ Frame �� Controls ��Ȱ��ȭ
        ActivateFrame Me.frameSpouse, False
        Me.cmdTransferSpouse.Enabled = False
    End If
    
End Sub

Private Sub cmdCopyLifeNo_Child1_Click()
    If Me.txtLifeNo_Child1 = Empty Then Exit Sub
    CopyStartAnimation Me.cmdCopyLifeNo_Child1
    CopyToClipboard Me.txtLifeNo_Child1
    CopyEndAnimation Me.cmdCopyLifeNo_Child1
End Sub

Private Sub cmdCopyLifeNo_Child2_Click()
    If Me.txtLifeNo_Child2 = Empty Then Exit Sub
    CopyStartAnimation Me.cmdCopyLifeNo_Child2
    CopyToClipboard Me.txtLifeNo_Child2
    CopyEndAnimation Me.cmdCopyLifeNo_Child2
End Sub

Private Sub cmdCopyLifeNo_Child3_Click()
    If Me.txtLifeNo_Child3 = Empty Then Exit Sub
    CopyStartAnimation Me.cmdCopyLifeNo_Child3
    CopyToClipboard Me.txtLifeNo_Child3
    CopyEndAnimation Me.cmdCopyLifeNo_Child3
End Sub

Private Sub cmdCopyLifeNo_Click()
    If Me.txtLifeNo = Empty Then Exit Sub
    CopyStartAnimation Me.cmdCopyLifeNo
    CopyToClipboard Me.txtLifeNo
    CopyEndAnimation Me.cmdCopyLifeNo
End Sub

Private Sub cmdCopyLifeNo_Spouse_Click()
    If Me.txtLifeNo_Spouse = Empty Then Exit Sub
    CopyStartAnimation Me.cmdCopyLifeNo_Spouse
    CopyToClipboard Me.txtLifeNo_Spouse
    CopyEndAnimation Me.cmdCopyLifeNo_Spouse
End Sub

Private Sub cmdTransfer_Click()

    Dim pStaff As New PastoralStaff
    Dim pWife As New PastoralWife
    Dim pStaffDao As New PastoralStaffDao
    Dim pWifeDao As New PastoralWifeDao
    
    pStaff.ParseFromForm Me
    pWife.ParseFromForm Me

    If Me.chkTransfer = False Then
        '--//������Ʈ �� Ȯ������
        If MsgBox("Ÿ�μ����� ������ �����ȣ�� �����ڸ� �ű� �߰��ϸ� ���� �����Ͱ� �̰�ó�� �˴ϴ�." & vbNewLine & _
            "��� �Ͻðڽ��ϱ�?", vbQuestion + vbYesNo, banner) = vbNo Then
            Exit Sub
        End If
        
        Me.chkTransfer.Value = True
        
        '--//pWife�� �Ѳ����� �����̴� ������ ����Ű�� �޽���
        If pWife.lifeNo <> "" Then
            Me.chkTransferSpouse.Value = True
            MsgBox "����ڵ� �Բ� �̵�ó�� �ǵ��� ��" & Me.cmdTransferSpouse.Caption & "���� �Բ� üũ �˴ϴ�.", vbOKOnly, banner
        End If
    Else
        '--//������Ʈ �� Ȯ������
        If MsgBox("Ÿ�μ� �̵�ó���� ��ҵ˴ϴ�. Ÿ�μ����� �����͸� �̰����� �� �����ϴ�." & vbNewLine & _
            "��� �Ͻðڽ��ϱ�?", vbQuestion + vbYesNo, banner) = vbNo Then
            Exit Sub
        End If
        
        '--//pStaff�� Suspend ������Ʈ
        Me.chkTransfer.Value = False
        
        '--//pWife�� �Ѳ����� �����̴� ������ ����Ű�� �޽���
        If pWife.lifeNo <> "" Then
            Me.chkTransferSpouse.Value = False
            MsgBox "��" & Me.cmdTransferSpouse.Caption & "���� �Բ� üũ ���� �Ǿ����ϴ�.", vbOKOnly, banner
        End If
    End If
    
    '--//pStaff�� Suspend ������Ʈ
    pStaff.Suspend = Me.chkTransfer.Value
    If pWife.lifeNo <> "" Then
        pWife.Suspend = Me.chkTransferSpouse.Value
    End If
    
    '--//����� ���� ����
    pStaffDao.Save pStaff
    pWifeDao.Save pWife
    
    '--//������ ���ΰ�ħ
    If Me.lstPStaff.listIndex <> -1 Then
        lstPStaff_Click
    End If

End Sub

Private Sub cmdTransferSpouse_Click()
    
    '--//ȥ�ΰ��谡 ������ ���¿��� ����ڸ� �̵��ϸ� ���� ���Ἲ�� ��ġ�Ƿ� ������� ����
    If Me.chkTransferSpouse = True Then
        If Me.chkTransfer.Value = True Then
            MsgBox "����ڸ� �ܵ����� ������ �� �����ϴ�." & vbNewLine & _
                    "����ڸ� üũ�ϱ� ���Ͻø� ������� üũ�� ���� ������ �� �����ϼ���.", vbCritical + vbOKOnly, banner
            Exit Sub
        End If
        
        If MsgBox("������� Ÿ�μ� �̵�ó���� ��ҵ˴ϴ�. Ÿ�μ����� �����͸� �̰����� �� �����ϴ�." & vbNewLine & _
            "��� �Ͻðڽ��ϱ�?", vbQuestion + vbYesNo, banner) = vbNo Then
            Exit Sub
        Else
            Me.chkTransferSpouse.Value = False
        End If
    Else
        If MsgBox("�� ����� ����� ������ ����Ǿ��� ��쿡�� ����մϴ�." & vbNewLine & _
                    "��� �����Ͻðڽ��ϱ�", vbQuestion + vbYesNo, banner) = vbNo Then
            Exit Sub
        Else
            Me.chkTransferSpouse.Value = True
        End If
    End If
    
    Dim pWife As New PastoralWife
    Dim pWifeDao As New PastoralWifeDao
    
    pWife.ParseFromForm Me
    pWife.Suspend = Me.chkTransferSpouse.Value
    
    pWifeDao.Save pWife
    
    '--//������ ���ΰ�ħ
    If Me.lstPStaff.listIndex <> -1 Then
        lstPStaff_Click
    End If
    
End Sub

Private Sub cmdDelete_Child1_Click()

    Dim pStaff As New PastoralStaff
    Dim pStaffDao As New PastoralStaffDao
    
    pStaff.ParseFromForm Me
    
    '--//üũ�ڽ��� üũ ���� �� ������ �ڳ� ���� �����Ƿ� ���ν��� ����
    If pStaff.LifeNoChild1 = "" Then
        MsgBox "������ �ڳ�1 ������ �����ϴ�.", vbCritical, banner
        Exit Sub
    End If
    
    '--//���� ���࿩�� ��Ȯ��
    If MsgBox("�ڳ�1 ������ �����˴ϴ�. ���� ���� �Ͻðڽ��ϱ�?", vbQuestion + vbYesNo, banner) = vbNo Then
        Exit Sub
    End If
    
    '--//�ڳ�1 ���� ����
    pStaffDao.DeleteChild1 pStaff
    
    '--//�޼��� �ڽ�
    MsgBox "�ڳ�1 ������ ���� �Ǿ����ϴ�.", , banner
    Call lstPStaff_Click
    
End Sub

Private Sub cmdDelete_Child2_Click()
    
    Dim pStaff As New PastoralStaff
    Dim pStaffDao As New PastoralStaffDao
    
    pStaff.ParseFromForm Me
    
    If pStaff.LifeNoChild2 = "" Then
        MsgBox "������ �ڳ�2 ������ �����ϴ�.", vbCritical, banner
        Exit Sub
    End If
    
    '--//���� ���࿩�� ��Ȯ��
    If MsgBox("�ڳ�2 ������ �����˴ϴ�. ���� ���� �Ͻðڽ��ϱ�?", vbQuestion + vbYesNo, banner) = vbNo Then
        Exit Sub
    End If
    
    '--//�ڳ�2 ���� ����
    pStaffDao.DeleteChild2 pStaff


    '--//�޼��� �ڽ�
    MsgBox "�ڳ�2 ������ ���� �Ǿ����ϴ�.", , banner
    Call lstPStaff_Click
End Sub

Private Sub cmdDelete_Child3_Click()
    
    Dim pStaff As New PastoralStaff
    Dim pStaffDao As New PastoralStaffDao
    
    pStaff.ParseFromForm Me
    
    If pStaff.LifeNoChild3 = "" Then
        MsgBox "������ �ڳ�3 ������ �����ϴ�.", vbCritical, banner
        Exit Sub
    End If
    
    '--//���� ���࿩�� ��Ȯ��
    If MsgBox("�ڳ�3 ������ �����˴ϴ�. ���� ���� �Ͻðڽ��ϱ�?", vbQuestion + vbYesNo, banner) = vbNo Then
        Exit Sub
    End If
    
    '--//�ڳ�3 ���� ����
    pStaffDao.DeleteChild3 pStaff

    '--//�޼��� �ڽ�
    MsgBox "�ڳ�3 ������ ���� �Ǿ����ϴ�.", , banner
    Call lstPStaff_Click
End Sub

Private Sub cmdDelete_Spouse_Click()
    
    Dim pWife As New PastoralWife
    Dim pWifeDao As New PastoralWifeDao
    
    pWife.ParseFromForm Me
    
    If pWife.lifeNo = "" Then
        MsgBox "������ ����� ������ �����ϴ�.", vbCritical, banner
        Exit Sub
    End If
    
    '--//���� ���࿩�� ��Ȯ��
    If MsgBox("����� ������ �����˴ϴ�. ���� ���� �Ͻðڽ��ϱ�?", vbQuestion + vbYesNo, banner) = vbNo Then
        Exit Sub
    End If
    
    '--//����� ���� ����
    pWifeDao.Delete pWife

    '--//�޼��� �ڽ�
    MsgBox "����� ������ ���� �Ǿ����ϴ�.", , banner
    lstPStaff_Click
    
End Sub

Private Sub cmdExtendChild_Click()
    Me.Width = 820
    Me.cmdExtendChild.Visible = False
    Me.cmdMinimize.Visible = True
End Sub

Private Sub cmdMinimize_Click()
    Me.Width = 545
    Me.cmdMinimize.Visible = False
    Me.cmdExtendChild.Visible = True
End Sub

Private Sub cmdFamily_Click()
    argShow2 = 1
    frm_Update_FamilyInfo.Show
End Sub

Private Sub cmdFamily_Spouse_Click()
    argShow2 = 2
    frm_Update_FamilyInfo.Show
End Sub

Private Sub savePhoto(lifeNo As String, filePath As String)

    Dim pPhoto As New PStaffPhoto
    Dim pPhotoDao As New PStaffPhotoDao
    Dim stream As New ADODB.stream
    
    stream.Type = adTypeBinary
    stream.Open
    stream.LoadFromFile filePath
    
    pPhoto.lifeNo = lifeNo
    pPhoto.photo = encodeBase64(stream.Read)
    pPhotoDao.Save pPhoto

End Sub

Private Sub lblPic_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    
    Dim filePath As String
    
    filePath = Application.GetOpenFilename("Image File Format, *.jpg;*.jpeg; *.bmp; *.tif; *.png", , "Photo Select")
    
    If filePath = "False" Then
        Exit Sub
    End If
    
    '--//DB�� ���� ����
    savePhoto Me.txtLifeNo, filePath
    
    '--//�����޽���
    MsgBox "������ ���������� ����Ǿ����ϴ�.", , banner
    
    '--//�󺧿� ��������
    InsertPicToLabel Me.lblPic, Me.txtLifeNo
    
End Sub

Private Sub lblPic_Spouse_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    
    If Me.txtLifeNo_Spouse = "" Then
        Exit Sub
    End If
    
    Dim filePath As String
    
    filePath = Application.GetOpenFilename("Image File Format, *.jpg;*.jpeg; *.bmp; *.tif; *.png", , "Photo Select")
    
    If filePath = "False" Then
        Exit Sub
    End If
    
    '--//DB�� ���� ����
    savePhoto Me.txtLifeNo_Spouse, filePath
    
    '--//�����޽���
    MsgBox "������ ���������� ����Ǿ����ϴ�.", , banner
    
    '--//�󺧿� ��������
    InsertPicToLabel Me.lblPic_Spouse, Me.txtLifeNo_Spouse
    
End Sub

Private Sub lblPic_Child1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim filePath As String
    
    filePath = Application.GetOpenFilename("Image File Format, *.jpg;*.jpeg; *.bmp; *.tif; *.png", , "Photo Select")
    
    If filePath = "False" Then
        Exit Sub
    End If
    
    '--//DB�� ���� ����
    savePhoto Me.txtLifeNo_Child1, filePath
    
    '--//�����޽���
    MsgBox "������ ���������� ����Ǿ����ϴ�.", , banner
    
    '--//�󺧿� ��������
    InsertPicToLabel Me.lblPic_Child1, Me.txtLifeNo_Child1
    
End Sub

Private Sub lblPic_Child2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim filePath As String
    
    filePath = Application.GetOpenFilename("Image File Format, *.jpg;*.jpeg; *.bmp; *.tif; *.png", , "Photo Select")
    
    If filePath = "False" Then
        Exit Sub
    End If
    
    '--//DB�� ���� ����
    savePhoto Me.txtLifeNo_Child2, filePath
    
    '--//�����޽���
    MsgBox "������ ���������� ����Ǿ����ϴ�.", , banner
    
    '--//�󺧿� ��������
    InsertPicToLabel Me.lblPic_Child2, Me.txtLifeNo_Child2
    
End Sub

Private Sub lblPic_Child3_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim filePath As String
    
    filePath = Application.GetOpenFilename("Image File Format, *.jpg;*.jpeg; *.bmp; *.tif; *.png", , "Photo Select")
    
    If filePath = "False" Then
        Exit Sub
    End If
    
    '--//DB�� ���� ����
    savePhoto Me.txtLifeNo_Child3, filePath
    
    '--//�����޽���
    MsgBox "������ ���������� ����Ǿ����ϴ�.", , banner
    
    '--//�󺧿� ��������
    InsertPicToLabel Me.lblPic_Child3, Me.txtLifeNo_Child3
    
End Sub

Private Sub lstPStaff_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Un'HookListBoxScroll
End Sub

Private Sub lstPStaff_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Me.lstPStaff.ListCount > 0 Then
        'HookListBoxScroll Me, Me.lstPStaff
    End If
End Sub

Private Sub txtBirthday_Change()
    Call Date_Format(Me.txtBirthday)
End Sub

Private Sub txtBirthday_Child1_Change()
    Call Date_Format(Me.txtBirthday_Child1)
End Sub

Private Sub txtBirthday_Child2_Change()
    Call Date_Format(Me.txtBirthday_Child2)
End Sub

Private Sub txtBirthday_Child3_Change()
    Call Date_Format(Me.txtBirthday_Child3)
End Sub

Private Sub txtBirthday_Spouse_Change()
    Call Date_Format(Me.txtBirthday_Spouse)
End Sub

Private Sub txtLifeNo_Change()
'    Me.chkTransfer.Enabled = True
End Sub

Private Sub CopyStartAnimation(cmdBox As MSForms.CommandButton)
    
    Dim i As Integer
    
    For i = 0 To 25
'        Sleep 8
        cmdBox.Left = cmdBox.Left - 1
        cmdBox.Width = cmdBox.Width + 1
        Me.Repaint
    Next
    
    With cmdBox
        .Caption = "Copied"
        .BackColor = &HC0C0FF
    End With
    Me.Repaint
    
    Application.Wait DateAdd("s", 1, Now)
'    Sleep 700
    
End Sub

Private Sub CopyEndAnimation(cmdBox As MSForms.CommandButton)

    Dim i As Integer
    
    For i = 0 To 25
'        Sleep 8
        cmdBox.Left = cmdBox.Left + 1
        cmdBox.Width = cmdBox.Width - 1
        Me.Repaint
    Next
    
    With cmdBox
        .Caption = "..."
        .BackColor = &HC0FFC0
    End With
    Me.Repaint

End Sub

Private Sub CopyToClipboard(text As String)
    On Error GoTo ErrHandler
    Static obj As DataObject
    
    If obj Is Nothing Then Set obj = New DataObject
    obj.SetText text
    obj.PutInClipboard
ErrHandler:
End Sub

Private Sub txtNationality_Enter()
    If Me.txtNationality = "" Then
        argShow = 3
        frm_Search_Country.Show
    End If
End Sub
Private Sub txtNationality_Spouse_Enter()
    If Me.txtNationality_Spouse = "" Then
        argShow = 4
        frm_Search_Country.Show
    End If
End Sub

Private Sub txtOrdinationPrayer_dt_Change()
    Call Date_Format(Me.txtOrdinationPrayer_dt)
End Sub

Private Sub txtOvs_dt_Change()
    Call Date_Format(Me.txtOvs_dt)
End Sub

Private Sub txtSalary_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Dim su As Double
    
    Me.txtSalary = fnExtract(Me.txtSalary)
    
    If Len(Me.txtSalary.Value) = 0 Then
        su = 0
    Else
        su = CDbl(Me.txtSalary.text)
    End If

    Me.txtSalary.text = Format(su, "#,##0")
    Me.txtSalary.SelStart = Len(Me.txtSalary.Value)
End Sub

Private Sub txtWedding_dt_Change()
    Call Date_Format(Me.txtWedding_dt)
End Sub

Private Sub UserForm_Initialize()
    If connIP = Empty Then setGlobalVariant '//�������� �缳��
    If checkLogin = 0 Then f_login.Show '//�α���üũ
    If checkLogin = 0 Then End '--//�α��� ���� �� ���ν��� ����
    
    '--//���ʼ���
    Me.cmdClose.Cancel = True
    Me.chkSpouse.Visible = False
    Me.chkChild1.Visible = False
    Me.chkChild2.Visible = False
    Me.chkChild3.Visible = False
    Me.cmdAdd.Visible = False
    Me.cmdCancel.Visible = False
    Me.cmdDelete_Spouse.Enabled = False
    Me.cmdDelete_Child1.Enabled = False
    Me.cmdDelete_Child2.Enabled = False
    Me.cmdDelete_Child3.Enabled = False
    Me.cmdTransfer.Enabled = False
    Me.cmdTransferSpouse.Enabled = False
    Me.chkTransfer.Enabled = False
    Me.chkTransferSpouse.Enabled = False
    
    '--//���ѿ� ���� ��Ʈ�� ����
    HideDeleteButtonByUserAuth
    ExtraControlsEnable False
    cmdMinimize_Click
    
    '--//�޺��ڽ� ����
    With Me.cboBaptism
        .AddItem "��"
        .AddItem "��"
    End With
    
    '--//����Ʈ�ڽ� ����
    With Me.lstPStaff
        .ColumnCount = 4
        .ColumnHeads = False
        .ColumnWidths = "0,120,70,50" '�����ȣ, ��ȸ��, �ѱ��̸�(����), ��å
        .TextAlign = fmTextAlignLeft
        .Font = "����"
    End With
    
    '--//�˻�â Setfocus
    Me.txtSearchText.SetFocus
    
End Sub
Private Sub lstPStaff_Click()
    
    Dim filePath As String
    Dim FileName As String
    Dim control As MSForms.control
    
    '--//��Ʈ�� ����
    ExtraControlsEnable True
    Me.chkSpouse.Visible = True
    Me.chkChild1.Visible = True
    Me.chkChild2.Visible = True
    Me.chkChild3.Visible = True
    HideDeleteButtonByUserAuth
    
    Me.lblPic.Picture = LoadPicture("")
    Me.lblPic_Spouse.Picture = LoadPicture("")
    Me.lblPic_Child1.Picture = LoadPicture("")
    Me.lblPic_Child2.Picture = LoadPicture("")
    Me.lblPic_Child3.Picture = LoadPicture("")
    
    '--//�λ絥���� �ؽ�Ʈ�ڽ��� �߰�
    If Me.lstPStaff.listIndex = -1 Then
        Exit Sub
    End If
    
    Dim pStaff As New PastoralStaff
    Dim pStaffDao As New PastoralStaffDao
    Dim pWife As New PastoralWife
    Dim pWifeDao As New PastoralWifeDao
    Dim lifeNo As String
    
    With Me.lstPStaff
        lifeNo = .List(.listIndex)
        Set pStaff = pStaffDao.FindByLifeNo(lifeNo)
        Set pWife = pWifeDao.FindByPStaff(pStaff)
    End With
    
    '--//��� ä���
    FillOutForm pStaff, pWife
    
    '--//üũ�ڽ� Ŭ�� ���ν��� �������� ��Ȱ�� �ؽ�Ʈ�ڽ� ó��
    ActivateFrame Me.framePStaff, True
    chkSpouse_Click
    chkChild1_Click
    chkChild2_Click
    chkChild3_Click
    
    '--//���̺� ���� ���󺹱�
    For Each control In Me.controls
        If TypeName(control) = "Label" And Not control.Name Like "*Info*" Then
            ChangeLabelColor control, vbBlack
        End If
    Next
    
    '--//Frame �ܺ� ��Ʈ�� Ȱ��ȭ/��Ȱ��ȭ ���� �����ϱ�
    ComboBoxEnable Me.cboBaptism, True
    TextBoxEnable Me.txtWedding_dt, True
    TextBoxEnable Me.txtOrdinationPrayer_dt, True
    TextBoxEnable Me.txtSalary, True
    
    '--//�ܱ��� �ؿܹ߷���,������� ���Ƴ���
    TextBoxEnable Me.txtOvs_dt, Me.txtNationality = "���ѹα�"
    TextBoxEnable Me.txtTheological_Order, Me.txtNationality = "���ѹα�"
    
    '--//�����ȣ �ִ� �е鸸 ���������� �߰���ư Ȱ��ȭ
    ControlEnable Me.cmdFamily, Me.txtLifeNo <> ""
    ControlEnable Me.cmdFamily_Spouse, Me.txtLifeNo_Spouse <> ""
    ControlEnable Me.cmdTransfer, Me.txtLifeNo <> ""
    
    '--//�����, �ڳ�1,2,3 �ִ� ��� üũ�ڽ� ��Ȱ��ȭ
    ControlEnable Me.chkSpouse, Me.txtLifeNo_Spouse = ""
    ControlEnable Me.chkChild1, Me.txtLifeNo_Child1 = ""
    ControlEnable Me.chkChild2, Me.txtLifeNo_Child2 = ""
    ControlEnable Me.chkChild3, Me.txtLifeNo_Child3 = ""
    ControlEnable Me.cmdDelete_Spouse, Me.txtLifeNo_Spouse <> ""
    ControlEnable Me.cmdDelete_Child1, Me.txtLifeNo_Child1 <> ""
    ControlEnable Me.cmdDelete_Child2, Me.txtLifeNo_Child2 <> ""
    ControlEnable Me.cmdDelete_Child3, Me.txtLifeNo_Child3 <> ""
    
    '--//������ �����߰�
    InsertPicToLabel Me.lblPic, pStaff.lifeNo
    
    '--//�ڳ�1 �����߰�
    If Me.txtLifeNo_Child1 <> "" Then
        InsertPicToLabel Me.lblPic_Child1, pStaff.LifeNoChild1
    End If
    
    '--//�ڳ�2 �����߰�
    If Me.txtLifeNo_Child2 <> "" Then
        InsertPicToLabel Me.lblPic_Child2, pStaff.LifeNoChild2
    End If
    
    '--//�ڳ�3 �����߰�
    If Me.txtLifeNo_Child3 <> "" Then
        InsertPicToLabel Me.lblPic_Child3, pStaff.LifeNoChild3
    End If
    
    '--//����� �����߰�
    If Me.txtLifeNo_Spouse <> "" Then
        InsertPicToLabel Me.lblPic_Spouse, pWife.lifeNo
    End If

End Sub

Private Sub cmdSearch_Click()
    
    Dim pStaffInfo As New PStaffInfoView
    Dim pStaffInfoDao As New PStaffInfoViewDao
    Dim pStaffInfoList As Object
    
    '--//�˻����ǿ� ���� pStaffInfo ��ü ����Ʈ �ҷ�����
    Set pStaffInfoList = pStaffInfoDao.FindBySearchText(Me.txtSearchText, Me.chkAll.Value)
    
    '--//����Ʈ �ʱ�ȭ
    Me.lstPStaff.Clear
    
    '--//�˻��� ��� ����Ʈ�� ä���ֱ�
    Dim pStaffInfoTemp As New PStaffInfoView
    With Me.lstPStaff
        For Each pStaffInfoTemp In pStaffInfoList
            Me.lstPStaff.AddItem pStaffInfoTemp.lifeNo
            .List(.ListCount - 1, 1) = pStaffInfoTemp.ChurchNameKo
            .List(.ListCount - 1, 2) = pStaffInfoTemp.NameKoAndTitle
            .List(.ListCount - 1, 3) = pStaffInfoTemp.position
        Next
    End With
End Sub
Private Sub txtChurch_Change()
    Me.txtChurch.BackColor = RGB(255, 255, 255)
End Sub
Private Sub cmdCancel_Click()
    
    Dim control As MSForms.control
    
    '--//��Ʈ�� ���󺹱�
'    HideCmdBtnForInput False
    ActivateInputMode Me.controls, False, INPUT_FOR_ALL
    
    If Me.txtSearchText <> "" Then
        lstPStaff_Click
    Else
        
'        ExtraControlsEnable False
'        '--//���̺� ���� ���󺹱�
'        For Each control In Me.controls
'            If TypeName(control) = "Label" And Not control.name Like "*Info*" Then
'                ChangeLabelColor control, vbBlack
'            End If
'        Next
    End If
    
    '--//�� ���ΰ�ħ
    cmdMinimize_Click
    lstPStaff_Click
    
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

'--//�����ڿ� ����� ��� �����մϴ�.
'--//�ڳ� ������ ������ ������ �� ���̺� ���� �����Ƿ� �Բ� ���� �˴ϴ�.
Private Sub cmdDelete_Click()
    
    If MsgBox("������ �����͸� �����Ͻðڽ��ϱ�?" & vbNewLine & "������ �����ʹ� ������ �� �����ϴ�.", vbQuestion + vbYesNo, banner) = vbNo Then
        Exit Sub
    End If
    
    Dim pStaff As New PastoralStaff
    Dim pWife As New PastoralWife
    Dim pStaffDao As New PastoralStaffDao
    Dim pWifeDao As New PastoralWifeDao
    
    '--//������ ��ü �Ľ�
    pStaff.ParseFromForm Me
    pWife.ParseFromForm Me
    
    If pWife.lifeNo <> "" Then
        If MsgBox("����ڵ� �Բ� ���� �˴ϴ�." & vbNewLine & "������ �����ʹ� ������ �� �����ϴ�." & vbNewLine & _
                  "��� �Ͻðڽ��ϱ�?", vbQuestion + vbYesNo, banner) = vbNo Then
            Exit Sub
        End If
    End If
    
    '--//DB���� ��ü ����
    pStaffDao.Delete pStaff
    pWifeDao.Delete pWife
    
    '--//�޼����ڽ�
    MsgBox "�ش� �����Ͱ� �����Ǿ����ϴ�.", , banner
    
    '--//����Ʈ�ڽ� ���ΰ�ħ
    InitializeTextBoxes Me.controls
    cmdSearch_Click
    lstPStaff_Click
    Me.lstPStaff.listIndex = -1
    
End Sub

'--//������ ������ ������ DB�� �ݿ��մϴ�.
Private Sub cmdEdit_Click()
    
    '--//������ ��ȿ�� �˻�
    If fnData_Validation(Me.controls) = False Then
On Error Resume Next
        txtBox_Focus.SetFocus
        txtBox_Focus.SelStart = 0
        txtBox_Focus.SelLength = Len(txtBox_Focus)
On Error GoTo 0
        Exit Sub
    End If
    
    '--//��������: pStaff1, pWife1 = ���� �� ��ü | pStaff2, pWife2 = ���� �� ��ü(DB��ü)
    '              pStaffDao, pWifeDao = ������ ������ ������Ʈ
    Dim pStaff1 As New PastoralStaff
    Dim pWife1 As New PastoralWife
    Dim pStaff2 As New PastoralStaff
    Dim pWife2 As New PastoralWife
    Dim pStaffDao As New PastoralStaffDao
    Dim pWifeDao As New PastoralWifeDao
    
    '--//���� �� ��ü(�� ���)
    pStaff1.ParseFromForm Me
    pWife1.ParseFromForm Me
    
    '--//���� �� ��ü(DB ���)
    Set pStaff2 = pStaffDao.FindByStaff(pStaff1)
    Set pWife2 = pWifeDao.FindByWife(pWife1)
    
    '--//PStaff ��ü�� PWife ��ü�� �ƹ��͵� �����Ȱ� ���ٸ� ���ν��� ����
    If pStaff1.IsEqual(pStaff2, True) And pWife1.IsEqual(pWife2, True) Then
        Exit Sub
    End If
    
    '--//PStaff ��ü ������Ʈ
    pStaffDao.Save pStaff1 '--//pStaff1�� ���� ������Ʈ
    
    If pWife1.lifeNo <> "" Then
        Dim pStaff3 As New PastoralStaff '--// pStaff3: pWife1�� ����ڷ� ��ϵ� ������
        Set pStaff3 = pStaffDao.FindByLifeNo(pWife2.lifeNoSpouse)
        If pStaff1.IsEqual(pStaff3) Or pStaff3.lifeNo = "" Then
            '--//���� ���� ���� �� ����� ������ �����ϴٸ�
            pWifeDao.Save pWife1 '--//pWife1�� ���� ������Ʈ
        Else
            '--//���� ���� ���� �� ����� ������ �ٸ��ٸ�
            Dim OvsDept As New OvsDepartment
            Dim ovsDeptDao As New OvsDepartmentDao
            Set OvsDept = ovsDeptDao.FindById(pStaff3.OvsDept) '--// �����μ�
            MsgBox "�Է��Ͻ� ���� �̹� �ٸ� ���� ����ڷ� ��ϵǾ� �ֽ��ϴ�." & vbNewLine & _
                    "����� ���Ͻø� ���� ���� ����ڿ��� ���� ���ּ���." & vbNewLine & vbNewLine & _
                    "�̸�: " & pStaff3.nameKo & vbNewLine & _
                    "�����ȣ: " & pStaff3.lifeNo & vbNewLine & _
                    "�����μ�: " & OvsDept.DeptName, vbYesNo, banner
            Exit Sub
        End If
    End If
    
    '--//�޼����ڽ�
    MsgBox "������Ʈ �Ǿ����ϴ�.", , banner
    
    '--//����Ʈ�ڽ� �ʱ�ȭ
    Call lstPStaff_Click
    
End Sub

Private Sub cmdNew_Click()
    
    Dim control As MSForms.control
    
    '--//��ǲ��� Ȱ��ȭ
    ActivateInputMode Me.controls, True, INPUT_FOR_ALL
    
End Sub

Private Sub cmdADD_Click()
    
    '--//������ ��ȿ�� �˻�
    If fnData_Validation(Me.controls) = False Then
On Error Resume Next
        txtBox_Focus.SetFocus
        txtBox_Focus.SelStart = 0
        txtBox_Focus.SelLength = Len(txtBox_Focus)
On Error GoTo 0
        Exit Sub
    End If
    
    Dim pStaff As New PastoralStaff '--// DB�� ���� �߰��� PastoralStaff ��ü
    Dim pWife As New PastoralWife '--// DB�� ��ü �߰��� PastoralWife ��ü
    Dim pWifeTemp As New PastoralWife '--// �ߺ�üũ�� ���� ��ü
    Dim pStaffTemp As New PastoralStaff '--// �ߺ�üũ�� ���� ��ü
    Dim pWifeDao As New PastoralWifeDao
    Dim pStaffDao As New PastoralStaffDao
    Dim iOk As Integer
    
    '--//��������� �̰��Ǿ��� ���, ���� �̰����� ���� �� �����Ƿ� �������� �ʰ� �̰���Ű�� ���� �÷��� ����
    Dim wasRunFlag As Boolean
    
    '--//�� �����κ��� ������, ��� ��ü �Ľ�
    pStaff.ParseFromForm Me
    pWife.ParseFromForm Me
    
    '--//������ �ߺ�üũ
    Set pStaffTemp = pStaffDao.FindByStaff(pStaff)
    If pStaff.IsEqual(pStaffTemp) Then
        '--//Suspend: False - �̵��Ұ� / True - �̵�����
        If pStaffTemp.Suspend = False Then
            '--//�ߺ��� ��� Ŀ���� �����߻�
            On Error GoTo PSTAFF_IS_ALREADY_REGISTERED
            Dim OvsDept As New OvsDepartment
            Dim ovsDeptDao As New OvsDepartmentDao
            Set OvsDept = ovsDeptDao.FindById(pStaffTemp.OvsDept)
            err.Raise vbObjectError + _
                    ERR_CODE_PSTAFF_IS_ALREADY_REGISTERED, , _
                    ERR_DESC_PSTAFF_IS_ALREADY_REGISTERED & _
                    "�����μ� ������ ����� Ȥ�� �����ڿ��� �����ϼ���." & vbNewLine & _
                    "�����μ�: " & OvsDept.DeptName
        Else
            '--//�̵������� �� �̵�ó��
            iOk = MsgBox("������ ��ϵ� ������ ������ �̹� �ֽ��ϴ�." & vbNewLine & vbNewLine & "�ҷ����ðڽ��ϱ�?", vbQuestion + vbYesNo, banner)
            If iOk = vbYes Then
                Set pStaff = pStaffTemp
                pStaff.OvsDept = USER_DEPT
                pStaff.Suspend = False
                wasRunFlag = True
            Else
                Exit Sub
            End If
        End If
    End If
    
    '--//pStaff ��ü DB�� ����
    pStaffDao.Save pStaff

'--------------------------------------------------------------------

    '--//����� �ߺ�üũ
    If Me.chkSpouse.Value = False Or Me.txtLifeNo_Spouse = "" Then GoTo PASS1
    
    pWifeTemp = pWifeDao.FindByWifeAndSpouseLifeNo(pWife)
    If pWife.IsEqual(pWifeTemp) <> "" Then
        '--//Suspend: False - �̵��Ұ� / True - �̵�����
        If pWifeTemp.Suspend = False Then
            '--//�ߺ��� ��� Ŀ���� �����߻�
            On Error GoTo WIFE_IS_ALREADY_REGISTERED_OF_OTHER
            Set pStaff = pStaffDao.FindByLifeNo(pWifeTemp.lifeNoSpouse)
            err.Raise vbObjectError + _
                ERR_CODE_WIFE_IS_ALREADY_REGISTERED_OF_OTHER, , _
                ERR_DESC_WIFE_IS_ALREADY_REGISTERED_OF_OTHER & _
                "��ϵ� �����: " & pStaff.nameKo & "(" & pStaff.lifeNo & ")"
        Else
            '--//�̵������� �� �̵�ó��
            If wasRunFlag Then
                iOk = vbYes
            Else
                iOk = MsgBox("������ ��ϵ� ��� ������ �̹� �ֽ��ϴ�." & vbNewLine & vbNewLine & "�ҷ����ðڽ��ϱ�?", vbQuestion + vbYesNo, banner)
            End If
            If iOk = vbYes Then
                Set pWife = pWifeTemp
                pWife.OvsDept = USER_DEPT
                pWife.lifeNoSpouse = pStaff.lifeNo
                pWife.Suspend = False
            Else
                Exit Sub
            End If
        End If
    End If
    
    '--//pWife ��ü DB�� ����
    pWifeDao.Save pWife
    
PASS1:

    '--//�����޼���
    MsgBox "�߰� �Ǿ����ϴ�." & vbNewLine & vbNewLine & "�߷� �̷��� �ݵ�� �߰����ּ���.", , banner
    
    '--//�߰��� �ο� �˻��Ͽ� ��Ͽ� ����
    Me.txtSearchText = pStaff.lifeNo
    Me.chkAll.Value = True
    Call cmdSearch_Click
    Call lstPStaff_Click
On Error Resume Next
    Me.lstPStaff.listIndex = 0
On Error GoTo 0
    
    '--//��ư���� �������
    HideCmdBtnForInput False
    HideDeleteButtonByUserAuth
    
DONE:
    Exit Sub
    
PSTAFF_IS_ALREADY_REGISTERED:
    MsgBox err.Description, vbCritical, banner
    Exit Sub
    
WIFE_IS_ALREADY_REGISTERED_OF_OTHER:
    MsgBox err.Description, vbCritical, banner
    queryKey = pWife.lifeNoSpouse
    Call returnListPosition3(Me, Me.lstPStaff.Name, queryKey)
    Exit Sub
End Sub

'---------------------------------------
'������ �Է°��� ���� ������ ��ȿ�� �˻�
'TRUE: �̻����, FALSE: �߸���.
'---------------------------------------
Private Function fnData_Validation(controls As MSForms.controls) As Boolean
    
    '--//��ȿ�� �����޼����� ���� Map ��ü ����
    Dim messageMap As New Collection
    messageMap.Add "������ �������", Me.txtBirthday.Name
    messageMap.Add "�ڳ�1 �������", Me.txtBirthday_Child1.Name
    messageMap.Add "�ڳ�2 �������", Me.txtBirthday_Child2.Name
    messageMap.Add "�ڳ�3 �������", Me.txtBirthday_Child3.Name
    messageMap.Add "����� �������", Me.txtBirthday_Spouse.Name
    messageMap.Add "�ؿ� ���� �߷���", Me.txtOvs_dt.Name
    messageMap.Add "ȥ����", Me.txtWedding_dt.Name
    messageMap.Add "�ȼ���", Me.txtOrdinationPrayer_dt.Name
    messageMap.Add "������ ����", Me.txtNationality.Name
    messageMap.Add "����� ����", Me.txtNationality_Spouse.Name
    messageMap.Add "������ �����ȣ", Me.txtLifeNo.Name
    messageMap.Add "�ڳ�1 �����ȣ", Me.txtLifeNo_Child1.Name
    messageMap.Add "�ڳ�2 �����ȣ", Me.txtLifeNo_Child2.Name
    messageMap.Add "�ڳ�3 �����ȣ", Me.txtLifeNo_Child3.Name
    messageMap.Add "����� �����ȣ", Me.txtLifeNo_Spouse.Name
    messageMap.Add "������ �ѱ��̸�", Me.txtName_ko.Name
    messageMap.Add "����� �ѱ��̸�", Me.txtName_Spouse_ko.Name
    messageMap.Add "�ڳ�1 �ѱ��̸�", Me.txtName_Child1_ko.Name
    messageMap.Add "�ڳ�2 �ѱ��̸�", Me.txtName_Child2_ko.Name
    messageMap.Add "�ڳ�3 �ѱ��̸�", Me.txtName_Child3_ko.Name
    messageMap.Add "������ �����̸�", Me.txtName_en.Name
    messageMap.Add "����� �����̸�", Me.txtName_Spouse_en.Name
    messageMap.Add "�ڳ�1 �����̸�", Me.txtName_Child1_en.Name
    messageMap.Add "�ڳ�2 �����̸�", Me.txtName_Child2_en.Name
    messageMap.Add "�ڳ�3 �����̸�", Me.txtName_Child3_en.Name
    
    
    '--//�����Ͱ� ��ȿ�ϴٴ� ���� �Ͽ� ����
    fnData_Validation = True
    
    Dim control As MSForms.control
    '--//control ��ȸ�ϸ鼭 ����
On Error GoTo INVALID_INPUT_DATA
    errorMessageAppeared = False
    For Each control In controls
        Select Case TypeName(control)
            Case "TextBox":
                Set txtBox_Focus = control
                '--//�ʼ����� �����Ͱ� �Էµ��� �ʾ��� ��
                If IsEmptyRequiredControl(control) Then
                    err.Raise vbObjectError + ERR_CODE_REQUIRED_INPUT_PINFORMATION, , _
                        StringFormat(ERR_DESC_REQUIRED_INPUT_PINFORMATION, messageMap.Item(control.Name))
                '--//������ ������ �߸� �Է� �Ǿ��� ��(�����ȣ, ��¥ ��)
                ElseIf IsInvalidInput(control) Then
                    err.Raise vbObjectError + ERR_CODE_INVALID_INPUT_PINFORMATION, , _
                        StringFormat(ERR_DESC_INVALID_INPUT_PINFORMATION, messageMap.Item(control.Name))
                End If
            Case "Frame":
                fnData_Validation control.controls
        End Select
    Next

SUCCEED:
    Exit Function

INVALID_INPUT_DATA:
    fnData_Validation = False
    If Not errorMessageAppeared Then
        MsgBox err.Description, vbCritical, banner
        errorMessageAppeared = True
    End If
End Function

Private Function GetRequiredMap() As Collection

    '--//�ʼ��� �Է´�� TextBox ����Ʈ ����
    Dim requiredMap As New Collection
    requiredMap.Add "������ �����ȣ", Me.txtLifeNo.Name
    requiredMap.Add "������ �ѱ��̸�", Me.txtName_ko.Name
    requiredMap.Add "������ �����̸�", Me.txtName_en.Name
    requiredMap.Add "������ �������", Me.txtBirthday.Name
    requiredMap.Add "������ ����", Me.txtNationality.Name
    requiredMap.Add "����� �����ȣ", Me.txtLifeNo_Spouse.Name
    requiredMap.Add "����� �ѱ��̸�", Me.txtName_Spouse_ko.Name
    requiredMap.Add "����� �����̸�", Me.txtName_Spouse_en.Name
    requiredMap.Add "����� �������", Me.txtBirthday_Spouse.Name
    requiredMap.Add "����� ����", Me.txtNationality_Spouse.Name
    requiredMap.Add "�ڳ�1 �����ȣ", Me.txtLifeNo_Child1.Name
    requiredMap.Add "�ڳ�1 �ѱ��̸�", Me.txtName_Child1_ko.Name
    requiredMap.Add "�ڳ�1 �����̸�", Me.txtName_Child1_en.Name
    requiredMap.Add "�ڳ�1 �������", Me.txtBirthday_Child1.Name
    requiredMap.Add "�ڳ�2 �����ȣ", Me.txtLifeNo_Child2.Name
    requiredMap.Add "�ڳ�2 �ѱ��̸�", Me.txtName_Child2_ko.Name
    requiredMap.Add "�ڳ�2 �����̸�", Me.txtName_Child2_en.Name
    requiredMap.Add "�ڳ�2 �������", Me.txtBirthday_Child2.Name
    requiredMap.Add "�ڳ�3 �����ȣ", Me.txtLifeNo_Child3.Name
    requiredMap.Add "�ڳ�3 �ѱ��̸�", Me.txtName_Child3_ko.Name
    requiredMap.Add "�ڳ�3 �����̸�", Me.txtName_Child3_en.Name
    requiredMap.Add "�ڳ�3 �������", Me.txtBirthday_Child3.Name
    
    Set GetRequiredMap = requiredMap

End Function

'--//�Էµ� �����Ͱ� ��ȿ���� ������ Ȯ��
'--//True: ��ȿ���� ����
'--//False: ��ȿ��
Private Function IsInvalidInput(control As MSForms.control) As Boolean
    
    '--//��ȿ�ϴٴ� ���� �Ͽ� ����
    IsInvalidInput = False
    
    '--//��Ʈ�� ���� ��������� ���ν��� ����
    If control.text = "" Then
        Exit Function
    End If
    
    '--//���� ������ ���� ��ü�غ�
    Dim CountryList As Object
    Dim CountryDao As New CountryDao
    Set CountryList = CountryDao.GetCountryList
    
    '--//��¥���� üũ
    If control.Name Like "*Birthday*" Or control.Name Like "*_dt*" Then
        If Not IsDate(control.text) Then
            IsInvalidInput = True
        End If
    End If
    
    '--//���� üũ
    If control.Name Like "*Nationality*" Then
        If Not CountryList.Contains(control.text) Then
            IsInvalidInput = True
        End If
    End If
    
    '--//�����ȣ ���� üũ
    If control.Name Like "*LifeNo*" Then
        If Not IsNumeric(fnExtract(control.text)) Or _
            Mid(control.text, 4, 1) <> "-" Or Mid(control.text, 11, 1) <> "-" Then
            IsInvalidInput = True
        End If
    End If
    
    '--//�̸� üũ
    If control.Name Like "*Name*" Then
        '--//�ѱ��̸� üũ
        If control.Name Like "*ko*" Then
            If Len(fnExtract(control.text, "E")) > 0 Then
                IsInvalidInput = True
            End If
        End If
        '--//�����̸� üũ
        If control.Name Like "*en*" Then
            If Len(fnExtract(control.text, "H")) > 0 Then
                IsInvalidInput = True
            End If
        End If
    End If
    
    '--//������� Ȯ��
    If control.Name = Me.txtTheological_Order.Name And _
        Not IsNumeric(Me.txtTheological_Order) And Me.txtTheological_Order <> "" Then
        IsInvalidInput = True
    End If
    
End Function

'--//�ʼ� �Է°��� ����ִ��� Ȯ��
'--//True: �������
'--//False: ������� ����
Private Function IsEmptyRequiredControl(control As MSForms.control) As Boolean

    '--//�ʼ��� ������ ���� ��ü�غ�
    Dim requiredMap As Collection
    Set requiredMap = GetRequiredMap

    IsEmptyRequiredControl = False
        
    '--//�ʼ��� �Է¿��� üũ
    If ExistsInCollection(requiredMap, control.Name) And control = "" Then
        If control.Name Like "*Spouse*" And Me.chkSpouse.Value = False Then
            Exit Function
        End If
        If control.Name Like "*Child1*" And Me.chkChild1.Value = False Then
            Exit Function
        End If
        If control.Name Like "*Child2*" And Me.chkChild2.Value = False Then
            Exit Function
        End If
        If control.Name Like "*Child3*" And Me.chkChild3.Value = False Then
            Exit Function
        End If
        
        IsEmptyRequiredControl = True
    End If

End Function

'@param frame: Frame ���� Control�鿡 ���Ͽ� Activate/DeActivate �մϴ�.
'@param blnActivate: True: Activate | False: Deactivate
Private Sub ActivateFrame(ByRef frame As MSForms.frame, blnActivate As Boolean)
    
    Dim control As MSForms.control
    Dim requiredList As Object
    
    '--//�ʼ��� ����Ʈ �ҷ�����
    Set requiredList = GetRequiredList
    
    For Each control In frame.controls
        Select Case TypeName(control)
            Case "TextBox":
                If InStr(control.Name, "LifeNo") > 0 And blnActivate Then
                    TextBoxEnable control, Not blnActivate
                Else
                    TextBoxEnable control, blnActivate
                End If
            Case "Label":
                If Not control.Name Like "*Info*" Then
                    If blnActivate Then
                        ChangeLabelColor control, vbRed
                    Else
                        ChangeLabelColor control, vbBlack
                    End If
                End If
                If control.Name Like "*Pic*" Then
                    If Not blnActivate Then
                        control.Picture = LoadPicture("")
'                        InsertPicToLabel control, ""
                    End If
                End If
            Case "Frame":
                '--//�̰� Ȱ��ȭ ��Ű�� ���ѷ��� ����
'                ActivateInputModeForFrame control, blnActivate
        End Select
    Next
    
End Sub

'@param controls: InputMode�� Ȱ��ȭ�� controls ��ü
'@param blnActivate: Ȱ��ȭ ���� ����
Private Sub ActivateInputMode(ByRef controls As MSForms.controls, blnActivate As Boolean, INPUTMODE)

    Dim control As MSForms.control
    Dim requiredList As Object
    
    If INPUTMODE = INPUT_FOR_ALL Then
        '--//�ű��Է� �ÿ��� �˻����� ��� ��Ȱ��ȭ
        LockControlsOfSearchFunction blnActivate
        
        '--//InputMode�� ���� Command ��ư����
        HideCmdBtnForInput blnActivate
        
        ExtraControlsEnable blnActivate
    End If
    
    '--//�ؽ�Ʈ�ڽ� �ʱ�ȭ
    InitializeTextBoxes controls
    
    
    '--//�ʼ��� ����Ʈ �ҷ�����
    Set requiredList = GetRequiredList
    
    For Each control In controls
        Select Case TypeName(control)
            Case "TextBox":
                If Not control.Name Like "*Search*" Then
                    TextBoxEnable control, blnActivate
                End If
            Case "Label":
                '--//Info���� �����̹Ƿ� ����
                If Not control.Name Like "*Info*" Then
                    If blnActivate And _
                        requiredList.Contains(control.Caption) Then
                        ChangeLabelColor control, vbRed
                    Else
                        ChangeLabelColor control, vbBlack
                    End If
                End If
            Case "Frame":
                If Not control.Name Like "*Transfer*" Then
                    ActivateInputMode control.controls, blnActivate, INPUT_FOR_PART
                End If
            Case "CheckBox":
                If (control.Name Like "*Child*" Or control.Name Like "*Spouse*") And _
                    Not control.Name Like "*Transfer*" Then
                    control.Visible = blnActivate
                End If
            Case "CommandButton":
                If control.Name Like "*Delete*" And _
                    (control.Name Like "*Child*" Or control.Name Like "*Spouse*") Then
                    control.Enabled = Not blnActivate
                End If
        End Select
    Next
    
    '--//������ �ʺ�����(�ڳడ ���̵���)
    cmdExtendChild_Click
    
    '--//üũ�ڽ� üũ���ο� ���� ��������
    '--//���� True�� ���� �������� ���� ����
    '--//���� True�� �ͱ��� �ϸ� ���ѷ���
'    If Not chkSpouse.Value Then chkSpouse_Click
'    If Not chkChild1.Value Then chkChild1_Click
'    If Not chkChild2.Value Then chkChild2_Click
'    If Not chkChild3.Value Then chkChild3_Click
    
    '--//���� �ʱ�ȭ
    For Each control In controls
        If TypeName(control) = "Label" Then
            If IsNumeric(InStr(control, "Pic")) Then
                control.Picture = LoadPicture("")
            End If
        End If
    Next
    
    '--//InputMode False �� �� �� �ʱ�ȭ
    If Not blnActivate Then
        lstPStaff_Click
    End If

End Sub

'--//�ű� �Է� �� �˻����� ��Ʈ���� ��Ȱ��ȭ �Ͽ� �ǵ�ġ ���� �̺�Ʈ�� �߻����� �ʵ��� �����մϴ�.
'@param blnActivate: Ȱ��ȭ ���� ����
Sub LockControlsOfSearchFunction(blnActivate As Boolean)
    
    '--//�˻��� ���õ� ��Ʈ�� ��Ȱ��ȭ
    Me.txtSearchText.Enabled = Not blnActivate
    Me.lstPStaff.Enabled = Not blnActivate
    Me.chkAll.Enabled = Not blnActivate
    Me.cmdSearch.Enabled = Not blnActivate

End Sub

'--//�ű� �Է� �� Input Control�� ������ �ʱ�ȭ �մϴ�.
'@param controls: TextBox�� �ʱ�ȭ �ϱ� ���ؼ� controls ��ü�� �޽��ϴ�.
Sub InitializeTextBoxes(ByRef controls As MSForms.controls)
    
    Dim control As MSForms.control
    
    For Each control In controls
        Select Case TypeName(control)
        Case "TextBox", "ComboBox":
            If Not control.Name Like "*Search*" Then
                control = ""
            End If
        Case "CommandButton":
            
        Case "CheckBox":
            '--//�����, �ڳ�1,2,3 üũ�ڽ��� Click �޼��带 ���� ���߰�
            '--//Transfer�� �ʱ�ȭ
            '--//cmdCancel_Click �ÿ��� lstPstaff_Click�� ���� ����
            If control.Name Like "*Transfer*" Then
                control.Value = False
            End If
        Case "Label":
            If control.Name Like "*Pic*" Then
                control.Picture = LoadPicture("")
            End If
        Case "Frame":
            InitializeTextBoxes control.controls
        End Select
    Next
    
End Sub

'--//DELETE_ITEM ������ �ִ� ����ڸ� ������ư�� ���̰� �մϴ�.
Private Sub HideDeleteButtonByUserAuth()
    
    Dim authList As Object
    Dim objUserDao As New UserDao
    
    Set authList = objUserDao.GetUserAuthorities
    
    Dim strAuth As Variant
    If authList.Contains("DELETE_ITEM") Then
        Me.cmdDelete.Visible = True
        Me.cmdDelete_Child1.Visible = True
        Me.cmdDelete_Child2.Visible = True
        Me.cmdDelete_Child3.Visible = True
        Me.cmdDelete_Spouse.Visible = True
    Else
        Me.cmdDelete.Visible = False
        Me.cmdDelete_Child1.Visible = False
        Me.cmdDelete_Child2.Visible = False
        Me.cmdDelete_Child3.Visible = False
        Me.cmdDelete_Spouse.Visible = False
    End If

End Sub

'--//�ű� �Է� �� ������ �� ��ư�� ������ �ʾƾ� �� ��ư�� �����մϴ�.
Private Sub HideCmdBtnForInput(argBoolean As Boolean)
    Me.cmdNew.Visible = Not argBoolean
    Me.cmdEdit.Visible = Not argBoolean
    Me.cmdDelete.Visible = Not argBoolean
    Me.cmdCancel.Visible = argBoolean
    Me.cmdAdd.Visible = argBoolean
    
    '--//������ư�� ���ѿ� ���� �ٽü���
    Call HideDeleteButtonByUserAuth
End Sub

'--//�ΰ����� ����� �ϴ� Control�鿡 ���� Ȱ��ȭ/��Ȱ��ȭ ���� ����
'@param blnActivate: Ȱ��ȭ/��Ȱ��ȭ ����
Private Sub ExtraControlsEnable(ByVal blnActivate As Boolean)

    '--//Add ��ư�� Ȱ��ȭ ���ο� ����
    Me.cmdAdd.Enabled = blnActivate
    If Me.lstPStaff.listIndex = -1 Then
        Me.cmdDelete.Enabled = Not blnActivate
        Me.cmdEdit.Enabled = Not blnActivate
    Else
        Me.cmdDelete.Enabled = blnActivate
        Me.cmdEdit.Enabled = blnActivate
    End If
    
    Dim control As MSForms.control
    
    '--//�� �� ��ü ��Ʈ�ε鿡 ���Ͽ�
    For Each control In Me.controls
        Select Case TypeName(control)
            Case "CheckBox":
                If Not (control.Name Like "*chkAll*" Or control.Name Like "*Transfer*") Then
                    ControlEnable control, blnActivate
                End If
            Case "TextBox":
                If control.Name <> "txtSearchText" And _
                    Not control.Name Like "*txtFamily*" Then
                    TextBoxEnable control, blnActivate
                End If
            Case "ComboBox":
                ComboBoxEnable control, blnActivate
            Case "Label":
                If control.Name Like "*Pic*" Then
                    ControlEnable control, blnActivate
                End If
            Case "CommandButton":
                If control.Caption Like "*������*" Then
                    ControlEnable control, Not blnActivate
                End If
                If control.Name Like "*Transfer*" Then
                    ControlEnable control, Not blnActivate
                End If
        End Select
    Next

End Sub

'--//Deprecated(2022-11-07)
'--//������ DB�� �����ϴ� ������ ������ ����Ǿ����Ƿ� �� �̻� ������� �ʽ��ϴ�.
Sub sbCopyPic(ByVal control As MSForms.control)

    Dim filePath As String
    Dim FileName As Variant
    Dim Class As String '--//��󱸺�
    Dim strExtension As String
    
    '--//��󱸺п� ���� ��������
    Select Case Replace(control.Name, "txtLifeNo_", "")
        Case "Spouse"
            Class = "�����"
        Case "Child1"
            Class = "�ڳ�1"
        Case "Child2"
            Class = "�ڳ�2"
        Case "Child3"
            Class = "�ڳ�3"
    Case Else
        Class = "������"
    End Select
    
    '--//��ȿ�� üũ
    If fnFindPicPath = False Then
        MsgBox "���̵�ũ�� ���� ������ �ּ���.", vbCritical, banner
        Exit Sub
    End If
    If control = "" Then
        MsgBox Class & " �����ȣ�� ���� �Է��� �ּ���.", vbCritical, banner
        Exit Sub
    End If
    
    '--//�������ϰ�� �ҷ�����
    filePath = fnFindPicPath
    
    '--//��������
    If FileName = False Then
        FileName = Application.GetOpenFilename(fileFilter:="Picturefile, *.jpg;*.jpeg; *.bmp; *.tif; *.png")
    End If
    
    '--//���� �̼��� �� ���ν��� ����
    If FileName = False Then
        Exit Sub
    End If
    
    '--//������ ���� Ȯ���� jpg,png�� �ƴ� �� ���ν��� ����
    If IsInArray(Right(FileName, Len(FileName) - InStrRev(FileName, ".")), Array("jpg"), , rtnSequence) = -1 Then
        MsgBox "������ JPG Ȯ���ڸ� ������ �� �ֽ��ϴ�."
        Exit Sub
    End If
    
    strExtension = "." & Right(FileName, Len(FileName) - InStrRev(FileName, "."))
    
    '--//������ ���� ���� �����ȣ�� �ٲ� ���� �������� ��η� �̵�����
    If Dir(filePath & control & strExtension) <> "" Then
        If MsgBox("������ �̹� ���� �մϴ�. ���� �Ͻðڽ��ϱ�?", vbQuestion + vbYesNo, banner) = vbYes Then
            Kill filePath & control & strExtension
            Name FileName As filePath & control & strExtension
        End If
    Else
        Name FileName As filePath & control & strExtension
    End If
    
    Call lstPStaff_Click

End Sub

'--//PStaff, PWife ��ü �Ľ��� ���ؼ� �������� �Ǿ� �ִ� ��Ʈ�ѿ� �⺻���� ä�� �ֽ��ϴ�.
Public Sub FillWithDefaultValue()
    If Me.txtBirthday = "" Then Me.txtBirthday = DateSerial(1900, 1, 1)
    If Me.txtBirthday_Spouse = "" Then Me.txtBirthday_Spouse = DateSerial(1900, 1, 1)
    If Me.txtBirthday_Child1 = "" Then Me.txtBirthday_Child1 = DateSerial(1900, 1, 1)
    If Me.txtBirthday_Child2 = "" Then Me.txtBirthday_Child2 = DateSerial(1900, 1, 1)
    If Me.txtBirthday_Child3 = "" Then Me.txtBirthday_Child3 = DateSerial(1900, 1, 1)
    If Me.txtOvs_dt = "" Then Me.txtOvs_dt = DateSerial(1900, 1, 1)
    If Me.txtWedding_dt = "" Then Me.txtWedding_dt = DateSerial(1900, 1, 1)
    If Me.txtOrdinationPrayer_dt = "" Then Me.txtOrdinationPrayer_dt = DateSerial(1900, 1, 1)
    
    If Me.txtTheological_Order = "" Then Me.txtTheological_Order = 0
    If Me.txtSalary = "" Then Me.txtSalary = 0
    
End Sub

'--//pStaff, pWife ��ü�� �޾� ���� ������ ä�� �ֽ��ϴ�.
Private Sub FillOutForm(ByRef pStaff As PastoralStaff, ByRef pWife As PastoralWife)

    '--//������
    Me.txtLifeNo = pStaff.lifeNo
    Me.txtName_ko = pStaff.nameKo
    Me.txtName_en = pStaff.NameEn
    Me.txtBirthday = pStaff.Birthday
    Me.txtPhone = pStaff.Phone
    Me.txtNationality = pStaff.Nationality
    Me.txtEducation = pStaff.Education
    Me.txtHome = pStaff.Home
    Me.txtFamily = pStaff.Family
    Me.txtHealth = pStaff.Health
    Me.txtOther = pStaff.Other
    Me.txtOvs_dt = IIf(pStaff.AppoOvs = DateSerial(1900, 1, 1), "", pStaff.AppoOvs)
    Me.cboBaptism = pStaff.Baptism
    Me.txtOrdinationPrayer_dt = IIf(pStaff.OrdinationPrayer = DateSerial(1900, 1, 1), "", pStaff.OrdinationPrayer)
    Me.txtWedding_dt = IIf(pStaff.WeddingDt = DateSerial(1900, 1, 1), "", pStaff.WeddingDt)
    Me.txtSalary = pStaff.Salary
    Me.txtTheological_Order = IIf(pStaff.TheologicalOrder = 0, "", pStaff.TheologicalOrder)
    Me.chkTransfer.Value = pStaff.Suspend
    
    '--//�ڳ�1
    Me.txtLifeNo_Child1 = pStaff.LifeNoChild1
    Me.txtName_Child1_ko = pStaff.NameKoChild1
    Me.txtName_Child1_en = pStaff.NameEnChild1
    Me.txtBirthday_Child1 = IIf(pStaff.BirthdayChild1 = DateSerial(1900, 1, 1), "", pStaff.BirthdayChild1)
    Me.txtPhone_Child1 = pStaff.PhoneChild1
    If pStaff.LifeNoChild1 <> "" Then
        Me.chkChild1.Value = True
    Else
        Me.chkChild1.Value = False
    End If
    
    '--//�ڳ�2
    Me.txtLifeNo_Child2 = pStaff.LifeNoChild2
    Me.txtName_Child2_ko = pStaff.NameKoChild2
    Me.txtName_Child2_en = pStaff.NameEnChild2
    Me.txtBirthday_Child2 = IIf(pStaff.BirthdayChild2 = DateSerial(1900, 1, 1), "", pStaff.BirthdayChild2)
    Me.txtPhone_Child2 = pStaff.PhoneChild2
    If pStaff.LifeNoChild2 <> "" Then
        Me.chkChild2.Value = True
    Else
        Me.chkChild2.Value = False
    End If
    
    '--//�ڳ�3
    Me.txtLifeNo_Child3 = pStaff.LifeNoChild3
    Me.txtName_Child3_ko = pStaff.NameKoChild3
    Me.txtName_Child3_en = pStaff.NameEnChild3
    Me.txtBirthday_Child3 = IIf(pStaff.BirthdayChild3 = DateSerial(1900, 1, 1), "", pStaff.BirthdayChild3)
    Me.txtPhone_Child3 = pStaff.PhoneChild3
    If pStaff.LifeNoChild3 <> "" Then
        Me.chkChild3.Value = True
    Else
        Me.chkChild3.Value = False
    End If
    
    '--//�����
    Me.txtLifeNo_Spouse = pWife.lifeNo
    Me.txtName_Spouse_ko = pWife.nameKo
    Me.txtName_Spouse_en = pWife.NameEn
    Me.txtBirthday_Spouse = IIf(pWife.Birthday = DateSerial(1900, 1, 1), "", pWife.Birthday)
    Me.txtPhone_Spouse = pWife.Phone
    Me.txtNationality_Spouse = pWife.Nationality
    Me.txtEducation_Spouse = pWife.Education
    Me.txtHome_Spouse = pWife.Home
    Me.txtFamily_Spouse = pWife.Family
    Me.txtHealth_Spouse = pWife.Health
    Me.txtOther_Spouse = pWife.Other
    If pWife.lifeNo <> "" Then
        Me.chkSpouse.Value = True
    Else
        Me.chkSpouse.Value = False
    End If
    Me.chkTransferSpouse.Value = pWife.Suspend

End Sub
