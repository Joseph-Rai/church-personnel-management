VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Search_Phone 
   Caption         =   "����ó ���� ������"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7275
   OleObjectBlob   =   "frm_Search_Phone.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "frm_Search_Phone"
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

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdCopy_All_Click()
    
    Dim txtCopy As String
    
    If Not (Me.txtLandLine = "" And Me.txtWMCPhone = "" And Me.txtPhone_PStaff = "" And Me.txtPhone_Spouse = "") Then
        With Me.lstPStaff
            txtCopy = IIf(.List(.listIndex, 2) <> .List(.listIndex, 3), "��ȸ��: " & .List(.listIndex, 3) & vbNewLine & "����ȸ��: " & Me.lblChurch.Caption, "��ȸ��: " & Me.lblChurch.Caption)
            If Me.txtLandLine <> "" Then
                txtCopy = txtCopy & vbNewLine & vbNewLine & Me.lblLandLine.Caption & vbNewLine & Me.txtLandLine
            End If
            If Me.txtWMCPhone <> "" Then
                txtCopy = txtCopy & vbNewLine & vbNewLine & Me.lblWMCPhone.Caption & vbNewLine & Me.txtWMCPhone
            End If
            If Me.txtPhone_PStaff <> "" Then
                txtCopy = txtCopy & vbNewLine & vbNewLine & Me.txtName & "/" & Me.txtPosition & vbNewLine & Me.txtPhone_PStaff
            End If
            If Me.txtPhone_Spouse <> "" Then
                txtCopy = txtCopy & vbNewLine & vbNewLine & Me.txtName_Spouse & "/" & Me.txtPosition_Spouse & vbNewLine & Me.txtPhone_Spouse
            End If
            If Me.txtAddress <> "" Then
                txtCopy = txtCopy & vbNewLine & vbNewLine & Me.txtAddress
            End If
            
            CopyText (Trim(txtCopy))
        End With
    Else
        MsgBox "������ ������ �����ϴ�.", vbInformation
    End If
End Sub

Private Sub cmdCopy_LandLine_Click()
    If Me.txtLandLine <> "" Then
        With Me.lstPStaff
            CopyText (IIf(.List(.listIndex, 2) <> .List(.listIndex, 3), "��ȸ��: " & .List(.listIndex, 3) & vbNewLine & "����ȸ��: " & Me.lblChurch.Caption, "��ȸ��: " & Me.lblChurch.Caption) & vbNewLine & vbNewLine & Me.lblLandLine.Caption & vbNewLine & Me.txtLandLine)
        End With
    Else
        MsgBox "������ ������ �����ϴ�.", vbInformation
    End If
End Sub

Private Sub cmdCopy_pstaff_Click()
    If Me.txtPhone_PStaff <> "" Then
        With Me.lstPStaff
            CopyText (IIf(.List(.listIndex, 2) <> .List(.listIndex, 3), "��ȸ��: " & .List(.listIndex, 3) & vbNewLine & "����ȸ��: " & Me.lblChurch.Caption, "��ȸ��: " & Me.lblChurch.Caption) & vbNewLine & vbNewLine & Me.txtName & "/" & Me.txtPosition & vbNewLine & Me.txtPhone_PStaff)
        End With
    Else
        MsgBox "������ ������ �����ϴ�.", vbInformation
    End If
End Sub

Private Sub cmdCopy_Spouse_Click()
    If Me.txtPhone_Spouse <> "" Then
        With Me.lstPStaff
            CopyText (IIf(.List(.listIndex, 2) <> .List(.listIndex, 3), "��ȸ��: " & .List(.listIndex, 3) & vbNewLine & "����ȸ��: " & Me.lblChurch.Caption, "��ȸ��: " & Me.lblChurch.Caption) & vbNewLine & vbNewLine & Me.txtName_Spouse & "/" & Me.txtPosition_Spouse & vbNewLine & Me.txtPhone_Spouse)
        End With
    Else
        MsgBox "������ ������ �����ϴ�.", vbInformation
    End If
End Sub

Private Sub cmdCopy_WMCPhone_Click()
    If Me.txtWMCPhone <> "" Then
        With Me.lstPStaff
            CopyText (IIf(.List(.listIndex, 2) <> .List(.listIndex, 3), "��ȸ��: " & .List(.listIndex, 3) & vbNewLine & "����ȸ��: " & Me.lblChurch.Caption, "��ȸ��: " & Me.lblChurch.Caption) & vbNewLine & vbNewLine & Me.lblWMCPhone.Caption & vbNewLine & Me.txtWMCPhone)
        End With
    Else
        MsgBox "������ ������ �����ϴ�.", vbInformation
    End If
End Sub

Private Sub cmdEdit_Click()
    Dim result As T_RESULT
    Dim lifeNo As String
    
    '--//������ ���� �ִ��� üũ
    strSql = makeSelectSQL(TB2)
    Call makeListData(strSql, TB2)
    If Me.txtLandLine = LISTDATA(0, 4) And Me.txtWMCPhone = LISTDATA(0, 5) And Me.txtPhone_PStaff = LISTDATA(0, 6) And Me.txtPhone_Spouse = LISTDATA(0, 7) And Me.txtAddress = LISTDATA(0, 18) Then
        Exit Sub
    End If
    
    '--//��ȸ����ó ������Ʈ SQL�� ����, ����, �αױ��
    strSql = makeUpdateSQL(TB3)
    result.strSql = strSql
    connectTaskDB
    result.affectedCount = executeSQL("cmdEdit_Click", TB3, strSql, Me.Name, "��ȸ����ó ������Ʈ")
    writeLog "cmdEdit_Click", TB3, strSql, 0, Me.Name, "��ȸ����ó ������Ʈ", result.affectedCount
    disconnectALL
    
    '--//�����ڿ���ó ������Ʈ SQL�� ����, ����, �αױ��
    lifeNo = Me.lstPStaff.List(Me.lstPStaff.listIndex)
    
    If Mid(lifeNo, InStr(5, lifeNo, "-") + 1, 1) = 1 Then
        strSql = makeUpdateSQL(TB4)
        result.strSql = strSql
        connectTaskDB
        result.affectedCount = executeSQL("cmdEdit_Click", TB4, strSql, Me.Name, "�����ڿ���ó ������Ʈ")
        writeLog "cmdEdit_Click", TB4, strSql, 0, Me.Name, "�����ڿ���ó ������Ʈ", result.affectedCount
        disconnectALL
    End If
    
    '--//��𿬶�ó ������Ʈ SQL�� ����, ����, �αױ��
    If Mid(lifeNo, InStr(5, lifeNo, "-") + 1, 1) = 2 Then
        strSql = makeUpdateSQL(TB5)
        result.strSql = strSql
        connectTaskDB
        result.affectedCount = executeSQL("cmdEdit_Click", TB5, strSql, Me.Name, "��𿬶�ó ������Ʈ")
        writeLog "cmdEdit_Click", TB5, strSql, 0, Me.Name, "��𿬶�ó ������Ʈ", result.affectedCount
        disconnectALL
    End If
    
    '--//�޼����ڽ�
    MsgBox "������Ʈ �Ǿ����ϴ�.", , banner
    
    '--//����Ʈ�ڽ� �ʱ�ȭ
    Call lstPStaff_Click
End Sub

Private Sub cmdExport_Click()

    Dim targetWB As Workbook
    Dim i As Long
    Dim lngColumnIndex As Long
    Dim arg As String
    
    Me.lblExport.Visible = True
    Me.optKo.Visible = False
    Me.optEn.Visible = False
    Me.Repaint
    
    '--//������ �ҷ�����
    strSql = makeSelectSQL(TB6)
    Call makeListData(strSql, TB6)
    lngColumnIndex = IsInArray("�����μ�", LISTFIELD, , rtnSequence)
   
    '--//�� ��ũ�� ���� �� ������ �ٿ��ֱ�
    Set targetWB = Workbooks.Add
    Call Optimization
    With targetWB.Sheets(1)
        .Cells(3, "A").Resize(, UBound(LISTFIELD) + 1) = LISTFIELD
        .Cells(4, "A").Resize(cntRecord, UBound(LISTFIELD) + 1) = LISTDATA
        
    
    '--//1�� ������ ���� �� ����
        lngColumnIndex = getColumIndex("�ѱ��̸�(����)", "�����̸�")
        .Cells(2, lngColumnIndex + 1) = "�� �� ��"
        lngColumnIndex = getColumIndex("����ѱ��̸�(����)", "��𿵹��̸�")
        .Cells(2, lngColumnIndex + 1) = "�� ��"
    
    '--//��������
        '--//1~2�� ����ó��
        lngColumnIndex = IsInArray("��å", LISTFIELD, , rtnSequence)
        For i = 1 To lngColumnIndex + 1
            .Cells(2, i).Resize(2).Merge
        Next
        lngColumnIndex = getColumIndex("�ѱ��̸�(����)", "�����̸�")
        .Cells(2, lngColumnIndex + 1).Resize(, 2).Merge
        lngColumnIndex = getColumIndex("����ѱ��̸�(����)", "��𿵹��̸�")
        .Cells(2, lngColumnIndex + 1).Resize(, 2).Merge
        
        '--//�����μ� ���߱�
        lngColumnIndex = IsInArray("�����μ�", LISTFIELD, , rtnSequence)
        .Cells(1, "A").Offset(, lngColumnIndex).Resize(, (UBound(LISTFIELD) + 1) - (lngColumnIndex)).EntireColumn.Group
        
        '--//�������
        .Columns("A:A").Resize(, UBound(LISTFIELD) + 1).EntireColumn.HorizontalAlignment = xlCenter
        .Columns("A:A").Resize(, UBound(LISTFIELD) + 1).EntireColumn.VerticalAlignment = xlCenter
        
        '--//�ʵ� �۲� �� ���� ����
        .Cells(2, "A").Resize(2, UBound(LISTFIELD) + 1).Interior.ThemeColor = xlThemeColorDark2
        .Cells(2, "A").Resize(2, UBound(LISTFIELD) + 1).Font.Bold = True
        
        '--//�׵θ�
        '--��ü
        lngColumnIndex = IsInArray("�������ȭ��ȣ", LISTFIELD, , rtnSequence)
        .Cells(2, "A").Resize(cntRecord + 2, lngColumnIndex + 1).Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Cells(2, "A").Resize(cntRecord + 2, lngColumnIndex + 1).Borders(xlEdgeTop).LineStyle = xlContinuous
        .Cells(2, "A").Resize(cntRecord + 2, lngColumnIndex + 1).Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Cells(2, "A").Resize(cntRecord + 2, lngColumnIndex + 1).Borders(xlEdgeRight).LineStyle = xlContinuous
        .Cells(2, "A").Resize(cntRecord + 2, lngColumnIndex + 1).Borders(xlInsideVertical).LineStyle = xlContinuous
        .Cells(2, "A").Resize(cntRecord + 2, lngColumnIndex + 1).Borders(xlInsideHorizontal).LineStyle = xlContinuous
        
        .Cells(2, "A").Resize(cntRecord + 2, lngColumnIndex + 1).Borders(xlEdgeLeft).Weight = xlHairline
        .Cells(2, "A").Resize(cntRecord + 2, lngColumnIndex + 1).Borders(xlEdgeTop).Weight = xlHairline
        .Cells(2, "A").Resize(cntRecord + 2, lngColumnIndex + 1).Borders(xlEdgeBottom).Weight = xlHairline
        .Cells(2, "A").Resize(cntRecord + 2, lngColumnIndex + 1).Borders(xlEdgeRight).Weight = xlHairline
        .Cells(2, "A").Resize(cntRecord + 2, lngColumnIndex + 1).Borders(xlInsideVertical).Weight = xlHairline
        .Cells(2, "A").Resize(cntRecord + 2, lngColumnIndex + 1).Borders(xlInsideHorizontal).Weight = xlHairline
        
        .Cells(2, "A").Resize(cntRecord + 2, lngColumnIndex + 1).Borders(xlEdgeLeft).Weight = xlMedium
        .Cells(2, "A").Resize(cntRecord + 2, lngColumnIndex + 1).Borders(xlEdgeTop).Weight = xlMedium
        .Cells(2, "A").Resize(cntRecord + 2, lngColumnIndex + 1).Borders(xlEdgeBottom).Weight = xlMedium
        .Cells(2, "A").Resize(cntRecord + 2, lngColumnIndex + 1).Borders(xlEdgeRight).Weight = xlMedium
        
        '--�߰� ���α��м�
        lngColumnIndex = getColumIndex("�ѱ��̸�(����)", "�����̸�")
        .Cells(2, lngColumnIndex + 1).Resize(cntRecord + 2, 3).Borders(xlEdgeLeft).Weight = xlMedium
        .Cells(2, lngColumnIndex + 1).Resize(cntRecord + 2, 3).Borders(xlEdgeRight).Weight = xlMedium
        .Cells(2, lngColumnIndex + 1).Resize(cntRecord + 2, 3).Borders(xlInsideVertical).Weight = xlHairline
        .Cells(2, lngColumnIndex + 1).Resize(cntRecord + 2, 3).Borders(xlInsideHorizontal).Weight = xlHairline
        
        '--�ʵ�
        lngColumnIndex = IsInArray("�������ȭ��ȣ", LISTFIELD, , rtnSequence)
        .Cells(2, "A").Resize(2, lngColumnIndex + 1).Borders(xlEdgeBottom).Weight = xlMedium
        
        '--//��ȭ��ȣ �÷� ����
        lngColumnIndex = IsInArray("���ͳ���ȭ", LISTFIELD, , rtnSequence)
        .Cells(4, lngColumnIndex + 1).Resize(cntRecord, 2).Interior.ThemeColor = xlThemeColorAccent3
        .Cells(4, lngColumnIndex + 1).Resize(cntRecord, 2).Interior.TintAndShade = 0.799981688894314
        .Cells(1, lngColumnIndex + 1).Resize(, 2).EntireColumn.ColumnWidth = 22
        
        lngColumnIndex = IsInArray("��������ȭ��ȣ", LISTFIELD, , rtnSequence)
        .Cells(4, lngColumnIndex + 1).Resize(cntRecord).Interior.ThemeColor = xlThemeColorAccent4
        .Cells(4, lngColumnIndex + 1).Resize(cntRecord).Interior.TintAndShade = 0.799981688894314
        .Cells(1, lngColumnIndex + 1).EntireColumn.ColumnWidth = 22
        
        lngColumnIndex = IsInArray("�������ȭ��ȣ", LISTFIELD, , rtnSequence)
        .Cells(4, lngColumnIndex + 1).Resize(cntRecord).Interior.ThemeColor = xlThemeColorAccent2
        .Cells(4, lngColumnIndex + 1).Resize(cntRecord).Interior.TintAndShade = 0.799981688894314
        .Cells(1, lngColumnIndex + 1).EntireColumn.ColumnWidth = 22
        
        '--//���ʺ�, ����� ����
        Columns("A:A").Resize(, UBound(LISTFIELD) + 1).EntireColumn.AutoFit
        Columns("A:A").Resize(cntRecord + 2).EntireRow.AutoFit
        'Columns("A:A").Resize(cntRecord + 2).RowHeight = 24
        
        Call Normal
        Call Optimization
        
        '--//����ȸ �۲� �� ��� ��Ÿ��
        Dim temp As Long
        Dim temp2 As Long
        lngColumnIndex = IsInArray("����ȸ�ڵ�", LISTFIELD, , rtnSequence)
        temp = IsInArray("���ͳ���ȭ", LISTFIELD, , rtnSequence)
        temp2 = IsInArray("����ȸ�ڵ�", LISTFIELD, , rtnSequence)
        For i = 4 To 3 + cntRecord
            If Cells(i, lngColumnIndex + 1) <> Cells(i - 1, lngColumnIndex + 1) Then
                Cells(i, "A").EntireRow.Font.Bold = True
                Cells(i, "A").Resize(, lngColumnIndex + 1).Interior.color = 13434879
            Else
                If InStr(Cells(i, temp2 + 1), "MC") > 0 Then
                    Cells(i, temp + 1).Resize(, 2).ClearContents
                End If
            End If
            If Cells(i, "A").EntireRow.RowHeight < 24 Then
                Cells(i, "A").EntireRow.RowHeight = 24
            End If
        Next
        
        '--//�μ��� �ۼ�
        Cells(1, "A").EntireRow.RowHeight = 25
        strSql = "SELECT dept_nm FROM db_ovs_dept WHERE dept_id = " & SText(USER_DEPT) & ";"
        Call makeListData(strSql, "op_system.db_ovs_dept")
        Cells(1, "C") = LISTDATA(0, 0) & " ����ó"
        Cells(1, "C").Font.Bold = True
        Cells(1, "C").Font.Size = 16
        Cells(1, "D") = Format(Now(), "yyyy-mm")
        Cells(1, "D").Font.Bold = True
        Cells(1, "D").Font.Size = 16
        Cells(1, "D").Font.ThemeColor = xlThemeColorAccent2
        
        ActiveSheet.Outline.ShowLevels RowLevels:=0, ColumnLevels:=1
        
        '--//����Ʈ���� ����
        With ActiveSheet.PageSetup
            .PrintTitleRows = "$2:$3"
            .PrintTitleColumns = ""
        End With
        ActiveSheet.PageSetup.PrintArea = ""
        With ActiveSheet.PageSetup
            .LeftMargin = Application.InchesToPoints(0.25)
            .RightMargin = Application.InchesToPoints(0.25)
            .TopMargin = Application.InchesToPoints(0.75)
            .BottomMargin = Application.InchesToPoints(0.75)
            .HeaderMargin = Application.InchesToPoints(0.3)
            .FooterMargin = Application.InchesToPoints(0.3)
            .Orientation = xlLandscape
            .CenterFooter = "&N������ �� &P������"
        End With
        ActiveWindow.View = xlPageBreakPreview
        ActiveSheet.VPageBreaks(1).DragOff Direction:=xlToRight, RegionIndex:=1
        
        Call sbClearVariant
    End With
    
    Call Normal
    Me.lblExport.Visible = False
    Me.optKo.Visible = True
    Me.optEn.Visible = True
    MsgBox "��¹��� �����Ǿ����ϴ�."
    
End Sub

Private Function getColumIndex(arg1 As String, arg2 As String)

    Dim lngColumnIndex As Long
    Dim lngColumnIndex_KO As Long
    Dim lngColumnIndex_EN As Long

    lngColumnIndex_KO = IsInArray(arg1, LISTFIELD, , rtnSequence)
    lngColumnIndex_EN = IsInArray(arg2, LISTFIELD, , rtnSequence)
    lngColumnIndex = WorksheetFunction.Max(lngColumnIndex_KO, lngColumnIndex_EN)
    
    getColumIndex = lngColumnIndex

End Function

Private Sub lstPStaff_Click()
    
    Dim filePath As String
    Dim FileName As String
    Dim lifeNo As String
    
    '--//�̷¸�� ������ ä���
    strSql = makeSelectSQL(TB2)
    Call makeListData(strSql, TB2)
    If cntRecord > 0 Then
        Me.txtLandLine = LISTDATA(0, 4)
        Me.txtWMCPhone = LISTDATA(0, 5)
        Me.txtPhone_PStaff = LISTDATA(0, 6)
        Me.txtPhone_Spouse = LISTDATA(0, 7)
        Me.lblChurch.Caption = LISTDATA(0, 3)
        Me.lblTime_different = Format(LISTDATA(0, 1), "����: hh��nn��")
        Me.txtName = LISTDATA(0, 9)
        Me.txtPosition = LISTDATA(0, 10)
        Me.txtName_Spouse = LISTDATA(0, 12)
        Me.txtPosition_Spouse = LISTDATA(0, 13)
        Me.txtAddress = LISTDATA(0, 18)
        
        '--//��ȸ�� �� ���� ����
        Me.lblChurch.Visible = True
        Me.lblTime_different.Visible = True
        
        '--//��ȸ�� ���� �ؽ�Ʈ�ڽ� ��Ȱ��ȭ
        If Me.lstPStaff.listIndex <> -1 Then
            If LISTDATA(0, 2) = "" Then '--//��ȸ�ڵ尡 ������
                Me.txtLandLine.Enabled = False
                Me.txtWMCPhone.Enabled = False
                Me.txtAddress.Enabled = False
            Else
                Me.txtLandLine.Enabled = True
                Me.txtWMCPhone.Enabled = True
                Me.txtAddress.Enabled = True
            End If
        Else
            Me.txtLandLine.Enabled = False
            Me.txtWMCPhone.Enabled = False
            Me.txtAddress.Enabled = False
        End If
    Else
        Call sbtxtBox_Init
    End If
    Call sbClearVariant
    
    '--//������ ���� �ؽ�Ʈ�ڽ� ��Ȱ��ȭ
    If Me.lstPStaff.listIndex <> -1 Then
        lifeNo = Me.lstPStaff.List(Me.lstPStaff.listIndex)
        
        If Mid(lifeNo, InStr(5, lifeNo, "-") + 1, 1) = 1 Then
            Me.txtPhone_PStaff.Enabled = True
            Me.txtPhone_Spouse.Enabled = False
        End If
        If Mid(lifeNo, InStr(5, lifeNo, "-") + 1, 1) = 2 Then
            Me.txtPhone_PStaff.Enabled = False
            Me.txtPhone_Spouse.Enabled = True
        End If
    Else
        Me.txtPhone_PStaff.Enabled = False
        Me.txtPhone_Spouse.Enabled = False
    End If
    
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
    If Me.lstPStaff.ListCount > 0 Then
        'HookListBoxScroll Me, Me.lstPStaff
    End If
End Sub

Private Sub txtChurchNM_Change()
    Me.txtChurchNM.BackColor = RGB(255, 255, 255)
End Sub

Private Sub UserForm_Initialize()
    If connIP = Empty Then setGlobalVariant '//�������� �缳��
    If checkLogin = 0 Then f_login.Show '//�α���üũ
    If checkLogin = 0 Then End '--//�α��� ���� �� ���ν��� ����
    
    '--//���ʼ���
    Me.cmdClose.Cancel = True
    TB1 = "op_system.v0_pstaff_information_all" '--//�����ڸ���Ʈ
    TB2 = "op_system.v_phone" '--//��ȭ��ȣ
    TB3 = "op_system.db_phone" '--//��ȸ����ó
    TB4 = "op_system.db_pastoralstaff" '--//����������
    TB5 = "op_system.db_pastoralwife" '--//���������
    TB6 = "op_system.v_phone_export" '--//��¹� �������
    
    '--//��Ʈ�� ����
    Me.lstPStaff.Enabled = False
    Me.lblChurch.Visible = False
    Me.lblTime_different.Visible = False
    Me.txtName.Visible = False
    Me.txtName_Spouse.Visible = False
    Me.txtPosition.Visible = False
    Me.txtPosition_Spouse.Visible = False
    Me.txtPhone_PStaff.Enabled = False
    Me.txtPhone_Spouse.Enabled = False
    Me.txtLandLine.Enabled = False
    Me.txtWMCPhone.Enabled = False
    Me.txtAddress.Enabled = False
    Me.lblExport.Visible = False
    Me.optKo.Value = True
    
    '--//����Ʈ�ڽ� ����
    With Me.lstPStaff
        .ColumnCount = 6
        .ColumnHeads = False
        .ColumnWidths = "0,0,120,70,50,50" '�����ȣ, ��ȸ�ڵ�, ��ȸ��, ������ȸ��, �ѱ��̸�(����), ��å
        .TextAlign = fmTextAlignLeft
        .Font = "����"
    End With
    
    '--//�μ��� ���� ������ȭ ī���ȣ ����
    strSql = "SELECT a.dept_phonecard FROM op_system.db_ovs_dept a WHERE a.dept_id = " & SText(USER_DEPT)
    Call makeListData(strSql, "op_system.db_ovs_dept")
    Me.lblCardNo.Caption = "�Ϲ���ȭ(���� 08216) : " & LISTDATA(0, 0)
    Call sbClearVariant
    
    Me.txtChurchNM.SetFocus
    
End Sub
Private Sub cmdSearch_Click()

    Call sbtxtBox_Init
    
    If Not Me.chkAll Then
        strSql = makeSelectSQL(TB1)
        Call makeListData(strSql, TB1)
    Else
        strSql = makeSelectSQL2(TB1)
        Call makeListData(strSql, TB1)
    End If
    
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
        strSql = "SELECT a.`�����ڻ����ȣ`,a.`��ȸ�ڵ�`,a.`��ȸ��`,a.`������ȸ��`,a.`�ѱ��̸�(����)`,a.`��å` " & _
                    "FROM " & TB2 & " a " & _
                    "WHERE a.`��å` is not null AND (a.`�ѱ��̸�(����)` LIKE '%" & Me.txtChurchNM & "%' OR a.`�����̸�` LIKE '%" & Me.txtChurchNM & "%' OR a.`��ȸ��` LIKE '%" & Me.txtChurchNM & "%'" & " OR a.`������ȸ��` LIKE '%" & Me.txtChurchNM & "%'" & " OR a.`�����ڻ����ȣ` LIKE '%" & Me.txtChurchNM & "%')" & _
                    " AND a.`�����μ�` = " & SText(USER_DEPT) & _
                 " Union " & _
                 "SELECT b.`����ڻ����ȣ`,b.`��ȸ�ڵ�`,b.`��ȸ��`,b.`������ȸ��`,b.`����ѱ��̸�(����)`,b.`�����å` " & _
                    "FROM " & TB2 & " b " & _
                    "WHERE b.`�����å` is not null AND (b.`����ѱ��̸�(����)` LIKE '%" & Me.txtChurchNM & "%' OR b.`��𿵹��̸�` LIKE '%" & Me.txtChurchNM & "%' OR b.`��ȸ��` LIKE '%" & Me.txtChurchNM & "%'" & " OR b.`������ȸ��` LIKE '%" & Me.txtChurchNM & "%'" & " OR b.`����ڻ����ȣ` LIKE '%" & Me.txtChurchNM & "%')" & _
                    " AND b.`�����μ�` = " & SText(USER_DEPT) & _
                    " ORDER BY `������ȸ��`,FIELD(`��å`,'��ȸ��','��ȸ��븮','����','��븮���','����','�����','����ȸ������','�����ڻ��','����Ұ�����','�����ڻ��','�������1�ܰ�','�������2�ܰ�','�������3�ܰ�','�������'," & getPosition2Joining & ",'');"
    Case TB2
        With Me.lstPStaff
            strSql = "SELECT * FROM " & TB2 & " a WHERE a.`�����ڻ����ȣ` = " & SText(.List(.listIndex)) & " OR a.`����ڻ����ȣ` = " & SText(.List(.listIndex)) & ";"
        End With
    Case TB6
        If Me.optKo.Value Then
            strSql = "SELECT" & _
                        " a.`��������`" & _
                        ",DATE_FORMAT(a.`����`,'%H:%i') '����'" & _
                        ",a.`��ȸ��`" & _
                        ",a.`���ͳ���ȭ`" & _
                        ",a.`������ȭ`" & _
                        ",a.`��å`" & _
                        ",a.`�ѱ��̸�(����)`" & _
                        ",a.`��������ȭ��ȣ`" & _
                        ",a.`����ѱ��̸�(����)`" & _
                        ",a.`�������ȭ��ȣ`" & _
                        ",a.`�����μ�`" & _
                        ",a.`����ȸ�ڵ�`" & _
                        ",a.`����ȸ�ڵ�`" & _
                    " FROM " & TB2 & " a" & _
                    " WHERE a.`��ȸ��` IS NOT NULL AND a.`��å` IS NOT NULL" & _
                        " AND a.`���ļ���` >= (SELECT sort_order FROM op_system.db_churchlist WHERE church_nm = '���� ������丣')" & _
                        " AND a.`�����μ�` = " & SText(USER_DEPT) & _
                    " ORDER BY a.`���ļ���`, a.`��å` IS NULL ASC, FIELD(`��å`,'��ȸ��','��ȸ��븮','����','��븮���','����','�����','����ȸ������','�����ڻ��','����Ұ�����','�����ڻ��','�������1�ܰ�','�������2�ܰ�','�������3�ܰ�','�������'," & getPosition2Joining & ",'')" & ";"
        Else
            strSql = "SELECT" & _
                        " a.`��������`" & _
                        ",DATE_FORMAT(a.`����`,'%H:%i') '����'" & _
                        ",a.`������ȸ��`" & _
                        ",a.`���ͳ���ȭ`" & _
                        ",a.`������ȭ`" & _
                        ",a.`��å`" & _
                        ",a.`�����̸�`" & _
                        ",a.`��������ȭ��ȣ`" & _
                        ",a.`��𿵹��̸�`" & _
                        ",a.`�������ȭ��ȣ`" & _
                        ",a.`�����μ�`" & _
                        ",a.`����ȸ�ڵ�`" & _
                        ",a.`����ȸ�ڵ�`" & _
                    " FROM " & TB2 & " a" & _
                    " WHERE a.`��ȸ��` IS NOT NULL AND a.`��å` IS NOT NULL" & _
                        " AND a.`���ļ���` >= (SELECT sort_order FROM op_system.db_churchlist WHERE church_nm = '���� ������丣')" & _
                        " AND a.`�����μ�` = " & SText(USER_DEPT) & _
                    " ORDER BY a.`���ļ���`, a.`��å` IS NULL ASC, FIELD(`��å`,'��ȸ��','��ȸ��븮','����','��븮���','����','�����','����ȸ������','�����ڻ��','����Ұ�����','�����ڻ��','�������1�ܰ�','�������2�ܰ�','�������3�ܰ�','�������'," & getPosition2Joining & ",'')" & ";"
        End If
    Case Else
    End Select
    makeSelectSQL = strSql
End Function
Private Function makeSelectSQL2(ByVal tableNM As String) As String
    
    Select Case tableNM
    Case TB1
        strSql = "SELECT a.`�����ȣ`,esta1.`church_sid_custom`,a.`����ȸ��`,a.`��ȸ��`,a.`�ѱ��̸�(����)`,a.`��å` " & _
                    "FROM " & TB1 & " a " & _
                    "LEFT JOIN op_system.db_history_church_establish esta1 ON IFNULL(a.`����ȸ�ڵ�`,a.`��ȸ�ڵ�`)=esta1.`church_sid` " & _
                    "WHERE (a.`�ѱ��̸�(����)` LIKE '%" & Me.txtChurchNM & "%' OR a.`�����̸�` LIKE '%" & Me.txtChurchNM & "%' OR a.`��ȸ��` LIKE '%" & Me.txtChurchNM & "%'" & " OR a.`����ȸ��` LIKE '%" & Me.txtChurchNM & "%'" & " OR a.`�����ȣ` LIKE '%" & Me.txtChurchNM & "%')" & _
                    " AND a.`�����μ�` = " & SText(USER_DEPT) & _
                 " Union " & _
                 "SELECT b.`����ڻ���`,esta2.`church_sid_custom`,b.`����ȸ��`,b.`��ȸ��`,b.`����ѱ��̸�(����)`,b.`�����å` " & _
                    "FROM " & TB1 & " b " & _
                    "LEFT JOIN op_system.db_history_church_establish esta2 ON IFNULL(b.`����ȸ�ڵ�`,b.`��ȸ�ڵ�`)=esta2.`church_sid` " & _
                    "WHERE (b.`����ѱ��̸�(����)` LIKE '%" & Me.txtChurchNM & "%' OR b.`��𿵹��̸�` LIKE '%" & Me.txtChurchNM & "%' OR b.`��ȸ��` LIKE '%" & Me.txtChurchNM & "%'" & " OR b.`����ȸ��` LIKE '%" & Me.txtChurchNM & "%'" & " OR b.`����ڻ���` LIKE '%" & Me.txtChurchNM & "%')" & _
                    " AND b.`�����μ�` = " & SText(USER_DEPT) & " AND b.`����ڻ���` IS NOT NULL " & _
                    " ORDER BY `��ȸ��`,FIELD(`��å`,'��ȸ��','��ȸ��븮','����','��븮���','����','�����','����ȸ������','�����ڻ��','����Ұ�����','�����ڻ��','�������1�ܰ�','�������2�ܰ�','�������3�ܰ�','�������'," & getPosition2Joining & ",'');"
    Case TB2
    Case Else
    End Select
    makeSelectSQL2 = strSql
End Function
Private Function makeUpdateSQL(ByVal tableNM As String) As String
    
    Select Case tableNM
    Case TB3
        strSql = "SELECT a.church_sid FROM " & TB3 & " a " & _
                " WHERE a.church_sid = " & SText(Me.lstPStaff.List(Me.lstPStaff.listIndex, 1)) & ";"
        Call makeListData(strSql, TB3)
        
        If cntRecord > 0 Then
            strSql = "UPDATE " & TB3 & " a " & _
                    "SET a.phone = " & SText(Me.txtLandLine) & ", a.wmcphone = " & SText(Me.txtWMCPhone) & ", a.address = " & SText(Me.txtAddress) & _
                    " WHERE a.church_sid = " & SText(Me.lstPStaff.List(Me.lstPStaff.listIndex, 1)) & ";"
        Else
            strSql = "INSERT INTO " & TB3 & _
                    " VALUES (" & SText(Me.lstPStaff.List(Me.lstPStaff.listIndex, 1)) & "," & SText(Me.txtLandLine) & "," & SText(Me.txtWMCPhone) & "," & SText(Me.txtAddress) & ");"
        End If
    Case TB4
        strSql = "UPDATE " & TB4 & " a " & _
                "SET a.phone = " & SText(Me.txtPhone_PStaff) & _
                " WHERE a.lifeno = " & SText(Me.lstPStaff.List(Me.lstPStaff.listIndex)) & ";"
    Case TB5
        strSql = "UPDATE " & TB5 & " a " & _
                "SET a.phone = " & SText(Me.txtPhone_Spouse) & _
                " WHERE a.lifeno = " & SText(Me.lstPStaff.List(Me.lstPStaff.listIndex)) & ";"
    Case Else
    End Select
    makeUpdateSQL = strSql
End Function
Private Sub sbClearVariant()
    Erase LISTFIELD
    Erase LISTDATA
    cntRecord = Empty
    strSql = vbNullString
End Sub
Sub sbtxtBox_Init()
    Me.txtLandLine = ""
    Me.txtPhone_PStaff = ""
    Me.txtPhone_Spouse = ""
    Me.txtWMCPhone = ""
    Me.txtName = ""
    Me.txtName_Spouse = ""
    Me.txtPosition = ""
    Me.txtPosition_Spouse = ""
    Me.lblChurch.Visible = False
    Me.lblTime_different.Visible = False
End Sub
Private Function CopyText(text As String) As Boolean
    
On Error GoTo nErr
    Dim MSForms_DataObject As DataObject
'    Set MSForms_DataObject = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    If MSForms_DataObject Is Nothing Then Set MSForms_DataObject = New DataObject
    MSForms_DataObject.SetText text
    MSForms_DataObject.PutInClipboard
'    Set MSForms_DataObject = Nothing
    CopyText = True
nErr:
End Function

Private Function getPosition2Joining()

    Dim strQuery As String
    strQuery = "SELECT * FROM op_system.a_position2;"
    Call makeListData(strQuery, "op_system.a_position2")
        
    Dim result As String
    Dim i As Integer
    For i = 0 To cntRecord - 1
        If i < cntRecord - 1 Then
            result = result & "'" & LISTDATA(i, 0) & "', "
        Else
            result = result & "'" & LISTDATA(i, 0) & "'"
        End If
    Next
    
    getPosition2Joining = result

End Function
