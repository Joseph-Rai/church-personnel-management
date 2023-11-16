VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Search_PStaff_Detail 
   Caption         =   "������ ������ �˻�������"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5850
   OleObjectBlob   =   "frm_Search_PStaff_Detail.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "frm_Search_PStaff_Detail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cntlst1Col As Integer '//����Ʈ �÷� ��
Dim TB1 As String, TB2 As String, TB3 As String, TB4 As String, TB5 As String, TB6 As String, TB7 As String, TB8 As String, TB9 As String, TB10 As String, TB11 As String, TB12 As String, TB13 As String, TB14 As String '//���� ����� DB ���̺�
Const strOrderBy As String = "currency_un ASC" '//DB���� sort_order �ʵ�
Dim caseSave As Integer '//1: �߰�, 2: ����, 3: ����(SUSPEND), 4: ��������
Dim queryKey As Integer '//����Ʈ ��ġ ��ȯ�� ���� id
Dim LISTDATA() As String '--//DB���� �޾ƿ� rs�� �迭�� ����
Dim LISTFIELD() As String '--//DB���� �޾ƿ� rs�� �ʵ带 �迭�� ����
Dim cntRecord As Integer '--//DB���� �޾ƿ� ���ڵ��� ����
Dim strSql As String '--//SQL ������
Dim ws As Worksheet

Private Sub lstPStaff_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Un'HookListBoxScroll
End Sub

Private Sub lstPStaff_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Me.lstPStaff.ListCount > 0 Then
        'HookListBoxScroll Me, Me.lstPStaff
    End If
End Sub

Private Sub UserForm_Initialize()
    
    If connIP = Empty Then setGlobalVariant '//�������� �缳��
    If checkLogin = 0 Then f_login.Show '//�α���üũ
    If checkLogin = 0 Then End '--//�α��� ���� �� ���ν��� ����
    
    '--//��Ʈ����
    Set ws = ActiveSheet
    
    '--//���ʼ���
    Me.cmdClose.Cancel = True
    TB1 = "op_system.v0_pstaff_information" '--//���������� ������ �˻��� ����
    
    TB3 = "op_system.v_pstaff_detail" '--//��ȸ�� �⺻����
    TB4 = "op_system.v_pstaff_detail_title" '--//�����̷�
    TB5 = "op_system.v_pstaff_detail_transfer" '--//�߷��̷�
    TB6 = "op_system.v_pstaff_detail_flight" '--//�װ�������
    TB7 = "op_system.v_pstaff_detail_accomplishment" '--//��������
    TB8 = "op_system.v_familyinfo" '--//��������
    TB9 = "op_system.v_pstaff_detail_accomplishment_main" '--//��������(����ȸ)
    TB10 = "op_system.v_pstaff_detail_accomplishment_both" '--//��������(��ü+����ȸ)
    TB11 = "op_system.v_pstaff_detail_transfer2" '--//�߷��̷�2
    TB12 = "op_system.v_pstaff_detail_concise_transfer_history" '--//���� ��ȸ���
    TB13 = "op_system.v_pstaff_detail_concise_transfer_history_main" '--//���� ��ȸ���
    TB14 = "op_system.v_pstaff_detail_concise_transfer_history_both" '--//���� ��ȸ���
    
    '--//����Ʈ�ڽ� ����
    With Me.lstPStaff
        .ColumnCount = 5
        .ColumnHeads = False
        .ColumnWidths = "0,120,70,50,0" '�����ȣ, ��ȸ��, �ѱ��̸�(����), ��å,����ڻ���
        .TextAlign = fmTextAlignLeft
        .Font = "����"
    End With
    
    Me.optAll.Value = True
    Me.cmdOk.Enabled = False
    Me.txtChurchNM.SetFocus

End Sub

Private Sub cmdSearch_Click()
    Me.lstPStaff.Clear
    strSql = makeSelectSQL(TB1)
    Call makeListData(strSql, TB1)
    If cntRecord > 0 Then
        Me.lstPStaff.List = LISTDATA
    End If
    Call sbClearVariant
End Sub

Private Sub txtChurchNM_Change()
    Me.txtChurchNM.BackColor = RGB(255, 255, 255)
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub
Private Sub lstPStaff_Click()
    Me.cmdOk.Enabled = True
End Sub
Private Sub cmdOK_Click()

    Dim i As Integer, j As Integer
    Dim filePath As String
    Dim FileName As String
    Dim rngTarget As Range
    
    '--//��Ʈ Ȱ��ȭ, ����ȭ, �������
    WB_ORIGIN.Activate
    ws.Activate
    Call Optimization
    Call shUnprotect(globalSheetPW)
    
    '--//���� ������ ����
    Range("PStaff_Detail_rngTarget").CurrentRegion.ClearContents
    Range("PStaff_Detail_Title").Offset(1).Resize(3, 6).ClearContents
    Range("PStaff_Detail_Transfer").Offset(1).Resize(10, 6).ClearContents
    Range("PStaff_Detail_Flight").Offset(1).Resize(5, 6).ClearContents
    Range("PStaff_Detail_rngAtten").CurrentRegion.ClearContents
    Range("PStaff_Detail_rngFamily").CurrentRegion.ClearContents
    
    '--//������ �⺻���� ����
        strSql = makeSelectSQL(TB3)
        connectTaskDB
        Call makeListData(strSql, TB3)
    '    callDBtoRS "cmdOK_Click", TB2, strSQL, Me.Name
        
        '--//��ȯ�� ListData�� ���� ��Ʈ�� ����
        Optimization
        Range("PStaff_Detail_rngTarget").Resize(, UBound(LISTFIELD) + 1) = LISTFIELD
        If cntRecord > 0 Then
            Range("PStaff_Detail_rngTarget").Offset(1).Resize(cntRecord, UBound(LISTFIELD) + 1) = LISTDATA
        End If
        Normal
        
        Call sbClearVariant
        disconnectALL
    
    '--//�������� ����
        strSql = makeSelectSQL(TB8) '--//��������
        connectTaskDB
        Call makeListData(strSql, TB8)
    '    callDBtoRS "cmdOK_Click", TB2, strSQL, Me.Name
    
        '--//��ȯ�� ListData�� ���� ��Ʈ�� ����
        Optimization
        Range("PStaff_Detail_rngFamily").Resize(, UBound(LISTFIELD) + 1) = LISTFIELD
        If cntRecord > 0 Then
            Range("PStaff_Detail_rngFamily").Offset(1).Resize(cntRecord, UBound(LISTFIELD) + 1) = LISTDATA
        End If
        Normal
        
        Call sbClearVariant
        disconnectALL
    
    '--//���� �Ӹ��̷� ����
        strSql = makeSelectSQL(TB4) '--//������ �����Ӹ� �̷�
        connectTaskDB
        Call makeListData(strSql, TB4)
    '    callDBtoRS "cmdOK_Click", TB2, strSQL, Me.Name
        
        '--//��ȯ�� ListData�� ���� ��Ʈ�� ����
        Optimization
        Range("PStaff_Detail_Title").Resize(, UBound(LISTFIELD) + 1) = LISTFIELD
        If cntRecord > 0 Then
            Range("PStaff_Detail_Title").Offset(1).Resize(cntRecord, UBound(LISTFIELD) + 1) = LISTDATA
        End If
        Normal
        
        Call sbClearVariant
        disconnectALL
        
        strSql = makeSelectSQL2(TB4) '--//��� �����Ӹ� �̷�
        connectTaskDB
        Call makeListData(strSql, TB4)
    '    callDBtoRS "cmdOK_Click", TB2, strSQL, Me.Name
        
        '--//��ȯ�� ListData�� ���� ��Ʈ�� ����
        Optimization
        Range("PStaff_Detail_Title").Offset(, 3).Resize(, UBound(LISTFIELD) + 1) = LISTFIELD
        If cntRecord > 0 Then
            Range("PStaff_Detail_Title").Offset(1, 3).Resize(cntRecord, UBound(LISTFIELD) + 1) = LISTDATA
        End If
        Normal
        
        Call sbClearVariant
        disconnectALL
    
    '--//��ȸ��� ����
        Dim objPastroalCareerDao As New PastoralCareerDao
        '--//����������
        
        Range("PStaff_Detail_PHistory").Offset(1) = objPastroalCareerDao.GetMinDateForAssistantOverseer(Range("PStaff_Detail_LifeNo"))
        '--//��ȸ�� ������
        Range("PStaff_Detail_PHistory").Offset(1, 1) = objPastroalCareerDao.GetMinDateForOverseer(Range("PStaff_Detail_LifeNo"))
        '--//�������
        Range("PStaff_Detail_PHistory").Offset(3) = objPastroalCareerDao.GetAssistantOverseerCareer(Range("PStaff_Detail_LifeNo"))
        '--//��ȸ����
        Range("PStaff_Detail_PHistory").Offset(3, 1) = objPastroalCareerDao.GetOverseerCareer(Range("PStaff_Detail_LifeNo"))
        '--//��ȸ���
        Range("PStaff_Detail_PHistory").Offset(3, 5) = objPastroalCareerDao.GetTotalPastoralCareer(Range("PStaff_Detail_LifeNo"))
    
    connectTaskDB
    '--//�߷��̷� ����
        If Range("D4") <> 0 Then '--//��å�� ������
            strSql = "SELECT * FROM (SELECT `�߷���`,`����/��å`,`��ȸ����`,`��ȸ��`,'',`�Ⱓ` FROM op_system.v_pstaff_detail_transfer WHERE `�����ȣ` = " & SText(Range("PStaff_Detail_LifeNo")) & " AND `����/��å` IS NOT NULL LIMIT 10) a ORDER BY `�߷���`;"
            Call makeListData(strSql, TB5)
        Else
            strSql = "SELECT * FROM (SELECT `�߷���`,`����/��å`,`��ȸ����`,`��ȸ��`,'',`�Ⱓ` FROM op_system.v_pstaff_detail_transfer WHERE `�����ȣ` = " & SText(Range("PStaff_Detail_LifeNo")) & " LIMIT 10) a ORDER BY `�߷���`;"
            Call makeListData(strSql, TB5)
        End If
        
        '--//��ȯ�� ListData�� ���� ��Ʈ�� ����
        Optimization
        
        Range("PStaff_Detail_Transfer").Resize(, UBound(LISTFIELD) + 1) = LISTFIELD
        If cntRecord > 0 Then
            Range("PStaff_Detail_Transfer").Offset(1).Resize(cntRecord, UBound(LISTFIELD) + 1) = LISTDATA
        End If
        
        Normal
        
        Call sbClearVariant
        disconnectALL
    
    '--//���Ա� �̷� ����
        strSql = makeSelectSQL(TB6) '--//������ ���Ա� �̷�
        connectTaskDB
        Call makeListData(strSql, TB6)
    '    callDBtoRS "cmdOK_Click", TB2, strSQL, Me.Name
        
        '--//��ȯ�� ListData�� ���� ��Ʈ�� ����
        Optimization
        Range("PStaff_Detail_Flight").Resize(, UBound(LISTFIELD) + 1) = LISTFIELD
        If cntRecord > 0 Then
            Range("PStaff_Detail_Flight").Offset(1).Resize(cntRecord, UBound(LISTFIELD) + 1) = LISTDATA
        End If
        Normal
        
        Call sbClearVariant
        disconnectALL
        
        strSql = makeSelectSQL2(TB6) '--//��� ���Ա� �̷�
        connectTaskDB
        Call makeListData(strSql, TB6)
    '    callDBtoRS "cmdOK_Click", TB2, strSQL, Me.Name
        
        '--//��ȯ�� ListData�� ���� ��Ʈ�� ����
        Optimization
        Range("PStaff_Detail_Flight").Offset(, 3).Resize(, UBound(LISTFIELD) + 1) = LISTFIELD
        If cntRecord > 0 Then
            Range("PStaff_Detail_Flight").Offset(1, 3).Resize(cntRecord, UBound(LISTFIELD) + 1) = LISTDATA
        End If
        Normal
        
        Call sbClearVariant
        disconnectALL
    
    Application.CalculateFullRebuild
    
    Range("9:9").EntireRow.AutoFit '--//�ǰ� ����� �ڵ�����
'    Range("15:15").EntireRow.AutoFit '--//�������� ����� �ڵ�����
    
    '--//�ٳణ ��ȸ��� ����
    If Me.optAll.Value Then
        strSql = makeSelectSQL(TB12)
    Call makeListData(strSql, TB12)
    ElseIf Me.optMain.Value Then
        strSql = makeSelectSQL(TB13)
    Call makeListData(strSql, TB13)
    ElseIf Me.optBoth Then
        strSql = makeSelectSQL(TB14)
    Call makeListData(strSql, TB14)
    End If
    Range("PStaff_Detail_cntChurch").Offset(1).Resize(15, UBound(LISTFIELD) + 1).ClearContents
    If cntRecord > 0 Then
        Range("PStaff_Detail_cntChurch").Offset(1).Resize(cntRecord, UBound(LISTFIELD)) = LISTDATA
        Range("PStaff_Detail_cntChurch").Offset(1, UBound(LISTFIELD)).Resize(cntRecord).FormulaR1C1 = _
        "=SUMIFS(OFFSET(PStaff_Detail_rngAtten,,2,1000,1),OFFSET(PStaff_Detail_rngAtten,,,1000,1),RC20,OFFSET(PStaff_Detail_rngAtten,,1,1000,1),RC21)"
        Range("PStaff_Detail_cntChurch").Offset(-1).Copy
        Range("PStaff_Detail_cntChurch").Offset(1, 1).Resize(cntRecord, 5).PasteSpecial xlPasteValues, xlPasteSpecialOperationAdd
        Application.CutCopyMode = False
    End If
    
    '--//�ߺ� ����
    Dim curChurchName As String
    Dim prevChurchName As String
    Dim curStartDate As Date
    Dim prevStartDate As Date
    For i = 1 To cntRecord
        curChurchName = Range("PStaff_Detail_cntChurch").Offset(i)
        prevChurchName = Range("PStaff_Detail_cntChurch").Offset(i + 1)
        If curChurchName = prevChurchName Then
            curStartDate = Range("PStaff_Detail_cntChurch").Offset(i, 1)
            prevStartDate = Range("PStaff_Detail_cntChurch").Offset(i + 1, 1)
            If prevStartDate >= curStartDate And prevStartDate <> 0 Then
                '--//2023.09.25 ������ �ܿ��� �����Ǹ� �ȵ�.
                '--//����,��å�� �����Ǹ� "���� ������å"�� "����" ������å�� ǥ�õ�
'                Range("PStaff_Detail_cntChurch").Offset(i, 1) = Range("PStaff_Detail_cntChurch").Offset(i + 1, 1) '//������ ����
'                Range("PStaff_Detail_cntChurch").Offset(i, 6) = Range("PStaff_Detail_cntChurch").Offset(i + 1, 6) '//���� ����
'                Range("PStaff_Detail_cntChurch").Offset(i, 7) = Range("PStaff_Detail_cntChurch").Offset(i + 1, 7) '//��å ����
'                Range("PStaff_Detail_cntChurch").Offset(i, 8) = Range("PStaff_Detail_cntChurch").Offset(i + 1, 8) '//��ȸ���� ����
                Range("PStaff_Detail_cntChurch").Offset(i, 2) = Range("PStaff_Detail_cntChurch").Offset(i + 1, 2) '//������ ����
                
                Range("PStaff_Detail_cntChurch").Offset(i, 5) = _
                    DateDiff("m", Range("PStaff_Detail_cntChurch").Offset(i, 1), Range("PStaff_Detail_cntChurch").Offset(i, 2)) '//�Ⱓ ����
'                Range("PStaff_Detail_cntChurch").Offset(i + 1).Resize(, UBound(listField) + 1).Delete Shift:=xlUp
                Range("PStaff_Detail_cntChurch").Offset(i + 2).Resize(100, UBound(LISTFIELD) + 1).Copy
                Range("PStaff_Detail_cntChurch").Offset(i + 1).Resize(100, UBound(LISTFIELD) + 1).PasteSpecial Paste:=xlPasteFormulas
                Application.CutCopyMode = False
                cntRecord = cntRecord - 1
            End If
        End If
    Next
    Range("PStaff_Detail_cntChurch").Offset(-3) = cntRecord
    Range("PStaff_Detail_cntChurch").Offset(-3, 2) = UBound(LISTFIELD)
    
    
    '--//�������� �⼮������ ����
    If Me.optAll.Value Then
        strSql = makeSelectSQL(TB7)
        connectTaskDB
        Call makeListData(strSql, TB7)
    ElseIf Me.optMain.Value Then
        strSql = makeSelectSQL(TB9)
        connectTaskDB
        Call makeListData(strSql, TB9)
    ElseIf Me.optBoth Then
        strSql = makeSelectSQL(TB10)
        connectTaskDB
        Call makeListData(strSql, TB10)
    End If
    
    '--//��ȯ�� ListData�� ���� ��Ʈ�� ����
    Optimization
    Range("PStaff_Detail_rngAtten").Resize(, UBound(LISTFIELD) + 1) = LISTFIELD
    If cntRecord > 0 Then
        Range("PStaff_Detail_rngAtten").Offset(1).Resize(cntRecord, UBound(LISTFIELD) + 1) = LISTDATA
    End If
    Normal
    
    Call sbClearVariant
    disconnectALL
    
    On Error Resume Next
    Range("PStaff_Detail_rngAtten").Offset(-1).Copy
    Range("PStaff_Detail_rngAtten").CurrentRegion.Offset(1, 1).PasteSpecial Paste:=xlPasteValues, operation:=xlPasteSpecialOperationAdd
    Range("PStaff_Detail_rngFamily").Offset(1, Range("PStaff_Detail_rngFamily_Rank") - 1).Resize(Range("PStaff_Detail_rngFamily_cntData")).PasteSpecial Paste:=xlPasteValues, operation:=xlPasteSpecialOperationAdd
    Range("PStaff_Detail_rngFamily").Offset(1, 0).Resize(Range("PStaff_Detail_rngFamily_cntData"), 2).PasteSpecial Paste:=xlPasteValues, operation:=xlPasteSpecialOperationAdd
    On Error GoTo 0
    
    Application.CutCopyMode = False
    
    '--//�⼮ 0���� �����ϸ� ���� �޷� �̵�
    Dim fieldCount As Integer
    i = 1
    fieldCount = Range("PStaff_Detail_cntChurch").Offset(-3, 2)
    Do While Range("PStaff_Detail_cntChurch").Offset(i) <> ""
        Do While Range("PStaff_Detail_cntChurch").Offset(i, fieldCount) = 0 And _
            Range("PStaff_Detail_cntChurch").Offset(i, 1) <= Range("PStaff_Detail_cntChurch").Offset(i, 2)
                Range("PStaff_Detail_cntChurch").Offset(i, 1) = WorksheetFunction.EDate(Range("PStaff_Detail_cntChurch").Offset(i, 1), 1)
        Loop
        i = i + 1
    Loop
    
    '--//��Ʈ����
    Application.CalculateFullRebuild
    Call sbArrangeChart_Atten
    
    '--//�Ⱦ��� �������� ��Ʈ �����
    ActiveSheet.Outline.ShowLevels RowLevels:=2
    On Error Resume Next
    Range(Range("PStaff_Detail_Church1").Offset(-1), Range("PStaff_Detail_Church8").Offset(20)).Rows.Ungroup
    Range("PStaff_Detail_Nationality").Resize(3).EntireRow.Rows.Ungroup
    Range("PStaff_Detail_rngFamily").Resize(11).EntireRow.Rows.Ungroup
    Range("PStaff_Detail_PHistory").Resize(4).EntireRow.Rows.Ungroup
    Range("PStaff_Detail_Transfer").Resize(11).EntireRow.Rows.Ungroup
    On Error GoTo 0
    '--//�������� �׷���
    For i = 0 To 8
        If Range("PStaff_Detail_Family").Offset(i + 2) = "" And Range("PStaff_Detail_Family").Offset(i + 2, 3) = "" Then
            Range("PStaff_Detail_Family").Offset(i + 2).Rows.Group
        End If
    Next
    On Error Resume Next
'    rngTarget.Rows.Group
    On Error GoTo 0
    '--//��ȸ��¿��� �׷���
    If Not (Range("PStaff_Detail_CurrentPosition") = "��ȸ��" Or _
            Range("PStaff_Detail_CurrentPosition") = "��ȸ��븮" Or _
            Range("PStaff_Detail_CurrentPosition") = "����" Or _
            Range("PStaff_Detail_CurrentPosition") Like "*������*" Or _
            Range("PStaff_Detail_CurrentPosition") Like "*����*") Then
            
        Range("PStaff_Detail_PHistory").Resize(4).EntireRow.Rows.Group
    End If
    
    '--//�߷��̷¿��� �׷���(�ּ� 4�� ���� ���̵��� ����)
    For i = 5 To 10
        If Range("PStaff_Detail_Transfer").Offset(i) = "" Then
            Range("PStaff_Detail_Transfer").Offset(i).EntireRow.Rows.Group
        End If
    Next
    
    '--//���ڿ��� �׷���
    '--//������, ������� ���������� ������ ��� ���� ��쿡�� �׷���(���ڰ� �ʿ� ���� ���)
    If Range("PStaff_Detail_GospelCountry") = Range("PStaff_Detail_Nationality").Offset(1) And _
        Range("PStaff_Detail_GospelCountry") = Range("PStaff_Detail_Nationality").Offset(1, 3) Then
        Range(Range("PStaff_Detail_Nationality"), Range("PStaff_Detail_Nationality").Offset(2)).EntireRow.Rows.Group
    End If
    '--//��Ʈ���� �׷���
    Select Case Range("PStaff_Detail_cntChurch").Offset(-3)
    Case 0
        Range(Range("PStaff_Detail_Church1").Offset(-1), Range("PStaff_Detail_Church8").Offset(20)).EntireRow.Rows.Group
    Case 1
        Range(Range("PStaff_Detail_Church2"), Range("PStaff_Detail_Church8").Offset(20)).EntireRow.Rows.Group
    Case 2
        Range(Range("PStaff_Detail_Church3").Offset(-1), Range("PStaff_Detail_Church8").Offset(20)).EntireRow.Rows.Group
    Case 3
        Range(Range("PStaff_Detail_Church4"), Range("PStaff_Detail_Church8").Offset(20)).EntireRow.Rows.Group
    Case 4
        Range(Range("PStaff_Detail_Church5").Offset(-1), Range("PStaff_Detail_Church8").Offset(20)).EntireRow.Rows.Group
    Case 5
        Range(Range("PStaff_Detail_Church6"), Range("PStaff_Detail_Church8").Offset(20)).EntireRow.Rows.Group
    Case 6
        Range(Range("PStaff_Detail_Church7").Offset(-1), Range("PStaff_Detail_Church8").Offset(20)).EntireRow.Rows.Group
    Case 7
        Range(Range("PStaff_Detail_Church8"), Range("PStaff_Detail_Church8").Offset(20)).EntireRow.Rows.Group
    Case Else
    End Select
    ActiveSheet.Outline.ShowLevels RowLevels:=1
    
    '--//����Ʈ���� ����
    ActiveSheet.PageSetup.Zoom = 100
    Set ActiveSheet.HPageBreaks(1).Location = Range("PStaff_Detail_Church1").Offset(-1)
    
    '--//��������
On Error Resume Next
    ActiveSheet.Pictures.Delete
    
    If Range("PStaff_Detail_LifeNo") <> "" Then
        InsertPStaffPic Range("PStaff_Detail_LifeNo"), Range("PStaff_Detail_Pic_M")
    End If

    If Not (Range("PStaff_Detail_LifeNo_Spouse") = "" Or Range("PStaff_Detail_LifeNo_Spouse") = "0") Then
        InsertPStaffPic Range("PStaff_Detail_LifeNo_Spouse"), Range("PStaff_Detail_Pic_F")
    End If
    
    InsertPStaffPic "", Range("J1")
    
    If ActiveSheet.Pictures.Count > 0 Then
        ActiveSheet.Pictures(ActiveSheet.Pictures.Count).Delete
    End If
Here:
On Error GoTo 0
    
    Sheets("������ ������").Range("A10").Select
    Sheets("������ ������").Range("A1").Select
    
    Call shProtect(globalSheetPW)
    Call Normal
    
    MsgBox "�۾��� �Ϸ�Ǿ����ϴ�."
    
End Sub

Private Function GetMinDateForAssistantOverseer(lifeNo As String) As String

    strSql = "" & _
        " SELECT MIN(p.Start_dt)" & _
        " FROM op_system.db_position p" & _
        " WHERE p.lifeno = " & SText(lifeNo) & _
        "     AND p.Position IN ('����');"
    makeListData strSql, "op_system.db_position"
    
    Dim result As String
    
    If cntRecord > 0 Then
        result = LISTDATA(0, 0)
    Else
        result = ""
    End If
    
    GetMinDateForAssistantOverseer = result

End Function

Private Function GetMinDateForOverseer(lifeNo As String) As String

    strSql = "" & _
        " SELECT MIN(p.Start_dt)" & _
        " FROM op_system.db_position p" & _
        " WHERE p.lifeno = " & SText(lifeNo) & _
        "     AND p.Position IN ('��ȸ��', '��ȸ��븮');"
    makeListData strSql, "op_system.db_position"
    
    Dim result As String
    
    If cntRecord > 0 Then
        result = LISTDATA(0, 0)
    Else
        result = ""
    End If
    
    GetMinDateForOverseer = result

End Function

Private Function GetAssistantOverseerCareer(lifeNo As String) As String

    '--//�������� Ȱ���� �̷� ����
    strSql = "" & _
        " SELECT p.Start_dt, p.End_dt" & _
        " FROM op_system.db_position p" & _
        " WHERE p.lifeno = " & SText(lifeNo) & _
        "     AND p.Position IN ('����');"
    makeListData strSql, "op_system.db_position"
    
    Dim result As String
    result = GetConvertedFormatPeriod
    
    GetAssistantOverseerCareer = result

End Function

Private Function GetOverseerCareer(lifeNo As String) As String

    '--//��ȸ������ Ȱ���� �̷� ����
    strSql = "" & _
        " SELECT p.Start_dt, p.End_dt" & _
        " FROM op_system.db_position p" & _
        " WHERE p.lifeno = " & SText(lifeNo) & _
        "     AND p.Position IN ('��ȸ��', '��ȸ��븮');"
    makeListData strSql, "op_system.db_position"
    
    Dim result As String
    result = GetConvertedFormatPeriod
    
    GetOverseerCareer = result

End Function

Private Function GetTotalPastoralCareer(lifeNo As String) As String
    
    '--//��ȸ�ڷ� Ȱ���� �̷� ����
    strSql = "" & _
        " SELECT p.Start_dt, p.End_dt" & _
        " FROM op_system.db_position p" & _
        " WHERE p.lifeno = " & SText(lifeNo) & _
        "     AND p.Position IN ('��ȸ��', '��ȸ��븮', '����');"
    makeListData strSql, "op_system.db_position"
    
    Dim result As String
    result = GetConvertedFormatPeriod
    
    GetTotalPastoralCareer = result

End Function

Private Function GetConvertedFormatPeriod() As String

    Dim year As Integer
    Dim month As Integer

    '--//Ž���� ���� minDate, maxDate ����
    Dim minDate As Date: minDate = DateSerial(9999, 12, 31)
    Dim maxDate As Date: maxDate = DateSerial(1900, 1, 1)
    
    If cntRecord <= 0 Then
        GetConvertedFormatPeriod = ""
    End If
    
    Dim i As Integer
    For i = 0 To cntRecord - 1
        minDate = WorksheetFunction.Min(minDate, LISTDATA(i, 0))
        maxDate = WorksheetFunction.Max(maxDate, WorksheetFunction.Min(Date, LISTDATA(i, 1)))
    Next
    
    '--//�� �޾� �ǳʶٸ� ��ȸ��¿� ���Ե� ��¥��� ������ �߰�
    Dim tempDate As Date: tempDate = WorksheetFunction.EoMonth(minDate, 0)
    Do
        For i = 0 To cntRecord - 1
            Dim startDate As Date: startDate = LISTDATA(i, 0)
            Dim endDate As Date: endDate = LISTDATA(i, 1)
            If startDate <= tempDate And tempDate <= endDate Then
                month = month + 1
                If month >= 12 Then
                    month = 0
                    year = year + 1
                    Exit For
                End If
            End If
        Next
        
        If tempDate = DateSerial(9999, 12, 31) Then
            tempDate = DateSerial(9999, 12, 31)
        Else
            tempDate = WorksheetFunction.EDate(tempDate, 1)
        End If
        
        If tempDate > maxDate Then
            Exit Do
        End If
    Loop

    '--//year, month => Y�� M���� �������� ��ȯ
    Dim result As String
    If year > 0 Then
        result = result & year & "��"
    End If
    
    If month > 0 Then
        If result = "" Then
            result = month & "����"
        Else
            result = result & " " & month & "����"
        End If
    End If
    
    GetConvertedFormatPeriod = result

End Function

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
        '�����ȣ, ��ȸ��, �ѱ��̸�(����), ��å
        strSql = "SELECT a.`�����ȣ`,a.`��ȸ��`,a.`�ѱ��̸�(����)`,a.`��å`,a.`����ڻ���` " & _
                    "FROM " & TB1 & " a " & _
                    "WHERE (a.`�ѱ��̸�(����)` LIKE '%" & Me.txtChurchNM & "%' OR a.`��ȸ��` LIKE '%" & Me.txtChurchNM & "%'" & _
                    " OR a.`�����̸�` LIKE '%" & Me.txtChurchNM & "%' OR a.`����ȸ��` LIKE '%" & Me.txtChurchNM & "%'" & _
                    " OR a.`������ȸ��` LIKE '%" & Me.txtChurchNM & "%' OR a.`��������ȸ��` LIKE '%" & Me.txtChurchNM & "%' " & _
                    " OR a.`�����ȣ` LIKE '%" & Me.txtChurchNM & "%') " & " AND a.`�����μ�` = " & SText(USER_DEPT) & ";"
    Case TB2
    Case TB3 '--//�⺻����
        With Me.lstPStaff
            strSql = "SELECT * FROM " & TB3 & " a WHERE a.`�����ȣ` = " & SText(.List(.listIndex))
        End With
    Case TB4 '--//�����̷�
        With Me.lstPStaff
            strSql = "SELECT * FROM (SELECT a.`��ȸ��`,a.`�Ӹ���`,a.`����` FROM " & TB4 & " a WHERE a.`�����ȣ` = " & SText(.List(.listIndex)) & " ORDER BY a.`�Ӹ���` DESC LIMIT 3) a ORDER BY a.`�Ӹ���`;"
        End With
    Case TB5 '--//�߷��̷�
        With Me.lstPStaff
            strSql = "SELECT * FROM (SELECT a.`�߷���`,a.`����/��å`,a.`��ȸ����`,a.`��ȸ��` FROM " & TB5 & " a WHERE a.`�����ȣ` = " & SText(.List(.listIndex)) & " ORDER BY a.`�߷���` DESC LIMIT 10) a ORDER BY a.`�߷���`, FIELD(a.`��ȸ����`,'MC','HBC','BC','PBC');"
        End With
    Case TB6 '--//���Ա��̷�
        With Me.lstPStaff
            strSql = "SELECT * FROM (SELECT a.`�湮����`,a.`�湮����` FROM " & TB6 & " a WHERE a.`�����ȣ` = " & SText(.List(.listIndex)) & " ORDER BY a.`�湮����` DESC LIMIT 5) a ORDER BY a.`�湮����`;"
        End With
    Case TB7 '--//��������
        With Me.lstPStaff
            strSql = "SELECT a.`��ȸ��` ,a.`��¥`,a.`��ü1ȸ`,a.`��ü4ȸ`,a.`�л�1ȸ`,a.`�л�4ȸ`,a.`����`,a.`ħ��`,a.`������`,a.`������`,a.`������`,a.`����`,a.`��å`,a.`����������`,a.`����������`,a.`��ȸ����` FROM " & TB7 & " a WHERE a.`�����ȣ` = " & SText(.List(.listIndex)) & ";"
        End With
    Case TB8 '--//��������
        
        If Range("F6") = 0 Then
            
            strSql = "SELECT DISTINCT a.family_cd FROM " & TB8 & " a WHERE a.lifeno = " & SText(Range("C6").Value) & " AND a.relations NOT IN (""��"",""��"");"
            Call makeListData(strSql, TB8)
            
            If cntRecord = 1 Then
                Range("PStaff_Detail_MaleFamilyCode") = LISTDATA(0, 0)
                strSql = "SELECT * FROM (SELECT a.family_id,a.family_cd,REPLACE(a.relations,'(����)','') AS 'relations',a.lifeno,a.name_ko,a.name_en,a.church_sid,a.church_nm,a.title,a.POSITION,a.birthday,a.education,a.religion,a.recognition,a.memo,a.SUSPEND,CONCAT(a.title,if(a.title<>'' AND a.position<>'','/',''),a.position) AS 'status',RANK() OVER (PARTITION BY a.family_cd,relations ORDER BY a.birthday) AS 'rank' FROM op_system.v_familyinfo a WHERE a.family_cd = " & SText(LISTDATA(0, 0)) & " ORDER BY family_cd,FIELD(relations,'��','��(����)','��','��(����)','����','�ڸ�'),birthday) a WHERE a.lifeno <> " & SText(Range("C6").Value) & ";"
            ElseIf cntRecord > 1 Then
                MsgBox "������ �������� �����Ϳ� �ߺ������� �ֽ��ϴ�. �ߺ��� �ڷḦ �����ϼ���.", vbCritical, banner
            End If
        Else
            strSql = "SELECT DISTINCT a.family_cd FROM " & TB8 & " a WHERE a.lifeno = " & SText(Range("C6").Value) & " AND a.relations NOT IN (""��"",""��"")"
            Call makeListData(strSql, TB8)
            
            If cntRecord = 0 Then
                strSql = "SELECT DISTINCT a.family_cd FROM " & TB8 & " a WHERE a.lifeno = " & SText(Range("F6").Value) & " AND a.relations NOT IN (""��"",""��"");"
                Call makeListData(strSql, TB8)
                
                If cntRecord = 1 Then
                    Range("PStaff_Detail_MaleFamilyCode") = LISTDATA(0, 0)
                    strSql = "SELECT * FROM (SELECT a.family_id,a.family_cd,REPLACE(a.relations,'(����)','') AS 'relations',a.lifeno,a.name_ko,a.name_en,a.church_sid,a.church_nm,a.title,a.POSITION,a.birthday,a.education,a.religion,a.recognition,a.memo,a.SUSPEND,CONCAT(a.title,if(a.title<>'' AND a.position<>'','/',''),a.position) AS 'status',RANK() OVER (PARTITION BY a.family_cd,relations ORDER BY a.birthday) AS 'rank' FROM op_system.v_familyinfo a WHERE a.family_cd = " & SText(LISTDATA(0, 0)) & " ORDER BY family_cd,FIELD(relations,'��','��(����)','��','��(����)','����','�ڸ�'),birthday) a WHERE a.lifeno NOT IN (" & SText(Range("F6").Value) & ");"
                End If
            ElseIf cntRecord = 1 Then
                strSql = "SELECT DISTINCT a.family_cd FROM " & TB8 & " a WHERE a.lifeno = " & SText(Range("F6").Value) & " AND a.relations NOT IN (""��"",""��"");"
                Call makeListData(strSql, TB8)
                
                If cntRecord = 1 Then
                    strSql = "SELECT DISTINCT a.family_cd FROM " & TB8 & " a WHERE a.lifeno = " & SText(Range("C6").Value) & " AND a.relations NOT IN (""��"",""��"")" & _
                                " UNION SELECT DISTINCT a.family_cd FROM " & TB8 & " a WHERE a.lifeno = " & SText(Range("F6").Value) & " AND a.relations NOT IN (""��"",""��"");"
                    Call makeListData(strSql, TB8)
                    
                    Range("PStaff_Detail_MaleFamilyCode") = LISTDATA(0, 0)
                    Range("PStaff_Detail_FemaleFamilyCode") = LISTDATA(1, 0)
                    strSql = "SELECT * FROM (SELECT a.family_id,a.family_cd,REPLACE(a.relations,'(����)','') AS 'relations',a.lifeno,a.name_ko,a.name_en,a.church_sid,a.church_nm,a.title,a.POSITION,a.birthday,a.education,a.religion,a.recognition,a.memo,a.SUSPEND,CONCAT(a.title,if(a.title<>'' AND a.position<>'','/',''),a.position) AS 'status',RANK() OVER (PARTITION BY a.family_cd,relations ORDER BY a.birthday) AS 'rank' FROM op_system.v_familyinfo a WHERE a.family_cd = " & SText(LISTDATA(0, 0)) & _
                            " UNION SELECT a.family_id,a.family_cd,REPLACE(a.relations,'(����)','') AS 'relations',a.lifeno,a.name_ko,a.name_en,a.church_sid,a.church_nm,a.title,a.POSITION,a.birthday,a.education,a.religion,a.recognition,a.memo,a.SUSPEND,CONCAT(a.title,if(a.title<>'' AND a.position<>'','/',''),a.position) AS 'status',RANK() OVER (PARTITION BY a.family_cd,relations ORDER BY a.birthday) AS 'rank' FROM op_system.v_familyinfo a WHERE a.family_cd = " & SText(LISTDATA(1, 0)) & " ORDER BY family_cd,FIELD(relations,'��','��(����)','��','��(����)','����','�ڸ�'),birthday) a WHERE a.lifeno NOT IN (" & SText(Range("C6").Value) & "," & SText(Range("F6").Value) & ");"
                ElseIf cntRecord = 0 Then
                    Range("PStaff_Detail_MaleFamilyCode") = LISTDATA(0, 0)
                    strSql = "SELECT * FROM (SELECT a.family_id,a.family_cd,REPLACE(a.relations,'(����)','') AS 'relations',a.lifeno,a.name_ko,a.name_en,a.church_sid,a.church_nm,a.title,a.POSITION,a.birthday,a.education,a.religion,a.recognition,a.memo,a.SUSPEND,CONCAT(a.title,if(a.title<>'' AND a.position<>'','/',''),a.position) AS 'status',RANK() OVER (PARTITION BY a.family_cd,relations ORDER BY a.birthday) AS 'rank' FROM op_system.v_familyinfo a WHERE a.family_cd = " & SText(LISTDATA(0, 0)) & " ORDER BY family_cd,FIELD(relations,'��','��(����)','��','��(����)','����','�ڸ�'),birthday) a WHERE a.lifeno NOT IN (" & SText(Range("C6").Value) & ");"
                End If
            ElseIf cntRecord > 2 Then
                MsgBox "������ Ȥ�� ��� �������� �����Ϳ� �ߺ������� �ֽ��ϴ�. �ߺ��� �ڷḦ �����ϼ���.", vbCritical, banner
            End If
        End If
        
        strSql = strSql & ";"
    Case TB9 '--//��������(����ȸ)
        With Me.lstPStaff
            strSql = " SELECT a.`��ȸ��` ,a.`��¥`,a.`��ü1ȸ`,a.`��ü4ȸ`,a.`�л�1ȸ`,a.`�л�4ȸ`,a.`����`,a.`ħ��`,a.`������`,a.`������`,a.`������`,a.`����`,a.`��å`,a.`����������`,a.`����������`,a.`��ȸ����` FROM " & TB9 & " a WHERE a.`�����ȣ` = " & SText(.List(.listIndex)) & ";"
        End With
    Case TB10 '--//��������(��ü+����ȸ)
        With Me.lstPStaff
            strSql = "SELECT a.`��ȸ��` ,a.`��¥`,a.`��ü1ȸ`,a.`��ü4ȸ`,a.`�л�1ȸ`,a.`�л�4ȸ`,a.`����`,a.`ħ��`,a.`������`,a.`������`,a.`������`,a.`����`,a.`��å`,a.`����������`,a.`����������`,a.`��ȸ����` FROM " & TB7 & " a WHERE a.`�����ȣ` = " & SText(.List(.listIndex)) & " UNION " & _
                    " SELECT a.`��ȸ��` ,a.`��¥`,a.`��ü1ȸ`,a.`��ü4ȸ`,a.`�л�1ȸ`,a.`�л�4ȸ`,a.`����`,a.`ħ��`,a.`������`,a.`������`,a.`������`,a.`����`,a.`��å`,a.`����������`,a.`����������`,a.`��ȸ����` FROM " & TB9 & " a WHERE a.`�����ȣ` = " & SText(.List(.listIndex)) & " ORDER BY `����������`,`��ȸ��`,`��¥`;"
        End With
    Case TB11 '--//�߷��̷�
        With Me.lstPStaff
            strSql = "SELECT * FROM (SELECT a.`�߷���`,a.`����/��å`,a.`��ȸ����`,a.`��ȸ��` FROM " & TB11 & " a WHERE a.`�����ȣ` = " & SText(.List(.listIndex)) & " ORDER BY a.`�߷���` DESC LIMIT 10) a ORDER BY a.`�߷���`, FIELD(a.`��ȸ����`,'MC','HBC','BC','PBC');"
        End With
    Case TB12 '--//���� ��ȸ���
        With Me.lstPStaff
            strSql = "SELECT * FROM " & TB12 & " a WHERE a.`�����ȣ` = " & SText(.List(.listIndex)) & ";"
        End With
    Case TB13
        With Me.lstPStaff
            strSql = "SELECT * FROM " & TB13 & " a WHERE a.`�����ȣ` = " & SText(.List(.listIndex)) & ";"
        End With
    Case TB14
        With Me.lstPStaff
            strSql = "SELECT * FROM " & TB14 & " a WHERE a.`�����ȣ` = " & SText(.List(.listIndex)) & ";"
        End With
    Case Else
        '//ȭ��id, ȭ���Ī, ��ȸ��, ��ȭȯ��, �޷�ȭȯ��
        'strSQL = "SELECT a.currency_id, a.currency_un, a.refer_dt, a.fx_rate_krw, a.fx_rate_usd " & _
                      "FROM " & TB1 & " a WHERE a.user_id = " & user_id & ";"
    End Select
    makeSelectSQL = strSql
End Function
'-----------------------------------------
'  ���Ǻ� Select SQL�� �ۼ�
'    - makeSelectSQL(�˻���, ����)
'    - DB���� ��ȯ�� ����Ʈ �ʵ����
'-----------------------------------------
Private Function makeSelectSQL2(ByVal tableNM As String) As String
    Select Case tableNM
    Case TB4 '--//������� �̷�
        With Me.lstPStaff
            strSql = "SELECT * FROM (SELECT a.`��ȸ��`,a.`�Ӹ���`,a.`����` FROM " & TB4 & " a WHERE a.`�����ȣ` = " & SText(.List(.listIndex, 4)) & " ORDER BY a.`�Ӹ���` DESC LIMIT 3) a ORDER BY a.`�Ӹ���`;"
        End With
    Case TB6 '--//��� ���Ա� �̷�
        With Me.lstPStaff
            strSql = "SELECT * FROM (SELECT a.`�湮����`,a.`�湮����` FROM " & TB6 & " a WHERE a.`�����ȣ` = " & SText(.List(.listIndex, 4)) & " ORDER BY a.`�湮����` DESC LIMIT 5) a ORDER BY a.`�湮����`;"
        End With
    Case Else
        '//ȭ��id, ȭ���Ī, ��ȸ��, ��ȭȯ��, �޷�ȭȯ��
        'strSQL = "SELECT a.currency_id, a.currency_un, a.refer_dt, a.fx_rate_krw, a.fx_rate_usd " & _
                      "FROM " & TB1 & " a WHERE a.user_id = " & user_id & ";"
    End Select
    makeSelectSQL2 = strSql
End Function

Private Sub sbClearVariant()
    Erase LISTFIELD
    Erase LISTDATA
    cntRecord = Empty
    strSql = vbNullString
End Sub

Private Sub sbArrangeChart_Atten()

    Dim noMax As Integer
    Dim noMin As Integer
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim Term As Long
    Dim rngTarget_above1_All As Range
    Dim rngTarget_above1_Stu As Range
    Dim rngTarget_4Rate As Range
    Dim rngTarget_TitheRate As Range
    Dim ChartNM As String
    Dim RangeNM As String
    Dim rngChartNM As String
    Dim cycle As Long
    Dim Rate_Above4() As Variant
    
    
On Error Resume Next
    
    For cycle = 1 To 8
        
        '--//��Ʈ�̸� ����
        ChartNM = "Chart_Church" & cycle
        
        '--//�����̸� ����
        RangeNM = "Result_rngChurch" & cycle
        
        '--//���� ����
        Set rngTarget_above1_All = Range(RangeNM & "_1All")
        Set rngTarget_above1_Stu = Range(RangeNM & "_1Stu")
        Set rngTarget_4Rate = Range(RangeNM & "_4Rate")
        Set rngTarget_TitheRate = Range(RangeNM & "_TitheRate")
        
        '--//��Ʈ�� �ִ밪, �ּҰ� ��������
        noMax = WorksheetFunction.Max(rngTarget_above1_All) '--//�л��̻� 1ȸ�⼮ �ִ밪
        noMin = WorksheetFunction.Min(rngTarget_above1_Stu) '--//�л��̻� 4ȸ�⼮ �ּҰ�
        i = 1: j = 1
          
          
          '--//�⼮ �׷��� ��Ʈ����
          With Sheets("������ ������").ChartObjects(ChartNM).Chart.Axes(xlValue)
            
            '--//�ı� �Ը� ���� �������� �޸� �մϴ�.
            Select Case noMax
                Case Is <= 100: Term = 10
                Case Is <= 500: Term = 50
                Case Is <= 1000: Term = 100
                Case Else: Term = 100
            End Select
            
            '--//������ �ִ밪�� ���մϴ�..
            Do
                If Term * i > noMax Then
                    .MaximumScale = Term * i
                    Exit Do
                End If
                i = i + 1
            Loop
            
            '--//������ �ּҰ��� ���մϴ�.
            Do
                If Term * j >= noMin * 0.3 Then
                    .MinimumScale = Term * (j - 1)
                    Exit Do
                End If
                j = j + 1
            Loop
            
            '--//������ �ִ밪�� �ּҰ��� ���̰� 4�� ����� �ƴϸ� �ִ밪 ����
            Do
                If (.MaximumScale - .MinimumScale) Mod 4 = 0 Then Exit Do
                i = i + 1
                .MaximumScale = Term * i
            Loop
            
            .MajorUnit = (.MaximumScale - .MinimumScale) / 4
            
          End With
          
          
          '--//4ȸ���� �迭�� ����
'          ReDim Rate_Above4(0 To Range(RangeNM).Columns.Count - 1)
'          For k = 0 To UBound(Rate_Above4)
'            Rate_Above4(k) = rngTarget_above4_Stu.Cells(k) / rngTarget_above1_Stu.Cells(k)
'          Next
          
          '--//������ ��������
          With Sheets("������ ������").ChartObjects(ChartNM).Chart.Axes(xlValue, xlSecondary)
'            .MaximumScale = Application.Max(WorksheetFunction.RoundUp(WorksheetFunction.Max(rngTarget_4Rate), 1), WorksheetFunction.RoundUp(WorksheetFunction.Max(rngTarget_TitheRate), 1))
'            .MinimumScale = Application.Min(WorksheetFunction.RoundDown(WorksheetFunction.Min(rngTarget_4Rate), 1), WorksheetFunction.RoundDown(WorksheetFunction.Min(rngTarget_TitheRate), 1))
            .MaximumScale = 3
            .MinimumScale = 0
          End With
          
          '--//��Ʈ ��ġ����
'          rngChartNM = "PStaff_Detail_Church" & cycle
'          Sheets("������ ������").ChartObjects(ChartNM).Top = Range(rngChartNM).Offset(3).Top + 4
      
      Next
On Error GoTo 0

End Sub

