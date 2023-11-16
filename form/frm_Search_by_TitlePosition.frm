VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Search_by_TitlePosition 
   Caption         =   "������å�� �˻� ������"
   ClientHeight    =   8580.001
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4800
   OleObjectBlob   =   "frm_Search_by_TitlePosition.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "frm_Search_by_TitlePosition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cntlst1Col As Integer '//����Ʈ �÷� ��
Dim TB1 As String, TB2 As String '//���� ����� DB ���̺�
Const strOrderBy As String = "currency_un ASC" '//DB���� sort_order �ʵ�
Dim caseSave As Integer '//1: �߰�, 2: ����, 3: ����(SUSPEND), 4: ��������
Dim queryKey As Integer '//����Ʈ ��ġ ��ȯ�� ���� id
Dim LISTDATA() As String '--//DB���� �޾ƿ� rs�� �迭�� ����
Dim LISTFIELD() As String '--//DB���� �޾ƿ� rs�� �ʵ带 �迭�� ����
Dim cntRecord As Integer '--//DB���� �޾ƿ� ���ڵ��� ����
Dim strSql As String '--//SQL ������
Dim ws As Worksheet

Private Sub cboCountry_Change()
    '--//cboUnion ������ ����
    If Me.cboCountry.listIndex <> -1 Then
        strSql = "SELECT DISTINCT a.`����ȸ` FROM op_system.v_search_titleposition a WHERE a.`��������`=" & SText(Me.cboCountry.List(Me.cboCountry.listIndex)) & " AND a.`�����μ�` = " & SText(USER_DEPT) & ";"
    Else
        strSql = "SELECT DISTINCT a.`����ȸ` FROM op_system.v_search_titleposition a WHERE a.`�����μ�` = " & SText(USER_DEPT) & ";"
    End If
    Call makeListData(strSql, TB1)
    
    If Me.cboUnion.listIndex <> -1 Then
        Me.cboUnion.Clear
    End If
    Me.cboUnion.List = LISTDATA
End Sub

Private Sub cboSort1_Change()
    
'    Dim i As Long
'
'    If Me.cboSort1.ListIndex <> -1 Then
'        Me.cboSort2.Enabled = True
'    Else
'        Me.cboSort2.Enabled = False
'    End If
    
    '--//Sync list items of sort combobox
    Call setListItemForSortComboBox
    
End Sub

Private Sub setListItemForSortComboBox()

    Dim listItems As Object
    Set listItems = CreateObject("System.Collections.ArrayList")

    If Me.MultiPage1.Value = 1 Then
        strSql = "SELECT a.`����ȸ`,a.`��������`,a.`��ȸ��(��ü)` AS '��ȸ��',a.`��ȸ��` AS `��౳ȸ��`,a.`����ȸ ���ļ���` AS '����ȸ �������',a.`����ȸ ���ļ���` AS '����ȸ �������',a.`�����ȣ`,a.`����ѱ��̸�(����)`,a.`��𿵹��̸�`,a.`�������`,a.`�����å`,a.`����� �������`," & _
                    "a.`�����`,a.`��ü1ȸ`,a.`����ȸ�߷���`,a.`(�ؿ�)���ʹ߷���`,a.`��ȸ����` FROM op_system.v_search_titleposition a WHERE a.`�����μ�` = " & SText(USER_DEPT) & ";"
        Call makeListData(strSql, TB1)
    Else
        strSql = "SELECT a.`����ȸ`,a.`��������`,a.`��ȸ��(��ü)` AS '��ȸ��',a.`��ȸ��` AS `��౳ȸ��`,a.`����ȸ ���ļ���` AS '����ȸ �������',a.`����ȸ ���ļ���` AS '����ȸ �������',a.`�����ȣ`,a.`�ѱ��̸�(����)`,a.`�����̸�`,a.`����`,a.`��å`,a.`�������`," & _
                    "a.`����`,a.`��ü1ȸ`,a.`����ȸ�߷���`,a.`(�ؿ�)���ʹ߷���`,a.`��ȸ����` FROM op_system.v_search_titleposition a WHERE a.`�����μ�` = " & SText(USER_DEPT) & ";"
        Call makeListData(strSql, TB1)
    End If
    
    Dim tmp As Variant
    For Each tmp In LISTFIELD
        listItems.Add tmp
    Next
    
    Dim cboBox As MSForms.control
    For Each cboBox In Me.Frame_Sort.controls
        If TypeName(cboBox) = "ComboBox" Then
            If cboBox.Value <> "" Then
                listItems.Remove cboBox.Value
            End If
        End If
    Next
    
    For Each cboBox In Me.Frame_Sort.controls
        If TypeName(cboBox) = "ComboBox" Then
            cboBox.List = listItems.ToArray
        End If
    Next
    
End Sub

Private Sub cboSort2_Change()
    
    '--//Sync list items of sort combobox
    Call setListItemForSortComboBox
    
End Sub

Private Sub cboSort3_Change()

    '--//Sync list items of sort combobox
    Call setListItemForSortComboBox

End Sub

Private Sub cboSort4_Change()
        
    '--//Sync list items of sort combobox
    Call setListItemForSortComboBox
    
End Sub

Private Sub MultiPage1_Change()
    
    Dim i As Long
    Dim chkBox As MSForms.control
    
    '--//������ ������ üũ�ڽ� �ʱ�ȭ
    If Me.MultiPage1.Value = 1 Then
        For Each chkBox In Me.Frame_Position.controls
            If TypeName(chkBox) = "CheckBox" Then
                chkBox.Value = 0
            End If
        Next
        For Each chkBox In Me.Frame_Title.controls
            If TypeName(chkBox) = "CheckBox" Then
                chkBox.Value = 0
            End If
        Next
        For Each chkBox In Me.Frame_Sort.controls
            If TypeName(chkBox) = "ComboBox" Then
                chkBox.Value = ""
            End If
        Next
    ElseIf Me.MultiPage1.Value = 0 Then
        For Each chkBox In Me.Frame_Position_Spouse.controls
            If TypeName(chkBox) = "CheckBox" Then
                chkBox.Value = 0
            End If
        Next
        For Each chkBox In Me.Frame_Title_Spouse.controls
            If TypeName(chkBox) = "CheckBox" Then
                chkBox.Value = 0
            End If
        Next
        For Each chkBox In Me.Frame_Sort.controls
            If TypeName(chkBox) = "ComboBox" Then
                chkBox.Value = ""
            End If
        Next
    End If
    
    '--//�����ؿ� ���� ���� ����
    If Me.MultiPage1.Value = 1 Then
        '--//cboNationality ������ �߰�
        strSql = "SELECT DISTINCT a.`�����` FROM op_system.v_search_titleposition a WHERE a.`�����` is not null AND a.`�����μ�` = " & SText(USER_DEPT) & " ORDER BY a.`�����`"
        Call makeListData(strSql, TB1)
        Me.cboNationality.List = LISTDATA
        Me.lblNationality.Caption = "�����"
    ElseIf Me.MultiPage1.Value = 0 Then
        '--//cboNationality ������ �߰�
        strSql = "SELECT DISTINCT a.`����` FROM op_system.v_search_titleposition a WHERE a.`�����μ�` = " & SText(USER_DEPT) & " ORDER BY a.`����`"
        Call makeListData(strSql, TB1)
        Me.cboNationality.List = LISTDATA
        Me.lblNationality.Caption = "����"
    End If
    
    '--//Sync list items of sort combobox
    Call setListItemForSortComboBox
    
End Sub

Private Sub UserForm_Initialize()
    
    Dim i As Long
    
    If connIP = Empty Then setGlobalVariant '//�������� �缳��
    If checkLogin = 0 Then f_login.Show '//�α���üũ
    If checkLogin = 0 Then End '--//�α��� ���� �� ���ν��� ����
    
    '--//��Ʈ����
    Set ws = ActiveSheet
    
    '--//���ʼ���
    Me.cmdClose.Cancel = True
    TB1 = "op_system.v_search_titleposition" '--//������å�� ����Ʈ
    
    '--//��Ʈ�� ����
    Me.cmdClose.Cancel = True
'    Me.cboSort2.Enabled = False '--//���ı���1�� ���õǰ� �� ���Ŀ� Ȱ��ȭ
'    Me.cboSort3.Enabled = False '--//���ı���2�� ���õǰ� �� ���Ŀ� Ȱ��ȭ
'    Me.cboSort4.Enabled = False '--//���ı���3�� ���õǰ� �� ���Ŀ� Ȱ��ȭ
    
    
'--//�޺��ڽ� ������ �߰�
    '--//cboCountry ������ �߰�
    strSql = "SELECT DISTINCT a.`��������` FROM op_system.v_search_titleposition a WHERE a.`�����μ�` = " & SText(USER_DEPT) & " ORDER BY a.`��������`"
    Call makeListData(strSql, TB1)
    Me.cboCountry.List = LISTDATA
    
    '--//���� �ʱ�ȭ
    Call sbClearVariant
    
    '--//cboUnion ������ �߰�
    strSql = "SELECT DISTINCT a.`����ȸ` FROM op_system.v_search_titleposition a WHERE a.�����μ� = " & SText(USER_DEPT) & " ORDER BY a.`����ȸ`"
    Call makeListData(strSql, TB1)
    Me.cboUnion.List = LISTDATA
    
    '--//cboNationality ������ �߰�
    strSql = "SELECT DISTINCT a.`����` FROM op_system.v_search_titleposition a WHERE a.�����μ� = " & SText(USER_DEPT) & " ORDER BY a.`����`"
    Call makeListData(strSql, TB1)
    Me.cboNationality.List = LISTDATA
    
    '--//���� �ʱ�ȭ
    Call sbClearVariant
    
    '--//cboSort1,2,3,4 ������ �߰�
    Call setListItemForSortComboBox
    
    '--//Set Multipage 0
    Me.MultiPage1.Value = 0
    
    '--//���� �ʱ�ȭ
    Call sbClearVariant
    
End Sub

Private Sub cmdSearch_Click()
    
    Dim i As Long
    
    '--//��ȿ�� �˻�
    If isSelected_Posision = False And isSelected_Title = False Then
        MsgBox "�˻��� ��å Ȥ�� ������ ��� �ϳ� �̻� ������ �ּ���.", vbCritical, "�˻����� ����"
        Exit Sub
    End If
    
    '--//��Ʈ Ȱ��ȭ, ����ȭ, �������
    WB_ORIGIN.Activate
    ws.Activate
    Call Optimization
    Call shUnprotect(globalSheetPW)
    
    '--//���� �ʱ�ȭ
    Call sbInitialize_From
    
    '--//strSQL�� ������� listData ����
    strSql = makeSelectSQL(TB1)
    Call makeListData(strSql, TB1)
    
    '--//��µ� Data ������ ���� ����Ʈ �ۼ�
    Call sbMakeReport
    
    '--//�ʿ��� ������ ����
    Range("TitlePosition_rngCntRecord") = cntRecord
    Range("TitlePosition_rngCntField") = UBound(LISTFIELD) + 1
    Range("TitlePosition_rngSearchCode") = Me.MultiPage1.Value
    
    '--//��ȯ�� ListData�� ���� ��Ʈ�� ����
    Optimization
    If cntRecord > 0 Then
        Range("TitlePosition_rngTarget").Resize(cntRecord, UBound(LISTFIELD) + 1) = LISTFIELD
        Range("TitlePosition_rngTarget").Offset(1).Resize(cntRecord, UBound(LISTFIELD) + 1) = LISTDATA
    End If
    
    Normal
    
    '--//��������
    Call sbInsertPic
    
    '--//��ȸ������ ���
    Range("TitlePosition_Date") = Format(DateSerial(year(Date), month(Date) - 1, 1), "yyyy�� mm��")
    
    '--//���ı��� ���
    Range("TitlePosition_rngSort").ClearContents
    If Me.cboSort1.Value <> "" Then
        Range("TitlePosition_rngSort") = "���ı���: 1. " & Me.cboSort1.Value
    End If
    If Me.cboSort2.Value <> "" Then
        Range("TitlePosition_rngSort") = Range("TitlePosition_rngSort") & ", 2. " & Me.cboSort2.Value
    End If
    If Me.cboSort3.Value <> "" Then
        Range("TitlePosition_rngSort") = Range("TitlePosition_rngSort") & ", 3. " & Me.cboSort3.Value
    End If
    If Me.cboSort4.Value <> "" Then
        Range("TitlePosition_rngSort") = Range("TitlePosition_rngSort") & ", 4. " & Me.cboSort4.Value
    End If
    
    '--//��� �� �� �����ְ�
    Application.CalculateFullRebuild
    Range("D3").Select
    
    '--//�μ⿵�� ����
    ActiveSheet.PageSetup.PrintArea = Range(Cells(1, "D"), Cells(Cells(Rows.Count, "A").End(xlUp).Row, "O")).Address
    
    '--//������
    Call sbMakeTitle
    If InStr(Range("TitlePosition_rngTitle"), "��å����") > 0 Then
        With Range("TitlePosition_rngTitle").Characters(Start:=WorksheetFunction.Find("��å����", Range("TitlePosition_rngTitle")), Length:=Len(Range("TitlePosition_rngTitle")) - 16).Font
            .color = vbBlue
            .FontStyle = "���� ����Ӳ�"
            .Size = 12
        End With
    Else
        With Range("TitlePosition_rngTitle").Characters(Start:=WorksheetFunction.Find("��������", Range("TitlePosition_rngTitle")), Length:=Len(Range("TitlePosition_rngTitle")) - 16).Font
            .color = vbBlue
            .FontStyle = "���� ����Ӳ�"
            .Size = 12
        End With
    End If
    With Range("TitlePosition_rngTitle").Characters(Start:=InStrRev(Range("TitlePosition_rngTitle"), " "), Length:=Len(Range("TitlePosition_rngTitle")) - InStrRev(Range("TitlePosition_rngTitle"), " ")).Font
        .color = vbRed
    End With
    
    '--//���� �ʱ�ȭ
    Call sbClearVariant
    
    shProtect globalSheetPW
    Call Normal
End Sub

Private Function isSelected_Posision()

    Dim control As MSForms.control
    Dim result As Boolean
    
    result = False
    
    If Me.MultiPage1.Value = 0 Then
        For Each control In Me.Frame_Position.controls
            If TypeName(control) = "CheckBox" Then
                If control.Value = True Then
                    result = True
                    Exit For
                End If
            End If
        Next
    Else
        For Each control In Me.Frame_Position_Spouse.controls
            If TypeName(control) = "CheckBox" Then
                If control.Value = True Then
                    result = True
                    Exit For
                End If
            End If
        Next
    End If
    
    isSelected_Posision = result

End Function

Private Function isSelected_Title()

    Dim control As MSForms.control
    Dim result As Boolean
    
    result = False
    
    If MultiPage1.Value = 0 Then
        For Each control In Me.Frame_Title.controls
            If TypeName(control) = "CheckBox" Then
                If control.Value = True Then
                    result = True
                    Exit For
                End If
            End If
        Next
    Else
        For Each control In Me.Frame_Title_Spouse.controls
            If TypeName(control) = "CheckBox" Then
                If control.Value = True Then
                    result = True
                    Exit For
                End If
            End If
        Next
    End If
    
    isSelected_Title = result

End Function

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub makeListData(ByVal strSql As String, ByVal tableNM As String)
    '#################################################
    'DB���� �޾ƿ� �����͸� listData �迭�� �����մϴ�.
    '#################################################
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

Private Function makeSelectSQL(ByVal tableNM As String) As String
    '########################################
    '���õ� ���ǿ� ���� Select���� �ۼ��մϴ�.
    'makeSelectSQL(���̺��)
    '########################################
    Dim conPosition As String
    Dim conTitle As String
    Dim conSub As String
    Dim conSort As String
    Dim WhereClause As String
    Dim OrderClause As String
    
    Select Case tableNM
    Case TB1
        '--//���Ǻ� SelectSQL�� ����
        strSql = "SELECT * FROM (SELECT a.* FROM " & TB1 & " a UNION " & _
                    "SELECT b.`����ڻ���`,b.`��ȸ��`,b.`������ȸ��`,b.`����ȸ��`,b.`��������ȸ��`,b.`��������`,b.`����ѱ��̸�(����)`,b.`��𿵹��̸�`,b.`�����å`,b.`�����å2`,b.`����� �������`,b.`�����`,b.`(�ؿ�)���ʹ߷���`,b.`����ȸ�߷���`,NULL,NULL,NULL,NULL,NULL,NULL,NULL,b.`�������`,NULL,NULL,b.`����ȸ`,b.`��ü1ȸ`,b.`�л�1ȸ`,b.`����ȸ��ü1ȸ`,b.`����ȸ�л�1ȸ`,b.`��������ȸ`,b.`���������`,b.`����`,b.`����ȸ������`,b.`����Ұ�����`,b.`�������`,NULL,b.`�����å2������`,NULL,b.`��ü1ȸ(2�� ��)`,b.`�л�1ȸ(2�� ��)`,b.`����ȸ��ü1ȸ(2�� ��)`,b.`����ȸ�л�1ȸ(2�� ��)`,b.`�����μ�`,b.`����ȸ ���ļ���`,b.`����ȸ ���ļ���`,b.`��ȸ����`,b.`����ȸ ���ļ���`,b.`��ȸ��(��ü)`, b.`����ȸ����` FROM " & TB1 & " b WHERE b.`�����å2` <> '') a "
        
        '--//��ȸ�� ��å�� ����
        If Me.chkOverseer.Value Then
            If conPosition = "" Then
                conPosition = "a.`��å`='��ȸ��'"
            Else
                conPosition = conPosition & " OR a.`��å`='��ȸ��'"
            End If
        End If
        
        If Me.chkOverseer_Temp.Value Then
            If conPosition = "" Then
                conPosition = "a.`��å`='��ȸ��븮'"
            Else
                conPosition = conPosition & " OR a.`��å`='��ȸ��븮'"
            End If
        End If
        
        If Me.chkAssistant.Value Then
            If conPosition = "" Then
                conPosition = "a.`��å`='����'"
            Else
                conPosition = conPosition & " OR a.`��å`='����'"
            End If
        End If
        
        If Me.chkTheological.Value Then
            If conPosition = "" Then
                conPosition = "a.`�������` is not null"
            Else
                conPosition = conPosition & " OR a.`�������` is not null"
            End If
        End If
        
        If Me.chkBCLeader.Value Then
            If conPosition = "" Then
                conPosition = "a.`��å`='����ȸ������'"
            Else
                conPosition = conPosition & " OR a.`��å`='����ȸ������'"
            End If
        End If
        
        If Me.chkPBCLeader.Value Then
            If conPosition = "" Then
                conPosition = "a.`��å`='����Ұ�����'"
            Else
                conPosition = conPosition & " OR a.`��å`='����Ұ�����'"
            End If
        End If
        
        If Me.chkBuildingManager.Value Then
            If conPosition = "" Then
                conPosition = "a.`��å2` LIKE '%�ǹ�����%'"
            Else
                conPosition = conPosition & " OR a.`��å2` LIKE '%�ǹ�����%'"
            End If
        End If
        
        If Me.chkTranslator.Value Then
            If conPosition = "" Then
                conPosition = "a.`��å2` LIKE '%����%'"
            Else
                conPosition = conPosition & " OR a.`��å2` LIKE '%����%'"
            End If
        End If
        
        If Me.chkGeneralAffair.Value Then
            If conPosition = "" Then
                conPosition = "a.`��å2` LIKE '%����%'"
            Else
                conPosition = conPosition & " OR a.`��å2` LIKE '%����%'"
            End If
        End If
        
        If Me.chkMission.Value Then
            If conPosition = "" Then
                conPosition = "a.`��å2` LIKE '%�ں�%'"
            Else
                conPosition = conPosition & " OR a.`��å2` LIKE '%�ں�%'"
            End If
        End If
        
        '--//��� ��å�� ����
        If Me.chkOverseerWife.Value Then
            If conPosition = "" Then
                conPosition = "a.`�����å`='����'"
            Else
                conPosition = conPosition & " OR a.`�����å`='����'"
            End If
        End If
        
        If Me.chkOverseerWife_Temp.Value Then
            If conPosition = "" Then
                conPosition = "a.`�����å`='��븮���'"
            Else
                conPosition = conPosition & " OR a.`�����å`='��븮���'"
            End If
        End If
        
        If Me.chkAssistantWife.Value Then
            If conPosition = "" Then
                conPosition = "a.`�����å`='�����'"
            Else
                conPosition = conPosition & " OR a.`�����å`='�����'"
            End If
        End If
        
        If Me.chkTheologicalWife.Value Then
            If conPosition = "" Then
                conPosition = "a.`�����å`='�������'"
            Else
                conPosition = conPosition & " OR a.`�����å`='�������'"
            End If
        End If
        
        If Me.chkBCLeaderWife.Value Then
            If conPosition = "" Then
                conPosition = "a.`�����å`='�����ڻ��'"
            Else
                conPosition = conPosition & " OR a.`�����å`='�����ڻ��'"
            End If
        End If
        
        If Me.chkPBCLeaderWife.Value Then
            If conPosition = "" Then
                conPosition = "a.`�����å`='�����ڻ��'"
            Else
                conPosition = conPosition & " OR a.`�����å`='�����ڻ��'"
            End If
        End If
        
        '--//��ȸ�� ���п� ����
        If Me.chkPastor.Value Then
            If conTitle = "" Then
                conTitle = "a.`����`='���'"
            Else
                conTitle = conTitle & " OR a.`����`='���'"
            End If
        End If
        
        If Me.chkElder.Value Then
            If conTitle = "" Then
                conTitle = "a.`����`='���'"
            Else
                conTitle = conTitle & " OR a.`����`='���'"
            End If
        End If
        
        If Me.chkMissionary.Value Then
            If conTitle = "" Then
                conTitle = "a.`����`='������'"
            Else
                conTitle = conTitle & " OR a.`����`='������'"
            End If
        End If
        
        If Me.chkDeacon.Value Then
            If conTitle = "" Then
                conTitle = "a.`����`='����'"
            Else
                conTitle = conTitle & " OR a.`����`='����'"
            End If
        End If
        
        If Me.chkBrother.Value Then
            If conTitle = "" Then
                conTitle = "a.`����` is null"
            Else
                conTitle = conTitle & " OR a.`����` is null"
            End If
        End If
        
        '--//��� ���п� ����
        If Me.chkSeniorDeaconess.Value Then
            If conTitle = "" Then
                conTitle = "a.`�������`='�ǻ�'"
            Else
                conTitle = conTitle & " OR a.`�������`='�ǻ�'"
            End If
        End If
        
        If Me.chkMissionaryF.Value Then
            If conTitle = "" Then
                conTitle = "a.`�������`='������'"
            Else
                conTitle = conTitle & " OR a.`�������`='������'"
            End If
        End If
        
        If Me.chkDeaconess.Value Then
            If conTitle = "" Then
                conTitle = "a.`�������`='����'"
            Else
                conTitle = conTitle & " OR a.`�������`='����'"
            End If
        End If
        
        If Me.chkSister.Value Then
            If conTitle = "" Then
                conTitle = "a.`�������` is null"
            Else
                conTitle = conTitle & " OR a.`�������` is null"
            End If
        End If
        
        '--//�������ǿ� ����
        If Me.cboCountry.listIndex <> -1 Then
            If conSub = "" Then
                conSub = "a.`��������`=" & SText(Me.cboCountry.List(Me.cboCountry.listIndex))
            Else
                conSub = conSub & "AND a.`��������`=" & SText(Me.cboCountry.List(Me.cboCountry.listIndex))
            End If
        End If
        
        If Me.cboUnion.listIndex <> -1 Then
            If conSub = "" Then
                conSub = "a.`����ȸ`=" & SText(Me.cboUnion.List(Me.cboUnion.listIndex))
            Else
                conSub = conSub & "AND a.`����ȸ`=" & SText(Me.cboUnion.List(Me.cboUnion.listIndex))
            End If
        End If
        
        If Me.MultiPage1.Value = 0 Then '--//��ȸ�� �˻��� ��
            If Me.cboNationality.listIndex <> -1 Then
                If conSub = "" Then
                    conSub = "a.`����`=" & SText(Me.cboNationality.List(Me.cboNationality.listIndex))
                Else
                    conSub = conSub & "AND a.`����`=" & SText(Me.cboNationality.List(Me.cboNationality.listIndex))
                End If
            End If
        ElseIf Me.MultiPage1.Value = 1 Then '--//��� �˻��� ��
            If Me.cboNationality.listIndex <> -1 Then
                If conSub = "" Then
                    conSub = "a.`�����`=" & SText(Me.cboNationality.List(Me.cboNationality.listIndex))
                Else
                    conSub = conSub & "AND a.`�����`=" & SText(Me.cboNationality.List(Me.cboNationality.listIndex))
                End If
            End If
        Else
        End If
        
        '--//���ı��ؿ� ����
        conSort = makeOrderByClause(Me.cboSort1, conSort)
        conSort = makeOrderByClause(Me.cboSort2, conSort)
        conSort = makeOrderByClause(Me.cboSort3, conSort)
        conSort = makeOrderByClause(Me.cboSort4, conSort)
        
        '--//WHERE�� ��������
        If conPosition <> "" Then
            If WhereClause = "" Then
                WhereClause = "WHERE (" & conPosition & ")"
            Else
                WhereClause = WhereClause & " AND (" & conPosition & ")"
            End If
        End If
        
        If conTitle <> "" Then
            If WhereClause = "" Then
                WhereClause = "WHERE (" & conTitle & ")"
            Else
                WhereClause = WhereClause & " AND (" & conTitle & ")"
            End If
        End If
        
        If conSub <> "" Then
            If WhereClause = "" Then
                WhereClause = "WHERE (" & conSub & ")"
            Else
                WhereClause = WhereClause & " AND (" & conSub & ")"
            End If
        End If
        
        If WhereClause = "" Then
            WhereClause = "WHERE" & " `�����μ�` = " & SText(USER_DEPT)
        Else
            WhereClause = WhereClause & " AND `�����μ�` = " & SText(USER_DEPT)
        End If
        
        '--//ORDER BY�� ����
        If conSort <> "" Then
            OrderClause = " ORDER BY " & conSort
        End If
        
        strSql = strSql & WhereClause & OrderClause & ";"
        
    Case Else
    End Select
    makeSelectSQL = strSql
End Function

Private Function makeOrderByClause(cboList As MSForms.comboBox, conSort As String)

    If cboList.Value <> "" Then
        Select Case cboList.Value
            Case "����ȸ":
                conSort = AppendText(conSort, "a.`" & cboList.Value & " ���ļ���`")
            Case "��ü1ȸ":
                conSort = AppendText(conSort, "a.`" & cboList.Value & "`" & " DESC, a.`��ü1ȸ(2�� ��)` DESC")
            Case "��౳ȸ��":
                conSort = AppendText(conSort, "a.`" & Replace(cboList.Value, "���", "") & "`")
            Case "��ȸ��":
                conSort = AppendText(conSort, "a.`��ȸ��(��ü)`")
            Case "����ȸ �������"
                conSort = AppendText(conSort, "a.`" & Replace(cboList.Value, "����", "����") & "`")
            Case "����ȸ �������"
                conSort = AppendText(conSort, "a.`" & Replace(cboList.Value, "����", "����") & "`")
            Case "����", "�������":
                If Me.chkAssistantWife.Value Then
                    conSort = AppendText(conSort, "a." & "`�������` IS NULL ASC, FIELD(`�������`,'���','���','�ǻ�','������','����','����','�ڸ�','')")
                Else
                    conSort = AppendText(conSort, "a." & "`����` IS NULL ASC, FIELD(`����`,'���','���','�ǻ�','������','����','����','�ڸ�','')")
                End If
            Case "��å", "�����å":
                If Me.chkAssistantWife.Value Then
                    conSort = AppendText(conSort, "a." & "`�����å` IS NULL ASC, FIELD(`�����å`,'����','��븮���','�����','�����ڻ��','�����ڻ��','�������'," & getPosition2Joining & ",'')")
                Else
                    conSort = AppendText(conSort, "a." & "`��å` IS NULL ASC, FIELD(`��å`,'��ȸ��','��ȸ��븮','����','����ȸ������','����Ұ�����','�������1�ܰ�','�������2�ܰ�','�������3�ܰ�','�������'," & getPosition2Joining & ",'')")
                End If
            Case "��ȸ����":
                conSort = AppendText(conSort, "FIELD(`��ȸ����`,'����ȸ','�����','')")
            Case Else
                conSort = AppendText(conSort, "a.`" & cboList.Value & "`")
        End Select
    End If
    
    makeOrderByClause = conSort

End Function

Private Function AppendText(sourceText As String, targetText As String)

    Dim result As String
    
    If sourceText = "" Then
        result = targetText
    Else
        result = sourceText & ", " & targetText
    End If
    
    AppendText = result

End Function

Public Sub sbInsertPic()
    '###############################
    '������ ��ġ�� ������ �����մϴ�.
    '###############################
    Dim lifeNo As String
    
    ActiveSheet.Pictures.Delete

    '--//���� �ֱ� ���μ���
On Error Resume Next
    Dim i As Long, j As Long
    For i = 4 To Cells(Rows.Count, "A").End(xlUp).Row Step 7: For j = 5 To 14 Step 2
        '--//��������
        lifeNo = Cells(i, j).Value
        
        '--//��������
        If lifeNo <> "" Then
            InsertPStaffPic lifeNo, Cells(i, j).Resize(4)
        End If
    Next j: Next i
    
    '--//�������� ��¥���� �ϳ� �߰�(����Ʋ���� ���� ����)
    InsertPStaffPic lifeNo, Cells(1, "D")
    
    '--//���� Ʋ���� ������ ���� ������ ���� ����
    If ActiveSheet.Pictures.Count > 0 Then
        ActiveSheet.Pictures(ActiveSheet.Pictures.Count).Delete
    End If
    
On Error GoTo 0

End Sub
Private Sub sbClearVariant()
    '##########################
    '���� ������ �ʱ�ȭ �մϴ�.
    '##########################
    Erase LISTFIELD
    Erase LISTDATA
    cntRecord = Empty
    strSql = vbNullString
End Sub
Private Sub sbSortData_By_Union()
    '######################################################################
    '���ı��ؿ� ����ȸ�� ���ԵǾ� ���� ��� ������ ����ȸ ������ �����մϴ�.
    '######################################################################
    
    Select Case USER_DEPT
    Case 10 '--//�ƽþ�3��
        ActiveWorkbook.Worksheets("������å�� ��Ȳ").Sort.SortFields.Clear
        ActiveWorkbook.Worksheets("������å�� ��Ȳ").Sort.SortFields.Add key:=Range("TitlePosition_rngTarget").Offset(1, 24).Resize(Range("TitlePosition_rngTarget").CurrentRegion.Rows.Count - 1) _
            , SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder:= _
            "īƮ����,���ȵ���,�����ߺ�,���ȼ���" _
            , DataOption:=xlSortNormal
        With ActiveWorkbook.Worksheets("������å�� ��Ȳ").Sort
            .SetRange Range("TitlePosition_rngTarget").CurrentRegion
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    Case Else
        ActiveWorkbook.Worksheets("������å�� ��Ȳ").Sort.SortFields.Clear
        ActiveWorkbook.Worksheets("������å�� ��Ȳ").Sort.SortFields.Add key:=Range("TitlePosition_rngTarget").Offset(1, 24).Resize(Range("TitlePosition_rngTarget").CurrentRegion.Rows.Count - 1) _
            , SortOn:=xlSortOnValues, Order:=xlAscending _
            , DataOption:=xlSortNormal
        With ActiveWorkbook.Worksheets("������å�� ��Ȳ").Sort
            .SetRange Range("TitlePosition_rngTarget").CurrentRegion
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    End Select
End Sub

Public Sub sbInitialize_From()
    '############################
    '����Ʈ ����� �ʱ�ȭ �մϴ�.
    '############################
    Dim R As Long
    
    '--//���� ������ ����
    Range("TitlePosition_rngTarget").CurrentRegion.Offset(1).ClearContents
    
    '--//������ �� ã��
    R = Cells(Rows.Count, "A").End(xlUp).Row
    Range("10:10").Resize(R).Delete Shift:=xlUp

End Sub

Private Sub sbMakeReport()
    '################################################
    '�˻��� ������ ������ ���� ����Ʈ ��� �����մϴ�.
    '################################################
    
    Dim rngTarget As Range
    Dim cntRow As Long
    
    '--//���Կ� �ʿ��� �� ����
    cntRow = ((cntRecord - 1) \ 5) * 7
    
    '--//ù ���� ������ ����Ʈ ����
    If cntRow <> 0 Then
        Set rngTarget = Range("10:10").Resize(cntRow)
    
    
        '--//ù �� ����
        Range("3:9").Copy
    
        '--//rngTarget�� �ٿ��ֱ�(1. ����, 2. ����)
        rngTarget.PasteSpecial Paste:=xlPasteFormats
        rngTarget.PasteSpecial Paste:=xlPasteFormulas
    End If
    
End Sub

Private Sub sbMakeTitle()

    '##################################
    '���Ǻ��� ����Ʈ ������ �����մϴ�.
    '##################################
    
    Dim strTitle As String
    Dim conPosition As String
    Dim conTitle As String
    Dim conSub As String
    
    strSql = "SELECT a.dept_nm FROM op_system.db_ovs_dept a WHERE a.dept_id = " & SText(USER_DEPT)
    Call makeListData(strSql, "op_system.db_ovs_dept")
    
    strTitle = LISTDATA(0, 0) & " ���Ǻ� ��ȸ�� ���" & vbNewLine
    
    '--//��ȸ�� ��å�� ����
    If Me.chkOverseer.Value Then
        If conPosition = "" Then
            conPosition = "��ȸ��" & "(" & WorksheetFunction.CountIf(Range("AD:AD"), "��ȸ��") & ")"
        Else
            conPosition = conPosition & ",��ȸ��" & "(" & WorksheetFunction.CountIf(Range("AD:AD"), "��ȸ��") & ")"
        End If
    End If
    
    If Me.chkOverseer_Temp.Value Then
        If conPosition = "" Then
            conPosition = "��븮" & "(" & WorksheetFunction.CountIf(Range("AD:AD"), "��ȸ��븮") & ")"
        Else
            conPosition = conPosition & ",��븮" & "(" & WorksheetFunction.CountIf(Range("AD:AD"), "��ȸ��븮") & ")"
        End If
    End If
    
    If Me.chkAssistant.Value Then
        If conPosition = "" Then
            conPosition = "����" & "(" & WorksheetFunction.CountIf(Range("AD:AD"), "����") & ")"
        Else
            conPosition = conPosition & ",����" & "(" & WorksheetFunction.CountIf(Range("AD:AD"), "����") & ")"
        End If
    End If
    
    If Me.chkTheological.Value Then
        If conPosition = "" Then
            conPosition = "�������" & "(" & WorksheetFunction.CountIf(Range("AP:AP"), "*�������*") & ")"
        Else
            conPosition = conPosition & ",�������" & "(" & WorksheetFunction.CountIf(Range("AP:AP"), "*�������*") & ")"
        End If
    End If
    
    If Me.chkBCLeader.Value Then
        If conPosition = "" Then
            conPosition = "������" & "(" & WorksheetFunction.CountIf(Range("AD:AD"), "����ȸ������") & ")"
        Else
            conPosition = conPosition & ",������" & "(" & WorksheetFunction.CountIf(Range("AD:AD"), "����ȸ������") & ")"
        End If
    End If
    
    If Me.chkPBCLeader.Value Then
        If conPosition = "" Then
            conPosition = "������" & "(" & WorksheetFunction.CountIf(Range("AD:AD"), "����Ұ�����") & ")"
        Else
            conPosition = conPosition & ",������" & "(" & WorksheetFunction.CountIf(Range("AD:AD"), "����Ұ�����") & ")"
        End If
    End If
    
    If Me.chkBuildingManager.Value Then
        If conPosition = "" Then
            conPosition = "�ǹ�����" & "(" & WorksheetFunction.CountIf(Range("AE:AE"), "*�ǹ�����*") & ")"
        Else
            conPosition = conPosition & ",�ǹ�����" & "(" & WorksheetFunction.CountIf(Range("AE:AE"), "*�ǹ�����*") & ")"
        End If
    End If
    
    If Me.chkTranslator.Value Then
        If conPosition = "" Then
            conPosition = "������" & "(" & WorksheetFunction.CountIf(Range("AE:AE"), "*������*") & ")"
        Else
            conPosition = conPosition & ",������" & "(" & WorksheetFunction.CountIf(Range("AE:AE"), "*������*") & ")"
        End If
    End If
    
    If Me.chkGeneralAffair.Value Then
        If conPosition = "" Then
            conPosition = "��������" & "(" & WorksheetFunction.CountIf(Range("AE:AE"), "*��������*") & ")"
        Else
            conPosition = conPosition & ",��������" & "(" & WorksheetFunction.CountIf(Range("AE:AE"), "*��������*") & ")"
        End If
    End If
    
    If Me.chkMission.Value Then
        If conPosition = "" Then
            conPosition = "�ں�" & "(" & WorksheetFunction.CountIf(Range("AE:AE"), "*�ں�*") & ")"
        Else
            conPosition = conPosition & ",�ں�" & "(" & WorksheetFunction.CountIf(Range("AE:AE"), "*�ں�*") & ")"
        End If
    End If
    
    '--//��� ��å�� ����
    If Me.chkOverseerWife.Value Then
        If conPosition = "" Then
            conPosition = "����" & "(" & WorksheetFunction.CountIf(Range("AM:AM"), "����") & ")"
        Else
            conPosition = conPosition & ",����" & "(" & WorksheetFunction.CountIf(Range("AM:AM"), "����") & ")"
        End If
    End If
    
    If Me.chkOverseerWife_Temp.Value Then
        If conPosition = "" Then
            conPosition = "��븮���" & "(" & WorksheetFunction.CountIf(Range("AM:AM"), "��븮���") & ")"
        Else
            conPosition = conPosition & ",��븮���" & "(" & WorksheetFunction.CountIf(Range("AM:AM"), "��븮���") & ")"
        End If
    End If
    
    If Me.chkAssistantWife.Value Then
        If conPosition = "" Then
            conPosition = "�����" & "(" & WorksheetFunction.CountIf(Range("AM:AM"), "�����") & ")"
        Else
            conPosition = conPosition & ",�����" & "(" & WorksheetFunction.CountIf(Range("AM:AM"), "�����") & ")"
        End If
    End If
    
    If Me.chkTheologicalWife.Value Then
        If conPosition = "" Then
            conPosition = "�������" & "(" & WorksheetFunction.CountIfs(Range("AM:AM"), "*�������*", Range("AJ:AJ"), "<>""""") & ")"
        Else
            conPosition = conPosition & ",�������" & "(" & WorksheetFunction.CountIfs(Range("AM:AM"), "*�������*", Range("AJ:AJ"), "<>""""") & ")"
        End If
    End If
    
    If Me.chkBCLeaderWife.Value Then
        If conPosition = "" Then
            conPosition = "�����ڻ��" & "(" & WorksheetFunction.CountIf(Range("AM:AM"), "�����ڻ��") & ")"
        Else
            conPosition = conPosition & ",�����ڻ��" & "(" & WorksheetFunction.CountIf(Range("AM:AM"), "�����ڻ��") & ")"
        End If
    End If
    
    If Me.chkPBCLeaderWife.Value Then
        If conPosition = "" Then
            conPosition = "�����ڻ��" & "(" & WorksheetFunction.CountIf(Range("AM:AM"), "�����ڻ��") & ")"
        Else
            conPosition = conPosition & ",�����ڻ��" & "(" & WorksheetFunction.CountIf(Range("AM:AM"), "�����ڻ��") & ")"
        End If
    End If
    
    '--//��ȸ�� ���п� ����
    If Me.chkPastor.Value Then
        If conTitle = "" Then
            conTitle = "���" & "(" & WorksheetFunction.CountIf(Range("AQ:AQ"), "���") & ")"
        Else
            conTitle = conTitle & ",���" & "(" & WorksheetFunction.CountIf(Range("AQ:AQ"), "���") & ")"
        End If
    End If
    
    If Me.chkElder.Value Then
        If conTitle = "" Then
            conTitle = "���" & "(" & WorksheetFunction.CountIf(Range("AQ:AQ"), "���") & ")"
        Else
            conTitle = conTitle & ",���" & "(" & WorksheetFunction.CountIf(Range("AQ:AQ"), "���") & ")"
        End If
    End If
    
    If Me.chkMissionary.Value Then
        If conTitle = "" Then
            conTitle = "������" & "(" & WorksheetFunction.CountIf(Range("AQ:AQ"), "������") & ")"
        Else
            conTitle = conTitle & ",������" & "(" & WorksheetFunction.CountIf(Range("AQ:AQ"), "������") & ")"
        End If
    End If
    
    If Me.chkDeacon.Value Then
        If conTitle = "" Then
            conTitle = "����" & "(" & WorksheetFunction.CountIf(Range("AQ:AQ"), "����") & ")"
        Else
            conTitle = conTitle & ",����" & "(" & WorksheetFunction.CountIf(Range("AQ:AQ"), "����") & ")"
        End If
    End If
    
    If Me.chkBrother.Value Then
        If conTitle = "" Then
            conTitle = "����" & "(" & WorksheetFunction.CountA(Range("V:V")) - WorksheetFunction.CountA(Range("AQ:AQ")) & ")"
        Else
            conTitle = conTitle & ",����" & "(" & WorksheetFunction.CountA(Range("V:V")) - WorksheetFunction.CountA(Range("AQ:AQ")) & ")"
        End If
    End If
    
    '--//��� ���п� ����
    If Me.chkSeniorDeaconess.Value Then
        If conTitle = "" Then
            conTitle = "�ǻ�" & "(" & WorksheetFunction.CountIf(Range("AR:AR"), "�ǻ�") & ")"
        Else
            conTitle = conTitle & ",�ǻ�" & "(" & WorksheetFunction.CountIf(Range("AR:AR"), "�ǻ�") & ")"
        End If
    End If
    
    If Me.chkMissionaryF.Value Then
        If conTitle = "" Then
            conTitle = "������" & "(" & WorksheetFunction.CountIf(Range("AR:AR"), "������") & ")"
        Else
            conTitle = conTitle & ",������" & "(" & WorksheetFunction.CountIf(Range("AR:AR"), "������") & ")"
        End If
    End If
    
    If Me.chkDeaconess.Value Then
        If conTitle = "" Then
            conTitle = "����" & "(" & WorksheetFunction.CountIf(Range("AR:AR"), "����") & ")"
        Else
            conTitle = conTitle & ",����" & "(" & WorksheetFunction.CountIf(Range("AR:AR"), "����") & ")"
        End If
    End If
    
    If Me.chkSister.Value Then
        If conTitle = "" Then
            conTitle = "�ڸ�" & "(" & WorksheetFunction.CountA(Range("V:V")) - WorksheetFunction.CountA(Range("AR:AR")) & ")"
        Else
            conTitle = conTitle & ",�ڸ�" & "(" & WorksheetFunction.CountA(Range("V:V")) - WorksheetFunction.CountA(Range("AR:AR")) & ")"
        End If
    End If
    
    '--//�������ǿ� ����
    If Me.cboCountry.listIndex <> -1 Then
        If conSub = "" Then
            conSub = "����: " & Me.cboCountry.List(Me.cboCountry.listIndex)
        Else
            conSub = conSub & " / ����: " & Me.cboCountry.List(Me.cboCountry.listIndex)
        End If
    End If
    
    If Me.cboUnion.listIndex <> -1 Then
        If conSub = "" Then
            conSub = "����ȸ: " & Me.cboUnion.List(Me.cboUnion.listIndex)
        Else
            conSub = conSub & " / ����ȸ: " & Me.cboUnion.List(Me.cboUnion.listIndex)
        End If
    End If
    
    If Me.cboNationality.listIndex <> -1 Then
        If conSub = "" Then
            conSub = "����: " & Me.cboNationality.List(Me.cboNationality.listIndex)
        Else
            conSub = conSub & " / ����: " & Me.cboNationality.List(Me.cboNationality.listIndex)
        End If
    End If
    
    '--//���� �����Ͽ� �������
    If conPosition <> "" Then
        strTitle = strTitle & _
                    "��å����: " & conPosition & vbNewLine
    End If
    
    If conTitle <> "" Then
        strTitle = strTitle & _
                    "��������: " & conTitle & vbNewLine
    End If
    
    If conSub <> "" Then
        strTitle = strTitle & _
                    conSub & vbNewLine
    End If
    
'    strTitle = Left(strTitle, Len(strTitle) - 1)
    strTitle = strTitle & "�˻� �� �ο�: " & WorksheetFunction.CountA(Range("V:V")) - 1 & "��"
    Range("TitlePosition_rngTitle") = strTitle
    
End Sub

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
