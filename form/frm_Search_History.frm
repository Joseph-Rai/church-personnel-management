VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Search_History 
   Caption         =   "������ �˻�"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5655
   OleObjectBlob   =   "frm_Search_History.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "frm_Search_History"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cntlst1Col As Integer '//����Ʈ �÷� ��
Dim TB1 As String, TB2 As String, TB3 As String, TB4 As String, TB5 As String, TB6 As String '//���� ����� DB ���̺�
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
    TB2 = "op_system.v_transfer_history" '--//�����ڿ��� ����Ʈ �߷��̷��� ����
    TB3 = "op_system.v_familyinfo" '--//��������
    
    '--//����Ʈ�ڽ� ����
    With Me.lstPStaff
        .ColumnCount = 4
        .ColumnHeads = False
        .ColumnWidths = "0,120,70,50" '�����ȣ, ��ȸ��, �ѱ��̸�(����), ��å
        .TextAlign = fmTextAlignLeft
        .Font = "����"
    End With
    
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
    
    '--//��Ʈ Ȱ��ȭ �� �������
    WB_ORIGIN.Activate
    ws.Activate
    Call shUnprotect(globalSheetPW)
    
    '--//��ȸ���ÿ��� Ȯ��
    If Me.lstPStaff.listIndex = -1 Then
        MsgBox "��Ͽ� ���õ� ���� �����ϴ�.", vbCritical, "���ÿ���"
        Exit Sub
    End If
    
    '--//���� ������ ����
    Range("His_rngTarget").CurrentRegion.ClearContents
    Range("His_rngFamily").CurrentRegion.ClearContents
    
    '--//�⺻���� �� �߷��̷� ����
        '--//SQL��
        strSql = makeSelectSQL(TB2)
        
        '--//DB���� �ڷ� ȣ���Ͽ� ���ڵ�� ��ȯ
        connectTaskDB
        Call makeListData(strSql, TB2)
    '    callDBtoRS "cmdOK_Click", TB2, strSQL, Me.Name
        
        '--//��ȯ�� ListData�� ���� ��Ʈ�� ����
        Optimization
        Range("His_rngTarget").Resize(, UBound(LISTFIELD) + 1) = LISTFIELD
        If cntRecord > 0 Then
            Range("His_rngTarget").Offset(1).Resize(cntRecord, UBound(LISTFIELD) + 1) = LISTDATA
        End If
        Normal
        
        Call sbClearVariant
        disconnectALL
    
    '--//�������� ����
        strSql = makeSelectSQL(TB3) '--//��������
        connectTaskDB
        Call makeListData(strSql, TB3)
    '    callDBtoRS "cmdOK_Click", TB2, strSQL, Me.Name
    
        '--//��ȯ�� ListData�� ���� ��Ʈ�� ����
        Optimization
        Range("His_rngFamily").Resize(, UBound(LISTFIELD) + 1) = LISTFIELD
        If cntRecord > 0 Then
            Range("His_rngFamily").Offset(1).Resize(cntRecord, UBound(LISTFIELD) + 1) = LISTDATA
        End If
        Normal
        
        Call sbClearVariant
        disconnectALL
    
    '--//���ڿ��� �� ����, ���ڵ����ͷ� ��ȯ
    On Error Resume Next
    Range("His_rngFamily").Offset(-1).Copy
    Range("His_rngFamily").Offset(1, Range("AL29") - 1).Resize(Range("AK26")).PasteSpecial Paste:=xlPasteValues, operation:=xlPasteSpecialOperationAdd
    Range("His_rngFamily").Offset(1, 1).Resize(Range("AK26")).PasteSpecial Paste:=xlPasteValues, operation:=xlPasteSpecialOperationAdd
    On Error GoTo 0
    
    Application.CutCopyMode = False
    
    '--//����,�ǰ�,��Ÿ ��������
    On Error Resume Next
    ActiveSheet.Outline.ShowLevels RowLevels:=2
    Range(Range("His_Family"), Range("His_Family").Offset(10)).Rows.Ungroup
    On Error GoTo 0
    For i = 0 To 8
        If Range("His_Family").Offset(i + 2) = "" And Range("His_Family").Offset(i + 2, 4) = "" Then
            Range("His_Family").Offset(i + 2).Rows.Group
        End If
    Next
    On Error Resume Next
'    rngTarget.Rows.Group
    On Error GoTo 0
    ActiveSheet.Outline.ShowLevels RowLevels:=1
    
    Range("23:23").EntireRow.AutoFit '--//�ǰ� ����� ����
    Range("24:24").EntireRow.AutoFit '--//��Ÿ ����� ����
    
    '--//��������
On Error Resume Next
    ActiveSheet.Pictures.Delete

    If Range("His_LifeNo") <> "" Then
        InsertPStaffPic Range("His_LifeNo"), Range("His_Pic_M")
    End If
    
    If Not (Range("His_LifeNo_Spouse") = "" Or Range("His_LifeNo_Spouse") = "0") Then
        InsertPStaffPic Range("His_LifeNo_Spouse"), Range("His_Pic_F")
    End If
    
    '--//�������� ��¥���� �߰� �� �����Ͽ� ��Ʋ���� ����
    InsertPStaffPic "", Range("Z9")
    If ActiveSheet.Pictures.Count > 0 Then
        ActiveSheet.Pictures(ActiveSheet.Pictures.Count).Delete
    End If
On Error GoTo 0

    Sheets("�����ڿ���").Range("C1").Select
    
    Call shProtect(globalSheetPW)
    
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
        strSql = "SELECT a.`�����ȣ`,a.`��ȸ��`,a.`�ѱ��̸�(����)`,a.`��å` " & _
                    "FROM " & TB1 & " a " & _
                    "WHERE (a.`�ѱ��̸�(����)` LIKE '%" & Me.txtChurchNM & "%' OR a.`��ȸ��` LIKE '%" & Me.txtChurchNM & "%'" & _
                    " OR a.`�����̸�` LIKE '%" & Me.txtChurchNM & "%' OR a.`����ȸ��` LIKE '%" & Me.txtChurchNM & "%'" & _
                    " OR a.`������ȸ��` LIKE '%" & Me.txtChurchNM & "%' OR a.`��������ȸ��` LIKE '%" & Me.txtChurchNM & "%')" & _
                    " AND a.`�����μ�` = " & SText(USER_DEPT) & ";"

    Case TB2
        strSql = "SELECT * FROM op_system.v_transfer_history a WHERE a.`�����ȣ` = " & SText(Me.lstPStaff.List(Me.lstPStaff.listIndex)) & ";"
    Case TB3
        If Range("T4") = 0 Then
            
            strSql = "SELECT DISTINCT a.family_cd FROM " & TB3 & " a WHERE a.lifeno = " & SText(Range("J4").Value) & " AND a.relations NOT IN (""��"",""��"");"
            Call makeListData(strSql, TB3)
            
            If cntRecord = 1 Then
                Range("AC26") = LISTDATA(0, 0)
                strSql = "SELECT * FROM (SELECT a.family_id,a.family_cd,REPLACE(a.relations,'(����)','') AS 'relations',a.lifeno,a.name_ko,a.name_en,a.church_sid,a.church_nm,a.title,a.POSITION,a.birthday,a.education,a.religion,a.recognition,a.memo,a.SUSPEND,CONCAT(a.title,if(a.title<>'' AND a.position<>'','/',''),a.position) AS 'status',RANK() OVER (PARTITION BY a.family_cd,relations ORDER BY a.birthday) AS 'rank' FROM op_system.v_familyinfo a WHERE a.family_cd = " & SText(LISTDATA(0, 0)) & " ORDER BY family_cd,FIELD(relations,'��','��','����','�ڸ�'),birthday) a WHERE a.lifeno <> " & SText(Range("J4").Value) & ";"
            ElseIf cntRecord > 1 Then
                MsgBox "������ �������� �����Ϳ� �ߺ������� �ֽ��ϴ�. �ߺ��� �ڷḦ �����ϼ���.", vbCritical, banner
            End If
        Else
            strSql = "SELECT DISTINCT a.family_cd FROM " & TB3 & " a WHERE a.lifeno = " & SText(Range("J4").Value) & " AND a.relations NOT IN (""��"",""��"")"
            Call makeListData(strSql, TB3)
            
            If cntRecord = 0 Then
                strSql = "SELECT DISTINCT a.family_cd FROM " & TB3 & " a WHERE a.lifeno = " & SText(Range("T4").Value) & " AND a.relations NOT IN (""��"",""��"");"
                Call makeListData(strSql, TB3)
                
                If cntRecord = 1 Then
                    Range("AC26") = LISTDATA(0, 0)
                    strSql = "SELECT * FROM (SELECT a.family_id,a.family_cd,REPLACE(a.relations,'(����)','') AS 'relations',a.lifeno,a.name_ko,a.name_en,a.church_sid,a.church_nm,a.title,a.POSITION,a.birthday,a.education,a.religion,a.recognition,a.memo,a.SUSPEND,CONCAT(a.title,if(a.title<>'' AND a.position<>'','/',''),a.position) AS 'status',RANK() OVER (PARTITION BY a.family_cd,relations ORDER BY a.birthday) AS 'rank' FROM op_system.v_familyinfo a WHERE a.family_cd = " & SText(LISTDATA(0, 0)) & " ORDER BY family_cd,FIELD(relations,'��','��','����','�ڸ�'),birthday) a WHERE a.lifeno NOT IN (" & SText(Range("T4").Value) & ");"
                End If
            ElseIf cntRecord = 1 Then
                strSql = "SELECT DISTINCT a.family_cd FROM " & TB3 & " a WHERE a.lifeno = " & SText(Range("T4").Value) & " AND a.relations NOT IN (""��"",""��"");"
                Call makeListData(strSql, TB3)
                
                If cntRecord = 1 Then
                    strSql = "SELECT DISTINCT a.family_cd FROM " & TB3 & " a WHERE a.lifeno = " & SText(Range("J4").Value) & " AND a.relations NOT IN (""��"",""��"")" & _
                                " UNION SELECT DISTINCT a.family_cd FROM " & TB3 & " a WHERE a.lifeno = " & SText(Range("T4").Value) & " AND a.relations NOT IN (""��"",""��"");"
                    Call makeListData(strSql, TB3)
                    
                    Range("AC26") = LISTDATA(0, 0)
                    Range("AD26") = LISTDATA(1, 0)
                    strSql = "SELECT * FROM (SELECT a.family_id,a.family_cd,REPLACE(a.relations,'(����)','') AS 'relations',a.lifeno,a.name_ko,a.name_en,a.church_sid,a.church_nm,a.title,a.POSITION,a.birthday,a.education,a.religion,a.recognition,a.memo,a.SUSPEND,CONCAT(a.title,if(a.title<>'' AND a.position<>'','/',''),a.position) AS 'status',RANK() OVER (PARTITION BY a.family_cd,relations ORDER BY a.birthday) AS 'rank' FROM op_system.v_familyinfo a WHERE a.family_cd = " & SText(LISTDATA(0, 0)) & _
                            " UNION SELECT a.family_id,a.family_cd,REPLACE(a.relations,'(����)','') AS 'relations',a.lifeno,a.name_ko,a.name_en,a.church_sid,a.church_nm,a.title,a.POSITION,a.birthday,a.education,a.religion,a.recognition,a.memo,a.SUSPEND,CONCAT(a.title,if(a.title<>'' AND a.position<>'','/',''),a.position) AS 'status',RANK() OVER (PARTITION BY a.family_cd,relations ORDER BY a.birthday) AS 'rank' FROM op_system.v_familyinfo a WHERE a.family_cd = " & SText(LISTDATA(1, 0)) & " ORDER BY family_cd,FIELD(relations,'��','��','����','�ڸ�'),birthday) a WHERE a.lifeno NOT IN (" & SText(Range("J4").Value) & "," & SText(Range("T4").Value) & ");"
                ElseIf cntRecord = 0 Then
                    Range("AC26") = LISTDATA(0, 0)
                    strSql = "SELECT * FROM (SELECT a.family_id,a.family_cd,REPLACE(a.relations,'(����)','') AS 'relations',a.lifeno,a.name_ko,a.name_en,a.church_sid,a.church_nm,a.title,a.POSITION,a.birthday,a.education,a.religion,a.recognition,a.memo,a.SUSPEND,CONCAT(a.title,if(a.title<>'' AND a.position<>'','/',''),a.position) AS 'status',RANK() OVER (PARTITION BY a.family_cd,relations ORDER BY a.birthday) AS 'rank' FROM op_system.v_familyinfo a WHERE a.family_cd = " & SText(LISTDATA(0, 0)) & " ORDER BY family_cd,FIELD(relations,'��','��','����','�ڸ�'),birthday) a WHERE a.lifeno NOT IN (" & SText(Range("J4").Value) & ");"
                End If
            ElseIf cntRecord > 2 Then
                MsgBox "������ Ȥ�� ��� �������� �����Ϳ� �ߺ������� �ֽ��ϴ�. �ߺ��� �ڷḦ �����ϼ���.", vbCritical, banner
            End If
        End If
        
        strSql = strSql & ";"
    Case Else
        '//ȭ��id, ȭ���Ī, ��ȸ��, ��ȭȯ��, �޷�ȭȯ��
        'strSQL = "SELECT a.currency_id, a.currency_un, a.refer_dt, a.fx_rate_krw, a.fx_rate_usd " & _
                      "FROM " & TB1 & " a WHERE a.user_id = " & user_id & ";"
    End Select
    makeSelectSQL = strSql
End Function
Private Sub sbClearVariant()
    Erase LISTFIELD
    Erase LISTDATA
    cntRecord = Empty
    strSql = vbNullString
End Sub

