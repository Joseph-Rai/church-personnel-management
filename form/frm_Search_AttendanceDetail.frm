VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Search_AttendanceDetail 
   Caption         =   "��ȸ �˻�"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7365
   OleObjectBlob   =   "frm_Search_AttendanceDetail.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "frm_Search_AttendanceDetail"
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

Private Sub chkAllListItems_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    SelectAllItems Me.lstChurch, Not Me.chkAllListItems.Value
End Sub

Private Sub lstChurch_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Me.optEntireMC.Value = True And Me.lstChurch.ListCount > 0 Then
        If CountSelectedItems(Me.lstChurch) = Me.lstChurch.ListCount Then
            Me.chkAllListItems.Value = True
        Else
            Me.chkAllListItems.Value = False
        End If
    End If
End Sub

Private Sub optEntireMC_Click()
    RelativeControlDisable
    
    strSql = "" & _
        "SELECT " & _
        "    DISTINCT Null " & _
        "    ,geo.country_nm_ko " & _
        "    ,NULL,NULL,NULL " & _
        "FROM op_system.db_geodata geo " & _
        "INNER JOIN op_system.db_ovs_dept dept " & _
        "    ON geo.division = dept.dept_nm " & _
        "INNER JOIN op_system.db_churchlist churchlist " & _
        "    ON geo.geo_cd = churchlist.geo_cd " & _
        "        AND churchlist.church_gb IN ('MC', 'HBC1', 'HBC2') " & _
        "WHERE Dept.DEPT_ID = " & SText(USER_DEPT) & ";"
    
    Call makeListData(strSql, "����ȸ ������ ��������Ʈ")
    
    Me.lstChurch.MultiSelect = fmMultiSelectMulti
    Me.lstChurch.Clear
    Me.lstChurch.List = LISTDATA
    
End Sub

Private Sub lstChurch_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Un'HookListBoxScroll
End Sub

Private Sub lstChurch_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Me.lstChurch.ListCount > 0 Then
        'HookListBoxScroll Me, Me.lstChurch
    End If
End Sub

Private Sub optMcBcPbc_Click()
    RelativeControlDisable
    
    Me.lstChurch.MultiSelect = fmMultiSelectSingle
    Me.lstChurch.Clear
    If Me.txtChurch <> "" Then
        cmdSearch_Click
    End If
    
End Sub

Private Sub RelativeControlDisable()

    TextBoxEnable Me.txtChurch, Not Me.optEntireMC.Value
    ControlEnable Me.cmdSearch, Not Me.optEntireMC.Value
    ControlEnable Me.chkAll, Not Me.optEntireMC.Value
    ControlEnable Me.frameSelectOption, Not Me.optEntireMC.Value
    ControlEnable Me.chkMM, Not Me.optEntireMC.Value
    ControlEnable Me.chkBC, Not Me.optEntireMC.Value
    ControlEnable Me.chkPBC, Not Me.optEntireMC.Value
    
    Me.chkMM.Value = Not Me.optEntireMC.Value
    Me.chkBC.Value = Not Me.optEntireMC.Value
    Me.chkPBC.Value = Not Me.optEntireMC.Value
    Me.chkAllListItems.Visible = Me.optEntireMC.Value
    Me.chkAllListItems.Value = Not Me.optEntireMC.Value
    
    If Me.optEntireMC.Value = True Then
        Me.lblChurchName.Caption = "������"
        Me.lblChurchClass.Caption = ""
    Else
        Me.lblChurchName.Caption = "��ȸ��"
        Me.lblChurchClass.Caption = "��ȸ����"
    End If

End Sub

Private Sub ListBoxFormatInitialize()

    With Me.lstChurch
        .ColumnCount = 4
        .ColumnHeads = False
        .ColumnWidths = "0,150,50,0,0" '��ȸ�ڵ�, ��ȸ��, ��ȸ����, ������ȸ��, Ŀ�����ڵ�
        .Width = 265.5
        .TextAlign = fmTextAlignLeft
        .Font = "����"
    End With

End Sub

Private Sub UserForm_Initialize()
    
    Dim intYear As Integer
    Dim intMonth As Integer
    
    If connIP = Empty Then setGlobalVariant '//�������� �缳��
    If checkLogin = 0 Then f_login.Show '//�α���üũ
    If checkLogin = 0 Then End '--//�α��� ���� �� ���ν��� ����
    
    '--//��Ʈ����
    Set ws = ActiveSheet
    
    '--//���ʼ���
    Me.cmdClose.Cancel = True
    TB1 = "op_system.v_atten_detail_churchlist" '--//��ȸ����Ʈ
    TB2 = "op_system.temp_atten_detail" '--//�⼮��Ȳ��
    TB4 = "op_system.temp_atten_detail_main" '--//����ȸ(MC) �⼮��Ȳ��
    
    '--//����Ʈ�ڽ� ����
    ListBoxFormatInitialize
    
    '--//��Ʈ�� �ʱ�ȭ
    Me.optMcBcPbc = True
    Me.chkMM = True
    Me.chkBC = True
    Me.chkPBC = True
    
    '--//�޺��ڽ� ������ �߰�
    For intYear = year(Date) To 2005 Step -1 '--//�⵵ ä���
        Me.cboYear.AddItem intYear
    Next
    For intMonth = 12 To 1 Step -1 '--//�� ä���
        Me.cboMonth.AddItem intMonth
    Next
    
    '--//�޺��ڽ� ���� ������ ��¥�� �����߰�
    If Range("AttenDetail_rngDate") = "" Then
        Me.cboYear = year(Date)
        Me.cboMonth = month(Date)
    Else
        Me.cboYear = year(WorksheetFunction.EDate(Date, -1))
        Me.cboMonth = month(WorksheetFunction.EDate(Date, -1))
    End If
    
    Me.txtChurch.SetFocus
End Sub

Private Sub cboYear_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Un'HookListBoxScroll
End Sub

Private Sub cboYear_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Me.cboYear.ListCount > 0 Then
        'HookListBoxScroll Me, Me.cboYear
    End If
End Sub

Private Sub cboMonth_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Un'HookListBoxScroll
End Sub

Private Sub cboMonth_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Me.cboMonth.ListCount > 0 Then
        'HookListBoxScroll Me, Me.cboMonth
    End If
End Sub

Private Sub cmdSearch_Click()
    strSql = makeSelectSQL(TB1)
    Call makeListData(strSql, TB1)
    If cntRecord > 0 Then
        Me.lstChurch.List = LISTDATA
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
    Unload Me
End Sub
Private Sub cmdPrint_Click()
    Dim printPages As Integer
    printPages = ActiveSheet.PageSetup.Pages.Count
    ActiveWindow.SelectedSheets.PrintOut FROM:=1, to:=printPages, Copies:=1
End Sub
Private Sub cmdOK_Click()
    
    Dim result As T_RESULT
    
    '--//��Ʈ Ȱ��ȭ �� �������
    WB_ORIGIN.Activate
    ws.Activate
    Call shUnprotect(globalSheetPW)
    
    '--//��ȸ ���ÿ��� �Ǵ�
    If Me.lstChurch.listIndex = -1 And Me.optEntireMC = False Then
        MsgBox "��ȸ�� �����ϼ���.", vbCritical, "����"
        Exit Sub
    End If
    
    '--//��¥ �Է¿���
    If Me.cboYear <> "" Or Me.cboMonth <> "" Then
        If Me.cboYear = "" And Me.cboMonth <> "" Then
            MsgBox "�˻� ���� �Է��� �ּ���.", vbCritical, banner
            Exit Sub
        End If
        
        If Me.cboYear <> "" And Me.cboMonth = "" Then
            MsgBox "�˻� ���� �Է��� �ּ���.", vbCritical, banner
            Exit Sub
        End If
        
        Range("AttenDetail_rngDate") = DateSerial(Me.cboYear, Me.cboMonth, 1)
    End If
    
    '--//������ �ʱ�ȭ
    initCurrentPage
    
    Call Optimization
    
    '--//�⼮������ ����
    If Range("AttenDetail_rngTarget").Cells(1, 1) <> "" Then
        Range("AttenDetail_rngTarget").CurrentRegion.ClearContents
    End If
    
    Dim tableName As String
    With Me.lstChurch
        If Me.optEntireMC = True Then
            strSql = "CALL `Routine_atten_detail_main`(" & SText(Format(WorksheetFunction.EoMonth(DateSerial(Me.cboYear, Me.cboMonth, 1), 0), "yyyy-mm-dd")) & ", " & _
                        SText(USER_ID) & ")"
            tableName = TB4 '--//��ü ��ȸ(MC) �⼮��ȸ
        Else
            strSql = "CALL `Routine_atten_detail`(" & SText(Format(WorksheetFunction.EoMonth(DateSerial(Me.cboYear, Me.cboMonth, 1), 0), "yyyy-mm-dd")) & ", " & _
                        SText(.List(.listIndex, 0)) & ", " & SText(USER_ID) & ");"
            tableName = TB2 '--//MC, MM, BC, PBC �⼮��ȸ
        End If
    End With
    connectTaskDB
    result.strSql = strSql
    result.affectedCount = executeSQL("cmdOK_Click", tableName, strSql, Me.Name, "�ӽ���ȸ ���̺� ����")
    disconnectALL
    
    strSql = makeSelectSQL(tableName)
    Call makeListData(strSql, tableName)

    Dim rngDataColumns As Integer
    rngDataColumns = UBound(LISTFIELD) + 1
    Range("AttenDetail_rngTarget").Resize(, rngDataColumns) = LISTFIELD
    If cntRecord > 0 Then
        Range("AttenDetail_rngTarget").Offset(1).Resize(cntRecord, rngDataColumns) = LISTDATA
    End If
    
    '--//������ ���� ��,�� �� ����
    Range("AttenDetail_rngTarget_Rows") = Range("AttenDetail_rngTarget").CurrentRegion.Rows.Count
    Range("AttenDetail_rngTarget_Columns") = Range("AttenDetail_rngTarget").CurrentRegion.Columns.Count
    If Me.optMcBcPbc.Value = True Then
        Range("attenDetail_SearchOption") = 1
    Else
        Range("attenDetail_SearchOption") = 2
    End If
    
    '--//�⼮������ ���ĺ�ȯ(string -> int)
    Range("A1").Copy
    Range("AttenDetail_rngTarget").Offset(1).Resize(cntRecord, rngDataColumns).PasteSpecial Paste:=xlPasteValues, operation:=xlPasteSpecialOperationAdd
    Application.CutCopyMode = False
    Call sbClearVariant
    
    If Range("AttenDetail_rngDate") > Range("AttenDetail_MaxDate") And IsDate(Range("AttenDetail_MaxDate")) Then
        Range("AttenDetail_rngDate") = Range("AttenDetail_MaxDate")
    End If
    
    '--//��ȸ, ����ȸ, ����� ����Ʈ ����
    Range("AttenDetail_ChurchList").CurrentRegion.ClearContents
    strSql = makeSelectSQL(TB3)
    Call makeListData(strSql, TB3)
    
    Range("AttenDetail_ChurchList").Resize(, 4) = LISTFIELD
    If cntRecord > 0 Then
        Range("AttenDetail_ChurchList").Offset(1).Resize(cntRecord, 4) = LISTDATA
    End If
    Range("AttenDetail_ChurchCount") = cntRecord
    
    Range("AttenDetail_ChurchList").Offset(, 4) = "DateMinNonZero"
    Range("AttenDetail_ChurchList").Offset(1, 4).Resize(cntRecord).FormulaR1C1 = _
        "=IFERROR(" & Chr(10) & "    EDATE(" & Chr(10) & "        LOOKUP(" & Chr(10) & "            0," & Chr(10) & _
        "            OFFSET(AttenDetail_rngTarget,MATCH(RC25,OFFSET(AttenDetail_rngTarget,1,,AttenDetail_rngTarget_Rows,1),0),MATCH(""�⼮(��ü 1ȸ)"",OFFSET(AttenDetail_rngTarget,,,1,AttenDetail_rngTarget_Columns),0)-1,COUNTIF(OFFSET(AttenDetail_rngTarget,,,AttenDetail_rngTarget_Rows,1),RC25),1)," & Chr(10) & _
        "            OFFSET(AttenDetail_rngTarget,MATCH(RC25,OFFSET(AttenDetail_rngTarget,1,,AttenDetail_rngTarget_Rows,1),0),MATCH(""attendance_dt"",OFFSET(AttenDetail_rngTarget,,,1,AttenDetail_rngTarget_Columns),0)-1,COUNTIF(OFFSET(AttenDetail_rngTarget,,,AttenDetail_rngTarget_Rows,1),RC25),1)" & Chr(10) & "        )," & Chr(10) & "    1)," & Chr(10) & _
        "    INDEX(OFFSET(AttenDetail_rngTarget,MATCH(RC25,OFFSET(AttenDetail_rngTarget,1,,AttenDetail_rngTarget_Rows,1),0),MATCH(""attendance_dt"",OFFSET(AttenDetail_rngTarget,,,1,AttenDetail_rngTarget_Columns),0)-1,COUNTIF(OFFSET(AttenDetail_rngTarget,,,AttenDetail_rngTarget_Rows,1),RC25),1),1)" & Chr(10) & ")"
    
    Range("A1").Copy
    Range("AttenDetail_ChurchList").Offset(1).Resize(cntRecord, 4).PasteSpecial Paste:=xlPasteValues, operation:=xlPasteSpecialOperationAdd
    Application.CutCopyMode = False
    Call sbClearVariant
    
On Error Resume Next
    '--//��ȸ ������ ���� ������ ����
    makeReportPage
    
    '--//��������
    attenDetailInsertPicture
On Error GoTo 0
    
    Call Normal
    
    Range("A2").Select
    
    Call shProtect(globalSheetPW)
    
End Sub
Public Sub initCurrentPage()

    Range(Range("17:17"), Range("17:17").End(xlDown)).Delete

End Sub
Private Sub makeReportPage()

    Dim basicTableRange As Range
    Dim basicRows As Integer
    Dim countChurch As Integer
    
    Set basicTableRange = Range("AttenDetail_BasicTableRange")
    basicRows = basicTableRange.Rows.Count
    countChurch = Range("AttenDetail_ChurchCount")
    
    basicTableRange.Copy
    basicTableRange.Offset(basicRows).Resize(basicRows * (countChurch - 1)).PasteSpecial Paste:=xlPasteAll
    
    '--//����Ʈ ���� ����
    Dim printRange As Range
    Set printRange = basicTableRange.Offset(, 1).Resize(, basicTableRange.Columns.Count - 1)
    ActiveSheet.PageSetup.PrintArea = Range(printRange.Offset(-2).Resize(printRange.Rows.Count + 2), printRange.Offset(basicRows).Resize(basicRows * (countChurch - 1))).Address
    If ActiveSheet.VPageBreaks.Count > 0 Then
        ActiveSheet.VPageBreaks(1).DragOff Direction:=xlToRight, RegionIndex:=1
    End If
    
    Dim index As Integer
    Dim pageNum As Integer
    For index = 17 To Cells(Rows.Count, "B").End(xlUp).Row
        If (index - 5) Mod 12 = 0 Then
            If (Cells(index, "B") - 1) Mod 4 = 0 Then
                pageNum = pageNum + 1
                Set ActiveSheet.HPageBreaks(pageNum).Location = Cells(index, "C")
            End If
        End If
    Next index
    
    '--//ȭ�� ������Ʈ �������� �ʵ��� 1�� ��ٸ�
    Application.ScreenUpdating = True
    Application.Wait Now + TimeValue("00:00:01")
    
End Sub

Public Sub attenDetailInsertPicture()

    Dim index As Integer
    Dim countChurch As Integer
    Dim FileName As String
    Dim filePath As String

    ActiveSheet.Pictures.Delete
    
    filePath = fnFindPicPath
    countChurch = Range("AttenDetail_ChurchCount")
    
'    If filePath = "" Then
'        MsgBox "���� ������Ʈ �����Դϴ�. ���̵�ũ ������ Ȯ���� �ּ���.", vbCritical, "���� ������Ʈ ����"
'        Exit Sub
'    End If
On Error Resume Next
    Dim basicRows As Integer
    basicRows = Range("AttenDetail_BasicTableRange").Rows.Count
    For index = 9 To basicRows * (countChurch - 1) + 9 Step basicRows
        
        InsertPStaffPic Cells(index, "C"), Cells(index, "C").Resize(5)
        
        If Not (Cells(index, "D") = "" Or Cells(index, "D") = 0) Then
            InsertPStaffPic Cells(index, "D"), Cells(index, "D").Resize(5)
        End If
    Next index
    
    '--//�������� ��¥���� �ϳ� �߰� �� �����Ͽ� ��Ʋ���� ����
    FileName = fnFindRepresentativePic
    InsertPStaffPic "", Range("A2")
'    Call sbInsertPicture2_WIS(fileName, Range("A2"))
    If ActiveSheet.Pictures.Count > 0 Then
        ActiveSheet.Pictures(ActiveSheet.Pictures.Count).Delete
    End If
On Error GoTo 0

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
    
    Dim selectedCountryList As String
    Dim selectedChurchClassList As String
    
    Select Case tableNM
    Case TB1
        '--//��ȸ�ڵ�, ��ȸ��
        If Me.chkAll Then
            strSql = "SELECT a.church_sid,a.church_nm,a.church_gb,null,a.church_sid_custom " & _
                        "FROM " & TB1 & " a " & _
                        "WHERE a.ovs_dept = " & USER_DEPT & " AND a.church_nm LIKE '%" & Me.txtChurch & "%';"
        Else
            strSql = "SELECT a.church_sid,a.church_nm,a.church_gb,null,a.church_sid_custom " & _
                        "FROM " & TB1 & " a " & _
                        "WHERE a.ovs_dept = " & USER_DEPT & " AND a.end_dt = '9999-12-31' AND a.church_nm LIKE '%" & Me.txtChurch & "%';"
        End If
    
    Case TB2
        selectedChurchClassList = GetSelectedChurchClassArrayString
        
        '--//�⼮ ������
        strSql = "SELECT * FROM " & TB2 & " a WHERE user_id = " & SText(USER_ID)
        strSql = strSql & _
                " AND a.church_gb IN (" & selectedChurchClassList & ");"
        
    Case TB3
        '--//����ȸ, ����ȸ, ����� ���
        Dim tableName As String
        If Me.optEntireMC = True Then
            tableName = TB4
            GetSelectedCountryArrayString
        Else
            tableName = TB2
        End If
        
        strSql = "SELECT " & _
                        "DISTINCT atten.CHURCH_SID_CUSTOM " & _
                        ",atten.church_sid " & _
                        ",atten.church_nm " & _
                        ",atten.church_gb " & _
                    "FROM " & tableName & " atten " & _
                    "WHERE atten.user_id = " & SText(USER_ID) & " "
        If Me.optEntireMC = True Then
            strSql = strSql & _
                    "    AND atten.country IN (" & GetSelectedCountryArrayString & ") "
        ElseIf Me.optMcBcPbc = True Then
            strSql = strSql & _
                    "    AND atten.church_gb IN (" & GetSelectedChurchClassArrayString & ") "
        End If
        
        strSql = strSql & _
            "ORDER BY FIELD(atten.church_gb,'MC','HBC','MM','BC','PBC')"
    Case TB4
        '--//��ü ��ȸ(MC) �⼮ ������
        strSql = "SELECT * FROM " & TB4 & " a WHERE user_id = " & SText(USER_ID)
        strSql = strSql & _
                    "    AND a.country IN (" & GetSelectedCountryArrayString & ") "
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

Function GetSelectedCountryArrayString()

    Dim i As Integer
    Dim result As String
    With Me.lstChurch
        For i = 0 To .ListCount - 1
            If .Selected(i) = True Then
                If result = "" Then
                    result = SText(.List(i, 1))
                Else
                    result = result & "," & SText(.List(i, 1))
                End If
            End If
        Next
    End With
    
    GetSelectedCountryArrayString = result

End Function

Function GetSelectedChurchClassArrayString()

    Dim result As String
    
    If result = "" Then
        result = SText("MC")
    Else
        result = result & SText("MC")
    End If
    
    If result = "" Then
        result = SText("HBC")
    Else
        result = result & "," & SText("HBC")
    End If
    
    If Me.chkMM.Value = True Then
        If result = "" Then
            result = SText("MM")
        Else
            result = result & "," & SText("MM")
        End If
    End If
    
    If Me.chkBC.Value = True Then
        If result = "" Then
            result = SText("BC")
        Else
            result = result & "," & SText("BC")
        End If
    End If
    
    If Me.chkPBC.Value = True Then
        If result = "" Then
            result = SText("PBC")
        Else
            result = result & "," & SText("PBC")
        End If
    End If
    
    GetSelectedChurchClassArrayString = result

End Function




