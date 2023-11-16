Attribute VB_Name = "sb_EXcelToDB"
Option Explicit

Sub insertDataToDB_Everything()
    
    Dim shtNM As String, tableNM As String, dbNM As String, strSql As String
    Dim affectedCount As Long
    Dim Values() As String
    Dim cntField As Integer, cntRecord As Long, i As Long, j As Long, k As Long
    
    '//Sheet��, Table��, db�� ����
    shtNM = "sheet1" '//�����ڡ�
    tableNM = "op_system.db_division" '//�����ڡ�
    dbNM = "op_system" '//�����ڡ�
    
    '//DB����
    Call connectTaskDB
    
    '//Table �ʱ�ȭ
'    strSQL = "TRUNCATE TABLE " & tableNM
'    affectedCount = executeSQL("insertDataToDB_Everything", tableNM, strSQL, "������ �ʱ�ȭ")
    
    '/�迭 ũ�� ����
    Sheets(shtNM).Activate
    cntField = Range("A1").CurrentRegion.Columns.Count
    cntRecord = Range("A1").CurrentRegion.Rows.Count - 1
    ReDim Values(cntRecord - 1, cntField - 1)
   
    '//DB�� �̵��� �ڷ� Values �迭�� ����
    For i = 0 To cntField - 1
        For j = 0 To cntRecord - 1
            Values(j, i) = Range("A2").Offset(j, i)
        Next j
    Next i
    
    '//���ڵ庰�� Insert(�����ڡ�)
    '  - �ʵ� �� ��ŭ �߰��ؼ� ����
    '  - ������ NULL���� �ְ��� �� ��� "''"�� "NULL"�� ����
    '  - �⺻���� �������� "Default" & "," & _
    '  - TimeStamp�� ��������: "CURRENT_TIMESTAMP()" & ");"
    For i = 0 To cntRecord - 1
        strSql = "INSERT INTO " & tableNM & " VALUES(" & _
                       convertStrToSQL(Values(i, 0)) & "," & _
                       convertStrToSQL(Values(i, 1)) & "," & _
                       convertStrToSQL(Values(i, 2)) & "," & _
                       convertStrToSQL(Values(i, 3)) & ");"
'        strSQL = "UPDATE " & tableNM & " a SET " & _
'                       "a.score_avg = " & convertStrToSQL(Values(i, 1)) & "," & _
'                       "a.subject_count = " & convertStrToSQL(Values(i, 2)) & _
'                        " WHERE a.lifeno = " & convertStrToSQL(Values(i, 0)) & ";"
        affectedCount = executeSQL("insertDataToDB_ChurchDivision", tableNM, strSql, , "���� ������ �о�ֱ�")
        k = k + affectedCount
    Next i
    MsgBox k & "���� ���ڵ尡 '" & tableNM & "'�� �߰��Ǿ����ϴ�.", vbInformation, banner
    
    '//���� ����
    disconnectDB
End Sub
'---------------------------------------------------
'  ���� �ڷḦ DB�� Insert
'    - Table �ʱ�ȭ
'    - ���� ���ڵ� �ϳ��� Insert
'----------------------------------------------------
Sub insertDataToDB_BranchList_Admin()
    
    Dim shtNM As String, tableNM As String, dbNM As String, strSql As String
    Dim affectedCount As Long
    Dim Values() As String
    Dim cntField As Integer, cntRecord As Long, i As Long, j As Long, k As Integer
    
    '//Sheet��, Table��, db�� ����
    shtNM = "Sheet1" '//�����ڡ�
    tableNM = "op_system.a_branch_admin" '//�����ڡ�
    dbNM = "op_system" '//�����ڡ�
    
    '//DB����
    Call connectTaskDB
    
    '//Table �ʱ�ȭ
    strSql = "TRUNCATE TABLE " & tableNM
    affectedCount = executeSQL("insertDataToDB_BranchList_Admin", tableNM, strSql, "������ �ʱ�ȭ")
    
    '/�迭 ũ�� ����
    Sheets(shtNM).Activate
    cntField = Range("A1").CurrentRegion.Columns.Count
    cntRecord = Range("A1").CurrentRegion.Rows.Count - 1
    ReDim Values(cntRecord - 1, cntField - 1)
   
    '//DB�� �̵��� �ڷ� Values �迭�� ����
    For i = 0 To cntField - 1
        For j = 0 To cntRecord - 1
            Values(j, i) = Range("A2").Offset(j, i)
        Next j
    Next i
    
    '//���ڵ庰�� Insert(�����ڡ�)
    '  - �ʵ� �� ��ŭ �߰��ؼ� ����
    '  - ������ NULL���� �ְ��� �� ��� "''"�� "NULL"�� ����
    '  - �⺻���� �������� "Default" & "," & _
    '  - TimeStamp�� ��������: "CURRENT_TIMESTAMP()" & ");"
    For i = 0 To cntRecord - 1
        strSql = "INSERT INTO " & tableNM & " VALUES(" & _
                       convertStrToSQL(Values(i, 0)) & "," & convertStrToSQL(Values(i, 1)) & "," & convertStrToSQL(Values(i, 2)) & "," & _
                       convertStrToSQL(Values(i, 3)) & "," & convertStrToSQL(Values(i, 4)) & "," & convertStrToSQL(Values(i, 5)) & "," & _
                       convertStrToSQL(Values(i, 6)) & "," & convertStrToSQL(Values(i, 7)) & "," & convertStrToSQL(Values(i, 8)) & "," & _
                       convertStrToSQL(Values(i, 9)) & "," & convertStrToSQL(Values(i, 10)) & "," & IIf(Values(i, 11) = "", "Null", convertStrToSQL(Values(i, 11))) & "," & _
                       IIf(Values(i, 12) = "", "Null", convertStrToSQL(Values(i, 12))) & "," & convertStrToSQL(Values(i, 13)) & "," & _
                       convertStrToSQL(Values(i, 14)) & "," & convertStrToSQL(Values(i, 15)) & "," & _
                       convertStrToSQL(Values(i, 16)) & "," & convertStrToSQL(Values(i, 17)) & "," & convertStrToSQL(Values(i, 18)) & "," & _
                       convertStrToSQL(Values(i, 19)) & "," & IIf(Values(i, 20) = "", "Null", convertStrToSQL(Values(i, 20))) & "," & _
                       convertStrToSQL(Values(i, 21)) & "," & convertStrToSQL(Values(i, 22)) & "," & _
                       convertStrToSQL(Values(i, 23)) & "," & convertStrToSQL(Values(i, 24)) & "," & _
                       convertStrToSQL(Values(i, 25)) & "," & _
                       convertStrToSQL(Values(i, 26)) & "," & convertStrToSQL(Values(i, 27)) & "," & IIf(Values(i, 28) = "", "Null", convertStrToSQL(Values(i, 28))) & "," & _
                       convertStrToSQL(Values(i, 29)) & "," & convertStrToSQL(Values(i, 30)) & "," & convertStrToSQL(Values(i, 31)) & "," & _
                       convertStrToSQL(Values(i, 32)) & ");"
        affectedCount = executeSQL("insertDataToDB_BranchList_Admin", tableNM, strSql, , "���� ������ �о�ֱ�")
        k = k + affectedCount
    Next i
    MsgBox k & "���� ���ڵ尡 '" & tableNM & "'�� �߰��Ǿ����ϴ�.", vbInformation, banner
    
    '//���� ����
    disconnectDB
End Sub
Sub insertDataToDB_ChurchList_Admin()
    
    Dim shtNM As String, tableNM As String, dbNM As String, strSql As String
    Dim affectedCount As Long
    Dim Values() As String
    Dim cntField As Integer, cntRecord As Long, i As Long, j As Long, k As Integer
    
    '//Sheet��, Table��, db�� ����
    shtNM = "Sheet1" '//�����ڡ�
    tableNM = "op_system.a_churchlist_admin" '//�����ڡ�
    dbNM = "op_system" '//�����ڡ�
    
    '//DB����
    Call connectTaskDB
    
    '//Table �ʱ�ȭ
    strSql = "TRUNCATE TABLE " & tableNM & ";"
'    affectedCount = executeSQL("insertDataToDB_ChurchList_Admin", tableNM, strSql, "������ �ʱ�ȭ")
    
    '/�迭 ũ�� ����
    Sheets(shtNM).Activate
    cntField = Range("A1").CurrentRegion.Columns.Count
    cntRecord = Range("A1").CurrentRegion.Rows.Count - 1
    ReDim Values(cntRecord - 1, cntField - 1)
   
    '//DB�� �̵��� �ڷ� Values �迭�� ����
    For i = 0 To cntField - 1
        For j = 0 To cntRecord - 1
            Values(j, i) = Range("A2").Offset(j, i)
        Next j
    Next i
    
    '//���ڵ庰�� Insert(�����ڡ�)
    '  - �ʵ� �� ��ŭ �߰��ؼ� ����
    '  - ������ NULL���� �ְ��� �� ��� "''"�� "NULL"�� ����
    '  - �⺻���� �������� "Default" & "," & _
    '  - TimeStamp�� ��������: "CURRENT_TIMESTAMP()" & ");"
    For i = 0 To cntRecord - 1
        strSql = strSql & "INSERT INTO " & tableNM & " VALUES(" & _
                       convertStrToSQL(Values(i, 0)) & "," & convertStrToSQL(Values(i, 1)) & "," & convertStrToSQL(Values(i, 2)) & "," & _
                       convertStrToSQL(Values(i, 3)) & "," & convertStrToSQL(Values(i, 4)) & "," & convertStrToSQL(Values(i, 5)) & "," & _
                       convertStrToSQL(Values(i, 6)) & "," & convertStrToSQL(Values(i, 7)) & "," & convertStrToSQL(Values(i, 8)) & "," & _
                       convertStrToSQL(Values(i, 9)) & "," & convertStrToSQL(Values(i, 10)) & "," & convertStrToSQL(Values(i, 11)) & "," & _
                       convertStrToSQL(Values(i, 12)) & "," & convertStrToSQL(Values(i, 13)) & "," & _
                       convertStrToSQL(Values(i, 14)) & "," & convertStrToSQL(Values(i, 15)) & "," & _
                       convertStrToSQL(Values(i, 16)) & "," & convertStrToSQL(Values(i, 17)) & "," & _
                       convertStrToSQL(Values(i, 18)) & "," & convertStrToSQL(Values(i, 19)) & "," & _
                       convertStrToSQL(Values(i, 20)) & "," & convertStrToSQL(Values(i, 21)) & "," & _
                       convertStrToSQL(Values(i, 22)) & "," & convertStrToSQL(Values(i, 23)) & "," & _
                       convertStrToSQL(Values(i, 24)) & "," & convertStrToSQL(Values(i, 25)) & ");"
        k = k + 1
        If k Mod 1000 = 0 Then
            affectedCount = executeSQL("insertDataToDB_ChurchList_Admin", tableNM, strSql, , "���� ������ �о�ֱ�")
            strSql = vbNullString
        End If
    Next i
    
    If strSql <> vbNullString Then
        affectedCount = executeSQL("insertDataToDB_ChurchList_Admin", tableNM, strSql, , "���� ������ �о�ֱ�")
    End If
    
    MsgBox k & "���� ���ڵ尡 '" & tableNM & "'�� �߰��Ǿ����ϴ�.", vbInformation, banner
    
    '//���� ����
    disconnectDB
End Sub

Sub UpdateAttendanceRecent12Month()

    Dim maxRowNum As Long
    maxRowNum = Cells(Rows.Count, "F").End(xlUp).Row
    
    '--//��ȸ�ڵ� ã�Ƽ� �ֱ����� LIst
    Dim churchList As Object
    Dim vChurchDao As New ChurchDao
    Set churchList = vChurchDao.getAllChurchListReverseOrder
    Dim churchCode As String
    
    Dim index As Long
    
    Application.ScreenUpdating = False
    
    '--//��ȸ�� ����
    For index = 2 To maxRowNum Step 16
        Do While Cells(index, "H") = "" And Cells(index, "F") <> ""
            Cells(index, "F").Resize(16).EntireRow.Delete Shift:=xlUp
            maxRowNum = maxRowNum - 16
            index = WorksheetFunction.Max(2, index - 16)
        Loop
        
        If Cells(index, "H") = "����ȸ" Or Cells(index, "H") = "��ȸ��������ȸ" Then
            If InStr(Cells(index, "F"), "����ȸ") <= 0 Then
                Cells(index, "F").Resize(16) = Cells(index, "F") & " ����ȸ"
            End If
        End If
        
        If InStr(Cells(index, "F"), "��ü") > 0 Then
            Cells(index, "F").Resize(16) = Replace(Cells(index, "F"), " ��ü", "")
        End If
        
        If Not (Left(Cells(index, "E"), 2) = "MC" Or Left(Cells(index, "E"), 2) = "BC") Then
            churchCode = getChurchCodeinList(churchList, Cells(index, "F"))
            Cells(index, "E").Resize(16) = churchCode
        End If
    Next
    
    '--//��ü, ����ȸ ��Ʈ�� �ȵǾ� �ִ� �� ����
    Dim targetRange As Range
    index = 2
    Do While index <= maxRowNum
        If InStr(Cells(index, "F"), "����ȸ") > 0 Then
            If WorksheetFunction.CountIf(Range("F:F"), Replace(Cells(index, "F"), " ����ȸ", "")) <= 0 Then
                Set targetRange = Cells(index, "F").Resize(16).EntireRow
                targetRange.Copy
                targetRange.Insert Shift:=xlDown
                Cells(index, "F").Resize(16) = Replace(Cells(index, "F"), " ����ȸ", "")
                Cells(index, "H").Resize(16) = "������ȸ"
                Cells(index, "E").Resize(16) = Replace(Cells(index, "E"), "MM", "MC")
                Application.CutCopyMode = False
                index = index + 16
                maxRowNum = maxRowNum + 16
            End If
        Else
            If Cells(index, "H") = "������ȸ" Then
                If WorksheetFunction.CountIf(Range("F:F"), Cells(index, "F") & " ����ȸ") <= 0 Then
                    Set targetRange = Cells(index, "F").Resize(16).EntireRow
                    targetRange.Copy
                    targetRange.Insert Shift:=xlDown
                    Cells(index, "F").Resize(16).Offset(16) = Cells(index, "F") & " ����ȸ"
                    Cells(index, "H").Resize(16).Offset(16) = "����ȸ"
                    Cells(index, "E").Resize(16).Offset(16) = Replace(Cells(index, "E"), "MC", "MM")
                    Application.CutCopyMode = False
                    index = index + 16
                    maxRowNum = maxRowNum + 16
                End If
            End If
        End If
        index = index + 1
    Loop
    
    Application.ScreenUpdating = True
    
    If Range("E1").End(xlDown).Row = maxRowNum Then
        Dim updateDate As Date
        Dim strUpdateDate As String
        strUpdateDate = Range("A1").End(xlToRight).Offset(, -1)
        updateDate = DateSerial(Left(strUpdateDate, 4), Right(strUpdateDate, 2), 1)
        ConvertAttendanceDataToTableFormat updateDate
    Else
        MsgBox "Some church code could not found in database." & vbCrLf & "Re-run this method after filling out missing code.", vbOKOnly, "Complete"
    End If
    
    SaveExcelFile ActiveWorkbook

End Sub

Sub SaveExcelFile(wb As Workbook)
    Dim filePath As String
    
    ' ���� �����ִ� ������ ���� ���� Ȯ��
    If wb.Path = "" Then
        ' ������� ���� ���, ���� ���̾�α� ����
        With Application.FileDialog(msoFileDialogSaveAs)
            .title = "���� ����"
            .Show
            If .SelectedItems.Count > 0 Then
                filePath = .SelectedItems(1)
                wb.SaveAs filePath
            End If
        End With
    Else
        ' �̹� ����� ���, ���� ��ο� �����ϱ�
        wb.Save
    End If
End Sub

Function getChurchCodeinList(ByRef churchList As Object, ByVal churchName As String) As String

    Dim tmpChurch As Church
    Dim churchCode As String
    
RETRY:
    For Each tmpChurch In churchList
        If tmpChurch.Name = churchName Then
            churchCode = tmpChurch.Id
            Exit For
        End If
    Next
    
    If churchCode = vbNullString And InStr(churchName, "_���") > 0 Then
        churchName = Replace(churchName, "_���", "")
        GoTo RETRY
    End If
    
    getChurchCodeinList = churchCode

End Function

Sub ConvertAttendanceDataToTableFormat(updateDate As Date)
    
    '--//1���� ��ȸ�ڵ� ���԰� �ߺ����Ÿ� �Ϸ��� ����Ʈ�� ���� �۾��� �����Ѵ�.
    '--//�ڡڡڷ� ǥ�õ� ����, Ư�� ��¥ ���� �ùٷ� ������ �� ���ν����� �����Ѵ�.
    
    Dim rngTarget As Range
    Dim intRotate As Integer
    Dim R As Long
    Dim intUpdateMonth As Integer
    Dim intUpdateYear As Integer
    Dim getDate As Date
    Dim i As Long
    
    '--//�⼮���� ������ ���� �ʿ� ������ ����
    Dim targetSheet As Worksheet
    Range("E:F").Copy
    Set targetSheet = Sheets.Add
    targetSheet.Paste
    targetSheet.Activate
    Range("A:B").RemoveDuplicates Columns:=Array(1, 2), Header:=xlYes
    
    '--//��������: �ڡڡ�
'    intUpdateMonth = 8 '��
'    intUpdateYear = 2023 '��
    intUpdateMonth = month(updateDate)
    intUpdateYear = year(updateDate)
    intRotate = 12 ''--//12�� �ݺ�
    R = Cells(Rows.Count, "A").End(xlUp).Row '������ �� ��ȣ
    Set rngTarget = Range(Cells(2, "A"), Cells(R, "C"))
    getDate = DateSerial(intUpdateYear, intUpdateMonth, 1)
    
    '--//����ι�
    '-->>�ʵ� �ؽ�Ʈ �Է�
    Range("C1") = "��¥"
    Range("D1") = "�⼮(��ü 1ȸ)"
    Range("E1") = "�⼮(��ü 4ȸ)"
    Range("F1") = "�⼮(�л��̻� 1ȸ)"
    Range("G1") = "�⼮(�л��̻� 4ȸ)"
    Range("H1") = "����(��ü)"
    Range("I1") = "����(�л��̻�)"
    Range("J1") = "ħ��(��ü)"
    Range("K1") = "����������"
    Range("L1") = "������ ��(�ӽ� ����)"
    Range("M1") = "������ ��(�ӽ� ����)"
    Range("C:C").NumberFormatLocal = "yyyy-mm-dd"
    
    '-->>������ ������ �������� ���� 12����ġ ����Ʈ ����
    Range(Cells(2, "C"), Cells(R, "C")) = getDate
    rngTarget.Copy
    For i = 1 To intRotate
        rngTarget.Offset((R - 1) * i).PasteSpecial Paste:=xlPasteValues
        rngTarget.Offset((R - 1) * i, 2).Resize(, 1) = WorksheetFunction.EDate(getDate, -1 * i)
        Application.Wait DateAdd("s", 1, Now)
    Next
    Range("C:C").Columns.AutoFit
    
    
    '-->>�⼮ ������ �ҷ����� ���� ���� / ������: �ð��� ���� �ɸ� �� ����.
    rngTarget.Offset(, 3).Resize((R - 1) * (intRotate + 1), 10).FormulaR1C1 = _
        "=SUMIFS(OFFSET(Sheet1!C1,,MATCH(TEXT(" & targetSheet.Name & "!RC3,""yyyy-mm""),Sheet1!R1,0)-1),Sheet1!C5," & targetSheet.Name & "!RC1,Sheet1!C14," & targetSheet.Name & "!R1C)"
    
    '--//������ ����
    ActiveSheet.UsedRange.Copy
    ActiveSheet.UsedRange.PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    ActiveSheet.Copy after:=Sheets(Sheets.Count)
    
    ActiveSheet.Name = "Backup"
    Range("B:B").Delete Shift:=xlLeft
    Range("A1").Select
    
    Call insertDataToDB_AttendanceData
    
    MsgBox "�۾��� �Ϸ��Ͽ����ϴ�."

End Sub

Sub insertDataToDB_AttendanceData()
    
    Dim tableNM As String, strSql As String
    Dim cntField As Integer, j As Long, k As Long
    
   
    '//DB�� �̵��� �ڷ� Values �迭�� ����
    Dim objAtten As Attendance
    Dim objAttenDao As New AttendanceDao
    Dim vAttenList As Object
    Set vAttenList = CreateObject("System.Collections.ArrayList")
    
    For j = 1 To Cells(Rows.Count, "A").End(xlUp).Row - 1
        Set objAtten = New Attendance
        objAtten.ParseFromRange Range("A1").Offset(j)
        vAttenList.Add objAtten
'        objAttenDao.Save objAtten
        k = k + 1
    Next j
    
    objAttenDao.SaveAll vAttenList
    
    MsgBox k & "���� ���ڵ尡 '" & TABLE_ATTENDANCE & "'�� �߰��Ǿ����ϴ�.", vbInformation, banner
    
End Sub

Sub insertDataToDB_ChurchList_from_ProjectTeam()
    
    Dim tableNM As String, strSql As String
    Dim cntField As Integer, j As Long, k As Long
   
    '//DB�� �̵��� �ڷ� Values �迭�� ����
    Dim objChurch As Church
    Dim objChurchDao As New ChurchDao
    Dim vChurchList As Object
    Set vChurchList = CreateObject("System.Collections.ArrayList")
    
    For j = 1 To Cells(Rows.Count, "A").End(xlUp).Row - 1
        Set objChurch = New Church
        objChurch.ParseFromRange Range("A1").Offset(j)
        vChurchList.Add objChurch
'        objAttenDao.Save objAtten
        k = k + 1
    Next j
    
    objChurchDao.SaveAll vChurchList
    
    MsgBox k & "���� ���ڵ尡 '" & TABLE_CHURCH & "'�� �߰��Ǿ����ϴ�.", vbInformation, banner
    
    '--//churchlist_custom ������Ʈ
    strSql = "CALL `Routine_make_churchlist_custom`();"
    connectTaskDB
    executeSQL "insertDataToDB_ChurchList_from_ProjectTeam", "db_churchlist_custom", strSql, , "Ŀ���� ��ȸ����Ʈ ����"
'    writeLog "cmdADD_Click", "temp_pstaff_by_time", strSQL, 0, Me.Name, "�ӽ���ȸ ���̺� ����", result.affectedCount
    disconnectALL
    
End Sub

Sub insertDataToDB_pastoralStaff()
    
    Dim shtNM As String, tableNM As String, dbNM As String, strSql As String
    Dim affectedCount As Long
    Dim Values() As String
    Dim cntField As Integer, cntRecord As Long, i As Long, j As Long, k As Integer
    
    '//Sheet��, Table��, db�� ����
    shtNM = "sheet1" '//�����ڡ�
    tableNM = "op_system.db_pastoralstaff" '//�����ڡ�
    dbNM = "op_system" '//�����ڡ�
    
    '//DB����
    Call connectTaskDB
    
    '//Table �ʱ�ȭ
'    strSQL = "TRUNCATE TABLE " & tableNM
'    affectedCount = executeSQL("insertDataToDB_Everything", tableNM, strSQL, "������ �ʱ�ȭ")
    
    '/�迭 ũ�� ����
    Sheets(shtNM).Activate
    cntField = Range("A1").CurrentRegion.Columns.Count
    cntRecord = Range("A1").CurrentRegion.Rows.Count - 1
    ReDim Values(cntRecord - 1, cntField - 1)
   
    '//DB�� �̵��� �ڷ� Values �迭�� ����
    For i = 0 To cntField - 1
        For j = 0 To cntRecord - 1
            Values(j, i) = Range("A2").Offset(j, i)
        Next j
    Next i
    
    '//���ڵ庰�� Insert(�����ڡ�)
    '  - �ʵ� �� ��ŭ �߰��ؼ� ����
    '  - ������ NULL���� �ְ��� �� ��� "''"�� "NULL"�� ����
    '  - �⺻���� �������� "Default" & "," & _
    '  - TimeStamp�� ��������: "CURRENT_TIMESTAMP()" & ");"
    For i = 0 To cntRecord - 1
        strSql = "INSERT INTO " & tableNM & " VALUES(" & _
                       convertStrToSQL(Values(i, 0)) & "," & convertStrToSQL(Values(i, 1)) & "," & convertStrToSQL(Values(i, 2)) & "," & _
                       IIf(Values(i, 3) = "", "Null", convertStrToSQL(Values(i, 3))) & "," & convertStrToSQL(Values(i, 4)) & "," & convertStrToSQL(Values(i, 5)) & "," & _
                       convertStrToSQL(Values(i, 6)) & "," & convertStrToSQL(Values(i, 7)) & "," & convertStrToSQL(Values(i, 8)) & "," & _
                       IIf(Values(i, 9) = "", "Null", convertStrToSQL(Values(i, 9))) & "," & convertStrToSQL(Values(i, 10)) & "," & convertStrToSQL(Values(i, 11)) & "," & _
                       convertStrToSQL(Values(i, 12)) & "," & convertStrToSQL(Values(i, 13)) & "," & IIf(Values(i, 14) = "", "Null", convertStrToSQL(Values(i, 14))) & "," & _
                       convertStrToSQL(Values(i, 15)) & "," & convertStrToSQL(Values(i, 16)) & "," & convertStrToSQL(Values(i, 17)) & "," & _
                       convertStrToSQL(Values(i, 18)) & "," & IIf(Values(i, 19) = "", "Null", convertStrToSQL(Values(i, 19))) & "," & convertStrToSQL(Values(i, 20)) & "," & _
                       convertStrToSQL(Values(i, 21)) & "," & convertStrToSQL(Values(i, 22)) & "," & convertStrToSQL(Values(i, 23)) & "," & _
                       convertStrToSQL(Values(i, 24)) & "," & convertStrToSQL(Values(i, 25)) & "," & IIf(Values(i, 26) = "", "Null", convertStrToSQL(Values(i, 26))) & "," & _
                       IIf(Values(i, 27) = "", "Null", convertStrToSQL(Values(i, 27))) & "," & IIf(Values(i, 28) = "", "Null", convertStrToSQL(Values(i, 28))) & "," & IIf(Values(i, 29) = "", "Null", convertStrToSQL(Values(i, 29))) & "," & _
                       convertStrToSQL(Values(i, 30)) & "," & IIf(Values(i, 31) = "", "Null", convertStrToSQL(Values(i, 31))) & "," & convertStrToSQL(Values(i, 32)) & "," & _
                       convertStrToSQL(Values(i, 33)) & ");"
                       
'        strSQL = "UPDATE " & tableNM & " a SET " & _
'                       "a.name_ko = " & convertStrToSQL(Values(i, 1)) & "," & "a.name_en = " & convertStrToSQL(Values(i, 2)) & "," & "a.nationality = " & convertStrToSQL(Values(i, 3)) & "," & _
'                       "a.birthday = " & IIf(Values(i, 4) = "", "Null", convertStrToSQL(Values(i, 4))) & "," & "a.phone = " & convertStrToSQL(Values(i, 5)) & "," & "a.lifeno_child1 = " & convertStrToSQL(Values(i, 6)) & "," & _
'                       "a.name_ko_child1 = " & convertStrToSQL(Values(i, 7)) & "," & "a.name_en_child1 = " & convertStrToSQL(Values(i, 8)) & "," & _
'                       "a.birthday_child1 = " & IIf(Values(i, 9) = "", "Null", convertStrToSQL(Values(i, 9))) & "," & "a.phone_child1 = " & convertStrToSQL(Values(i, 10)) & "," & _
'                       "a.lifeno_child2 = " & convertStrToSQL(Values(i, 11)) & "," & "a.name_ko_child2 = " & convertStrToSQL(Values(i, 12)) & "," & "a.name_en_child2 = " & convertStrToSQL(Values(i, 13)) & "," & _
'                       "a.birthday_child2 = " & IIf(Values(i, 14) = "", "Null", convertStrToSQL(Values(i, 14))) & "," & "a.phone_child2 = " & convertStrToSQL(Values(i, 15)) & "," & _
'                       "a.lifeno_child3 = " & convertStrToSQL(Values(i, 16)) & "," & "a.name_ko_child3 = " & convertStrToSQL(Values(i, 17)) & "," & "a.name_en_child3 = " & convertStrToSQL(Values(i, 18)) & "," & _
'                       "a.birthday_child3 = " & IIf(Values(i, 19) = "", "Null", convertStrToSQL(Values(i, 19))) & "," & "a.phone_child3 = " & convertStrToSQL(Values(i, 20)) & "," & _
'                       "a.home = " & convertStrToSQL(Values(i, 21)) & "," & "a.family = " & convertStrToSQL(Values(i, 22)) & "," & "a.health = " & convertStrToSQL(Values(i, 23)) & "," & _
'                       "a.other = " & convertStrToSQL(Values(i, 24)) & "," & "a.baptism = " & convertStrToSQL(Values(i, 25)) & "," & _
'                       "a.ordination_prayer = " & IIf(Values(i, 26) = "", "Null", convertStrToSQL(Values(i, 26))) & "," & _
'                       "a.appo_ovs = " & IIf(Values(i, 27) = "", "Null", convertStrToSQL(Values(i, 27))) & "," & _
'                       "a.wedding_dt = " & IIf(Values(i, 28) = "", "Null", convertStrToSQL(Values(i, 28))) & "," & _
'                       "a.theological_order = " & IIf(Values(i, 29) = "", "Null", convertStrToSQL(Values(i, 29))) & "," & _
'                       "a.education = " & convertStrToSQL(Values(i, 30)) & "," & _
'                       "a.salary = " & IIf(Values(i, 31) = "", "Null", convertStrToSQL(Values(i, 31))) & "," & _
'                       "a.suspend = " & convertStrToSQL(Values(i, 32)) & "," & "a.ovs_dept = " & convertStrToSQL(Values(i, 33)) & _
'                       " WHERE a.lifeno = " & convertStrToSQL(Values(i, 0)) & ";"
        affectedCount = executeSQL("insertDataToDB_ChurchList_Admin", tableNM, strSql, , "���� ������ �о�ֱ�")
        k = k + affectedCount
    Next i
    MsgBox k & "���� ���ڵ尡 '" & tableNM & "'�� �߰��Ǿ����ϴ�.", vbInformation, banner
    
    '//���� ����
    disconnectDB
End Sub

Sub insertDataToDB_pastoralWife()
    
    Dim shtNM As String, tableNM As String, dbNM As String, strSql As String
    Dim affectedCount As Long
    Dim Values() As String
    Dim cntField As Integer, cntRecord As Long, i As Long, j As Long, k As Integer
    
    '//Sheet��, Table��, db�� ����
    shtNM = "sheet1" '//�����ڡ�
    tableNM = "op_system.db_pastoralwife" '//�����ڡ�
    dbNM = "op_system" '//�����ڡ�
    
    '//DB����
    Call connectTaskDB
    
    '//Table �ʱ�ȭ
'    strSQL = "TRUNCATE TABLE " & tableNM
'    affectedCount = executeSQL("insertDataToDB_Everything", tableNM, strSQL, "������ �ʱ�ȭ")
    
    '/�迭 ũ�� ����
    Sheets(shtNM).Activate
    cntField = Range("A1").CurrentRegion.Columns.Count
    cntRecord = Range("A1").CurrentRegion.Rows.Count - 1
    ReDim Values(cntRecord - 1, cntField - 1)
   
    '//DB�� �̵��� �ڷ� Values �迭�� ����
    For i = 0 To cntField - 1
        For j = 0 To cntRecord - 1
            Values(j, i) = Range("A2").Offset(j, i)
        Next j
    Next i
    
    '//���ڵ庰�� Insert(�����ڡ�)
    '  - �ʵ� �� ��ŭ �߰��ؼ� ����
    '  - ������ NULL���� �ְ��� �� ��� "''"�� "NULL"�� ����
    '  - �⺻���� �������� "Default" & "," & _
    '  - TimeStamp�� ��������: "CURRENT_TIMESTAMP()" & ");"
    For i = 0 To cntRecord - 1
        strSql = "INSERT INTO " & tableNM & " VALUES(" & _
                       convertStrToSQL(Values(i, 0)) & "," & convertStrToSQL(Values(i, 1)) & "," & convertStrToSQL(Values(i, 2)) & "," & _
                       convertStrToSQL(Values(i, 3)) & "," & convertStrToSQL(Values(i, 4)) & "," & _
                       IIf(Values(i, 5) = "", "Null", convertStrToSQL(Values(i, 5))) & "," & IIf(Values(i, 6) = "", "Null", convertStrToSQL(Values(i, 6))) & "," & "Null" & "," & _
                       IIf(Values(i, 8) = "", "Null", convertStrToSQL(Values(i, 8))) & "," & IIf(Values(i, 9) = "", "Null", convertStrToSQL(Values(i, 9))) & "," & IIf(Values(i, 10) = "", "Null", convertStrToSQL(Values(i, 10))) & "," & _
                       IIf(Values(i, 11) = "", "Null", convertStrToSQL(Values(i, 11))) & "," & convertStrToSQL(Values(i, 12)) & "," & convertStrToSQL(Values(i, 13)) & ");"
                       
'        strSQL = "UPDATE " & tableNM & " a SET " & _
'                       "a.name_ko = " & convertStrToSQL(Values(i, 1)) & "," & "a.name_en = " & convertStrToSQL(Values(i, 2)) & "," & "a.nationality = " & convertStrToSQL(Values(i, 3)) & "," & _
'                       "a.birthday = " & IIf(Values(i, 4) = "", "Null", convertStrToSQL(Values(i, 4))) & "," & "a.phone = " & convertStrToSQL(Values(i, 5)) & "," & "a.lifeno_child1 = " & convertStrToSQL(Values(i, 6)) & "," & _
'                       "a.name_ko_child1 = " & convertStrToSQL(Values(i, 7)) & "," & "a.name_en_child1 = " & convertStrToSQL(Values(i, 8)) & "," & _
'                       "a.birthday_child1 = " & IIf(Values(i, 9) = "", "Null", convertStrToSQL(Values(i, 9))) & "," & "a.phone_child1 = " & convertStrToSQL(Values(i, 10)) & "," & _
'                       "a.lifeno_child2 = " & convertStrToSQL(Values(i, 11)) & "," & "a.name_ko_child2 = " & convertStrToSQL(Values(i, 12)) & "," & "a.name_en_child2 = " & convertStrToSQL(Values(i, 13)) & "," & _
'                       "a.birthday_child2 = " & IIf(Values(i, 14) = "", "Null", convertStrToSQL(Values(i, 14))) & "," & "a.phone_child2 = " & convertStrToSQL(Values(i, 15)) & "," & _
'                       "a.lifeno_child3 = " & convertStrToSQL(Values(i, 16)) & "," & "a.name_ko_child3 = " & convertStrToSQL(Values(i, 17)) & "," & "a.name_en_child3 = " & convertStrToSQL(Values(i, 18)) & "," & _
'                       "a.birthday_child3 = " & IIf(Values(i, 19) = "", "Null", convertStrToSQL(Values(i, 19))) & "," & "a.phone_child3 = " & convertStrToSQL(Values(i, 20)) & "," & _
'                       "a.home = " & convertStrToSQL(Values(i, 21)) & "," & "a.family = " & convertStrToSQL(Values(i, 22)) & "," & "a.health = " & convertStrToSQL(Values(i, 23)) & "," & _
'                       "a.other = " & convertStrToSQL(Values(i, 24)) & "," & "a.baptism = " & convertStrToSQL(Values(i, 25)) & "," & _
'                       "a.ordination_prayer = " & IIf(Values(i, 26) = "", "Null", convertStrToSQL(Values(i, 26))) & "," & _
'                       "a.appo_ovs = " & IIf(Values(i, 27) = "", "Null", convertStrToSQL(Values(i, 27))) & "," & _
'                       "a.wedding_dt = " & IIf(Values(i, 28) = "", "Null", convertStrToSQL(Values(i, 28))) & "," & _
'                       "a.theological_order = " & IIf(Values(i, 29) = "", "Null", convertStrToSQL(Values(i, 29))) & "," & _
'                       "a.education = " & convertStrToSQL(Values(i, 30)) & "," & _
'                       "a.salary = " & IIf(Values(i, 31) = "", "Null", convertStrToSQL(Values(i, 31))) & "," & _
'                       "a.suspend = " & convertStrToSQL(Values(i, 32)) & "," & "a.ovs_dept = " & convertStrToSQL(Values(i, 33)) & _
'                       " WHERE a.lifeno = " & convertStrToSQL(Values(i, 0)) & ";"
        affectedCount = executeSQL("insertDataToDB_ChurchList_Admin", tableNM, strSql, , "���� ������ �о�ֱ�")
        k = k + affectedCount
    Next i
    MsgBox k & "���� ���ڵ尡 '" & tableNM & "'�� �߰��Ǿ����ϴ�.", vbInformation, banner
    
    '//���� ����
    disconnectDB
End Sub
Sub insertDataToDB_sermon()
    
    Dim shtNM As String, tableNM As String, dbNM As String, strSql As String
    Dim affectedCount As Long
    Dim Values() As String
    Dim cntField As Integer, cntRecord As Long, i As Long, j As Long, k As Integer
    
    '//Sheet��, Table��, db�� ����
    shtNM = "insert" '//�����ڡ�
    tableNM = "op_system.db_sermon" '//�����ڡ�
    dbNM = "op_system" '//�����ڡ�
    
    '//DB����
    Call connectTaskDB
    
    '//Table �ʱ�ȭ
'    strSQL = "TRUNCATE TABLE " & tableNM
'    affectedCount = executeSQL("insertDataToDB_Everything", tableNM, strSQL, "������ �ʱ�ȭ")
    
    '/�迭 ũ�� ����
    Sheets(shtNM).Activate
    cntField = Range("A1").CurrentRegion.Columns.Count
    cntRecord = Range("A1").CurrentRegion.Rows.Count - 1
    ReDim Values(cntRecord - 1, cntField - 1)
   
    '//DB�� �̵��� �ڷ� Values �迭�� ����
    For i = 0 To cntField - 1
        For j = 0 To cntRecord - 1
            Values(j, i) = Range("A2").Offset(j, i)
        Next j
    Next i
    
    '//���ڵ庰�� Insert(�����ڡ�)
    '  - �ʵ� �� ��ŭ �߰��ؼ� ����
    '  - ������ NULL���� �ְ��� �� ��� "''"�� "NULL"�� ����
    '  - �⺻���� �������� "Default" & "," & _
    '  - TimeStamp�� ��������: "CURRENT_TIMESTAMP()" & ");"
    For i = 0 To cntRecord - 1
        strSql = "INSERT INTO " & tableNM & " VALUES(" & _
                       convertStrToSQL(Values(i, 0)) & "," & _
                       convertStrToSQL(Values(i, 1)) & "," & _
                       convertStrToSQL(Values(i, 2)) & ");"
'        strSQL = "UPDATE " & tableNM & " a SET " & _
'                       "a.score_avg = " & convertStrToSQL(Values(i, 1)) & "," & _
'                       "a.subject_count = " & convertStrToSQL(Values(i, 2)) & _
'                        " WHERE a.lifeno = " & convertStrToSQL(Values(i, 0)) & ";"
        affectedCount = executeSQL("insertDataToDB_ChurchList_Admin", tableNM, strSql, , "���� ������ �о�ֱ�")
        k = k + affectedCount
    Next i
    MsgBox k & "���� ���ڵ尡 '" & tableNM & "'�� �߰��Ǿ����ϴ�.", vbInformation, banner
    
    '//���� ����
    disconnectDB
End Sub

Sub insertDataToDB_GeoData()
    
    Dim shtNM As String, tableNM As String, dbNM As String, strSql As String
    Dim affectedCount As Long
    Dim Values() As String
    Dim cntField As Integer, cntRecord As Long, i As Long, j As Long, k As Long
    
    '//Sheet��, Table��, db�� ����
    shtNM = "geo_data" '//�����ڡ�
    tableNM = "op_system.db_geodata" '//�����ڡ�
    dbNM = "op_system" '//�����ڡ�
    
    '//DB����
    Call connectTaskDB
    
    '//Table �ʱ�ȭ
    strSql = "TRUNCATE TABLE " & tableNM
    affectedCount = executeSQL("insertDataToDB_GeoData", tableNM, strSql, "������ �ʱ�ȭ")
    
    '/�迭 ũ�� ����
    Sheets(shtNM).Activate
    cntField = Range("A1").CurrentRegion.Columns.Count
    cntRecord = Range("A1").CurrentRegion.Rows.Count - 1
    ReDim Values(cntRecord - 1, cntField - 1)
   
    '//DB�� �̵��� �ڷ� Values �迭�� ����
    For i = 0 To cntField - 1
        For j = 0 To cntRecord - 1
            Values(j, i) = Range("A2").Offset(j, i)
        Next j
    Next i
    
    '//���ڵ庰�� Insert(�����ڡ�)
    '  - �ʵ� �� ��ŭ �߰��ؼ� ����
    '  - ������ NULL���� �ְ��� �� ��� "''"�� "NULL"�� ����
    '  - �⺻���� �������� "Default" & "," & _
    '  - TimeStamp�� ��������: "CURRENT_TIMESTAMP()" & ");"
    For i = 0 To cntRecord - 1
        strSql = "INSERT INTO " & tableNM & " VALUES(" & _
                       convertStrToSQL(Values(i, 0)) & "," & convertStrToSQL(Values(i, 1)) & "," & _
                       convertStrToSQL(Values(i, 2)) & "," & IIf(Values(i, 3) = "", "Null", convertStrToSQL(Values(i, 3))) & "," & _
                       IIf(Values(i, 4) = "", "Null", convertStrToSQL(Values(i, 4))) & "," & IIf(Values(i, 5) = "", "Null", convertStrToSQL(Values(i, 5))) & "," & _
                       IIf(Values(i, 6) = "", "Null", convertStrToSQL(Values(i, 6))) & "," & IIf(Values(i, 7) = "", "Null", convertStrToSQL(Values(i, 7))) & "," & _
                       IIf(Values(i, 8) = "", "Null", convertStrToSQL(Values(i, 8))) & "," & IIf(Values(i, 9) = "", "Null", convertStrToSQL(Values(i, 9))) & "," & _
                       IIf(Values(i, 10) = "", "Null", convertStrToSQL(Values(i, 10))) & "," & IIf(Values(i, 11) = "", "Null", convertStrToSQL(Values(i, 11))) & "," & _
                       IIf(Values(i, 12) = "", "Null", convertStrToSQL(Values(i, 12))) & "," & IIf(Values(i, 13) = "", "Null", convertStrToSQL(Values(i, 13))) & "," & _
                       IIf(Values(i, 14) = "", "Null", convertStrToSQL(Values(i, 14))) & "," & IIf(Values(i, 15) = "", "Null", convertStrToSQL(Values(i, 15))) & "," & _
                       IIf(Values(i, 16) = "", "Null", convertStrToSQL(Values(i, 16))) & "," & IIf(Values(i, 17) = "", "Null", convertStrToSQL(Values(i, 17))) & "," & _
                       IIf(Values(i, 18) = "", "Null", convertStrToSQL(Values(i, 18))) & "," & IIf(Values(i, 19) = "", "Null", convertStrToSQL(Values(i, 19))) & "," & _
                       IIf(Values(i, 20) = "", "Null", convertStrToSQL(Values(i, 20))) & "," & IIf(Values(i, 21) = "", "Null", convertStrToSQL(Values(i, 21))) & "," & _
                       IIf(Values(i, 22) = "", "Null", convertStrToSQL(Values(i, 22))) & "," & IIf(Values(i, 23) = "", "Null", convertStrToSQL(Values(i, 23))) & "," & _
                       IIf(Values(i, 24) = "", "Null", convertStrToSQL(Values(i, 24))) & "," & IIf(Values(i, 25) = "", "Null", convertStrToSQL(Values(i, 25))) & ");"
                       
        affectedCount = executeSQL("insertDataToDB_ChurchList_Admin", tableNM, strSql, , "���� ������ �о�ֱ�")
        k = k + affectedCount
    Next i
    MsgBox k & "���� ���ڵ尡 '" & tableNM & "'�� �߰��Ǿ����ϴ�.", vbInformation, banner
    
    '//���� ����
    disconnectDB
End Sub
