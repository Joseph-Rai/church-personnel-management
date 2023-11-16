Attribute VB_Name = "sb_EXcelToDB"
Option Explicit

Sub insertDataToDB_Everything()
    
    Dim shtNM As String, tableNM As String, dbNM As String, strSql As String
    Dim affectedCount As Long
    Dim Values() As String
    Dim cntField As Integer, cntRecord As Long, i As Long, j As Long, k As Long
    
    '//Sheet명, Table명, db명 설정
    shtNM = "sheet1" '//수정★★
    tableNM = "op_system.db_division" '//수정★★
    dbNM = "op_system" '//수정★★
    
    '//DB연결
    Call connectTaskDB
    
    '//Table 초기화
'    strSQL = "TRUNCATE TABLE " & tableNM
'    affectedCount = executeSQL("insertDataToDB_Everything", tableNM, strSQL, "데이터 초기화")
    
    '/배열 크기 지정
    Sheets(shtNM).Activate
    cntField = Range("A1").CurrentRegion.Columns.Count
    cntRecord = Range("A1").CurrentRegion.Rows.Count - 1
    ReDim Values(cntRecord - 1, cntField - 1)
   
    '//DB로 이동할 자료 Values 배열에 저장
    For i = 0 To cntField - 1
        For j = 0 To cntRecord - 1
            Values(j, i) = Range("A2").Offset(j, i)
        Next j
    Next i
    
    '//레코드별로 Insert(수정★★)
    '  - 필드 수 만큼 추가해서 진행
    '  - 공백대신 NULL값을 넣고자 할 경우 "''"을 "NULL"로 수정
    '  - 기본값을 넣으려면 "Default" & "," & _
    '  - TimeStamp를 넣으려면: "CURRENT_TIMESTAMP()" & ");"
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
        affectedCount = executeSQL("insertDataToDB_ChurchDivision", tableNM, strSql, , "엑셀 데이터 밀어넣기")
        k = k + affectedCount
    Next i
    MsgBox k & "개의 레코드가 '" & tableNM & "'에 추가되었습니다.", vbInformation, banner
    
    '//연결 끊기
    disconnectDB
End Sub
'---------------------------------------------------
'  엑셀 자료를 DB에 Insert
'    - Table 초기화
'    - 엑셀 레코드 하나씩 Insert
'----------------------------------------------------
Sub insertDataToDB_BranchList_Admin()
    
    Dim shtNM As String, tableNM As String, dbNM As String, strSql As String
    Dim affectedCount As Long
    Dim Values() As String
    Dim cntField As Integer, cntRecord As Long, i As Long, j As Long, k As Integer
    
    '//Sheet명, Table명, db명 설정
    shtNM = "Sheet1" '//수정★★
    tableNM = "op_system.a_branch_admin" '//수정★★
    dbNM = "op_system" '//수정★★
    
    '//DB연결
    Call connectTaskDB
    
    '//Table 초기화
    strSql = "TRUNCATE TABLE " & tableNM
    affectedCount = executeSQL("insertDataToDB_BranchList_Admin", tableNM, strSql, "데이터 초기화")
    
    '/배열 크기 지정
    Sheets(shtNM).Activate
    cntField = Range("A1").CurrentRegion.Columns.Count
    cntRecord = Range("A1").CurrentRegion.Rows.Count - 1
    ReDim Values(cntRecord - 1, cntField - 1)
   
    '//DB로 이동할 자료 Values 배열에 저장
    For i = 0 To cntField - 1
        For j = 0 To cntRecord - 1
            Values(j, i) = Range("A2").Offset(j, i)
        Next j
    Next i
    
    '//레코드별로 Insert(수정★★)
    '  - 필드 수 만큼 추가해서 진행
    '  - 공백대신 NULL값을 넣고자 할 경우 "''"을 "NULL"로 수정
    '  - 기본값을 넣으려면 "Default" & "," & _
    '  - TimeStamp를 넣으려면: "CURRENT_TIMESTAMP()" & ");"
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
        affectedCount = executeSQL("insertDataToDB_BranchList_Admin", tableNM, strSql, , "엑셀 데이터 밀어넣기")
        k = k + affectedCount
    Next i
    MsgBox k & "개의 레코드가 '" & tableNM & "'에 추가되었습니다.", vbInformation, banner
    
    '//연결 끊기
    disconnectDB
End Sub
Sub insertDataToDB_ChurchList_Admin()
    
    Dim shtNM As String, tableNM As String, dbNM As String, strSql As String
    Dim affectedCount As Long
    Dim Values() As String
    Dim cntField As Integer, cntRecord As Long, i As Long, j As Long, k As Integer
    
    '//Sheet명, Table명, db명 설정
    shtNM = "Sheet1" '//수정★★
    tableNM = "op_system.a_churchlist_admin" '//수정★★
    dbNM = "op_system" '//수정★★
    
    '//DB연결
    Call connectTaskDB
    
    '//Table 초기화
    strSql = "TRUNCATE TABLE " & tableNM & ";"
'    affectedCount = executeSQL("insertDataToDB_ChurchList_Admin", tableNM, strSql, "데이터 초기화")
    
    '/배열 크기 지정
    Sheets(shtNM).Activate
    cntField = Range("A1").CurrentRegion.Columns.Count
    cntRecord = Range("A1").CurrentRegion.Rows.Count - 1
    ReDim Values(cntRecord - 1, cntField - 1)
   
    '//DB로 이동할 자료 Values 배열에 저장
    For i = 0 To cntField - 1
        For j = 0 To cntRecord - 1
            Values(j, i) = Range("A2").Offset(j, i)
        Next j
    Next i
    
    '//레코드별로 Insert(수정★★)
    '  - 필드 수 만큼 추가해서 진행
    '  - 공백대신 NULL값을 넣고자 할 경우 "''"을 "NULL"로 수정
    '  - 기본값을 넣으려면 "Default" & "," & _
    '  - TimeStamp를 넣으려면: "CURRENT_TIMESTAMP()" & ");"
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
            affectedCount = executeSQL("insertDataToDB_ChurchList_Admin", tableNM, strSql, , "엑셀 데이터 밀어넣기")
            strSql = vbNullString
        End If
    Next i
    
    If strSql <> vbNullString Then
        affectedCount = executeSQL("insertDataToDB_ChurchList_Admin", tableNM, strSql, , "엑셀 데이터 밀어넣기")
    End If
    
    MsgBox k & "개의 레코드가 '" & tableNM & "'에 추가되었습니다.", vbInformation, banner
    
    '//연결 끊기
    disconnectDB
End Sub

Sub UpdateAttendanceRecent12Month()

    Dim maxRowNum As Long
    maxRowNum = Cells(Rows.Count, "F").End(xlUp).Row
    
    '--//교회코드 찾아서 넣기위한 LIst
    Dim churchList As Object
    Dim vChurchDao As New ChurchDao
    Set churchList = vChurchDao.getAllChurchListReverseOrder
    Dim churchCode As String
    
    Dim index As Long
    
    Application.ScreenUpdating = False
    
    '--//교회명 정리
    For index = 2 To maxRowNum Step 16
        Do While Cells(index, "H") = "" And Cells(index, "F") <> ""
            Cells(index, "F").Resize(16).EntireRow.Delete Shift:=xlUp
            maxRowNum = maxRowNum - 16
            index = WorksheetFunction.Max(2, index - 16)
        Loop
        
        If Cells(index, "H") = "본교회" Or Cells(index, "H") = "총회관리지교회" Then
            If InStr(Cells(index, "F"), "본교회") <= 0 Then
                Cells(index, "F").Resize(16) = Cells(index, "F") & " 본교회"
            End If
        End If
        
        If InStr(Cells(index, "F"), "전체") > 0 Then
            Cells(index, "F").Resize(16) = Replace(Cells(index, "F"), " 전체", "")
        End If
        
        If Not (Left(Cells(index, "E"), 2) = "MC" Or Left(Cells(index, "E"), 2) = "BC") Then
            churchCode = getChurchCodeinList(churchList, Cells(index, "F"))
            Cells(index, "E").Resize(16) = churchCode
        End If
    Next
    
    '--//전체, 본교회 세트로 안되어 있는 것 정리
    Dim targetRange As Range
    index = 2
    Do While index <= maxRowNum
        If InStr(Cells(index, "F"), "본교회") > 0 Then
            If WorksheetFunction.CountIf(Range("F:F"), Replace(Cells(index, "F"), " 본교회", "")) <= 0 Then
                Set targetRange = Cells(index, "F").Resize(16).EntireRow
                targetRange.Copy
                targetRange.Insert Shift:=xlDown
                Cells(index, "F").Resize(16) = Replace(Cells(index, "F"), " 본교회", "")
                Cells(index, "H").Resize(16) = "관리교회"
                Cells(index, "E").Resize(16) = Replace(Cells(index, "E"), "MM", "MC")
                Application.CutCopyMode = False
                index = index + 16
                maxRowNum = maxRowNum + 16
            End If
        Else
            If Cells(index, "H") = "관리교회" Then
                If WorksheetFunction.CountIf(Range("F:F"), Cells(index, "F") & " 본교회") <= 0 Then
                    Set targetRange = Cells(index, "F").Resize(16).EntireRow
                    targetRange.Copy
                    targetRange.Insert Shift:=xlDown
                    Cells(index, "F").Resize(16).Offset(16) = Cells(index, "F") & " 본교회"
                    Cells(index, "H").Resize(16).Offset(16) = "본교회"
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
    
    ' 현재 열려있는 파일의 저장 여부 확인
    If wb.Path = "" Then
        ' 저장되지 않은 경우, 저장 다이얼로그 띄우기
        With Application.FileDialog(msoFileDialogSaveAs)
            .title = "파일 저장"
            .Show
            If .SelectedItems.Count > 0 Then
                filePath = .SelectedItems(1)
                wb.SaveAs filePath
            End If
        End With
    Else
        ' 이미 저장된 경우, 현재 경로에 저장하기
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
    
    If churchCode = vbNullString And InStr(churchName, "_폐쇄") > 0 Then
        churchName = Replace(churchName, "_폐쇄", "")
        GoTo RETRY
    End If
    
    getChurchCodeinList = churchCode

End Function

Sub ConvertAttendanceDataToTableFormat(updateDate As Date)
    
    '--//1차로 교회코드 삽입과 중복제거를 완료한 리스트를 토대로 작업을 진행한다.
    '--//★★★로 표시된 변수, 특히 날짜 등을 올바로 수정한 후 프로시저를 실행한다.
    
    Dim rngTarget As Range
    Dim intRotate As Integer
    Dim R As Long
    Dim intUpdateMonth As Integer
    Dim intUpdateYear As Integer
    Dim getDate As Date
    Dim i As Long
    
    '--//출석정보 정리를 위한 필요 데이터 복사
    Dim targetSheet As Worksheet
    Range("E:F").Copy
    Set targetSheet = Sheets.Add
    targetSheet.Paste
    targetSheet.Activate
    Range("A:B").RemoveDuplicates Columns:=Array(1, 2), Header:=xlYes
    
    '--//변수설정: ★★★
'    intUpdateMonth = 8 '월
'    intUpdateYear = 2023 '년
    intUpdateMonth = month(updateDate)
    intUpdateYear = year(updateDate)
    intRotate = 12 ''--//12번 반복
    R = Cells(Rows.Count, "A").End(xlUp).Row '마지막 행 번호
    Set rngTarget = Range(Cells(2, "A"), Cells(R, "C"))
    getDate = DateSerial(intUpdateYear, intUpdateMonth, 1)
    
    '--//실행부문
    '-->>필드 텍스트 입력
    Range("C1") = "날짜"
    Range("D1") = "출석(전체 1회)"
    Range("E1") = "출석(전체 4회)"
    Range("F1") = "출석(학생이상 1회)"
    Range("G1") = "출석(학생이상 4회)"
    Range("H1") = "반차(전체)"
    Range("I1") = "반차(학생이상)"
    Range("J1") = "침례(전체)"
    Range("K1") = "고정전도인"
    Range("L1") = "지역장 수(임시 포함)"
    Range("M1") = "구역장 수(임시 포함)"
    Range("C:C").NumberFormatLocal = "yyyy-mm-dd"
    
    '-->>설정된 연월을 기준으로 이전 12개월치 리스트 제작
    Range(Cells(2, "C"), Cells(R, "C")) = getDate
    rngTarget.Copy
    For i = 1 To intRotate
        rngTarget.Offset((R - 1) * i).PasteSpecial Paste:=xlPasteValues
        rngTarget.Offset((R - 1) * i, 2).Resize(, 1) = WorksheetFunction.EDate(getDate, -1 * i)
        Application.Wait DateAdd("s", 1, Now)
    Next
    Range("C:C").Columns.AutoFit
    
    
    '-->>출석 데이터 불러오는 수식 삽입 / ★주의: 시간이 오래 걸릴 수 있음.
    rngTarget.Offset(, 3).Resize((R - 1) * (intRotate + 1), 10).FormulaR1C1 = _
        "=SUMIFS(OFFSET(Sheet1!C1,,MATCH(TEXT(" & targetSheet.Name & "!RC3,""yyyy-mm""),Sheet1!R1,0)-1),Sheet1!C5," & targetSheet.Name & "!RC1,Sheet1!C14," & targetSheet.Name & "!R1C)"
    
    '--//값으로 고정
    ActiveSheet.UsedRange.Copy
    ActiveSheet.UsedRange.PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    ActiveSheet.Copy after:=Sheets(Sheets.Count)
    
    ActiveSheet.Name = "Backup"
    Range("B:B").Delete Shift:=xlLeft
    Range("A1").Select
    
    Call insertDataToDB_AttendanceData
    
    MsgBox "작업을 완료하였습니다."

End Sub

Sub insertDataToDB_AttendanceData()
    
    Dim tableNM As String, strSql As String
    Dim cntField As Integer, j As Long, k As Long
    
   
    '//DB로 이동할 자료 Values 배열에 저장
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
    
    MsgBox k & "개의 레코드가 '" & TABLE_ATTENDANCE & "'에 추가되었습니다.", vbInformation, banner
    
End Sub

Sub insertDataToDB_ChurchList_from_ProjectTeam()
    
    Dim tableNM As String, strSql As String
    Dim cntField As Integer, j As Long, k As Long
   
    '//DB로 이동할 자료 Values 배열에 저장
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
    
    MsgBox k & "개의 레코드가 '" & TABLE_CHURCH & "'에 추가되었습니다.", vbInformation, banner
    
    '--//churchlist_custom 업데이트
    strSql = "CALL `Routine_make_churchlist_custom`();"
    connectTaskDB
    executeSQL "insertDataToDB_ChurchList_from_ProjectTeam", "db_churchlist_custom", strSql, , "커스텀 교회리스트 생성"
'    writeLog "cmdADD_Click", "temp_pstaff_by_time", strSQL, 0, Me.Name, "임시조회 테이블 생성", result.affectedCount
    disconnectALL
    
End Sub

Sub insertDataToDB_pastoralStaff()
    
    Dim shtNM As String, tableNM As String, dbNM As String, strSql As String
    Dim affectedCount As Long
    Dim Values() As String
    Dim cntField As Integer, cntRecord As Long, i As Long, j As Long, k As Integer
    
    '//Sheet명, Table명, db명 설정
    shtNM = "sheet1" '//수정★★
    tableNM = "op_system.db_pastoralstaff" '//수정★★
    dbNM = "op_system" '//수정★★
    
    '//DB연결
    Call connectTaskDB
    
    '//Table 초기화
'    strSQL = "TRUNCATE TABLE " & tableNM
'    affectedCount = executeSQL("insertDataToDB_Everything", tableNM, strSQL, "데이터 초기화")
    
    '/배열 크기 지정
    Sheets(shtNM).Activate
    cntField = Range("A1").CurrentRegion.Columns.Count
    cntRecord = Range("A1").CurrentRegion.Rows.Count - 1
    ReDim Values(cntRecord - 1, cntField - 1)
   
    '//DB로 이동할 자료 Values 배열에 저장
    For i = 0 To cntField - 1
        For j = 0 To cntRecord - 1
            Values(j, i) = Range("A2").Offset(j, i)
        Next j
    Next i
    
    '//레코드별로 Insert(수정★★)
    '  - 필드 수 만큼 추가해서 진행
    '  - 공백대신 NULL값을 넣고자 할 경우 "''"을 "NULL"로 수정
    '  - 기본값을 넣으려면 "Default" & "," & _
    '  - TimeStamp를 넣으려면: "CURRENT_TIMESTAMP()" & ");"
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
        affectedCount = executeSQL("insertDataToDB_ChurchList_Admin", tableNM, strSql, , "엑셀 데이터 밀어넣기")
        k = k + affectedCount
    Next i
    MsgBox k & "개의 레코드가 '" & tableNM & "'에 추가되었습니다.", vbInformation, banner
    
    '//연결 끊기
    disconnectDB
End Sub

Sub insertDataToDB_pastoralWife()
    
    Dim shtNM As String, tableNM As String, dbNM As String, strSql As String
    Dim affectedCount As Long
    Dim Values() As String
    Dim cntField As Integer, cntRecord As Long, i As Long, j As Long, k As Integer
    
    '//Sheet명, Table명, db명 설정
    shtNM = "sheet1" '//수정★★
    tableNM = "op_system.db_pastoralwife" '//수정★★
    dbNM = "op_system" '//수정★★
    
    '//DB연결
    Call connectTaskDB
    
    '//Table 초기화
'    strSQL = "TRUNCATE TABLE " & tableNM
'    affectedCount = executeSQL("insertDataToDB_Everything", tableNM, strSQL, "데이터 초기화")
    
    '/배열 크기 지정
    Sheets(shtNM).Activate
    cntField = Range("A1").CurrentRegion.Columns.Count
    cntRecord = Range("A1").CurrentRegion.Rows.Count - 1
    ReDim Values(cntRecord - 1, cntField - 1)
   
    '//DB로 이동할 자료 Values 배열에 저장
    For i = 0 To cntField - 1
        For j = 0 To cntRecord - 1
            Values(j, i) = Range("A2").Offset(j, i)
        Next j
    Next i
    
    '//레코드별로 Insert(수정★★)
    '  - 필드 수 만큼 추가해서 진행
    '  - 공백대신 NULL값을 넣고자 할 경우 "''"을 "NULL"로 수정
    '  - 기본값을 넣으려면 "Default" & "," & _
    '  - TimeStamp를 넣으려면: "CURRENT_TIMESTAMP()" & ");"
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
        affectedCount = executeSQL("insertDataToDB_ChurchList_Admin", tableNM, strSql, , "엑셀 데이터 밀어넣기")
        k = k + affectedCount
    Next i
    MsgBox k & "개의 레코드가 '" & tableNM & "'에 추가되었습니다.", vbInformation, banner
    
    '//연결 끊기
    disconnectDB
End Sub
Sub insertDataToDB_sermon()
    
    Dim shtNM As String, tableNM As String, dbNM As String, strSql As String
    Dim affectedCount As Long
    Dim Values() As String
    Dim cntField As Integer, cntRecord As Long, i As Long, j As Long, k As Integer
    
    '//Sheet명, Table명, db명 설정
    shtNM = "insert" '//수정★★
    tableNM = "op_system.db_sermon" '//수정★★
    dbNM = "op_system" '//수정★★
    
    '//DB연결
    Call connectTaskDB
    
    '//Table 초기화
'    strSQL = "TRUNCATE TABLE " & tableNM
'    affectedCount = executeSQL("insertDataToDB_Everything", tableNM, strSQL, "데이터 초기화")
    
    '/배열 크기 지정
    Sheets(shtNM).Activate
    cntField = Range("A1").CurrentRegion.Columns.Count
    cntRecord = Range("A1").CurrentRegion.Rows.Count - 1
    ReDim Values(cntRecord - 1, cntField - 1)
   
    '//DB로 이동할 자료 Values 배열에 저장
    For i = 0 To cntField - 1
        For j = 0 To cntRecord - 1
            Values(j, i) = Range("A2").Offset(j, i)
        Next j
    Next i
    
    '//레코드별로 Insert(수정★★)
    '  - 필드 수 만큼 추가해서 진행
    '  - 공백대신 NULL값을 넣고자 할 경우 "''"을 "NULL"로 수정
    '  - 기본값을 넣으려면 "Default" & "," & _
    '  - TimeStamp를 넣으려면: "CURRENT_TIMESTAMP()" & ");"
    For i = 0 To cntRecord - 1
        strSql = "INSERT INTO " & tableNM & " VALUES(" & _
                       convertStrToSQL(Values(i, 0)) & "," & _
                       convertStrToSQL(Values(i, 1)) & "," & _
                       convertStrToSQL(Values(i, 2)) & ");"
'        strSQL = "UPDATE " & tableNM & " a SET " & _
'                       "a.score_avg = " & convertStrToSQL(Values(i, 1)) & "," & _
'                       "a.subject_count = " & convertStrToSQL(Values(i, 2)) & _
'                        " WHERE a.lifeno = " & convertStrToSQL(Values(i, 0)) & ";"
        affectedCount = executeSQL("insertDataToDB_ChurchList_Admin", tableNM, strSql, , "엑셀 데이터 밀어넣기")
        k = k + affectedCount
    Next i
    MsgBox k & "개의 레코드가 '" & tableNM & "'에 추가되었습니다.", vbInformation, banner
    
    '//연결 끊기
    disconnectDB
End Sub

Sub insertDataToDB_GeoData()
    
    Dim shtNM As String, tableNM As String, dbNM As String, strSql As String
    Dim affectedCount As Long
    Dim Values() As String
    Dim cntField As Integer, cntRecord As Long, i As Long, j As Long, k As Long
    
    '//Sheet명, Table명, db명 설정
    shtNM = "geo_data" '//수정★★
    tableNM = "op_system.db_geodata" '//수정★★
    dbNM = "op_system" '//수정★★
    
    '//DB연결
    Call connectTaskDB
    
    '//Table 초기화
    strSql = "TRUNCATE TABLE " & tableNM
    affectedCount = executeSQL("insertDataToDB_GeoData", tableNM, strSql, "데이터 초기화")
    
    '/배열 크기 지정
    Sheets(shtNM).Activate
    cntField = Range("A1").CurrentRegion.Columns.Count
    cntRecord = Range("A1").CurrentRegion.Rows.Count - 1
    ReDim Values(cntRecord - 1, cntField - 1)
   
    '//DB로 이동할 자료 Values 배열에 저장
    For i = 0 To cntField - 1
        For j = 0 To cntRecord - 1
            Values(j, i) = Range("A2").Offset(j, i)
        Next j
    Next i
    
    '//레코드별로 Insert(수정★★)
    '  - 필드 수 만큼 추가해서 진행
    '  - 공백대신 NULL값을 넣고자 할 경우 "''"을 "NULL"로 수정
    '  - 기본값을 넣으려면 "Default" & "," & _
    '  - TimeStamp를 넣으려면: "CURRENT_TIMESTAMP()" & ");"
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
                       
        affectedCount = executeSQL("insertDataToDB_ChurchList_Admin", tableNM, strSql, , "엑셀 데이터 밀어넣기")
        k = k + affectedCount
    Next i
    MsgBox k & "개의 레코드가 '" & tableNM & "'에 추가되었습니다.", vbInformation, banner
    
    '//연결 끊기
    disconnectDB
End Sub
