Attribute VB_Name = "sb_DBtoExcel"
Option Explicit
'-----------------------------------------------
'  DB 자료를 엑셀에 반환
'    - excel_export(파일오픈여부)
'    - 워크북을 만들어서 바탕화면에 저장
'    - 파일오픈이 true면 저장 후 열기
'------------------------------------------------
Sub excel_export() 'Optional FileOpen As Boolean = False)
    
    Dim tableNM As String, dbNM As String
    Dim strSql As String
    Dim i As Integer
    Dim fileNM As String
    Dim fileSNM As String
    Dim FileOpen As Boolean
    FileOpen = False
    
    '//db명 설정
    tableNM = "op_system.db_churchlist_custom" '//db명.테이블명 - 수정★★
    dbNM = "op_system" '//수정★★

    
    '//DB연결
    connectTaskDB
    
    '//Select문
    strSql = "SELECT " & _
                    "CONCAT(a.church_nm, IFNULL(b.church_nm, REPLACE(a.church_nm,' 본교회', ''))) AS 'id' " & _
                    ",a.* " & _
                    ",IFNULL(b.church_nm, REPLACE(a.church_nm,' 본교회', '')) AS 'main_church_name' " & _
                "FROM op_system.db_churchlist_custom a " & _
                "LEFT JOIN op_system.db_churchlist b " & _
                    "ON a.main_church_cd = b.church_sid " & _
                "ORDER BY a.sort_order DESC"
    
    
    '//SQL문 실행하고 조회된 자료를 레코드셋에 담음
    Call callDBtoRS("excel_export", tableNM, strSql)
    If rs.EOF = True Then
        MsgBox "조회 조건에 맞는 자료가 없습니다.", vbInformation, banner
        disconnectALL
        Exit Sub
    End If
    
    '//엑셀로 자료 내보내기
    Optimization
    Workbooks.Add
    For i = i To rs.Fields.Count - 1
        Cells(1, 1).Offset(0, i).Value = rs.Fields(i).Name
    Next i
    Cells(2, 1).CopyFromRecordset rs
    Cells(1.1).CurrentRegion.Columns.AutoFit
    fileNM = GetDesktopPath() & tableNM & "(" & Format(Date, "yyyymmdd") & "_" & Format(Time, "hhmm") & ")" & ".xlsx"
    Application.DisplayAlerts = False
        ActiveWorkbook.SaveAs FileName:=fileNM
    Application.DisplayAlerts = True
    If FileOpen = False Then
        ActiveWorkbook.Close
    End If
    Normal
    
    '//결과보고, 마무리
    fileSNM = Right(fileNM, Len(fileNM) - InStrRev(fileNM, "\"))
    MsgBox "바탕화면에 파일이 생성되었습니다." & vbNewLine & vbNewLine & _
        " - 파일이름: " & fileSNM, vbInformation, banner
    
    disconnectALL
End Sub

Sub Excel_Export_For_Atten_Update() 'Optional FileOpen As Boolean = False)
    
    Dim tableNM As String, dbNM As String
    Dim strSql As String
    Dim i As Integer
    Dim fileNM As String
    Dim fileSNM As String
    Dim FileOpen As Boolean
    FileOpen = False
    
    '//db명 설정
    tableNM = "op_system.db_churchlist_custom" '//db명.테이블명 - 수정★★
    dbNM = "op_system" '//수정★★

    
    '//DB연결
    connectTaskDB
    
    '//Select문
    strSql = "SELECT " & _
                    "CONCAT(a.church_nm, IFNULL(b.church_nm, REPLACE(a.church_nm,' 본교회', ''))) AS 'id' " & _
                    ",a.* " & _
                    ",IFNULL(b.church_nm, REPLACE(a.church_nm,' 본교회', '')) AS 'main_church_name' " & _
                "FROM op_system.db_churchlist_custom a " & _
                "LEFT JOIN op_system.db_churchlist b " & _
                    "ON a.main_church_cd = b.church_sid " & _
                "ORDER BY a.sort_order DESC"
    
    
    '//SQL문 실행하고 조회된 자료를 레코드셋에 담음
    Call callDBtoRS("excel_export", tableNM, strSql)
    If rs.EOF = True Then
        MsgBox "조회 조건에 맞는 자료가 없습니다.", vbInformation, banner
        disconnectALL
        Exit Sub
    End If
    
    '//엑셀로 자료 내보내기
    Optimization
    Workbooks.Add
    For i = i To rs.Fields.Count - 1
        Cells(1, 1).Offset(0, i).Value = rs.Fields(i).Name
    Next i
    Cells(2, 1).CopyFromRecordset rs
    Cells(1.1).CurrentRegion.Columns.AutoFit
    fileNM = GetDesktopPath() & tableNM & "(" & Format(Date, "yyyymmdd") & "_" & Format(Time, "hhmm") & ")" & ".xlsx"
    Application.DisplayAlerts = False
        ActiveWorkbook.SaveAs FileName:=fileNM
    Application.DisplayAlerts = True
    If FileOpen = False Then
        ActiveWorkbook.Close
    End If
    Normal
    
    '//결과보고, 마무리
    fileSNM = Right(fileNM, Len(fileNM) - InStrRev(fileNM, "\"))
    MsgBox "바탕화면에 파일이 생성되었습니다." & vbNewLine & vbNewLine & _
        " - 파일이름: " & fileSNM, vbInformation, banner
    
    disconnectALL
End Sub

'-----------------------------------------------
'  DB연결
'    - connectDB(서버 IP, 스키마, ID, PW)
'-----------------------------------------------
Sub connectDB(argIP As String, argDB As String, argId As String, argPW As String)
    Set conn = New ADODB.Connection
    conn.ConnectionString = "Driver={" & ODBCDriver & "};Server=" & argIP & ";Port=3306;Database=" & argDB & ";User=" & argId & ";Password=" & argPW & ";Option=2;"
    conn.Open
End Sub

