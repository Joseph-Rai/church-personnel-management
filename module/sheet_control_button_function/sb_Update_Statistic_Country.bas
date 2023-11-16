Attribute VB_Name = "sb_Update_Statistic_Country"
Option Explicit
Dim rngStart As Range '--//리포트 시작셀
Dim rngEnd As Range '--//리포트 종료셀
Dim LISTDATA() As String '--//DB에서 받아온 rs를 배열로 저장
Dim LISTFIELD() As String '--//DB에서 받아온 rs의 필드를 배열로 저장
Dim cntRecord As Integer '--//DB에서 받아온 레코드의 개수
Dim strSql As String '--//SQL 쿼리문
Const TB1 As String = "op_system.v_statistic_by_country"
Sub Sheet_Init()

    Dim R As Long
    Dim i As Long

    If connIP = Empty Then setGlobalVariant '//전역변수 재설정
    If checkLogin = 0 Then f_login.Show '//로그인체크
    
    '--//변수설정
    Set rngStart = Range("A3")
    Set rngEnd = Range("A:A").Find("합계", lookat:=xlWhole)
    
    '--//기존 데이터 모두 삭제
    If rngEnd.Row - rngStart.Row > 2 Then
        Range(Cells(rngStart.Row + 2, "A"), Cells(rngEnd.Row - 1, "A")).EntireRow.Delete
    End If
    
    '--//데이터 불러오기
    strSql = "SELECT * FROM " & TB1 & ";"
    Call makeListData(strSql, TB1)
    
    '--//데이터 수만큼 행추가
    rngStart.Offset(2).Resize(cntRecord - 1).EntireRow.Insert Shift:=xlDown
    
    '--//불러온 데이터 삽입
    rngStart.Offset(1, 1).Resize(cntRecord, UBound(LISTFIELD) + 1) = LISTDATA
    
    '--//서식복사
    rngStart.Offset(1).EntireRow.Copy
    rngStart.Offset(1).Resize(cntRecord).EntireRow.PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
    
    '--//A열 번호 채우기
    rngStart.Offset(1).Resize(cntRecord).Formula = "=ROW()-3"
    
    '--//텍스트 형식의 숫자 숫자형식으로 바꾸기
    Range("A2").Copy
    rngStart.Offset(1, 2).Resize(cntRecord, UBound(LISTFIELD) + 1).PasteSpecial Paste:=xlPasteValues, operation:=xlPasteSpecialOperationAdd
    Application.CutCopyMode = False
    
    '--//0은 보이지 않게 처리
    rngStart.Offset(1, 2).Resize(cntRecord, UBound(LISTFIELD) + 1).Replace 0, "", lookat:=xlWhole
    
    '--//합계 행 수식처리
    Set rngEnd = Range("A:A").Find("합계", lookat:=xlWhole)
    rngEnd.Offset(, 2).Resize(, UBound(LISTFIELD) + 1).FormulaR1C1 = "=SUM(R[" & rngStart.Row - rngEnd.Row + 1 & "]C:R[-1]C)"
    
    '--//rngStart 선택 후 종료
    rngStart.Select
    
End Sub
Private Sub makeListData(ByVal strSql As String, ByVal tableNM As String)

    Dim i As Integer, j As Integer
    
    '//DB에서 자료 호출하여 레코드셋에 반환
    connectTaskDB
    callDBtoRS "makeListData", tableNM, strSql, "sb_Update_Statistic_Country"
    
    '//레코드셋의 데이터를 listData 배열에 반환
    If Not rs.EOF Then
        ReDim LISTDATA(0 To rs.RecordCount - 1, 0 To rs.Fields.Count - 1) '//DB에서 반환할 배열의 크기 지정: 레코드셋의 레코드 수, 필드 수
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
    
    '--//필드명 배열 채우기
    ReDim LISTFIELD(0 To rs.Fields.Count - 1)
    For i = 0 To rs.Fields.Count - 1
        LISTFIELD(i) = rs.Fields(i).Name
    Next i
    cntRecord = rs.RecordCount '--//레코드 수 검토
    disconnectALL
    
End Sub
