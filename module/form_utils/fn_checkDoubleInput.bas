Attribute VB_Name = "fn_checkDoubleInput"
Option Explicit

'--------------------------------------------------------------------------------------------------
'  특정 필드 중복 검토
'    - checkDoubleInput(필드명, 데이터, DB테이블명, 유저폼명, 이전데이터) True: 중복
'    - 수정의 경우 이전 데이터를 넣어서 중복체크 우회
'--------------------------------------------------------------------------------------------------
Public Function checkDoubleInput(fieldNM As String, Data As Variant, tableNM As String, formNM As String, Optional ByVal beforeData As Variant = Empty) As Boolean
    Dim strSql As String
    Dim cntRecord As Integer
    
    '//특정 필드에 특정 데이터 갯수 반환
    Call connectTaskDB
    strSql = "SELECT COUNT(" & fieldNM & ") record_cnt " & _
                  "FROM " & tableNM & " " & _
                  "WHERE " & fieldNM & " = " & SText(Data) & ";"
    callDBtoRS "checkDoubleInput", tableNM, strSql, formNM
    If rs.EOF = True Then
        cntRecord = 0
    Else
        cntRecord = rs("record_cnt").Value
    End If
    Call disconnectALL
    
    '//중복 입력 검증
    If beforeData <> Empty And beforeData = Data Then Exit Function '//수정의 경우 기존 데이터와 동일해서 통과
    If cntRecord >= 1 Then
        checkDoubleInput = True
    Else
        checkDoubleInput = False
    End If
End Function

'----------------------------------------------------------------------------------------------------------------------------
'  관계 데이터 중복 검토
'    - checkDoubleInput(데이터유형, 필드명1, 필드명2, 데이터1, 데이터2, DB테이블명, 유저폼이름) True: 중복
'----------------------------------------------------------------------------------------------------------------------------
Public Function checkDoubleInput2(dataType As Integer, fieldNM1 As String, fieldNM2 As String, Data1 As Variant, Data2 As Variant, _
                                                      tableNM As String, formNM As String) As Boolean
    Dim strSql As String
    Dim cntRecord As Integer
    
    '//특정 필드에 특정 데이터 갯수 반환
    Call connectTaskDB
    strSql = "SELECT COUNT(*) record_cnt " & _
                  "FROM " & tableNM & " " & _
                  "WHERE " & fieldNM1 & " = " & SText(Data1) & " AND " & _
                  fieldNM2 & " = " & SText(Data2) & ";"
    callDBtoRS "checkDoubleInput2", tableNM, strSql, formNM
    If rs.EOF = True Then
        cntRecord = 0
    Else
        cntRecord = rs("record_cnt").Value
    End If
    Call disconnectALL
    
    '//중복 입력 검증
    Select Case dataType
        Case 1 '//신규입력
            If cntRecord > 0 Then
                checkDoubleInput2 = True
            Else
                checkDoubleInput2 = False
            End If
        Case 2 '//수정입력
            If cntRecord >= 2 Then
                checkDoubleInput2 = True
            Else
                checkDoubleInput2 = False
            End If
        Case 4 '//완전삭제
            checkDoubleInput2 = False
    End Select
End Function

'------------------------------------------------------------------------------------------------------------------------------------------------
'  기간 관계 데이터 중복 검토
'    - checkDoubleInput3( 데이터유형, 필드명1, 필드명2, 데이터1, 데이터2, 시작일, 종료일, DB테이블명, 유저폼이름) True: 중복
'------------------------------------------------------------------------------------------------------------------------------------------------
Public Function checkDoubleInput3(dataType As Integer, fieldNM1 As String, fieldNM2 As String, Data1 As Variant, Data2 As Variant, _
                                                      START_DT As Date, END_DT As Date, _
                                                      tableNM As String, formNM As String) As Boolean
    Dim strSql As String
    Dim cntRecord As Integer
    
    '//특정 필드에 특정 데이터 갯수 반환
    Call connectTaskDB
    strSql = "SELECT COUNT(*) record_cnt " & _
                  "FROM " & tableNM & " " & _
                  "WHERE " & fieldNM1 & " = " & SText(Data1) & " AND " & _
                  fieldNM2 & " = " & SText(Data2) & " AND " & _
                  "start_dt <= " & SText(END_DT) & " AND " & _
                  "end_dt >= " & SText(START_DT) & ";"
    callDBtoRS "checkDoubleInput3", tableNM, strSql, formNM
    If rs.EOF = True Then
        cntRecord = 0
    Else
        cntRecord = rs("record_cnt").Value
    End If
    
    Call disconnectALL
    
    '//중복 입력 검증
    Select Case dataType
        Case 1 '//신규입력
            If cntRecord > 0 Then
                checkDoubleInput3 = True
            Else
                checkDoubleInput3 = False
            End If
        Case 2 '//수정입력
            If cntRecord > 1 Then
                checkDoubleInput3 = True
            Else
                checkDoubleInput3 = False
            End If
        Case 4 '//완전삭제
            checkDoubleInput3 = False
    End Select
End Function

'-------------------------------------------------------------------------------------------------------------------------
'  기간 데이터 중복 검토
'    - checkDoubleInput4( 데이터유형, 필드명, 데이터, 시작일, 종료일, DB테이블명, 유저폼이름) True: 중복
'-------------------------------------------------------------------------------------------------------------------------
Public Function checkDoubleInput4(dataType As Integer, fieldNM As String, Data As Variant, _
                                                      START_DT As Date, END_DT As Date, _
                                                      tableNM As String, formNM As String) As Boolean
    Dim strSql As String
    Dim cntRecord As Integer
    
    '//특정 필드에 특정 데이터 갯수 반환
    Call connectTaskDB
    strSql = "SELECT COUNT(*) record_cnt " & _
                  "FROM " & tableNM & " " & _
                  "WHERE " & fieldNM & " = " & SText(Data) & " AND " & _
                  "start_dt <= " & SText(END_DT) & " AND " & _
                  "end_dt >= " & SText(START_DT) & ";"
    callDBtoRS "checkDoubleInput4", tableNM, strSql, formNM
    If rs.EOF = True Then
        cntRecord = 0
    Else
        cntRecord = rs("record_cnt").Value
    End If
    
    Call disconnectALL
    
    '//중복 입력 검증
    Select Case dataType
        Case 1 '//신규입력
            If cntRecord > 0 Then
                checkDoubleInput4 = True
            Else
                checkDoubleInput4 = False
            End If
        Case 2 '//수정입력
            If cntRecord > 1 Then
                checkDoubleInput4 = True
            Else
                checkDoubleInput4 = False
            End If
        Case 4 '//완전삭제
            checkDoubleInput4 = False
    End Select
End Function

