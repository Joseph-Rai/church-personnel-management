VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} f_login 
   Caption         =   "로그인"
   ClientHeight    =   2835
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4500
   OleObjectBlob   =   "f_login.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "f_login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LISTDATA() As String '--//DB에서 받아온 rs를 배열로 저장
Dim LISTFIELD() As String '--//DB에서 받아온 rs의 필드를 배열로 저장
Dim cntRecord As Integer '--//DB에서 받아온 레코드의 개수

'------------------------------------
'  로그인창 종료 시 로그인검증
'------------------------------------
Private Sub UserForm_Terminate()
    If checkLogin = 0 Then
        MsgBox "로그인 정보가 확인되지 않았습니다." & space(7), vbCritical, banner
'        ThisWorkbook.Close savechanges:=False
    End If
    disconnectALL
End Sub

'------------------------------------------------------
'  로그인 폼(common)
'  - ID, PW체크
'  - 프로그램 버전 체크
'  - IP체크
'------------------------------------------------------
Private Sub UserForm_Initialize()
On Error GoTo ErrHandler
    Dim strSql As String
       
    '//기초설정
    Me.cmd_close.Width = 0
    Me.cmd_close.Cancel = True
    Me.Caption = banner
    txt_ID.Value = Application.UserName
    Me.lbl_pv = programv
        
    '//등록된 사용자 체크
    If checkUserNm(txt_ID.Value) = False Then
        MsgBox "등록된 사용자가 아닙니다." & space(7) & vbNewLine & _
                "로그인 창에서 이름을 변경해 주세요." & space(7) & vbNewLine & _
                "사용자 등록이 필요하면 관리자에게 요청해 주세요.", vbInformation, banner
        GoTo n
    End If
    
    '//비밀번호 설정 여부 체크
    Call checkInitialPW
n:
    txt_PW.SetFocus
    Exit Sub
ErrHandler:
    End
End Sub

'---------------------------------------------------------------------------------------
'  등록된 사용자 체크
'    - txt_ID에 입력된 사용자가 등록된 사용자인지 검토하여 true / false 값 반환
'---------------------------------------------------------------------------------------
Private Function checkUserNm(ByVal argUserNM As String) As Boolean
    Dim strSql As String
    
    connectCommonDB
    strSql = "SELECT * FROM common.users WHERE user_id = (SELECT user_id FROM common.users WHERE user_nm = " & SText(Me.txt_ID.Value) & ");"
    callDBtoRS "checkUserNm", "common.users", strSql, "f_login", "사용자확인"
    
    If rs.RecordCount = 0 Then
        checkUserNm = False
    Else
        checkUserNm = True
    End If
    
    disconnectALL
End Function

'---------------------------------------
'  txt_ID에서 exit 할 경우
'    - 사용자 이름 등록여부 체크
'    - 비밀번호 초기 설정 여부 체크
'---------------------------------------
Private Sub txt_ID_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If txt_ID = Empty Then
        Exit Sub
    End If
    
    '//사용자 이름 등록 여부 체크
    If checkUserNm(txt_ID.Value) = False Then
        MsgBox "등록된 사용자가 아닙니다." & space(7) & vbNewLine & _
                "로그인 창에서 이름을 변경해 주세요." & space(7) & vbNewLine & _
                "사용자 등록이 필요하면 관리자에게 요청해 주세요.", vbInformation, banner
        txt_ID.SetFocus
        Exit Sub
    End If
    If txt_ID.Value <> Application.UserName Then
        Application.UserName = txt_ID.Value
    End If
    
    '//비밀번호 초기 설정 여부 체크
    Call checkInitialPW
    
End Sub

'----------------------------------------------------------------------------------------
'  등록된 사용자의 경우 비밀번호가 설정되어 있었는지 체크하고 설정하도록 진행
'----------------------------------------------------------------------------------------
Private Sub checkInitialPW()
    Dim strSql As String
    Dim strPW As Integer
    
    connectCommonDB
    strSql = "SELECT pw_initialize FROM common.users WHERE user_id = (SELECT user_id FROM common.users WHERE user_nm = " & SText(Me.txt_ID.Value) & ");"
    callDBtoRS "checkInitialPW", "common.users", strSql, "f_login", "사용자비밀번호등록여부조회"
    
    strPW = rs("pw_initialize").Value
    If strPW = 1 Then '//PW 입력 이력이 없으면 PW 설정
        MsgBox "비밀번호가 설정되어 있지 않습니다." & vbNewLine & _
                     "비빌번호 설정화면으로 이동합니다.", vbInformation, banner
        Call registerNewPW
    End If
    disconnectALL
End Sub

'--------------------------------------------------------------------------------------------------------------------
'  사용자는 등록된 사용자이지만 비밀번호 설정이 안되어 있는 경우(pw_initialize = 1) 신규비밀번호 등록
'--------------------------------------------------------------------------------------------------------------------
Private Sub registerNewPW()
    Dim strSql As String
    Dim strPW As Integer
    Dim USER_PW As Variant
    Dim affectedCount As Long
    
    '//비밀번호 입력 받기
    Do
        USER_PW = InputBoxPW("신규 비밀번호를 대소문자를 구분하여 4자리 이상으로 설정해 주세요.", banner)
    Loop Until USER_PW <> Empty And Len(USER_PW) > 3
    
    '//비밀번호 등록(암호화)
    connectCommonDB
    strSql = "UPDATE common.users SET user_pw = SHA2(" & SText(USER_PW) & ", 512) WHERE user_id = (SELECT user_id FROM common.users WHERE user_nm = " & SText(Me.txt_ID.Value) & ");"
    affectedCount = executeSQL("registerNewPW", "common.users", strSql, "f_login", "초기비밀번호설정")
    If affectedCount > 0 Then
         writeLog "registerNewPW", "common.users", strSql, 0, Me.Name, "사용자PW등록", affectedCount
    End If
    disconnectALL
    
    '//비밀번호 등록 확인
    If affectedCount = 0 Then
        MsgBox "비밀번호가 설정되지 않았습니다." & space(7) & vbNewLine & _
            "관리자에게 문의하여 주시기 바랍니다.", vbInformation, banner
    Else
         '//비밀번호 초기화 비활성화
         connectCommonDB
        strSql = "UPDATE common.users SET pw_initialize = 0 WHERE user_id = (SELECT user_id FROM common.users WHERE user_nm = " & SText(Me.txt_ID.Value) & ");"
        executeSQL "registerNewPW", "common.users", strSql, "f_login", "비밀번호초기화비활성화"
        writeLog "registerNewPW", "common.users", strSql, 0, Me.Name, "비밀번호초기화비활성화", 1
        MsgBox "비밀번호 설정이 완료되었습니다." & space(7), vbInformation, banner
    End If
    disconnectALL
End Sub

'--------------------------------------
'  확인버튼 시
'    - 사용자 이름 등록 여부 체크
'    - 프로그램 최신버전 확인
'    - IP체크
'    - 비밀번호 체크
'    - 환영인사
'---------------------------------------
Private Sub cmd_query_Click()
    Dim strSql As String
    Dim affectedCount As Long
    Dim ipRng As Integer
    
    '//사용자 이름 등록 여부 체크
    If txt_ID = Empty Then
        MsgBox "사용자 이름을 입력하세요.", vbInformation, banner
        Exit Sub
    End If
    If checkUserNm(txt_ID.Value) = False Then
        MsgBox "등록된 사용자가 아닙니다." & space(7) & vbNewLine & _
                "로그인 창에서 이름을 변경하세요." & space(7) & vbNewLine & _
                "사용자 등록이 필요하면 관리자에게 요청하세요.", vbInformation, banner
        txt_ID.SetFocus
        Exit Sub
    End If
    If txt_ID.Value <> Application.UserName Then
        Application.UserName = txt_ID.Value
    End If
    
    '//비밀번호 입력 여부 체크
    If txt_PW = Empty Then
        MsgBox "비밀번호를 입력하세요.", vbInformation, banner
        txt_PW.SetFocus
        Exit Sub
    End If
    
    '//프로그램 버전 확인
    connectCommonDB
    strSql = "SELECT programv FROM common.users WHERE user_id = (SELECT user_id FROM common.users WHERE user_nm = " & SText(Me.txt_ID.Value) & ");"
    callDBtoRS "cmd_query_Click", "common.users", strSql, Me.Name, "프로그램버전 확인"
    If UCase(rs("programv").Value) <> UCase(programv) Then
        MsgBox "사용하려는 프로그램이 최신버전이 아닙니다." & vbNewLine & _
                     "프로그램 오류 방지를 위해 마이디스크에서 최신버전으로 복사 후 사용해 주세요.", vbInformation, banner
        disconnectALL
        cmd_close_Click
    End If
    
    '//IP확인
    'IP입력 여부 확인
    strSql = "SELECT user_ip FROM common.users WHERE user_id = (SELECT user_id FROM common.users WHERE user_nm = " & SText(Me.txt_ID.Value) & ");"
    callDBtoRS "cmd_query_Click", "common.users", strSql, "f_login", "사용자IP확인"
    
    If IsNull(rs("user_ip").Value) = True Or Len(rs("user_ip")) = 0 Then '최초 접속이면 IP 기록
        If MsgBox("현재의 PC를 사용자의 PC로 등록합니다." & vbNewLine & _
                         "등록된 PC외 다른 PC에서는 프로그램 사용이 제한됩니다." & vbNewLine & _
                         "진행하겠습니까?", vbQuestion + vbYesNo, banner) = vbNo Then
            disconnectALL
            Exit Sub
        Else
            '[신규IP넣기]
            strSql = "UPDATE common.users SET user_ip = " & SText(GetLocalIPaddress) & " WHERE user_id = (SELECT user_id FROM common.users WHERE user_nm = " & SText(Me.txt_ID.Value) & ");"
            executeSQL "cmd_query_Click", "common.users", strSql, Me.Name, "사용자IP기록"
            writeLog "cmd_query_Click", "common.users", strSql, 0, Me.Name, "사용자IP기록", 1
        End If
    Else '최초 접속 아닐 경우 IP 체크
        strSql = "SELECT user_ip FROM common.users WHERE user_id = (SELECT user_id FROM common.users WHERE user_nm = " & SText(Me.txt_ID.Value) & ");"
        callDBtoRS "cmd_query_Click", "common.users", strSql, Me.Name, "사용자IP확인"
        If rs("user_ip").Value <> GetLocalIPaddress Then
            MsgBox "이 프로그램은 허용된 PC에서만 사용 가능합니다." & vbNewLine & _
                         "사용자의 PC 등록 변경이 필요하면 관리자에게 요청하세요." & vbNewLine & _
                         "프로그램을 종료합니다.", vbInformation, banner
            disconnectALL
            cmd_close_Click
        End If
    End If
    
    '//비밀번호 맞는 지 검토
    If checkPW(txt_PW.Value) = True Then
        '로그인 값 1, 글로벌 변수 설정
        checkLogin = 1
        setGlobalVariant
        '접속시간 업데이트
        connectCommonDB
        strSql = "UPDATE common.users SET time_stamp = CURRENT_TIMESTAMP() WHERE user_id = " & USER_ID & ";"
        executeSQL "cmd_query_Click", "common.users", strSql, Me.Name, "사용자접속시간업데이트"
        disconnectALL
        '환영인사
        MsgBox Application.UserName & "님 복많이 받으세요." & space(7) & vbNewLine & vbNewLine & _
                 "오늘은 " & Format(Date, "YYYY-MM-DD") & "일 입니다." & vbNewLine & _
                 "오늘도 ANIMO!", vbInformation, banner
        'today에 오늘 날짜 설정
        today = Date
        Unload Me
    Else
        '비밀번호가 다르면 다시 입력
        MsgBox "비밀번호가 틀렸습니다." & space(7) & vbNewLine & _
            "비밀번호를 다시 입력하여 주세요.", vbInformation, banner
        txt_PW.Value = Empty
        txt_PW.SetFocus
        Exit Sub
    End If
    
    '--//선지자 상세정보 시트 조정
    Dim sql As String
    
    sql = "SELECT b.authority FROM op_system.a_auth_table a" & _
          " LEFT JOIN op_system.a_authority b " & _
          "     ON a.authority_id = b.id" & _
          " WHERE a.user_id = " & USER_ID & ";"
    
    Dim i As Integer, j As Integer
    
    '//DB에서 자료 호출하여 레코드셋에 반환
    connectTaskDB
    callDBtoRS "makeListData", "op_system.a_auth_table", sql, "f_login"
    
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
    
    If cntRecord <= 0 Then
        Exit Sub
    End If
    
    If IsInArray("PSTAFF_DETAIL_SHEET_VIEW", LISTDATA) <> -1 Then
        Sheets("선지자 상세정보").Visible = True
    End If
    
    If IsInArray("A3_APPOINTMENT_FORM", LISTDATA) <> -1 Then
        Sheets("A3인사발령").Visible = True
    End If
    
    
End Sub

'------------------------------------------------------------------------
'  입력된 비밀번호가 맞는지 틀린지 검토하여 true / false 값 반환
'------------------------------------------------------------------------
Private Function checkPW(ByVal argPW As String) As Boolean
    Dim strSql As String
    Dim strPW As Variant
    
    connectCommonDB
    strSql = "SELECT user_pw FROM common.users WHERE user_id = (SELECT user_id FROM common.users WHERE user_nm = " & SText(Application.UserName) & ");"
    callDBtoRS "checkPW", "common.users", strSql, "f_login"
    
    strPW = rs("user_pw").Value
    If strPW <> to_SHA512(argPW) Then
        checkPW = False
    Else
        checkPW = True
    End If
End Function

Private Sub cmd_close_Click()
    Unload Me
End Sub

'---------------------------------------
'  비밀번호 변경
'    - 기존 비밀번호 확인
'    - 신규 비밀번호 입력
'---------------------------------------
Private Sub cmd_chgPW_Click()
    Dim oldPW As String
    Dim newPW As String
    
    If MsgBox("비밀번호를 변경하겠습니까?", vbQuestion + vbYesNo, banner) = vbNo Then
        Exit Sub
    Else
        oldPW = InputBoxPW("기존 비밀번호를 입력하세요.", banner)
        If oldPW = "" Then
            MsgBox "기존 비밀번호 입력이 취소되었습니다.", vbInformation, banner
            Exit Sub
        Else
            If checkPW(oldPW) = True Then
                registerNewPW
            Else
                MsgBox "기존 비밀번호가 일치하지 않습니다." & vbNewLine & _
                             "관리자에게 문의하여 주세요.", vbInformation, banner
            End If
        End If
    End If
End Sub
