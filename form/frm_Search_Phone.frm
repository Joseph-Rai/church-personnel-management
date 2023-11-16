VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Search_Phone 
   Caption         =   "연락처 열람 마법사"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7275
   OleObjectBlob   =   "frm_Search_Phone.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "frm_Search_Phone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cntlst1Col As Integer '//리스트 컬럼 수
Dim TB1 As String, TB2 As String, TB3 As String, TB4 As String, TB5 As String, TB6 As String, TB7 As String '//폼에 연결된 DB 테이블
Const strOrderBy As String = "currency_un ASC" '//DB에서 sort_order 필드
Dim caseSave As Integer '//1: 추가, 2: 수정, 3: 삭제(SUSPEND), 4: 완전삭제
Dim queryKey As Integer '//리스트 위치 반환에 사용될 id
Dim LISTDATA() As String '--//DB에서 받아온 rs를 배열로 저장
Dim LISTFIELD() As String '--//DB에서 받아온 rs의 필드를 배열로 저장
Dim cntRecord As Integer '--//DB에서 받아온 레코드의 개수
Dim strSql As String '--//SQL 쿼리문
Dim txtBox_Focus As MSForms.textBox

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdCopy_All_Click()
    
    Dim txtCopy As String
    
    If Not (Me.txtLandLine = "" And Me.txtWMCPhone = "" And Me.txtPhone_PStaff = "" And Me.txtPhone_Spouse = "") Then
        With Me.lstPStaff
            txtCopy = IIf(.List(.listIndex, 2) <> .List(.listIndex, 3), "교회명: " & .List(.listIndex, 3) & vbNewLine & "지교회명: " & Me.lblChurch.Caption, "교회명: " & Me.lblChurch.Caption)
            If Me.txtLandLine <> "" Then
                txtCopy = txtCopy & vbNewLine & vbNewLine & Me.lblLandLine.Caption & vbNewLine & Me.txtLandLine
            End If
            If Me.txtWMCPhone <> "" Then
                txtCopy = txtCopy & vbNewLine & vbNewLine & Me.lblWMCPhone.Caption & vbNewLine & Me.txtWMCPhone
            End If
            If Me.txtPhone_PStaff <> "" Then
                txtCopy = txtCopy & vbNewLine & vbNewLine & Me.txtName & "/" & Me.txtPosition & vbNewLine & Me.txtPhone_PStaff
            End If
            If Me.txtPhone_Spouse <> "" Then
                txtCopy = txtCopy & vbNewLine & vbNewLine & Me.txtName_Spouse & "/" & Me.txtPosition_Spouse & vbNewLine & Me.txtPhone_Spouse
            End If
            If Me.txtAddress <> "" Then
                txtCopy = txtCopy & vbNewLine & vbNewLine & Me.txtAddress
            End If
            
            CopyText (Trim(txtCopy))
        End With
    Else
        MsgBox "복사할 내용이 없습니다.", vbInformation
    End If
End Sub

Private Sub cmdCopy_LandLine_Click()
    If Me.txtLandLine <> "" Then
        With Me.lstPStaff
            CopyText (IIf(.List(.listIndex, 2) <> .List(.listIndex, 3), "교회명: " & .List(.listIndex, 3) & vbNewLine & "지교회명: " & Me.lblChurch.Caption, "교회명: " & Me.lblChurch.Caption) & vbNewLine & vbNewLine & Me.lblLandLine.Caption & vbNewLine & Me.txtLandLine)
        End With
    Else
        MsgBox "복사할 내용이 없습니다.", vbInformation
    End If
End Sub

Private Sub cmdCopy_pstaff_Click()
    If Me.txtPhone_PStaff <> "" Then
        With Me.lstPStaff
            CopyText (IIf(.List(.listIndex, 2) <> .List(.listIndex, 3), "교회명: " & .List(.listIndex, 3) & vbNewLine & "지교회명: " & Me.lblChurch.Caption, "교회명: " & Me.lblChurch.Caption) & vbNewLine & vbNewLine & Me.txtName & "/" & Me.txtPosition & vbNewLine & Me.txtPhone_PStaff)
        End With
    Else
        MsgBox "복사할 내용이 없습니다.", vbInformation
    End If
End Sub

Private Sub cmdCopy_Spouse_Click()
    If Me.txtPhone_Spouse <> "" Then
        With Me.lstPStaff
            CopyText (IIf(.List(.listIndex, 2) <> .List(.listIndex, 3), "교회명: " & .List(.listIndex, 3) & vbNewLine & "지교회명: " & Me.lblChurch.Caption, "교회명: " & Me.lblChurch.Caption) & vbNewLine & vbNewLine & Me.txtName_Spouse & "/" & Me.txtPosition_Spouse & vbNewLine & Me.txtPhone_Spouse)
        End With
    Else
        MsgBox "복사할 내용이 없습니다.", vbInformation
    End If
End Sub

Private Sub cmdCopy_WMCPhone_Click()
    If Me.txtWMCPhone <> "" Then
        With Me.lstPStaff
            CopyText (IIf(.List(.listIndex, 2) <> .List(.listIndex, 3), "교회명: " & .List(.listIndex, 3) & vbNewLine & "지교회명: " & Me.lblChurch.Caption, "교회명: " & Me.lblChurch.Caption) & vbNewLine & vbNewLine & Me.lblWMCPhone.Caption & vbNewLine & Me.txtWMCPhone)
        End With
    Else
        MsgBox "복사할 내용이 없습니다.", vbInformation
    End If
End Sub

Private Sub cmdEdit_Click()
    Dim result As T_RESULT
    Dim lifeNo As String
    
    '--//수정된 내용 있는지 체크
    strSql = makeSelectSQL(TB2)
    Call makeListData(strSql, TB2)
    If Me.txtLandLine = LISTDATA(0, 4) And Me.txtWMCPhone = LISTDATA(0, 5) And Me.txtPhone_PStaff = LISTDATA(0, 6) And Me.txtPhone_Spouse = LISTDATA(0, 7) And Me.txtAddress = LISTDATA(0, 18) Then
        Exit Sub
    End If
    
    '--//교회연락처 업데이트 SQL문 생성, 실행, 로그기록
    strSql = makeUpdateSQL(TB3)
    result.strSql = strSql
    connectTaskDB
    result.affectedCount = executeSQL("cmdEdit_Click", TB3, strSql, Me.Name, "교회연락처 업데이트")
    writeLog "cmdEdit_Click", TB3, strSql, 0, Me.Name, "교회연락처 업데이트", result.affectedCount
    disconnectALL
    
    '--//선지자연락처 업데이트 SQL문 생성, 실행, 로그기록
    lifeNo = Me.lstPStaff.List(Me.lstPStaff.listIndex)
    
    If Mid(lifeNo, InStr(5, lifeNo, "-") + 1, 1) = 1 Then
        strSql = makeUpdateSQL(TB4)
        result.strSql = strSql
        connectTaskDB
        result.affectedCount = executeSQL("cmdEdit_Click", TB4, strSql, Me.Name, "선지자연락처 업데이트")
        writeLog "cmdEdit_Click", TB4, strSql, 0, Me.Name, "선지자연락처 업데이트", result.affectedCount
        disconnectALL
    End If
    
    '--//사모연락처 업데이트 SQL문 생성, 실행, 로그기록
    If Mid(lifeNo, InStr(5, lifeNo, "-") + 1, 1) = 2 Then
        strSql = makeUpdateSQL(TB5)
        result.strSql = strSql
        connectTaskDB
        result.affectedCount = executeSQL("cmdEdit_Click", TB5, strSql, Me.Name, "사모연락처 업데이트")
        writeLog "cmdEdit_Click", TB5, strSql, 0, Me.Name, "사모연락처 업데이트", result.affectedCount
        disconnectALL
    End If
    
    '--//메세지박스
    MsgBox "업데이트 되었습니다.", , banner
    
    '--//리스트박스 초기화
    Call lstPStaff_Click
End Sub

Private Sub cmdExport_Click()

    Dim targetWB As Workbook
    Dim i As Long
    Dim lngColumnIndex As Long
    Dim arg As String
    
    Me.lblExport.Visible = True
    Me.optKo.Visible = False
    Me.optEn.Visible = False
    Me.Repaint
    
    '--//데이터 불러오기
    strSql = makeSelectSQL(TB6)
    Call makeListData(strSql, TB6)
    lngColumnIndex = IsInArray("관리부서", LISTFIELD, , rtnSequence)
   
    '--//새 워크북 생성 및 데이터 붙여넣기
    Set targetWB = Workbooks.Add
    Call Optimization
    With targetWB.Sheets(1)
        .Cells(3, "A").Resize(, UBound(LISTFIELD) + 1) = LISTFIELD
        .Cells(4, "A").Resize(cntRecord, UBound(LISTFIELD) + 1) = LISTDATA
        
    
    '--//1행 데이터 삽입 및 정리
        lngColumnIndex = getColumIndex("한글이름(직분)", "영문이름")
        .Cells(2, lngColumnIndex + 1) = "관 리 자"
        lngColumnIndex = getColumIndex("사모한글이름(직분)", "사모영문이름")
        .Cells(2, lngColumnIndex + 1) = "사 모"
    
    '--//서식정리
        '--//1~2행 병합처리
        lngColumnIndex = IsInArray("직책", LISTFIELD, , rtnSequence)
        For i = 1 To lngColumnIndex + 1
            .Cells(2, i).Resize(2).Merge
        Next
        lngColumnIndex = getColumIndex("한글이름(직분)", "영문이름")
        .Cells(2, lngColumnIndex + 1).Resize(, 2).Merge
        lngColumnIndex = getColumIndex("사모한글이름(직분)", "사모영문이름")
        .Cells(2, lngColumnIndex + 1).Resize(, 2).Merge
        
        '--//관리부서 감추기
        lngColumnIndex = IsInArray("관리부서", LISTFIELD, , rtnSequence)
        .Cells(1, "A").Offset(, lngColumnIndex).Resize(, (UBound(LISTFIELD) + 1) - (lngColumnIndex)).EntireColumn.Group
        
        '--//가운데정렬
        .Columns("A:A").Resize(, UBound(LISTFIELD) + 1).EntireColumn.HorizontalAlignment = xlCenter
        .Columns("A:A").Resize(, UBound(LISTFIELD) + 1).EntireColumn.VerticalAlignment = xlCenter
        
        '--//필드 글꼴 및 배경색 변경
        .Cells(2, "A").Resize(2, UBound(LISTFIELD) + 1).Interior.ThemeColor = xlThemeColorDark2
        .Cells(2, "A").Resize(2, UBound(LISTFIELD) + 1).Font.Bold = True
        
        '--//테두리
        '--전체
        lngColumnIndex = IsInArray("배우자전화번호", LISTFIELD, , rtnSequence)
        .Cells(2, "A").Resize(cntRecord + 2, lngColumnIndex + 1).Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Cells(2, "A").Resize(cntRecord + 2, lngColumnIndex + 1).Borders(xlEdgeTop).LineStyle = xlContinuous
        .Cells(2, "A").Resize(cntRecord + 2, lngColumnIndex + 1).Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Cells(2, "A").Resize(cntRecord + 2, lngColumnIndex + 1).Borders(xlEdgeRight).LineStyle = xlContinuous
        .Cells(2, "A").Resize(cntRecord + 2, lngColumnIndex + 1).Borders(xlInsideVertical).LineStyle = xlContinuous
        .Cells(2, "A").Resize(cntRecord + 2, lngColumnIndex + 1).Borders(xlInsideHorizontal).LineStyle = xlContinuous
        
        .Cells(2, "A").Resize(cntRecord + 2, lngColumnIndex + 1).Borders(xlEdgeLeft).Weight = xlHairline
        .Cells(2, "A").Resize(cntRecord + 2, lngColumnIndex + 1).Borders(xlEdgeTop).Weight = xlHairline
        .Cells(2, "A").Resize(cntRecord + 2, lngColumnIndex + 1).Borders(xlEdgeBottom).Weight = xlHairline
        .Cells(2, "A").Resize(cntRecord + 2, lngColumnIndex + 1).Borders(xlEdgeRight).Weight = xlHairline
        .Cells(2, "A").Resize(cntRecord + 2, lngColumnIndex + 1).Borders(xlInsideVertical).Weight = xlHairline
        .Cells(2, "A").Resize(cntRecord + 2, lngColumnIndex + 1).Borders(xlInsideHorizontal).Weight = xlHairline
        
        .Cells(2, "A").Resize(cntRecord + 2, lngColumnIndex + 1).Borders(xlEdgeLeft).Weight = xlMedium
        .Cells(2, "A").Resize(cntRecord + 2, lngColumnIndex + 1).Borders(xlEdgeTop).Weight = xlMedium
        .Cells(2, "A").Resize(cntRecord + 2, lngColumnIndex + 1).Borders(xlEdgeBottom).Weight = xlMedium
        .Cells(2, "A").Resize(cntRecord + 2, lngColumnIndex + 1).Borders(xlEdgeRight).Weight = xlMedium
        
        '--중간 세로구분선
        lngColumnIndex = getColumIndex("한글이름(직분)", "영문이름")
        .Cells(2, lngColumnIndex + 1).Resize(cntRecord + 2, 3).Borders(xlEdgeLeft).Weight = xlMedium
        .Cells(2, lngColumnIndex + 1).Resize(cntRecord + 2, 3).Borders(xlEdgeRight).Weight = xlMedium
        .Cells(2, lngColumnIndex + 1).Resize(cntRecord + 2, 3).Borders(xlInsideVertical).Weight = xlHairline
        .Cells(2, lngColumnIndex + 1).Resize(cntRecord + 2, 3).Borders(xlInsideHorizontal).Weight = xlHairline
        
        '--필드
        lngColumnIndex = IsInArray("배우자전화번호", LISTFIELD, , rtnSequence)
        .Cells(2, "A").Resize(2, lngColumnIndex + 1).Borders(xlEdgeBottom).Weight = xlMedium
        
        '--//전화번호 컬럼 배경색
        lngColumnIndex = IsInArray("인터넷전화", LISTFIELD, , rtnSequence)
        .Cells(4, lngColumnIndex + 1).Resize(cntRecord, 2).Interior.ThemeColor = xlThemeColorAccent3
        .Cells(4, lngColumnIndex + 1).Resize(cntRecord, 2).Interior.TintAndShade = 0.799981688894314
        .Cells(1, lngColumnIndex + 1).Resize(, 2).EntireColumn.ColumnWidth = 22
        
        lngColumnIndex = IsInArray("선지자전화번호", LISTFIELD, , rtnSequence)
        .Cells(4, lngColumnIndex + 1).Resize(cntRecord).Interior.ThemeColor = xlThemeColorAccent4
        .Cells(4, lngColumnIndex + 1).Resize(cntRecord).Interior.TintAndShade = 0.799981688894314
        .Cells(1, lngColumnIndex + 1).EntireColumn.ColumnWidth = 22
        
        lngColumnIndex = IsInArray("배우자전화번호", LISTFIELD, , rtnSequence)
        .Cells(4, lngColumnIndex + 1).Resize(cntRecord).Interior.ThemeColor = xlThemeColorAccent2
        .Cells(4, lngColumnIndex + 1).Resize(cntRecord).Interior.TintAndShade = 0.799981688894314
        .Cells(1, lngColumnIndex + 1).EntireColumn.ColumnWidth = 22
        
        '--//열너비, 행높이 조정
        Columns("A:A").Resize(, UBound(LISTFIELD) + 1).EntireColumn.AutoFit
        Columns("A:A").Resize(cntRecord + 2).EntireRow.AutoFit
        'Columns("A:A").Resize(cntRecord + 2).RowHeight = 24
        
        Call Normal
        Call Optimization
        
        '--//본교회 글꼴 및 배경 스타일
        Dim temp As Long
        Dim temp2 As Long
        lngColumnIndex = IsInArray("본교회코드", LISTFIELD, , rtnSequence)
        temp = IsInArray("인터넷전화", LISTFIELD, , rtnSequence)
        temp2 = IsInArray("지교회코드", LISTFIELD, , rtnSequence)
        For i = 4 To 3 + cntRecord
            If Cells(i, lngColumnIndex + 1) <> Cells(i - 1, lngColumnIndex + 1) Then
                Cells(i, "A").EntireRow.Font.Bold = True
                Cells(i, "A").Resize(, lngColumnIndex + 1).Interior.color = 13434879
            Else
                If InStr(Cells(i, temp2 + 1), "MC") > 0 Then
                    Cells(i, temp + 1).Resize(, 2).ClearContents
                End If
            End If
            If Cells(i, "A").EntireRow.RowHeight < 24 Then
                Cells(i, "A").EntireRow.RowHeight = 24
            End If
        Next
        
        '--//부서명 작성
        Cells(1, "A").EntireRow.RowHeight = 25
        strSql = "SELECT dept_nm FROM db_ovs_dept WHERE dept_id = " & SText(USER_DEPT) & ";"
        Call makeListData(strSql, "op_system.db_ovs_dept")
        Cells(1, "C") = LISTDATA(0, 0) & " 연락처"
        Cells(1, "C").Font.Bold = True
        Cells(1, "C").Font.Size = 16
        Cells(1, "D") = Format(Now(), "yyyy-mm")
        Cells(1, "D").Font.Bold = True
        Cells(1, "D").Font.Size = 16
        Cells(1, "D").Font.ThemeColor = xlThemeColorAccent2
        
        ActiveSheet.Outline.ShowLevels RowLevels:=0, ColumnLevels:=1
        
        '--//프린트영역 설정
        With ActiveSheet.PageSetup
            .PrintTitleRows = "$2:$3"
            .PrintTitleColumns = ""
        End With
        ActiveSheet.PageSetup.PrintArea = ""
        With ActiveSheet.PageSetup
            .LeftMargin = Application.InchesToPoints(0.25)
            .RightMargin = Application.InchesToPoints(0.25)
            .TopMargin = Application.InchesToPoints(0.75)
            .BottomMargin = Application.InchesToPoints(0.75)
            .HeaderMargin = Application.InchesToPoints(0.3)
            .FooterMargin = Application.InchesToPoints(0.3)
            .Orientation = xlLandscape
            .CenterFooter = "&N페이지 중 &P페이지"
        End With
        ActiveWindow.View = xlPageBreakPreview
        ActiveSheet.VPageBreaks(1).DragOff Direction:=xlToRight, RegionIndex:=1
        
        Call sbClearVariant
    End With
    
    Call Normal
    Me.lblExport.Visible = False
    Me.optKo.Visible = True
    Me.optEn.Visible = True
    MsgBox "출력물이 생성되었습니다."
    
End Sub

Private Function getColumIndex(arg1 As String, arg2 As String)

    Dim lngColumnIndex As Long
    Dim lngColumnIndex_KO As Long
    Dim lngColumnIndex_EN As Long

    lngColumnIndex_KO = IsInArray(arg1, LISTFIELD, , rtnSequence)
    lngColumnIndex_EN = IsInArray(arg2, LISTFIELD, , rtnSequence)
    lngColumnIndex = WorksheetFunction.Max(lngColumnIndex_KO, lngColumnIndex_EN)
    
    getColumIndex = lngColumnIndex

End Function

Private Sub lstPStaff_Click()
    
    Dim filePath As String
    Dim FileName As String
    Dim lifeNo As String
    
    '--//이력목록 데이터 채우기
    strSql = makeSelectSQL(TB2)
    Call makeListData(strSql, TB2)
    If cntRecord > 0 Then
        Me.txtLandLine = LISTDATA(0, 4)
        Me.txtWMCPhone = LISTDATA(0, 5)
        Me.txtPhone_PStaff = LISTDATA(0, 6)
        Me.txtPhone_Spouse = LISTDATA(0, 7)
        Me.lblChurch.Caption = LISTDATA(0, 3)
        Me.lblTime_different = Format(LISTDATA(0, 1), "시차: hh시nn분")
        Me.txtName = LISTDATA(0, 9)
        Me.txtPosition = LISTDATA(0, 10)
        Me.txtName_Spouse = LISTDATA(0, 12)
        Me.txtPosition_Spouse = LISTDATA(0, 13)
        Me.txtAddress = LISTDATA(0, 18)
        
        '--//교회명 및 시차 띄우기
        Me.lblChurch.Visible = True
        Me.lblTime_different.Visible = True
        
        '--//교회에 따른 텍스트박스 비활성화
        If Me.lstPStaff.listIndex <> -1 Then
            If LISTDATA(0, 2) = "" Then '--//교회코드가 없으면
                Me.txtLandLine.Enabled = False
                Me.txtWMCPhone.Enabled = False
                Me.txtAddress.Enabled = False
            Else
                Me.txtLandLine.Enabled = True
                Me.txtWMCPhone.Enabled = True
                Me.txtAddress.Enabled = True
            End If
        Else
            Me.txtLandLine.Enabled = False
            Me.txtWMCPhone.Enabled = False
            Me.txtAddress.Enabled = False
        End If
    Else
        Call sbtxtBox_Init
    End If
    Call sbClearVariant
    
    '--//성별에 따른 텍스트박스 비활성화
    If Me.lstPStaff.listIndex <> -1 Then
        lifeNo = Me.lstPStaff.List(Me.lstPStaff.listIndex)
        
        If Mid(lifeNo, InStr(5, lifeNo, "-") + 1, 1) = 1 Then
            Me.txtPhone_PStaff.Enabled = True
            Me.txtPhone_Spouse.Enabled = False
        End If
        If Mid(lifeNo, InStr(5, lifeNo, "-") + 1, 1) = 2 Then
            Me.txtPhone_PStaff.Enabled = False
            Me.txtPhone_Spouse.Enabled = True
        End If
    Else
        Me.txtPhone_PStaff.Enabled = False
        Me.txtPhone_Spouse.Enabled = False
    End If
    
    '--//사진추가
    filePath = fnFindPicPath
    FileName = Me.lstPStaff.List(Me.lstPStaff.listIndex) & ".jpg"
'    If Not Len(Dir(FilePath & FileName)) > 0 Then
'        FileName = Me.lstPStaff.List(Me.lstPStaff.ListIndex) & ".png"
'    End If
On Error Resume Next
    Me.lblPic.Picture = LoadPicture(filePath & FileName)
    If err.Number <> 0 Then
        Me.lblPic.Picture = LoadPicture("")
    End If
On Error GoTo 0
    
End Sub

Private Sub lstPStaff_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Un'HookListBoxScroll
End Sub

Private Sub lstPStaff_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Me.lstPStaff.ListCount > 0 Then
        'HookListBoxScroll Me, Me.lstPStaff
    End If
End Sub

Private Sub txtChurchNM_Change()
    Me.txtChurchNM.BackColor = RGB(255, 255, 255)
End Sub

Private Sub UserForm_Initialize()
    If connIP = Empty Then setGlobalVariant '//전역변수 재설정
    If checkLogin = 0 Then f_login.Show '//로그인체크
    If checkLogin = 0 Then End '--//로그인 실패 시 프로시저 종료
    
    '--//기초설정
    Me.cmdClose.Cancel = True
    TB1 = "op_system.v0_pstaff_information_all" '--//선지자리스트
    TB2 = "op_system.v_phone" '--//전화번호
    TB3 = "op_system.db_phone" '--//교회연락처
    TB4 = "op_system.db_pastoralstaff" '--//선지자정보
    TB5 = "op_system.db_pastoralwife" '--//배우자정보
    TB6 = "op_system.v_phone_export" '--//출력물 생성양식
    
    '--//컨트롤 설정
    Me.lstPStaff.Enabled = False
    Me.lblChurch.Visible = False
    Me.lblTime_different.Visible = False
    Me.txtName.Visible = False
    Me.txtName_Spouse.Visible = False
    Me.txtPosition.Visible = False
    Me.txtPosition_Spouse.Visible = False
    Me.txtPhone_PStaff.Enabled = False
    Me.txtPhone_Spouse.Enabled = False
    Me.txtLandLine.Enabled = False
    Me.txtWMCPhone.Enabled = False
    Me.txtAddress.Enabled = False
    Me.lblExport.Visible = False
    Me.optKo.Value = True
    
    '--//리스트박스 설정
    With Me.lstPStaff
        .ColumnCount = 6
        .ColumnHeads = False
        .ColumnWidths = "0,0,120,70,50,50" '생명번호, 교회코드, 교회명, 관리교회명, 한글이름(직분), 직책
        .TextAlign = fmTextAlignLeft
        .Font = "굴림"
    End With
    
    '--//부서에 따른 국제전화 카드번호 설정
    strSql = "SELECT a.dept_phonecard FROM op_system.db_ovs_dept a WHERE a.dept_id = " & SText(USER_DEPT)
    Call makeListData(strSql, "op_system.db_ovs_dept")
    Me.lblCardNo.Caption = "일반전화(국제 08216) : " & LISTDATA(0, 0)
    Call sbClearVariant
    
    Me.txtChurchNM.SetFocus
    
End Sub
Private Sub cmdSearch_Click()

    Call sbtxtBox_Init
    
    If Not Me.chkAll Then
        strSql = makeSelectSQL(TB1)
        Call makeListData(strSql, TB1)
    Else
        strSql = makeSelectSQL2(TB1)
        Call makeListData(strSql, TB1)
    End If
    
    If cntRecord = 0 Then
            MsgBox "반환할 DB 데이터가 없습니다.", vbInformation, banner
            Call sbClearVariant
            Exit Sub
        End If
    
    Me.lstPStaff.List = LISTDATA
    Call sbClearVariant
    Me.lstPStaff.Enabled = True
End Sub
Private Sub makeListData(ByVal strSql As String, ByVal tableNM As String)

    Dim i As Integer, j As Integer
    
    '//DB에서 자료 호출하여 레코드셋에 반환
    connectTaskDB
    callDBtoRS "makeListData", tableNM, strSql, Me.Name
    
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
'-----------------------------------------
'  조건별 Select SQL문 작성
'    - makeSelectSQL(검색어, 필터)
'    - DB에서 반환할 리스트 필드수정
'-----------------------------------------
Private Function makeSelectSQL(ByVal tableNM As String) As String
    
    Select Case tableNM
    Case TB1
        strSql = "SELECT a.`선지자생명번호`,a.`교회코드`,a.`교회명`,a.`관리교회명`,a.`한글이름(직분)`,a.`직책` " & _
                    "FROM " & TB2 & " a " & _
                    "WHERE a.`직책` is not null AND (a.`한글이름(직분)` LIKE '%" & Me.txtChurchNM & "%' OR a.`영문이름` LIKE '%" & Me.txtChurchNM & "%' OR a.`교회명` LIKE '%" & Me.txtChurchNM & "%'" & " OR a.`관리교회명` LIKE '%" & Me.txtChurchNM & "%'" & " OR a.`선지자생명번호` LIKE '%" & Me.txtChurchNM & "%')" & _
                    " AND a.`관리부서` = " & SText(USER_DEPT) & _
                 " Union " & _
                 "SELECT b.`배우자생명번호`,b.`교회코드`,b.`교회명`,b.`관리교회명`,b.`사모한글이름(직분)`,b.`사모직책` " & _
                    "FROM " & TB2 & " b " & _
                    "WHERE b.`사모직책` is not null AND (b.`사모한글이름(직분)` LIKE '%" & Me.txtChurchNM & "%' OR b.`사모영문이름` LIKE '%" & Me.txtChurchNM & "%' OR b.`교회명` LIKE '%" & Me.txtChurchNM & "%'" & " OR b.`관리교회명` LIKE '%" & Me.txtChurchNM & "%'" & " OR b.`배우자생명번호` LIKE '%" & Me.txtChurchNM & "%')" & _
                    " AND b.`관리부서` = " & SText(USER_DEPT) & _
                    " ORDER BY `관리교회명`,FIELD(`직책`,'당회장','당회장대리','당사모','당대리사모','동역','동사모','지교회관리자','지관자사모','예배소관리자','예관자사모','예비생도1단계','예비생도2단계','예비생도3단계','생도사모'," & getPosition2Joining & ",'');"
    Case TB2
        With Me.lstPStaff
            strSql = "SELECT * FROM " & TB2 & " a WHERE a.`선지자생명번호` = " & SText(.List(.listIndex)) & " OR a.`배우자생명번호` = " & SText(.List(.listIndex)) & ";"
        End With
    Case TB6
        If Me.optKo.Value Then
            strSql = "SELECT" & _
                        " a.`선교국가`" & _
                        ",DATE_FORMAT(a.`시차`,'%H:%i') '시차'" & _
                        ",a.`교회명`" & _
                        ",a.`인터넷전화`" & _
                        ",a.`유선전화`" & _
                        ",a.`직책`" & _
                        ",a.`한글이름(직분)`" & _
                        ",a.`선지자전화번호`" & _
                        ",a.`사모한글이름(직분)`" & _
                        ",a.`배우자전화번호`" & _
                        ",a.`관리부서`" & _
                        ",a.`본교회코드`" & _
                        ",a.`지교회코드`" & _
                    " FROM " & TB2 & " a" & _
                    " WHERE a.`교회명` IS NOT NULL AND a.`직책` IS NOT NULL" & _
                        " AND a.`정렬순서` >= (SELECT sort_order FROM op_system.db_churchlist WHERE church_nm = '몽골 울란바토르')" & _
                        " AND a.`관리부서` = " & SText(USER_DEPT) & _
                    " ORDER BY a.`정렬순서`, a.`직책` IS NULL ASC, FIELD(`직책`,'당회장','당회장대리','당사모','당대리사모','동역','동사모','지교회관리자','지관자사모','예배소관리자','예관자사모','예비생도1단계','예비생도2단계','예비생도3단계','생도사모'," & getPosition2Joining & ",'')" & ";"
        Else
            strSql = "SELECT" & _
                        " a.`선교국가`" & _
                        ",DATE_FORMAT(a.`시차`,'%H:%i') '시차'" & _
                        ",a.`영문교회명`" & _
                        ",a.`인터넷전화`" & _
                        ",a.`유선전화`" & _
                        ",a.`직책`" & _
                        ",a.`영문이름`" & _
                        ",a.`선지자전화번호`" & _
                        ",a.`사모영문이름`" & _
                        ",a.`배우자전화번호`" & _
                        ",a.`관리부서`" & _
                        ",a.`본교회코드`" & _
                        ",a.`지교회코드`" & _
                    " FROM " & TB2 & " a" & _
                    " WHERE a.`교회명` IS NOT NULL AND a.`직책` IS NOT NULL" & _
                        " AND a.`정렬순서` >= (SELECT sort_order FROM op_system.db_churchlist WHERE church_nm = '몽골 울란바토르')" & _
                        " AND a.`관리부서` = " & SText(USER_DEPT) & _
                    " ORDER BY a.`정렬순서`, a.`직책` IS NULL ASC, FIELD(`직책`,'당회장','당회장대리','당사모','당대리사모','동역','동사모','지교회관리자','지관자사모','예배소관리자','예관자사모','예비생도1단계','예비생도2단계','예비생도3단계','생도사모'," & getPosition2Joining & ",'')" & ";"
        End If
    Case Else
    End Select
    makeSelectSQL = strSql
End Function
Private Function makeSelectSQL2(ByVal tableNM As String) As String
    
    Select Case tableNM
    Case TB1
        strSql = "SELECT a.`생명번호`,esta1.`church_sid_custom`,a.`지교회명`,a.`교회명`,a.`한글이름(직분)`,a.`직책` " & _
                    "FROM " & TB1 & " a " & _
                    "LEFT JOIN op_system.db_history_church_establish esta1 ON IFNULL(a.`지교회코드`,a.`교회코드`)=esta1.`church_sid` " & _
                    "WHERE (a.`한글이름(직분)` LIKE '%" & Me.txtChurchNM & "%' OR a.`영문이름` LIKE '%" & Me.txtChurchNM & "%' OR a.`교회명` LIKE '%" & Me.txtChurchNM & "%'" & " OR a.`지교회명` LIKE '%" & Me.txtChurchNM & "%'" & " OR a.`생명번호` LIKE '%" & Me.txtChurchNM & "%')" & _
                    " AND a.`관리부서` = " & SText(USER_DEPT) & _
                 " Union " & _
                 "SELECT b.`배우자생번`,esta2.`church_sid_custom`,b.`지교회명`,b.`교회명`,b.`사모한글이름(직분)`,b.`사모직책` " & _
                    "FROM " & TB1 & " b " & _
                    "LEFT JOIN op_system.db_history_church_establish esta2 ON IFNULL(b.`지교회코드`,b.`교회코드`)=esta2.`church_sid` " & _
                    "WHERE (b.`사모한글이름(직분)` LIKE '%" & Me.txtChurchNM & "%' OR b.`사모영문이름` LIKE '%" & Me.txtChurchNM & "%' OR b.`교회명` LIKE '%" & Me.txtChurchNM & "%'" & " OR b.`지교회명` LIKE '%" & Me.txtChurchNM & "%'" & " OR b.`배우자생번` LIKE '%" & Me.txtChurchNM & "%')" & _
                    " AND b.`관리부서` = " & SText(USER_DEPT) & " AND b.`배우자생번` IS NOT NULL " & _
                    " ORDER BY `교회명`,FIELD(`직책`,'당회장','당회장대리','당사모','당대리사모','동역','동사모','지교회관리자','지관자사모','예배소관리자','예관자사모','예비생도1단계','예비생도2단계','예비생도3단계','생도사모'," & getPosition2Joining & ",'');"
    Case TB2
    Case Else
    End Select
    makeSelectSQL2 = strSql
End Function
Private Function makeUpdateSQL(ByVal tableNM As String) As String
    
    Select Case tableNM
    Case TB3
        strSql = "SELECT a.church_sid FROM " & TB3 & " a " & _
                " WHERE a.church_sid = " & SText(Me.lstPStaff.List(Me.lstPStaff.listIndex, 1)) & ";"
        Call makeListData(strSql, TB3)
        
        If cntRecord > 0 Then
            strSql = "UPDATE " & TB3 & " a " & _
                    "SET a.phone = " & SText(Me.txtLandLine) & ", a.wmcphone = " & SText(Me.txtWMCPhone) & ", a.address = " & SText(Me.txtAddress) & _
                    " WHERE a.church_sid = " & SText(Me.lstPStaff.List(Me.lstPStaff.listIndex, 1)) & ";"
        Else
            strSql = "INSERT INTO " & TB3 & _
                    " VALUES (" & SText(Me.lstPStaff.List(Me.lstPStaff.listIndex, 1)) & "," & SText(Me.txtLandLine) & "," & SText(Me.txtWMCPhone) & "," & SText(Me.txtAddress) & ");"
        End If
    Case TB4
        strSql = "UPDATE " & TB4 & " a " & _
                "SET a.phone = " & SText(Me.txtPhone_PStaff) & _
                " WHERE a.lifeno = " & SText(Me.lstPStaff.List(Me.lstPStaff.listIndex)) & ";"
    Case TB5
        strSql = "UPDATE " & TB5 & " a " & _
                "SET a.phone = " & SText(Me.txtPhone_Spouse) & _
                " WHERE a.lifeno = " & SText(Me.lstPStaff.List(Me.lstPStaff.listIndex)) & ";"
    Case Else
    End Select
    makeUpdateSQL = strSql
End Function
Private Sub sbClearVariant()
    Erase LISTFIELD
    Erase LISTDATA
    cntRecord = Empty
    strSql = vbNullString
End Sub
Sub sbtxtBox_Init()
    Me.txtLandLine = ""
    Me.txtPhone_PStaff = ""
    Me.txtPhone_Spouse = ""
    Me.txtWMCPhone = ""
    Me.txtName = ""
    Me.txtName_Spouse = ""
    Me.txtPosition = ""
    Me.txtPosition_Spouse = ""
    Me.lblChurch.Visible = False
    Me.lblTime_different.Visible = False
End Sub
Private Function CopyText(text As String) As Boolean
    
On Error GoTo nErr
    Dim MSForms_DataObject As DataObject
'    Set MSForms_DataObject = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    If MSForms_DataObject Is Nothing Then Set MSForms_DataObject = New DataObject
    MSForms_DataObject.SetText text
    MSForms_DataObject.PutInClipboard
'    Set MSForms_DataObject = Nothing
    CopyText = True
nErr:
End Function

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
