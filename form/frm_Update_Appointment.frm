VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Update_Appointment 
   Caption         =   "소속교회, 직분, 직책 컨트롤러"
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7305
   OleObjectBlob   =   "frm_Update_Appointment.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "frm_Update_Appointment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim txtBox_Focus As MSForms.textBox

'--//직분 콤보박스 변경 이벤트
Private Sub cboTitle_Change()
    Me.txtChurchNow = Me.cboTitle
End Sub

'--//직분 콤보박스를 위한 마우스 스크롤
Private Sub cboTitle_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Un'HookListBoxScroll
End Sub

'--//직분 콤보박스를 위한 마우스 스크롤
Private Sub cboTitle_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Me.cboTitle.ListCount > 0 Then
        'HookListBoxScroll Me, Me.cboTitle
    End If
End Sub

'--//무기한 체크박스 변경 이벤트
Private Sub chkPresent_Change()
    Select Case Me.chkPresent.Value
        Case True
            TextBoxEnable Me.txtEnd, False
            Me.txtEnd.Value = "현재"
        Case False
            TextBoxEnable Me.txtEnd, True
            If Me.lstHistory.listIndex = -1 Then
                Me.txtEnd = ""
            Else
                If Me.txtEnd = "현재" Then
                    Me.txtEnd.Value = Date - 1
                End If
            End If
    Case Else
    End Select
End Sub

'--//cmdCancel 버튼 클릭 이벤트
Private Sub cmdCancel_Click()
    InputModeActivate False
    HideDeleteButtonByUserAuth
    lstHistory_Click
End Sub

'--//cmdClose 버튼 클릭 이벤트
Private Sub cmdClose_Click()
    Unload Me
End Sub

'--//cmdDelete 버튼 클릭 이벤트
Private Sub cmdDelete_Click()
    
    Dim result As T_RESULT
    
    If MsgBox("선택한 데이터를 삭제하시겠습니까?", vbQuestion + vbYesNo, banner) = vbNo Then
        Exit Sub
    End If
    
    Select Case TASK_CODE
        Case 1
            Dim tmpTransfer As New Transfer
            Dim objTransferDao As New TransferDao
            With Me.lstHistory
                Set tmpTransfer = objTransferDao.FindByCode(.List(.listIndex))
            End With
            objTransferDao.Delete tmpTransfer
        Case 2
            Dim tmpTitle As New title
            Dim objTitleDao As New TitleDao
            With Me.lstHistory
                Set tmpTitle = objTitleDao.FindByCode(.List(.listIndex))
            End With
            objTitleDao.Delete tmpTitle
        Case 3
            Dim tmpPosition As New position
            Dim objPositionDao As New PositionDao
            With Me.lstHistory
                Set tmpPosition = objPositionDao.FindByCode(.List(.listIndex))
            End With
            objPositionDao.Delete tmpPosition
        Case 4
            Dim tmpSpecialPosition As New SpecialPosition
            Dim objSpecialPositionDao As New SpecialPositionDao
            With Me.lstHistory
                Set tmpSpecialPosition = objSpecialPositionDao.FindByCode(.List(.listIndex))
            End With
            objSpecialPositionDao.Delete tmpSpecialPosition
    Case Else
    End Select
    
    '--//메세지박스
    MsgBox "해당 데이터가 삭제되었습니다.", , banner
    
    '--//리스트박스 새로고침
    Call lstPStaff_Click
    Me.lstHistory.listIndex = -1
    
End Sub

'--//cmdEdit 버튼 클릭 이벤트
Private Sub cmdEdit_Click()
    
    '--//데이터 유효성 검사
    If fnData_Validation = False Then
On Error Resume Next
        txtBox_Focus.SetFocus
        txtBox_Focus.SelStart = 0
        txtBox_Focus.SelLength = Len(txtBox_Focus)
On Error GoTo 0
        Exit Sub
    End If
    
    Dim objTransfer As New Transfer
    Dim objTitle As New title
    Dim objPosition As New position
    Dim objSpecialPosition As New SpecialPosition
    Dim objTransferDao As New TransferDao
    Dim objTitleDao As New TitleDao
    Dim objPositionDao As New PositionDao
    Dim objSpecialPositionDao As New SpecialPositionDao
    
    '--//수정된 내용 있는지 체크
On Error GoTo ERR_TIME_OVERLAPPED
    Select Case TASK_CODE
    Case 1
        Dim tmpTransfer As New Transfer
        objTransfer.ParseFromForm Me
        Set tmpTransfer = objTransferDao.FindByTrans(objTransfer)
        
        If objTransfer.IsEqual(tmpTransfer, True) Then
            Exit Sub
        End If
        
        If objTransferDao.CheckTimeOverlapped(objTransfer) Then
            err.Raise ERR_CODE_TIME_OVERLAPPED, , ERR_DESC_TIME_OVERLAPPED
            returnListPosition Me, Me.lstHistory.Name, objTransfer.Code
        End If
        
        objTransferDao.Save objTransfer
        
    Case 2
        Dim tmpTitle As New title
        objTitle.ParseFromForm Me
        Set tmpTitle = objTitleDao.FindByTitle(objTitle)
        
        If objTitle.IsEqual(tmpTitle, True) Then
            Exit Sub
        End If
        
        If objTitleDao.CheckTimeOverlapped(objTitle) Then
            err.Raise ERR_CODE_TIME_OVERLAPPED, , ERR_DESC_TIME_OVERLAPPED
            returnListPosition Me, Me.lstHistory.Name, objTitle.Code
        End If
        
        objTitleDao.Save objTitle
        
    Case 3
        Dim tmpPosition As New position
        objPosition.ParseFromForm Me
        Set tmpPosition = objPositionDao.FindByPosition(objPosition)
        
        If objPosition.IsEqual(tmpPosition, True) Then
            Exit Sub
        End If
        
        If objPositionDao.CheckTimeOverlapped(objPosition) Then
            err.Raise ERR_CODE_TIME_OVERLAPPED, , ERR_DESC_TIME_OVERLAPPED
            returnListPosition Me, Me.lstHistory.Name, objPosition.Code
        End If
        
        objPositionDao.Save objPosition
        
    Case 4
        Dim tmpSpecialPosition As New SpecialPosition
        objSpecialPosition.ParseFromForm Me
        Set tmpSpecialPosition = objSpecialPositionDao.FindBySpecialPosition(objSpecialPosition)
        
        If objSpecialPosition.IsEqual(tmpSpecialPosition, True) Then
            Exit Sub
        End If
        
        If objSpecialPositionDao.CheckTimeOverlapped(objSpecialPosition) Then
            err.Raise ERR_CODE_TIME_OVERLAPPED, , ERR_DESC_TIME_OVERLAPPED
            returnListPosition Me, Me.lstHistory.Name, objSpecialPosition.Code
        End If
        
        objSpecialPositionDao.Save objSpecialPosition
    
    End Select
    
    '--//메세지박스
    MsgBox "업데이트 되었습니다.", , banner
    
    '--//리스트박스 초기화
    Dim queryKey As Integer
    queryKey = Me.lstHistory.listIndex
    lstPStaff_Click
    Me.lstHistory.listIndex = queryKey

DONE:
    Exit Sub
ERR_TIME_OVERLAPPED:
    MsgBox err.Description, vbCritical, banner
End Sub

'--//cmdNew버튼 클릭 이벤트
Private Sub cmdNew_Click()
    
    If lstHistory.ListCount = 0 Then
        lstHistory_Click '--//입력 컨트롤 활성화를 위한 lstHistory 클릭이벤트
    End If
    
    Me.lstHistory.listIndex = Me.lstHistory.ListCount - 1
    
    InputModeActivate True
    HideDeleteButtonByUserAuth
    
End Sub

'--//cmdAdd 버튼 클릭 이벤트
Private Sub cmdADD_Click()
    
    '--//데이터 유효성 검사
    If fnData_Validation = False Then
On Error Resume Next
        txtBox_Focus.SetFocus
        txtBox_Focus.SelStart = 0
        txtBox_Focus.Setlength = Len(txtBox_Focus)
On Error GoTo 0
        Exit Sub
    End If
    
    Dim objTransfer As New Transfer
    Dim objTitle As New title
    Dim objPosition As New position
    Dim objSpecialPosition As New SpecialPosition
    Dim objTransferDao As New TransferDao
    Dim objTitleDao As New TitleDao
    Dim objPositionDao As New PositionDao
    Dim objSpecialPositionDao As New SpecialPositionDao
    
    '--//중복체크
On Error GoTo ERR_TIME_OVERLAPPED
    Select Case TASK_CODE
        Case 1
            objTransfer.ParseFromForm Me
            
            If objTransferDao.CheckTimeOverlapped(objTransfer) Then
                err.Raise ERR_CODE_TIME_OVERLAPPED, , ERR_DESC_TIME_OVERLAPPED
                returnListPosition Me, Me.lstHistory.Name, objTransfer.Code
            End If
            
            objTransferDao.Save objTransfer
        Case 2
            objTitle.ParseFromForm Me
            
            If objTitleDao.CheckTimeOverlapped(objTitle) Then
                err.Raise ERR_CODE_TIME_OVERLAPPED, , ERR_DESC_TIME_OVERLAPPED
                returnListPosition Me, Me.lstHistory.Name, objTitle.Code
            End If
            
            objTitleDao.Save objTitle
        Case 3
            objPosition.ParseFromForm Me
            
            If objPositionDao.CheckTimeOverlapped(objPosition) Then
                err.Raise ERR_CODE_TIME_OVERLAPPED, , ERR_DESC_TIME_OVERLAPPED
                returnListPosition Me, Me.lstHistory.Name, objPosition.Code
            End If
            
            objPositionDao.Save objPosition
        Case 4
            objSpecialPosition.ParseFromForm Me
            
            If objSpecialPositionDao.CheckTimeOverlapped(objSpecialPosition) Then
                err.Raise ERR_CODE_TIME_OVERLAPPED, , ERR_DESC_TIME_OVERLAPPED
                returnListPosition Me, Me.lstHistory.Name, objSpecialPosition.Code
            End If
            
            objSpecialPositionDao.Save objSpecialPosition
    End Select
    
    '--//버튼설정 원래대로
    InputModeActivate False
    HideDeleteButtonByUserAuth
    
    '--//메세지박스
    MsgBox "추가 되었습니다.", , banner
    lstPStaff_Click
    Me.lstHistory.listIndex = Me.lstHistory.ListCount - 1
    lstHistory_Click
    
DONE:
    Exit Sub
ERR_TIME_OVERLAPPED:
    MsgBox err.Description, vbCritical, banner
End Sub

'--//cmdSearch 버튼 클릭 이벤트
Private Sub cmdSearch_Church_Click()
    argShow = 1
    argShow3 = 1
    frm_Update_Appointment_1.Show
End Sub

'--//lstHistory 클릭 이벤트
Private Sub lstHistory_Click()
    
    Dim IsEmpty As Boolean
    
    '--//컨트롤 설정
    If Me.lstHistory.listIndex <> -1 Then
        IsEmpty = True
    Else
        IsEmpty = False
    End If
    
    Me.txtStart.Enabled = IsEmpty
    Me.txtEnd.Enabled = IsEmpty
    Me.cmdEdit.Enabled = IsEmpty
    Me.cmdAdd.Enabled = IsEmpty
    Me.cmdDelete.Enabled = IsEmpty
    Me.chkPresent.Visible = IsEmpty
    If TASK_CODE = 1 Then
        Me.cmdSearch_Church.Enabled = IsEmpty
    Else
        Me.txtChurchNow.Enabled = IsEmpty
    End If
    Me.cboTitle.Enabled = IsEmpty
    
    
    '--//리스트 클릭 시 시작일, 종료일, 내용 표시
    If Me.lstHistory.listIndex <> -1 Then
        With Me.lstHistory
            Select Case TASK_CODE
            Case 1
                Me.txtStart = .List(.listIndex, 2)
                Me.txtEnd = .List(.listIndex, 3)
                Me.txtChurchNow_sid = .List(.listIndex, 4)
                Me.txtChurchNow = .List(.listIndex, 5)
            Case 2
                Me.txtStart = .List(.listIndex, 2)
                Me.txtEnd = .List(.listIndex, 3)
                If CDate(.List(.listIndex, 5)) <> DateSerial(1900, 1, 1) Then
                    Me.txtTitleOrdinaryDate = .List(.listIndex, 5)
                Else
                    Me.txtTitleOrdinaryDate = ""
                End If
                Me.txtChurchNow = .List(.listIndex, 4)
                Me.cboTitle = .List(.listIndex, 4)
            Case 3, 4
                Me.txtStart = .List(.listIndex, 2)
                Me.txtEnd = .List(.listIndex, 3)
                Me.txtChurchNow = .List(.listIndex, 4)
                Me.cboTitle = .List(.listIndex, 4)
            End Select
        End With
    End If
    
    If Me.txtEnd = "현재" Then
        Me.chkPresent.Value = True
        Me.txtEnd.Enabled = False
    Else
        Me.chkPresent.Value = False
    End If
    
End Sub

'--//lstHistory를 위한 마우스 스크롤
Private Sub lstHistory_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Un'HookListBoxScroll
End Sub

'--//lstHistory를 위한 마우스 스크롤
Private Sub lstHistory_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Me.lstHistory.ListCount > 0 Then
        'HookListBoxScroll Me, Me.lstHistory
    End If
End Sub

'--//lstPStaff 클릭 이벤트
Private Sub lstPStaff_Click()
    
    '--//컨트롤 초기화
    TextBox_Initialize
    Me.lstHistory.Enabled = True
    Me.cmdNew.Enabled = True
    Me.lstHistory.Clear
    
    '--//이력 목록상자 조정
    With Me.lstHistory
        Select Case TASK_CODE
        Case 1
            .ColumnCount = 6
            .ColumnHeads = False
            .ColumnWidths = "0,0,80,100,0,250" '발령코드, 생명번호, 시작일, 종료일, 교회코드, 교회명
            .TextAlign = fmTextAlignLeft
            .Font = "굴림"
        Case 2
            .ColumnCount = 6
            .ColumnHeads = False
            .ColumnWidths = "0,0,80,100,250, 100" '직분(직책)코드, 생명번호, 시작일, 종료일, 직분(또는 직책), 안수일
            .TextAlign = fmTextAlignLeft
            .Font = "굴림"
        Case 3, 4
            .ColumnCount = 5
            .ColumnHeads = False
            .ColumnWidths = "0,0,80,100,250" '직분(직책)코드, 생명번호, 시작일, 종료일, 직분(또는 직책)
            .TextAlign = fmTextAlignLeft
            .Font = "굴림"
        End Select
    End With
    
    Dim tmpList As Object
    Dim strLifeNo As String
    With Me.lstPStaff
        strLifeNo = .List(.listIndex)
    End With
    
    '--//이력목록 데이터 채우기
    Dim i As Integer
    Select Case TASK_CODE
        Case 1
            Dim objTransfer As New Transfer
            Dim objTransferDao As New TransferDao
            
            Set tmpList = objTransferDao.FindByLifeNo(strLifeNo)
            
            If Not tmpList Is Nothing Then
                Dim tmpTransfer As New Transfer
                With Me.lstHistory
                    For Each tmpTransfer In tmpList
                        .AddItem tmpTransfer.Code
                        .List(.ListCount - 1, 1) = tmpTransfer.lifeNo
                        .List(.ListCount - 1, 2) = tmpTransfer.startDate
                        If tmpTransfer.endDate = DateSerial(9999, 12, 31) Then
                            .List(.ListCount - 1, 3) = "현재"
                        Else
                            .List(.ListCount - 1, 3) = tmpTransfer.endDate
                        End If
                        .List(.ListCount - 1, 4) = tmpTransfer.ChurchID
                        .List(.ListCount - 1, 5) = tmpTransfer.churchName
                    Next
                End With
            End If
        Case 2
            Dim objTitle As New title
            Dim objTitleDao As New TitleDao
            
            Set tmpList = objTitleDao.FindByLifeNo(strLifeNo)
            
            If Not tmpList Is Nothing Then
                Dim tmpTitle As New title
                With Me.lstHistory
                    For Each tmpTitle In tmpList
                        .AddItem tmpTitle.Code
                        .List(.ListCount - 1, 1) = tmpTitle.lifeNo
                        .List(.ListCount - 1, 2) = tmpTitle.startDate
                        If tmpTitle.endDate = DateSerial(9999, 12, 31) Then
                            .List(.ListCount - 1, 3) = "현재"
                        Else
                            .List(.ListCount - 1, 3) = tmpTitle.endDate
                        End If
                        .List(.ListCount - 1, 4) = tmpTitle.title
                        .List(.ListCount - 1, 5) = tmpTitle.TitleOrdinaryDate
                    Next
                End With
            End If
        Case 3
            Dim objPosition As New position
            Dim objPositionDao As New PositionDao
            
            Set tmpList = objPositionDao.FindByLifeNo(strLifeNo)
            
            If Not tmpList Is Nothing Then
                Dim tmpPosition As New position
                With Me.lstHistory
                    For Each tmpPosition In tmpList
                        .AddItem tmpPosition.Code
                        .List(.ListCount - 1, 1) = tmpPosition.lifeNo
                        .List(.ListCount - 1, 2) = tmpPosition.startDate
                        If tmpPosition.endDate = DateSerial(9999, 12, 31) Then
                            .List(.ListCount - 1, 3) = "현재"
                        Else
                            .List(.ListCount - 1, 3) = tmpPosition.endDate
                        End If
                        .List(.ListCount - 1, 4) = tmpPosition.position
                    Next
                End With
            End If
        Case 4
            Dim objSpecialPosition As New SpecialPosition
            Dim objSpecialPositionDao As New SpecialPositionDao
            
            Set tmpList = objSpecialPositionDao.FindByLifeNo(strLifeNo)
            
            If Not tmpList Is Nothing Then
                Dim tmpSpecialPosition As New SpecialPosition
                With Me.lstHistory
                    For Each tmpSpecialPosition In tmpList
                        .AddItem tmpSpecialPosition.Code
                        .List(.ListCount - 1, 1) = tmpSpecialPosition.lifeNo
                        .List(.ListCount - 1, 2) = tmpSpecialPosition.startDate
                        If tmpSpecialPosition.endDate = DateSerial(9999, 12, 31) Then
                            .List(.ListCount - 1, 3) = "현재"
                        Else
                            .List(.ListCount - 1, 3) = tmpSpecialPosition.endDate
                        End If
                        .List(.ListCount - 1, 4) = tmpSpecialPosition.SpecialPosition
                    Next
                End With
            End If
    End Select
    
    '--//이력 리스트박스가 비어있지 않으면 마지막 값을 클릭
    If Me.lstHistory.ListCount > 0 Then
        Me.lstHistory.listIndex = Me.lstHistory.ListCount - 1
    End If
    
    '--//컨트롤 설정
    If Me.lstHistory.listIndex <> -1 Then
        Me.txtStart.Enabled = True
        Me.txtEnd.Enabled = True
        Me.txtTitleOrdinaryDate.Enabled = True
        Me.cmdEdit.Enabled = True
        Me.cmdAdd.Enabled = True
        Me.cmdDelete.Enabled = True
        Me.chkPresent.Visible = True
        If TASK_CODE = 1 Then
            Me.cmdSearch_Church.Enabled = True
        Else
            Me.txtChurchNow.Enabled = True
        End If
        Me.cboTitle.Enabled = True
    Else
        Me.txtStart.Enabled = False
        Me.txtEnd.Enabled = False
        Me.txtTitleOrdinaryDate.Enabled = False
        Me.cmdEdit.Enabled = False
        Me.cmdAdd.Enabled = False
        Me.cmdDelete.Enabled = False
        Me.chkPresent.Visible = False
        If TASK_CODE = 1 Then
            Me.cmdSearch_Church.Enabled = False
        Else
            Me.txtChurchNow.Enabled = False
        End If
        Me.cboTitle.Enabled = False
    End If
    
    If Me.txtEnd = "현재" Then
        Me.chkPresent.Value = -1
        Me.chkPresent.Value = True
    End If
    
    '--//사진추가
    InsertPicToLabel Me.lblPic, strLifeNo
    
End Sub

'--//lstPStaff를 위한 마우스 스크롤
Private Sub lstPStaff_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Un'HookListBoxScroll
End Sub

'--//lstPStaff를 위한 마우스 스크롤
Private Sub lstPStaff_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Me.lstPStaff.ListCount > 0 Then
        'HookListBoxScroll Me, Me.lstPStaff
    End If
End Sub

'--//옵션 공통 클릭 이벤트
Private Sub RadioOpstionClickEvent(strValue As String)

    Me.lstHistory.Clear
    If Me.lstPStaff.listIndex <> -1 Then
        Call lstPStaff_Click
    Else
        Me.lstPStaff.Clear
    End If
    Me.lblKind.Caption = strValue
    Me.lblKind2.Caption = strValue
    Me.txtChurchNM.Enabled = True
    Me.cmdSearch.Enabled = True
    Me.txtChurchNM.BackColor = &HC0FFFF
    
    '--//이력 리스트박스가 비어있지 않으면 마지막 값을 클릭
    If Me.lstHistory.ListCount > 0 Then
        Me.lstHistory.listIndex = Me.lstHistory.ListCount - 1
    End If

End Sub

'--//직책이력옵션 변경 이벤트
Private Sub optPosition_Change()
    If Me.optPosition Then
        Me.cboTitle.Visible = True
        Me.txtChurchNow.Enabled = False
        
        Dim objPositionDao As New PositionDao
        Dim positionList As Object
        
        Set positionList = objPositionDao.GetPositionList
        
        Dim strPosition As Variant
        For Each strPosition In positionList
            Me.cboTitle.AddItem strPosition
        Next
        
        Me.cmdSearch_Church.Visible = False
    Else
        Me.cboTitle.Visible = False
        Me.txtChurchNow.Enabled = True
        Me.cboTitle.Clear
        Me.cmdSearch_Church.Visible = True
    End If
End Sub

'--//직책이력옵션 클릭 이벤트
Private Sub optPosition_Click()
    TASK_CODE = 3
    RadioOpstionClickEvent "직책"
End Sub

'--//특수직책이력옵션 변경 이벤트
Private Sub optPosition2_Change()
    If Me.optPosition2 Then
        Me.cboTitle.Visible = True
        Me.txtChurchNow.Enabled = False
        
        Dim objSpecialPositionDao As New SpecialPositionDao
        Dim specialPositionList As Object
        
        Set specialPositionList = objSpecialPositionDao.GetSpecialPositionList
        
        Dim strSpecialPosition As Variant
        For Each strSpecialPosition In specialPositionList
            Me.cboTitle.AddItem strSpecialPosition
        Next
        
        Me.cmdSearch_Church.Visible = False
    Else
        Me.cboTitle.Visible = False
        Me.txtChurchNow.Enabled = True
        Me.cboTitle.Clear
        Me.cmdSearch_Church.Visible = True
    End If
End Sub

'--//특수직책이력옵션 클릭 이벤트
Private Sub optPosition2_Click()
    TASK_CODE = 4
    RadioOpstionClickEvent "특수직책"
End Sub

'--//직분이력옵션 변경 이벤트
Private Sub optTitle_Change()
    If Me.optTitle Then
        Me.cboTitle.Visible = True
        Me.txtChurchNow.Enabled = False
        
        Me.lblTitleOrdinaryDate.Visible = True
        Me.txtTitleOrdinaryDate.Visible = True
        
        Dim objTitleDao As New TitleDao
        Dim titleList As Object
        
        Set titleList = objTitleDao.GetTitleList
        
        Dim strTitle As Variant
        For Each strTitle In titleList
            Me.cboTitle.AddItem strTitle
        Next
        
        Me.cmdSearch_Church.Visible = False
    Else
        Me.cboTitle.Visible = False
        Me.txtChurchNow.Enabled = True
        Me.cboTitle.Clear
        Me.cmdSearch_Church.Visible = True
        Me.lblTitleOrdinaryDate.Visible = False
        Me.txtTitleOrdinaryDate.Visible = False
    End If
End Sub

'--//직분이력옵션 클릭 이벤트
Private Sub optTitle_Click()
    TASK_CODE = 2
    RadioOpstionClickEvent "직분"
End Sub

'--//발령이력옵션 클릭 이벤트
Private Sub optTransfer_Click()
    TASK_CODE = 1
    RadioOpstionClickEvent "소속교회"
    Me.lblKind2.Caption = "교회명"
End Sub

'--//검색어 변경 이벤트
Private Sub txtChurchNM_Change()
    Me.txtChurchNM.BackColor = RGB(255, 255, 255)
End Sub

'--//종료일 변경 이벤트
Private Sub txtEnd_Change()
    Call Date_Format(Me.txtEnd)
End Sub

'--//시작일 변경 이벤트
Private Sub txtStart_Change()
    Call Date_Format(Me.txtStart)
End Sub

'--//직분안수일 변경 이벤트
Private Sub txtTitleOrdinaryDate_Change()
    Call Date_Format(Me.txtTitleOrdinaryDate)
End Sub

'--//유저폼 초기화
Private Sub UserForm_Initialize()
    If connIP = Empty Then setGlobalVariant '//전역변수 재설정
    If checkLogin = 0 Then f_login.Show '//로그인체크
    If checkLogin = 0 Then End '--//로그인 실패 시 프로시저 종료
    
    '--//기초설정
    Me.cmdClose.Cancel = True
    
    '--//권한에 따른 컨트롤 설정
    HideDeleteButtonByUserAuth
    
    '--//컨트롤 설정
'    Me.txtChurchNM.Enabled = False
'    Me.txtChurchNM.BackColor = &HE0E0E0
    TextBoxEnable Me.txtChurchNM, False
    Me.cmdSearch.Enabled = False
    Me.lstPStaff.Enabled = False
    Me.lstHistory.Enabled = False
    Me.txtStart.Enabled = False
    Me.txtEnd.Enabled = False
    Me.txtChurchNow.Enabled = False
    Me.cmdEdit.Enabled = False
    Me.cmdAdd.Enabled = False
    Me.cmdAdd.Visible = False
    Me.cmdCancel.Visible = False
    Me.cmdNew.Enabled = False
    Me.txtChurchNow_sid.Enabled = False
    Me.cmdSearch_Church.Enabled = False
    Me.cmdDelete.Enabled = False
    Me.chkPresent.Visible = False
    Me.cboTitle.Visible = False
    Me.cboTitle.Enabled = False
    
    '--//리스트박스 설정
    With Me.lstPStaff
        .ColumnCount = 4
        .ColumnHeads = False
        .ColumnWidths = "0,120,70,50" '생명번호, 교회명, 한글이름(직분), 직책
        .TextAlign = fmTextAlignLeft
        .Font = "굴림"
    End With
    
End Sub

'--//검색 버튼 클릭 이벤트
Private Sub cmdSearch_Click()
    
    Me.lstPStaff.Clear
    Me.lstHistory.Clear
    
    Dim objPStaffInfo As New PStaffInfoView
    Dim objPStaffInfoDao As New PStaffInfoViewDao
    Dim pStaffInfoList As Object
    '--//DB에서 목록을 받아옵니다.
    Set pStaffInfoList = objPStaffInfoDao.FindBySearchText(Me.txtChurchNM, Me.chkAll.Value)
    
    '--//받아온 목록이 없다면
    If pStaffInfoList.Count = 0 Then
        MsgBox "반환할 DB 데이터가 없습니다.", vbInformation, banner
        Exit Sub
    End If
    
    Dim tmpPStaffInfo As New PStaffInfoView
    With Me.lstPStaff
        '--//받아온 목록을 lstPStaff에 추가 합니다.
        For Each tmpPStaffInfo In pStaffInfoList
            Me.lstPStaff.AddItem tmpPStaffInfo.lifeNo
            .List(.ListCount - 1, 1) = tmpPStaffInfo.ChurchNameKo
            .List(.ListCount - 1, 2) = tmpPStaffInfo.NameKoAndTitle
            .List(.ListCount - 1, 3) = tmpPStaffInfo.position
            If Me.optTransfer.Value = False Then '--//발령이력은 배우자 나타나지 않기
                If tmpPStaffInfo.lifeNoSpouse <> "" Then '--//배우자가 있는 경우 리스트에 삽입
                    Me.lstPStaff.AddItem tmpPStaffInfo.lifeNoSpouse
                    .List(.ListCount - 1, 1) = tmpPStaffInfo.ChurchNameKo
                    .List(.ListCount - 1, 2) = tmpPStaffInfo.NameKoAndTitleSpouse
                    .List(.ListCount - 1, 3) = tmpPStaffInfo.PositionSpouse
                End If
            End If
        Next
    End With
    Me.lstPStaff.Enabled = True
    
End Sub

'--//입력 컨트롤을 초기화 합니다.
Sub TextBox_Initialize()
    Me.chkPresent.Value = False
    Me.txtStart.Value = ""
    Me.txtEnd.Value = ""
    Me.txtTitleOrdinaryDate.Value = ""
    Me.txtChurchNow.Value = ""
    Me.txtChurchNow_sid.Value = ""
    Me.cboTitle.Value = ""
End Sub

'--//입력된 값을 검증합니다.
Private Function fnData_Validation()
'---------------------------------------
'유저폼 입력값에 대한 데이터 유효성 검사
'TRUE: 이상없음, FALSE: 잘못됨.
'---------------------------------------
    Dim sql As String
    Dim tRecordSet As T_RECORD_SET
    Dim tmpList As Object
    Set tmpList = CreateObject("System.Collections.ArrayList")
    
    fnData_Validation = True '데이터가 유효하다는 가정 하에 시작
    
    Select Case TASK_CODE
        Case 2
            Dim objTitleDao As New TitleDao
            Set tmpList = objTitleDao.GetTitleList
            
            '--//지정되지 않은 값을 입력하면
            If Not tmpList.Contains(Me.txtChurchNow.text) Then
                MsgBox "직분을 잘못 입력하였습니다. 다시 확인 해주세요.", vbCritical, banner
                fnData_Validation = False: Exit Function
            End If
        Case 3
            Dim objPositionDao As New PositionDao
            Set tmpList = objPositionDao.GetPositionList
            
            '--//지정되지 않은 값을 입력하면
            If Not tmpList.Contains(Me.txtChurchNow.text) Then
                MsgBox "직책을 잘못 입력하였습니다. 다시 확인 해주세요.", vbCritical, banner
                fnData_Validation = False: Exit Function
            End If
        Case 4
            Dim objSpecialPosition As New SpecialPositionDao
            Set tmpList = objSpecialPosition.GetSpecialPositionList
            
            '--//지정되지 않은 값을 입력하면
            If Not tmpList.Contains(Me.txtChurchNow.text) Then
                MsgBox "특수직책을 잘못 입력하였습니다. 다시 확인 해주세요.", vbCritical, banner
                fnData_Validation = False: Exit Function
            End If
    Case Else
    End Select
    
    '--//날짜형식 검증
    If Not IsDate(Me.txtStart) Then
        MsgBox "올바른 날짜 형태가 아닙니다. 시작일을 다시 확인 해주세요.", vbCritical, banner
        Set txtBox_Focus = Me.txtStart: fnData_Validation = False: Exit Function
    End If
    
    If Not IsDate(Me.txtEnd) And Me.txtEnd <> "현재" Then
        MsgBox "올바른 날짜 형태가 아닙니다. 종료일을 다시 확인 해주세요.", vbCritical, banner
        Set txtBox_Focus = Me.txtEnd: fnData_Validation = False: Exit Function
    End If
    
    If Me.txtEnd <> "현재" Then
        If CDate(Me.txtEnd) <= CDate(Me.txtStart) Then
            MsgBox "종료일은 시작일보다 작거나 같을 수 없습니다.", vbCritical, banner
            fnData_Validation = False: Exit Function
        End If
    End If
    
End Function

'--//DELETE_ITEM 권한이 있는 USER에게만 삭제 버튼이 보이도록 합니다.
Private Sub HideDeleteButtonByUserAuth()
    
    Dim authList As Object
    Dim objUserDao As New UserDao
    
    Set authList = objUserDao.GetUserAuthorities
    
    Dim strAuth As Variant
    If authList.Contains("DELETE_ITEM") Then
        Me.cmdDelete.Visible = True
    Else
        Me.cmdDelete.Visible = False
    End If
    
End Sub

'--//입력모드 활성화/비활성화
Private Sub InputModeActivate(ByVal blnActivate As Boolean)
    
    '--//버튼 활성화/비활성화
    Me.cmdNew.Visible = Not blnActivate
    Me.cmdEdit.Visible = Not blnActivate
    Me.cmdDelete.Visible = Not blnActivate
    Me.cmdCancel.Visible = blnActivate
    Me.cmdAdd.Visible = blnActivate
    Me.cmdAdd.Enabled = blnActivate
    If Me.optTransfer.Value = True Then
        Me.cmdSearch_Church.Enabled = blnActivate
        Me.txtChurchNow.Enabled = False
    Else
        Me.txtChurchNow.Enabled = blnActivate
    End If
    
    '--//컨트롤 활성화/비활성화
    Me.Frame1.Enabled = Not blnActivate
    Me.txtChurchNM.Enabled = Not blnActivate
    Me.cmdSearch.Enabled = Not blnActivate
    Me.lstPStaff.Enabled = Not blnActivate
    Me.lstHistory.Enabled = Not blnActivate
    Me.chkAll.Enabled = Not blnActivate
    Me.txtStart.Enabled = blnActivate
    Me.txtEnd.Enabled = blnActivate
    Me.chkPresent.Visible = blnActivate
    Me.cboTitle.Enabled = blnActivate
    
    '--//입력 컨트롤 초기화
    If blnActivate Then
        TextBox_Initialize
        Me.chkPresent.Value = blnActivate
    End If
    
End Sub
