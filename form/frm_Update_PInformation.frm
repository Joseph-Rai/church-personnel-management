VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Update_PInformation 
   Caption         =   "인사자료 관리마법사"
   ClientHeight    =   10530
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   16365
   OleObjectBlob   =   "frm_Update_PInformation.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "frm_Update_PInformation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim queryKey As Integer '//리스트 위치 반환에 사용될 id
Dim txtBox_Focus As MSForms.textBox
Dim errorMessageAppeared As Boolean '--//재귀함수에서 오류메시지가 호출 깊이 수만큼 출력되는 것을 방지

Public Enum INPUTMODE
    INPUT_FOR_ALL
    INPUT_FOR_PART
End Enum

Private Sub chkChild1_Click()
    
    Dim control As MSForms.control

    If Me.chkChild1 = True Then
        If Me.txtLifeNo_Child1 = "" Then
            '--//Child1 생명번호가 없으면 Input Mode
            ActivateInputMode Me.frameChild1.controls, True, INPUT_FOR_PART
        Else
            '--//Child1 생명번호가 있으면 Frame 내 Controls 단순 활성화
            ActivateFrame Me.frameChild1, True
        End If
    Else
        '--//Child1 생명번호가 없으면 Frame 내 Controls 비활성화
        ActivateFrame Me.frameChild1, False
    End If
    
End Sub

Private Sub chkChild2_Click()

    Dim control As MSForms.control
    
    If Me.chkChild2 = True Then
        If Me.txtLifeNo_Child2 = "" Then
            '--//Child2 생명번호가 없으면 Input Mode
            ActivateInputMode Me.frameChild2.controls, True, INPUT_FOR_PART
        Else
            '--//Child2 생명번호가 있으면 Frame 내 Controls 단순 활성화
            ActivateFrame Me.frameChild2, True
        End If
    Else
        '--//Child2 생명번호가 없으면 Frame 내 Controls 비활성화
        ActivateFrame Me.frameChild2, False
    End If
    
End Sub

Private Sub chkChild3_Click()
    
    Dim control As MSForms.control
    
    If Me.chkChild3 = True Then
        If Me.txtLifeNo_Child3 = "" Then
            '--//Child3 생명번호가 없으면 Input Mode
            ActivateInputMode Me.frameChild3.controls, True, INPUT_FOR_PART
        Else
            '--//Child3 생명번호가 있으면 Frame 내 Controls 단순 활성화
            ActivateFrame Me.frameChild3, True
        End If
    Else
        '--//Child3 생명번호가 없으면 Frame 내 Controls 비활성화
        ActivateFrame Me.frameChild3, False
    End If
    
End Sub

Private Sub chkSpouse_Click()
    
    Dim control As MSForms.control
    
    If Me.chkSpouse = True Then
        If Me.txtLifeNo_Spouse = "" Then
            '--//Spouse 생명번호가 없으면 Input Mode
            ActivateInputMode Me.frameSpouse.controls, True, INPUT_FOR_PART
        Else
            '--//Spouse 생명번호가 있으면 Frame 내 Controls 단순 활성화
            ActivateFrame Me.frameSpouse, True
            Me.cmdTransferSpouse.Enabled = True
        End If
    Else
        '--//Spouse 생명번호가 없으면 Frame 내 Controls 비활성화
        ActivateFrame Me.frameSpouse, False
        Me.cmdTransferSpouse.Enabled = False
    End If
    
End Sub

Private Sub cmdCopyLifeNo_Child1_Click()
    If Me.txtLifeNo_Child1 = Empty Then Exit Sub
    CopyStartAnimation Me.cmdCopyLifeNo_Child1
    CopyToClipboard Me.txtLifeNo_Child1
    CopyEndAnimation Me.cmdCopyLifeNo_Child1
End Sub

Private Sub cmdCopyLifeNo_Child2_Click()
    If Me.txtLifeNo_Child2 = Empty Then Exit Sub
    CopyStartAnimation Me.cmdCopyLifeNo_Child2
    CopyToClipboard Me.txtLifeNo_Child2
    CopyEndAnimation Me.cmdCopyLifeNo_Child2
End Sub

Private Sub cmdCopyLifeNo_Child3_Click()
    If Me.txtLifeNo_Child3 = Empty Then Exit Sub
    CopyStartAnimation Me.cmdCopyLifeNo_Child3
    CopyToClipboard Me.txtLifeNo_Child3
    CopyEndAnimation Me.cmdCopyLifeNo_Child3
End Sub

Private Sub cmdCopyLifeNo_Click()
    If Me.txtLifeNo = Empty Then Exit Sub
    CopyStartAnimation Me.cmdCopyLifeNo
    CopyToClipboard Me.txtLifeNo
    CopyEndAnimation Me.cmdCopyLifeNo
End Sub

Private Sub cmdCopyLifeNo_Spouse_Click()
    If Me.txtLifeNo_Spouse = Empty Then Exit Sub
    CopyStartAnimation Me.cmdCopyLifeNo_Spouse
    CopyToClipboard Me.txtLifeNo_Spouse
    CopyEndAnimation Me.cmdCopyLifeNo_Spouse
End Sub

Private Sub cmdTransfer_Click()

    Dim pStaff As New PastoralStaff
    Dim pWife As New PastoralWife
    Dim pStaffDao As New PastoralStaffDao
    Dim pWifeDao As New PastoralWifeDao
    
    pStaff.ParseFromForm Me
    pWife.ParseFromForm Me

    If Me.chkTransfer = False Then
        '--//업데이트 전 확인절차
        If MsgBox("타부서에서 동일한 생명번호의 선지자를 신규 추가하면 현재 데이터가 이관처리 됩니다." & vbNewLine & _
            "계속 하시겠습니까?", vbQuestion + vbYesNo, banner) = vbNo Then
            Exit Sub
        End If
        
        Me.chkTransfer.Value = True
        
        '--//pWife도 한꺼번에 움직이는 것임을 상기시키는 메시지
        If pWife.lifeNo <> "" Then
            Me.chkTransferSpouse.Value = True
            MsgBox "배우자도 함께 이동처리 되도록 『" & Me.cmdTransferSpouse.Caption & "』도 함께 체크 됩니다.", vbOKOnly, banner
        End If
    Else
        '--//업데이트 전 확인절차
        If MsgBox("타부서 이동처리가 취소됩니다. 타부서에서 데이터를 이관받을 수 없습니다." & vbNewLine & _
            "계속 하시겠습니까?", vbQuestion + vbYesNo, banner) = vbNo Then
            Exit Sub
        End If
        
        '--//pStaff의 Suspend 업데이트
        Me.chkTransfer.Value = False
        
        '--//pWife도 한꺼번에 움직이는 것임을 상기시키는 메시지
        If pWife.lifeNo <> "" Then
            Me.chkTransferSpouse.Value = False
            MsgBox "『" & Me.cmdTransferSpouse.Caption & "』도 함께 체크 해제 되었습니다.", vbOKOnly, banner
        End If
    End If
    
    '--//pStaff의 Suspend 업데이트
    pStaff.Suspend = Me.chkTransfer.Value
    If pWife.lifeNo <> "" Then
        pWife.Suspend = Me.chkTransferSpouse.Value
    End If
    
    '--//변경된 내용 저장
    pStaffDao.Save pStaff
    pWifeDao.Save pWife
    
    '--//페이지 새로고침
    If Me.lstPStaff.listIndex <> -1 Then
        lstPStaff_Click
    End If

End Sub

Private Sub cmdTransferSpouse_Click()
    
    '--//혼인관계가 유지된 상태에서 배우자만 이동하면 참조 무결성을 해치므로 허용하지 않음
    If Me.chkTransferSpouse = True Then
        If Me.chkTransfer.Value = True Then
            MsgBox "배우자만 단독으로 해제할 수 없습니다." & vbNewLine & _
                    "배우자만 체크하길 원하시면 관리대상 체크를 먼저 해제한 후 진행하세요.", vbCritical + vbOKOnly, banner
            Exit Sub
        End If
        
        If MsgBox("배우자의 타부서 이동처리가 취소됩니다. 타부서에서 데이터를 이관받을 수 없습니다." & vbNewLine & _
            "계속 하시겠습니까?", vbQuestion + vbYesNo, banner) = vbNo Then
            Exit Sub
        Else
            Me.chkTransferSpouse.Value = False
        End If
    Else
        If MsgBox("이 기능은 사모의 남편이 변경되었을 경우에만 사용합니다." & vbNewLine & _
                    "계속 진행하시겠습니까", vbQuestion + vbYesNo, banner) = vbNo Then
            Exit Sub
        Else
            Me.chkTransferSpouse.Value = True
        End If
    End If
    
    Dim pWife As New PastoralWife
    Dim pWifeDao As New PastoralWifeDao
    
    pWife.ParseFromForm Me
    pWife.Suspend = Me.chkTransferSpouse.Value
    
    pWifeDao.Save pWife
    
    '--//페이지 새로고침
    If Me.lstPStaff.listIndex <> -1 Then
        lstPStaff_Click
    End If
    
End Sub

Private Sub cmdDelete_Child1_Click()

    Dim pStaff As New PastoralStaff
    Dim pStaffDao As New PastoralStaffDao
    
    pStaff.ParseFromForm Me
    
    '--//체크박스에 체크 해제 시 삭제할 자녀 정보 없으므로 프로시저 종료
    If pStaff.LifeNoChild1 = "" Then
        MsgBox "삭제할 자녀1 정보가 없습니다.", vbCritical, banner
        Exit Sub
    End If
    
    '--//삭제 진행여부 재확인
    If MsgBox("자녀1 정보가 삭제됩니다. 정말 진행 하시겠습니까?", vbQuestion + vbYesNo, banner) = vbNo Then
        Exit Sub
    End If
    
    '--//자녀1 정보 삭제
    pStaffDao.DeleteChild1 pStaff
    
    '--//메세지 박스
    MsgBox "자녀1 정보가 삭제 되었습니다.", , banner
    Call lstPStaff_Click
    
End Sub

Private Sub cmdDelete_Child2_Click()
    
    Dim pStaff As New PastoralStaff
    Dim pStaffDao As New PastoralStaffDao
    
    pStaff.ParseFromForm Me
    
    If pStaff.LifeNoChild2 = "" Then
        MsgBox "삭제할 자녀2 정보가 없습니다.", vbCritical, banner
        Exit Sub
    End If
    
    '--//삭제 진행여부 재확인
    If MsgBox("자녀2 정보가 삭제됩니다. 정말 진행 하시겠습니까?", vbQuestion + vbYesNo, banner) = vbNo Then
        Exit Sub
    End If
    
    '--//자녀2 정보 삭제
    pStaffDao.DeleteChild2 pStaff


    '--//메세지 박스
    MsgBox "자녀2 정보가 삭제 되었습니다.", , banner
    Call lstPStaff_Click
End Sub

Private Sub cmdDelete_Child3_Click()
    
    Dim pStaff As New PastoralStaff
    Dim pStaffDao As New PastoralStaffDao
    
    pStaff.ParseFromForm Me
    
    If pStaff.LifeNoChild3 = "" Then
        MsgBox "삭제할 자녀3 정보가 없습니다.", vbCritical, banner
        Exit Sub
    End If
    
    '--//삭제 진행여부 재확인
    If MsgBox("자녀3 정보가 삭제됩니다. 정말 진행 하시겠습니까?", vbQuestion + vbYesNo, banner) = vbNo Then
        Exit Sub
    End If
    
    '--//자녀3 정보 삭제
    pStaffDao.DeleteChild3 pStaff

    '--//메세지 박스
    MsgBox "자녀3 정보가 삭제 되었습니다.", , banner
    Call lstPStaff_Click
End Sub

Private Sub cmdDelete_Spouse_Click()
    
    Dim pWife As New PastoralWife
    Dim pWifeDao As New PastoralWifeDao
    
    pWife.ParseFromForm Me
    
    If pWife.lifeNo = "" Then
        MsgBox "삭제할 배우자 정보가 없습니다.", vbCritical, banner
        Exit Sub
    End If
    
    '--//삭제 진행여부 재확인
    If MsgBox("배우자 정보가 삭제됩니다. 정말 진행 하시겠습니까?", vbQuestion + vbYesNo, banner) = vbNo Then
        Exit Sub
    End If
    
    '--//배우자 정보 삭제
    pWifeDao.Delete pWife

    '--//메세지 박스
    MsgBox "배우자 정보가 삭제 되었습니다.", , banner
    lstPStaff_Click
    
End Sub

Private Sub cmdExtendChild_Click()
    Me.Width = 820
    Me.cmdExtendChild.Visible = False
    Me.cmdMinimize.Visible = True
End Sub

Private Sub cmdMinimize_Click()
    Me.Width = 545
    Me.cmdMinimize.Visible = False
    Me.cmdExtendChild.Visible = True
End Sub

Private Sub cmdFamily_Click()
    argShow2 = 1
    frm_Update_FamilyInfo.Show
End Sub

Private Sub cmdFamily_Spouse_Click()
    argShow2 = 2
    frm_Update_FamilyInfo.Show
End Sub

Private Sub savePhoto(lifeNo As String, filePath As String)

    Dim pPhoto As New PStaffPhoto
    Dim pPhotoDao As New PStaffPhotoDao
    Dim stream As New ADODB.stream
    
    stream.Type = adTypeBinary
    stream.Open
    stream.LoadFromFile filePath
    
    pPhoto.lifeNo = lifeNo
    pPhoto.photo = encodeBase64(stream.Read)
    pPhotoDao.Save pPhoto

End Sub

Private Sub lblPic_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    
    Dim filePath As String
    
    filePath = Application.GetOpenFilename("Image File Format, *.jpg;*.jpeg; *.bmp; *.tif; *.png", , "Photo Select")
    
    If filePath = "False" Then
        Exit Sub
    End If
    
    '--//DB에 사진 저장
    savePhoto Me.txtLifeNo, filePath
    
    '--//성공메시지
    MsgBox "사진이 성공적으로 저장되었습니다.", , banner
    
    '--//라벨에 사진삽입
    InsertPicToLabel Me.lblPic, Me.txtLifeNo
    
End Sub

Private Sub lblPic_Spouse_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    
    If Me.txtLifeNo_Spouse = "" Then
        Exit Sub
    End If
    
    Dim filePath As String
    
    filePath = Application.GetOpenFilename("Image File Format, *.jpg;*.jpeg; *.bmp; *.tif; *.png", , "Photo Select")
    
    If filePath = "False" Then
        Exit Sub
    End If
    
    '--//DB에 사진 저장
    savePhoto Me.txtLifeNo_Spouse, filePath
    
    '--//성공메시지
    MsgBox "사진이 성공적으로 저장되었습니다.", , banner
    
    '--//라벨에 사진삽입
    InsertPicToLabel Me.lblPic_Spouse, Me.txtLifeNo_Spouse
    
End Sub

Private Sub lblPic_Child1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim filePath As String
    
    filePath = Application.GetOpenFilename("Image File Format, *.jpg;*.jpeg; *.bmp; *.tif; *.png", , "Photo Select")
    
    If filePath = "False" Then
        Exit Sub
    End If
    
    '--//DB에 사진 저장
    savePhoto Me.txtLifeNo_Child1, filePath
    
    '--//성공메시지
    MsgBox "사진이 성공적으로 저장되었습니다.", , banner
    
    '--//라벨에 사진삽입
    InsertPicToLabel Me.lblPic_Child1, Me.txtLifeNo_Child1
    
End Sub

Private Sub lblPic_Child2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim filePath As String
    
    filePath = Application.GetOpenFilename("Image File Format, *.jpg;*.jpeg; *.bmp; *.tif; *.png", , "Photo Select")
    
    If filePath = "False" Then
        Exit Sub
    End If
    
    '--//DB에 사진 저장
    savePhoto Me.txtLifeNo_Child2, filePath
    
    '--//성공메시지
    MsgBox "사진이 성공적으로 저장되었습니다.", , banner
    
    '--//라벨에 사진삽입
    InsertPicToLabel Me.lblPic_Child2, Me.txtLifeNo_Child2
    
End Sub

Private Sub lblPic_Child3_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim filePath As String
    
    filePath = Application.GetOpenFilename("Image File Format, *.jpg;*.jpeg; *.bmp; *.tif; *.png", , "Photo Select")
    
    If filePath = "False" Then
        Exit Sub
    End If
    
    '--//DB에 사진 저장
    savePhoto Me.txtLifeNo_Child3, filePath
    
    '--//성공메시지
    MsgBox "사진이 성공적으로 저장되었습니다.", , banner
    
    '--//라벨에 사진삽입
    InsertPicToLabel Me.lblPic_Child3, Me.txtLifeNo_Child3
    
End Sub

Private Sub lstPStaff_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Un'HookListBoxScroll
End Sub

Private Sub lstPStaff_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Me.lstPStaff.ListCount > 0 Then
        'HookListBoxScroll Me, Me.lstPStaff
    End If
End Sub

Private Sub txtBirthday_Change()
    Call Date_Format(Me.txtBirthday)
End Sub

Private Sub txtBirthday_Child1_Change()
    Call Date_Format(Me.txtBirthday_Child1)
End Sub

Private Sub txtBirthday_Child2_Change()
    Call Date_Format(Me.txtBirthday_Child2)
End Sub

Private Sub txtBirthday_Child3_Change()
    Call Date_Format(Me.txtBirthday_Child3)
End Sub

Private Sub txtBirthday_Spouse_Change()
    Call Date_Format(Me.txtBirthday_Spouse)
End Sub

Private Sub txtLifeNo_Change()
'    Me.chkTransfer.Enabled = True
End Sub

Private Sub CopyStartAnimation(cmdBox As MSForms.CommandButton)
    
    Dim i As Integer
    
    For i = 0 To 25
'        Sleep 8
        cmdBox.Left = cmdBox.Left - 1
        cmdBox.Width = cmdBox.Width + 1
        Me.Repaint
    Next
    
    With cmdBox
        .Caption = "Copied"
        .BackColor = &HC0C0FF
    End With
    Me.Repaint
    
    Application.Wait DateAdd("s", 1, Now)
'    Sleep 700
    
End Sub

Private Sub CopyEndAnimation(cmdBox As MSForms.CommandButton)

    Dim i As Integer
    
    For i = 0 To 25
'        Sleep 8
        cmdBox.Left = cmdBox.Left + 1
        cmdBox.Width = cmdBox.Width - 1
        Me.Repaint
    Next
    
    With cmdBox
        .Caption = "..."
        .BackColor = &HC0FFC0
    End With
    Me.Repaint

End Sub

Private Sub CopyToClipboard(text As String)
    On Error GoTo ErrHandler
    Static obj As DataObject
    
    If obj Is Nothing Then Set obj = New DataObject
    obj.SetText text
    obj.PutInClipboard
ErrHandler:
End Sub

Private Sub txtNationality_Enter()
    If Me.txtNationality = "" Then
        argShow = 3
        frm_Search_Country.Show
    End If
End Sub
Private Sub txtNationality_Spouse_Enter()
    If Me.txtNationality_Spouse = "" Then
        argShow = 4
        frm_Search_Country.Show
    End If
End Sub

Private Sub txtOrdinationPrayer_dt_Change()
    Call Date_Format(Me.txtOrdinationPrayer_dt)
End Sub

Private Sub txtOvs_dt_Change()
    Call Date_Format(Me.txtOvs_dt)
End Sub

Private Sub txtSalary_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Dim su As Double
    
    Me.txtSalary = fnExtract(Me.txtSalary)
    
    If Len(Me.txtSalary.Value) = 0 Then
        su = 0
    Else
        su = CDbl(Me.txtSalary.text)
    End If

    Me.txtSalary.text = Format(su, "#,##0")
    Me.txtSalary.SelStart = Len(Me.txtSalary.Value)
End Sub

Private Sub txtWedding_dt_Change()
    Call Date_Format(Me.txtWedding_dt)
End Sub

Private Sub UserForm_Initialize()
    If connIP = Empty Then setGlobalVariant '//전역변수 재설정
    If checkLogin = 0 Then f_login.Show '//로그인체크
    If checkLogin = 0 Then End '--//로그인 실패 시 프로시저 종료
    
    '--//기초설정
    Me.cmdClose.Cancel = True
    Me.chkSpouse.Visible = False
    Me.chkChild1.Visible = False
    Me.chkChild2.Visible = False
    Me.chkChild3.Visible = False
    Me.cmdAdd.Visible = False
    Me.cmdCancel.Visible = False
    Me.cmdDelete_Spouse.Enabled = False
    Me.cmdDelete_Child1.Enabled = False
    Me.cmdDelete_Child2.Enabled = False
    Me.cmdDelete_Child3.Enabled = False
    Me.cmdTransfer.Enabled = False
    Me.cmdTransferSpouse.Enabled = False
    Me.chkTransfer.Enabled = False
    Me.chkTransferSpouse.Enabled = False
    
    '--//권한에 따른 컨트롤 설정
    HideDeleteButtonByUserAuth
    ExtraControlsEnable False
    cmdMinimize_Click
    
    '--//콤보박스 설정
    With Me.cboBaptism
        .AddItem "유"
        .AddItem "무"
    End With
    
    '--//리스트박스 설정
    With Me.lstPStaff
        .ColumnCount = 4
        .ColumnHeads = False
        .ColumnWidths = "0,120,70,50" '생명번호, 교회명, 한글이름(직분), 직책
        .TextAlign = fmTextAlignLeft
        .Font = "굴림"
    End With
    
    '--//검색창 Setfocus
    Me.txtSearchText.SetFocus
    
End Sub
Private Sub lstPStaff_Click()
    
    Dim filePath As String
    Dim FileName As String
    Dim control As MSForms.control
    
    '--//컨트롤 설정
    ExtraControlsEnable True
    Me.chkSpouse.Visible = True
    Me.chkChild1.Visible = True
    Me.chkChild2.Visible = True
    Me.chkChild3.Visible = True
    HideDeleteButtonByUserAuth
    
    Me.lblPic.Picture = LoadPicture("")
    Me.lblPic_Spouse.Picture = LoadPicture("")
    Me.lblPic_Child1.Picture = LoadPicture("")
    Me.lblPic_Child2.Picture = LoadPicture("")
    Me.lblPic_Child3.Picture = LoadPicture("")
    
    '--//인사데이터 텍스트박스에 추가
    If Me.lstPStaff.listIndex = -1 Then
        Exit Sub
    End If
    
    Dim pStaff As New PastoralStaff
    Dim pStaffDao As New PastoralStaffDao
    Dim pWife As New PastoralWife
    Dim pWifeDao As New PastoralWifeDao
    Dim lifeNo As String
    
    With Me.lstPStaff
        lifeNo = .List(.listIndex)
        Set pStaff = pStaffDao.FindByLifeNo(lifeNo)
        Set pWife = pWifeDao.FindByPStaff(pStaff)
    End With
    
    '--//양식 채우기
    FillOutForm pStaff, pWife
    
    '--//체크박스 클릭 프로시저 실행으로 비활성 텍스트박스 처리
    ActivateFrame Me.framePStaff, True
    chkSpouse_Click
    chkChild1_Click
    chkChild2_Click
    chkChild3_Click
    
    '--//레이블 색깔 원상복귀
    For Each control In Me.controls
        If TypeName(control) = "Label" And Not control.Name Like "*Info*" Then
            ChangeLabelColor control, vbBlack
        End If
    Next
    
    '--//Frame 외부 컨트롤 활성화/비활성화 따로 조절하기
    ComboBoxEnable Me.cboBaptism, True
    TextBoxEnable Me.txtWedding_dt, True
    TextBoxEnable Me.txtOrdinationPrayer_dt, True
    TextBoxEnable Me.txtSalary, True
    
    '--//외국인 해외발령일,생도기수 막아놓기
    TextBoxEnable Me.txtOvs_dt, Me.txtNationality = "대한민국"
    TextBoxEnable Me.txtTheological_Order, Me.txtNationality = "대한민국"
    
    '--//생명번호 있는 분들만 가족구성원 추가버튼 활성화
    ControlEnable Me.cmdFamily, Me.txtLifeNo <> ""
    ControlEnable Me.cmdFamily_Spouse, Me.txtLifeNo_Spouse <> ""
    ControlEnable Me.cmdTransfer, Me.txtLifeNo <> ""
    
    '--//배우자, 자녀1,2,3 있는 경우 체크박스 비활성화
    ControlEnable Me.chkSpouse, Me.txtLifeNo_Spouse = ""
    ControlEnable Me.chkChild1, Me.txtLifeNo_Child1 = ""
    ControlEnable Me.chkChild2, Me.txtLifeNo_Child2 = ""
    ControlEnable Me.chkChild3, Me.txtLifeNo_Child3 = ""
    ControlEnable Me.cmdDelete_Spouse, Me.txtLifeNo_Spouse <> ""
    ControlEnable Me.cmdDelete_Child1, Me.txtLifeNo_Child1 <> ""
    ControlEnable Me.cmdDelete_Child2, Me.txtLifeNo_Child2 <> ""
    ControlEnable Me.cmdDelete_Child3, Me.txtLifeNo_Child3 <> ""
    
    '--//선지자 사진추가
    InsertPicToLabel Me.lblPic, pStaff.lifeNo
    
    '--//자녀1 사진추가
    If Me.txtLifeNo_Child1 <> "" Then
        InsertPicToLabel Me.lblPic_Child1, pStaff.LifeNoChild1
    End If
    
    '--//자녀2 사진추가
    If Me.txtLifeNo_Child2 <> "" Then
        InsertPicToLabel Me.lblPic_Child2, pStaff.LifeNoChild2
    End If
    
    '--//자녀3 사진추가
    If Me.txtLifeNo_Child3 <> "" Then
        InsertPicToLabel Me.lblPic_Child3, pStaff.LifeNoChild3
    End If
    
    '--//배우자 사진추가
    If Me.txtLifeNo_Spouse <> "" Then
        InsertPicToLabel Me.lblPic_Spouse, pWife.lifeNo
    End If

End Sub

Private Sub cmdSearch_Click()
    
    Dim pStaffInfo As New PStaffInfoView
    Dim pStaffInfoDao As New PStaffInfoViewDao
    Dim pStaffInfoList As Object
    
    '--//검색조건에 따라 pStaffInfo 객체 리스트 불러오기
    Set pStaffInfoList = pStaffInfoDao.FindBySearchText(Me.txtSearchText, Me.chkAll.Value)
    
    '--//리스트 초기화
    Me.lstPStaff.Clear
    
    '--//검색된 목록 리스트에 채워넣기
    Dim pStaffInfoTemp As New PStaffInfoView
    With Me.lstPStaff
        For Each pStaffInfoTemp In pStaffInfoList
            Me.lstPStaff.AddItem pStaffInfoTemp.lifeNo
            .List(.ListCount - 1, 1) = pStaffInfoTemp.ChurchNameKo
            .List(.ListCount - 1, 2) = pStaffInfoTemp.NameKoAndTitle
            .List(.ListCount - 1, 3) = pStaffInfoTemp.position
        Next
    End With
End Sub
Private Sub txtChurch_Change()
    Me.txtChurch.BackColor = RGB(255, 255, 255)
End Sub
Private Sub cmdCancel_Click()
    
    Dim control As MSForms.control
    
    '--//컨트롤 원상복귀
'    HideCmdBtnForInput False
    ActivateInputMode Me.controls, False, INPUT_FOR_ALL
    
    If Me.txtSearchText <> "" Then
        lstPStaff_Click
    Else
        
'        ExtraControlsEnable False
'        '--//레이블 색깔 원상복귀
'        For Each control In Me.controls
'            If TypeName(control) = "Label" And Not control.name Like "*Info*" Then
'                ChangeLabelColor control, vbBlack
'            End If
'        Next
    End If
    
    '--//폼 새로고침
    cmdMinimize_Click
    lstPStaff_Click
    
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

'--//선지자와 배우자 모두 삭제합니다.
'--//자녀 정보도 선지자 정보와 한 테이블에 묶여 있으므로 함께 삭제 됩니다.
Private Sub cmdDelete_Click()
    
    If MsgBox("선택한 데이터를 삭제하시겠습니까?" & vbNewLine & "삭제된 데이터는 복구할 수 없습니다.", vbQuestion + vbYesNo, banner) = vbNo Then
        Exit Sub
    End If
    
    Dim pStaff As New PastoralStaff
    Dim pWife As New PastoralWife
    Dim pStaffDao As New PastoralStaffDao
    Dim pWifeDao As New PastoralWifeDao
    
    '--//폼에서 객체 파싱
    pStaff.ParseFromForm Me
    pWife.ParseFromForm Me
    
    If pWife.lifeNo <> "" Then
        If MsgBox("배우자도 함께 삭제 됩니다." & vbNewLine & "삭제된 데이터는 복구할 수 없습니다." & vbNewLine & _
                  "계속 하시겠습니까?", vbQuestion + vbYesNo, banner) = vbNo Then
            Exit Sub
        End If
    End If
    
    '--//DB에서 객체 삭제
    pStaffDao.Delete pStaff
    pWifeDao.Delete pWife
    
    '--//메세지박스
    MsgBox "해당 데이터가 삭제되었습니다.", , banner
    
    '--//리스트박스 새로고침
    InitializeTextBoxes Me.controls
    cmdSearch_Click
    lstPStaff_Click
    Me.lstPStaff.listIndex = -1
    
End Sub

'--//수정된 내용이 있으면 DB에 반영합니다.
Private Sub cmdEdit_Click()
    
    '--//데이터 유효성 검사
    If fnData_Validation(Me.controls) = False Then
On Error Resume Next
        txtBox_Focus.SetFocus
        txtBox_Focus.SelStart = 0
        txtBox_Focus.SelLength = Len(txtBox_Focus)
On Error GoTo 0
        Exit Sub
    End If
    
    '--//변수설정: pStaff1, pWife1 = 수정 후 객체 | pStaff2, pWife2 = 수정 전 객체(DB객체)
    '              pStaffDao, pWifeDao = 데이터 엑세스 오브젝트
    Dim pStaff1 As New PastoralStaff
    Dim pWife1 As New PastoralWife
    Dim pStaff2 As New PastoralStaff
    Dim pWife2 As New PastoralWife
    Dim pStaffDao As New PastoralStaffDao
    Dim pWifeDao As New PastoralWifeDao
    
    '--//수정 후 객체(폼 기반)
    pStaff1.ParseFromForm Me
    pWife1.ParseFromForm Me
    
    '--//수정 전 객체(DB 기반)
    Set pStaff2 = pStaffDao.FindByStaff(pStaff1)
    Set pWife2 = pWifeDao.FindByWife(pWife1)
    
    '--//PStaff 객체와 PWife 객체가 아무것도 수정된게 없다면 프로시저 종료
    If pStaff1.IsEqual(pStaff2, True) And pWife1.IsEqual(pWife2, True) Then
        Exit Sub
    End If
    
    '--//PStaff 객체 업데이트
    pStaffDao.Save pStaff1 '--//pStaff1의 정보 업데이트
    
    If pWife1.lifeNo <> "" Then
        Dim pStaff3 As New PastoralStaff '--// pStaff3: pWife1의 배우자로 등록된 선지자
        Set pStaff3 = pStaffDao.FindByLifeNo(pWife2.lifeNoSpouse)
        If pStaff1.IsEqual(pStaff3) Or pStaff3.lifeNo = "" Then
            '--//수정 전과 수정 후 배우자 정보가 동일하다면
            pWifeDao.Save pWife1 '--//pWife1의 정보 업데이트
        Else
            '--//수정 전과 수정 후 배우자 정보가 다르다면
            Dim OvsDept As New OvsDepartment
            Dim ovsDeptDao As New OvsDepartmentDao
            Set OvsDept = ovsDeptDao.FindById(pStaff3.OvsDept) '--// 관리부서
            MsgBox "입력하신 사모는 이미 다른 분의 배우자로 등록되어 있습니다." & vbNewLine & _
                    "등록을 원하시면 먼저 기존 배우자에서 삭제 해주세요." & vbNewLine & vbNewLine & _
                    "이름: " & pStaff3.nameKo & vbNewLine & _
                    "생명번호: " & pStaff3.lifeNo & vbNewLine & _
                    "관리부서: " & OvsDept.DeptName, vbYesNo, banner
            Exit Sub
        End If
    End If
    
    '--//메세지박스
    MsgBox "업데이트 되었습니다.", , banner
    
    '--//리스트박스 초기화
    Call lstPStaff_Click
    
End Sub

Private Sub cmdNew_Click()
    
    Dim control As MSForms.control
    
    '--//인풋모드 활성화
    ActivateInputMode Me.controls, True, INPUT_FOR_ALL
    
End Sub

Private Sub cmdADD_Click()
    
    '--//데이터 유효성 검사
    If fnData_Validation(Me.controls) = False Then
On Error Resume Next
        txtBox_Focus.SetFocus
        txtBox_Focus.SelStart = 0
        txtBox_Focus.SelLength = Len(txtBox_Focus)
On Error GoTo 0
        Exit Sub
    End If
    
    Dim pStaff As New PastoralStaff '--// DB에 실제 추가할 PastoralStaff 객체
    Dim pWife As New PastoralWife '--// DB에 실체 추가할 PastoralWife 객체
    Dim pWifeTemp As New PastoralWife '--// 중복체크를 위한 객체
    Dim pStaffTemp As New PastoralStaff '--// 중복체크를 위한 객체
    Dim pWifeDao As New PastoralWifeDao
    Dim pStaffDao As New PastoralStaffDao
    Dim iOk As Integer
    
    '--//관리대상이 이관되었을 경우, 사모는 이관하지 않을 수 없으므로 질문하지 않고 이관시키기 위한 플래그 변수
    Dim wasRunFlag As Boolean
    
    '--//폼 정보로부터 선지자, 사모 객체 파싱
    pStaff.ParseFromForm Me
    pWife.ParseFromForm Me
    
    '--//선지자 중복체크
    Set pStaffTemp = pStaffDao.FindByStaff(pStaff)
    If pStaff.IsEqual(pStaffTemp) Then
        '--//Suspend: False - 이동불가 / True - 이동가능
        If pStaffTemp.Suspend = False Then
            '--//중복된 경우 커스텀 에러발생
            On Error GoTo PSTAFF_IS_ALREADY_REGISTERED
            Dim OvsDept As New OvsDepartment
            Dim ovsDeptDao As New OvsDepartmentDao
            Set OvsDept = ovsDeptDao.FindById(pStaffTemp.OvsDept)
            err.Raise vbObjectError + _
                    ERR_CODE_PSTAFF_IS_ALREADY_REGISTERED, , _
                    ERR_DESC_PSTAFF_IS_ALREADY_REGISTERED & _
                    "관리부서 데이터 담당자 혹은 관리자에게 문의하세요." & vbNewLine & _
                    "관리부서: " & OvsDept.DeptName
        Else
            '--//이동가능할 시 이동처리
            iOk = MsgBox("기존에 등록된 선지자 정보가 이미 있습니다." & vbNewLine & vbNewLine & "불러오시겠습니까?", vbQuestion + vbYesNo, banner)
            If iOk = vbYes Then
                Set pStaff = pStaffTemp
                pStaff.OvsDept = USER_DEPT
                pStaff.Suspend = False
                wasRunFlag = True
            Else
                Exit Sub
            End If
        End If
    End If
    
    '--//pStaff 객체 DB에 저장
    pStaffDao.Save pStaff

'--------------------------------------------------------------------

    '--//배우자 중복체크
    If Me.chkSpouse.Value = False Or Me.txtLifeNo_Spouse = "" Then GoTo PASS1
    
    pWifeTemp = pWifeDao.FindByWifeAndSpouseLifeNo(pWife)
    If pWife.IsEqual(pWifeTemp) <> "" Then
        '--//Suspend: False - 이동불가 / True - 이동가능
        If pWifeTemp.Suspend = False Then
            '--//중복된 경우 커스텀 에러발생
            On Error GoTo WIFE_IS_ALREADY_REGISTERED_OF_OTHER
            Set pStaff = pStaffDao.FindByLifeNo(pWifeTemp.lifeNoSpouse)
            err.Raise vbObjectError + _
                ERR_CODE_WIFE_IS_ALREADY_REGISTERED_OF_OTHER, , _
                ERR_DESC_WIFE_IS_ALREADY_REGISTERED_OF_OTHER & _
                "등록된 배우자: " & pStaff.nameKo & "(" & pStaff.lifeNo & ")"
        Else
            '--//이동가능할 시 이동처리
            If wasRunFlag Then
                iOk = vbYes
            Else
                iOk = MsgBox("기존에 등록된 사모 정보가 이미 있습니다." & vbNewLine & vbNewLine & "불러오시겠습니까?", vbQuestion + vbYesNo, banner)
            End If
            If iOk = vbYes Then
                Set pWife = pWifeTemp
                pWife.OvsDept = USER_DEPT
                pWife.lifeNoSpouse = pStaff.lifeNo
                pWife.Suspend = False
            Else
                Exit Sub
            End If
        End If
    End If
    
    '--//pWife 객체 DB에 저장
    pWifeDao.Save pWife
    
PASS1:

    '--//성공메세지
    MsgBox "추가 되었습니다." & vbNewLine & vbNewLine & "발령 이력을 반드시 추가해주세요.", , banner
    
    '--//추가된 인원 검색하여 목록에 띄우기
    Me.txtSearchText = pStaff.lifeNo
    Me.chkAll.Value = True
    Call cmdSearch_Click
    Call lstPStaff_Click
On Error Resume Next
    Me.lstPStaff.listIndex = 0
On Error GoTo 0
    
    '--//버튼설정 원래대로
    HideCmdBtnForInput False
    HideDeleteButtonByUserAuth
    
DONE:
    Exit Sub
    
PSTAFF_IS_ALREADY_REGISTERED:
    MsgBox err.Description, vbCritical, banner
    Exit Sub
    
WIFE_IS_ALREADY_REGISTERED_OF_OTHER:
    MsgBox err.Description, vbCritical, banner
    queryKey = pWife.lifeNoSpouse
    Call returnListPosition3(Me, Me.lstPStaff.Name, queryKey)
    Exit Sub
End Sub

'---------------------------------------
'유저폼 입력값에 대한 데이터 유효성 검사
'TRUE: 이상없음, FALSE: 잘못됨.
'---------------------------------------
Private Function fnData_Validation(controls As MSForms.controls) As Boolean
    
    '--//유효성 오류메세지를 위한 Map 객체 생성
    Dim messageMap As New Collection
    messageMap.Add "선지자 생년월일", Me.txtBirthday.Name
    messageMap.Add "자녀1 생년월일", Me.txtBirthday_Child1.Name
    messageMap.Add "자녀2 생년월일", Me.txtBirthday_Child2.Name
    messageMap.Add "자녀3 생년월일", Me.txtBirthday_Child3.Name
    messageMap.Add "배우자 생년월일", Me.txtBirthday_Spouse.Name
    messageMap.Add "해외 최초 발령일", Me.txtOvs_dt.Name
    messageMap.Add "혼인일", Me.txtWedding_dt.Name
    messageMap.Add "안수일", Me.txtOrdinationPrayer_dt.Name
    messageMap.Add "선지자 국적", Me.txtNationality.Name
    messageMap.Add "배우자 국적", Me.txtNationality_Spouse.Name
    messageMap.Add "선지자 생명번호", Me.txtLifeNo.Name
    messageMap.Add "자녀1 생명번호", Me.txtLifeNo_Child1.Name
    messageMap.Add "자녀2 생명번호", Me.txtLifeNo_Child2.Name
    messageMap.Add "자녀3 생명번호", Me.txtLifeNo_Child3.Name
    messageMap.Add "배우자 생명번호", Me.txtLifeNo_Spouse.Name
    messageMap.Add "선지자 한글이름", Me.txtName_ko.Name
    messageMap.Add "배우자 한글이름", Me.txtName_Spouse_ko.Name
    messageMap.Add "자녀1 한글이름", Me.txtName_Child1_ko.Name
    messageMap.Add "자녀2 한글이름", Me.txtName_Child2_ko.Name
    messageMap.Add "자녀3 한글이름", Me.txtName_Child3_ko.Name
    messageMap.Add "선지자 영문이름", Me.txtName_en.Name
    messageMap.Add "배우자 영문이름", Me.txtName_Spouse_en.Name
    messageMap.Add "자녀1 영문이름", Me.txtName_Child1_en.Name
    messageMap.Add "자녀2 영문이름", Me.txtName_Child2_en.Name
    messageMap.Add "자녀3 영문이름", Me.txtName_Child3_en.Name
    
    
    '--//데이터가 유효하다는 가정 하에 시작
    fnData_Validation = True
    
    Dim control As MSForms.control
    '--//control 순회하면서 점검
On Error GoTo INVALID_INPUT_DATA
    errorMessageAppeared = False
    For Each control In controls
        Select Case TypeName(control)
            Case "TextBox":
                Set txtBox_Focus = control
                '--//필수값인 데이터가 입력되지 않았을 때
                If IsEmptyRequiredControl(control) Then
                    err.Raise vbObjectError + ERR_CODE_REQUIRED_INPUT_PINFORMATION, , _
                        StringFormat(ERR_DESC_REQUIRED_INPUT_PINFORMATION, messageMap.Item(control.Name))
                '--//데이터 형식이 잘못 입력 되었을 때(생명번호, 날짜 등)
                ElseIf IsInvalidInput(control) Then
                    err.Raise vbObjectError + ERR_CODE_INVALID_INPUT_PINFORMATION, , _
                        StringFormat(ERR_DESC_INVALID_INPUT_PINFORMATION, messageMap.Item(control.Name))
                End If
            Case "Frame":
                fnData_Validation control.controls
        End Select
    Next

SUCCEED:
    Exit Function

INVALID_INPUT_DATA:
    fnData_Validation = False
    If Not errorMessageAppeared Then
        MsgBox err.Description, vbCritical, banner
        errorMessageAppeared = True
    End If
End Function

Private Function GetRequiredMap() As Collection

    '--//필수값 입력대상 TextBox 리스트 생성
    Dim requiredMap As New Collection
    requiredMap.Add "관리자 생명번호", Me.txtLifeNo.Name
    requiredMap.Add "관리자 한글이름", Me.txtName_ko.Name
    requiredMap.Add "관리자 영문이름", Me.txtName_en.Name
    requiredMap.Add "관리자 생년월일", Me.txtBirthday.Name
    requiredMap.Add "관리자 국적", Me.txtNationality.Name
    requiredMap.Add "배우자 생명번호", Me.txtLifeNo_Spouse.Name
    requiredMap.Add "배우자 한글이름", Me.txtName_Spouse_ko.Name
    requiredMap.Add "배우자 영문이름", Me.txtName_Spouse_en.Name
    requiredMap.Add "배우자 생년월일", Me.txtBirthday_Spouse.Name
    requiredMap.Add "배우자 국적", Me.txtNationality_Spouse.Name
    requiredMap.Add "자녀1 생명번호", Me.txtLifeNo_Child1.Name
    requiredMap.Add "자녀1 한글이름", Me.txtName_Child1_ko.Name
    requiredMap.Add "자녀1 영문이름", Me.txtName_Child1_en.Name
    requiredMap.Add "자녀1 생년월일", Me.txtBirthday_Child1.Name
    requiredMap.Add "자녀2 생명번호", Me.txtLifeNo_Child2.Name
    requiredMap.Add "자녀2 한글이름", Me.txtName_Child2_ko.Name
    requiredMap.Add "자녀2 영문이름", Me.txtName_Child2_en.Name
    requiredMap.Add "자녀2 생년월일", Me.txtBirthday_Child2.Name
    requiredMap.Add "자녀3 생명번호", Me.txtLifeNo_Child3.Name
    requiredMap.Add "자녀3 한글이름", Me.txtName_Child3_ko.Name
    requiredMap.Add "자녀3 영문이름", Me.txtName_Child3_en.Name
    requiredMap.Add "자녀3 생년월일", Me.txtBirthday_Child3.Name
    
    Set GetRequiredMap = requiredMap

End Function

'--//입력된 데이터가 유효하지 않은지 확인
'--//True: 유효하지 않음
'--//False: 유효함
Private Function IsInvalidInput(control As MSForms.control) As Boolean
    
    '--//유효하다는 가정 하에 시작
    IsInvalidInput = False
    
    '--//컨트롤 값이 비어있으면 프로시저 종료
    If control.text = "" Then
        Exit Function
    End If
    
    '--//국적 점검을 위한 객체준비
    Dim CountryList As Object
    Dim CountryDao As New CountryDao
    Set CountryList = CountryDao.GetCountryList
    
    '--//날짜형식 체크
    If control.Name Like "*Birthday*" Or control.Name Like "*_dt*" Then
        If Not IsDate(control.text) Then
            IsInvalidInput = True
        End If
    End If
    
    '--//국적 체크
    If control.Name Like "*Nationality*" Then
        If Not CountryList.Contains(control.text) Then
            IsInvalidInput = True
        End If
    End If
    
    '--//생명번호 형식 체크
    If control.Name Like "*LifeNo*" Then
        If Not IsNumeric(fnExtract(control.text)) Or _
            Mid(control.text, 4, 1) <> "-" Or Mid(control.text, 11, 1) <> "-" Then
            IsInvalidInput = True
        End If
    End If
    
    '--//이름 체크
    If control.Name Like "*Name*" Then
        '--//한글이름 체크
        If control.Name Like "*ko*" Then
            If Len(fnExtract(control.text, "E")) > 0 Then
                IsInvalidInput = True
            End If
        End If
        '--//영문이름 체크
        If control.Name Like "*en*" Then
            If Len(fnExtract(control.text, "H")) > 0 Then
                IsInvalidInput = True
            End If
        End If
    End If
    
    '--//생도기수 확인
    If control.Name = Me.txtTheological_Order.Name And _
        Not IsNumeric(Me.txtTheological_Order) And Me.txtTheological_Order <> "" Then
        IsInvalidInput = True
    End If
    
End Function

'--//필수 입력값이 비어있는지 확인
'--//True: 비어있음
'--//False: 비어있지 않음
Private Function IsEmptyRequiredControl(control As MSForms.control) As Boolean

    '--//필수값 점검을 위한 객체준비
    Dim requiredMap As Collection
    Set requiredMap = GetRequiredMap

    IsEmptyRequiredControl = False
        
    '--//필수값 입력여부 체크
    If ExistsInCollection(requiredMap, control.Name) And control = "" Then
        If control.Name Like "*Spouse*" And Me.chkSpouse.Value = False Then
            Exit Function
        End If
        If control.Name Like "*Child1*" And Me.chkChild1.Value = False Then
            Exit Function
        End If
        If control.Name Like "*Child2*" And Me.chkChild2.Value = False Then
            Exit Function
        End If
        If control.Name Like "*Child3*" And Me.chkChild3.Value = False Then
            Exit Function
        End If
        
        IsEmptyRequiredControl = True
    End If

End Function

'@param frame: Frame 내부 Control들에 대하여 Activate/DeActivate 합니다.
'@param blnActivate: True: Activate | False: Deactivate
Private Sub ActivateFrame(ByRef frame As MSForms.frame, blnActivate As Boolean)
    
    Dim control As MSForms.control
    Dim requiredList As Object
    
    '--//필수값 리스트 불러오기
    Set requiredList = GetRequiredList
    
    For Each control In frame.controls
        Select Case TypeName(control)
            Case "TextBox":
                If InStr(control.Name, "LifeNo") > 0 And blnActivate Then
                    TextBoxEnable control, Not blnActivate
                Else
                    TextBoxEnable control, blnActivate
                End If
            Case "Label":
                If Not control.Name Like "*Info*" Then
                    If blnActivate Then
                        ChangeLabelColor control, vbRed
                    Else
                        ChangeLabelColor control, vbBlack
                    End If
                End If
                If control.Name Like "*Pic*" Then
                    If Not blnActivate Then
                        control.Picture = LoadPicture("")
'                        InsertPicToLabel control, ""
                    End If
                End If
            Case "Frame":
                '--//이거 활성화 시키면 무한루프 빠짐
'                ActivateInputModeForFrame control, blnActivate
        End Select
    Next
    
End Sub

'@param controls: InputMode를 활성화할 controls 객체
'@param blnActivate: 활성화 할지 여부
Private Sub ActivateInputMode(ByRef controls As MSForms.controls, blnActivate As Boolean, INPUTMODE)

    Dim control As MSForms.control
    Dim requiredList As Object
    
    If INPUTMODE = INPUT_FOR_ALL Then
        '--//신규입력 시에는 검색관련 기능 비활성화
        LockControlsOfSearchFunction blnActivate
        
        '--//InputMode를 위한 Command 버튼설정
        HideCmdBtnForInput blnActivate
        
        ExtraControlsEnable blnActivate
    End If
    
    '--//텍스트박스 초기화
    InitializeTextBoxes controls
    
    
    '--//필수값 리스트 불러오기
    Set requiredList = GetRequiredList
    
    For Each control In controls
        Select Case TypeName(control)
            Case "TextBox":
                If Not control.Name Like "*Search*" Then
                    TextBoxEnable control, blnActivate
                End If
            Case "Label":
                '--//Info라벨은 설명문이므로 제외
                If Not control.Name Like "*Info*" Then
                    If blnActivate And _
                        requiredList.Contains(control.Caption) Then
                        ChangeLabelColor control, vbRed
                    Else
                        ChangeLabelColor control, vbBlack
                    End If
                End If
            Case "Frame":
                If Not control.Name Like "*Transfer*" Then
                    ActivateInputMode control.controls, blnActivate, INPUT_FOR_PART
                End If
            Case "CheckBox":
                If (control.Name Like "*Child*" Or control.Name Like "*Spouse*") And _
                    Not control.Name Like "*Transfer*" Then
                    control.Visible = blnActivate
                End If
            Case "CommandButton":
                If control.Name Like "*Delete*" And _
                    (control.Name Like "*Child*" Or control.Name Like "*Spouse*") Then
                    control.Enabled = Not blnActivate
                End If
        End Select
    Next
    
    '--//유저폼 너비조정(자녀가 보이도록)
    cmdExtendChild_Click
    
    '--//체크박스 체크여부에 따른 서식조정
    '--//값이 True인 것은 조정하지 말고 생략
    '--//값이 True인 것까지 하면 무한루프
'    If Not chkSpouse.Value Then chkSpouse_Click
'    If Not chkChild1.Value Then chkChild1_Click
'    If Not chkChild2.Value Then chkChild2_Click
'    If Not chkChild3.Value Then chkChild3_Click
    
    '--//사진 초기화
    For Each control In controls
        If TypeName(control) = "Label" Then
            If IsNumeric(InStr(control, "Pic")) Then
                control.Picture = LoadPicture("")
            End If
        End If
    Next
    
    '--//InputMode False 일 때 폼 초기화
    If Not blnActivate Then
        lstPStaff_Click
    End If

End Sub

'--//신규 입력 시 검색관련 컨트롤을 비활성화 하여 의도치 않은 이벤트가 발생하지 않도록 방지합니다.
'@param blnActivate: 활성화 할지 여부
Sub LockControlsOfSearchFunction(blnActivate As Boolean)
    
    '--//검색과 관련된 컨트롤 비활성화
    Me.txtSearchText.Enabled = Not blnActivate
    Me.lstPStaff.Enabled = Not blnActivate
    Me.chkAll.Enabled = Not blnActivate
    Me.cmdSearch.Enabled = Not blnActivate

End Sub

'--//신규 입력 시 Input Control의 내용을 초기화 합니다.
'@param controls: TextBox를 초기화 하기 위해서 controls 객체를 받습니다.
Sub InitializeTextBoxes(ByRef controls As MSForms.controls)
    
    Dim control As MSForms.control
    
    For Each control In controls
        Select Case TypeName(control)
        Case "TextBox", "ComboBox":
            If Not control.Name Like "*Search*" Then
                control = ""
            End If
        Case "CommandButton":
            
        Case "CheckBox":
            '--//배우자, 자녀1,2,3 체크박스는 Click 메서드를 통해 맞추고
            '--//Transfer만 초기화
            '--//cmdCancel_Click 시에는 lstPstaff_Click을 통해 맞춤
            If control.Name Like "*Transfer*" Then
                control.Value = False
            End If
        Case "Label":
            If control.Name Like "*Pic*" Then
                control.Picture = LoadPicture("")
            End If
        Case "Frame":
            InitializeTextBoxes control.controls
        End Select
    Next
    
End Sub

'--//DELETE_ITEM 권한이 있는 사용자만 삭제버튼이 보이게 합니다.
Private Sub HideDeleteButtonByUserAuth()
    
    Dim authList As Object
    Dim objUserDao As New UserDao
    
    Set authList = objUserDao.GetUserAuthorities
    
    Dim strAuth As Variant
    If authList.Contains("DELETE_ITEM") Then
        Me.cmdDelete.Visible = True
        Me.cmdDelete_Child1.Visible = True
        Me.cmdDelete_Child2.Visible = True
        Me.cmdDelete_Child3.Visible = True
        Me.cmdDelete_Spouse.Visible = True
    Else
        Me.cmdDelete.Visible = False
        Me.cmdDelete_Child1.Visible = False
        Me.cmdDelete_Child2.Visible = False
        Me.cmdDelete_Child3.Visible = False
        Me.cmdDelete_Spouse.Visible = False
    End If

End Sub

'--//신규 입력 시 보여야 할 버튼과 보이지 않아야 할 버튼을 구분합니다.
Private Sub HideCmdBtnForInput(argBoolean As Boolean)
    Me.cmdNew.Visible = Not argBoolean
    Me.cmdEdit.Visible = Not argBoolean
    Me.cmdDelete.Visible = Not argBoolean
    Me.cmdCancel.Visible = argBoolean
    Me.cmdAdd.Visible = argBoolean
    
    '--//삭제버튼은 권한에 따라 다시설정
    Call HideDeleteButtonByUserAuth
End Sub

'--//부가적인 기능을 하는 Control들에 대해 활성화/비활성화 여부 조정
'@param blnActivate: 활성화/비활성화 여부
Private Sub ExtraControlsEnable(ByVal blnActivate As Boolean)

    '--//Add 버튼이 활성화 여부에 따라
    Me.cmdAdd.Enabled = blnActivate
    If Me.lstPStaff.listIndex = -1 Then
        Me.cmdDelete.Enabled = Not blnActivate
        Me.cmdEdit.Enabled = Not blnActivate
    Else
        Me.cmdDelete.Enabled = blnActivate
        Me.cmdEdit.Enabled = blnActivate
    End If
    
    Dim control As MSForms.control
    
    '--//폼 내 전체 컨트로들에 대하여
    For Each control In Me.controls
        Select Case TypeName(control)
            Case "CheckBox":
                If Not (control.Name Like "*chkAll*" Or control.Name Like "*Transfer*") Then
                    ControlEnable control, blnActivate
                End If
            Case "TextBox":
                If control.Name <> "txtSearchText" And _
                    Not control.Name Like "*txtFamily*" Then
                    TextBoxEnable control, blnActivate
                End If
            Case "ComboBox":
                ComboBoxEnable control, blnActivate
            Case "Label":
                If control.Name Like "*Pic*" Then
                    ControlEnable control, blnActivate
                End If
            Case "CommandButton":
                If control.Caption Like "*구성원*" Then
                    ControlEnable control, Not blnActivate
                End If
                If control.Name Like "*Transfer*" Then
                    ControlEnable control, Not blnActivate
                End If
        End Select
    Next

End Sub

'--//Deprecated(2022-11-07)
'--//사진을 DB에 저장하는 것으로 로직이 변경되었으므로 더 이상 사용하지 않습니다.
Sub sbCopyPic(ByVal control As MSForms.control)

    Dim filePath As String
    Dim FileName As Variant
    Dim Class As String '--//대상구분
    Dim strExtension As String
    
    '--//대상구분에 따른 변수설정
    Select Case Replace(control.Name, "txtLifeNo_", "")
        Case "Spouse"
            Class = "배우자"
        Case "Child1"
            Class = "자녀1"
        Case "Child2"
            Class = "자녀2"
        Case "Child3"
            Class = "자녀3"
    Case Else
        Class = "선지자"
    End Select
    
    '--//유효성 체크
    If fnFindPicPath = False Then
        MsgBox "마이디스크를 먼저 연결해 주세요.", vbCritical, banner
        Exit Sub
    End If
    If control = "" Then
        MsgBox Class & " 생명번호를 먼저 입력해 주세요.", vbCritical, banner
        Exit Sub
    End If
    
    '--//사진파일경로 불러오기
    filePath = fnFindPicPath
    
    '--//사진선택
    If FileName = False Then
        FileName = Application.GetOpenFilename(fileFilter:="Picturefile, *.jpg;*.jpeg; *.bmp; *.tif; *.png")
    End If
    
    '--//사진 미선택 시 프로시저 종료
    If FileName = False Then
        Exit Sub
    End If
    
    '--//선택한 파일 확장자 jpg,png가 아닐 시 프로시저 종료
    If IsInArray(Right(FileName, Len(FileName) - InStrRev(FileName, ".")), Array("jpg"), , rtnSequence) = -1 Then
        MsgBox "사진은 JPG 확장자만 선택할 수 있습니다."
        Exit Sub
    End If
    
    strExtension = "." & Right(FileName, Len(FileName) - InStrRev(FileName, "."))
    
    '--//선택한 파일 명을 생면번호로 바꾼 이후 사진파일 경로로 이동복사
    If Dir(filePath & control & strExtension) <> "" Then
        If MsgBox("파일이 이미 존재 합니다. 변경 하시겠습니까?", vbQuestion + vbYesNo, banner) = vbYes Then
            Kill filePath & control & strExtension
            Name FileName As filePath & control & strExtension
        End If
    Else
        Name FileName As filePath & control & strExtension
    End If
    
    Call lstPStaff_Click

End Sub

'--//PStaff, PWife 객체 파싱을 위해서 공란으로 되어 있는 컨트롤에 기본값을 채워 넣습니다.
Public Sub FillWithDefaultValue()
    If Me.txtBirthday = "" Then Me.txtBirthday = DateSerial(1900, 1, 1)
    If Me.txtBirthday_Spouse = "" Then Me.txtBirthday_Spouse = DateSerial(1900, 1, 1)
    If Me.txtBirthday_Child1 = "" Then Me.txtBirthday_Child1 = DateSerial(1900, 1, 1)
    If Me.txtBirthday_Child2 = "" Then Me.txtBirthday_Child2 = DateSerial(1900, 1, 1)
    If Me.txtBirthday_Child3 = "" Then Me.txtBirthday_Child3 = DateSerial(1900, 1, 1)
    If Me.txtOvs_dt = "" Then Me.txtOvs_dt = DateSerial(1900, 1, 1)
    If Me.txtWedding_dt = "" Then Me.txtWedding_dt = DateSerial(1900, 1, 1)
    If Me.txtOrdinationPrayer_dt = "" Then Me.txtOrdinationPrayer_dt = DateSerial(1900, 1, 1)
    
    If Me.txtTheological_Order = "" Then Me.txtTheological_Order = 0
    If Me.txtSalary = "" Then Me.txtSalary = 0
    
End Sub

'--//pStaff, pWife 객체를 받아 폼에 내용을 채워 넣습니다.
Private Sub FillOutForm(ByRef pStaff As PastoralStaff, ByRef pWife As PastoralWife)

    '--//관리자
    Me.txtLifeNo = pStaff.lifeNo
    Me.txtName_ko = pStaff.nameKo
    Me.txtName_en = pStaff.NameEn
    Me.txtBirthday = pStaff.Birthday
    Me.txtPhone = pStaff.Phone
    Me.txtNationality = pStaff.Nationality
    Me.txtEducation = pStaff.Education
    Me.txtHome = pStaff.Home
    Me.txtFamily = pStaff.Family
    Me.txtHealth = pStaff.Health
    Me.txtOther = pStaff.Other
    Me.txtOvs_dt = IIf(pStaff.AppoOvs = DateSerial(1900, 1, 1), "", pStaff.AppoOvs)
    Me.cboBaptism = pStaff.Baptism
    Me.txtOrdinationPrayer_dt = IIf(pStaff.OrdinationPrayer = DateSerial(1900, 1, 1), "", pStaff.OrdinationPrayer)
    Me.txtWedding_dt = IIf(pStaff.WeddingDt = DateSerial(1900, 1, 1), "", pStaff.WeddingDt)
    Me.txtSalary = pStaff.Salary
    Me.txtTheological_Order = IIf(pStaff.TheologicalOrder = 0, "", pStaff.TheologicalOrder)
    Me.chkTransfer.Value = pStaff.Suspend
    
    '--//자녀1
    Me.txtLifeNo_Child1 = pStaff.LifeNoChild1
    Me.txtName_Child1_ko = pStaff.NameKoChild1
    Me.txtName_Child1_en = pStaff.NameEnChild1
    Me.txtBirthday_Child1 = IIf(pStaff.BirthdayChild1 = DateSerial(1900, 1, 1), "", pStaff.BirthdayChild1)
    Me.txtPhone_Child1 = pStaff.PhoneChild1
    If pStaff.LifeNoChild1 <> "" Then
        Me.chkChild1.Value = True
    Else
        Me.chkChild1.Value = False
    End If
    
    '--//자녀2
    Me.txtLifeNo_Child2 = pStaff.LifeNoChild2
    Me.txtName_Child2_ko = pStaff.NameKoChild2
    Me.txtName_Child2_en = pStaff.NameEnChild2
    Me.txtBirthday_Child2 = IIf(pStaff.BirthdayChild2 = DateSerial(1900, 1, 1), "", pStaff.BirthdayChild2)
    Me.txtPhone_Child2 = pStaff.PhoneChild2
    If pStaff.LifeNoChild2 <> "" Then
        Me.chkChild2.Value = True
    Else
        Me.chkChild2.Value = False
    End If
    
    '--//자녀3
    Me.txtLifeNo_Child3 = pStaff.LifeNoChild3
    Me.txtName_Child3_ko = pStaff.NameKoChild3
    Me.txtName_Child3_en = pStaff.NameEnChild3
    Me.txtBirthday_Child3 = IIf(pStaff.BirthdayChild3 = DateSerial(1900, 1, 1), "", pStaff.BirthdayChild3)
    Me.txtPhone_Child3 = pStaff.PhoneChild3
    If pStaff.LifeNoChild3 <> "" Then
        Me.chkChild3.Value = True
    Else
        Me.chkChild3.Value = False
    End If
    
    '--//배우자
    Me.txtLifeNo_Spouse = pWife.lifeNo
    Me.txtName_Spouse_ko = pWife.nameKo
    Me.txtName_Spouse_en = pWife.NameEn
    Me.txtBirthday_Spouse = IIf(pWife.Birthday = DateSerial(1900, 1, 1), "", pWife.Birthday)
    Me.txtPhone_Spouse = pWife.Phone
    Me.txtNationality_Spouse = pWife.Nationality
    Me.txtEducation_Spouse = pWife.Education
    Me.txtHome_Spouse = pWife.Home
    Me.txtFamily_Spouse = pWife.Family
    Me.txtHealth_Spouse = pWife.Health
    Me.txtOther_Spouse = pWife.Other
    If pWife.lifeNo <> "" Then
        Me.chkSpouse.Value = True
    Else
        Me.chkSpouse.Value = False
    End If
    Me.chkTransferSpouse.Value = pWife.Suspend

End Sub
