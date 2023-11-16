VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PastoralWife"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'###################################
'TABLE: DB_PASTORALWIFE에 대한 CLASS
'###################################

Dim vLifeNo As String
Dim vNationality As String
Dim vNameKo As String
Dim vNameEn As String
Dim vBirthDay As Date
Dim vPhone As String
Dim vHome As String
Dim vFamily As String
Dim vHealth As String
Dim vOther As String
Dim vLifeNoSpouse As String
Dim vEducation As String
Dim vSuspend As Integer
Dim vOvsDept As Integer

Private Sub class_initialize()
    vBirthDay = DateSerial(1900, 1, 1)
End Sub

Public Property Get lifeNo() As String
    lifeNo = vLifeNo
End Property

Public Property Let lifeNo(argLifeNo As String)
    vLifeNo = argLifeNo
End Property

Public Property Get Nationality() As String
    Nationality = vNationality
End Property

Public Property Let Nationality(argNationality As String)
    vNationality = argNationality
End Property

Public Property Get nameKo() As String
    nameKo = vNameKo
End Property

Public Property Let nameKo(argNameKo As String)
    vNameKo = argNameKo
End Property

Public Property Get NameEn() As String
    NameEn = vNameEn
End Property

Public Property Let NameEn(argNameEn As String)
    vNameEn = argNameEn
End Property

Public Property Get Birthday() As Date
    Birthday = vBirthDay
End Property

Public Property Let Birthday(argBirthday As Date)
    vBirthDay = argBirthday
End Property

Public Property Get Phone() As String
    Phone = vPhone
End Property

Public Property Let Phone(argPhone As String)
    vPhone = argPhone
End Property

Public Property Get Home() As String
    Home = vHome
End Property

Public Property Let Home(argHome As String)
    vHome = argHome
End Property

Public Property Get Family() As String
    Family = vFamily
End Property

Public Property Let Family(argFamily As String)
    vFamily = argFamily
End Property

Public Property Get Health() As String
    Health = vHealth
End Property

Public Property Let Health(argHealth As String)
    vHealth = argHealth
End Property

Public Property Get Other() As String
    Other = vOther
End Property

Public Property Let Other(argOther As String)
    vOther = argOther
End Property

Public Property Get lifeNoSpouse() As String
    lifeNoSpouse = vLifeNoSpouse
End Property

Public Property Let lifeNoSpouse(argLifeNoSpouse As String)
    vLifeNoSpouse = argLifeNoSpouse
End Property

Public Property Get Education() As String
    Education = vEducation
End Property

Public Property Let Education(argEducation As String)
    vEducation = argEducation
End Property

Public Property Get Suspend() As Boolean
    If vSuspend = 0 Then
        Suspend = False
    Else
        Suspend = True
    End If
End Property

Public Property Let Suspend(argSuspend As Boolean)
    If argSuspend Then
        vSuspend = 1
    Else
        vSuspend = 0
    End If
End Property

Public Property Get OvsDept() As String
    OvsDept = vOvsDept
End Property

Public Property Let OvsDept(argOvsDept As String)
    vOvsDept = argOvsDept
End Property

'@param form: 선지자정보 폼을 변수로 받아 객체 속성에 값을 채웁니다.
Public Sub ParseFromForm(ByRef form As frm_Update_PInformation)
    Call frm_Update_PInformation.FillWithDefaultValue
    vLifeNo = form.txtLifeNo_Spouse
    vNationality = form.txtNationality_Spouse
    vNameKo = form.txtName_Spouse_ko
    vNameEn = form.txtName_Spouse_en
    vBirthDay = form.txtBirthday_Spouse
    vPhone = form.txtPhone_Spouse
    vHome = form.txtHome_Spouse
    vFamily = form.txtFamily_Spouse
    vHealth = form.txtHealth_Spouse
    vOther = form.txtOther_Spouse
    vLifeNoSpouse = form.txtLifeNo
    vEducation = form.txtEducation_Spouse
    vSuspend = IIf(form.chkTransferSpouse, 1, 0)
    vOvsDept = USER_DEPT
End Sub

'@param tRecordSet: DB에서 받은 RecordSet 타입을 변수로 받아 객체 속성에 값을 채웁니다.
Public Sub ParseFromRecordSet(ByRef tRecordSet As T_RECORD_SET, Optional index As Long = 0)
    With tRecordSet
        vLifeNo = .LISTDATA(index, 0)
        vNationality = .LISTDATA(index, 1)
        vNameKo = .LISTDATA(index, 2)
        vNameEn = .LISTDATA(index, 3)
        If .LISTDATA(index, 4) <> "" Then
            vBirthDay = .LISTDATA(index, 4)
        Else
            vBirthDay = DateSerial(1900, 1, 1)
        End If
        vPhone = .LISTDATA(index, 5)
        vHome = .LISTDATA(index, 6)
        vFamily = .LISTDATA(index, 7)
        vHealth = .LISTDATA(index, 8)
        vOther = .LISTDATA(index, 9)
        vLifeNoSpouse = .LISTDATA(index, 10)
        vEducation = .LISTDATA(index, 11)
        vSuspend = .LISTDATA(index, 12)
        vOvsDept = USER_DEPT
    End With
End Sub

'@param argPWife: argPWife 객체를 변수로 받아 생명번호 기준으로 동일한지 여부를 점검합니다.
'@param blnStrictly: 변수 값이 True는 강한 검증을 말하며, 생명번호를 제외한 나머지 정보까지 모두 같을 때 True를 반환합니다.
'                    선택사항이며 기본값은 False 입니다.
Public Function IsEqual(ByRef argPWife As PastoralWife, Optional blnStrictly As Boolean = False) As Boolean
            
    If Me.lifeNo <> argPWife.lifeNo Then IsEqual = False: Exit Function
    
    If blnStrictly And argPWife.lifeNo <> "" Then
        If Me.Nationality <> argPWife.Nationality Then IsEqual = False: Exit Function
        If Me.nameKo <> argPWife.nameKo Then IsEqual = False: Exit Function
        If Me.NameEn <> argPWife.NameEn Then IsEqual = False: Exit Function
        If Me.Birthday <> argPWife.Birthday Then IsEqual = False: Exit Function
        If Me.Phone <> argPWife.Phone Then IsEqual = False: Exit Function
        If Me.Home <> argPWife.Home Then IsEqual = False: Exit Function
        If Me.Family <> argPWife.Family Then IsEqual = False: Exit Function
        If Me.Health <> argPWife.Health Then IsEqual = False: Exit Function
        If Me.Other <> argPWife.Other Then IsEqual = False: Exit Function
        If Me.lifeNoSpouse <> argPWife.lifeNoSpouse Then IsEqual = False: Exit Function
        If Me.Education <> argPWife.Education Then IsEqual = False: Exit Function
        If Me.Suspend <> argPWife.Suspend Then IsEqual = False: Exit Function
        If Me.OvsDept <> argPWife.OvsDept Then IsEqual = False: Exit Function
    End If
    IsEqual = True
End Function
