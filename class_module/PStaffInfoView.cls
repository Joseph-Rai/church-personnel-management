VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PStaffInfoView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'#########################################
'TABLE: V0_PSTAFF_INFORMATION에 대한 CLASS
'#########################################

Dim vChurchCode As String
Dim vChurchNameKo As String
Dim vChurchNameEn As String
Dim vBranchNameKo As String
Dim vBranchNameEn As String
Dim vMissionCountry As String
Dim vLifeNo As String
Dim vNameKoAndTitle As String
Dim vNameEn As String
Dim vPosition As String
Dim vPosition2 As String
Dim vBirthDay As Date
Dim vNationality As String
Dim vHome As String
Dim vHomeSpouse As String
Dim vAppoOvs As Date
Dim vAppoCur As Date
Dim vLifeNoSpouse As String
Dim vNameKoAndTitleSpouse As String
Dim vNameEnSpouse As String
Dim vPositionSpouse As String
Dim vBirthdaySpouse As Date
Dim vTheologicalStep As String
Dim vOvsDept As Integer
Dim vNationalitySpouse As String
Dim vNameKo As String
Dim vNameKoSpouse As String
Dim vEducation As String
Dim vEducationSpouse As String
Dim vTitle As String
Dim vTitleSpouse As String
Dim vBranchCode As String

Public Property Get churchCode() As String
    churchCode = vChurchCode
End Property

Public Property Let churchCode(argChurchCode As String)
    vChurchCode = argChurchCode
End Property

Public Property Get ChurchNameKo() As String
    ChurchNameKo = vChurchNameKo
End Property

Public Property Let ChurchNameKo(argChurchNameKo As String)
    vChurchNameKo = argChurchNameKo
End Property

Public Property Get ChurchNameEn() As String
    ChurchNameEn = vChurchNameEn
End Property

Public Property Let ChurchNameEn(argChurchNameEn As String)
    vChurchNameEn = argChurchNameEn
End Property

Public Property Get BranchNameKo() As String
    BranchNameKo = vBranchNameKo
End Property

Public Property Let BranchNameKo(argBranchNameKo As String)
    vBranchNameKo = argBranchNameKo
End Property

Public Property Get BranchNameEn() As String
    BranchNameEn = vBranchNameEn
End Property

Public Property Let BranchNameEn(argBranchNameEn As String)
    vBranchNameEn = argBranchNameEn
End Property

Public Property Get MissionCountry() As String
    MissionCountry = vMissionCountry
End Property

Public Property Let MissionCountry(argMissionCountry As String)
    vMissionCountry = argMissionCountry
End Property

Public Property Get lifeNo() As String
    lifeNo = vLifeNo
End Property

Public Property Let lifeNo(argLifeNo As String)
    vLifeNo = argLifeNo
End Property

Public Property Get NameKoAndTitle() As String
    NameKoAndTitle = vNameKoAndTitle
End Property

Public Property Let NameKoAndTitle(argNameKoAndTitle As String)
    vNameKoAndTitle = argNameKoAndTitle
End Property

Public Property Get NameEn() As String
    NameEn = vNameEn
End Property

Public Property Let NameEn(argNameEn As String)
    vNameEn = argNameEn
End Property

Public Property Get position() As String
    position = vPosition
End Property

Public Property Let position(argPosition As String)
    vPosition = argPosition
End Property

Public Property Get Position2() As String
    Position2 = vPosition2
End Property

Public Property Let Position2(argPosition2 As String)
    vPosition2 = argPosition2
End Property

Public Property Get Birthday() As Date
    Birthday = vBirthDay
End Property

Public Property Let Birthday(argBirthday As Date)
    vBirthDay = argBirthday
End Property

Public Property Get Nationality() As String
    Nationality = vNationality
End Property

Public Property Let Nationality(argNationality As String)
    vNationality = argNationality
End Property

Public Property Get Home() As String
    Home = vHome
End Property

Public Property Let Home(argHome As String)
    vHome = argHome
End Property

Public Property Get HomeSpouse() As String
    HomeSpouse = vHomeSpouse
End Property

Public Property Let HomeSpouse(argHomeSpouse As String)
    vHomeSpouse = argHomeSpouse
End Property

Public Property Get AppoOvs() As Date
    AppoOvs = vAppoOvs
End Property

Public Property Let AppoOvs(argAppoOvs As Date)
    vAppoOvs = argAppoOvs
End Property

Public Property Get AppoCur() As Date
    AppoCur = vAppoCur
End Property

Public Property Let AppoCur(argAppoCur As Date)
    vAppoCur = argAppoCur
End Property

Public Property Get lifeNoSpouse() As String
    lifeNoSpouse = vLifeNoSpouse
End Property

Public Property Let lifeNoSpouse(argLifeNoSpouse As String)
    vLifeNoSpouse = argLifeNoSpouse
End Property

Public Property Get NameKoAndTitleSpouse() As String
    NameKoAndTitleSpouse = vNameKoAndTitleSpouse
End Property

Public Property Let NameKoAndTitleSpouse(argNameKoAndTitleSpouse As String)
    vNameKoAndTitleSpouse = argNameKoAndTitleSpouse
End Property

Public Property Get NameEnSpouse() As String
    NameEnSpouse = vNameEnSpouse
End Property

Public Property Let NameEnSpouse(argNameEnSpouse As String)
    vNameEnSpouse = argNameEnSpouse
End Property

Public Property Get PositionSpouse() As String
    PositionSpouse = vPositionSpouse
End Property

Public Property Let PositionSpouse(argPositionSpouse As String)
    vPositionSpouse = argPositionSpouse
End Property

Public Property Get BirthdaySpouse() As Date
    BirthdaySpouse = vBirthdaySpouse
End Property

Public Property Let BirthdaySpouse(argBirthdaySpouse As Date)
    vBirthdaySpouse = argBirthdaySpouse
End Property

Public Property Get TheologicalStep() As String
    TheologicalStep = vTheologicalStep
End Property

Public Property Let TheologicalStep(argTheologicalStep As String)
    vTheologicalStep = argTheologicalStep
End Property

Public Property Get OvsDept() As String
    OvsDept = vOvsDept
End Property

Public Property Let OvsDept(argOvsDept As String)
    vOvsDept = argOvsDept
End Property

Public Property Get NationalitySpouse() As String
    NationalitySpouse = vNationalitySpouse
End Property

Public Property Let NationalitySpouse(argNationalitySpouse As String)
    vNationalitySpouse = argNationalitySpouse
End Property

Public Property Get nameKo() As String
    nameKo = vNameKo
End Property

Public Property Let nameKo(argNameKo As String)
    vNameKo = argNameKo
End Property

Public Property Get NameKoSpouse() As String
    NameKoSpouse = vNameKoSpouse
End Property

Public Property Let NameKoSpouse(argNameKoSpouse As String)
    vNameKoSpouse = argNameKoSpouse
End Property

Public Property Get Education() As String
    Education = vEducation
End Property

Public Property Let Education(argEducation As String)
    vEducation = argEducation
End Property

Public Property Get EducationSpouse() As String
    EducationSpouse = vEducationSpouse
End Property

Public Property Let EducationSpouse(argEducationSpouse As String)
    vEducationSpouse = argEducationSpouse
End Property

Public Property Get title() As String
    title = vTitle
End Property

Public Property Let title(argTitle As String)
    vTitle = argTitle
End Property

Public Property Get TitleSpouse() As String
    TitleSpouse = vTitleSpouse
End Property

Public Property Let TitleSpouse(argTitleSpouse As String)
    vTitleSpouse = argTitleSpouse
End Property

Public Property Get BranchCode() As String
    BranchCode = vBranchCode
End Property

Public Property Let BranchCode(argBranchCode As String)
    vBranchCode = argBranchCode
End Property

'@param tRecordSet: DB에서 받은 RecordSet 타입을 변수로 받아 객체 속성에 값을 채웁니다.
Public Sub ParseFromRecordSet(ByRef tRecordSet As T_RECORD_SET, Optional index As Long = 0)
    With tRecordSet
        vChurchCode = .LISTDATA(index, 0)
        vChurchNameKo = .LISTDATA(index, 1)
        vChurchNameEn = .LISTDATA(index, 2)
        vBranchNameKo = .LISTDATA(index, 3)
        vBranchNameEn = .LISTDATA(index, 4)
        vMissionCountry = .LISTDATA(index, 5)
        vLifeNo = .LISTDATA(index, 6)
        vNameKoAndTitle = .LISTDATA(index, 7)
        vNameEn = .LISTDATA(index, 8)
        vPosition = .LISTDATA(index, 9)
        vPosition2 = .LISTDATA(index, 10)
        If .LISTDATA(index, 11) <> "" Then
            vBirthDay = .LISTDATA(index, 11)
        Else
            vBirthDay = DateSerial(1900, 1, 1)
        End If
        vNationality = .LISTDATA(index, 12)
        vHome = .LISTDATA(index, 13)
        vHomeSpouse = .LISTDATA(index, 14)
        If .LISTDATA(index, 15) <> "" Then
            vAppoOvs = .LISTDATA(index, 15)
        Else
            vAppoOvs = DateSerial(1900, 1, 1)
        End If
        If .LISTDATA(index, 16) <> "" Then
            vAppoCur = .LISTDATA(index, 16)
        Else
            vAppoCur = DateSerial(1900, 1, 1)
        End If
        vLifeNoSpouse = .LISTDATA(index, 17)
        vNameKoAndTitleSpouse = .LISTDATA(index, 18)
        vNameEnSpouse = .LISTDATA(index, 19)
        vPositionSpouse = .LISTDATA(index, 20)
        If .LISTDATA(index, 21) <> "" Then
            vBirthdaySpouse = .LISTDATA(index, 21)
        Else
            vBirthdaySpouse = DateSerial(1900, 1, 1)
        End If
        vTheologicalStep = .LISTDATA(index, 22)
        vOvsDept = .LISTDATA(index, 23)
        vNationalitySpouse = .LISTDATA(index, 24)
        vNameKo = .LISTDATA(index, 25)
        vNameKoSpouse = .LISTDATA(index, 26)
        vEducation = .LISTDATA(index, 27)
        vEducationSpouse = .LISTDATA(index, 28)
        vTitle = .LISTDATA(index, 29)
        vTitleSpouse = .LISTDATA(index, 30)
        vBranchCode = .LISTDATA(index, 31)
    End With
End Sub
