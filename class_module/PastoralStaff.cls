VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PastoralStaff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'####################################
'TABLE: DB_PASTORALSTAFF에 대한 CLASS
'####################################

Dim vLifeNo As String
Dim vNameKo As String
Dim vNameEn As String
Dim vNationality As String
Dim vBirthDay As Date
Dim vPhone As String
Dim vLifenoChild1 As String
Dim vNameKoChild1 As String
Dim vNameEnChild1 As String
Dim vBirthdayChild1 As Date
Dim vPhoneChild1 As String
Dim vLifenoChild2 As String
Dim vNameKoChild2 As String
Dim vNameEnChild2 As String
Dim vBirthdayChild2 As Date
Dim vPhoneChild2 As String
Dim vLifenoChild3 As String
Dim vNameKoChild3 As String
Dim vNameEnChild3 As String
Dim vBirthdayChild3 As Date
Dim vPhoneChild3 As String
Dim vHome As String
Dim vFamily As String
Dim vHealth As String
Dim vOther As String
Dim vBaptism As String
Dim vOrdinationPrayer As Date
Dim vAppoOvs As Date
Dim vWeddingDt As Date
Dim vTheologicalOrder As Integer
Dim vEducation As String
Dim vSalary As Long
Dim vSuspend As Integer
Dim vOvsDept As Integer

Public Property Get lifeNo() As String
    lifeNo = vLifeNo
End Property

Public Property Let lifeNo(argLifeNo As String)
    vLifeNo = argLifeNo
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

Public Property Get Nationality() As String
    Nationality = vNationality
End Property

Public Property Let Nationality(argNationality As String)
    vNationality = argNationality
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

Public Property Get LifeNoChild1() As String
    LifeNoChild1 = vLifenoChild1
End Property

Public Property Let LifeNoChild1(argLifeNoChild1 As String)
    vLifenoChild1 = argLifeNoChild1
End Property

Public Property Get NameKoChild1() As String
    NameKoChild1 = vNameKoChild1
End Property

Public Property Let NameKoChild1(argNameKoChild1 As String)
    vNameKoChild1 = argNameKoChild1
End Property

Public Property Get NameEnChild1() As String
    NameEnChild1 = vNameEnChild1
End Property

Public Property Let NameEnChild1(argNameEnChild1 As String)
    vNameEnChild1 = argNameEnChild1
End Property

Public Property Get BirthdayChild1() As Date
    BirthdayChild1 = vBirthdayChild1
End Property

Public Property Let BirthdayChild1(argBirthdayChild1 As Date)
    vBirthdayChild1 = argBirthdayChild1
End Property

Public Property Get PhoneChild1() As String
    PhoneChild1 = vPhoneChild1
End Property

Public Property Let PhoneChild1(argPhoneChild1 As String)
    vPhoneChild1 = argPhoneChild1
End Property

Public Property Get LifeNoChild2() As String
    LifeNoChild2 = vLifenoChild2
End Property

Public Property Let LifeNoChild2(argLifeNoChild2 As String)
    vLifenoChild2 = argLifeNoChild2
End Property

Public Property Get NameKoChild2() As String
    NameKoChild2 = vNameKoChild2
End Property

Public Property Let NameKoChild2(argNameKoChild2 As String)
    vNameKoChild2 = argNameKoChild2
End Property

Public Property Get NameEnChild2() As String
    NameEnChild2 = vNameEnChild2
End Property

Public Property Let NameEnChild2(argNameEnChild2 As String)
    vNameEnChild2 = argNameEnChild2
End Property

Public Property Get BirthdayChild2() As Date
    BirthdayChild2 = vBirthdayChild2
End Property

Public Property Let BirthdayChild2(argBirthdayChild2 As Date)
    vBirthdayChild2 = argBirthdayChild2
End Property

Public Property Get PhoneChild2() As String
    PhoneChild2 = vPhoneChild2
End Property

Public Property Let PhoneChild2(argPhoneChild2 As String)
    vPhoneChild2 = argPhoneChild2
End Property

Public Property Get LifeNoChild3() As String
    LifeNoChild3 = vLifenoChild3
End Property

Public Property Let LifeNoChild3(argLifeNoChild3 As String)
    vLifenoChild3 = argLifeNoChild3
End Property

Public Property Get NameKoChild3() As String
    NameKoChild3 = vNameKoChild3
End Property

Public Property Let NameKoChild3(argNameKoChild3 As String)
    vNameKoChild3 = argNameKoChild3
End Property

Public Property Get NameEnChild3() As String
    NameEnChild3 = vNameEnChild3
End Property

Public Property Let NameEnChild3(argNameEnChild3 As String)
    vNameEnChild3 = argNameEnChild3
End Property

Public Property Get BirthdayChild3() As Date
    BirthdayChild3 = vBirthdayChild3
End Property

Public Property Let BirthdayChild3(argBirthdayChild3 As Date)
    vBirthdayChild3 = argBirthdayChild3
End Property

Public Property Get PhoneChild3() As String
    PhoneChild3 = vPhoneChild3
End Property

Public Property Let PhoneChild3(argPhoneChild3 As String)
    vPhoneChild3 = argPhoneChild3
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

Public Property Get Baptism() As String
    Baptism = vBaptism
End Property

Public Property Let Baptism(argBaptism As String)
    vBaptism = argBaptism
End Property

Public Property Get OrdinationPrayer() As Date
    OrdinationPrayer = vOrdinationPrayer
End Property

Public Property Let OrdinationPrayer(argOrdinationPrayer As Date)
    vOrdinationPrayer = argOrdinationPrayer
End Property

Public Property Get AppoOvs() As Date
    AppoOvs = vAppoOvs
End Property

Public Property Let AppoOvs(argAppoOvs As Date)
    vAppoOvs = argAppoOvs
End Property

Public Property Get WeddingDt() As Date
    WeddingDt = vWeddingDt
End Property

Public Property Let WeddingDt(argWeddingDt As Date)
    vWeddingDt = argWeddingDt
End Property

Public Property Get TheologicalOrder() As Integer
    TheologicalOrder = vTheologicalOrder
End Property

Public Property Let TheologicalOrder(argTheologicalOrder As Integer)
    vTheologicalOrder = argTheologicalOrder
End Property

Public Property Get Education() As String
    Education = vEducation
End Property

Public Property Let Education(argEducation As String)
    vEducation = argEducation
End Property

Public Property Get Salary() As Long
    Salary = vSalary
End Property

Public Property Let Salary(argSalary As Long)
    vSalary = argSalary
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

Public Property Get OvsDept() As Integer
    OvsDept = vOvsDept
End Property

Public Property Let OvsDept(argOvsDept As Integer)
    vOvsDept = argOvsDept
End Property

'@param form: 선지자정보 폼을 변수로 받아 객체 속성에 값을 채웁니다.
Public Sub ParseFromForm(ByRef form As frm_Update_PInformation)
    Call frm_Update_PInformation.FillWithDefaultValue
    vLifeNo = form.txtLifeNo
    vNameKo = form.txtName_ko
    vNameEn = form.txtName_en
    vNationality = form.txtNationality
    vBirthDay = form.txtBirthday
    vPhone = form.txtPhone
    vLifenoChild1 = form.txtLifeNo_Child1
    vNameKoChild1 = form.txtName_Child1_ko
    vNameEnChild1 = form.txtName_Child1_en
    vBirthdayChild1 = form.txtBirthday_Child1
    vPhoneChild1 = form.txtPhone_Child1
    vLifenoChild2 = form.txtLifeNo_Child2
    vNameKoChild2 = form.txtName_Child2_ko
    vNameEnChild2 = form.txtName_Child2_en
    vBirthdayChild2 = form.txtBirthday_Child2
    vPhoneChild2 = form.txtPhone_Child2
    vLifenoChild3 = form.txtLifeNo_Child3
    vNameKoChild3 = form.txtName_Child3_ko
    vNameEnChild3 = form.txtName_Child3_en
    vBirthdayChild3 = form.txtBirthday_Child3
    vPhoneChild3 = form.txtPhone_Child3
    vHome = form.txtHome
    vFamily = form.txtFamily
    vHealth = form.txtHealth
    vOther = form.txtOther
    vBaptism = form.cboBaptism
    vOrdinationPrayer = form.txtOrdinationPrayer_dt
    vAppoOvs = form.txtOvs_dt
    vWeddingDt = form.txtWedding_dt
    vTheologicalOrder = form.txtTheological_Order
    vEducation = form.txtEducation
    vSalary = form.txtSalary
    vSuspend = IIf(form.chkTransfer, 1, 0)
    vOvsDept = USER_DEPT
End Sub

'@param tRecordSet: DB에서 받은 RecordSet 타입을 변수로 받아 객체 속성에 값을 채웁니다.
Public Sub ParseFromRecordSet(ByRef tRecordSet As T_RECORD_SET, Optional index As Long = 0)
    With tRecordSet
        vLifeNo = .LISTDATA(index, 0)
        vNameKo = .LISTDATA(index, 1)
        vNameEn = .LISTDATA(index, 2)
        vNationality = .LISTDATA(index, 3)
        If .LISTDATA(index, 4) <> "" Then
            vBirthDay = .LISTDATA(index, 4)
        End If
        vPhone = .LISTDATA(index, 5)
        vLifenoChild1 = .LISTDATA(index, 6)
        vNameKoChild1 = .LISTDATA(index, 7)
        vNameEnChild1 = .LISTDATA(index, 8)
        If .LISTDATA(index, 9) <> "" Then
            vBirthdayChild1 = .LISTDATA(index, 9)
        Else
            vBirthdayChild1 = DateSerial(1900, 1, 1)
        End If
        vPhoneChild1 = .LISTDATA(index, 10)
        vLifenoChild2 = .LISTDATA(index, 11)
        vNameKoChild2 = .LISTDATA(index, 12)
        vNameEnChild2 = .LISTDATA(index, 13)
        If .LISTDATA(index, 14) <> "" Then
            vBirthdayChild2 = .LISTDATA(index, 14)
        Else
            vBirthdayChild2 = DateSerial(1900, 1, 1)
        End If
        vPhoneChild2 = .LISTDATA(index, 15)
        vLifenoChild3 = .LISTDATA(index, 16)
        vNameKoChild3 = .LISTDATA(index, 17)
        vNameEnChild3 = .LISTDATA(index, 18)
        If .LISTDATA(index, 19) <> "" Then
            vBirthdayChild3 = .LISTDATA(index, 19)
        Else
            vBirthdayChild3 = DateSerial(1900, 1, 1)
        End If
        vPhoneChild3 = .LISTDATA(index, 20)
        vHome = .LISTDATA(index, 21)
        vFamily = .LISTDATA(index, 22)
        vHealth = .LISTDATA(index, 23)
        vOther = .LISTDATA(index, 24)
        vBaptism = .LISTDATA(index, 25)
        If .LISTDATA(index, 26) <> "" Then
            vOrdinationPrayer = .LISTDATA(index, 26)
        Else
            vOrdinationPrayer = DateSerial(1900, 1, 1)
        End If
        If .LISTDATA(index, 27) <> "" Then
            vAppoOvs = .LISTDATA(index, 27)
        Else
            vAppoOvs = DateSerial(1900, 1, 1)
        End If
        If .LISTDATA(index, 28) <> "" Then
            vWeddingDt = .LISTDATA(index, 28)
        Else
            vWeddingDt = DateSerial(1900, 1, 1)
        End If
        If .LISTDATA(index, 29) <> "" Then
            vTheologicalOrder = .LISTDATA(index, 29)
        Else
            vTheologicalOrder = 0
        End If
        vEducation = .LISTDATA(index, 30)
        vSalary = .LISTDATA(index, 31)
        vSuspend = .LISTDATA(index, 32)
        vOvsDept = .LISTDATA(index, 33)
    End With
End Sub

'@param argPStaff: argPStaff 객체를 변수로 받아 생명번호 기준으로 동일한지 여부를 점검합니다.
'@param blnStrictly: 변수 값이 True는 강한 검증을 말하며, 생명번호를 제외한 나머지 정보까지 모두 같을 때 True를 반환합니다.
'                    선택사항이며 기본값은 False 입니다.
Public Function IsEqual(ByRef argPStaff As PastoralStaff, Optional blnStrictly As Boolean = False) As Boolean
    
    If Me.lifeNo <> argPStaff.lifeNo Then IsEqual = False: Exit Function
    
    If blnStrictly And argPStaff.lifeNo <> "" Then
        If Me.nameKo <> argPStaff.nameKo Then IsEqual = False: Exit Function
        If Me.NameEn <> argPStaff.NameEn Then IsEqual = False: Exit Function
        If Me.Nationality <> argPStaff.Nationality Then IsEqual = False: Exit Function
        If Me.Birthday <> argPStaff.Birthday Then IsEqual = False: Exit Function
        If Me.Phone <> argPStaff.Phone Then IsEqual = False: Exit Function
        If Me.LifeNoChild1 <> argPStaff.LifeNoChild1 Then IsEqual = False: Exit Function
        If Me.NameKoChild1 <> argPStaff.NameKoChild1 Then IsEqual = False: Exit Function
        If Me.NameEnChild1 <> argPStaff.NameEnChild1 Then IsEqual = False: Exit Function
        If Me.BirthdayChild1 <> argPStaff.BirthdayChild1 Then IsEqual = False: Exit Function
        If Me.PhoneChild1 <> argPStaff.PhoneChild1 Then IsEqual = False: Exit Function
        If Me.LifeNoChild2 <> argPStaff.LifeNoChild2 Then IsEqual = False: Exit Function
        If Me.NameKoChild2 <> argPStaff.NameKoChild2 Then IsEqual = False: Exit Function
        If Me.NameEnChild2 <> argPStaff.NameEnChild2 Then IsEqual = False: Exit Function
        If Me.BirthdayChild2 <> argPStaff.BirthdayChild2 Then IsEqual = False: Exit Function
        If Me.PhoneChild2 <> argPStaff.PhoneChild2 Then IsEqual = False: Exit Function
        If Me.LifeNoChild3 <> argPStaff.LifeNoChild3 Then IsEqual = False: Exit Function
        If Me.NameKoChild3 <> argPStaff.NameKoChild3 Then IsEqual = False: Exit Function
        If Me.NameEnChild3 <> argPStaff.NameEnChild3 Then IsEqual = False: Exit Function
        If Me.BirthdayChild3 <> argPStaff.BirthdayChild3 Then IsEqual = False: Exit Function
        If Me.PhoneChild3 <> argPStaff.PhoneChild3 Then IsEqual = False: Exit Function
        If Me.Home <> argPStaff.Home Then IsEqual = False: Exit Function
        If Me.Family <> argPStaff.Family Then IsEqual = False: Exit Function
        If Me.Health <> argPStaff.Health Then IsEqual = False: Exit Function
        If Me.Other <> argPStaff.Other Then IsEqual = False: Exit Function
        If Me.Baptism <> argPStaff.Baptism Then IsEqual = False: Exit Function
        If Me.OrdinationPrayer <> argPStaff.OrdinationPrayer Then IsEqual = False: Exit Function
        If Me.AppoOvs <> argPStaff.AppoOvs Then IsEqual = False: Exit Function
        If Me.WeddingDt <> argPStaff.WeddingDt Then IsEqual = False: Exit Function
        If Me.TheologicalOrder <> argPStaff.TheologicalOrder Then IsEqual = False: Exit Function
        If Me.Education <> argPStaff.Education Then IsEqual = False: Exit Function
        If Me.Salary <> argPStaff.Salary Then IsEqual = False: Exit Function
        If Me.Suspend <> argPStaff.Suspend Then IsEqual = False: Exit Function
        If Me.OvsDept <> argPStaff.OvsDept Then IsEqual = False: Exit Function
    End If
    IsEqual = True
End Function
