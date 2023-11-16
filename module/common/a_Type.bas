Attribute VB_Name = "A_Type"
Option Explicit

'//구조체 정의
Type T_USERS '//common.users
    USER_ID As Integer
    USER_NM As String
    USER_GB As String
    USER_PW As String
    USER_DEPT As String
    argIP As String
    argDB As String
    argUN As String
    argPW As String
    SUSPENDED As Integer '1: SUSPENDED
End Type
 
Type T_RESULT
    strSql As String
    affectedCount As Long
End Type

Type T_ATTENDANCE
    church_sid As String
    ATTENDANCE_DT As Date
    ONCE_ALL As Integer
    FORTH_ALL As Integer
    ONCE_STU As Integer
    FORTH_STU As Integer
    TITHE_ALL As Integer
    TITHE_STU As Integer
    BAPTISM_ALL As Integer
    Evangelist As Integer
    UL As Integer
    GL As Integer
End Type

Type T_CHURCHLIST
    church_sid As String
    CHURCH_NM As String
    CHURCH_GB As String
    MANAGER_CD As String
    MAIN_CHURCH_CD As String
    ovs_dept As String
    Suspend As Integer
    SORT_ORDER As Integer
End Type

Type T_HISTORY_CHURCH
    HIS_CD As Integer
    church_sid As String
    HIS_DT As Date
    HISTORY As String
End Type

Type T_OVS_DEPT
    DEPT_ID As Integer
    DEPT_LV1 As String
    DEPT_LV2 As String
    DEPT_LV3 As String
    DEPT_NM As String
    DEPT_PHONECARD As String
    DEPT_PICPATH As String
    SORT_ORDER As Integer
    SUSPEDED As Integer
End Type

Type T_PASTORALSTAFF
    lifeNo As String
    NAME_KO As String
    name_en As String
    Nationality As String
    Birthday As Date
    Phone As String
    lifeno_child1 As String
    name_ko_child1 As String
    name_en_child1 As String
    birthday_child1 As Date
    phone_child1 As String
    lifeno_child2 As String
    name_ko_child2 As String
    name_en_child2 As String
    birthday_child2 As Date
    phone_child2 As String
    lifeno_child3 As String
    name_ko_child3 As String
    name_en_child3 As String
    birthday_child3 As Date
    phone_child3 As String
    Home As String
    Family As String
    Health As String
    Other As String
    Baptism As String
    ordination_prayer As String
    appo_ovs As Date
    wedding_dt As Date
    theological_order As String
    Education As String
    Salary As Long
    Suspend As Integer
    ovs_dept As Integer
End Type

Type T_PASTORALWIFE
    lifeNo As String
    Nationality As String
    NAME_KO As String
    name_en As String
    Birthday As Date
    Phone As String
    Home As String
    Family As String
    Health As String
    Other As String
    LIFENO_SPOUSE As String
    Education As String
    Suspend As Integer
    ovs_dept As Integer
End Type

Type T_POSITION
    POSITION_CD As Integer
    lifeNo As String
    START_DT As Date
    END_DT As Date
    position As String
End Type

Type T_POSITION2
    POSITION2_CD As Integer
    lifeNo As String
    START_DT As Date
    END_DT As Date
    Position2 As String
End Type

Type T_THEOLOGICAL
    THEOLOGICAL_CD As Integer
    lifeNo As String
    LEVEL As String
    START_DT As Date
    END_DT As Date
    RESIGN_DT As Date
    RECOMMAND_CHURCH As String
End Type

Type T_TITLE
    TITLE_CD As Integer
    lifeNo As String
    START_DT As Date
    END_DT As Date
    title As String
End Type

Type T_TRANSFER
    TRANSFER_CD As Integer
    lifeNo As String
    START_DT As Date
    END_DT As Date
    church_sid As String
End Type

Type T_a_POSITION_SPOUSE
    position As String
    POSITION_SPOUSE As String
End Type

Type T_CHURCH_ESTA
    CHURCH_ESTA_CD As String
    CHURCH_SID_CUSTOM As String
    START_DT As Date
    END_DT As Date
    church_sid As String
End Type

Type T_FLIGHT_SCHEDULE
    FLIGHT_CD As String
    lifeNo As String
    FLIGHT_DT As Date
    DEPARTURE As String
    Destination As String
    VISIT_PURPOSE As String
End Type

Type T_BC_LEADER
    church_sid As String
    START_DT As Date
    END_DT As Date
    lifeNo As String
    RESPONSIBILITY As String
End Type

Type T_UNION
    CHURCH_SID_CUSTOM As String
    START_DT As Date
    END_DT As Date
    UNION As Integer
End Type

Type T_SERMON
    lifeNo As String
    SCORE_AVG As Double
    SUBJECT_COUNT As Integer
End Type

Type T_VISA
    visa_cd As Integer
    lifeNo As String
    START_DT As Date
    END_DT As Date
    Visa As String
    memo As String
End Type

Type T_FAMILY
    FAMILY_ID As Integer
    FAMILY_CD As Integer
    RELATIONS As String
    lifeNo As String
    NAME_KO As String
    name_en As String
    church_sid As String
    title As String
    position As String
    Birthday As Date
    Education As String
    RELIGION As String
    RECOGNITION As String
    memo As String
    Suspend As Integer
End Type

Type T_COUNSEL
    COUNSEL_ID As Integer
    LIFE_NO As String
    COUNSEL_DT As Date
    CATEGORY As String
    title As String
    CONTENT As String
    result As String
    REMARK As String
    STATUS As String
    ovs_dept As Integer
End Type

Type T_AUTHORITY
    TABLE_ID As Integer
    USER_ID As Integer
    AUTHORITY_ID As Integer
End Type

Type T_RECORD_SET
    LISTDATA() As String
    LISTFIELD() As String
    CNT_RECORD As Long
End Type
