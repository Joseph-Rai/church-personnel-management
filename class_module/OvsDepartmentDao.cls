VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OvsDepartmentDao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'@param departmentName: 부서명을 기준으로 DB에서 검색합니다.
Public Function FindByDepartmentName(ByVal departmentName As String) As OvsDepartment
    Dim sql As String
    Dim tRecordSet As New T_RECORD_SET
    
    sql = "SELECT * FROM " & TABLE_OVS_DEPT & _
                    " a WHERE a.dept_nm = " & SText(departmentName) & ";"
    
    Set tRecordSet = makeListData(sql, TABLE_OVS_DEPT)
    
    Dim OvsDept As New OvsDepartment
    If tRecordSet.CNT_RECORD > 0 Then
        OvsDept.ParseFromRecordSet tRecordSet
    Else
        Set OvsDept = Nothing
    End If
    
    Set FindByDepartmentName = OvsDept
End Function

'@param id: 부서 id를 기준으로 DB에서 검색합니다.
Public Function FindById(ByVal Id As Integer) As OvsDepartment
    Dim sql As String
    Dim tRecordSet As T_RECORD_SET
    
    sql = "SELECT * FROM " & TABLE_OVS_DEPT & _
                    " a WHERE a.dept_id = " & SText(Id) & ";"
    
     tRecordSet = makeListData(sql, TABLE_OVS_DEPT)
    
    Dim OvsDept As New OvsDepartment
    If tRecordSet.CNT_RECORD > 0 Then
        OvsDept.ParseFromRecordSet tRecordSet
    Else
        Set OvsDept = Nothing
    End If
    
    Set FindById = OvsDept
End Function
