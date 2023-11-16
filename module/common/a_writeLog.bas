Attribute VB_Name = "a_WriteLog"
Option Explicit

'----------------------------------------------------------------------------------------------------
'  �αױ��
'    - �����α�: SQL�� ���� �� �߻��� �α׸� ���(executeSQL, callDBtoRS)
'    - �׼Ƿα�: DB�� ������ ���ν���(Insert, Update, Delete) ���� �� �α� ���
'    - writelog(���ν�����, ���̺��, SQL, �����ڵ�, ���̸�, ���̸�, ����������ڵ��)
'-----------------------------------------------------------------------------------------------------
Sub writeLog(procedureNM As String, tableNM As String, SQLScript As String, ErrorCD As Integer, Optional formNM As String = "NULL", Optional jobNM As String = "NULL", _
                     Optional affectedCount As Long = 0)
    Dim strSql As String
    connectCommonDB
    
    strSql = "INSERT INTO common.logs(procedure_nm, table_nm, sql_script, error_cd, form_nm, job_nm, affectedCount, user_id) " & _
                  "Values(" & SText(procedureNM) & ", " & _
                                    SText(tableNM) & ", " & _
                                    SText(Replace(SQLScript, ";", "")) & ", " & _
                                    ErrorCD & ", " & _
                                    SText(formNM) & ", " & _
                                    SText(jobNM) & ", " & _
                                    affectedCount & ", " & _
                                    USER_ID & ");"

    executeSQL "writeLog", "common.logs", strSql, , "�αױ��"
    disconnectDB
End Sub


