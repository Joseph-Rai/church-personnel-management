Attribute VB_Name = "a_ErrHandler"
Option Explicit

'-----------------------------------------------------------------------------------------------------
'  ����ó��: errhandle(���ν�����, ���̺��, SQL��, ���̸�, �۾���)
'    - ���� �߻� ������ ����� �ϱ� ���� �޽��� �ڽ��� ǥ��
'    - ���� �߻��� ���� �α� ����� DB�� ����� ���븸 callDBroRS, executeSQL���� ����
'-----------------------------------------------------------------------------------------------------
Sub ErrHandle(procedureNM As String, Optional tableNM As String = "NULL", Optional SQLScript As String = "NULL", Optional formNM As String = "NULL", Optional jobNM As String = "��Ÿ")
    If err.Number <> 0 Then
        MsgBox "������ �߻��߽��ϴ�." & space(7) & vbNewLine & _
            " �� ������ �߻��� ������ ĸó�Ͽ� �����ڿ��� �����ּ���." & vbNewLine & vbNewLine & _
            "  �� �۾��� : " & Application.UserName & vbNewLine & _
            "  �� �۾��Ͻ� : " & Now & vbNewLine & _
            "  �� �۾����� : " & jobNM & vbNewLine & vbNewLine & _
            "  �� ���� �߻� vba : " & procedureNM & vbNewLine & _
            "  �� ���� �߻� �� : " & formNM & vbNewLine & _
            "  �� ���� �߻� DB : " & tableNM & vbNewLine & _
            "  �� ���� �߻� Script : " & SQLScript & vbNewLine & vbNewLine & vbNewLine & _
            "  �� ���� �ڵ� : " & err.Number & vbNewLine & _
            "  �� ���� ���� : " & err.Description & vbNewLine & _
            "  �� ���� �ҽ� : " & err.Source
    End If
End Sub

