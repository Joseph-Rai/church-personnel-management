Attribute VB_Name = "sb_setCBox"
Option Explicit

'-----------------------------------------------------------------------------------
'  �޺��ڽ� ����: ��ü���
'    - setCBox(�޺��ڽ��̸�, �޺��ڽ����� ���, ���̸�, ��ȸ��, ��ȸ����)
'-----------------------------------------------------------------------------------
Sub setCBox(ByRef argCBox As MSForms.comboBox, kindCBox As String, formNM As String, Optional referDate As Date, Optional user_authority As Integer = 1)
    Dim strSql As String
    If referDate = Empty Then referDate = today
    
    Select Case kindCBox
        Case "FX" '//ȭ�� �޺�
            With argCBox
                .ColumnCount = 3
                .ColumnHeads = False
                .ColumnWidths = "0,50,100" 'ȭ��id, ȭ���Ī, ȭ���Ī
                .TextColumn = 2
                .ListWidth = "150"
                .TextAlign = fmTextAlignLeft
                .IMEMode = fmIMEModeAlpha
                .Style = fmStyleDropDownCombo
            End With
            strSql = "SELECT currency_id, currency_un, currency_nm FROM fx_calculator.currencies ORDER BY sort_order;"
            loadDataToCBox argCBox, strSql, "fx_calculator.currencies", formNM
    End Select
End Sub
