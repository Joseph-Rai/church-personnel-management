Attribute VB_Name = "sb_setCBox"
Option Explicit

'-----------------------------------------------------------------------------------
'  ÄÞº¸¹Ú½º ¼³Á¤: ÀüÃ¼¸ñ·Ï
'    - setCBox(ÄÞº¸¹Ú½ºÀÌ¸§, ÄÞº¸¹Ú½º³»¿ë ¾à¾î, ÆûÀÌ¸§, Á¶È¸ÀÏ, Á¶È¸±ÇÇÑ)
'-----------------------------------------------------------------------------------
Sub setCBox(ByRef argCBox As MSForms.comboBox, kindCBox As String, formNM As String, Optional referDate As Date, Optional user_authority As Integer = 1)
    Dim strSql As String
    If referDate = Empty Then referDate = today
    
    Select Case kindCBox
        Case "FX" '//È­Æó ÄÞº¸
            With argCBox
                .ColumnCount = 3
                .ColumnHeads = False
                .ColumnWidths = "0,50,100" 'È­Æóid, È­Æó¾àÄª, È­Æó¸íÄª
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
