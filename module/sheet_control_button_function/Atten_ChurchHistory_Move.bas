Attribute VB_Name = "Atten_ChurchHistory_Move"
Option Explicit

Sub shp_MoveUp_Click()
'    shUnprotect globalSheetPW
    Range("Atten_rngHistory_Index") = IIf(Range("Atten_rngHistory_Index") - 1 < 1, 1, Range("Atten_rngHistory_Index") - 1)
'    shProtect globalSheetPW
End Sub
Sub shp_MoveDown_Click()
'    shUnprotect globalSheetPW
    Range("Atten_rngHistory_Index") = IIf(Range("Atten_rngHistory_cntRecord") < 10, 1, IIf(Range("Atten_rngHistory_Index") + 1 > Range("Atten_rngHistory_cntRecord").Value - 9, Range("Atten_rngHistory_cntRecord") - 9, Range("Atten_rngHistory_Index") + 1))
'    shProtect globalSheetPW
End Sub
Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
'

'
    ActiveSheet.Outline.ShowLevels RowLevels:=0, ColumnLevels:=1
    ActiveSheet.Outline.ShowLevels RowLevels:=0, ColumnLevels:=2

    Columns("DB:DB").Select
    Range("DB50").Activate
    Selection.Columns.Group
    ActiveSheet.Outline.ShowLevels RowLevels:=0, ColumnLevels:=1
End Sub
