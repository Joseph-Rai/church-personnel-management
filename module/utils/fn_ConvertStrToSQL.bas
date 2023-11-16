Attribute VB_Name = "fn_ConvertStrToSQL"
Option Explicit

Public Function convertStrToSQL(argString As String)

    If IsNull(argString) Or argString = "" Then
        convertStrToSQL = "''"
        Exit Function
    End If
    
    If IsNumeric(argString) Then
        convertStrToSQL = CDbl(argString)
    Else
        convertStrToSQL = SText(argString)
    End If
    
    'convertStrToSQL = IIf((IsNull(argString) Or argString = ""), "''", (IIf(IsNumeric(argString), argString, SText(argString))))

End Function

