Attribute VB_Name = "fn_IsZero"
Option Explicit

Public Function IsZero(num As Variant) As Boolean

    If num = 0 Then
        IsZero = True
    Else
        IsZero = False
    End If

End Function

Public Function IfZero(num As Variant, argReplacement As Variant)

    If num = 0 Then
        IfZero = argReplacement
    Else
        IfZero = num
    End If

End Function
