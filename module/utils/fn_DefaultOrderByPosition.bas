Attribute VB_Name = "fn_DefaultOrderByPosition"
Option Explicit

Public Function GetDefaultOrderByPosition() As String

    GetDefaultOrderByPosition = _
        "FIELD(`직책`,'당회장','당회장대리','당사모','당대리사모','동역','동사모','지교회관리자','지관자사모','예배소관리자','예관자사모','예비생도1단계','예비생도2단계','예비생도3단계','생도사모'," & getPosition2Joining & ",'')"

End Function
