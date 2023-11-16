Attribute VB_Name = "AttenDetail_Date_Move"
Option Explicit

Sub sbMoveRight_AttenDetail()

    If Range("AttenDetail_rngDate") < Range("AttenDetail_MaxDate") Then
        Range("AttenDetail_rngDate") = WorksheetFunction.EDate(Range("AttenDetail_rngDate"), 1)
    End If
End Sub

Sub sbMoveLeftAttenDetail()

    If WorksheetFunction.EDate(Range("AttenDetail_rngDate"), -12) > Range("AttenDetail_MinDate") Then
        Range("AttenDetail_rngDate") = WorksheetFunction.EDate(Range("AttenDetail_rngDate"), -1)
    End If
End Sub
