Attribute VB_Name = "sb_returnListPosition"
Option Explicit

'------------------------------------------------------------------------------------------------------------------
'  원래의 리스트 항목으로 이동
'    - ReturnListPosition(폼이름, 리스트이름, key값)
'------------------------------------------------------------------------------------------------------------------
Sub returnListPosition(ByRef argForm As UserForm, ByVal argList As String, ByVal argKey As String)
    Dim i As Long
    Dim colKey As Integer 'list에서 key값의 컬럼위치
    
    '//list.BoundColumn의 기본값은 1
    colKey = argForm.controls(argList).BoundColumn - 1
    
    '//listbox에서 지정된 item위치로 이동
    With argForm.controls(argList)
        For i = 0 To .ListCount - 1
            If CStr(.List(i, colKey)) = CStr(argKey) Then
               .listIndex = i
                Exit For
            End If
        Next i
    End With
End Sub
Sub returnListPosition2(ByRef argForm As UserForm, ByVal argList As String, ByVal argKey As String)
    Dim i As Long
    Dim colKey As Integer 'list에서 key값의 컬럼위치
    Dim ArrKey As String
    
    '//list.BoundColumn의 기본값은 1
    colKey = argForm.controls(argList).BoundColumn - 1
    ArrKey = argKey
    
    '//listbox에서 지정된 item위치로 이동
    With argForm.controls(argList)
        For i = 0 To .ListCount - 1
            If CStr(.List(i, colKey + 1)) = CStr(ArrKey) Then
               .listIndex = i
                Exit For
            End If
        Next i
    End With
End Sub
Sub returnListPosition3(ByRef argForm As UserForm, ByVal argList As String, ByVal argKey As String)
    Dim i As Long
    Dim colKey As Integer 'list에서 key값의 컬럼위치
    
    '//list.BoundColumn의 기본값은 1
    colKey = argForm.controls(argList).BoundColumn - 1
    
    '//listbox에서 지정된 item위치로 이동
    With argForm.controls(argList)
        For i = 0 To .ListCount - 1
            If CStr(.List(i, colKey + 11)) = CStr(argKey) Then
               .listIndex = i
                Exit For
            End If
        Next i
    End With
End Sub
