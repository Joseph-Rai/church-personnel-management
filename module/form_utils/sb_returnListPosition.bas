Attribute VB_Name = "sb_returnListPosition"
Option Explicit

'------------------------------------------------------------------------------------------------------------------
'  ������ ����Ʈ �׸����� �̵�
'    - ReturnListPosition(���̸�, ����Ʈ�̸�, key��)
'------------------------------------------------------------------------------------------------------------------
Sub returnListPosition(ByRef argForm As UserForm, ByVal argList As String, ByVal argKey As String)
    Dim i As Long
    Dim colKey As Integer 'list���� key���� �÷���ġ
    
    '//list.BoundColumn�� �⺻���� 1
    colKey = argForm.controls(argList).BoundColumn - 1
    
    '//listbox���� ������ item��ġ�� �̵�
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
    Dim colKey As Integer 'list���� key���� �÷���ġ
    Dim ArrKey As String
    
    '//list.BoundColumn�� �⺻���� 1
    colKey = argForm.controls(argList).BoundColumn - 1
    ArrKey = argKey
    
    '//listbox���� ������ item��ġ�� �̵�
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
    Dim colKey As Integer 'list���� key���� �÷���ġ
    
    '//list.BoundColumn�� �⺻���� 1
    colKey = argForm.controls(argList).BoundColumn - 1
    
    '//listbox���� ������ item��ġ�� �̵�
    With argForm.controls(argList)
        For i = 0 To .ListCount - 1
            If CStr(.List(i, colKey + 11)) = CStr(argKey) Then
               .listIndex = i
                Exit For
            End If
        Next i
    End With
End Sub
