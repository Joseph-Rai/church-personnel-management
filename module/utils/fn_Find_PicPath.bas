Attribute VB_Name = "fn_Find_PicPath"
Option Explicit
'------------------------------------------------------------------------------------------
'   ��� :  ������ ���� ��θ� ã���ִ� �Լ�
'           ��θ� ã�� ������ ��� FALSE
'------------------------------------------------------------------------------------------
Public Function fnFindPicPath()

    Dim i As Long
    Dim filePath As String
    Dim FileName As String
    
    FileName = "*.jpg"
    
    '--//�������� ���ã��
    On Error Resume Next
    For i = 1 To 24
        filePath = Chr(66 + i) & Right(PICPATH_PSTAFF, Len(PICPATH_PSTAFF) - 1)
        FileName = Dir(filePath)
        If FileName <> "" Then Exit For
    Next
    On Error GoTo 0
    If FileName = "" Then
        'MsgBox "���� ������Ʈ �����Դϴ�. ���̵�ũ ������ Ȯ���� �ּ���.", vbCritical, "���� ������Ʈ ����"
        fnFindPicPath = ""
        Exit Function
    End If
'    FilePath = Left(Left(FilePath, Len(FilePath) - 1), InStrRev(Left(FilePath, Len(FilePath) - 1), "\") - 1)
    If Right(filePath, 1) <> Application.PathSeparator Then
        filePath = filePath & Application.PathSeparator
    End If
    
    fnFindPicPath = filePath

End Function
Public Function fnFindRepresentativePic()

    Dim i As Long
    Dim filePath As String
    Dim FileName As String
    
    filePath = fnFindPicPath
    FileName = Dir(filePath & "*.jpg")
    
    Do
        If FileName <> "" Then
            If Not InStr(FileName, "~") > 0 Then
                Exit Do
            Else
                FileName = Dir
            End If
        Else
            Exit Do
        End If
    Loop
    
    fnFindRepresentativePic = filePath & FileName

End Function
