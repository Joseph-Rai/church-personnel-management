Attribute VB_Name = "fn_Find_PicPath"
Option Explicit
'------------------------------------------------------------------------------------------
'   기능 :  선지자 사진 경로를 찾아주는 함수
'           경로를 찾지 못했을 경우 FALSE
'------------------------------------------------------------------------------------------
Public Function fnFindPicPath()

    Dim i As Long
    Dim filePath As String
    Dim FileName As String
    
    FileName = "*.jpg"
    
    '--//사진파일 경로찾기
    On Error Resume Next
    For i = 1 To 24
        filePath = Chr(66 + i) & Right(PICPATH_PSTAFF, Len(PICPATH_PSTAFF) - 1)
        FileName = Dir(filePath)
        If FileName <> "" Then Exit For
    Next
    On Error GoTo 0
    If FileName = "" Then
        'MsgBox "사진 업데이트 오류입니다. 마이디스크 연결을 확인해 주세요.", vbCritical, "사진 업데이트 오류"
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
