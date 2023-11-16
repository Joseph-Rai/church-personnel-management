Attribute VB_Name = "sb_PicFileToDB"
Option Explicit

'--//이미지 파일 하나를 선택하여 교회지도 이미지 파일을 DB에 저장합니다.
Public Sub SaveChurchMap()

    Dim churchCode As String
    Dim filePath As String
    Dim objChurchMap As New ChurchMap
    Dim objChurchMapDao As New ChurchMapDao
    Dim stream As New ADODB.stream
    
    '--//파일선택
    With Application.FileDialog(msoFileDialogFilePicker)  '폴더선택 창에서
        .Filters.Clear
        .Filters.Add "Images", "*.jpg; *.jpeg; *.bmp; *.tif; *.png"
        .Show
        If .SelectedItems.Count = 0 Then Exit Sub
        filePath = .SelectedItems(1)
    End With
    
    '--//DB에 저장
    stream.Type = adTypeBinary
    stream.Open
    stream.LoadFromFile filePath
    
    objChurchMap.Id = Range("Atten_ChurchCode")
    objChurchMap.map = encodeBase64(stream.Read)
    objChurchMapDao.Save objChurchMap
    
    stream.Close
    
    MsgBox "교회지도가 저장되었습니다.", vbInformation, banner
    
    '--//저장된 사진 삽입
    Dim ws As Worksheet
    Set ws = Range("Atten_Church_Map").Parent
    ws.Unprotect globalSheetPW
    InsertChurchMap objChurchMap.Id, Range("Atten_Church_Map")
    ws.Protect globalSheetPW

End Sub

'--//지정된 폴더에서 교회지도 파일을 일괄 DB에 저장합니다.
Public Sub PicFileToDBForChurchMap()

    Dim FileName As String
    Dim objChurchMap As New ChurchMap
    Dim objChurchMapDao As New ChurchMapDao
    Dim stream As New ADODB.stream
    Dim cntAffected As Integer
    Dim currentPath As String
    
    '--//폴더선택
    With Application.FileDialog(msoFileDialogFolderPicker)  '폴더선택 창에서
        .Show
        If .SelectedItems.Count = 0 Then Exit Sub
        currentPath = .SelectedItems(1) & Application.PathSeparator
    End With
    
    '--//폴더 내 파일 유무 검사
    FileName = Dir(currentPath & "*.*")

    If FileName = "" Then
        MsgBox "폴더에 파일이 없음", vbInformation, banner
        Exit Sub
    End If
    
    stream.Type = adTypeBinary
    stream.Open
        
    '--//DB에 저장
    FileName = Dir(currentPath)
    Do While FileName <> ""
        If IsImageFile(FileName) Then
        
            stream.LoadFromFile currentPath & FileName
            
            Dim churchName As String
            Dim objChurch As New Church
            Dim objChurchDao As New ChurchDao
            churchName = Left(FileName, InStr(FileName, ".") - 1)
            Set objChurch = objChurchDao.FindByChurchName(churchName)
            If objChurch.Id = "" Then '--//폐쇄교회 처리
                Set objChurch = objChurchDao.FindByChurchName(churchName & "_폐쇄")
            End If
            
            objChurchMap.Id = objChurch.Id
            objChurchMap.map = encodeBase64(stream.Read)
            objChurchMapDao.Save objChurchMap
            cntAffected = cntAffected + 1
        End If
        
        FileName = Dir
    Loop
    
    '--//성공메시지
    MsgBox "지도 " & cntAffected & "건이 성공적으로 저장되었습니다.", , banner

End Sub


'--//지정된 폴더에서 선지자사진을 일괄 DB에 저장합니다.
Public Sub PicFileToDBForPStaff()
    
    Dim FileName As String
    Dim pPhoto As New PStaffPhoto
    Dim pPhotoDao As New PStaffPhotoDao
    Dim stream As New ADODB.stream
    Dim cntAffected As Integer
    Dim currentPath As String
    
    '--//폴더선택
    With Application.FileDialog(msoFileDialogFolderPicker)  '폴더선택 창에서
        .Show
        If .SelectedItems.Count = 0 Then Exit Sub
        currentPath = .SelectedItems(1) & Application.PathSeparator
    End With
    
    '--//폴더 내 파일 유무 검사
    FileName = Dir(currentPath & "*.*")

    If FileName = "" Then
        MsgBox "폴더에 파일이 없음", vbInformation, banner
        Exit Sub
    End If
    
    stream.Type = adTypeBinary
    stream.Open
        
    '--//DB에 저장
    FileName = Dir(currentPath)
    Do While FileName <> ""
        If IsImageFile(FileName) Then
        
            stream.LoadFromFile currentPath & FileName
            
            pPhoto.lifeNo = Left(FileName, InStr(FileName, ".") - 1)
            pPhoto.photo = encodeBase64(stream.Read)
            pPhotoDao.Save pPhoto
            cntAffected = cntAffected + 1
        End If
        
        FileName = Dir
    Loop
    
    '--//성공메시지
    MsgBox "사진 " & cntAffected & "건이 성공적으로 저장되었습니다.", , banner
    
End Sub

Private Function IsImageFile(FileName As String)
    
    Dim strExtension As String
    Dim ValidExtensions As Object
    
    Set ValidExtensions = CreateObject("System.Collections.ArrayList")
    ValidExtensions.Add "jpg"
    ValidExtensions.Add "jpeg"
'    ValidExtensions.Add "png"
'    ValidExtensions.Add "tif"
'    ValidExtensions.Add "bmp" JPG는 UserForm에서 불러오지 못하므로 등록하지 않는 것으로 함
    
    strExtension = Right(FileName, Len(FileName) - InStrRev(FileName, "."))
    
    If ValidExtensions.Contains(LCase(strExtension)) Then
        IsImageFile = True
    Else
        IsImageFile = False
    End If
    
End Function



