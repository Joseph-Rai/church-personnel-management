Attribute VB_Name = "sb_PicFileToDB"
Option Explicit

'--//�̹��� ���� �ϳ��� �����Ͽ� ��ȸ���� �̹��� ������ DB�� �����մϴ�.
Public Sub SaveChurchMap()

    Dim churchCode As String
    Dim filePath As String
    Dim objChurchMap As New ChurchMap
    Dim objChurchMapDao As New ChurchMapDao
    Dim stream As New ADODB.stream
    
    '--//���ϼ���
    With Application.FileDialog(msoFileDialogFilePicker)  '�������� â����
        .Filters.Clear
        .Filters.Add "Images", "*.jpg; *.jpeg; *.bmp; *.tif; *.png"
        .Show
        If .SelectedItems.Count = 0 Then Exit Sub
        filePath = .SelectedItems(1)
    End With
    
    '--//DB�� ����
    stream.Type = adTypeBinary
    stream.Open
    stream.LoadFromFile filePath
    
    objChurchMap.Id = Range("Atten_ChurchCode")
    objChurchMap.map = encodeBase64(stream.Read)
    objChurchMapDao.Save objChurchMap
    
    stream.Close
    
    MsgBox "��ȸ������ ����Ǿ����ϴ�.", vbInformation, banner
    
    '--//����� ���� ����
    Dim ws As Worksheet
    Set ws = Range("Atten_Church_Map").Parent
    ws.Unprotect globalSheetPW
    InsertChurchMap objChurchMap.Id, Range("Atten_Church_Map")
    ws.Protect globalSheetPW

End Sub

'--//������ �������� ��ȸ���� ������ �ϰ� DB�� �����մϴ�.
Public Sub PicFileToDBForChurchMap()

    Dim FileName As String
    Dim objChurchMap As New ChurchMap
    Dim objChurchMapDao As New ChurchMapDao
    Dim stream As New ADODB.stream
    Dim cntAffected As Integer
    Dim currentPath As String
    
    '--//��������
    With Application.FileDialog(msoFileDialogFolderPicker)  '�������� â����
        .Show
        If .SelectedItems.Count = 0 Then Exit Sub
        currentPath = .SelectedItems(1) & Application.PathSeparator
    End With
    
    '--//���� �� ���� ���� �˻�
    FileName = Dir(currentPath & "*.*")

    If FileName = "" Then
        MsgBox "������ ������ ����", vbInformation, banner
        Exit Sub
    End If
    
    stream.Type = adTypeBinary
    stream.Open
        
    '--//DB�� ����
    FileName = Dir(currentPath)
    Do While FileName <> ""
        If IsImageFile(FileName) Then
        
            stream.LoadFromFile currentPath & FileName
            
            Dim churchName As String
            Dim objChurch As New Church
            Dim objChurchDao As New ChurchDao
            churchName = Left(FileName, InStr(FileName, ".") - 1)
            Set objChurch = objChurchDao.FindByChurchName(churchName)
            If objChurch.Id = "" Then '--//��ⱳȸ ó��
                Set objChurch = objChurchDao.FindByChurchName(churchName & "_���")
            End If
            
            objChurchMap.Id = objChurch.Id
            objChurchMap.map = encodeBase64(stream.Read)
            objChurchMapDao.Save objChurchMap
            cntAffected = cntAffected + 1
        End If
        
        FileName = Dir
    Loop
    
    '--//�����޽���
    MsgBox "���� " & cntAffected & "���� ���������� ����Ǿ����ϴ�.", , banner

End Sub


'--//������ �������� �����ڻ����� �ϰ� DB�� �����մϴ�.
Public Sub PicFileToDBForPStaff()
    
    Dim FileName As String
    Dim pPhoto As New PStaffPhoto
    Dim pPhotoDao As New PStaffPhotoDao
    Dim stream As New ADODB.stream
    Dim cntAffected As Integer
    Dim currentPath As String
    
    '--//��������
    With Application.FileDialog(msoFileDialogFolderPicker)  '�������� â����
        .Show
        If .SelectedItems.Count = 0 Then Exit Sub
        currentPath = .SelectedItems(1) & Application.PathSeparator
    End With
    
    '--//���� �� ���� ���� �˻�
    FileName = Dir(currentPath & "*.*")

    If FileName = "" Then
        MsgBox "������ ������ ����", vbInformation, banner
        Exit Sub
    End If
    
    stream.Type = adTypeBinary
    stream.Open
        
    '--//DB�� ����
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
    
    '--//�����޽���
    MsgBox "���� " & cntAffected & "���� ���������� ����Ǿ����ϴ�.", , banner
    
End Sub

Private Function IsImageFile(FileName As String)
    
    Dim strExtension As String
    Dim ValidExtensions As Object
    
    Set ValidExtensions = CreateObject("System.Collections.ArrayList")
    ValidExtensions.Add "jpg"
    ValidExtensions.Add "jpeg"
'    ValidExtensions.Add "png"
'    ValidExtensions.Add "tif"
'    ValidExtensions.Add "bmp" JPG�� UserForm���� �ҷ����� ���ϹǷ� ������� �ʴ� ������ ��
    
    strExtension = Right(FileName, Len(FileName) - InStrRev(FileName, "."))
    
    If ValidExtensions.Contains(LCase(strExtension)) Then
        IsImageFile = True
    Else
        IsImageFile = False
    End If
    
End Function



