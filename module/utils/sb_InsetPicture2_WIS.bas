Attribute VB_Name = "sb_InsetPicture2_WIS"
Option Explicit
Sub InsertChurchMap(Code As String, ByRef targetRange As Range)

    Dim objChurchMap As New ChurchMap
    Dim objChurchMapDao As New ChurchMapDao
    Dim filePath As String
    
    Set objChurchMap = objChurchMapDao.FindByChurchId(Code)
    filePath = saveBase64ToFileForMap(objChurchMap.map)
    sbInsertPicture2_WIS filePath, targetRange
    Kill filePath

End Sub

Sub InsertPStaffPic(lifeNo As String, ByRef targetRange As Range)

    Dim pPhoto As New PStaffPhoto
    Dim pPhotoDao As New PStaffPhotoDao
    Dim filePath As String
    
    Set pPhoto = pPhotoDao.FindByLifeNo(lifeNo)
    filePath = saveBase64ToFile(pPhoto.photo)
    sbInsertPicture2_WIS filePath, targetRange
    Kill filePath

End Sub

Sub sbInsertPicture2_WIS(Optional pic As Variant = False, Optional ByRef Pb As Range)

'----------------------------------------------------------------------------------------------------
'�� ���ν����� �������� ������ ������ ���·� ���õ� ��(Ȥ�� ����)�� ũ�⿡ ������ �����ϴ� �ڵ��Դϴ�.
'----------------------------------------------------------------------------------------------------

Dim kk

'--//Pb���� ������ ������ ���� Pb�� ����
If Pb Is Nothing Then Set Pb = Selection


'--//Pic ���� �����Ǿ� ���� ������ ���� ���̾�α� ��� ����
If pic = False Then
    pic = Application.GetOpenFilename(fileFilter:="Picturefile, *.jpg;*.jpeg; *.bmp; *.tif; *.png")
End If

'--//������ �������� ���� �� ���ν��� ����
If pic = False Then
    Exit Sub
End If

'--//���� ����
ActiveSheet.Pictures.Insert(pic).Select

'--//ũ�� �� ��ġ ����
With Selection
    .ShapeRange.LockAspectRatio = msoFalse '�׸� �¿�������� ����
    .Left = Pb.Left + 2 '��ġ���� �� �¿� ��ġ ����
    .Top = Pb.Top + 2 '��ġ���� �� ������ġ ����
    .Height = Pb.Height - 2 'ũ������ �� ��������
    .Width = Pb.Width - 4 'ũ������ �� �ʺ�����
End With

embed_Pics_Permanently Selection.ShapeRange.Item(1)

End Sub


Sub sbInsertPic_Call()
'
'��ȸ�� ��������Ȳ ��Ʈ�� ������ ������Ʈ �ϴ� �ڵ�
'
shUnprotect globalSheetPW

Call frm_Search_PStaff.sbInsertPic
'Selection.Delete
Range("B1").Select

shProtect globalSheetPW
MsgBox "������Ʈ �Ϸ�."
    
End Sub

