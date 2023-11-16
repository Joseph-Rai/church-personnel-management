Attribute VB_Name = "sb_InsertPicToLabel"
Option Explicit

'@param argLabel: ������ �ְ��� �ϴ� Label Control ��ü
'@param lifeNo: �����ȣ(�����ڵ�)

Public Const PSTAFF_PHOTO = "pstaff_photo"
Public Const PASSPORT_PHOTO = "passport_photo"
Public Const VISA_PHOTO = "visa_photo"

Public Sub InsertPicToLabel(ByRef argLabel As MSForms.label, lifeNo As String, Optional ByVal kind As String = PSTAFF_PHOTO)
    
    Select Case kind
        Case PSTAFF_PHOTO
            '--//�����߰�
            Dim pPhoto As New PStaffPhoto
            Dim pPhotoDao As New PStaffPhotoDao
            
            '--//������ �����߰�
            Set pPhoto = pPhotoDao.FindByLifeNo(lifeNo)
            argLabel.Picture = convertBase64toImage(pPhoto.photo)
        Case PASSPORT_PHOTO
            '--//�����߰�
            Dim ppPhoto As New PassportPhoto
            Dim ppPhotoDao As New PassportPhotoDao
            
            '--//������ �����߰�
            Set ppPhoto = ppPhotoDao.FindByLifeNo(lifeNo)
            If ppPhoto.photo <> "" Then
                argLabel.Picture = convertBase64toImage(ppPhoto.photo, PASSPORT_PHOTO)
            Else
                argLabel.Picture = LoadPicture("")
            End If
        Case VISA_PHOTO
            '--//�����߰�
            Dim vPhoto As New VisaPhoto
            Dim vPhotoDao As New VisaPhotoDao
            
            '--//������ �����߰�
            Set vPhoto = vPhotoDao.FindByVisaCode(lifeNo)
            If vPhoto.photo <> "" Then
                argLabel.Picture = convertBase64toImage(vPhoto.photo, VISA_PHOTO)
            Else
                argLabel.Picture = LoadPicture("")
            End If
    End Select
    

End Sub

