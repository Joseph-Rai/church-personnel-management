Attribute VB_Name = "sb_InsertPicToLabel"
Option Explicit

'@param argLabel: 사진을 넣고자 하는 Label Control 객체
'@param lifeNo: 생명번호(사진코드)

Public Const PSTAFF_PHOTO = "pstaff_photo"
Public Const PASSPORT_PHOTO = "passport_photo"
Public Const VISA_PHOTO = "visa_photo"

Public Sub InsertPicToLabel(ByRef argLabel As MSForms.label, lifeNo As String, Optional ByVal kind As String = PSTAFF_PHOTO)
    
    Select Case kind
        Case PSTAFF_PHOTO
            '--//사진추가
            Dim pPhoto As New PStaffPhoto
            Dim pPhotoDao As New PStaffPhotoDao
            
            '--//선지자 사진추가
            Set pPhoto = pPhotoDao.FindByLifeNo(lifeNo)
            argLabel.Picture = convertBase64toImage(pPhoto.photo)
        Case PASSPORT_PHOTO
            '--//사진추가
            Dim ppPhoto As New PassportPhoto
            Dim ppPhotoDao As New PassportPhotoDao
            
            '--//선지자 사진추가
            Set ppPhoto = ppPhotoDao.FindByLifeNo(lifeNo)
            If ppPhoto.photo <> "" Then
                argLabel.Picture = convertBase64toImage(ppPhoto.photo, PASSPORT_PHOTO)
            Else
                argLabel.Picture = LoadPicture("")
            End If
        Case VISA_PHOTO
            '--//사진추가
            Dim vPhoto As New VisaPhoto
            Dim vPhotoDao As New VisaPhotoDao
            
            '--//선지자 사진추가
            Set vPhoto = vPhotoDao.FindByVisaCode(lifeNo)
            If vPhoto.photo <> "" Then
                argLabel.Picture = convertBase64toImage(vPhoto.photo, VISA_PHOTO)
            Else
                argLabel.Picture = LoadPicture("")
            End If
    End Select
    

End Sub

