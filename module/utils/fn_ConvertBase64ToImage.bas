Attribute VB_Name = "fn_ConvertBase64ToImage"
Option Explicit

Public Function convertImageToBase64(filePath As String) As String
    Dim inputStream
    Set inputStream = CreateObject("ADODB.Stream")
    
    inputStream.Open
    inputStream.Type = adTypeBinary
    inputStream.LoadFromFile filePath
    
    Dim Bytes: Bytes = inputStream.Read
    Dim dom: Set dom = CreateObject("Microsoft.XMLDOM")
    Dim elem: Set elem = dom.createElement("tmp")
    
    elem.dataType = "bin.base64"
    elem.nodeTypedValue = Bytes
    convertImageToBase64 = elem.text
End Function

Public Function convertBase64toImage(base64 As String, Optional ByVal kind As String = PSTAFF_PHOTO)

    Dim defaultFilePath As String
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("temp")
    
    Select Case kind
        Case PSTAFF_PHOTO
            defaultFilePath = saveBase64ToFile(base64)
        Case PASSPORT_PHOTO
            defaultFilePath = saveBase64ToFileForPassport(base64)
        Case VISA_PHOTO
            defaultFilePath = saveBase64ToFileForVisa(base64)
    End Select
    
    
On Error GoTo ErrHandler
    '--//Office 2010 이상 사용자는 File로 Save하는 과정에서 ShadowCube 잠김으로 인한 오류발생
    Set convertBase64toImage = LoadPicture(defaultFilePath)
    
    Kill defaultFilePath
    
DONE:
    Exit Function
    
ErrHandler:
    
    '--//에러나면 사진을 Shape 형태로 Import하고 나서 클립보드를 이용해 IPicture 객체 획득
    Dim shp As Shape
    
    Application.ScreenUpdating = False
    Set shp = ws.Shapes.AddPicture(defaultFilePath, msoFalse, msoTrue, 0, 0, 300, 400)
    Set convertBase64toImage = PictureFromShape(shp)
    shp.Delete
    Application.ScreenUpdating = True
    
    Kill defaultFilePath

On Error GoTo 0

End Function

Public Function saveBase64ToFileDefault(base64 As String)

    Dim stream As New ADODB.stream
    Dim defaultFilePath As String: defaultFilePath = "C:\Temp\temp.dat"
    Dim parentDir As String
    
    '--//임시 저장경로 없는 경우 생성
    parentDir = Left(defaultFilePath, InStrRev(defaultFilePath, "\"))
    If Dir(parentDir, vbDirectory) = "" Then
        MkDir parentDir
    End If
    
    stream.Type = adTypeBinary
    stream.Open
    stream.Write decodeBase64(base64)
    stream.saveToFile defaultFilePath, adSaveCreateOverWrite
    stream.Close
    
    saveBase64ToFileDefault = defaultFilePath

End Function

Public Function saveBase64ToFile(base64 As String)

    If base64 = "" Then
        Dim pPhoto As New PStaffPhoto
        Dim pPhotoDao As New PStaffPhotoDao
        
        Set pPhoto = pPhotoDao.FindByLifeNo("사진없음")
        base64 = pPhoto.photo
    End If
    
    saveBase64ToFile = saveBase64ToFileDefault(base64)
    
End Function

Public Function saveBase64ToFileForMap(base64 As String)
    
    If base64 = "" Then
        Dim objChurchMap As New ChurchMap
        Dim objChurchMapDao As New ChurchMapDao
        
        Set objChurchMap = objChurchMapDao.FindByChurchId("지도없음")
        base64 = objChurchMap.map
    End If
    
    saveBase64ToFileForMap = saveBase64ToFileDefault(base64)
    
End Function

'@param: base64가 ""인 경우는 입력되지 않는다
Public Function saveBase64ToFileForPassport(base64 As String)
    
'    If base64 = "" Then
'        Dim objPassportPhoto As New PassportPhoto
'        Dim objPassportPhotoDao As New PassportPhotoDao
'
'        Set objPassportPhoto = objPassportPhotoDao.FindByLifeNo("여권없음")
'        base64 = objPassportPhoto.photo
'    End If
    
    saveBase64ToFileForPassport = saveBase64ToFileDefault(base64)
    
End Function

'@param: base64가 ""인 경우는 입력되지 않는다
Public Function saveBase64ToFileForVisa(base64 As String)
    
'    If base64 = "" Then
'        Dim objVisaPhoto As New VisaPhoto
'        Dim objVisaPhotoDao As New VisaPhotoDao
'
'        Set objVisaPhoto = objVisaPhotoDao.FindByVisaCode("사증없음")
'        base64 = objVisaPhoto.photo
'    End If
    
    saveBase64ToFileForVisa = saveBase64ToFileDefault(base64)
    
End Function
