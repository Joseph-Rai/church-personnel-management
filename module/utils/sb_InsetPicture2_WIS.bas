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
'이 프로시저는 사진비율 고정을 해제한 상태로 선택된 셀(혹은 범위)의 크기에 사진을 삽입하는 코드입니다.
'----------------------------------------------------------------------------------------------------

Dim kk

'--//Pb값이 없으면 선택한 셀을 Pb로 설정
If Pb Is Nothing Then Set Pb = Selection


'--//Pic 값이 설정되어 있지 않으면 파일 다이얼로그 열어서 선택
If pic = False Then
    pic = Application.GetOpenFilename(fileFilter:="Picturefile, *.jpg;*.jpeg; *.bmp; *.tif; *.png")
End If

'--//파일을 선택하지 않을 시 프로시저 종료
If pic = False Then
    Exit Sub
End If

'--//사진 삽입
ActiveSheet.Pictures.Insert(pic).Select

'--//크기 및 위치 조정
With Selection
    .ShapeRange.LockAspectRatio = msoFalse '그림 좌우고정비율 해제
    .Left = Pb.Left + 2 '위치조정 중 좌우 위치 조정
    .Top = Pb.Top + 2 '위치조정 중 상하위치 조정
    .Height = Pb.Height - 2 '크기조정 중 높이조정
    .Width = Pb.Width - 4 '크기조정 중 너비조정
End With

embed_Pics_Permanently Selection.ShapeRange.Item(1)

End Sub


Sub sbInsertPic_Call()
'
'교회별 선지자현황 시트의 사진만 업데이트 하는 코드
'
shUnprotect globalSheetPW

Call frm_Search_PStaff.sbInsertPic
'Selection.Delete
Range("B1").Select

shProtect globalSheetPW
MsgBox "업데이트 완료."
    
End Sub

