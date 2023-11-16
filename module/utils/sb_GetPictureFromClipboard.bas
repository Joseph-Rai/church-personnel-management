Attribute VB_Name = "sb_GetPictureFromClipboard"
Option Explicit

'#################################################
'
'#################################################

Private Const SRCCOPY As Long = &HCC0020
Private Const RC_PALETTE As Long = &H100
Private Const SIZEPALETTE As Long = 104
Private Const RASTERCAPS As Long = 38
Private Type PALETTEENTRY
  peRed As Byte
  peGreen As Byte
  peBlue As Byte
  peFlags As Byte
End Type
Private Type LOGPALETTE
  palVersion As Integer
  palNumEntries As Integer
  palPalEntry(255) As PALETTEENTRY    ' Enough for 256 colors
End Type
Private Type GUID
  Data1 As Long
  Data2 As Integer
  Data3 As Integer
  Data4(7) As Byte
End Type
Private Type PICTDESC
  Size As Long
  Typ As Long
#If Win64 Then
  hPic As LongPtr
  hPal As LongPtr
#Else
  hPic As Long
  hPal As Long
#End If
End Type

#If VBA7 Then
Private Declare PtrSafe Function OleCreatePictureIndirect Lib "oleaut32" ( _
    PICDESC As PICTDESC, RefIID As GUID, ByVal fPictureOwnsHandle As Long, _
    IPic As IPicture) As Long
#Else
Private Declare Function OleCreatePictureIndirect Lib "oleaut32" ( _
    PICDESC As PICTDESC, RefIID As GUID, ByVal fPictureOwnsHandle As Long, _
    IPic As IPicture) As Long
#End If

Private Enum PictureType
  CF_BITMAP = 2
  CF_ENHMETAFILE = 14
End Enum

#If Win64 Then
Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Private Declare PtrSafe Function OpenClipboard Lib "user32" ( _
    ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function GetClipboardData Lib "user32" ( _
    ByVal wFormat As Long) As LongPtr
Private Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32" ( _
    ByVal wFormat As Long) As Long
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare PtrSafe Function CopyEnhMetaFile Lib "gdi32" Alias "CopyEnhMetaFileA" ( _
    ByVal hemfSrc As LongPtr, ByVal lpszFile As String) As LongPtr
Private Declare PtrSafe Function CopyImage Lib "user32" ( _
    ByVal Handle As LongPtr, ByVal imageType As Long, ByVal NewWidth As Long, _
    ByVal NewHeight As Long, ByVal lFlags As Long) As LongPtr
#Else
Private Declare Function CloseClipboard Lib "User32" () As Long
Private Declare Function OpenClipboard Lib "User32" (ByVal hwnd As Long) As Long
Private Declare Function GetClipboardData Lib "User32" (ByVal wFormat As Long) As Long
Private Declare Function IsClipboardFormatAvailable Lib "User32" ( _
    ByVal wFormat As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function CopyEnhMetaFile Lib "gdi32" Alias "CopyEnhMetaFileA" ( _
    ByVal hemfSrc As Long, ByVal lpszFile As String) As Long
Private Declare Function CopyImage Lib "User32" ( _
    ByVal Handle As Long, ByVal imageType As Long, ByVal NewWidth As Long, _
    ByVal NewHeight As Long, ByVal lFlags As Long) As Long
#End If

Public Function PictureFromShape(ByVal S As Shape) As IPicture
  'Wandelt ein Shape uber die Zwischenablage in ein Picture
  S.CopyPicture xlScreen, xlBitmap
  Set PictureFromShape = PictureFromClipboard
End Function

Public Function PictureFromClipboard() As IPicture
  'Return a bitmap or metafile picture from clipboard (type is auto detected)
  Const IMAGE_BITMAP = 0
  Const LR_COPYRETURNORG = &H4
#If VBA7 Then
  Dim hPic As LongPtr, hCopy As LongPtr
#Else
  Dim hPic As Long, hCopy As Long
#End If
  Dim result As Long, PicType As PictureType
  Dim Count As Integer

  'Check if the clipboard contains a possible format
  If IsClipboardFormatAvailable(CF_BITMAP) <> 0 Then
    PicType = CF_BITMAP
  ElseIf IsClipboardFormatAvailable(CF_ENHMETAFILE) <> 0 Then
    PicType = CF_ENHMETAFILE
  End If
  If PicType = 0 Then err.Raise 70, "PictureFromClipboard", "No valid picture in " & _
    "clipboard"

  'Get access to the clipboard
  Do
    result = OpenClipboard(0&)
    If result <> 1 Then
      CloseClipboard
      DoEvents
      Sleep 10
    End If
    Count = Count + 1
  Loop Until Count = 10 Or result = 1
  If result <> 1 Then err.Raise 70, "PictureFromClipboard", "Can not open the clipboard"

  'Get a handle to the image data
  hPic = GetClipboardData(PicType)
  If hPic = 0 Then
    CloseClipboard
    err.Raise err.LastDllError, "PictureFromClipboard"
  End If
  'Create our own copy of the image on the clipboard, in the appropriate format.
  If PicType = CF_BITMAP Then
    hCopy = CopyImage(hPic, IMAGE_BITMAP, 0, 0, LR_COPYRETURNORG)
  Else
    hCopy = CopyEnhMetaFile(hPic, vbNullString)
  End If
  If hCopy = 0 Then err.Raise err.LastDllError, "PictureFromClipboard"
  'Release the clipboard to other programs
  CloseClipboard
  'Convert it into a Picture object and return it
  Set PictureFromClipboard = CreatePicture(hCopy, 0, PicType)
End Function

#If VBA7 Then
Private Function CreatePicture(ByVal hPic As LongPtr, ByVal hPal As LongPtr, _
    Optional ByVal PicType As PictureType = CF_BITMAP) As IPicture
#Else
Private Function CreatePicture(ByVal hPic As Long, ByVal hPal As Long, _
    Optional ByVal PicType As PictureType = CF_BITMAP) As IPicture
#End If
  Const PICTYPE_BITMAP As Long = 1
  Const PICTYPE_ENHMETAFILE As Long = 4
  Dim IPictureIID As GUID
  Dim IPic As IPicture
  Dim tagPic As PICTDESC

  'Fill in the IPicture GUID {7BF80980-BF32-101A-8BBB-00AA00300CAB}
  With IPictureIID
    .Data1 = &H7BF80980
    .Data2 = &HBF32
    .Data3 = &H101A
    .Data4(0) = &H8B
    .Data4(1) = &HBB
    .Data4(2) = &H0
    .Data4(3) = &HAA
    .Data4(4) = &H0
    .Data4(5) = &H30
    .Data4(6) = &HC
    .Data4(7) = &HAB
  End With

  'Set the properties on the picture object
  With tagPic
    .Size = Len(tagPic)
    .hPic = hPic
    Select Case PicType
      Case CF_BITMAP
        .Typ = PICTYPE_BITMAP
        .hPal = hPal
      Case CF_ENHMETAFILE
        .Typ = PICTYPE_ENHMETAFILE
        .hPal = 0
      Case Else
        err.Raise 51, "CreatePicture", "Invalid picture type"
    End Select
  End With

  'Create a picture that will delete it's bitmap when it is finished with it
  OleCreatePictureIndirect tagPic, IPictureIID, 1, IPic
  If IPic Is Nothing Then err.Raise err.LastDllError, "CreatePicture"
  Set CreatePicture = IPic
End Function

