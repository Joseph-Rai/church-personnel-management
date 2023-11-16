Attribute VB_Name = "sb_embed_pics_permanently"
Option Explicit

Public Sub embed_Pics_Permanently(shpC As Shape)
  
    Dim picLeft As Single                                   '각 그림(사진)의 왼쪽 위치 넣을 변수
    Dim picTop As Single                                   '각 그림의 윗쪽 위치 넣을 변수
           
    picLeft = shpC.Left                       '링크된 그림의 왼쪽 위치를 변수에
    picTop = shpC.Top                      '링크된 그림의 윗쪽 위치를 변수에
    
    shpC.CopyPicture Format:=xlBitmap                                 '링크된 그림을 복사 Appearance:=xlPrinter,
'    ActiveSheet.PasteSpecial Link:=False    '그림 링크깨고 붙여넣기

    ActiveSheet.Paste
    shpC.Delete                                '링크된 그림을 삭제
    Selection.Left = picLeft                 '복사된 그림 왼쪽 위치를 링크된 그림 왼쪽위치에
    Selection.Top = picTop                 '복사된 그림 윗쪽 위치를 링크된 그림 왼쪽위치에
    
'On Error Resume Next
'    Dim cnt As Integer
'RETRY:
'    ActiveSheet.Paste
'    If err.Number = 0 Then
'        shpC.Delete                                '링크된 그림을 삭제
'        Selection.Left = picLeft                 '복사된 그림 왼쪽 위치를 링크된 그림 왼쪽위치에
'        Selection.Top = picTop                 '복사된 그림 윗쪽 위치를 링크된 그림 왼쪽위치에
'    Else
'        cnt = cnt + 1
'        If cnt < 100 Then
'            err.Number = 0
'            GoTo RETRY
'        End If
'    End If
'On Error GoTo 0
   
End Sub
