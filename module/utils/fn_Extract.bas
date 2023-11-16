Attribute VB_Name = "fn_Extract"
'------------------------------------------------------------------------
' 주제 : 도서출판 길벗 엑셀 2016 VBA - 유용한 함수&매크로 기능
'------------------------------------------------------------------------
Option Explicit
Option Compare Text
Option Base 1

'------------------------------------------------------------------------------------------
'   기능 :  구분은
'            숫자(구분:N, 생략), 영문자(구분:E), 한글(구분:H), 기타문자(구분:O)
'            만 추출하여 반환
'------------------------------------------------------------------------------------------
Function fnExtract(문자열 As String, Optional 구분 As String = "N") As String
Attribute fnExtract.VB_Description = "텍스트에서 숫자, 영문자, 한글, 특수문자만 추출하는 함수"
Attribute fnExtract.VB_ProcData.VB_Invoke_Func = " \n17"
  Dim i As Integer
  Dim k As String
  '--// 숫자, 영문, 한글, 기타 문자를 저장할 변수
  Dim NumStr As String, EngStr As String, HanStr As String, EtcStr As String  '기타 글자들을 기억함
                                 
  Application.Volatile
  
  For i = 1 To Len(문자열)
      k = Mid(문자열, i, 1)
      Select Case k
         Case "0" To "9"
           NumStr = NumStr & k
         Case "."
           NumStr = NumStr & k
         Case "A" To "Z"
           EngStr = EngStr & k
         Case "a" To "z"
           EngStr = EngStr & k
         Case "ㄱ" To "홓"    '한글은 'ㄱ'이 가장 작고 '홓'이 가장 큰 글자
           HanStr = HanStr & k
         Case Else
           EtcStr = EtcStr & k
      End Select
  Next
  
  Select Case 구분
      Case "N":          fnExtract = NumStr
      Case "E":          fnExtract = EngStr
      Case "H":          fnExtract = HanStr
      Case "O":          fnExtract = EtcStr
      Case Else:         fnExtract = NumStr
  End Select
End Function


