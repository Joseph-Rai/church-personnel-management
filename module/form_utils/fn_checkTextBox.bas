Attribute VB_Name = "fn_checkTextBox"
Option Explicit

'------------------------------------------------------------------------------------------------------------------
'  TextBox 입력 값 검증
'    - checkTextBox(텍스트박스 이름, 텍스트박스 타이틀, 필수여부, 데이터형, 길이 제한, 입력 후 포커싱)
'    - 예: If checkTextBox(Me.txt1, "계정과목", True, "STRING", , False) = False Then Exit Sub
'------------------------------------------------------------------------------------------------------------------
Public Function checkTextBox(ByRef argTxtbox As MSForms.textBox, ByVal title As String, _
                                                Optional ByVal isEssencial As Boolean = False, _
                                                Optional ByVal dataType As String = "STRING", _
                                                Optional ByVal Length As Integer = 1000, _
                                                Optional ByVal isSetFocus As Boolean = True) As Boolean
    Dim checkResult As Boolean
    checkResult = True
    
   '필수입력 검증
    If isEssencial = True And (IsNull(argTxtbox.text) Or argTxtbox.text = "") Then
        MsgBox title & "을(를) 입력하세요." & space(7), vbInformation, banner
        If isSetFocus = True Then
            argTxtbox.SetFocus
        End If
        argTxtbox = Empty
        checkResult = False
    End If
    
    '데이터형 검증: 숫자, 날짜
    If (dataType = "NUMERIC") Then
        If Not IsNumeric(argTxtbox.text) Then
            MsgBox title & "에는 숫자만 입력할 수 있습니다." & space(7), vbInformation, banner
            If isSetFocus = True Then
                argTxtbox.SetFocus
            End If
            argTxtbox = Empty
            checkResult = False
        End If
    Else
        If (dataType = "DATE") Then
            If Not IsDate(argTxtbox.text) Then
                MsgBox title & "에는 날짜만 입력할 수 있습니다." & _
                vbNewLine & "YYYY-MM-DD 형태로 입력하세요." & space(7), vbInformation, banner
                If isSetFocus = True Then
                    argTxtbox.SetFocus
                End If
                argTxtbox = Empty
                checkResult = False
            End If
        End If
    End If
    
     '입력길이 검증
     If (Not IsNull(argTxtbox.text)) And Len(argTxtbox.text) > Length Then
        MsgBox title & "의 입력 최대 길이는 " & CStr(Length) & "를 넘을 수 없습니다." & space(7), vbInformation, banner
        If isSetFocus = True Then
            argTxtbox.SetFocus
        End If
        argTxtbox = Empty
        checkResult = False
    End If
    
    '함수결과의 True, False로 반환
    checkTextBox = checkResult
End Function

