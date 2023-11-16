Attribute VB_Name = "fn_checkTextBox"
Option Explicit

'------------------------------------------------------------------------------------------------------------------
'  TextBox �Է� �� ����
'    - checkTextBox(�ؽ�Ʈ�ڽ� �̸�, �ؽ�Ʈ�ڽ� Ÿ��Ʋ, �ʼ�����, ��������, ���� ����, �Է� �� ��Ŀ��)
'    - ��: If checkTextBox(Me.txt1, "��������", True, "STRING", , False) = False Then Exit Sub
'------------------------------------------------------------------------------------------------------------------
Public Function checkTextBox(ByRef argTxtbox As MSForms.textBox, ByVal title As String, _
                                                Optional ByVal isEssencial As Boolean = False, _
                                                Optional ByVal dataType As String = "STRING", _
                                                Optional ByVal Length As Integer = 1000, _
                                                Optional ByVal isSetFocus As Boolean = True) As Boolean
    Dim checkResult As Boolean
    checkResult = True
    
   '�ʼ��Է� ����
    If isEssencial = True And (IsNull(argTxtbox.text) Or argTxtbox.text = "") Then
        MsgBox title & "��(��) �Է��ϼ���." & space(7), vbInformation, banner
        If isSetFocus = True Then
            argTxtbox.SetFocus
        End If
        argTxtbox = Empty
        checkResult = False
    End If
    
    '�������� ����: ����, ��¥
    If (dataType = "NUMERIC") Then
        If Not IsNumeric(argTxtbox.text) Then
            MsgBox title & "���� ���ڸ� �Է��� �� �ֽ��ϴ�." & space(7), vbInformation, banner
            If isSetFocus = True Then
                argTxtbox.SetFocus
            End If
            argTxtbox = Empty
            checkResult = False
        End If
    Else
        If (dataType = "DATE") Then
            If Not IsDate(argTxtbox.text) Then
                MsgBox title & "���� ��¥�� �Է��� �� �ֽ��ϴ�." & _
                vbNewLine & "YYYY-MM-DD ���·� �Է��ϼ���." & space(7), vbInformation, banner
                If isSetFocus = True Then
                    argTxtbox.SetFocus
                End If
                argTxtbox = Empty
                checkResult = False
            End If
        End If
    End If
    
     '�Է±��� ����
     If (Not IsNull(argTxtbox.text)) And Len(argTxtbox.text) > Length Then
        MsgBox title & "�� �Է� �ִ� ���̴� " & CStr(Length) & "�� ���� �� �����ϴ�." & space(7), vbInformation, banner
        If isSetFocus = True Then
            argTxtbox.SetFocus
        End If
        argTxtbox = Empty
        checkResult = False
    End If
    
    '�Լ������ True, False�� ��ȯ
    checkTextBox = checkResult
End Function

