Attribute VB_Name = "sb_ControlEnabler"
Option Explicit

'@param control: 활성화 하고싶은 Control 객체
'@param blnEnabled: true: 활성화 / false: 비활성화
Public Sub ControlEnable(ByRef control As MSForms.control, blnEnabled As Boolean)
    
    control.Enabled = blnEnabled

End Sub

'@param textBox: 활성화 하고싶은 TextBox 객체
'@param blnEnabled: true: 활성화 / false: 비활성화
Public Sub TextBoxEnable(ByRef textBox As MSForms.textBox, blnEnabled As Boolean)

    If blnEnabled Then
        textBox.BackColor = RGB(255, 255, 255)
    Else
        textBox.BackColor = &HE0E0E0
    End If
    
    ControlEnable textBox, blnEnabled

End Sub

'@param comboBox: 활성화 하고싶은 ComboBox 객체
'@param blnEnabled: true: 활성화 / false: 비활성화
Public Sub ComboBoxEnable(ByRef argComboBox As MSForms.comboBox, blnEnabled As Boolean)

    If blnEnabled Then
        argComboBox.BackColor = RGB(255, 255, 255)
    Else
        argComboBox.BackColor = &HE0E0E0
    End If
    
    ControlEnable argComboBox, blnEnabled

End Sub

'@param argLabel: 색깔을 변경하고 싶은 Label 객체
'@param color: 필수값일 경우 vbRed / 보통의 경우 vbBlack
Public Sub ChangeLabelColor(ByRef argLabel As MSForms.label, color As Long)

    argLabel.ForeColor = color

End Sub

Public Function GetRequiredList() As Object

    Dim requiredList As Object
    Set requiredList = CreateObject("System.Collections.ArrayList")
    
    requiredList.Add "생명번호"
    requiredList.Add "생명번호"
    requiredList.Add "한글이름"
    requiredList.Add "영문이름"
    requiredList.Add "생년월일"
    requiredList.Add "국적"
    
    Set GetRequiredList = requiredList

End Function
