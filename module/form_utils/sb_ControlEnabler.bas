Attribute VB_Name = "sb_ControlEnabler"
Option Explicit

'@param control: Ȱ��ȭ �ϰ���� Control ��ü
'@param blnEnabled: true: Ȱ��ȭ / false: ��Ȱ��ȭ
Public Sub ControlEnable(ByRef control As MSForms.control, blnEnabled As Boolean)
    
    control.Enabled = blnEnabled

End Sub

'@param textBox: Ȱ��ȭ �ϰ���� TextBox ��ü
'@param blnEnabled: true: Ȱ��ȭ / false: ��Ȱ��ȭ
Public Sub TextBoxEnable(ByRef textBox As MSForms.textBox, blnEnabled As Boolean)

    If blnEnabled Then
        textBox.BackColor = RGB(255, 255, 255)
    Else
        textBox.BackColor = &HE0E0E0
    End If
    
    ControlEnable textBox, blnEnabled

End Sub

'@param comboBox: Ȱ��ȭ �ϰ���� ComboBox ��ü
'@param blnEnabled: true: Ȱ��ȭ / false: ��Ȱ��ȭ
Public Sub ComboBoxEnable(ByRef argComboBox As MSForms.comboBox, blnEnabled As Boolean)

    If blnEnabled Then
        argComboBox.BackColor = RGB(255, 255, 255)
    Else
        argComboBox.BackColor = &HE0E0E0
    End If
    
    ControlEnable argComboBox, blnEnabled

End Sub

'@param argLabel: ������ �����ϰ� ���� Label ��ü
'@param color: �ʼ����� ��� vbRed / ������ ��� vbBlack
Public Sub ChangeLabelColor(ByRef argLabel As MSForms.label, color As Long)

    argLabel.ForeColor = color

End Sub

Public Function GetRequiredList() As Object

    Dim requiredList As Object
    Set requiredList = CreateObject("System.Collections.ArrayList")
    
    requiredList.Add "�����ȣ"
    requiredList.Add "�����ȣ"
    requiredList.Add "�ѱ��̸�"
    requiredList.Add "�����̸�"
    requiredList.Add "�������"
    requiredList.Add "����"
    
    Set GetRequiredList = requiredList

End Function
