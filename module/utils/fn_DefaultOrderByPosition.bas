Attribute VB_Name = "fn_DefaultOrderByPosition"
Option Explicit

Public Function GetDefaultOrderByPosition() As String

    GetDefaultOrderByPosition = _
        "FIELD(`��å`,'��ȸ��','��ȸ��븮','����','��븮���','����','�����','����ȸ������','�����ڻ��','����Ұ�����','�����ڻ��','�������1�ܰ�','�������2�ܰ�','�������3�ܰ�','�������'," & getPosition2Joining & ",'')"

End Function
