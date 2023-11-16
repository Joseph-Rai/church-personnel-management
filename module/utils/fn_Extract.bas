Attribute VB_Name = "fn_Extract"
'------------------------------------------------------------------------
' ���� : �������� ��� ���� 2016 VBA - ������ �Լ�&��ũ�� ���
'------------------------------------------------------------------------
Option Explicit
Option Compare Text
Option Base 1

'------------------------------------------------------------------------------------------
'   ��� :  ������
'            ����(����:N, ����), ������(����:E), �ѱ�(����:H), ��Ÿ����(����:O)
'            �� �����Ͽ� ��ȯ
'------------------------------------------------------------------------------------------
Function fnExtract(���ڿ� As String, Optional ���� As String = "N") As String
Attribute fnExtract.VB_Description = "�ؽ�Ʈ���� ����, ������, �ѱ�, Ư�����ڸ� �����ϴ� �Լ�"
Attribute fnExtract.VB_ProcData.VB_Invoke_Func = " \n17"
  Dim i As Integer
  Dim k As String
  '--// ����, ����, �ѱ�, ��Ÿ ���ڸ� ������ ����
  Dim NumStr As String, EngStr As String, HanStr As String, EtcStr As String  '��Ÿ ���ڵ��� �����
                                 
  Application.Volatile
  
  For i = 1 To Len(���ڿ�)
      k = Mid(���ڿ�, i, 1)
      Select Case k
         Case "0" To "9"
           NumStr = NumStr & k
         Case "."
           NumStr = NumStr & k
         Case "A" To "Z"
           EngStr = EngStr & k
         Case "a" To "z"
           EngStr = EngStr & k
         Case "��" To "�P"    '�ѱ��� '��'�� ���� �۰� '�P'�� ���� ū ����
           HanStr = HanStr & k
         Case Else
           EtcStr = EtcStr & k
      End Select
  Next
  
  Select Case ����
      Case "N":          fnExtract = NumStr
      Case "E":          fnExtract = EngStr
      Case "H":          fnExtract = HanStr
      Case "O":          fnExtract = EtcStr
      Case Else:         fnExtract = NumStr
  End Select
End Function


