Attribute VB_Name = "sb_embed_pics_permanently"
Option Explicit

Public Sub embed_Pics_Permanently(shpC As Shape)
  
    Dim picLeft As Single                                   '�� �׸�(����)�� ���� ��ġ ���� ����
    Dim picTop As Single                                   '�� �׸��� ���� ��ġ ���� ����
           
    picLeft = shpC.Left                       '��ũ�� �׸��� ���� ��ġ�� ������
    picTop = shpC.Top                      '��ũ�� �׸��� ���� ��ġ�� ������
    
    shpC.CopyPicture Format:=xlBitmap                                 '��ũ�� �׸��� ���� Appearance:=xlPrinter,
'    ActiveSheet.PasteSpecial Link:=False    '�׸� ��ũ���� �ٿ��ֱ�

    ActiveSheet.Paste
    shpC.Delete                                '��ũ�� �׸��� ����
    Selection.Left = picLeft                 '����� �׸� ���� ��ġ�� ��ũ�� �׸� ������ġ��
    Selection.Top = picTop                 '����� �׸� ���� ��ġ�� ��ũ�� �׸� ������ġ��
    
'On Error Resume Next
'    Dim cnt As Integer
'RETRY:
'    ActiveSheet.Paste
'    If err.Number = 0 Then
'        shpC.Delete                                '��ũ�� �׸��� ����
'        Selection.Left = picLeft                 '����� �׸� ���� ��ġ�� ��ũ�� �׸� ������ġ��
'        Selection.Top = picTop                 '����� �׸� ���� ��ġ�� ��ũ�� �׸� ������ġ��
'    Else
'        cnt = cnt + 1
'        If cnt < 100 Then
'            err.Number = 0
'            GoTo RETRY
'        End If
'    End If
'On Error GoTo 0
   
End Sub
