Attribute VB_Name = "a_Custom_Error_Information"
Option Explicit

'Custom Error Code & Description
Public Const ERR_CODE_WIFE_IS_ALREADY_REGISTERED_OF_OTHER = 1000
Public Const ERR_DESC_WIFE_IS_ALREADY_REGISTERED_OF_OTHER = "�Էµ� ����ڴ� �̹� �ٸ� �������� ����ڷ� ��ϵǾ� �ֽ��ϴ�. ����� �����ȣ�� �ٽ� Ȯ�����ּ���." & vbNewLine

Public Const ERR_CODE_PSTAFF_IS_ALREADY_REGISTERED = 1001
Public Const ERR_DESC_PSTAFF_IS_ALREADY_REGISTERED = "�Էµ� �����ڴ� �̹� ��ϵǾ� �ֽ��ϴ�. �������� �����ȣ�� �ٽ� Ȯ�����ּ���." & vbNewLine

'--//Error Description ��� �� StringFormat �޼��带 �̿��� ��
'{0} - Invalid Control Name
Public Const ERR_CODE_INVALID_INPUT_PINFORMATION = 1002
Public Const ERR_DESC_INVALID_INPUT_PINFORMATION = "{0}��(��) �߸��Ǿ����ϴ�." & vbNewLine & "�ٽ� Ȯ���� �ּ���." & vbNewLine

'--//Error Description ��� �� StringFormat �޼��带 �̿��� ��
'{0} - Required Control Name
Public Const ERR_CODE_REQUIRED_INPUT_PINFORMATION = 1003
Public Const ERR_DESC_REQUIRED_INPUT_PINFORMATION = "�ʼ� �Է°��� �����Ǿ����ϴ�." & vbNewLine & "{0}��(��) �ٽ� Ȯ���� �ּ���."

Public Const ERR_CODE_TIME_OVERLAPPED = 1004
Public Const ERR_DESC_TIME_OVERLAPPED = "�ߺ��� �Ⱓ�� ������ �� �����ϴ�. �Է� ���� �ٽ� Ȯ�����ּ���."
