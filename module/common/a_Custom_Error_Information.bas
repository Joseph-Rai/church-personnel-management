Attribute VB_Name = "a_Custom_Error_Information"
Option Explicit

'Custom Error Code & Description
Public Const ERR_CODE_WIFE_IS_ALREADY_REGISTERED_OF_OTHER = 1000
Public Const ERR_DESC_WIFE_IS_ALREADY_REGISTERED_OF_OTHER = "입력된 배우자는 이미 다른 관리자의 배우자로 등록되어 있습니다. 배우자 생명번호를 다시 확인해주세요." & vbNewLine

Public Const ERR_CODE_PSTAFF_IS_ALREADY_REGISTERED = 1001
Public Const ERR_DESC_PSTAFF_IS_ALREADY_REGISTERED = "입력된 선지자는 이미 등록되어 있습니다. 선지자의 생명번호를 다시 확인해주세요." & vbNewLine

'--//Error Description 사용 시 StringFormat 메서드를 이용할 것
'{0} - Invalid Control Name
Public Const ERR_CODE_INVALID_INPUT_PINFORMATION = 1002
Public Const ERR_DESC_INVALID_INPUT_PINFORMATION = "{0}이(가) 잘못되었습니다." & vbNewLine & "다시 확인해 주세요." & vbNewLine

'--//Error Description 사용 시 StringFormat 메서드를 이용할 것
'{0} - Required Control Name
Public Const ERR_CODE_REQUIRED_INPUT_PINFORMATION = 1003
Public Const ERR_DESC_REQUIRED_INPUT_PINFORMATION = "필수 입력값이 누락되었습니다." & vbNewLine & "{0}를(을) 다시 확인해 주세요."

Public Const ERR_CODE_TIME_OVERLAPPED = 1004
Public Const ERR_DESC_TIME_OVERLAPPED = "중복된 기간은 존재할 수 없습니다. 입력 값을 다시 확인해주세요."
