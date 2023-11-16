Attribute VB_Name = "fn_GetLocalIPaddress"
Option Explicit
Public Const WMISql As String = "SELECT IPAddress, IPSubnet, DefaultIPGateway FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True "

'------------------------------------------------------------------------------------------------------------------
'  PC의 IP주소를 구하는 함수
'    - GetLocalIPaddress()
'------------------------------------------------------------------------------------------------------------------
Public Function GetLocalIPaddress() As String
    Dim myWMI As Object
    Dim myItms As Object
    Dim myItm As Object
    
    Set myWMI = GetObject("winmgmts:\\" & Environ("ComputerName") & "\root\CIMV2")
    Set myItms = myWMI.ExecQuery(WMISql, , 48)
    
    For Each myItm In myItms
        GetLocalIPaddress = myItm.IPAddress(0)
        If Left(GetLocalIPaddress, 3) = 172 Then '가상IP대역 호출 방지
            Exit For
        End If
    Next
    
    Set myWMI = Nothing
    Set myItms = Nothing
    Set myItm = Nothing
End Function
