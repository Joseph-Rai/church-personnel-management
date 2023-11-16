Attribute VB_Name = "fn_Base64"
Option Explicit

Public Function encodeBase64(Bytes) As String
    Dim DM, EL
    Set DM = CreateObject("Microsoft.XMLDOM")
    ' Create temporary node with Base64 data type
    Set EL = DM.createElement("tmp")
    EL.dataType = "bin.base64"
    ' Set bytes, get encoded String
    ' Convert byte string to byte array
    EL.nodeTypedValue = Bytes
    encodeBase64 = EL.text
End Function
  
Public Function decodeBase64(base64 As String)
    Dim DM, EL
    Set DM = CreateObject("Microsoft.XMLDOM")
    ' Create temporary node with Base64 data type
    Set EL = DM.createElement("tmp")
    EL.dataType = "bin.base64"
    ' Set encoded String, get bytes
    EL.text = base64
    decodeBase64 = EL.nodeTypedValue
End Function


