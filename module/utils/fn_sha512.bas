Attribute VB_Name = "fn_sha512"
Option Explicit

Public Function to_SHA512(mypassword As String) As String
    'Requires a reference to mscorlib 4.0 64-bit
    Dim text As Object
    Dim SHA512 As Object
    
    Set text = CreateObject("System.Text.UTF8Encoding")
    Set SHA512 = CreateObject("System.Security.Cryptography.SHA512Managed")
    'Set SHA256 = CreateObject("System.Security.Cryptography.SHA256Managed")
    'Debug.Print ToBase64String(SHA512.ComputeHash_2((text.GetBytes_4("mypassword"))))
    to_SHA512 = ToHexString(SHA512.ComputeHash_2((text.GetBytes_4(mypassword))))
End Function

Public Function ToHexString(rabyt)

  'Ref: http://stackoverflow.com/questions/1118947/converting-binary-file-to-base64-string
  With CreateObject("MSXML2.DOMDocument")
    .LoadXML "<root />"
    .DocumentElement.dataType = "bin.Hex"
    .DocumentElement.nodeTypedValue = rabyt
    ToHexString = Replace(.DocumentElement.text, vbLf, "")
  End With
  
End Function

