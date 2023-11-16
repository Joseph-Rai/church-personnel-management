Attribute VB_Name = "fn_inputboxPW"
' Module level variable for holding the hook
  Private InputBoxHook As Long
' API Functions
  #If VBA7 And Win64 Then
      Private Declare PtrSafe Function GetCurrentThreadId Lib "kernel32" () As Long
      Private Declare PtrSafe Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
      Private Declare PtrSafe Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
      Private Declare PtrSafe Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long
      Private Declare PtrSafe Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
      Private Declare PtrSafe Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
      Private Declare PtrSafe Function SendDlgItemMessage Lib "user32" Alias "SendDlgItemMessageA" (ByVal hDlg As Long, ByVal nIDDlgItem As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  #Else
      Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
      Private Declare Function UnhookWindowsHookEx Lib "User32" (ByVal hHook As Long) As Long
      Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
      Private Declare Function CallNextHookEx Lib "User32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long
      Private Declare Function GetClassName Lib "User32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
      Private Declare Function SetWindowsHookEx Lib "User32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
      Private Declare Function SendDlgItemMessage Lib "User32" Alias "SendDlgItemMessageA" (ByVal hDlg As Long, ByVal nIDDlgItem As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  #End If
Option Explicit

Function InputBoxPW(ByVal IB_Prompt As String, _
           Optional ByVal IB_Title As String = "Microsoft Excel", _
           Optional ByVal IB_Default As String = vbNullString) As String
' This function will request an input from the user using a hooked InputBox.
' The hooked InputBox is hooked to show all chars as asterisk (*),
' thus the hooked InputBox is very suitable for requesting confidential information, like passwords.
'
' As arguments, this function accepts the same first three arguments, as the normal InputBox function (Prompt, Title, Default Value),
' but you can just expand it to also accept the rest of the arguments, if you want to (after all it is just the normal InputBox function).

' * ' Initialize
      On Error Resume Next

' * ' Define variables
      Dim ThreadID As Long
      ThreadID = GetCurrentThreadId                                                             ' API function call

      Dim ModuleHandle As Long
      ModuleHandle = GetModuleHandle(vbNullString)                                              ' API function call

' * ' Set the hook
      InputBoxHook = SetWindowsHookEx(5, AddressOf InputBoxPW_Hook, ModuleHandle, ThreadID)     ' Assign hook handle/pointer to variable

' * ' Request masked input, using the hooked InputBox function
      InputBoxPW = InputBox(IB_Prompt, " " & WorksheetFunction.Trim(IB_Title), IB_Default)      ' Use hooked InputBox to request masked input

EF: ' End of Function
      UnhookWindowsHookEx InputBoxHook                                                          ' Release the hook
End Function

Private Function InputBoxPW_Hook(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
' This private function is the hook sub-function of the above 'InputBoxPW()' function.
' The hook is simply set to this function.
'
' Arguments is from the API function, 'CallNextHookEx' (https://msdn.microsoft.com/en-us/library/windows/desktop/ms644974(v=vs.85).aspx).

' * ' Initialize
      Const PasswordChar As String = "*"                                                        ' You can set this to another char, if you want to

      On Error Resume Next

' * ' Define variables
      Dim ClassName As String
      ClassName = String$(256, " ")
      GetClassName wParam, ClassName, 255

' * ' Hook the InputBox
      If nCode < 0 Then
            InputBoxPW_Hook = CallNextHookEx(InputBoxHook, nCode, wParam, lParam)
            Exit Function
      ElseIf nCode = 5 And Left$(ClassName, 6) = "#32770" Then                                  ' A window with the class name of InputBox has been activated
            SendDlgItemMessage wParam, 4900, 204, Asc(PasswordChar), 0
      End If

' * ' Make sure that any other hooks that may be in place are called correctly
      CallNextHookEx InputBoxHook, nCode, wParam, lParam
End Function

Sub InputBoxPW_Test()
' * ' Initialize
      On Error Resume Next

' * ' Define variable
      Dim mypassword As String
      mypassword = InputBoxPW(" Enter your password:")

' * ' Display result
      If StrPtr(mypassword) = 0 Then
            MsgBox "User pressed [Cancel]", vbOKOnly + vbCritical
      ElseIf Len(mypassword) < 1 Then
            MsgBox "Nothing was entered!", vbOKOnly + vbExclamation
      Else
            MsgBox mypassword, vbOKOnly + vbInformation
      End If
End Sub


