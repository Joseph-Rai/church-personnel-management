Attribute VB_Name = "z_MouseScroll"
''''' How to use : 'HookListBoxScroll Me, Me.LstBox
''''' 'Un'HookListBoxScroll when listbox Exit

'''''' normal module code

Option Explicit

Private Type POINTAPI
        x As Long
        y As Long
End Type

Private Type MOUSEHOOKSTRUCT
        pt As POINTAPI
        hWnd As Long
        wHitTestCode As Long
        dwExtraInfo As Long
End Type

#If VBA7 Then
    Private Declare PtrSafe Function FindWindow Lib "user32" _
        Alias "FindWindowA" ( _
        ByVal lpClassName As String, _
        ByVal lpWindowName As String) As LongPtr
    
    Private Declare PtrSafe Function GetWindowLong Lib "user32.dll" _
        Alias "GetWindowLongPtrA" ( _
        ByVal hWnd As LongPtr, _
        ByVal nIndex As Long) As LongPtr
    
    Private Declare PtrSafe Function SetWindowsHookEx Lib "user32" _
        Alias "SetWindowsHookExA" ( _
        ByVal idHook As Long, _
        ByVal lpfn As LongPtr, _
        ByVal hmod As LongPtr, _
        ByVal dwThreadId As Long) As LongPtr
    
    Private Declare PtrSafe Function CallNextHookEx Lib "user32" ( _
        ByVal hHook As LongPtr, _
        ByVal nCode As Long, _
        ByVal wParam As LongPtr, _
        lParam As Any) As LongPtr
    
    Private Declare PtrSafe Function UnhookWindowsHookEx Lib "user32" ( _
        ByVal hHook As LongPtr) As LongPtr
    
    Private Declare PtrSafe Function WindowFromPoint Lib "user32" ( _
        ByVal xPoint As Long, _
        ByVal yPoint As Long) As LongPtr
    
    Private Declare PtrSafe Function GetCursorPos Lib "user32.dll" ( _
        ByRef lpPoint As POINTAPI) As LongPtr
#Else
    Private Declare Function FindWindow Lib "User32" _
                                            Alias "FindWindowA" ( _
                                                            ByVal lpClassName As String, _
                                                            ByVal lpWindowName As String) As Long
    
    Private Declare Function GetWindowLong Lib "user32.dll" _
                                            Alias "GetWindowLongA" ( _
                                                            ByVal hwnd As Long, _
                                                            ByVal nIndex As Long) As Long
    
    Private Declare Function SetWindowsHookEx Lib "User32" _
                                            Alias "SetWindowsHookExA" ( _
                                                            ByVal idHook As Long, _
                                                            ByVal lpfn As Long, _
                                                            ByVal hmod As Long, _
                                                            ByVal dwThreadId As Long) As Long
    
    Private Declare Function CallNextHookEx Lib "User32" ( _
                                                            ByVal hHook As Long, _
                                                            ByVal nCode As Long, _
                                                            ByVal wParam As Long, _
                                                            lParam As Any) As Long
    
    Private Declare Function UnhookWindowsHookEx Lib "User32" ( _
                                                            ByVal hHook As Long) As Long
    
    Private Declare Function PostMessage Lib "user32.dll" _
                                             Alias "PostMessageA" ( _
                                                             ByVal hwnd As Long, _
                                                             ByVal wMsg As Long, _
                                                             ByVal wParam As Long, _
                                                             ByVal lParam As Long) As Long
    
    Private Declare Function WindowFromPoint Lib "User32" ( _
                                                            ByVal xPoint As Long, _
                                                            ByVal yPoint As Long) As Long
    
    Private Declare Function GetCursorPos Lib "user32.dll" ( _
                                                            ByRef lpPoint As POINTAPI) As Long
#End If

Private Const WH_MOUSE_LL As Long = 14
Private Const WM_MOUSEWHEEL As Long = &H20A
Private Const HC_ACTION As Long = 0
Private Const GWL_HINSTANCE As Long = (-6)

'Private Const WM_KEYDOWN As long = &H100
'Private Const WM_KEYUP As long = &H101
'Private Const VK_UP As long = &H26
'Private Const VK_DOWN As long = &H28
'Private Const WM_LBUTTONDOWN As long = &H201

#If VBA7 Then
    Private mListBoxHwnd As LongPtr
    Private mLngMouseHook As LongPtr
#Else
    Private mListBoxHwnd As Long
    Private mLngMouseHook As Long
#End If
Private mbHook As Boolean
Private mCtl As MSForms.control
Dim n As Long

Sub HookListBoxScroll(frm As Object, ctl As MSForms.control)
#If VBA7 Then
    Dim hwndUnderCursor As LongPtr
    Dim lngAppInst As LongPtr
#Else
    Dim hwndUnderCursor As Long
    Dim lngAppInst As Long
#End If
Dim tPT As POINTAPI
     GetCursorPos tPT
     hwndUnderCursor = WindowFromPoint(tPT.x, tPT.y)
     If Not frm.ActiveControl Is ctl Then
             ctl.SetFocus
     End If
     If mListBoxHwnd <> hwndUnderCursor Then
             'Un'HookListBoxScroll
             Set mCtl = ctl
             mListBoxHwnd = hwndUnderCursor
             lngAppInst = GetWindowLong(mListBoxHwnd, GWL_HINSTANCE)
             ' PostMessage mListBoxHwnd, WM_LBUTTONDOWN, 0&, 0&
             If Not mbHook Then
                     mLngMouseHook = SetWindowsHookEx( _
                                                     WH_MOUSE_LL, AddressOf MouseProc, lngAppInst, 0)
                     mbHook = mLngMouseHook <> 0
             End If
     End If
End Sub

Sub UnHookListBoxScroll()
     If mbHook Then
                Set mCtl = Nothing
             UnhookWindowsHookEx mLngMouseHook
             mLngMouseHook = 0
             mListBoxHwnd = 0
             mbHook = False
        End If
End Sub

#If VBA7 Then
Private Function MouseProc( _
             ByVal nCode As Long, ByVal wParam As Long, _
             ByRef lParam As MOUSEHOOKSTRUCT) As LongPtr
#Else
Private Function MouseProc( _
             ByVal nCode As Long, ByVal wParam As Long, _
             ByRef lParam As MOUSEHOOKSTRUCT) As Long
#End If
Dim idx As Long
        On Error GoTo errH
     If (nCode = HC_ACTION) Then
             If WindowFromPoint(lParam.pt.x, lParam.pt.y) = mListBoxHwnd Then
                     If wParam = WM_MOUSEWHEEL Then
                                MouseProc = True
                                If lParam.hWnd > 0 Then idx = -1 Else idx = 1
                             idx = idx + mCtl.TopIndex
                             If idx >= 0 Then mCtl.TopIndex = idx
                                Exit Function
                     End If
             Else
                     'Un'HookListBoxScroll
             End If
     End If
     MouseProc = CallNextHookEx( _
                             mLngMouseHook, nCode, wParam, ByVal lParam)
     Exit Function
errH:
     'Un'HookListBoxScroll
End Function
'''''''' end normal module code





