Attribute VB_Name = "z"
Option Explicit

Public Const sCorelDRAW$ = "CorelDRAW"

#If VBA7 Then

Declare PtrSafe Function GetKeyState& Lib "user32" (ByVal vKey&)
Declare PtrSafe Function GetCursorPos& Lib "user32" (Page As wPOINT)
Declare PtrSafe Function CountClipboardFormats& Lib "user32" ()
Declare PtrSafe Function GetSystemMetrics& Lib "user32" (ByVal Index&)
Declare PtrSafe Function GetWindowRect& Lib "user32" (ByVal hWnd&, r As wRECT)
Declare PtrSafe Function GetClientRect& Lib "user32" (ByVal hWnd&, r As wRECT)
Declare PtrSafe Function FindWindowW& Lib "user32" (ByVal lpClassW As Any, ByVal lpTitleW As LongPtr)
Declare PtrSafe Function FindWindowExW& Lib "user32" (ByVal hParent&, ByVal hChildAfter&, ByVal lpClassW As Any, ByVal lpTitleW As LongPtr)
Declare PtrSafe Function SendMessageW& Lib "user32" (ByVal hWnd&, ByVal msg&, ByVal wParam&, ByVal lParam&)
Declare PtrSafe Function RedrawWindow& Lib "user32" (ByVal hWnd&, r As Any, ByVal rgn&, ByVal flags As wEnumRedrawWindow)
Declare PtrSafe Sub mouse_event Lib "user32" (ByVal dwFlags&, ByVal dx&, ByVal dy&, ByVal cButtons&, ByVal dwExtraInfo&)

Type wRECT: L As Long: T As Long: r As Long: b As Long: End Type
Type wPOINT: x As Long: y As Long: End Type
Public Const VK_SCROLL& = &H91
Enum wEnumRedrawWindow: RDW_INVALIDATE& = 1: RDW_INTERNALPAINT& = 2: RDW_ERASE& = 4: RDW_VALIDATE& = 8: RDW_NOINTERNALPAINT& = 16: RDW_NOERASE& = 32: _
              RDW_NOCHILDREN& = 64: RDW_ALLCHILDREN& = 128: RDW_UPDATENOW& = 256: RDW_ERASENOW& = 512: RDW_FRAME& = 1024: RDW_NOFRAME& = 2048: End Enum

Declare PtrSafe Function SetTimer& Lib "user32" (ByVal hWnd&, ByVal nIDEvent&, ByVal uElapse&, ByVal lpTimerFunc As LongPtr)
Declare PtrSafe Function KillTimer& Lib "user32" (ByVal hWnd&, ByVal uIDEvent&)

Private Declare PtrSafe Function LoadStringW& Lib "user32" (ByVal hInstance&, ByVal uID&, ByVal lpBufferW As Any, ByVal nBufferMax&)
Private Declare PtrSafe Function GetModuleHandleW& Lib "kernel32" (ByVal fileNameW As Any)

Private MsgBoxExBtnNames$, MsgBoxExAtCursor As Boolean, MsgBoxExCBTHook&
Private Declare PtrSafe Function SetWindowsHookExW& Lib "user32" (ByVal idHook&, ByVal lpfn As LongPtr, ByVal hmod&, ByVal dwThreadId&)
Private Declare PtrSafe Function CallNextHookEx& Lib "user32" (ByVal hhk&, ByVal nCode%, ByVal wParam&, ByVal lParam&)
Private Declare PtrSafe Function UnhookWindowsHookEx& Lib "user32" (ByVal m_zlhHook&)
Private Declare PtrSafe Function GetCurrentThreadId& Lib "kernel32" ()
Private Declare PtrSafe Function GetWindowLongW& Lib "user32" (ByVal hWnd&, ByVal nIndex%)
Private Declare PtrSafe Function GetForegroundWindow& Lib "user32" ()
Private Declare PtrSafe Function SetForegroundWindow& Lib "user32" (ByVal hWnd&)
Private Declare PtrSafe Function SetWindowTextW& Lib "user32" (ByVal hWnd&, ByVal lpTextW As Any)
Private Declare PtrSafe Function SetWindowPos& Lib "user32" (ByVal hWnd&, ByVal hAfter&, ByVal x&, ByVal y&, ByVal cx&, ByVal cy&, ByVal wFlags As wEnumSetWindowPos)
Private Enum wEnumSetWindowPos: _
    SWP_NOSIZE = 1: SWP_NOMOVE = 2: SWP_NOZORDER = 4: _
    SWP_NOREDRAW = 8: SWP_NOACTIVATE = 16: SWP_DRAWFRAME = 32: _
    SWP_SHOWWINDOW = 64: SWP_HIDEWINDOW = 128: SWP_NOCOPYBITS = 256: _
    SWP_NOREPOSITION = 512: SWP_NOSENDCHANGING = 1024: SWP_DEFERERASE = 8192: _
    SWP_ASYNCWINDOWPOS = 16384
End Enum
                              
#Else

Declare Function GetKeyState& Lib "user32" (ByVal vKey&)
Declare Function GetCursorPos& Lib "user32" (Page As wPOINT)
Declare Function CountClipboardFormats& Lib "user32" ()
Declare Function GetSystemMetrics& Lib "user32" (ByVal Index&)
Declare Function GetWindowRect& Lib "user32" (ByVal hWnd&, r As wRECT)
Declare Function GetClientRect& Lib "user32" (ByVal hWnd&, r As wRECT)
Declare Function FindWindowW& Lib "user32" (ByVal lpClassW As Any, ByVal lpTitleW&)
Declare Function FindWindowExW& Lib "user32" (ByVal hParent&, ByVal hChildAfter&, ByVal lpClassW As Any, ByVal lpTitleW&)
Declare Function SendMessageW& Lib "user32" (ByVal hWnd&, ByVal msg&, ByVal wParam&, ByVal lParam&)
Declare Function RedrawWindow& Lib "user32" (ByVal hWnd&, r As Any, ByVal rgn&, ByVal flags As wEnumRedrawWindow)
Declare Sub mouse_event Lib "user32" (ByVal dwFlags&, ByVal dx&, ByVal dy&, ByVal cButtons&, ByVal dwExtraInfo&)

Type wRECT: L As Long: T As Long: r As Long: b As Long: End Type
Type wPOINT: x As Long: y As Long: End Type
Public Const VK_SCROLL& = &H91
Enum wEnumRedrawWindow: RDW_INVALIDATE& = 1: RDW_INTERNALPAINT& = 2: RDW_ERASE& = 4: RDW_VALIDATE& = 8: RDW_NOINTERNALPAINT& = 16: RDW_NOERASE& = 32: _
              RDW_NOCHILDREN& = 64: RDW_ALLCHILDREN& = 128: RDW_UPDATENOW& = 256: RDW_ERASENOW& = 512: RDW_FRAME& = 1024: RDW_NOFRAME& = 2048: End Enum

Declare Function SetTimer& Lib "user32" (ByVal hWnd&, ByVal nIDEvent&, ByVal uElapse&, ByVal lpTimerFunc&)
Declare Function KillTimer& Lib "user32" (ByVal hWnd&, ByVal uIDEvent&)

Private Declare Function LoadStringW& Lib "user32" (ByVal hInstance&, ByVal uID&, ByVal lpBufferW As Any, ByVal nBufferMax&)
Private Declare Function GetModuleHandleW& Lib "kernel32" (ByVal fileNameW As Any)

Private MsgBoxExBtnNames$, MsgBoxExAtCursor As Boolean, MsgBoxExCBTHook&
Private Declare Function SetWindowsHookExW& Lib "user32" (ByVal idHook&, ByVal lpfn&, ByVal hmod&, ByVal dwThreadId&)
Private Declare Function CallNextHookEx& Lib "user32" (ByVal hhk&, ByVal nCode%, ByVal wParam&, ByVal lParam&)
Private Declare Function UnhookWindowsHookEx& Lib "user32" (ByVal m_zlhHook&)
Private Declare Function GetCurrentThreadId& Lib "kernel32" ()
Private Declare Function GetWindowLongW& Lib "user32" (ByVal hWnd&, ByVal nIndex%)
Private Declare Function GetForegroundWindow& Lib "user32" ()
Private Declare Function SetForegroundWindow& Lib "user32" (ByVal hWnd&)
Private Declare Function SetWindowTextW& Lib "user32" (ByVal hWnd&, ByVal lpTextW As Any)
Private Declare Function SetWindowPos& Lib "user32" (ByVal hWnd&, ByVal hAfter&, ByVal x&, ByVal y&, ByVal cx&, ByVal cy&, ByVal wFlags As wEnumSetWindowPos)
Private Enum wEnumSetWindowPos: _
    SWP_NOSIZE = 1: SWP_NOMOVE = 2: SWP_NOZORDER = 4: _
    SWP_NOREDRAW = 8: SWP_NOACTIVATE = 16: SWP_DRAWFRAME = 32: _
    SWP_SHOWWINDOW = 64: SWP_HIDEWINDOW = 128: SWP_NOCOPYBITS = 256: _
    SWP_NOREPOSITION = 512: SWP_NOSENDCHANGING = 1024: SWP_DEFERERASE = 8192: _
    SWP_ASYNCWINDOWPOS = 16384
End Enum

#End If

'===============================================================================

Sub TimerTransparentEdge(ByVal hWnd&, ByVal uMsg&, ByVal idEvent&, ByVal dwTime&)
    UITranspEdge.onTimer hWnd, idEvent
End Sub

Public Sub BoostStart( _
               Optional ByVal UndoGroupName As String = "", _
               Optional ByVal Optimize As Boolean = True _
           )
    If Not UndoGroupName = "" And Not ActiveDocument Is Nothing Then _
        ActiveDocument.BeginCommandGroup UndoGroupName
    If Optimize And Not Optimization Then Optimization = True
    If EventsEnabled Then EventsEnabled = False
    If Not ActiveDocument Is Nothing Then
        With ActiveDocument
            .SaveSettings
            .PreserveSelection = False
            .Unit = cdrMillimeter
            .WorldScale = 1
            .ReferencePoint = cdrCenter
        End With
    End If
End Sub
Public Sub BoostFinish(Optional ByVal EndUndoGroup As Boolean = True)
    If Not EventsEnabled Then EventsEnabled = True
    If Optimization Then Optimization = False
    If Not ActiveDocument Is Nothing Then
        With ActiveDocument
            .RestoreSettings
            .PreserveSelection = True
            If EndUndoGroup Then .EndCommandGroup
        End With
        ActiveWindow.Refresh
    End If
    Application.Windows.Refresh
End Sub

Function UIrefresh&()
    SendMessageW AppWindow.Handle, &H111, &H19D09, 0 'c109a053-50ab-4943-8edc-d374ec153e7e
    SendMessageW AppWindow.Handle, &H111, &H19B78, 0 'f1aee54d-c9aa-4e6f-9193-82f496b0b72b
End Function

Function UIPositionFromRegistry(Form As Object, sUIname$) As Long 'returns handle
    Dim s$, r As wRECT
    On Error Resume Next
    s = Form.Caption: Form.Caption = Form.Caption & Hex$(Timer)
    UIPositionFromRegistry = FindWindowW(StrPtr("ThunderDFrame"), StrPtr(Form.Caption))
    GetWindowRect UIPositionFromRegistry, r
    Form.Caption = s
    
    s = Trim$(GetSetting(sCorelDRAW, sUIname, "WindowPos"))
    If Len(s) = 0 Then Exit Function
    Form.StartUpPosition = 0
    Form.Move Split(s, " ")(0), Split(s, " ")(1)
    If r.L < 0 Or r.L > GetSystemMetrics(0) - 30 Or _
       r.T < 0 Or r.T > GetSystemMetrics(1) - 100 Then _
          Form.StartUpPosition = 1
End Function

Function getCMSprofiles$()
    Dim s$, ss$, i&, o As Object, cms$, hFile&
    On Error Resume Next
    If VersionMajor < 13 Then
       s = Left$(UserDataPath, Len(UserDataPath) - 1)
       s = Left$(s, InStrRev(s, "\"))
       ss = s & "User Workspace\CorelDRAW" & VersionMajor & "\" & ActiveWorkspace.Name & "\CorelDRAW.ini"
       hFile = FreeFile(): Err.Clear: Open ss For Binary Access Read Shared As #hFile
       If Err.Number Then Exit Function
       ss = Space$(FileLen(ss)): Get #hFile, , ss: Close #hFile
       ss = Replace$(Replace$(ss, vbCrLf, vbCr), vbLf, vbCr)
       i = InStr(1, ss, "CurrentCMStyle", vbTextCompare): If i = 0 Then Exit Function
       ss = Mid$(ss, i + Len("CurrentCMStyle")): ss = Mid$(ss, InStr(ss, "=") + 1)
       i = InStr(ss, vbCr): cms = Left$(ss, IIf(i, i - 1, 999))
       
       s = s & "User Config\color.ini"
       hFile = FreeFile(): Err.Clear: Open s For Binary Access Read Shared As #hFile
       If Err.Number Then Exit Function
       s = Space$(FileLen(s)): Get #hFile, , s: Close #hFile
       s = Replace$(Replace$(s, vbCrLf, vbCr), vbLf, vbCr)
       
       i = InStr(1, s, "[CMStyle" + cms + "]", vbTextCompare): If i = 0 Then Exit Function
       s = Mid$(s, i + 1): i = InStr(1, s, "[CMStyle", vbTextCompare): If i Then s = Left$(s, i - 1)
       
       i = InStr(1, s, "IntRGBProfileName", vbTextCompare)
       If i Then
          ss = Mid$(s, i + Len("IntRGBProfileName")): ss = Mid$(ss, InStr(ss, "=") + 1)
          i = InStr(ss, vbCr): ss = Left$(ss, IIf(i, i - 1, 999))
          End If
       i = InStr(1, s, "SepsPrnProfileName", vbTextCompare)
       If i Then
          s = Mid$(s, i + Len("SepsPrnProfileName")): s = Mid$(s, InStr(s, "=") + 1)
          i = InStr(s, vbCr): s = Left$(s, IIf(i, i - 1, 999))
       End If
    Else
       Set o = Application
       ss = o.ColorManager.CurrentProfile(4) 'clrInternalRGB=4
       s = o.ColorManager.CurrentProfile(2) 'clrSeparationPrinter = 2
    End If
    getCMSprofiles = "RGB:" & vbTab & ss & vbCr & "CMYK:" & vbTab & s
End Function

Function sLayerDesktop$()
    Static s$
    If Len(s) = 0 Then s = VGDllString(6524): If Len(s) = 0 Then s = "Desktop"
    sLayerDesktop = s
End Function

Function sLayerGuides$()
    Static s$
    If Len(s) = 0 Then s = VGDllString(6530): If Len(s) = 0 Then s = "Guides"
    sLayerGuides = s
End Function

Function VGDllString$(ByVal nID&)
    Static hVGdll&: Const MAXBUF& = 1024: Dim i&
    If hVGdll = 0 Then hVGdll = GetModuleHandleW(StrPtr(IIf(VersionMajor > 11, "VGCoreIntl.dll", "DrawIntl.dll")))
    VGDllString = String$(MAXBUF, vbNullChar)
    i = LoadStringW(hVGdll, nID, StrPtr(VGDllString), MAXBUF)
    VGDllString = Left$(VGDllString, i)
End Function
Function vMaxL _
    (ByVal a&, _
     ByVal b&)
                If a > b Then vMaxL = a Else vMaxL = b
End Function

Function vMinL _
    (ByVal a&, _
     ByVal b&)
                If a < b Then vMinL = a Else vMinL = b
End Function

Function vSqueezeL _
    (ByVal v&, _
     ByVal min&, _
     ByVal max&)
                If v < min Then vSqueezeL = min _
                           Else If v > max Then vSqueezeL = max _
                                           Else vSqueezeL = v
End Function

Sub wmSetRedraw(ByVal state%, Optional ByVal hWnd&)
    SendMessageW IIf(hWnd, hWnd, AppWindow.Handle), 11, state, 0 'WM_SETREDRAW = 11
End Sub

Function UnitName$(ByVal Unit)
    If Unit > 0 And Unit <= 16 Then _
       UnitName = Choose(Unit, "in", "ft", "mm", "cm", "px", "mil", "m", "km", "did", "ag", "y", "pi", "cic", "pt", "uQ", "uH")
End Function

Function MsgBoxEx _
    (sPrompt$, _
     Optional ByVal Buttons As VbMsgBoxStyle, _
     Optional sTitle$, _
     Optional sBtnNames$, _
     Optional ByVal bAtCursor As Boolean = True) As VbMsgBoxResult
    
    Const GWL_HINSTANCE = -6, WH_CBT = 5
    Dim hwndPrev&
    MsgBoxExCBTHook = SetWindowsHookExW(WH_CBT, AddressOf msgboxExProc, _
                                        GetWindowLongW(AppWindow.Handle, GWL_HINSTANCE), GetCurrentThreadId())
    MsgBoxExAtCursor = bAtCursor
    MsgBoxExBtnNames = sBtnNames
    hwndPrev = GetForegroundWindow()
    MsgBoxEx = MsgBox(sPrompt, Buttons, sTitle)
    UnhookWindowsHookEx MsgBoxExCBTHook
    SetForegroundWindow hwndPrev
End Function

Private _
Function msgboxExProc& _
    (ByVal Lmsg&, ByVal wParam&, ByVal lParam&)
    Const HCBT_CREATEWND = 3, HCBT_ACTIVATE = 5
    Static hwndMsg&
    'Debug.Print Hex$(Lmsg) & "  " & Hex$(wParam)
    Select Case Lmsg
       Case Is < 0: msgboxExProc = CallNextHookEx(MsgBoxExCBTHook, Lmsg, wParam, lParam): Exit Function
       Case HCBT_CREATEWND:
             If hwndMsg = 0 Then hwndMsg = wParam: Exit Function
       Case HCBT_ACTIVATE
             If hwndMsg <> wParam Then Exit Function
             UnhookWindowsHookEx MsgBoxExCBTHook
             If MsgBoxExAtCursor Then
                Dim m As wPOINT, cp As wPOINT
                GetCursorPos cp: m.x = cp.x - 50: m.y = cp.y - 100
                SetWindowPos hwndMsg, 0, m.x, m.y, 0, 0, SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE
             End If
             If Len(MsgBoxExBtnNames) Then
                Dim h&, btns$(), i&
                btns = Split(MsgBoxExBtnNames, "|")
                Do
                   h = FindWindowExW(hwndMsg, h, StrPtr("Button"), 0): If h = 0 Then Exit Do
                   SetWindowTextW h, StrPtr(btns(i))
                   i = i + 1
                Loop While i <= UBound(btns)
             End If
             hwndMsg = 0
    End Select
End Function
