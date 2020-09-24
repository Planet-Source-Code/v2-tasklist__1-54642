Attribute VB_Name = "Declarations"
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Public Declare Function GetProcAddress Lib "kernel32.dll" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Public Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
Public Declare Function SetWindowsHookEx Lib "user32.dll" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Function UnhookWindowsHookEx Lib "user32.dll" (ByVal hHook As Long) As Long
Public Declare Function CallNextHookEx Lib "user32.dll" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SystemParametersInfo Lib "user32.dll" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long
Public Declare Function RegisterShellHook Lib "Shell32" Alias "#181" (ByVal hwnd As Long, ByVal nAction As Long) As Long
Public Declare Function IsWindowVisible Lib "user32.dll" (ByVal hwnd As Long) As Long
Public Declare Function GetParent Lib "user32.dll" (ByVal hwnd As Long) As Long
Public Declare Function GetWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

Public Const GWL_WNDPROC As Long = -4
Public Const SPI_GETMINIMIZEDMETRICS As Long = 43
Public Const SPI_SETMINIMIZEDMETRICS As Long = 44
Public Const SPIF_SENDWININICHANGE As Long = &H2

Public Const WH_CBT As Long = 5
Public Const WH_SHELL As Long = 10

Public Const REDRAW As Long = 6
Public Const ACTIVATED As Long = 4
Public Const CREATED As Long = 1
Public Const DESTROYED As Long = 2

Public Const RSH_DEREGISTER = 0
Public Const RSH_REGISTER = 1
Public Const RSH_REGISTER_PROGMAN = 2
Public Const RSH_REGISTER_TASKMAN = 3

Public Const GW_OWNER As Long = 4
Public Const GWL_EXSTYLE As Long = -20
Public Const WS_EX_TOOLWINDOW As Long = &H80&
Public Const WS_EX_APPWINDOW As Long = &H40000


Public Type MINIMIZEDMETRICS
    cbSize As Long
    iWidth As Long
    iHorzGap As Long
    iVertGap As Long
    iArrange As Long
End Type

Enum HookAction
    Install = 1
    Release = 0
End Enum
    

Public lHook As Long
Public lPrevProc As Long
Public lModule As Long
Public lProc As Long
Public uRegMsg As Long
