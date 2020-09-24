Attribute VB_Name = "Functions"
Public Function SetMM() As Boolean

Dim MM As MINIMIZEDMETRICS
    SetMM = SystemParametersInfo(SPI_GETMINIMIZEDMETRICS, Len(MM), MM, SPIF_SENDWININICHANGE)
    MM.iWidth = 154
    MM.iHorzGap = 100
    MM.iVertGap = 100
    MM.iArrange = 8
    MM.cbSize = Len(MM)

    SetMM = SystemParametersInfo(SPI_SETMINIMIZEDMETRICS, Len(MM), MM, SPIF_SENDWININICHANGE)
 
End Function




Public Function SetHook(Action As HookAction) As Boolean
    
    
    Select Case Action
        Case Install
        
            uRegMsg = RegisterWindowMessage(ByVal "SHELLHOOK")
            Call RegisterShellHook(frmMain.hwnd, RSH_REGISTER_TASKMAN)
            lPrevProc = SetWindowLong(frmMain.hwnd, GWL_WNDPROC, AddressOf WndProc)
            SetHook = True
            
        Case Release
            Call RegisterShellHook(frmMain.hwnd, RSH_DEREGISTER)
            SetHook = SetWindowLong(frmMain.hwnd, GWL_WNDPROC, lPrevProc)
    End Select
    
End Function




Public Function WndProc(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    If wMsg = uRegMsg Then
        Select Case wParam
            Case CREATED
                Call DoTask(CREATED, lParam)
            Case DESTROYED
                Call DoTask(DESTROYED, lParam)
            Case ACTIVATED
                Call DoTask(ACTIVATED, lParam)
            Case REDRAW
                Call DoTask(REDRAW, lParam)
            Case Else
                WndProc = CallWindowProc(lPrevProc, hwnd, wMsg, wParam, lParam)
        End Select
        If Len(Status) > 0 Then frmMain.lstTasks.AddItem Status
    Else
        WndProc = CallWindowProc(lPrevProc, hwnd, wMsg, wParam, lParam)
    End If
End Function

Public Function DoTask(nTask As Long, hwnd As Long)
    Select Case nTask
        Case CREATED
            frmMain.WindowCreated hwnd
        Case DESTROYED
            frmMain.WindowDestroyed hwnd
        Case REDRAW
            frmMain.WindowRedraw hwnd
        Case ACTIVATED
            frmMain.WindowActivated hwnd
    End Select
End Function
