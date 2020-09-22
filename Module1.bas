Attribute VB_Name = "Module1"
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Global Const SWP_NOSIZE = &H1
Global Const SWP_NOMOVE = &H2
Global Const SWP_NOZORDER = &H4
Global Const SWP_NOREDRAW = &H8
Global Const SWP_NOACTIVATE = &H10
Global Const SWP_FRAMECHANGED = &H20
Global Const SWP_SHOWWINDOW = &H40
Global Const SWP_HIDEWINDOW = &H80
Global Const SWP_NOCOPYBITS = &H100
Global Const SWP_NOOWNERZORDER = &H200
Global Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Global Const SWP_NOREPOSITION = SWP_NOOWNERZORDER

Global Const HWND_TOP = 0
Global Const HWND_BOTTOM = 1
Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Const SPI_SCREENSAVERRUNNING = 97&
Const EWX_LOGOFF = 0
Const EWX_SHUTDOWN = 1
Const EWX_REBOOT = 2
Const EWX_FORCE = 4
Const READAPI = 0
Const WRITEAPI = 1
Const READ_WRITE = 2

Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long




Public Sub DisableKeys(dKeys As Boolean)
Dim lReturn As Long
Dim bPre As Boolean
lReturn = SystemParametersInfo(SPI_SCREENSAVERRUNNING, dKeys, bPre, 0&)

End Sub
Public Sub RestartWindows(Optional msgPrompt As Boolean = False, Optional ShowErrors As Boolean = False)
Dim Ret As Long
Ret& = ExitWindowsEx(EWX_SHUTDOWN, 0)
Exit Sub
End Sub
