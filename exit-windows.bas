
Option Explicit
 
Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
 
Public Const EWX_FORCE = 4
Public Const EWX_LOGOFF = 0
Public Const EWX_REBOOT = 2
Public Const EWX_SHUTDOWN = 1
'Simple Program to shutdown windows
 
Public Sub Main()
	Dim Res As Long
	 
	'Warn User
	MsgBox ("Your System Will Now Shutdown")
	 
	'Call Function
	Res = ExitWindowsEx(EWX_SHUTDOWN, 0)
	 
	'No Need To Check For Success!!
End Sub