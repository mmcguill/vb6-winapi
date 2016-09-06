
Option Explicit
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

Private Sub Command1_Click()
	Dim Result As Long
	Dim x As Long
	Dim y As Long
	x = Val(Text1.Text)
	y = Val(Text2.Text)
	Result = WindowFromPoint(x, y)
	If Result <> 0 Then
		MsgBox Result, 0, "Window Handle"
	Else
		MsgBox "Error", 48, "Window Handle"
	End If
End Sub