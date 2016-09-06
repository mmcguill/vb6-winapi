
Option Explicit
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private Sub Form_Load()
	Dim Result As Long
	Dim ComputerName As String
	Dim LenComputerName As Long
	LenComputerName = 256
	Computername = Space(LenComputerName)
	Result = GetComputerName(ComputerName, LenComputerName)
	If Result <> 0 Then
		Label1.Caption = Left$(ComputerName, LenComputerName)
	Else
		Label1.Caption = "Error"
	End If
End Sub