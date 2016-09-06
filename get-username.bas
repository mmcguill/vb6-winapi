Option Explicit

Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private Sub Form_Load()
	Dim Result As Long
	Dim Username As String
	Dim LenUserName As Long

	LenUserName = 256
	Username = Space(LenUserName)
	Result = GetUserName(Username, LenUserName)

	If Result <> 0 Then
		Label1.Caption = Left$(Username, LenUserName)
	Else
		Label1.Caption = "Error"
	End If
End Sub
