'Require Variable Declaration
Option Explicit
 
'Registry types
Public Const REG_BINARY As Long = 3
Public Const REG_DWORD As Long = 4
Public Const REG_NONE As Long = 0
Public Const REG_LINK As Long = 6
Public Const REG_MULTI_SZ As Long = 7
Public Const REG_SZ As Long = 1
 
'Registry Main Keys
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_DYN_DATA = &H80000006
 
'Registry errors
Public Const ERROR_BADDB = 1
Public Const ERROR_BADKEY = 2
Public Const ERROR_CANTOPEN = 3
Public Const ERROR_CANTREAD = 4
Public Const ERROR_CANTWRITE = 5
Public Const ERROR_OUTOFMEMORY = 6
Public Const ERROR_INVALID_PARAMETER = 7
Public Const ERROR_ACCESS_DENIED = 8
Public Const ERROR_INVALID_PARAMETERS = 87
Public Const ERROR_NO_MORE_ITEMS = 259
Public Const ERROR_MORE_DATA = 234
Public Const ERROR_SUCCESS = 0&
 
'Registry access
Public Const KEY_CREATE_LINK = &H20
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const SYNCHRONIZE = &H100000
Public Const STANDARD_RIGHTS_ALL = &H1F0000
Public Const READ_CONTROL = &H20000
Public Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Public Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Public Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Public Const KEY_EXECUTE = (KEY_READ)
Public Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
Public Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
 
'Registry Time needed for RegEnumKeyEx
Type FILETIME
dwLowDateTime As Long
dwHighDateTime As Long
End Type
'Registry Options and Status needed for RegCreateKeyEx
Public Const REG_OPTION_NON_VOLATILE = 0
Public Const REG_OPTION_VOLATILE = 1
Public Const REG_CREATED_NEW_KEY = &H1
Public Const REG_OPENED_EXISTING_KEY = &H2
 
'The Maximum data length
Public Const MAX_LENGTH As Long = 2048
 
'WIN32 Registry Function Declarations
Declare Function RegCloseKey Lib "advapi32.dll" Alias "RegCloseKey" (ByVal hKey As Long) As Long
Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long
Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" (ByVal hKey As Long, ByVal lpClass As String, lpcbClass As Long, lpReserved As Long, lpcSubKeys As Long, lpcbMaxSubKeyLen As Long, lpcbMaxClassLen As Long, lpcValues As Long, lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long, lpcbSecurityDescriptor As Long, lpftLastWriteTime As FILETIME) As Long
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long

Public Function MyRegCloseKey(hKey As Long)
	'This is the return Variable
	Dim res As Long
	 
	'Here we call the Function
	res = RegCloseKey(hKey)
	 
	'Here we check for an error
	if res<>ERROR_SUCCESS then
		MsgBox("Error")
	else
		MsgBox("Success")
End if
End Function


Public Function MyRegCreateKeyEx(hKey as long, SubKey As String) As Long
	'This Function takes in one string - "SubKey" and creates under the given Key
	 
	'An example of what SubKey should look like:
	'"Sublevel1\Sublevel2" etc.
	 
	'This variable takes the result of the function
	Dim res As Long
	 
	'This is the resulting Key Handle if successful
	Dim ReshKey As Long
	 
	'This is the variable that recieves the disposition
	Dim Disposition As Long
	
	'Call the Function
	res = RegCreateKeyEx(hKey, SubKey, 0, "", REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, vbNull, ReshKey, Disposition)
	 
	'Test if it was successful
	If res <> ERROR_SUCCESS Then
		MsgBox ("Error")
		MyRegCreateKeyEx = res
	Else
		MsgBox ("Success - hKey:" & Str(ReshKey))
		MyRegCreateKeyEx = ERROR_SUCCESS
	End If
End Function

Public Function MyRegDeleteKey(hKey As Long, SubKey As String) As Long
	'This Function takes in one string - "SubKey" and Deletes it
	'An example of what SubKey should look like:
	'"Sublevel1\Sublevel2" etc.
	'This variable takes the result of the function
	
	Dim res As Long
	
	'Call the Function
	
	res = RegDeleteKey(hKey, SubKey)
	
	'Test if it was successful
	
	If res <> ERROR_SUCCESS Then
		MsgBox ("Error")
		MyRegDeleteKey = res
	Else
		MsgBox ("Success")
		MyRegDeleteKey = ERROR_SUCCESS
	End If
End Function

Public Function MyRegDeleteValue(hKey As Long, ValueName As String) As Long
	'This Function takes in one string - "ValueName" and Deletes the value asscoiated with it
	 'This variable takes the result of the function
	
	Dim Res As Long
	 
	'Call the Function
	Res = RegDeleteValue(hKey, ValueName)
	 
	'Test if it was successful
	
	If Res <> ERROR_SUCCESS Then
		MsgBox ("Error")
		MyRegDeleteValue = Res
	Else
		MsgBox ("Success")
		MyRegDeleteValue = ERROR_SUCCESS
	End If
End Function

Public Function MyRegEnumKeyEx(hKey As Long) As Long
	Dim Res As Long
	Dim Index As Long
	Dim KeyString As String
	Dim KeyStringLen As Long
	Dim ClassString As String
	Dim ClassStringLen As Long
	Dim LastWriteTime As FILETIME
	
	'Initialise the variable
	Index = 0
	
	'Call the Function repeatedly until it returns the NO_MORE_ITEMS error or another error

	Do
		'Initialise Strings each time
		KeyStringLen = MAX_LENGTH
		KeyString = Space(KeyStringLen)
		ClassStringLen = MAX_LENGTH
		ClassString = Space(ClassStringLen)
		Res = RegEnumKeyEx(hKey, Index, KeyString, KeyStringLen, 0, ClassString, ClassStringLen, LastWriteTime)
		
		'Check For Error
		
		If Res <> ERROR_SUCCESS Then
			'There was an Error
			'Make Sure its not the NO_MORE_ITEMS error	
			If Res <> ERROR_NO_MORE_ITEMS Then
				MsgBox ("An error has occured")
				MyRegEnumKeyEx = Res
				Exit Function
			End If
		Else
			'Function Was successful, Tell user the keyname and and class
			'and Go around once again to see if there are any more keys.
			Call MsgBox("KeyName: " & Left(KeyString, KeyStringLen), 0, "Success")
			Call MsgBox("Class: " & Left(ClassString, ClassStringLen), 0, "Success")
		End If
		
		'This increases the index and checks to see if we have run out
		'of items meaning keys and if we have then it will exit the loop
		
		Index = Index + 1
	Loop While Res <> ERROR_NO_MORE_ITEMS
	'Tell User we have no more items(Keys) left to enumerate
	MsgBox ("No More Items")
	MyRegEnumKeyEx = ERROR_SUCCESS
End Function

'This Function is a demonstration of RegEnumValue it is limited to displaying string data
Public Function MyRegEnumValue(hKey As Long) As Long
	Dim res As Long
	 
	Dim Index As Long
	Dim ValueString As String
	Dim ValueStringLen As Long
	 
	Dim DataType As Long
	 
	Dim DataString As String
	Dim DataStringLen As Long
	 
	'Initialise the variable
	Index = 0
	'Call the Function repeatedly until it returns the NO_MORE_ITEMS error or another error
	Do
		'Initialise Strings each time
		ValueStringLen = MAX_LENGTH
		ValueString = Space(ValueStringLen)
		 
		DataType = 0
		 
		DataStringLen = MAX_LENGTH
		DataString = Space(DataStringLen)
		
		'The ByVal Keyword before the DataString is essential
		res = RegEnumValue(hKey, Index, ValueString, ValueStringLen, 0, DataType, ByVal DataString, DataStringLen)
		
		'Check For Error
		If res <> ERROR_SUCCESS Then
			'There was an Error
			'Make Sure its not the NO_MORE_ITEMS error
			If res <> ERROR_NO_MORE_ITEMS Then
				MsgBox ("An error has occured")
				MyRegEnumValue = res
				Exit Function
			End If
		Else
			'Function Was successful, Check What Variable Type it is and if it is a string then
			'display the data aswell The Datatype is checked using a select case statement. This
			'Can also be used for formating the datatype to make a universal function which will
			'Display any data no matter what its type.
			 
			'Display ValueName
			Call MsgBox("ValueName: " & Left(ValueString, ValueStringLen), 0, "Success")
			 
			'Here is the Select Case Statement
			Select Case DataType
				Case REG_BINARY
					'Binary
					MsgBox ("Data Type: Binary")
				Case REG_DWORD
					'DWORD
					MsgBox ("Data Type: DWORD")
				Case REG_NONE
					'None
					MsgBox ("Data Type: None")
				Case REG_LINK
					'Link
					MsgBox ("Data Type: Link")
				Case REG_MULTI_SZ
					'MultiString
					MsgBox ("Data Type: Multi-String")
				Case REG_SZ
					'String data we can display this.
					MsgBox ("Data Type: String")
					MsgBox ("Data: " & Left(DataString, DataStringLen))
				Case Else
					'UnKnown
					MsgBox ("Unknown Data Type")
			End Select
		End If
		 
		'This increases the index and checks to see if we have run out
		'of items meaning Values and if we have then it will exit the loop
		Index = Index + 1
	Loop While res <> ERROR_NO_MORE_ITEMS
	'Tell User we have no more items(Values) left to enumerate
	MsgBox ("No More Items")
	MyRegEnumValue = ERROR_SUCCESS
End Function

Public Function MyRegOpenKeyEx(BaseKey As Long, SubKey As String) As Long
	Dim res As Long 'Result Variable
	Dim hOpenKey As Long 'Key Handle Variable
	 
	'Call The Function
	res = RegOpenKeyEx(BaseKey, SubKey, 0, KEY_ALL_ACCESS, hOpenKey)
	 
	'Check For Error
	If res <> ERROR_SUCCESS Then
		MsgBox ("Error")
		MyRegOpenKeyEx = res
		Exit Function
	Else
		'Function Was Successful
		MsgBox ("Success")
		MsgBox ("Key Handle: " & hOpenKey)
	End If
	MyRegOpenKeyEx = ERROR_SUCCESS
End Function


Public Function MyRegQueryInfoKey(hKey As Long)
	Dim res As Long
	Dim ClassString As String
	Dim ClassStringSize As Long
	 
	Dim NSubKeys As Long
	Dim MaxSubKeySize As Long
	Dim MaxClassSize As Long
	Dim NValues As Long
	Dim MaxValueNameSize As Long
	Dim MaxDataSize As Long
	Dim SecurityDesSize As Long
	 
	Dim LastWriteTime As FILETIME
	 
	'Initialise the String
	ClassStringSize = MAX_LENGTH
	ClassString = Space(ClassStringSize)
	 
	'Call The Function
	res = RegQueryInfoKey(hKey, ClassString, ClassStringSize, vbNull, NSubKeys, MaxSubKeySize, MaxClassSize, NValues, MaxValueNameSize, MaxDataSize, SecurityDesSize, LastWriteTime)
	 
	'Check For Error
	If res <> ERROR_SUCCESS Then
		MsgBox ("Error")
		MyRegQueryInfoKey = res
		Exit Function
	Else
		'Success Tell User the Info
		MsgBox ("Success")
		MsgBox ("Classname: " & Left(ClassString, ClassStringSize))
		MsgBox ("Number Of Subkeys: " & NSubKeys)
		MsgBox ("Max Subkey Size: " & MaxSubKeySize)
		MsgBox ("Max Class Size: " & MaxClassSize)
		MsgBox ("Number Of Values: " & NValues)
		MsgBox ("Max Value Name Size: " & MaxValueNameSize)
		MsgBox ("Max Data Size: " & MaxDataSize)
		MsgBox ("Security Descriptor Size: " & SecurityDesSize)
	End If
	MyRegQueryInfoKey = res
End Function


Public Function MyRegSetValueEx(hKey As Long, ValueName As String, Data As String) As Long
	Dim res As Long
	Dim DataType As Long
	Dim DataSize As Long
	 
	'Initialise
	DataSize = Len(Data)
	DataType = REG_SZ
	 
	'Call Function
	'The ByVal Keyword is very important here
	res = RegSetValueEx(hKey, ValueName, 0, DataType, ByVal Data, DataSize)
	 
	'Check For Error
	If res <> ERROR_SUCCESS Then
		MsgBox ("Error")
		MyRegSetValueEx = res
		Exit Function
	Else
		MsgBox ("Success")
	End If
	 
	MyRegSetValueEx = ERROR_SUCCESS
End Function