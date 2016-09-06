Option Explicit
Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Public Const FILE_FILE_COMPRESSION = &H10
Public Const FILE_PERSISTENT_ACLS = &H8
Public Const FILE_UNICODE_ON_DISK = &H4
Public Const FILE_CASE_SENSITIVE_SEARCH = &H1
Public Const FILE_CASE_PRESERVED_NAMES = &H2
Public Const FILE_VOLUME_IS_COMPRESSED = &H8000
'This Program Call the Function as soon as it loads up and Gives you the Information in a series of MessageBoxes
 
Public Sub Main()
	Dim Res As Long
	Dim RootPathName As String
	Dim VolumeNameBuffer As String
	Dim VolumeNameSize As Long
	Dim VolumeSerialNumber As Long
	Dim MaximumComponentLength As Long
	Dim FileSystemFlags As Long
	Dim FileSystemNameBuffer As String
	Dim FileSystemNameSize As Long
	 
	'This is used Later
	Dim FlagString As String
	 
	'This Takes the Rootpath of the volume you want info on
	'C:\
	'D:\
	'A:\ (etc...)
	RootPathName = "C:\"
	 
	'Initialise VolumeName string
	VolumeNameSize = 255
	VolumeNameBuffer = Space(VolumeNameSize)
	 
	'Initialise Filesystem string
	FileSystemNameSize = 255
	FileSystemNameBuffer = Space(FileSystemNameSize)
	 
	'Call the Function
	Res = GetVolumeInformation(RootPathName, VolumeNameBuffer, VolumeNameSize, VolumeSerialNumber, MaximumComponentLength, FileSystemFlags, FileSystemNameBuffer, FileSystemNameSize)
	 
	'Test For Error
	If Res <> 0 Then
		'This is the path that you supplied
		MsgBox ("Root Path: " & RootPathName)
		 
		'This is the name of the volume
		MsgBox ("Volume Name: " & Left(VolumeNameBuffer, VolumeNameSize))
		 
		'Serial Number
		MsgBox ("Volume Serial Number: " & VolumeSerialNumber)
		 
		'Filename Length - Maximum
		MsgBox ("Maximum Component Length: " & MaximumComponentLength)
		 
		'Flag Setup
		If (FileSystemFlags And FILE_FILE_COMPRESSION) Then FlagString = "File Compression Capability"
		If (FileSystemFlags And FILE_PERSISTENT_ACLS) Then FlagString = FlagString & Chr(13) & Chr(10) & "Persisant ACLS"
		If (FileSystemFlags And FILE_UNICODE_ON_DISK) Then FlagString = FlagString & Chr(13) & Chr(10) & "Unicode On Disk"
		If (FileSystemFlags And FILE_CASE_SENSITIVE_SEARCH) Then FlagString = FlagString & Chr(13) & Chr(10) & "Case-Sensitive"
		If (FileSystemFlags And FILE_CASE_PRESERVED_NAMES) Then FlagString = FlagString & Chr(13) & Chr(10) & "Case is Preserved"
		If (FileSystemFlags And FILE_VOLUME_IS_COMPRESSED) Then FlagString = FlagString & Chr(13) & Chr(10) & "Volume is Compressed"
		 
		'Test if any flags were set then display messagebox if they were
		If FlagString <> "" Then MsgBox (FlagString)
		 
		'File System Type
		MsgBox ("File System: " & Left(FileSystemNameBuffer, FileSystemNameSize))
	Else
		'Error
		MsgBox ("Error")
	End If
End Sub

'This Programs shows the basic implementation of getdiskfreespace
Public Sub Main()
	Dim Res As Long
	 
	'address of root path
	Dim RootPathName As String
	 
	'sectors per cluster
	Dim SectorsPerCluster As Long
	 
	'bytes per sector
	Dim BytesPerSector As Long
	 
	'number of free clusters
	Dim NumberOfFreeClusters As Long
	 
	'total number of clusters
	Dim TotalNumberOfClusters As Long
	 
	'String variable not used in call but for holding result
	Dim FormatStr As String
	'Initialise
	RootPathName = "C:\"
	 
	'Call Her
	Res = GetDiskFreeSpace(RootPathName, SectorsPerCluster, BytesPerSector, NumberOfFreeClusters, TotalNumberOfClusters)
	 
	If Res <> 0 Then
		FormatStr = "Success" & Chr(13) & Chr(10) & "Root Path: " & RootPathName & Chr(13) & Chr(10)
		FormatStr = FormatStr & "Sectors Per Cluster: " & SectorsPerCluster & Chr(13) & Chr(10)
		FormatStr = FormatStr & "Bytes Per Sector: " & BytesPerSector & Chr(13) & Chr(10)
		FormatStr = FormatStr & "Free Clusters: " & NumberOfFreeClusters & Chr(13) & Chr(10)
		FormatStr = FormatStr & "Total Number Of Clusters: " & TotalNumberOfClusters & Chr(13) & Chr(10) & Chr(13) & Chr(10)
		FormatStr = FormatStr & "Free Space: " & NumberOfFreeClusters * SectorsPerCluster * BytesPerSector
		MsgBox (FormatStr)
	Else
		MsgBox ("Error")
	End If
End Sub