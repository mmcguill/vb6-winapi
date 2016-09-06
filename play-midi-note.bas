Option Explicit

Private Declare Function midiOutClose Lib "winmm.dll" (ByVal hMidiOut As Long) As Long
Private Declare Function midiOutShortMsg Lib "winmm.dll" (ByVal hMidiOut As Long, ByVal dwMsg As Long) As Long
Private Declare Function midiOutOpen Lib "winmm.dll" (lphMidiOut As Long, ByVal uDeviceID As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Private Declare Function midiOutReset Lib "winmm.dll" (ByVal hMidiOut As Long) As Long

Dim hMidiOut As Long

Private Sub Command1_Click()
	Call SendMidiOut(144, 55, 100)
End Sub

Private Sub Form_Load()
	Dim X As Long
	Dim hMidiOut As Long
	X = midiOutOpen(hMidiOut, -1&, 0&, 0&, 0&)
	If X <> 0 Then End
End Sub

Private Sub Form_Unload(Cancel As Integer)
	Dim X As Long
	X = midiOutClose(hMidiOut)
	X = midiOutReset(hMidiOut)
End Sub

Private Sub SendMidiOut(MidiEventOut As Long, MidiNoteOut As Long, MidiVelOut As Long)
	Dim X As Long
	Dim lowint As Long
	Dim highint As Long
	Dim Velout As Long
	Dim MidiMessage As Long
	lowint = (MidiNoteOut * 256) + MidiEventOut
	Velout = MidiVelOut * 256
	highint = Velout * 256
	MidiMessage = lowint + highint
	X = midiOutShortMsg(hMidiOut, MidiMessage)
End Sub
