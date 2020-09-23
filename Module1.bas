Attribute VB_Name = "Module1"
Option Explicit

Public Const CALLBACK_FUNCTION = &H30000
Public Const MM_WIM_DATA = &H3C0
Public Const WHDR_DONE = &H1
Public Const GMEM_FIXED = &H0

Type WAVEHDR
   lpData As Long
   dwBufferLength As Long
   dwBytesRecorded As Long
   dwUser As Long
   dwFlags As Long
   dwLoops As Long
   lpNext As Long
   Reserved As Long
End Type
Type WAVEINCAPS
   wMid As Integer
   wPid As Integer
   vDriverVersion As Long
   szPname As String * 32
   dwFormats As Long
   wChannels As Integer
End Type
Type WAVEFORMAT
   wFormatTag As Integer
   nChannels As Integer
   nSamplesPerSec As Long
   nAvgBytesPerSec As Long
   nBlockAlign As Integer
   wBitsPerSample As Integer
   cbSize As Integer
End Type

Declare Function waveInOpen Lib "winmm.dll" (lphWaveIn As Long, ByVal uDeviceID As Long, lpFormat As WAVEFORMAT, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Declare Function waveInPrepareHeader Lib "winmm.dll" (ByVal hWaveIn As Long, lpWaveInHdr As WAVEHDR, ByVal uSize As Long) As Long
Declare Function waveInReset Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Declare Function waveInStart Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Declare Function waveInStop Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Declare Function waveInUnprepareHeader Lib "winmm.dll" (ByVal hWaveIn As Long, lpWaveInHdr As WAVEHDR, ByVal uSize As Long) As Long
Declare Function waveInClose Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Declare Function waveInGetDevCaps Lib "winmm.dll" Alias "waveInGetDevCapsA" (ByVal uDeviceID As Long, lpCaps As WAVEINCAPS, ByVal uSize As Long) As Long
Declare Function waveInGetNumDevs Lib "winmm.dll" () As Long
Declare Function waveInGetErrorText Lib "winmm.dll" Alias "waveInGetErrorTextA" (ByVal err As Long, ByVal lpText As String, ByVal uSize As Long) As Long
Declare Function waveInAddBuffer Lib "winmm.dll" (ByVal hWaveIn As Long, lpWaveInHdr As WAVEHDR, ByVal uSize As Long) As Long

Public Const MMSYSERR_NOERROR = 0
Public Const MAXPNAMELEN = 32

Public Const MIXER_LONG_NAME_CHARS = 64
Public Const MIXER_SHORT_NAME_CHARS = 16
Public Const MIXER_GETLINEINFOF_COMPONENTTYPE = &H3&
Public Const MIXER_GETCONTROLDETAILSF_VALUE = &H0&
Public Const MIXER_GETLINECONTROLSF_ONEBYTYPE = &H2&

Public Const MIXERLINE_COMPONENTTYPE_DST_FIRST = &H0&
Public Const MIXERLINE_COMPONENTTYPE_SRC_FIRST = &H1000&
Public Const MIXERLINE_COMPONENTTYPE_DST_SPEAKERS = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 4)
Public Const MIXERLINE_COMPONENTTYPE_SRC_MICROPHONE = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 3)
Public Const MIXERLINE_COMPONENTTYPE_SRC_LINE = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 2)

Public Const MIXERCONTROL_CT_CLASS_FADER = &H50000000
Public Const MIXERCONTROL_CT_UNITS_UNSIGNED = &H30000
Public Const MIXERCONTROL_CT_UNITS_SIGNED = &H20000
Public Const MIXERCONTROL_CT_CLASS_METER = &H10000000
Public Const MIXERCONTROL_CT_SC_METER_POLLED = &H0&
Public Const MIXERCONTROL_CONTROLTYPE_FADER = (MIXERCONTROL_CT_CLASS_FADER Or MIXERCONTROL_CT_UNITS_UNSIGNED)
Public Const MIXERCONTROL_CONTROLTYPE_VOLUME = (MIXERCONTROL_CONTROLTYPE_FADER + 1)
Public Const MIXERLINE_COMPONENTTYPE_DST_WAVEIN = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 7)
Public Const MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 8)
Public Const MIXERCONTROL_CONTROLTYPE_SIGNEDMETER = (MIXERCONTROL_CT_CLASS_METER Or MIXERCONTROL_CT_SC_METER_POLLED Or MIXERCONTROL_CT_UNITS_SIGNED)
Public Const MIXERCONTROL_CONTROLTYPE_PEAKMETER = (MIXERCONTROL_CONTROLTYPE_SIGNEDMETER + 1)

Declare Function mixerClose Lib "winmm.dll" (ByVal hmx As Long) As Long
Declare Function mixerGetControlDetails Lib "winmm.dll" Alias "mixerGetControlDetailsA" (ByVal hmxobj As Long, pmxcd As MIXERCONTROLDETAILS, ByVal fdwDetails As Long) As Long
Declare Function mixerGetDevCaps Lib "winmm.dll" Alias "mixerGetDevCapsA" (ByVal uMxId As Long, ByVal pmxcaps As MIXERCAPS, ByVal cbmxcaps As Long) As Long
Declare Function mixerGetID Lib "winmm.dll" (ByVal hmxobj As Long, pumxID As Long, ByVal fdwId As Long) As Long
Declare Function mixerGetLineInfo Lib "winmm.dll" Alias "mixerGetLineInfoA" (ByVal hmxobj As Long, pmxl As MIXERLINE, ByVal fdwInfo As Long) As Long
Declare Function mixerGetLineControls Lib "winmm.dll" Alias "mixerGetLineControlsA" (ByVal hmxobj As Long, pmxlc As MIXERLINECONTROLS, ByVal fdwControls As Long) As Long
Declare Function mixerGetNumDevs Lib "winmm.dll" () As Long
Declare Function mixerMessage Lib "winmm.dll" (ByVal hmx As Long, ByVal uMsg As Long, ByVal dwParam1 As Long, ByVal dwParam2 As Long) As Long
Declare Function mixerOpen Lib "winmm.dll" (phmx As Long, ByVal uMxId As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal fdwOpen As Long) As Long
Declare Function mixerSetControlDetails Lib "winmm.dll" (ByVal hmxobj As Long, pmxcd As MIXERCONTROLDETAILS, ByVal fdwDetails As Long) As Long

Declare Sub CopyStructFromPtr Lib "kernel32" Alias "RtlMoveMemory" (struct As Any, ByVal ptr As Long, ByVal cb As Long)
Declare Sub CopyPtrFromStruct Lib "kernel32" Alias "RtlMoveMemory" (ByVal ptr As Long, struct As Any, ByVal cb As Long)
Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Declare Function GlobalLock Lib "kernel32" (ByVal hmem As Long) As Long
Declare Function GlobalFree Lib "kernel32" (ByVal hmem As Long) As Long

Type MIXERCAPS
   wMid As Integer
   wPid As Integer
   vDriverVersion As Long
   szPname As String * MAXPNAMELEN
   fdwSupport As Long
   cDestinations As Long
End Type
Type MIXERCONTROL
   cbStruct As Long
   dwControlID As Long
   dwControlType As Long
   fdwControl As Long
   cMultipleItems As Long
   szShortName As String * MIXER_SHORT_NAME_CHARS
   szName As String * MIXER_LONG_NAME_CHARS
   lMinimum As Long
   lMaximum As Long
   Reserved(10) As Long
End Type
Type MIXERCONTROLDETAILS
   cbStruct As Long
   dwControlID As Long
   cChannels As Long
   item As Long
   cbDetails As Long
   paDetails As Long
End Type
Type MIXERCONTROLDETAILS_SIGNED
   lValue As Long
End Type
Type MIXERLINE
   cbStruct As Long
   dwDestination As Long
   dwSource As Long
   dwLineID As Long
   fdwLine As Long
   dwUser As Long
   dwComponentType As Long
   cChannels As Long
   cConnections As Long
   cControls As Long
   szShortName As String * MIXER_SHORT_NAME_CHARS
   szName As String * MIXER_LONG_NAME_CHARS
   dwType As Long
   dwDeviceID As Long
   wMid  As Integer
   wPid As Integer
   vDriverVersion As Long
   szPname As String * MAXPNAMELEN
End Type
Type MIXERLINECONTROLS
   cbStruct As Long
   dwLineID As Long
   dwControl As Long
   cControls As Long
   cbmxctrl As Long
   pamxctrl As Long
End Type

Public i As Integer, j As Integer, rc As Long, msg As String * 200, hWaveIn As Long
Public Const NUM_BUFFERS = 2
Public format As WAVEFORMAT, hmem(NUM_BUFFERS) As Long, inHdr(NUM_BUFFERS) As WAVEHDR
Public Const BUFFER_SIZE = 8192
Public Const DEVICEID = 0
Public fRecording As Boolean
Function GetControl(ByVal hmixer As Long, ByVal componentType As Long, ByVal ctrlType As Long, ByRef mxc As MIXERCONTROL) As Boolean
' This function attempts to obtain a mixer control. Returns True if successful.

   Dim mxlc As MIXERLINECONTROLS
   Dim mxl As MIXERLINE
   Dim hmem As Long
   Dim rc As Long
    
   mxl.cbStruct = Len(mxl)
   mxl.dwComponentType = componentType
   
   rc = mixerGetLineInfo(hmixer, mxl, MIXER_GETLINEINFOF_COMPONENTTYPE)
   
   If (MMSYSERR_NOERROR = rc) Then
      mxlc.cbStruct = Len(mxlc)
      mxlc.dwLineID = mxl.dwLineID
      mxlc.dwControl = ctrlType
      mxlc.cControls = 1
      mxlc.cbmxctrl = Len(mxc)
      hmem = GlobalAlloc(GMEM_FIXED, Len(mxc))
      mxlc.pamxctrl = GlobalLock(hmem)
      mxc.cbStruct = Len(mxc)
      rc = mixerGetLineControls(hmixer, mxlc, MIXER_GETLINECONTROLSF_ONEBYTYPE)
      If (MMSYSERR_NOERROR = rc) Then
         GetControl = True
         CopyStructFromPtr mxc, mxlc.pamxctrl, Len(mxc)
      Else
         GetControl = False
      End If
      GlobalFree (hmem)
      Exit Function
   End If
   
   GetControl = False
End Function

Sub waveInProc(ByVal hwi As Long, ByVal uMsg As Long, ByVal dwInstance As Long, ByRef hdr As WAVEHDR, ByVal dwParam2 As Long)
   If (uMsg = MM_WIM_DATA) Then
      If fRecording Then
         rc = waveInAddBuffer(hwi, hdr, Len(hdr))
      End If
   End If
End Sub

Function StartInput() As Boolean

    If fRecording Then
        StartInput = True
        Exit Function
    End If
    
    format.wFormatTag = 1
    format.nChannels = 1
    format.wBitsPerSample = 8
    format.nSamplesPerSec = 8000
    format.nBlockAlign = format.nChannels * format.wBitsPerSample / 8
    format.nAvgBytesPerSec = format.nSamplesPerSec * format.nBlockAlign
    format.cbSize = 0
    
    For i = 0 To NUM_BUFFERS - 1
        hmem(i) = GlobalAlloc(&H40, BUFFER_SIZE)
        inHdr(i).lpData = GlobalLock(hmem(i))
        inHdr(i).dwBufferLength = BUFFER_SIZE
        inHdr(i).dwFlags = 0
        inHdr(i).dwLoops = 0
    Next

    rc = waveInOpen(hWaveIn, DEVICEID, format, 0, 0, 0)
    If rc <> 0 Then
        waveInGetErrorText rc, msg, Len(msg)
        MsgBox msg
        StartInput = False
        Exit Function
    End If

    For i = 0 To NUM_BUFFERS - 1
        rc = waveInPrepareHeader(hWaveIn, inHdr(i), Len(inHdr(i)))
        If (rc <> 0) Then
            waveInGetErrorText rc, msg, Len(msg)
            MsgBox msg
        End If
    Next

    For i = 0 To NUM_BUFFERS - 1
        rc = waveInAddBuffer(hWaveIn, inHdr(i), Len(inHdr(i)))
        If (rc <> 0) Then
            waveInGetErrorText rc, msg, Len(msg)
            MsgBox msg
        End If
    Next

    fRecording = True
    rc = waveInStart(hWaveIn)
    StartInput = True
End Function

Sub StopInput()
    fRecording = False
    waveInReset hWaveIn
    waveInStop hWaveIn
    For i = 0 To NUM_BUFFERS - 1
        waveInUnprepareHeader hWaveIn, inHdr(i), Len(inHdr(i))
        GlobalFree hmem(i)
    Next
    waveInClose hWaveIn
End Sub

