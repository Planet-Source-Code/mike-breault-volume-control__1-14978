VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVolume"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const MMSYSERR_NOERROR = 0
Private Const MAXPNAMELEN = 32
Private Const MIXER_LONG_NAME_CHARS = 64
Private Const MIXER_SHORT_NAME_CHARS = 16
Private Const MIXER_GETLINEINFOF_COMPONENTTYPE = &H3&
Private Const MIXER_GETCONTROLDETAILSF_VALUE = &H0&
Private Const MIXER_GETLINECONTROLSF_ONEBYTYPE = &H2&
Private Const MIXER_GETLINECONTROLSF_ALL = &H0&
Private Const MIXER_GETLINECONTROLSF_ONEBYID = &H1&
Private Const MIXER_GETLINECONTROLSF_QUERYMASK = &HF&
Private Const MIXERLINE_COMPONENTTYPE_DST_FIRST = &H0&
Private Const MIXERLINE_COMPONENTTYPE_SRC_FIRST = &H1000&
Private Const MIXER_SETCONTROLDETAILSF_VALUE = &H0&
      
Private Const MIXERCONTROL_CT_CLASS_FADER = &H50000000
Private Const MIXERCONTROL_CT_UNITS_UNSIGNED = &H30000
Private Const MIXERCONTROL_CT_CLASS_LIST = &H70000000
Private Const MIXERCONTROL_CT_SC_LIST_MULTIPLE = &H1000000
Private Const MIXERCONTROL_CT_UNITS_BOOLEAN = &H10000
Private Const MIXERCONTROL_CT_CLASS_SWITCH = &H20000000
Private Const MIXERCONTROL_CT_SC_SWITCH_BOOLEAN = &H0&
Private Const MIXERCONTROL_CT_CLASS_METER = &H10000000
Private Const MIXERCONTROL_CT_SC_METER_POLLED = &H0&
Private Const MIXERCONTROL_CT_UNITS_SIGNED = &H20000
Private Const MIXERCONTROL_CT_CLASS_SLIDER = &H40000000

Private Const MIXERLINE_COMPONENTTYPE_DST_SPEAKERS = _
    (MIXERLINE_COMPONENTTYPE_DST_FIRST + 4)
                     
Private Const MIXERLINE_COMPONENTTYPE_SRC_MICROPHONE = _
    (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 3)
      
Private Const MIXERLINE_COMPONENTTYPE_SRC_LINE = _
    (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 2)
      
Private Const MIXERLINE_COMPONENTTYPE_SRC_COMPACTDISC = _
    (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 5)
    
Private Const MIXERLINE_COMPONENTTYPE_SRC_SYNTHESIZER = _
    (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 4)

Private Const MIXERLINE_COMPONENTTYPE_DST_WAVEIN = _
    (MIXERLINE_COMPONENTTYPE_DST_FIRST + 7)
Private Const MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT = _
    (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 8)

Private Const MIXERCONTROL_CONTROLTYPE_SIGNEDMETER = _
    (MIXERCONTROL_CT_CLASS_METER Or _
    MIXERCONTROL_CT_SC_METER_POLLED Or MIXERCONTROL_CT_UNITS_SIGNED)

Private Const MIXERCONTROL_CONTROLTYPE_BOOLEAN = _
    (MIXERCONTROL_CT_CLASS_SWITCH Or _
    MIXERCONTROL_CT_SC_SWITCH_BOOLEAN Or MIXERCONTROL_CT_UNITS_BOOLEAN)

Private Const MIXERCONTROL_CONTROLTYPE_MULTIPLESELECT = _
    (MIXERCONTROL_CT_CLASS_LIST Or _
    MIXERCONTROL_CT_SC_LIST_MULTIPLE Or MIXERCONTROL_CT_UNITS_BOOLEAN)

Private Const MIXERCONTROL_CONTROLTYPE_FADER = _
    (MIXERCONTROL_CT_CLASS_FADER Or _
    MIXERCONTROL_CT_UNITS_UNSIGNED)
      
Private Const MIXERCONTROL_CONTROLTYPE_VOLUME = _
    (MIXERCONTROL_CONTROLTYPE_FADER + 1)

Private Const MIXERCONTROL_CONTROLTYPE_BASS = _
    (MIXERCONTROL_CONTROLTYPE_FADER + 2)
    
Private Const MIXERCONTROL_CONTROLTYPE_EQUALIZER = _
    (MIXERCONTROL_CONTROLTYPE_FADER + 4)
    
Private Const MIXERCONTROL_CONTROLTYPE_MIXER = _
    (MIXERCONTROL_CONTROLTYPE_MULTIPLESELECT + 1)
    
Private Const MIXERCONTROL_CONTROLTYPE_MONO = _
    (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 3)
    
Private Const MIXERCONTROL_CONTROLTYPE_MUTE = _
    (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 2)
    
Private Const MIXERCONTROL_CONTROLTYPE_ONOFF = _
    (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 1)
    
Private Const MIXERCONTROL_CONTROLTYPE_PEAKMETER = _
    (MIXERCONTROL_CONTROLTYPE_SIGNEDMETER + 1)
    
Private Const MIXERCONTROL_CONTROLTYPE_TREBLE = _
    (MIXERCONTROL_CONTROLTYPE_FADER + 3)

Private Const MIXERCONTROL_CONTROLTYPE_SLIDER = _
    (MIXERCONTROL_CT_CLASS_SLIDER Or MIXERCONTROL_CT_UNITS_SIGNED)

Private Const MIXERCONTROL_CONTROLTYPE_PAN = _
    (MIXERCONTROL_CONTROLTYPE_SLIDER + 1)
    
Private Declare Function mixerClose Lib "winmm.dll" _
    (ByVal hmx As Long) As Long
         
Private Declare Function mixerGetControlDetails Lib "winmm.dll" _
    Alias "mixerGetControlDetailsA" _
    (ByVal hmxobj As Long, _
    pMxcd As MIXERCONTROLDETAILS, _
    ByVal fdwDetails As Long) As Long
         
Private Declare Function mixerGetDevCaps Lib "winmm.dll" _
    Alias "mixerGetDevCapsA" _
    (ByVal uMxId As Long, _
    ByVal pmxcaps As MIXERCAPS, _
    ByVal cbmxcaps As Long) As Long
         
Private Declare Function mixerGetID Lib "winmm.dll" _
    (ByVal hmxobj As Long, _
    pumxID As Long, _
    ByVal fdwId As Long) As Long
                     
Private Declare Function mixerGetLineControls Lib "winmm.dll" _
    Alias "mixerGetLineControlsA" _
    (ByVal hmxobj As Long, _
    pmxlc As MIXERLINECONTROLS, _
    ByVal fdwControls As Long) As Long
                     
Private Declare Function mixerGetLineInfo Lib "winmm.dll" _
    Alias "mixerGetLineInfoA" _
    (ByVal hmxobj As Long, _
    pmxl As MIXERLINE, _
    ByVal fdwInfo As Long) As Long
                     
Private Declare Function mixerGetNumDevs Lib "winmm.dll" () As Long
      
Private Declare Function mixerMessage Lib "winmm.dll" _
    (ByVal hmx As Long, _
    ByVal uMsg As Long, _
    ByVal dwParam1 As Long, _
    ByVal dwParam2 As Long) As Long
                     
Private Declare Function mixerOpen Lib "winmm.dll" _
    (phmx As Long, _
    ByVal uMxId As Long, _
    ByVal dwCallback As Long, _
    ByVal dwInstance As Long, _
    ByVal fdwOpen As Long) As Long
                     
Private Declare Function mixerSetControlDetails Lib "winmm.dll" _
    (ByVal hmxobj As Long, _
    pMxcd As MIXERCONTROLDETAILS, _
    ByVal fdwDetails As Long) As Long
              
Private Declare Sub CopyStructFromPtr Lib "kernel32" _
    Alias "RtlMoveMemory" _
    (struct As Any, _
    ByVal ptr As Long, ByVal cb As Long)
                     
Private Declare Sub CopyPtrFromStruct Lib "kernel32" _
    Alias "RtlMoveMemory" _
    (ByVal ptr As Long, _
    struct As Any, _
    ByVal cb As Long)
                     
Private Declare Function GlobalAlloc Lib "kernel32" _
    (ByVal wFlags As Long, _
    ByVal dwBytes As Long) As Long
                     
Private Declare Function GlobalLock Lib "kernel32" _
    (ByVal hmem As Long) As Long
                     
Private Declare Function GlobalFree Lib "kernel32" _
    (ByVal hmem As Long) As Long
      
Private Type MIXERCAPS
    wMid As Integer
    wPid As Integer
    vDriverVersion As Long
    szPname As String * MAXPNAMELEN
    fdwSupport As Long
    cDestinations As Long
End Type
      
Private Type MIXERCONTROL
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
      
Private Type MIXERCONTROLDETAILS
    cbStruct As Long
    dwControlID As Long
    cChannels As Long
    item As Long
    cbDetails As Long
    paDetails As Long
End Type
      
Private Type MIXERCONTROLDETAILS_UNSIGNED
    dwValue As Long
End Type

Private Type MIXERCONTROLDETAILS_BOOLEAN
    fValue As Long
End Type

Private Type MIXERCONTROLDETAILS_SIGNED
   lValue As Long
End Type

Private Type MIXERCONTROLDETAILS_LISTTEXT
    dwParam1 As Long
    dwParam2 As Long
    szName As String * MIXER_LONG_NAME_CHARS
End Type

Private Type MIXERLINE
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
      
Private Type MIXERLINECONTROLS
    cbStruct As Long
    dwLineID As Long
    dwControl As Long
    cControls As Long
    cbmxctrl As Long
    pamxctrl As Long
End Type

Private hmixer As Long
Private volCtrl As MIXERCONTROL
Private wavCtrl As MIXERCONTROL
Private volMute As MIXERCONTROL
Private wavMute As MIXERCONTROL
Private volFad As MIXERCONTROL
Private wavFad As MIXERCONTROL
Private volBass As MIXERCONTROL
Private wavBass As MIXERCONTROL
Private volTreb As MIXERCONTROL
Private wavTreb As MIXERCONTROL
Private volPan As MIXERCONTROL
Private wavPan As MIXERCONTROL
Private rc As Long
Private ok As Boolean
Private Const CALLBACK_FUNCTION = &H30000
Private Const MM_WIM_DATA = &H3C0
Private Const WHDR_DONE = &H1
Private Const GMEM_FIXED = &H0

Private Type WAVEHDR
   lpData As Long
   dwBufferLength As Long
   dwBytesRecorded As Long
   dwUser As Long
   dwFlags As Long
   dwLoops As Long
   lpNext As Long
   Reserved As Long
End Type

Private Type WAVEINCAPS
   wMid As Integer
   wPid As Integer
   vDriverVersion As Long
   szPname As String * 32
   dwFormats As Long
   wChannels As Integer
End Type

Private Type WAVEFORMAT
   wFormatTag As Integer
   nChannels As Integer
   nSamplesPerSec As Long
   nAvgBytesPerSec As Long
   nBlockAlign As Integer
   wBitsPerSample As Integer
   cbSize As Integer
End Type

Private Declare Function waveInOpen Lib "winmm.dll" (lphWaveIn As Long, _
                                             ByVal uDeviceID As Long, _
                                             lpFormat As WAVEFORMAT, _
                                             ByVal dwCallback As Long, _
                                             ByVal dwInstance As Long, _
                                             ByVal dwFlags As Long) As Long
Private Declare Function waveInPrepareHeader Lib "winmm.dll" (ByVal hWaveIn As Long, _
                                                      lpWaveInHdr As WAVEHDR, _
                                                      ByVal uSize As Long) As Long
Private Declare Function waveInReset Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Private Declare Function waveInStart Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Private Declare Function waveInStop Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Private Declare Function waveInUnprepareHeader Lib "winmm.dll" _
                                          (ByVal hWaveIn As Long, _
                                          lpWaveInHdr As WAVEHDR, _
                                          ByVal uSize As Long) As Long
Private Declare Function waveInClose Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Private Declare Function waveInGetDevCaps Lib "winmm.dll" Alias "waveInGetDevCapsA" _
                  (ByVal uDeviceID As Long, _
                  lpCaps As WAVEINCAPS, _
                  ByVal uSize As Long) As Long
Private Declare Function waveInGetNumDevs Lib "winmm.dll" () As Long
Private Declare Function waveInGetErrorText Lib "winmm.dll" Alias "waveInGetErrorTextA" _
                     (ByVal err As Long, _
                     ByVal lpText As String, _
                     ByVal uSize As Long) As Long
Private Declare Function waveInAddBuffer Lib "winmm.dll" (ByVal hWaveIn As Long, _
                                                   lpWaveInHdr As WAVEHDR, _
                                                   ByVal uSize As Long) As Long

Private i As Integer
Private j As Integer
Private msg As String * 200
Private hWaveIn As Long
Private format As WAVEFORMAT

Private Const NUM_BUFFERS = 2
Private Const BUFFER_SIZE = 8192
Private Const DEVICEID = 0
Private hmem(NUM_BUFFERS) As Long
Private inHdr(NUM_BUFFERS) As WAVEHDR

Private fRecording As Boolean

Private inputVolCtrl As MIXERCONTROL
Private outputVolCtrl As MIXERCONTROL

Private vol As MIXERCONTROLDETAILS_SIGNED
Private volume As Long
Private volHmem As Long
Private mxcd As MIXERCONTROLDETAILS

Private Sub Class_Initialize()
    Debug.Print "Volume Initialize"
    rc = mixerOpen(hmixer, 0, 0, 0, 0)
    ok = GetVolumeControl(hmixer, _
        MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, _
        MIXERCONTROL_CONTROLTYPE_VOLUME, _
        volCtrl)
    ok = GetVolumeControl(hmixer, _
        MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT, _
        MIXERCONTROL_CONTROLTYPE_VOLUME, _
        wavCtrl)
    ok = GetVolumeControl(hmixer, _
        MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, _
        MIXERCONTROL_CONTROLTYPE_MUTE, _
        volMute)
    ok = GetVolumeControl(hmixer, _
        MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT, _
        MIXERCONTROL_CONTROLTYPE_MUTE, _
        wavMute)
    ok = GetVolumeControl(hmixer, _
        MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, _
        MIXERCONTROL_CONTROLTYPE_FADER, _
        volFad)
    ok = GetVolumeControl(hmixer, _
        MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT, _
        MIXERCONTROL_CONTROLTYPE_FADER, _
        wavFad)
    ok = GetVolumeControl(hmixer, _
        MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, _
        MIXERCONTROL_CONTROLTYPE_BASS, _
        volBass)
    ok = GetVolumeControl(hmixer, _
        MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT, _
        MIXERCONTROL_CONTROLTYPE_BASS, _
        wavBass)
    ok = GetVolumeControl(hmixer, _
        MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, _
        MIXERCONTROL_CONTROLTYPE_TREBLE, _
        volTreb)
    ok = GetVolumeControl(hmixer, _
        MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT, _
        MIXERCONTROL_CONTROLTYPE_TREBLE, _
        wavTreb)
    ok = GetVolumeControl(hmixer, _
        MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, _
        MIXERCONTROL_CONTROLTYPE_PAN, _
        volPan)
    ok = GetVolumeControl(hmixer, _
        MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT, _
        MIXERCONTROL_CONTROLTYPE_PAN, _
        wavPan)
   mxcd.cbStruct = Len(mxcd)
   volHmem = GlobalAlloc(&H0, Len(volume))
   mxcd.paDetails = GlobalLock(volHmem)
   mxcd.cbDetails = Len(volume)
   mxcd.cChannels = 1
   ok = GetVolumeControl(hmixer, MIXERLINE_COMPONENTTYPE_DST_WAVEIN, _
   MIXERCONTROL_CONTROLTYPE_PEAKMETER, inputVolCtrl)
   If (ok <> True) Then
       ok = GetVolumeControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_MICROPHONE, MIXERCONTROL_CONTROLTYPE_PEAKMETER, inputVolCtrl)
   End If
   ok = GetVolumeControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT, _
   MIXERCONTROL_CONTROLTYPE_PEAKMETER, outputVolCtrl)
End Sub

Private Sub Class_Terminate()
    Debug.Print "Volume Terminate"
    mixerClose hmixer
    If (fRecording = True) Then
        StopInput
    End If
    GlobalFree volHmem

End Sub


Private Function SetVolume(mxc As MIXERCONTROL, _
                              ByVal volume As Long) As Boolean
    Dim mxcd As MIXERCONTROLDETAILS
    Dim vol As MIXERCONTROLDETAILS_UNSIGNED
    Dim hmem As Long
    mxcd.item = 0
    mxcd.dwControlID = mxc.dwControlID
    mxcd.cbStruct = Len(mxcd)
    mxcd.cbDetails = Len(vol)
    hmem = GlobalAlloc(&H40, Len(vol))
    mxcd.paDetails = GlobalLock(hmem)
    mxcd.cChannels = 1
    vol.dwValue = volume
    CopyPtrFromStruct mxcd.paDetails, vol, Len(vol)
    rc = mixerSetControlDetails(hmixer, _
        mxcd, _
        MIXER_SETCONTROLDETAILSF_VALUE)
    GlobalFree (hmem)
    If (MMSYSERR_NOERROR = rc) Then
        SetVolume = True
    Else
        SetVolume = False
    End If
End Function

Private Function GetVolume(mxc As MIXERCONTROL) As Long
    Dim mxcd As MIXERCONTROLDETAILS
    Dim vol As MIXERCONTROLDETAILS_UNSIGNED
    Dim hmem As Long
    mxcd.item = 0
    mxcd.dwControlID = mxc.dwControlID
    mxcd.cbStruct = Len(mxcd)
    mxcd.cbDetails = Len(vol)
    hmem = GlobalAlloc(&H40, Len(vol))
    mxcd.paDetails = GlobalLock(hmem)
    mxcd.cChannels = 1
    rc = mixerGetControlDetails(hmixer, _
        mxcd, _
        MIXER_GETCONTROLDETAILSF_VALUE)
    CopyStructFromPtr vol, mxcd.paDetails, Len(vol)
    GetVolume = vol.dwValue
    GlobalFree (hmem)
End Function

Private Function GetTreble(mxc As MIXERCONTROL) As Long
    Dim mxcd As MIXERCONTROLDETAILS
    Dim Treb As MIXERCONTROLDETAILS_UNSIGNED
    Dim hmem As Long
    mxcd.item = 0
    mxcd.dwControlID = mxc.dwControlID
    mxcd.cbStruct = Len(mxcd)
    mxcd.cbDetails = Len(Treb)
    hmem = GlobalAlloc(&H40, Len(Treb))
    mxcd.paDetails = GlobalLock(hmem)
    mxcd.cChannels = 1
    rc = mixerGetControlDetails(hmixer, _
        mxcd, _
        MIXER_GETCONTROLDETAILSF_VALUE)
    CopyStructFromPtr Treb, mxcd.paDetails, Len(Treb)
    GetTreble = Treb.dwValue
    GlobalFree (hmem)
End Function


Private Function SetTreble(mxc As MIXERCONTROL, _
                              ByVal Treble As Long)
    Dim mxcd As MIXERCONTROLDETAILS
    Dim Treb As MIXERCONTROLDETAILS_UNSIGNED
    Dim hmem As Long
    mxcd.item = 0
    mxcd.dwControlID = mxc.dwControlID
    mxcd.cbStruct = Len(mxcd)
    mxcd.cbDetails = Len(Treb)
    hmem = GlobalAlloc(&H40, Len(Treb))
    mxcd.paDetails = GlobalLock(hmem)
    mxcd.cChannels = 1
    Treb.dwValue = Treble
    CopyPtrFromStruct mxcd.paDetails, Treb, Len(Treb)
    rc = mixerSetControlDetails(hmixer, _
        mxcd, _
        MIXER_SETCONTROLDETAILSF_VALUE)
    GlobalFree (hmem)
End Function

Private Function GetBass(mxc As MIXERCONTROL) As Long
    Dim mxcd As MIXERCONTROLDETAILS
    Dim Bas As MIXERCONTROLDETAILS_UNSIGNED
    Dim hmem As Long
    mxcd.item = 0
    mxcd.dwControlID = mxc.dwControlID
    mxcd.cbStruct = Len(mxcd)
    mxcd.cbDetails = Len(Bas)
    hmem = GlobalAlloc(&H40, Len(Bas))
    mxcd.paDetails = GlobalLock(hmem)
    mxcd.cChannels = 1
    rc = mixerGetControlDetails(hmixer, _
        mxcd, _
        MIXER_GETCONTROLDETAILSF_VALUE)
    CopyStructFromPtr Bas, mxcd.paDetails, Len(Bas)
    GetBass = Bas.dwValue
    GlobalFree (hmem)
End Function

Private Function SetBass(mxc As MIXERCONTROL, _
                              ByVal Bass As Long)
    Dim mxcd As MIXERCONTROLDETAILS
    Dim Bas As MIXERCONTROLDETAILS_UNSIGNED
    Dim hmem As Long
    mxcd.item = 0
    mxcd.dwControlID = mxc.dwControlID
    mxcd.cbStruct = Len(mxcd)
    mxcd.cbDetails = Len(Bas)
    hmem = GlobalAlloc(&H40, Len(Bas))
    mxcd.paDetails = GlobalLock(hmem)
    mxcd.cChannels = 1
    Bas.dwValue = Bass
    CopyPtrFromStruct mxcd.paDetails, Bas, Len(Bas)
    rc = mixerSetControlDetails(hmixer, _
        mxcd, _
        MIXER_SETCONTROLDETAILSF_VALUE)
    GlobalFree (hmem)
End Function


Private Function GetPan(mxc As MIXERCONTROL) As Boolean
    Dim mxcd As MIXERCONTROLDETAILS
    Dim Pan As MIXERCONTROLDETAILS_UNSIGNED
    Dim hmem As Long
    mxcd.item = 0
    mxcd.dwControlID = mxc.dwControlID
    mxcd.cbStruct = Len(mxcd)
    mxcd.cbDetails = Len(Pan)
    hmem = GlobalAlloc(&H40, Len(Pan))
    mxcd.paDetails = GlobalLock(hmem)
    mxcd.cChannels = 1
    rc = mixerGetControlDetails(hmixer, _
        mxcd, _
        MIXER_GETCONTROLDETAILSF_VALUE)
    CopyStructFromPtr Pan, mxcd.paDetails, Len(Pan)
    GetPan = Pan.dwValue
    GlobalFree (hmem)
End Function


Private Function SetPan(mxc As MIXERCONTROL, _
                              ByVal Pano As Long)
    Dim mxcd As MIXERCONTROLDETAILS
    Dim Pan As MIXERCONTROLDETAILS_UNSIGNED
    Dim hmem As Long
    mxcd.item = 0
    mxcd.dwControlID = mxc.dwControlID
    mxcd.cbStruct = Len(mxcd)
    mxcd.cbDetails = Len(Pan)
    hmem = GlobalAlloc(&H40, Len(Pan))
    mxcd.paDetails = GlobalLock(hmem)
    mxcd.cChannels = 1
    Pan.dwValue = Pano
    CopyPtrFromStruct mxcd.paDetails, Pan, Len(Pan)
    rc = mixerSetControlDetails(hmixer, _
        mxcd, _
        MIXER_SETCONTROLDETAILSF_VALUE)
    GlobalFree (hmem)
End Function

Private Function GetMute(mxc As MIXERCONTROL) As Boolean
    Dim mxcd As MIXERCONTROLDETAILS
    Dim Mut As MIXERCONTROLDETAILS_BOOLEAN
    Dim hmem As Long
    mxcd.item = 0
    mxcd.dwControlID = mxc.dwControlID
    mxcd.cbStruct = Len(mxcd)
    mxcd.cbDetails = Len(Mut)
    hmem = GlobalAlloc(&H40, Len(Mut))
    mxcd.paDetails = GlobalLock(hmem)
    mxcd.cChannels = 1
    rc = mixerGetControlDetails(hmixer, _
        mxcd, _
        MIXER_GETCONTROLDETAILSF_VALUE)
    CopyStructFromPtr Mut, mxcd.paDetails, Len(Mut)
    GetMute = IIf((Mut.fValue = 1), True, False)
    GlobalFree (hmem)
End Function

Private Function SetMute(mxc As MIXERCONTROL, _
                              ByVal Mute As Boolean)
    Dim mxcd As MIXERCONTROLDETAILS
    Dim Mut As MIXERCONTROLDETAILS_BOOLEAN
    Dim hmem As Long
    mxcd.item = 0
    mxcd.dwControlID = mxc.dwControlID
    mxcd.cbStruct = Len(mxcd)
    mxcd.cbDetails = Len(Mut)
    hmem = GlobalAlloc(&H40, Len(Mut))
    mxcd.paDetails = GlobalLock(hmem)
    mxcd.cChannels = 1
    Mut.fValue = IIf(Mute, 1, 0)
    CopyPtrFromStruct mxcd.paDetails, Mut, Len(Mut)
    rc = mixerSetControlDetails(hmixer, _
        mxcd, _
        MIXER_SETCONTROLDETAILSF_VALUE)
    GlobalFree (hmem)
End Function


Private Function GetFader(mxc As MIXERCONTROL) As Long
    Dim mxcd As MIXERCONTROLDETAILS
    Dim Fad As MIXERCONTROLDETAILS_UNSIGNED
    Dim hmem As Long
    mxcd.item = 0
    mxcd.dwControlID = mxc.dwControlID
    mxcd.cbStruct = Len(mxcd)
    mxcd.cbDetails = Len(Fad)
    hmem = GlobalAlloc(&H40, Len(Fad))
    mxcd.paDetails = GlobalLock(hmem)
    mxcd.cChannels = 1
    rc = mixerGetControlDetails(hmixer, _
        mxcd, _
        MIXER_GETCONTROLDETAILSF_VALUE)
    CopyStructFromPtr Fad, mxcd.paDetails, Len(Fad)
    GetFader = Fad.dwValue
    GlobalFree (hmem)
End Function

Private Function SetFader(mxc As MIXERCONTROL, _
                              ByVal Fader As Long)
    Dim mxcd As MIXERCONTROLDETAILS
    Dim Fad As MIXERCONTROLDETAILS_UNSIGNED
    Dim hmem As Long
    mxcd.item = 0
    mxcd.dwControlID = mxc.dwControlID
    mxcd.cbStruct = Len(mxcd)
    mxcd.cbDetails = Len(Fad)
    hmem = GlobalAlloc(&H40, Len(Fad))
    mxcd.paDetails = GlobalLock(hmem)
    mxcd.cChannels = 1
    Fad.dwValue = Fader
    CopyPtrFromStruct mxcd.paDetails, Fad, Len(Fad)
    rc = mixerSetControlDetails(hmixer, _
        mxcd, _
        MIXER_SETCONTROLDETAILSF_VALUE)
    GlobalFree (hmem)
End Function

Public Property Get VolumeMin() As Long
    VolumeMin = volCtrl.lMinimum
End Property

Public Property Get WaveMin() As Long
    WaveMin = wavCtrl.lMinimum
End Property

Public Property Get VolTrebleMin() As Long
    VolTrebleMin = volTreb.lMinimum
End Property

Public Property Get WaveTrebleMin() As Long
    WaveTrebleMin = wavTreb.lMinimum
End Property

Public Property Get VolBassMin() As Long
    VolBassMin = volBass.lMinimum
End Property

Public Property Get WaveBassMin() As Long
    WaveBassMin = wavBass.lMinimum
End Property

Public Property Get VolumeMax() As Long
    VolumeMax = volCtrl.lMaximum
End Property

Public Property Get WaveMax() As Long
    WaveMax = wavCtrl.lMaximum
End Property

Public Property Get VolTrebleMax() As Long
    VolTrebleMax = volTreb.lMaximum
End Property

Public Property Get VolBassMax() As Long
    VolBassMax = volBass.lMaximum
End Property

Public Property Get VolumeLevel() As Long
    VolumeLevel = GetVolume(volCtrl)
End Property

Public Property Get WaveLevel() As Long
    WaveLevel = GetVolume(wavCtrl)
End Property

Public Property Get VolumeLevelBass() As Long
    VolumeLevelBass = GetBass(volBass)
End Property

Public Property Get VolumeLevelTreble() As Long
    VolumeLevelTreble = GetTreble(volTreb)
End Property

Public Property Let WaveLevel(ByVal New_Volume As Long)
    SetVolume wavCtrl, New_Volume
End Property

Public Property Let VolumeLevel(ByVal New_Volume As Long)
    SetVolume volCtrl, New_Volume
End Property

Public Property Let WaveLevelBass(ByVal New_Bass As Long)
    SetBass wavBass, New_Bass
End Property

Public Property Let VolumeLevelBass(ByVal New_Bass As Long)
    SetBass volBass, New_Bass
End Property

Public Property Let WaveLevelTreble(ByVal New_Treble As Long)
    SetTreble wavTreb, New_Treble
End Property

Public Property Let VolumeLevelTreble(ByVal New_Treble As Long)
    SetTreble volTreb, New_Treble
End Property

Public Property Get VolumeMute() As Boolean
    VolumeMute = GetMute(volMute)
End Property

Public Property Get WaveMute() As Boolean
    WaveMute = GetMute(wavMute)
End Property

Public Property Let VolumeMute(New_Mute As Boolean)
    SetMute volMute, New_Mute
End Property

Public Property Let WaveMute(New_Mute As Boolean)
    SetMute wavMute, New_Mute
End Property

Private Function StartInput() As Boolean
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
        MsgBox msg, vbCritical + vbOKOnly, _
        App.Title & " - Mixercontrol Fehler"
        StartInput = False
        Exit Function
    End If
    For i = 0 To NUM_BUFFERS - 1
        rc = waveInPrepareHeader(hWaveIn, inHdr(i), Len(inHdr(i)))
        If (rc <> 0) Then
            waveInGetErrorText rc, msg, Len(msg)
            MsgBox msg, vbCritical + vbOKOnly, _
            App.Title & " - Mixercontrol Fehler"
        End If
    Next
    For i = 0 To NUM_BUFFERS - 1
        rc = waveInAddBuffer(hWaveIn, inHdr(i), Len(inHdr(i)))
        If (rc <> 0) Then
            waveInGetErrorText rc, msg, Len(msg)
            MsgBox msg, vbCritical + vbOKOnly, _
            App.Title & " - Mixercontrol Fehler"
        End If
    Next
    fRecording = True
    rc = waveInStart(hWaveIn)
    StartInput = True
End Function

Private Sub StopInput()
    fRecording = False
    waveInReset hWaveIn
    waveInStop hWaveIn
    For i = 0 To NUM_BUFFERS - 1
        waveInUnprepareHeader hWaveIn, inHdr(i), Len(inHdr(i))
        GlobalFree hmem(i)
    Next
    waveInClose hWaveIn
End Sub

Private Sub waveInProc(ByVal hwi As Long, ByVal uMsg As Long, ByVal dwInstance As Long, ByRef hdr As WAVEHDR, ByVal dwParam2 As Long)
   If (uMsg = MM_WIM_DATA) Then
      If fRecording Then
         rc = waveInAddBuffer(hwi, hdr, Len(hdr))
      End If
   End If
End Sub

Private Function GetVolumeControl(ByVal hmixer As Long, _
                              ByVal componentType As Long, _
                              ByVal ctrlType As Long, _
                              ByRef mxc As MIXERCONTROL) As Boolean
    Dim mxlc As MIXERLINECONTROLS
    Dim mxl As MIXERLINE
    Dim hmem As Long
    Dim rc As Long
    mxl.cbStruct = Len(mxl)
    mxl.dwComponentType = componentType
    rc = mixerGetLineInfo(hmixer, mxl, _
        MIXER_GETLINEINFOF_COMPONENTTYPE)
    If (MMSYSERR_NOERROR = rc) Then
        mxlc.cbStruct = Len(mxlc)
        mxlc.dwLineID = mxl.dwLineID
        mxlc.dwControl = ctrlType
        mxlc.cControls = 1
        mxlc.cbmxctrl = Len(mxc)
        hmem = GlobalAlloc(&H40, Len(mxc))
        mxlc.pamxctrl = GlobalLock(hmem)
        mxc.cbStruct = Len(mxc)
        rc = mixerGetLineControls(hmixer, _
            mxlc, _
            MIXER_GETLINECONTROLSF_ONEBYTYPE)
        If (MMSYSERR_NOERROR = rc) Then
            GetVolumeControl = True
            CopyStructFromPtr mxc, mxlc.pamxctrl, Len(mxc)
        Else
            GetVolumeControl = False
        End If
        GlobalFree (hmem)
        Exit Function
    End If
    GetVolumeControl = False
End Function


Public Property Get MaxVolumeMeterInput() As Long
    MaxVolumeMeterInput = inputVolCtrl.lMaximum
End Property

Public Property Get MinVolumeMeterInput() As Long
    MinVolumeMeterInput = inputVolCtrl.lMinimum
End Property

Public Property Get MaxVolumeMeterOutput() As Long
    MaxVolumeMeterOutput = outputVolCtrl.lMaximum
End Property

Public Property Get MinVolumeMeterOutput() As Long
    MinVolumeMeterOutput = outputVolCtrl.lMinimum
End Property

Public Property Get CurrentVolumeMeterOutput() As Long
    If (fRecording) Then
       For i = 0 To (NUM_BUFFERS - 1)
          If inHdr(i).dwFlags And WHDR_DONE Then
             rc = waveInAddBuffer(hWaveIn, inHdr(i), Len(inHdr(i)))
          End If
       Next
    End If
    mxcd.dwControlID = outputVolCtrl.dwControlID
    mxcd.item = outputVolCtrl.cMultipleItems
    rc = mixerGetControlDetails(hmixer, mxcd, _
    MIXER_GETCONTROLDETAILSF_VALUE)
    CopyStructFromPtr volume, mxcd.paDetails, Len(volume)
    If (volume < 0) Then volume = -volume
    CurrentVolumeMeterOutput = volume
End Property

Public Property Get CurrentVolumeMeterInput() As Long
    If Not StartInput Then StartInput
    If (fRecording) Then
        For i = 0 To (NUM_BUFFERS - 1)
            If inHdr(i).dwFlags And WHDR_DONE Then
                rc = waveInAddBuffer(hWaveIn, inHdr(i), Len(inHdr(i)))
            End If
        Next
    End If
    mxcd.dwControlID = inputVolCtrl.dwControlID
    mxcd.item = inputVolCtrl.cMultipleItems
    rc = mixerGetControlDetails(hmixer, mxcd, _
    MIXER_GETCONTROLDETAILSF_VALUE)
    CopyStructFromPtr volume, mxcd.paDetails, Len(volume)
    If (volume < 0) Then
        volume = -volume
    End If
    CurrentVolumeMeterInput = volume
End Property
