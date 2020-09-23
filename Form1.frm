VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H8000000B&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Volume Control"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   4425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Tag             =   "Created By: Mike Breault"
   Begin VB.CheckBox chkVolMute 
      Caption         =   "Mute"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1905
      TabIndex        =   2
      Top             =   2040
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Left            =   1920
      Top             =   480
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Align           =   1  'Align Top
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4425
      _ExtentX        =   7805
      _ExtentY        =   450
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin VB.PictureBox Pic1 
      DrawWidth       =   2
      FillStyle       =   0  'Solid
      ForeColor       =   &H8000000E&
      Height          =   2055
      Left            =   0
      ScaleHeight     =   1995
      ScaleWidth      =   4275
      TabIndex        =   1
      Top             =   2520
      Width           =   4335
   End
   Begin ComctlLib.Slider sldMainTreb 
      Height          =   1815
      Left            =   465
      TabIndex        =   3
      Top             =   600
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   3201
      _Version        =   327682
      BorderStyle     =   1
      Orientation     =   1
      TickStyle       =   1
   End
   Begin ComctlLib.Slider sldMainBas 
      Height          =   1815
      Left            =   3345
      TabIndex        =   4
      Top             =   600
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   3201
      _Version        =   327682
      BorderStyle     =   1
      Orientation     =   1
      TickStyle       =   1
   End
   Begin ComctlLib.Slider sldMain 
      Height          =   615
      Left            =   1305
      TabIndex        =   5
      Top             =   1320
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1085
      _Version        =   327682
      BorderStyle     =   1
      TickStyle       =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Main Volume"
      Height          =   255
      Left            =   1785
      TabIndex        =   8
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Treble"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   5
      Left            =   465
      TabIndex        =   7
      Top             =   360
      Width           =   450
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bass"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   5
      Left            =   3345
      TabIndex        =   6
      Top             =   360
      Width           =   345
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim hmixer As Long
Dim inputVolCtrl As MIXERCONTROL
Dim outputVolCtrl As MIXERCONTROL
Dim rc As Long
Dim ok As Boolean

Dim mxcd As MIXERCONTROLDETAILS
Dim vol2 As MIXERCONTROLDETAILS_SIGNED
Dim volume As Long
Dim volHmem As Long
Dim PicX As Integer
Dim PicY As Integer
Dim NewX As Integer
Dim NewY As Integer
Dim OldX As Integer
Dim OldY As Integer
Dim Started As Boolean

Private vol As New clsVolume

Private Sub chkVolMute_Click()
     vol.VolumeMute = IIf((chkVolMute.Value = 1), True, False)
End Sub

Private Sub Form_Load()
    Me.ScaleMode = vbTwips
    Timer1.Interval = 50
    Timer1.Enabled = True
    rc = mixerOpen(hmixer, DEVICEID, 0, 0, 0)
    If ((MMSYSERR_NOERROR <> rc)) Then
        MsgBox "Couldn't open the mixer."
        Exit Sub
    End If
    ok = GetControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT, MIXERCONTROL_CONTROLTYPE_PEAKMETER, outputVolCtrl)
    If (ok = True) Then
       ProgressBar1.Min = 0
       ProgressBar1.Max = outputVolCtrl.lMaximum
    Else
       ProgressBar1.Enabled = False
       MsgBox "Couldn't get waveout meter"
    End If
    mxcd.cbStruct = Len(mxcd)
    volHmem = GlobalAlloc(&H0, Len(volume))
    mxcd.paDetails = GlobalLock(volHmem)
    mxcd.cbDetails = Len(volume)
    mxcd.cChannels = 1
    PicX = 0
    PicY = 1950
    NewX = 0
    NewY = 1950
    Started = False
    Set vol = New clsVolume
    With sldMain
        .Min = vol.VolumeMin
        .Max = vol.VolumeMax
        .TickFrequency = (.Max - .Min) \ 10
        .LargeChange = .TickFrequency
    End With
    With sldMainTreb
        If vol.VolTrebleMax = 0 Then
            .Visible = False
            Label2(5).Visible = False
        Else
            .Min = vol.VolTrebleMin
            .Max = vol.VolTrebleMax
            .TickFrequency = (.Max - .Min) \ 10
            .LargeChange = .TickFrequency
        End If
    End With
    With sldMainBas
        If vol.VolBassMax = 0 Then
            .Visible = False
            Label3(5).Visible = False
        Else
            .Min = vol.VolBassMin
            .Max = vol.VolBassMax
            .TickFrequency = (.Max - .Min) \ 10
            .LargeChange = .TickFrequency
        End If
    End With
End Sub

Private Sub Form_Paint()
    sldMain.Value = sldMain.Max - vol.VolumeLevel
    chkVolMute.Value = IIf(vol.VolumeMute, 1, 0)
    If sldMainTreb.Visible Then sldMainTreb.Value = vol.VolumeLevelTreble
    If sldMainBas.Visible Then sldMainBas.Value = vol.VolumeLevelBass
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set vol = Nothing
    GlobalFree volHmem
End Sub

Private Sub sldMain_Scroll()
    vol.VolumeLevel = sldMain.Max - sldMain.Value
End Sub

Private Sub sldMainBas_Scroll()
    vol.VolumeLevelBass = sldMainBas.Value
End Sub

Private Sub sldMainTreb_Scroll()
    vol.VolumeLevelTreble = sldMainTreb.Value
End Sub

Private Sub Timer1_Timer()
   On Error Resume Next
   If (ProgressBar1.Enabled = True) Then
      mxcd.dwControlID = outputVolCtrl.dwControlID
      mxcd.item = outputVolCtrl.cMultipleItems
      rc = mixerGetControlDetails(hmixer, mxcd, MIXER_GETCONTROLDETAILSF_VALUE)
      CopyStructFromPtr volume, mxcd.paDetails, Len(volume)
      If (volume < 0) Then volume = -volume
      ProgressBar1.Value = volume
      Dim Val As Long
      Val = Int(ProgressBar1.Value * (100 / ProgressBar1.Max))
      PicY = 1950 - (1950 * (Val / 100))
      If (PicX + 40) < 3730 Then
      PicX = PicX + 40
      Else
      Pic1.Cls
      PicX = 0
      Started = False
      End If
      Pic1.Circle (PicX, PicY), 1, RGB(0, 0, 0)
      OldX = NewX
      OldY = NewY
      NewX = PicX
      NewY = PicY
      If Started = True Then
      Pic1.Line (OldX, OldY)-(NewX, NewY), RGB(0, 0, 0)
      Else
      Started = True
      End If
   End If
End Sub
