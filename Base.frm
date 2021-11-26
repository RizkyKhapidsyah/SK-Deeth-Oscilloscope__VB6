VERSION 5.00
Begin VB.Form Base 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stereo DeethScope"
   ClientHeight    =   1665
   ClientLeft      =   3735
   ClientTop       =   1530
   ClientWidth     =   3600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   111
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   240
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Scope 
      BackColor       =   &H80000009&
      ForeColor       =   &H80000002&
      Height          =   768
      Index           =   1
      Left            =   1656
      ScaleHeight     =   256
      ScaleMode       =   0  'User
      ScaleWidth      =   123
      TabIndex        =   6
      Top             =   468
      Width           =   1524
   End
   Begin VB.Frame Stuff 
      BorderStyle     =   0  'None
      Height          =   336
      Left            =   72
      TabIndex        =   2
      Top             =   1296
      Width           =   3360
      Begin VB.CommandButton StartButton 
         Caption         =   "&Start"
         Height          =   336
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   804
      End
      Begin VB.CommandButton StopButton 
         Caption         =   "S&top"
         Enabled         =   0   'False
         Height          =   336
         Left            =   864
         TabIndex        =   4
         Top             =   0
         Width           =   804
      End
      Begin VB.CheckBox Flicker 
         Caption         =   "Flickerless"
         Height          =   300
         Left            =   1800
         TabIndex        =   3
         Top             =   36
         Width           =   1632
      End
   End
   Begin VB.PictureBox Scope 
      BackColor       =   &H80000009&
      ForeColor       =   &H80000002&
      Height          =   768
      Index           =   0
      Left            =   72
      ScaleHeight     =   256
      ScaleMode       =   0  'User
      ScaleWidth      =   123
      TabIndex        =   1
      Top             =   468
      Width           =   1524
   End
   Begin VB.ComboBox DevicesBox 
      Height          =   288
      Left            =   72
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   72
      Width           =   3108
   End
   Begin VB.Shape Shape 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1188
      Left            =   0
      Top             =   0
      Width           =   1812
   End
End
Attribute VB_Name = "Base"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private DevHandle As Long
Private InData(0 To 511) As Byte
Private Inited As Boolean
Public MinHeight As Long, MinWidth As Long

Private Type WaveFormatEx
    FormatTag As Integer
    Channels As Integer
    SamplesPerSec As Long
    AvgBytesPerSec As Long
    BlockAlign As Integer
    BitsPerSample As Integer
    ExtraDataSize As Integer
End Type

Private Type WaveHdr
    lpData As Long
    dwBufferLength As Long
    dwBytesRecorded As Long
    dwUser As Long
    dwFlags As Long
    dwLoops As Long
    lpNext As Long 'wavehdr_tag
    Reserved As Long
End Type

Private Type WaveInCaps
    ManufacturerID As Integer      'wMid
    ProductID As Integer       'wPid
    DriverVersion As Long       'MMVERSIONS vDriverVersion
    ProductName(1 To 32) As Byte 'szPname[MAXPNAMELEN]
    Formats As Long
    Channels As Integer
    Reserved As Integer
End Type

Private Const WAVE_INVALIDFORMAT = &H0&                 '/* invalid format */
Private Const WAVE_FORMAT_1M08 = &H1&                   '/* 11.025 kHz, Mono,   8-bit
Private Const WAVE_FORMAT_1S08 = &H2&                   '/* 11.025 kHz, Stereo, 8-bit
Private Const WAVE_FORMAT_1M16 = &H4&                   '/* 11.025 kHz, Mono,   16-bit
Private Const WAVE_FORMAT_1S16 = &H8&                   '/* 11.025 kHz, Stereo, 16-bit
Private Const WAVE_FORMAT_2M08 = &H10&                  '/* 22.05  kHz, Mono,   8-bit
Private Const WAVE_FORMAT_2S08 = &H20&                  '/* 22.05  kHz, Stereo, 8-bit
Private Const WAVE_FORMAT_2M16 = &H40&                  '/* 22.05  kHz, Mono,   16-bit
Private Const WAVE_FORMAT_2S16 = &H80&                  '/* 22.05  kHz, Stereo, 16-bit
Private Const WAVE_FORMAT_4M08 = &H100&                 '/* 44.1   kHz, Mono,   8-bit
Private Const WAVE_FORMAT_4S08 = &H200&                 '/* 44.1   kHz, Stereo, 8-bit
Private Const WAVE_FORMAT_4M16 = &H400&                 '/* 44.1   kHz, Mono,   16-bit
Private Const WAVE_FORMAT_4S16 = &H800&                 '/* 44.1   kHz, Stereo, 16-bit

Private Const WAVE_FORMAT_PCM = 1

Private Const WHDR_DONE = &H1&              '/* done bit */
Private Const WHDR_PREPARED = &H2&          '/* set if this header has been prepared */
Private Const WHDR_BEGINLOOP = &H4&         '/* loop start block */
Private Const WHDR_ENDLOOP = &H8&           '/* loop end block */
Private Const WHDR_INQUEUE = &H10&          '/* reserved for driver */

Private Const WIM_OPEN = &H3BE
Private Const WIM_CLOSE = &H3BF
Private Const WIM_DATA = &H3C0

Private Declare Function waveInAddBuffer Lib "winmm" (ByVal InputDeviceHandle As Long, ByVal WaveHdrPointer As Long, ByVal WaveHdrStructSize As Long) As Long
Private Declare Function waveInPrepareHeader Lib "winmm" (ByVal InputDeviceHandle As Long, ByVal WaveHdrPointer As Long, ByVal WaveHdrStructSize As Long) As Long
Private Declare Function waveInUnprepareHeader Lib "winmm" (ByVal InputDeviceHandle As Long, ByVal WaveHdrPointer As Long, ByVal WaveHdrStructSize As Long) As Long

Private Declare Function waveInGetNumDevs Lib "winmm" () As Long
Private Declare Function waveInGetDevCaps Lib "winmm" Alias "waveInGetDevCapsA" (ByVal uDeviceID As Long, ByVal WaveInCapsPointer As Long, ByVal WaveInCapsStructSize As Long) As Long

Private Declare Function waveInOpen Lib "winmm" (WaveDeviceInputHandle As Long, ByVal WhichDevice As Long, ByVal WaveFormatExPointer As Long, ByVal CallBack As Long, ByVal CallBackInstance As Long, ByVal Flags As Long) As Long
Private Declare Function waveInClose Lib "winmm" (ByVal WaveDeviceInputHandle As Long) As Long

Private Declare Function waveInStart Lib "winmm" (ByVal WaveDeviceInputHandle As Long) As Long
Private Declare Function waveInReset Lib "winmm" (ByVal WaveDeviceInputHandle As Long) As Long
Private Declare Function waveInStop Lib "winmm" (ByVal WaveDeviceInputHandle As Long) As Long


Sub InitDevices()
    Dim Caps As WaveInCaps, Which As Long
    DevicesBox.Clear
    For Which = 0 To waveInGetNumDevs - 1
        Call waveInGetDevCaps(Which, VarPtr(Caps), Len(Caps))
        'If Caps.Formats And WAVE_FORMAT_1M08 Then
        If Caps.Formats And WAVE_FORMAT_1S08 Then 'Now is 1S08 -- Check for devices that can do stereo 8-bit 11kHz
            Call DevicesBox.AddItem(StrConv(Caps.ProductName, vbUnicode), Which)
        End If
    Next
    If DevicesBox.ListCount = 0 Then
        MsgBox "You have no audio input devices!", vbCritical, "Ack!"
        End
    End If
    DevicesBox.ListIndex = 0
End Sub


Private Sub Flicker_Click()
    Scope(0).Cls
    Scope(1).Cls
    If Flicker.Value = vbChecked Then
        Scope(0).AutoRedraw = True
        Scope(1).AutoRedraw = True
    Else
        Scope(0).AutoRedraw = False
        Scope(1).AutoRedraw = False
    End If
End Sub


Private Sub Form_Load()
    Call InitDevices
    
    'Set MinWidth and MinHeight based on Shape...
    Dim XAdjust As Long, YAdjust As Long
    XAdjust = Me.Width \ Screen.TwipsPerPixelX - Me.ScaleWidth
    YAdjust = Me.Height \ Screen.TwipsPerPixelY - Me.ScaleHeight
    
    MinWidth = Shape.Width + XAdjust
    MinHeight = Shape.Height + YAdjust
    
    Shape.BackStyle = vbTransparent
    
    
    'Set the window proceedure to my own (which restricts the
    'minimum size of the form...
    'Comment out the SetWindowLong line if you're working with it
    'in the development environment since it'll hang in stop mode.
    MinMaxProc.Proc = GetWindowLong(Me.HWnd, GWL_WNDPROC)
    SetWindowLong Me.HWnd, GWL_WNDPROC, AddressOf WindowProc

End Sub


Private Sub Form_Resize()
    Scope(0).Cls
    Scope(1).Cls
    
    Stuff.Top = Me.ScaleHeight - Stuff.Height - 3
    Scope(0).Height = Me.ScaleHeight - 75
    Scope(1).Height = Scope(0).Height
    Scope(0).Width = (Me.ScaleWidth - 13) \ 2
    Scope(1).Width = Scope(0).Width
    Scope(1).Left = Scope(0).Left + Scope(0).Width + 1
    
    DevicesBox.Width = Me.ScaleWidth - 13
    
    Scope(0).ScaleHeight = 256
    Scope(0).ScaleWidth = 255
    Scope(1).ScaleHeight = 256
    Scope(1).ScaleWidth = 255
   
    'Make the window resize now so that it doesn't interfere with redrawing the data
    DoEvents
    
    'Redraw the data at the new size
    If Inited = True Then
        Call DrawData
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If DevHandle <> 0 Then
        Call DoStop
    End If
End Sub


Private Sub StartButton_Click()
    Static WaveFormat As WaveFormatEx
    With WaveFormat
        .FormatTag = WAVE_FORMAT_PCM
        .Channels = 2 'Two channels -- left and right
        .SamplesPerSec = 11025 '11khz
        .BitsPerSample = 8
        .BlockAlign = (.Channels * .BitsPerSample) \ 8
        .AvgBytesPerSec = .BlockAlign * .SamplesPerSec
        .ExtraDataSize = 0
    End With
    
    Debug.Print "waveInOpen:"; waveInOpen(DevHandle, DevicesBox.ListIndex, VarPtr(WaveFormat), 0, 0, 0)
    
    If DevHandle = 0 Then
        Call MsgBox("Wave input device didn't open!", vbExclamation, "Ack!")
        Exit Sub
    End If
    Debug.Print " "; DevHandle
    Call waveInStart(DevHandle)
    
    Inited = True
       
    StopButton.Enabled = True
    StartButton.Enabled = False
    
    Call Visualize
End Sub


Private Sub StopButton_Click()
    Call DoStop
End Sub


Private Sub DoStop()
    Call waveInReset(DevHandle)
    Call waveInClose(DevHandle)
    DevHandle = 0
    StopButton.Enabled = False
    StartButton.Enabled = True
End Sub


Private Sub Visualize()
    Static Wave As WaveHdr
    
    Wave.lpData = VarPtr(InData(0))
    Wave.dwBufferLength = 512 'This is now 512 so there's still 256 samples per channel
    Wave.dwFlags = 0
    
    Do
    
        Call waveInPrepareHeader(DevHandle, VarPtr(Wave), Len(Wave))
        Call waveInAddBuffer(DevHandle, VarPtr(Wave), Len(Wave))
    
        Do
            'Nothing -- we're waiting for the audio driver to mark
            'this wave chunk as done.
        Loop Until ((Wave.dwFlags And WHDR_DONE) = WHDR_DONE) Or DevHandle = 0
        
        Call waveInUnprepareHeader(DevHandle, VarPtr(Wave), Len(Wave))
        
        If DevHandle = 0 Then
            'The device has closed...
            Exit Do
        End If
        
        Scope(0).Cls
        Scope(1).Cls
        
        Call DrawData
        
        DoEvents
    Loop While DevHandle <> 0 'While the audio device is open

End Sub


Private Sub DrawData()
    Static X As Long
    
    Scope(0).CurrentX = -1
    Scope(0).CurrentY = Scope(0).ScaleHeight \ 2
    Scope(1).CurrentX = -1
    Scope(1).CurrentY = Scope(0).ScaleHeight \ 2
    
    'Plot the data...
    For X = 0 To 255
        Scope(0).Line Step(0, 0)-(X, InData(X * 2))
        Scope(1).Line Step(0, 0)-(X, InData(X * 2 + 1)) 'For a good soundcard...
        
        'Use these to plot dots instead of lines...
        'Scope(0).PSet (X, InData(X * 2))
        'Scope(1).PSet (X, InData(X * 2 + 1)) 'For a good soundcard...

        'My soundcard is pretty cheap... the right is
        'noticably less loud than the left... so I add five to it.
        'Scope(1).Line Step(0, 0)-(X, InData(X * 2 + 1) + 5)
    Next
            
    Scope(0).CurrentY = Scope(0).Width
    Scope(1).CurrentY = Scope(0).Width
End Sub
