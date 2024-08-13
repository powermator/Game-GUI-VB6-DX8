Attribute VB_Name = "DX7Sound"
'****************************************************************************
'* DirectX 7 Direct Sound Module for Visual Basic                           *
'****************************************************************************
'* By, Jonathan Enzinna                                                     *
'*     enzinna@hotpop.com                                                   *
'* May 14, 2000                                                             *
'****************************************************************************
'* Copyright (c) 2000: Lazer Light Ltd.                                     *
'*                     http://lazer-light.0pi.com/                          *
'****************************************************************************
'* If you have any questions, comments, or need help concerning this code,  *
'* then please e-mail me at enzinna@hotpop.com. If you have a new way to do *
'* the following functions (that should be shorter or more generic) then    *
'* feel free to e-mail them to me and get your name next to mine also. For  *
'* the newest version of this code, visit the Lazer Light Ltd. web page, or *
'* go to www.zarr.net/vb/                                                   *
'****************************************************************************


' To Initialize the buffers, use the following strings in the form:
'  cSoundBuffers=10 'Select the number of buffers you wish to use
'  SetupDX7Sound Me              ' Assign DX7 Direct Sound to the current form
'  CreateBuffers cSoundBuffers, "Default.wav"

Global cSoundBuffers As Integer
Option Explicit
Private m_dx As New DirectX7
Private m_dxs As DirectSound
Type dxBuffers
  isLoaded As Boolean
  Buffer As DirectSoundBuffer
End Type
Private SB() As dxBuffers

Public Sub CreateBuffers(AmountOfBuffer As Integer, DefaultFile As String)
  ReDim SB(AmountOfBuffer)
  For AmountOfBuffer = 0 To AmountOfBuffer
    DX7LoadSound AmountOfBuffer, DefaultFile
    VolumeLevel AmountOfBuffer, 50
  Next AmountOfBuffer
End Sub
Public Sub SetupDX7Sound(CurrentForm As Form)
  Set m_dxs = m_dx.DirectSoundCreate("")
  If Err.Number <> 0 Then
    MsgBox "Unable to start DirectSound. Check to see that your sound card is properly installed"
    End
  End If
  m_dxs.SetCooperativeLevel CurrentForm.hWnd, DSSCL_PRIORITY
End Sub
Public Sub DX7LoadSound(Buffer As Integer, sfile As String)
  Dim Filename As String
  Dim bufferDesc As DSBUFFERDESC
  Dim waveFormat As WAVEFORMATEX
  bufferDesc.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN _
  Or DSBCAPS_CTRLVOLUME Or DSBCAPS_STATIC
  waveFormat.nFormatTag = WAVE_FORMAT_PCM ' The sounds must be in ADPCM header format
  ' PLEASE NOTE: The following three lines MUST be set to accomidate your sounds
  ' or else DX7 will not run the sound properly, not run at all, or crash the user's
  ' computer (and we don't want that at all).
  waveFormat.nChannels = 2    '2 channels (Mono/Stereo)
  waveFormat.lSamplesPerSec = 22050 ' 22050 khz Sample Rate
  waveFormat.nBitsPerSample = 16  '16 bit
  waveFormat.nBlockAlign = waveFormat.nBitsPerSample / 8 * waveFormat.nChannels
  waveFormat.lAvgBytesPerSec = waveFormat.lSamplesPerSec * waveFormat.nBlockAlign
  Filename = sfile
  On Error GoTo Continue
  Set SB(Buffer).Buffer = m_dxs.CreateSoundBufferFromFile(Filename, bufferDesc, waveFormat)
  SB(Buffer).isLoaded = True
  Exit Sub
Continue:
  MsgBox "Error can't find file: " & Filename
End Sub

Public Sub PlaySoundWithPan(Buffer As Integer, Filename As String, Optional Volume As Byte, Optional PanValue As Byte)
  DX7LoadSound Buffer, Filename
  If PanValue <> 50 And PanValue < 100 Then PanSound Buffer, PanValue
  If Volume < 100 Then VolumeLevel Buffer, Volume
  If SB(Buffer).isLoaded Then SB(Buffer).Buffer.Play 0
End Sub
Public Sub PanSound(Buffer As Integer, PanValue As Byte)
' PLEASE NOTE:
' 0 is full left and 100 is full right, where 50 is center. Use this gradually in
' a while loop or on a timer to smoothly pan a sound from left to right or vice
' versa
  Select Case PanValue
    Case 0
      SB(Buffer).Buffer.SetPan -10000
    Case 100
      SB(Buffer).Buffer.SetPan 10000
    Case Else
      SB(Buffer).Buffer.SetPan (100 * PanValue) - 5000
  End Select
End Sub
Public Sub VolumeLevel(Buffer As Integer, Volume As Byte)
' PLEASE NOTE:
' 0 is silent and 100 is full strength, BE CAREFUL WHEN SETTING VALUE
  If Volume > 0 Then
    SB(Buffer).Buffer.SetVolume (60 * Volume) - 6000
  Else
    SB(Buffer).Buffer.SetVolume -6000
  End If
End Sub

Public Function IsPlaying(Buffer As Integer) As Long
  IsPlaying = SB(Buffer).Buffer.GetStatus
End Function
