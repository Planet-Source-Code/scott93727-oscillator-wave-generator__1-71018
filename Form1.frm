VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Audio Waveform Generator"
   ClientHeight    =   4605
   ClientLeft      =   60
   ClientTop       =   675
   ClientWidth     =   4605
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4605
   ScaleWidth      =   4605
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      Caption         =   "Frequency"
      Height          =   1695
      Left            =   120
      TabIndex        =   15
      Top             =   2760
      Width           =   4335
      Begin VB.CommandButton Command20 
         Caption         =   "10KHZ"
         Height          =   255
         Left            =   3480
         TabIndex        =   35
         Top             =   1320
         Width           =   735
      End
      Begin VB.CommandButton Command19 
         Caption         =   "9000HZ"
         Height          =   255
         Left            =   2640
         TabIndex        =   34
         Top             =   1320
         Width           =   735
      End
      Begin VB.CommandButton Command18 
         Caption         =   "8000HZ"
         Height          =   255
         Left            =   1800
         TabIndex        =   33
         Top             =   1320
         Width           =   735
      End
      Begin VB.CommandButton Command17 
         Caption         =   "7000HZ"
         Height          =   255
         Left            =   960
         TabIndex        =   32
         Top             =   1320
         Width           =   735
      End
      Begin VB.CommandButton Command16 
         Caption         =   "6000HZ"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   1320
         Width           =   735
      End
      Begin VB.CommandButton Command15 
         Caption         =   "0500HZ"
         Height          =   255
         Left            =   3480
         TabIndex        =   30
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton Command14 
         Caption         =   "0400HZ"
         Height          =   255
         Left            =   2640
         TabIndex        =   29
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton Command13 
         Caption         =   "0300HZ"
         Height          =   255
         Left            =   1800
         TabIndex        =   28
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton Command12 
         Caption         =   "0200HZ"
         Height          =   255
         Left            =   960
         TabIndex        =   27
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton Command11 
         Caption         =   "0100HZ"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton Command10 
         Caption         =   "5000HZ"
         Height          =   255
         Left            =   3480
         TabIndex        =   25
         Top             =   960
         Width           =   735
      End
      Begin VB.CommandButton Command9 
         Caption         =   "4500HZ"
         Height          =   255
         Left            =   2640
         TabIndex        =   24
         Top             =   960
         Width           =   735
      End
      Begin VB.CommandButton Command8 
         Caption         =   "4000HZ"
         Height          =   255
         Left            =   1800
         TabIndex        =   23
         Top             =   960
         Width           =   735
      End
      Begin VB.CommandButton Command7 
         Caption         =   "3500HZ"
         Height          =   255
         Left            =   960
         TabIndex        =   22
         Top             =   960
         Width           =   735
      End
      Begin VB.CommandButton Command6 
         Caption         =   "3000HZ"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   960
         Width           =   735
      End
      Begin VB.CommandButton Command5 
         Caption         =   "2500HZ"
         Height          =   255
         Left            =   3480
         TabIndex        =   20
         Top             =   600
         Width           =   735
      End
      Begin VB.CommandButton Command4 
         Caption         =   "2000HZ"
         Height          =   255
         Left            =   2640
         TabIndex        =   19
         Top             =   600
         Width           =   735
      End
      Begin VB.CommandButton Command3 
         Caption         =   "1500HZ"
         Height          =   255
         Left            =   1800
         TabIndex        =   18
         Top             =   600
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "1000HZ"
         Height          =   255
         Left            =   960
         TabIndex        =   17
         Top             =   600
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "0750HZ"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frequency HZ"
      Height          =   735
      Left            =   120
      TabIndex        =   13
      Top             =   840
      Width           =   1815
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   360
         TabIndex        =   14
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Frequency"
      Height          =   975
      Left            =   120
      TabIndex        =   10
      Top             =   1680
      Width           =   4335
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   240
         Max             =   20000
         Min             =   20
         TabIndex        =   11
         Top             =   480
         Value           =   20
         Width           =   3855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Frequency"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   3855
      End
   End
   Begin VB.VScrollBar VScroll4 
      Height          =   30
      Left            =   600
      TabIndex        =   9
      Top             =   1200
      Width           =   255
   End
   Begin VB.Frame Frame3 
      Caption         =   "Function"
      Height          =   735
      Left            =   2160
      TabIndex        =   6
      Top             =   840
      Width           =   2295
      Begin VB.CommandButton cmdStop 
         Caption         =   "Stop"
         Height          =   375
         Left            =   1200
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdGenerate 
         Caption         =   "Generate"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Waveform"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin VB.OptionButton optNoise 
         Caption         =   "Noise"
         Height          =   255
         Left            =   3480
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton optSaw 
         Caption         =   "Saw"
         Height          =   255
         Left            =   2760
         TabIndex        =   4
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton optTriangle 
         Caption         =   "Triangle"
         Height          =   255
         Left            =   1800
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optSquare 
         Caption         =   "Square"
         Height          =   255
         Left            =   840
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optSine 
         Caption         =   "Sine"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Menu m0 
      Caption         =   "Mixer"
      Begin VB.Menu m1 
         Caption         =   "Mixer"
      End
   End
   Begin VB.Menu h0 
      Caption         =   "Help"
      Begin VB.Menu h1 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'A simple audio waveform generator using DirectX8
'Text boxes and vertical scrollers are used to
'create an array of up/down buttons for frequency
'selection.

'Lower frequency limit is set to 20 Hz.

'Upper frequency limit is 9999.9 Hz which is the
'practical limit when using a 44,100 Hz samplerate.

Option Explicit

'Create the DirectSound8 Object

Dim dx As New DirectX8
Dim ds As DirectSound8
Dim dsBuffer As DirectSoundSecondaryBuffer8

'Declare the variables that will be used
'Dim amplitude
Dim frequency As Double
Dim increment As Double
Dim fileName As String
Dim fileSize As Double
Dim sample As Long
Dim period As Double
Dim state As Integer
Dim bufferptr As Long
Dim inputValue As Double
Dim sw
Const Pi = 3.141592654
Const sampleRate = 44100
Const amplitude = 127
'frequency = 300
'HScroll1.Value = frequency
Private Sub Command1_Click()
Text1.Text = 750
frequency = 750
HScroll1.Value = 750
If sw = 1 Then cmdGenerate_Click
End Sub
Private Sub Command10_Click()
Text1.Text = 5000
frequency = 5000
HScroll1.Value = 5000
If sw = 1 Then cmdGenerate_Click
End Sub
Private Sub Command11_Click()
Text1.Text = 100
frequency = 100
HScroll1.Value = 100
If sw = 1 Then cmdGenerate_Click
End Sub
Private Sub Command12_Click()
Text1.Text = 200
frequency = 200
HScroll1.Value = 200
If sw = 1 Then cmdGenerate_Click
End Sub
Private Sub Command13_Click()
Text1.Text = 300
frequency = 300
HScroll1.Value = 300
If sw = 1 Then cmdGenerate_Click
End Sub
Private Sub Command14_Click()
Text1.Text = 400
frequency = 400
HScroll1.Value = 400
If sw = 1 Then cmdGenerate_Click
End Sub
Private Sub Command15_Click()
Text1.Text = 500
frequency = 500
HScroll1.Value = 500
If sw = 1 Then cmdGenerate_Click
End Sub

Private Sub Command16_Click()
Text1.Text = 6000
frequency = 6000
HScroll1.Value = 6000
If sw = 1 Then cmdGenerate_Click
End Sub

Private Sub Command17_Click()
Text1.Text = 7000
frequency = 7000
HScroll1.Value = 7000
If sw = 1 Then cmdGenerate_Click
End Sub
Private Sub Command18_Click()
Text1.Text = 8000
frequency = 8000
HScroll1.Value = 8000
If sw = 1 Then cmdGenerate_Click
End Sub
Private Sub Command19_Click()
Text1.Text = 9000
frequency = 9000
HScroll1.Value = 9000
If sw = 1 Then cmdGenerate_Click
End Sub
Private Sub Command2_Click()
Text1.Text = 1000
frequency = 1000
HScroll1.Value = 1000
If sw = 1 Then cmdGenerate_Click
End Sub
Private Sub Command20_Click()
Text1.Text = 10000
frequency = 10000
HScroll1.Value = 10000
If sw = 1 Then cmdGenerate_Click
End Sub
Private Sub Command3_Click()
Text1.Text = 1500
frequency = 1500
HScroll1.Value = 1500
If sw = 1 Then cmdGenerate_Click
End Sub
Private Sub Command4_Click()
Text1.Text = 2000
frequency = 2000
HScroll1.Value = 2000
If sw = 1 Then cmdGenerate_Click
End Sub
Private Sub Command5_Click()
Text1.Text = 2500
frequency = 2500
HScroll1.Value = 2500
If sw = 1 Then cmdGenerate_Click
End Sub
Private Sub Command6_Click()
Text1.Text = 3000
frequency = 3000
HScroll1.Value = 3000
If sw = 1 Then cmdGenerate_Click
End Sub
Private Sub Command7_Click()
Text1.Text = 3500
frequency = 3500
HScroll1.Value = 3500
If sw = 1 Then cmdGenerate_Click
End Sub
Private Sub Command8_Click()
Text1.Text = 4000
frequency = 4000
HScroll1.Value = 4000
If sw = 1 Then cmdGenerate_Click
End Sub
Private Sub Command9_Click()
Text1.Text = 4500
frequency = 4500
HScroll1.Value = 4500
If sw = 1 Then cmdGenerate_Click
End Sub

'Initialize Waveform and Frequency Selectors - Note this
'can also be done in the Properties Editor. I did it this
'way to allow seeing all the initial settings at a glance.

Private Sub Form_Load()
sw = 0
    'Initialize the Waveform Selection option
    
    optSine.Value = True
    
   
    
    Me.Show
    On Local Error Resume Next
    Set ds = dx.DirectSoundCreate("")
    If Err.Number <> 0 Then
        MsgBox "Unable to start DirectSound"
        End
    End If
    ds.SetCooperativeLevel Me.hWnd, DSSCL_PRIORITY
    
    'Set the default startup frequency
    frequency = HScroll1.Value
    'conFreq
    
End Sub
'Cleanup on program exit
Private Sub cmdExit_Click()

    Cleanup
    Unload Me
    
End Sub
'Dispose of the DirectSound Object and its buffer
Private Sub Cleanup()

    If Not (dsBuffer Is Nothing) Then dsBuffer.Stop
    Set dsBuffer = Nothing
    Set ds = Nothing
    Set dx = Nothing
    
End Sub
Private Sub h1_Click()
Dim Msg, Title
Title = "OscillatorPro Ver. 1.0"
Msg = "(c) 2007 PWS software, Tech support scott93727@aol.com"
MsgBox Msg, , Title
End Sub
Private Sub HScroll1_Change()
sw = 1
Text1.Text = HScroll1.Value
frequency = HScroll1.Value
If optSine.Value = True Then
        sineWave
    End If
    
    If optSquare.Value = True Then
        squareWave
    End If
    
    If optTriangle.Value = True Then
        triangleWave
    End If
    
    If optSaw.Value = True Then
        sawWave
    End If
    
    If optNoise = True Then
        noise
    End If
End Sub
Private Sub cmdGenerate_Click()
sw = 1
    If optSine.Value = True Then
        sineWave
    End If
    
    If optSquare.Value = True Then
        squareWave
    End If
    
    If optTriangle.Value = True Then
        triangleWave
    End If
    
    If optSaw.Value = True Then
        sawWave
    End If
    
    If optNoise = True Then
        noise
    End If
    
End Sub
Private Sub sineWave()
    makeFile                                            'Create the file and write header
    bufferptr = 45                                      'Offset to beginning of waveform
        increment = Pi / (sampleRate / frequency)
        For inputValue = 0 To (2 * Pi) Step increment   'Step around the circle
            sample = Int(amplitude * Sin(inputValue))   'Calculate the sample value
            Put #1, bufferptr, sample                   'Write sample value to file
            bufferptr = bufferptr + 1                   'Increment buffer pointer
        Next inputValue                                 'Loop to the next sample
  
    closeFile                                           'Fill in the rest of the file data
                 'and close the file.
    
End Sub
Private Sub squareWave()

    makeFile
    
    bufferptr = 45
    period = (sampleRate / frequency)
    state = 1
        If state = 1 Then                   'Positive half cycle
            For inputValue = 0 To period
                sample = amplitude * state
                Put #1, bufferptr, sample
                bufferptr = bufferptr + 1
            Next inputValue
        End If
        
     state = -1
        If state = -1 Then                  'Negative half cycle
            For inputValue = 0 To period
                sample = amplitude * state
                Put #1, bufferptr, sample
                bufferptr = bufferptr + 1
            Next inputValue
        End If
        
    closeFile
    
End Sub
Private Sub sawWave()

    makeFile
    
    bufferptr = 45
        period = sampleRate / (frequency / 2)
        For inputValue = 0 To period
            sample = Int(2 * amplitude * (inputValue / period))
            Put #1, bufferptr, sample
            bufferptr = bufferptr + 1
        Next inputValue
        
    closeFile
    
End Sub
Private Sub triangleWave()

    makeFile
    
    state = 0
    bufferptr = 45
        period = sampleRate / frequency
        If state = 0 Then
        For inputValue = 0 To period / 2    'Generate Positive Slope
            sample = Int(2 * amplitude * (inputValue / period))
            Put #1, bufferptr, sample
            bufferptr = bufferptr + 1
        Next inputValue
        
        state = 1
        End If
        If state = 1 Then
        For inputValue = 0 To period        'Generate Negative Slope
            sample = Int((amplitude - 2 * amplitude) - 2 * amplitude * (inputValue - period) / period)
            Put #1, bufferptr, sample
            bufferptr = bufferptr + 1
        Next inputValue
        
        state = 2
        End If
        If state = 2 Then
        For inputValue = 0 To period / 2    'Positive Slope to finish cycle
            sample = Int(amplitude + (2 * amplitude * (inputValue / period)))
            Put #1, bufferptr, sample
            bufferptr = bufferptr + 1
        Next inputValue
        End If
        
    closeFile
    
End Sub
Private Sub noise()

    Randomize                               'Seed random # generator
    
    makeFile
    
    bufferptr = 45
    period = sampleRate
        For inputValue = 0 To period        'Create 44,100 random samples
            sample = Rnd(amplitude) * 254
            Put #1, bufferptr, sample
            bufferptr = bufferptr + 1
        Next inputValue
        
    closeFile
    
End Sub
'Create the .wav file and write header data

Private Sub makeFile()

    fileName = App.Path & "\temp.wav"
    On Error Resume Next
    Kill fileName                   'REM this line if file does not exist
    
    Open fileName For Binary Access Write As #1
        Put #1, 1, "RIFF"           '"RIFF" header
        Put #1, 5, CInt(0)          'Filesize - 8, will write later
        Put #1, 9, "WAVEfmt "       '"WAVEfmt " header - not space after fmt
        Put #1, 17, CLng(16)        'Lenth of format data
        Put #1, 21, CInt(1)         'Wave type PCM
        Put #1, 23, CInt(1)         '1 channel
        Put #1, 25, CLng(44100)     '44.1 kHz SampleRate
        Put #1, 29, CLng(88200)     '(SampleRate * BitsPerSample * Channels) / 8
        Put #1, 33, CInt(2)         '(BitsPerSample * Channels) / 8
        Put #1, 35, CInt(16)        'BitsPerSample
        Put #1, 37, "data"          '"data" Chunkheader
        Put #1, 41, CInt(0)         'Filesize - 44, will write later

End Sub
'Get the file length, write it into the header and close the file.
Private Sub closeFile()

    fileSize = LOF(1)
    Put #1, 5, CLng(fileSize - 8)
    Put #1, 41, CLng(fileSize - 44)
    Close #1
    
    Play
    
End Sub

'Define the DirectSound8 buffer, create it and set the play mode

Private Sub Play()

    Dim bufferDesc As DSBUFFERDESC
    bufferDesc.lFlags = DSBCAPS_STATIC Or DSBCAPS_STICKYFOCUS
    fileName = App.Path & "\temp.wav"
    Set dsBuffer = ds.CreateSoundBufferFromFile(fileName, bufferDesc)
    dsBuffer.Play DSBPLAY_LOOPING
    
End Sub
'Stop playing and clear the DirectSound8 buffer
Private Sub cmdStop_Click()
sw = 2
    dsBuffer.Stop
    Set dsBuffer = Nothing
    
End Sub
Private Sub m1_Click()
Dim retval
On Error Resume Next
retval = Shell("c:\windows\sndvol32", 1)
On Error Resume Next
retval = Shell("c:\windows\system32\sndvol32", 1)
End Sub
Private Sub optNoise_Click()
If sw = 1 Then cmdGenerate_Click
End Sub
Private Sub optSaw_Click()
If sw = 1 Then cmdGenerate_Click
End Sub
Private Sub optSine_click()
If sw = 1 Then cmdGenerate_Click
End Sub
Private Sub optSquare_Click()
If sw = 1 Then cmdGenerate_Click
End Sub
Private Sub optTriangle_Click()
If sw = 1 Then cmdGenerate_Click
End Sub
