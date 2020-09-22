VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "DTMF Detector"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4485
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   4485
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Timer tmr 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3750
      Top             =   1725
   End
   Begin MSComctlLib.StatusBar sbar 
      Align           =   2  'Unten ausrichten
      Height          =   240
      Left            =   0
      TabIndex        =   10
      Top             =   2970
      Width           =   4485
      _ExtentX        =   7911
      _ExtentY        =   423
      Style           =   1
      SimpleText      =   "Ready"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Frame frmNumber 
      Caption         =   "Dialed:"
      Height          =   915
      Left            =   225
      TabIndex        =   8
      Top             =   1875
      Width           =   3990
      Begin VB.Label lblNumber 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   150
         TabIndex        =   9
         Top             =   300
         Width           =   3690
      End
   End
   Begin VB.CommandButton cmdStop 
      Cancel          =   -1  'True
      Caption         =   "Stop Recording"
      Enabled         =   0   'False
      Height          =   345
      Left            =   2925
      TabIndex        =   7
      Top             =   1350
      Width           =   1290
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start Recording"
      Default         =   -1  'True
      Height          =   345
      Left            =   1575
      TabIndex        =   6
      Top             =   1350
      Width           =   1290
   End
   Begin MSComctlLib.Slider sldVol 
      Height          =   270
      Left            =   975
      TabIndex        =   5
      Top             =   900
      Width           =   3240
      _ExtentX        =   5715
      _ExtentY        =   476
      _Version        =   393216
      Max             =   65535
      TickStyle       =   3
   End
   Begin VB.ComboBox cboRecLine 
      Height          =   315
      Left            =   975
      Style           =   2  'Dropdown-Liste
      TabIndex        =   3
      Top             =   525
      Width           =   3240
   End
   Begin VB.ComboBox cboRecDev 
      Height          =   315
      Left            =   975
      Style           =   2  'Dropdown-Liste
      TabIndex        =   1
      Top             =   150
      Width           =   3240
   End
   Begin VB.Label lblVolume 
      Caption         =   "Volume:"
      Height          =   240
      Left            =   225
      TabIndex        =   4
      Top             =   900
      Width           =   765
   End
   Begin VB.Label lblLine 
      Caption         =   "Line:"
      Height          =   240
      Left            =   225
      TabIndex        =   2
      Top             =   525
      Width           =   690
   End
   Begin VB.Label lblDev 
      Caption         =   "Device:"
      Height          =   240
      Left            =   225
      TabIndex        =   0
      Top             =   150
      Width           =   840
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const samplerate            As Long = 22050
Private Const Channels              As Long = 1
Private Const SILENCE_THRESHOLD     As Long = 30
Private Const BUFSILENCE_THRESHOLD  As Long = 60

Private Const BUFFER_LENGTH         As Long = 10

Private Const PI                    As Single = 3.14159265358979
Private Const PI2                   As Single = PI * 2
Private Const LOG10                 As Single = 0.434294481903251
Private Const NODIVZ                As Single = 0.000000000000001

Private DTMF_F1()                   As Single
Private DTMF_F2()                   As Single
Private DTMF_NUM()                  As Single

Private WithEvents m_clsRecorder    As WaveInRecorder
Attribute m_clsRecorder.VB_VarHelpID = -1

Private m_lngRecorded               As Long
Private m_blnGotSilence             As Boolean

' Dual Tone Multi-Frequency (DTMF) Tabelle (ITU-T Q.23)
'
' Hz          1209    1336    1477    1633
' 697          1       2       3       A
' 770          4       5       6       B
' 852          7       8       9       C
' 941          *       0       #       D
'
Private Sub FillDTMFTable()
    Dim i   As Integer
    Dim j   As Integer
    Dim k   As Integer

    ReDim DTMF_F1(3) As Single
    ReDim DTMF_F2(3) As Single
    ReDim DTMF_NUM(15, 3) As Single

    DTMF_F1(0) = 697: DTMF_F2(0) = 1209
    DTMF_F1(1) = 770: DTMF_F2(1) = 1336
    DTMF_F1(2) = 852: DTMF_F2(2) = 1477
    DTMF_F1(3) = 941: DTMF_F2(3) = 1633

    ReDim DTMF_NUM(15, 2) As Single

    For i = 0 To UBound(DTMF_F1)
        For j = 0 To UBound(DTMF_F2)
            DTMF_NUM(k, 0) = DTMF_F1(i)
            DTMF_NUM(k, 1) = DTMF_F2(j)
            k = k + 1
        Next
    Next

    DTMF_NUM(0, 2) = Asc("1")
    DTMF_NUM(1, 2) = Asc("2")
    DTMF_NUM(2, 2) = Asc("3")
    DTMF_NUM(3, 2) = Asc("A")
    DTMF_NUM(4, 2) = Asc("4")
    DTMF_NUM(5, 2) = Asc("5")
    DTMF_NUM(6, 2) = Asc("6")
    DTMF_NUM(7, 2) = Asc("B")
    DTMF_NUM(8, 2) = Asc("7")
    DTMF_NUM(9, 2) = Asc("8")
    DTMF_NUM(10, 2) = Asc("9")
    DTMF_NUM(11, 2) = Asc("C")
    DTMF_NUM(12, 2) = Asc("*")
    DTMF_NUM(13, 2) = Asc("0")
    DTMF_NUM(14, 2) = Asc("#")
    DTMF_NUM(15, 2) = Asc("D")
End Sub

Private Sub cboRecDev_Click()
    Dim i   As Long
    
    cboRecLine.Clear
    
    If Not m_clsRecorder.SelectDevice(cboRecDev.ListIndex) Then
        MsgBox "Couldn't select device!", vbExclamation
        Exit Sub
    End If
    
    For i = 0 To m_clsRecorder.MixerLineCount - 1
        cboRecLine.AddItem m_clsRecorder.MixerLineName(i)
    Next
    cboRecLine.ListIndex = 0
End Sub

Private Sub cboRecLine_Click()
    If Not m_clsRecorder.SelectMixerLine(cboRecLine.ListIndex) Then
        MsgBox "Couldn't select Mixer Line!", vbExclamation
        Exit Sub
    End If
    
    sldVol.value = m_clsRecorder.MixerLineVolume
End Sub

Private Sub cmdStart_Click()
    lblNumber.Caption = ""
    
    If Not m_clsRecorder.StartRecord(samplerate, Channels) Then
        MsgBox "Couldn't start recording with " & samplerate & " Hz, " & Channels & " Channels!", vbExclamation
        Exit Sub
    End If
    
    tmr.Enabled = True
    
    cmdStart.Enabled = False
    cmdStop.Enabled = True
End Sub

Private Sub cmdStop_Click()
    If Not m_clsRecorder.StopRecord() Then
        MsgBox "Couldn't properly stop recording!"
    End If
    
    m_lngRecorded = 0
    tmr.Enabled = False
    sbar.SimpleText = "Ready"
    
    cmdStart.Enabled = True
    cmdStop.Enabled = False
End Sub

Private Sub Form_Load()
    Dim i   As Long
    
    Set m_clsRecorder = New WaveInRecorder
    
    ' 10 ms buffersize so we can detect silence between numbers
    m_clsRecorder.BufferSize = MSToBytes(BUFFER_LENGTH)
    
    For i = 0 To m_clsRecorder.DeviceCount - 1
        cboRecDev.AddItem m_clsRecorder.DeviceName(i)
    Next
    cboRecDev.ListIndex = 0
    
    FillDTMFTable
End Sub

Private Sub Form_Unload(Cancel As Integer)
    m_clsRecorder.StopRecord
End Sub

Private Sub m_clsRecorder_GotData(intBuffer() As Integer, lngLen As Long)
    Dim lngAmplitude    As Single, lngPartCount     As Long
    Dim lngMaxAmplitude As Single, lngMaxAmplitude2 As Single
    Dim sngMaxF1        As Single, sngMaxF2         As Single
    Dim sngSamples()    As Single
    Dim i               As Long
    
    m_lngRecorded = m_lngRecorded + BytesToMS(lngLen)
    
    ReDim sngSamples(lngLen \ 2 - 1) As Single
    
    For i = 0 To lngLen \ 2 - 1
        ' normalize all 16 bit samples to floats ([-1;+1])
        sngSamples(i) = intBuffer(i) / 32768
        
        ' sum up all the positive amplitudes in the signal buffer
        If intBuffer(i) > 0 Then
            lngMaxAmplitude = lngMaxAmplitude + intBuffer(i)
            lngPartCount = lngPartCount + 1
        End If
    Next
    
    ' no positive amplitudes?
    If lngPartCount = 0 Then Exit Sub
    
    ' get the power of the average amplitude of the signal to detect silence
    If power(lngMaxAmplitude / lngPartCount) < BUFSILENCE_THRESHOLD Then
        m_blnGotSilence = True
        Exit Sub
    End If
    lngMaxAmplitude = 0

    ' find the first DTMF frequency
    For i = 0 To UBound(DTMF_F1)
        lngAmplitude = power(Goertzel(sngSamples, 0, lngLen \ 2, DTMF_F1(i), samplerate))

        If lngAmplitude > lngMaxAmplitude Then
            lngMaxAmplitude = lngAmplitude
            sngMaxF1 = DTMF_F1(i)
        End If
    Next

    ' find the second DTMF frequency
    For i = 0 To UBound(DTMF_F2)
        lngAmplitude = power(Goertzel(sngSamples, 0, lngLen \ 2, DTMF_F2(i), samplerate))

        If lngAmplitude > lngMaxAmplitude2 Then
            lngMaxAmplitude2 = lngAmplitude
            sngMaxF2 = DTMF_F2(i)
        End If
    Next
    
    ' if we just had silence then this could be valid
    If m_blnGotSilence Then
        ' check if the found frequencies are powerful
        If lngMaxAmplitude > SILENCE_THRESHOLD And lngMaxAmplitude2 > SILENCE_THRESHOLD Then
            ' get the sign the frequencies match with
            For i = 0 To UBound(DTMF_NUM)
                If DTMF_NUM(i, 0) = sngMaxF1 Then
                    If DTMF_NUM(i, 1) = sngMaxF2 Then
                        lblNumber.Caption = lblNumber.Caption & Chr$(DTMF_NUM(i, 2))
                        Exit For
                    End If
                End If
            Next
            ' expect silence
            m_blnGotSilence = False
        End If
    End If
End Sub

Private Sub sldVol_Scroll()
    m_clsRecorder.MixerLineVolume = sldVol.value
End Sub

Private Function BytesToMS(ByVal bytes As Long) As Long
    BytesToMS = bytes / (samplerate * Channels * 2) * 1000
End Function

Private Function MSToBytes(ByVal ms As Long) As Long
    MSToBytes = ms / 1000 * (samplerate * Channels * 2)
End Function

Private Sub tmr_Timer()
    sbar.SimpleText = "Recording... " & FmtTime(m_lngRecorded)
End Sub

Private Function FmtTime(ByVal ms As Long) As String
    FmtTime = ((ms / 1000) \ 60) & ":" & Format((ms / 1000) Mod 60, "00")
End Function

' amplitude to Decibel
Private Function power(ByVal value As Single) As Single
    power = 20 * Log(Abs(value) + NODIVZ) * LOG10
End Function

' like a Fourier Transformation for 1 frequency
'
' source:
' http://www.musicdsp.org/archive.php?classid=0#107
Function Goertzel(sngData() As Single, ByVal S As Long, ByVal N As Long, ByVal freq As Single, ByVal sampr As Long) As Single
    Dim Skn     As Single
    Dim Skn1    As Single
    Dim Skn2    As Single
    Dim c       As Single
    Dim c2      As Single
    Dim i       As Long

    c = PI2 * freq / sampr
    c2 = Cos(c)

    For i = S To S + N - 1
        Skn2 = Skn1
        Skn1 = Skn
        Skn = 2 * c2 * Skn1 - Skn2 + sngData(i)
    Next

    Goertzel = Skn - Exp(-c) * Skn1
End Function
