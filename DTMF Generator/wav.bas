Attribute VB_Name = "modWAV"
'-----------------------------------------------------------------------------------
' modWAV - WAV File Routines
'
' Written By: David M.Hitchner
'             k5dmh@bellsouth.net
'             k5dmh@arrl.net
'
' This VB module is a collection of routines to uses to create a play WAV
' format files.
'-----------------------------------------------------------------------------------
Option Explicit

'-----------------------------------------------------------------------------------
' Wave File Format
'-----------------------------------------------------------------------------------
' RIFF Chunk   ( 12 bytes)
' 00 00 - 03  "RIFF"
' 04 04 - 07  Total Length to Follow  (Length of File - 8)
' 08 08 - 11  "WAVE"
'
' FORMAT Chunk ( 24 bytes )
' 0C 12 - 15  "fmt_"
' 10 16 - 19  Length of FORMAT Chunk  Always 0x10
' 14 20 - 21  Audio Format            Always 0x01
' 16 22 - 23  Channels                1 = Mono, 2 = Stereo
' 18 24 - 27  Sample Rate             In Hertz
' 1C 28 - 31  Bytes per Second        Sample Rate * Channels * Bits per Sample / 8
' 20 32 - 33  Bytes per Sample        Channels * Bits per Sample / 8
'                                       1 = 8 bit Mono
'                                       2 = 8 bit Stereo or 16 bit Mono
'                                       4 = 16 bit Stereo
' 22 34 - 35  Bits per Sample
'
' DATA Chunk
' 24 36 - 39  "data"
' 28 40 - 43  Length of Data          Samples * Channels * Bits per Sample / 8
' 2C 44 - End Data Samples
'              8 Bit = 0 to 255             unsigned bytes
'             16 Bit = -32,768 to 32,767    2's-complement signed integers
'-----------------------------------------------------------------------------------

Public Const MODE_MONO = 0      ' Mono
Public Const MODE_LR = 1        ' Stereo L+R
Public Const MODE_L = 2         ' Stereo L
Public Const MODE_R = 3         ' Stereo R

Public Const RATE_8000 = 8000
Public Const RATE_11025 = 11025
Public Const RATE_22050 = 22050
Public Const RATE_32000 = 32000
Public Const RATE_44100 = 44100
Public Const RATE_48000 = 48000
Public Const RATE_88000 = 88000
Public Const RATE_96000 = 96000

Public Const BITS_8 = 8
Public Const BITS_16 = 16

Public Type SINEWAVE
    dblFrequency As Double
    dblDataSlice As Double
    dblAmplitudeL As Double
    dblAmplitudeR As Double
End Type

Private PI As Double
Private intBits As Integer
Private lngSampleRate As Long
Private intSampleBytes As Integer
Private intAudioMode As Integer
Private dblFrequency As Double
Private dblVolumeL As Double
Private dblVolumeR As Double
Private intAudioWidth As Integer

Private Const SND_ALIAS = &H10000
Private Const SND_ALIAS_ID = &H110000
Private Const SND_ALIAS_START = 0
Private Const SND_APPLICATION = &H80
Private Const SND_ASYNC = &H1
Private Const SND_FILENAME = &H20000
Private Const SND_LOOP = &H8
Private Const SND_MEMORY = &H4
Private Const SND_NODEFAULT = &H2
Private Const SND_NOSTOP = &H10
Private Const SND_NOWAIT = &H2000
Private Const SND_PURGE = &H40
Private Const SND_RESERVED = &HFF000000
Private Const SND_RESOURCE = &H40004
Private Const SND_SYNC = &H0
Private Const SND_TYPE_MASK = &H170007
Private Const SND_VALID = &H1F
Private Const SND_VALIDFLAGS = &H17201F

Private Declare Function PlaySoundFile Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Private Declare Function PlaySoundMemory Lib "winmm.dll" Alias "PlaySoundA" (ptrMemory As Any, ByVal hModule As Long, ByVal dwFlags As Long) As Long

'-----------------------------------------------------------------------------------
' Wav_Play - Plays the wav file from memory.
'-----------------------------------------------------------------------------------
Public Function Wav_Play(WavArray() As Byte) As Boolean
    Dim lngStatus As Long
    
    lngStatus = PlaySoundMemory(WavArray(0), 0, SND_MEMORY Or SND_APPLICATION Or _
        SND_SYNC Or SND_NODEFAULT)

    If lngStatus = 0 Then
        Wav_Play = False
    Else
        Wav_Play = True
    End If
End Function

'-----------------------------------------------------------------------------------
' Wav_BuildHeader - Builds the WAV file header based on the sample rate, resolution,
'                   audio mode.  Also sets the volume level for other routines.
'-----------------------------------------------------------------------------------
Public Sub Wav_BuildHeader(WavArray() As Byte, SampleRate As Long, _
    Resolution As Integer, AudioMode As Integer, VolumeL As Double, VolumeR As Double)
    Dim lngBytesASec As Long
    
    PI = 4# * Atn(1#)
    
    ' Save parameters.
    lngSampleRate = SampleRate
    intBits = Resolution
    intAudioMode = AudioMode
    dblVolumeL = VolumeL
    dblVolumeR = VolumeR
    
    ReDim WavArray(0 To 43)
    
    '-------------------------------------------------------------------------------
    ' Fixed Data
    '-------------------------------------------------------------------------------
    WavArray(0) = 82   ' R
    WavArray(1) = 73   ' I
    WavArray(2) = 70   ' F
    WavArray(3) = 70   ' F
    WavArray(8) = 87   ' W
    WavArray(9) = 65   ' A
    WavArray(10) = 86  ' V
    WavArray(11) = 69  ' E
    WavArray(12) = 102 ' f
    WavArray(13) = 109 ' m
    WavArray(14) = 116 ' t
    WavArray(15) = 32  ' .
    WavArray(16) = 16  ' Length of Format Chunk
    WavArray(17) = 0   ' Length of Format Chunk
    WavArray(18) = 0   ' Length of Format Chunk
    WavArray(19) = 0   ' Length of Format Chunk
    WavArray(20) = 1   ' Audio Format
    WavArray(21) = 0   ' Audio Format
    WavArray(36) = 100 ' d
    WavArray(37) = 97  ' a
    WavArray(38) = 116 ' t
    WavArray(39) = 97  ' a
    
    '-------------------------------------------------------------------------------
    ' Bytes 22 - 23  Channels   1 = Mono, 2 = Stereo
    '-------------------------------------------------------------------------------
    Select Case intAudioMode
        Case MODE_MONO:
            WavArray(22) = 1
            WavArray(23) = 0
            intAudioWidth = 1
        Case MODE_LR:
            WavArray(22) = 2
            WavArray(23) = 0
            intAudioWidth = 2
        Case MODE_L:
            WavArray(22) = 2
            WavArray(23) = 0
            intAudioWidth = 2
        Case MODE_R:
            WavArray(22) = 2
            WavArray(23) = 0
            intAudioWidth = 2
    End Select

    '-------------------------------------------------------------------------------
    ' 24 - 27  Sample Rate             In Hertz
    '-------------------------------------------------------------------------------
    WavArray(24) = ExtractByte(lngSampleRate, 0)
    WavArray(25) = ExtractByte(lngSampleRate, 1)
    WavArray(26) = ExtractByte(lngSampleRate, 2)
    WavArray(27) = ExtractByte(lngSampleRate, 3)

    '-------------------------------------------------------------------------------
    ' Bytes 34 - 35  Bits per Sample
    '-------------------------------------------------------------------------------
    Select Case intBits
        Case 8:
            WavArray(34) = 8
            WavArray(35) = 0
            intSampleBytes = 1
        Case 16:
            WavArray(34) = 16
            WavArray(35) = 0
            intSampleBytes = 2
    End Select
    
    '-------------------------------------------------------------------------------
    ' Bytes 28 - 31  Bytes per Second   Sample Rate * Channels * Bits per Sample / 8
    '-------------------------------------------------------------------------------
    lngBytesASec = lngSampleRate * intAudioWidth * intSampleBytes

    WavArray(28) = ExtractByte(lngBytesASec, 0)
    WavArray(29) = ExtractByte(lngBytesASec, 1)
    WavArray(30) = ExtractByte(lngBytesASec, 2)
    WavArray(31) = ExtractByte(lngBytesASec, 3)
    
    '-------------------------------------------------------------------------------
    ' Bytes 32 - 33 Bytes per Sample     Channels * Bits per Sample / 8
    '                                       1 = 8 bit Mono
    '                                       2 = 8 bit Stereo or 16 bit Mono
    '                                       4 = 16 bit Stereo
    '-------------------------------------------------------------------------------
    If (intAudioMode = MODE_MONO) And (intBits = 8) Then
        WavArray(32) = 1
        WavArray(33) = 0
    End If
    
    If ((intAudioMode = MODE_LR) Or (intAudioMode = MODE_L) Or _
        (intAudioMode = MODE_R)) And (intBits = 8) Then
        WavArray(32) = 2
        WavArray(33) = 0
    End If
    
    If (intAudioMode = MODE_MONO) And (intBits = 16) Then
        WavArray(32) = 2
        WavArray(33) = 0
    End If
    
    If ((intAudioMode = MODE_LR) Or (intAudioMode = MODE_L) Or _
        (intAudioMode = MODE_R)) And (intBits = 16) Then
        WavArray(32) = 4
        WavArray(33) = 0
    End If

End Sub

'-----------------------------------------------------------------------------------
' Wav_MultiSineWave - Builds a complex wave form from one or more sine waves.
'-----------------------------------------------------------------------------------
Public Sub Wav_MultiSineWave(WavArray() As Byte, SineWaves() As SINEWAVE, _
    Seconds As Double)
    
    Dim i As Long, j As Long
    Dim lngLimit As Long
    Dim lngDataL As Long, lngDataR As Long
    Dim dblDataPtL As Double, dblDataPtR As Double
    Dim dblWaveTime As Double
    Dim dblSampleTime As Double
    Dim lngSamples As Long
    Dim lngFileSize As Long, lngDataSize As Long
    
    Dim intSineCount As Integer
    
    intSineCount = UBound(SineWaves)
    
    For i = 1 To intSineCount
        dblWaveTime = 1 / SineWaves(i).dblFrequency
        dblSampleTime = 1 / CDbl(lngSampleRate)
        SineWaves(i).dblDataSlice = (2 * PI) / (dblWaveTime / dblSampleTime)
    Next
    
    lngSamples = CLng(Seconds / dblSampleTime)

    '-------------------------------------------------------------------------------
    ' Bytes 40 - 43  Length of Data   Samples * Channels * Bits per Sample / 8
    '-------------------------------------------------------------------------------
    lngDataSize = lngSamples * intAudioWidth * (intBits / 8)
    ReDim Preserve WavArray(0 To 43 + lngDataSize)

    WavArray(40) = ExtractByte(lngDataSize, 0)
    WavArray(41) = ExtractByte(lngDataSize, 1)
    WavArray(42) = ExtractByte(lngDataSize, 2)
    WavArray(43) = ExtractByte(lngDataSize, 3)

    '-------------------------------------------------------------------------------
    ' Bytes 04 - 07  Total Length to Follow  (Length of File - 8)
    '-------------------------------------------------------------------------------
    lngFileSize = lngDataSize + 36

    WavArray(4) = ExtractByte(lngFileSize, 0)
    WavArray(5) = ExtractByte(lngFileSize, 1)
    WavArray(6) = ExtractByte(lngFileSize, 2)
    WavArray(7) = ExtractByte(lngFileSize, 3)

    '-------------------------------------------------------------------------------
    ' Bytes 44 - End   Data Samples
    '-------------------------------------------------------------------------------
    
    If intBits = 8 Then
        lngLimit = 127
    Else
        lngLimit = 32767
    End If

    For i = 0 To lngSamples - 1

        If intBits = 8 Then
            '-----------------------------------------------------------------------
            ' 8 Bit Data
            '-----------------------------------------------------------------------
            dblDataPtL = 0
            dblDataPtR = 0
            For j = 1 To intSineCount
                dblDataPtL = dblDataPtL + (Sin(i * SineWaves(j).dblDataSlice) * _
                    SineWaves(j).dblAmplitudeL)
                dblDataPtR = dblDataPtR + (Sin(i * SineWaves(j).dblDataSlice) * _
                    SineWaves(j).dblAmplitudeR)
            Next
                
            lngDataL = Int(dblDataPtL * dblVolumeL * lngLimit) + lngLimit
            lngDataR = Int(dblDataPtL * dblVolumeR * lngLimit) + lngLimit
            
            If intAudioMode = MODE_MONO Then _
                WavArray(i + 44) = ExtractByte(lngDataL, 0)

            If intAudioMode = MODE_LR Then       'L+R stereo
                WavArray((2 * i) + 44) = ExtractByte(lngDataL, 0)
                WavArray((2 * i) + 45) = ExtractByte(lngDataR, 0)
            End If

            If intAudioMode = MODE_L Then       ' L only stereo
                WavArray((2 * i) + 44) = ExtractByte(lngDataL, 0)
                WavArray((2 * i) + 45) = 0
            End If

            If intAudioMode = MODE_R Then       ' R only stereo
                WavArray((2 * i) + 44) = 0
                WavArray((2 * i) + 45) = ExtractByte(lngDataR, 0)
            End If

        Else
            
            '-----------------------------------------------------------------------
            ' 16 Bit Data
            '-----------------------------------------------------------------------
            dblDataPtL = 0
            dblDataPtR = 0
            For j = 1 To intSineCount
                dblDataPtL = dblDataPtL + (Sin(i * SineWaves(j).dblDataSlice) * _
                    SineWaves(j).dblAmplitudeL)
                dblDataPtR = dblDataPtR + (Sin(i * SineWaves(j).dblDataSlice) * _
                    SineWaves(j).dblAmplitudeR)
            Next
                
            lngDataL = Int(dblDataPtL * dblVolumeL * lngLimit)
            lngDataR = Int(dblDataPtL * dblVolumeR * lngLimit)
    
            If intAudioMode = MODE_MONO Then
                WavArray((2 * i) + 44) = ExtractByte(lngDataL, 0)
                WavArray((2 * i) + 45) = ExtractByte(lngDataL, 1)
            End If
        
            If intAudioMode = MODE_LR Then
                WavArray((4 * i) + 44) = ExtractByte(lngDataL, 0)
                WavArray((4 * i) + 45) = ExtractByte(lngDataL, 1)
                WavArray((4 * i) + 46) = ExtractByte(lngDataR, 0)
                WavArray((4 * i) + 47) = ExtractByte(lngDataR, 1)
            End If
        
            If intAudioMode = MODE_L Then
                WavArray((4 * i) + 44) = ExtractByte(lngDataL, 0)
                WavArray((4 * i) + 45) = ExtractByte(lngDataL, 1)
                WavArray((4 * i) + 46) = 0
                WavArray((4 * i) + 47) = 0
            End If
        
            If intAudioMode = MODE_R Then
                WavArray((4 * i) + 44) = 0
                WavArray((4 * i) + 45) = 0
                WavArray((4 * i) + 46) = ExtractByte(lngDataR, 0)
                WavArray((4 * i) + 47) = ExtractByte(lngDataR, 1)
            End If
        
        End If
    
    Next
End Sub
'-------------------------------------------------------------------------------------
' ExtractByte - Extracts the high or low byte from a short (16 bit) VB integer.
'
'   intWord     - VB Integer from which to extract byte.
'   intByte     - Returned high or low byte.
'   intPosition - |                    Word                   |
'                 | Byte = 3 | Byte = 2 | Byte = 1 | Byte = 0 |
'-------------------------------------------------------------------------------------
Private Function ExtractByte(lngWord As Long, intPosition As Integer) As Byte
    Dim lngTemp As Long
    Dim intByte As Integer
    
    If intPosition = 3 Then
        ' Byte 2
        lngTemp = lngWord
        
        ' Mask off byte and shift right 24 bits.
        '   Mask  -> 2130706432 = &H7F000000
        '   Shift -> Divide by 16777216
        lngTemp = (lngTemp And 2130706432) / 16777216
        
        ' Cast back to integer.
        intByte = lngTemp
    
    ElseIf intPosition = 2 Then
        ' Byte 2
        lngTemp = lngWord
        
        ' Mask off byte and shift right 16 bits.
        '   Mask  -> 16711680 = &HFF0000
        '   Shift -> Divide by 65536
        lngTemp = (lngTemp And 16711680) / 65536
        
        ' Cast back to integer.
        intByte = lngTemp
    
    ElseIf intPosition = 1 Then
        ' Byte 1
        lngTemp = lngWord
        
        ' Mask off high byte and shift right 8 bits.
        '   Mask  -> 65290 = &HFF00
        '   Shift -> Divide by 256
        lngTemp = (lngTemp And 65290) / 256
        
        ' Cast back to integer.
        intByte = lngTemp
    Else
        ' Byte 0
        intByte = lngWord And &HFF
    End If
    
    ExtractByte = intByte
End Function

'-----------------------------------------------------------------------------------
' Wav_Stop - Stop the currently playing wav.
'-----------------------------------------------------------------------------------
Public Sub Wav_Stop()
    Dim lngStatus As Long

    lngStatus = PlaySoundMemory(ByVal 0&, ByVal 0&, SND_PURGE Or SND_NODEFAULT)
End Sub



