Attribute VB_Name = "modBox"
Option Explicit

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Type DTMF
    bytTones() As Byte
End Type

Dim bytSound() As Byte

Public Sub BoxTone(Freq1 As Integer, Freq2 As Integer, Duration As Double)

Dim intAudioMode As Integer
Dim lngSampleRate As Long
Dim intBits As Integer
Dim intIndex As Integer
Dim udtSineWaves(1 To 2) As SINEWAVE
 
Dim udtBOX(16) As DTMF
  
intAudioMode = 1
lngSampleRate = 8000
intBits = 16
Wav_Stop

With udtSineWaves(1)
    .dblAmplitudeL = 0.25
    .dblAmplitudeR = 0.25
End With
        
With udtSineWaves(2)
    .dblAmplitudeL = 0.25
    .dblAmplitudeR = 0.25
End With
    
udtSineWaves(1).dblFrequency = Freq1
    
udtSineWaves(2).dblFrequency = Freq2
    
Wav_BuildHeader udtBOX(0).bytTones, lngSampleRate, intBits, intAudioMode, 0.5, 0.5
Wav_MultiSineWave udtBOX(0).bytTones, udtSineWaves, Duration

Wav_Play udtBOX(0).bytTones

End Sub
Public Sub Silence(Duration As Double)
Sleep (Duration * 1000)
End Sub
