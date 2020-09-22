VERSION 5.00
Begin VB.Form frmTones 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "DTMF Tones Generator"
   ClientHeight    =   4935
   ClientLeft      =   5610
   ClientTop       =   1560
   ClientWidth     =   7815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   7815
   Begin VB.CommandButton cmdABCD 
      Caption         =   "What to use the ABCD for?"
      Height          =   315
      Left            =   60
      TabIndex        =   58
      Top             =   4500
      Width           =   2295
   End
   Begin VB.CommandButton cmdWhat 
      Caption         =   "What is This?"
      Height          =   315
      Left            =   60
      TabIndex        =   25
      Top             =   3420
      Width           =   2295
   End
   Begin VB.CommandButton cmdCred 
      Caption         =   "> >     Credits     < <"
      Height          =   315
      Left            =   60
      TabIndex        =   24
      Top             =   3960
      Width           =   2295
   End
   Begin VB.Frame Frame2 
      Caption         =   "Boxes"
      Height          =   4815
      Left            =   2460
      TabIndex        =   20
      Top             =   60
      Width           =   5295
      Begin VB.Frame frmBox 
         Caption         =   "What are the ABCD tones for?"
         Height          =   4095
         Index           =   2
         Left            =   120
         TabIndex        =   59
         Top             =   600
         Visible         =   0   'False
         Width           =   5055
         Begin VB.Label lblABCD 
            Height          =   3495
            Left            =   180
            TabIndex        =   60
            Top             =   420
            Width           =   4695
         End
      End
      Begin VB.ComboBox cboBoxes 
         Height          =   315
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   1875
      End
      Begin VB.Frame frmBox 
         Height          =   4095
         Index           =   0
         Left            =   120
         TabIndex        =   22
         Top             =   600
         Visible         =   0   'False
         Width           =   5055
         Begin VB.Frame Frame4 
            Caption         =   "Canada"
            Height          =   1575
            Left            =   120
            TabIndex        =   33
            Top             =   2400
            Width           =   4815
            Begin VB.CommandButton cmdCANq 
               Caption         =   "Quarter"
               Height          =   315
               Left            =   120
               TabIndex        =   36
               Top             =   1140
               Width           =   735
            End
            Begin VB.CommandButton cmdCANd 
               Caption         =   "Dime"
               Height          =   315
               Left            =   120
               TabIndex        =   35
               Top             =   720
               Width           =   735
            End
            Begin VB.CommandButton cmdCANn 
               Caption         =   "Nickel"
               Height          =   315
               Left            =   120
               TabIndex        =   34
               Top             =   300
               Width           =   735
            End
            Begin VB.Label Label7 
               Caption         =   "5 x [33ms(2200), 33ms(Silence)]"
               Height          =   255
               Left            =   1080
               TabIndex        =   39
               Top             =   1200
               Width           =   2355
            End
            Begin VB.Label Label6 
               Caption         =   "2 x [60ms(2200), 60ms(Silence)]"
               Height          =   255
               Left            =   1080
               TabIndex        =   38
               Top             =   780
               Width           =   2295
            End
            Begin VB.Label Label4 
               Caption         =   "60ms(2200), 60ms(Silence)"
               Height          =   255
               Left            =   1020
               TabIndex        =   37
               Top             =   360
               Width           =   2295
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "US"
            Height          =   2175
            Left            =   120
            TabIndex        =   26
            Top             =   120
            Width           =   4815
            Begin VB.CommandButton cmdUSq 
               Caption         =   "Quarter"
               Height          =   315
               Left            =   120
               TabIndex        =   29
               Top             =   1560
               Width           =   735
            End
            Begin VB.CommandButton cmdUSd 
               Caption         =   "Dime"
               Height          =   315
               Left            =   120
               TabIndex        =   28
               Top             =   840
               Width           =   735
            End
            Begin VB.CommandButton cmdUSn 
               Caption         =   "Nickel"
               Height          =   315
               Left            =   120
               TabIndex        =   27
               Top             =   300
               Width           =   735
            End
            Begin VB.Label Label1 
               Caption         =   "35-160ms(1700hz+2200hz), 240ms(Silence) "
               Height          =   255
               Left            =   960
               TabIndex        =   32
               Top             =   360
               Width           =   3615
            End
            Begin VB.Label Label2 
               Caption         =   "2 x [35-160ms(1700hz+2200hz), 20-110ms(Silence)] 165ms(Silence)"
               Height          =   435
               Left            =   960
               TabIndex        =   31
               Top             =   780
               Width           =   3795
            End
            Begin VB.Label Label3 
               Caption         =   $"Tones.frx":0000
               Height          =   615
               Left            =   960
               TabIndex        =   30
               Top             =   1380
               Width           =   3615
            End
         End
      End
      Begin VB.Frame frmBox 
         Height          =   4095
         Index           =   1
         Left            =   120
         TabIndex        =   23
         Top             =   600
         Visible         =   0   'False
         Width           =   5055
         Begin VB.Frame Frame5 
            Caption         =   "Domestic"
            Height          =   2655
            Left            =   120
            TabIndex        =   42
            Top             =   600
            Width           =   3075
            Begin VB.CommandButton cmdBlue 
               Caption         =   "ST"
               Height          =   375
               Index           =   14
               Left            =   1860
               TabIndex        =   57
               Top             =   2040
               Width           =   1035
            End
            Begin VB.CommandButton cmdBlue 
               Caption         =   "ST2P"
               Height          =   375
               Index           =   13
               Left            =   1860
               TabIndex        =   56
               Top             =   1620
               Width           =   1035
            End
            Begin VB.CommandButton cmdBlue 
               Caption         =   "KP"
               Height          =   375
               Index           =   12
               Left            =   1860
               TabIndex        =   55
               Top             =   1200
               Width           =   1035
            End
            Begin VB.CommandButton cmdBlue 
               Caption         =   "STP"
               Height          =   375
               Index           =   11
               Left            =   1860
               TabIndex        =   54
               Top             =   780
               Width           =   1035
            End
            Begin VB.CommandButton cmdBlue 
               Caption         =   "ST3P"
               Height          =   375
               Index           =   10
               Left            =   1860
               TabIndex        =   53
               Top             =   360
               Width           =   1035
            End
            Begin VB.CommandButton cmdBlue 
               Caption         =   "9"
               Height          =   375
               Index           =   9
               Left            =   1020
               TabIndex        =   52
               Top             =   1200
               Width           =   375
            End
            Begin VB.CommandButton cmdBlue 
               Caption         =   "8"
               Height          =   375
               Index           =   8
               Left            =   600
               TabIndex        =   51
               Top             =   1200
               Width           =   375
            End
            Begin VB.CommandButton cmdBlue 
               Caption         =   "7"
               Height          =   375
               Index           =   7
               Left            =   180
               TabIndex        =   50
               Top             =   1200
               Width           =   375
            End
            Begin VB.CommandButton cmdBlue 
               Caption         =   "6"
               Height          =   375
               Index           =   6
               Left            =   1020
               TabIndex        =   49
               Top             =   780
               Width           =   375
            End
            Begin VB.CommandButton cmdBlue 
               Caption         =   "5"
               Height          =   375
               Index           =   5
               Left            =   600
               TabIndex        =   48
               Top             =   780
               Width           =   375
            End
            Begin VB.CommandButton cmdBlue 
               Caption         =   "4"
               Height          =   375
               Index           =   4
               Left            =   180
               TabIndex        =   47
               Top             =   780
               Width           =   375
            End
            Begin VB.CommandButton cmdBlue 
               Caption         =   "3"
               Height          =   375
               Index           =   3
               Left            =   1020
               TabIndex        =   46
               Top             =   360
               Width           =   375
            End
            Begin VB.CommandButton cmdBlue 
               Caption         =   "2"
               Height          =   375
               Index           =   2
               Left            =   600
               TabIndex        =   45
               Top             =   360
               Width           =   375
            End
            Begin VB.CommandButton cmdBlue 
               Caption         =   "1"
               Height          =   375
               Index           =   1
               Left            =   180
               TabIndex        =   44
               Top             =   360
               Width           =   375
            End
            Begin VB.CommandButton cmdBlue 
               Caption         =   "0"
               Height          =   375
               Index           =   0
               Left            =   600
               TabIndex        =   43
               Top             =   1620
               Width           =   375
            End
         End
         Begin VB.CommandButton cmdBlue 
            Caption         =   "International"
            Height          =   375
            Index           =   16
            Left            =   1140
            TabIndex        =   41
            Top             =   180
            Width           =   1095
         End
         Begin VB.CommandButton cmdBlue 
            Caption         =   "Domestic"
            Height          =   375
            Index           =   15
            Left            =   120
            TabIndex        =   40
            Top             =   180
            Width           =   915
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Phone DTMF Tones"
      Height          =   3195
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   2295
      Begin VB.CommandButton cmdDial 
         Caption         =   "Dial"
         Height          =   315
         Left            =   1440
         TabIndex        =   18
         Top             =   720
         Width           =   675
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   120
         TabIndex        =   17
         Top             =   300
         Width           =   2055
      End
      Begin VB.CommandButton cmdTone 
         Caption         =   "1"
         Height          =   435
         Index           =   1
         Left            =   180
         TabIndex        =   16
         Top             =   1140
         Width           =   435
      End
      Begin VB.CommandButton cmdTone 
         Caption         =   "2"
         Height          =   435
         Index           =   2
         Left            =   660
         TabIndex        =   15
         Top             =   1140
         Width           =   435
      End
      Begin VB.CommandButton cmdTone 
         Caption         =   "3"
         Height          =   435
         Index           =   3
         Left            =   1140
         TabIndex        =   14
         Top             =   1140
         Width           =   435
      End
      Begin VB.CommandButton cmdTone 
         Caption         =   "A"
         Height          =   435
         Index           =   4
         Left            =   1680
         TabIndex        =   13
         Top             =   1140
         Width           =   435
      End
      Begin VB.CommandButton cmdTone 
         Caption         =   "4"
         Height          =   435
         Index           =   5
         Left            =   180
         TabIndex        =   12
         Top             =   1620
         Width           =   435
      End
      Begin VB.CommandButton cmdTone 
         Caption         =   "5"
         Height          =   435
         Index           =   6
         Left            =   660
         TabIndex        =   11
         Top             =   1620
         Width           =   435
      End
      Begin VB.CommandButton cmdTone 
         Caption         =   "6"
         Height          =   435
         Index           =   7
         Left            =   1140
         TabIndex        =   10
         Top             =   1620
         Width           =   435
      End
      Begin VB.CommandButton cmdTone 
         Caption         =   "B"
         Height          =   435
         Index           =   8
         Left            =   1680
         TabIndex        =   9
         Top             =   1620
         Width           =   435
      End
      Begin VB.CommandButton cmdTone 
         Caption         =   "7"
         Height          =   435
         Index           =   9
         Left            =   180
         TabIndex        =   8
         Top             =   2100
         Width           =   435
      End
      Begin VB.CommandButton cmdTone 
         Caption         =   "8"
         Height          =   435
         Index           =   10
         Left            =   660
         TabIndex        =   7
         Top             =   2100
         Width           =   435
      End
      Begin VB.CommandButton cmdTone 
         Caption         =   "9"
         Height          =   435
         Index           =   11
         Left            =   1140
         TabIndex        =   6
         Top             =   2100
         Width           =   435
      End
      Begin VB.CommandButton cmdTone 
         Caption         =   "C"
         Height          =   435
         Index           =   12
         Left            =   1680
         TabIndex        =   5
         Top             =   2100
         Width           =   435
      End
      Begin VB.CommandButton cmdTone 
         Caption         =   "*"
         Height          =   435
         Index           =   13
         Left            =   180
         TabIndex        =   4
         Top             =   2580
         Width           =   435
      End
      Begin VB.CommandButton cmdTone 
         Caption         =   "0"
         Height          =   435
         Index           =   14
         Left            =   660
         TabIndex        =   3
         Top             =   2580
         Width           =   435
      End
      Begin VB.CommandButton cmdTone 
         Caption         =   "#"
         Height          =   435
         Index           =   15
         Left            =   1140
         TabIndex        =   2
         Top             =   2580
         Width           =   435
      End
      Begin VB.CommandButton cmdTone 
         Caption         =   "D"
         Height          =   435
         Index           =   16
         Left            =   1680
         TabIndex        =   1
         Top             =   2580
         Width           =   435
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fest Einfach
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   660
         Width           =   105
      End
   End
End
Attribute VB_Name = "frmTones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bytSound() As Byte
Private Type DTMF
    bytTones() As Byte
End Type
Dim cboBoxOLD

Private Sub cboBoxes_Click()
frmBox(2).Visible = False
frmBox(cboBoxOLD).Visible = False
frmBox(cboBoxes.ListIndex).Visible = True
cboBoxOLD = cboBoxes.ListIndex
End Sub

Private Sub cmdABCD_Click()
frmBox(2).Visible = True
lblABCD.Caption = "The ABCD tones are simply additional DTFM tones that may be used in any way the standard (0-9) tones are used." & vbNewLine _
& "The ABCD tones are used in the U.S. military telephone network (AutoVon), in some Automatic Call Distributor (ACD) systems, for control messages in some PBX systems, and in some amateur radio auto-patches." & vbNewLine _
& "In the AutoVon network, special telephones are equipped with ABCD keys. The ABCD keys are defined as such:" & vbNewLine _
& vbNewLine _
& "A - Flash" & vbNewLine _
& "B - Flash override priority" & vbNewLine _
& "C - Priority communication" & vbNewLine _
& "D - Priority override" & vbNewLine _
& vbNewLine _
& "Using a built-in maintenance mode of the Automatic Call Distributor (ACD) systems once used by Directory Assistance operators, you could connect two callers together."


End Sub

Private Sub cmdBlue_Click(Index As Integer)
'Frequencies(Hz) Domestic International
'1300+1500          0           0
'700+900            1           1
'700+1100           2           2
'900+1100           3           3
'700+1300           4           4
'900+1300           5           5
'1100+1300          6           6
'700+1500           7           7
'900+1500           8           8
'1100+1500          9           9
'700+1700           ST3p        Code 1
'900+1700           STp         Code 1
'1100+1700          KP          KP1
'1300+1700          ST2p        KP2
'1500+1700          ST          ST

Select Case Index
    Case 0
        BoxTone 1300, 1500, 0.25
        
    Case 1
        BoxTone 700, 900, 0.25
        
    Case 2
        BoxTone 700, 1100, 0.25
        
    Case 3
        BoxTone 900, 1100, 0.25
        
    Case 4
        BoxTone 700, 1300, 0.25
        
    Case 5
        BoxTone 900, 1300, 0.25
        
    Case 6
        BoxTone 1100, 1300, 0.25
        
    Case 7
        BoxTone 700, 1500, 0.25
        
    Case 8
        BoxTone 900, 1500, 0.25
        
    Case 9
        BoxTone 1100, 1500, 0.25
        
    Case 10
        BoxTone 700, 1700, 0.25
        
    Case 11
        BoxTone 900, 1700, 0.25
        
    Case 12
        BoxTone 1100, 1700, 0.25
        
    Case 13
        BoxTone 1300, 1700, 0.25
        
    Case 14
        BoxTone 1500, 1700, 0.25
        
    Case 15
        Frame5.Caption = "Domestic"
        cmdBlue(10).Caption = "ST3P"
        cmdBlue(11).Caption = "STP"
        cmdBlue(12).Caption = "KP"
        cmdBlue(13).Caption = "ST2P"
    
    Case 16
        Frame5.Caption = "International"
        cmdBlue(10).Caption = "Code1"
        cmdBlue(11).Caption = "Code"
        cmdBlue(12).Caption = "KP1"
        cmdBlue(13).Caption = "KP2"

End Select
End Sub

Private Sub cmdCANd_Click()
    BoxTone 2200, 2200, 0.06
    Silence 0.06
    BoxTone 2200, 2200, 0.06
    Silence 0.06
End Sub

Private Sub cmdCANn_Click()
    BoxTone 2200, 2200, 0.06
    Silence 0.06
End Sub

Private Sub cmdCANq_Click()
Dim i
For i = 1 To 5
    BoxTone 2200, 2200, 0.033
    Silence 0.033
Next i
End Sub

Private Sub cmdCred_Click()
MsgBox "modWAV - WAV File Routines" & vbNewLine _
                                    & vbNewLine & _
       "Written By: David M.Hitchner" & vbNewLine & _
       "             k5dmh@bellsouth.net" & vbNewLine & _
       "             k5dmh@arrl.net" & vbNewLine _
                                    & vbNewLine & _
       "UI - User Interface" & vbNewLine _
                                    & vbNewLine & _
       "Written By: Bobbek" & vbNewLine & _
       "             sug_bollarna@hotmail.com" & vbNewLine _
                                    & vbNewLine & _
       "Other Information" & vbNewLine _
                                    & vbNewLine & _
       "Boxing Freq. from:" & vbNewLine & _
       "             www.hackfaq.org", vbOKOnly, frmTones.Caption

End Sub

Private Sub cmdTone_Click(Index As Integer)

    Dim intAudioMode As Integer
    Dim lngSampleRate As Long
    Dim intBits As Integer
    Dim intIndex As Integer
    Dim udtSineWaves(1 To 2) As SINEWAVE
    
    Dim udtDTMF(16) As DTMF
    
    intAudioMode = 1
    lngSampleRate = 8000
    intBits = 16
    Wav_Stop

    '---------------------------------------------------
    '              DTMF Tones
    ' Freq  1209  1336  1477  1633
    '  697    1     2     3     A
    '  770    4     5     6     B
    '  852    7     8     9     C
    '  941    *     0     #     D
    '---------------------------------------------------
    
        With udtSineWaves(1)
            .dblAmplitudeL = 0.25
            .dblAmplitudeR = 0.25
        End With
        
        With udtSineWaves(2)
            .dblAmplitudeL = 0.25
            .dblAmplitudeR = 0.25
        End With
    
        udtSineWaves(1).dblFrequency = Choose(Index, _
            697, 697, 697, 697, _
            770, 770, 770, 770, _
            852, 852, 852, 852, _
            941, 941, 941, 941)
    
        udtSineWaves(2).dblFrequency = Choose(Index, _
            1209, 1336, 1477, 1633, _
            1209, 1336, 1477, 1633, _
            1209, 1336, 1477, 1633, _
            1209, 1336, 1477, 1633)
    
        Wav_BuildHeader udtDTMF(Index).bytTones, lngSampleRate, intBits, intAudioMode, 0.5, 0.5
        Wav_MultiSineWave udtDTMF(Index).bytTones, udtSineWaves, 0.25

        Wav_Play udtDTMF(Index).bytTones

End Sub

Private Sub cmdDial_Click()
Dim i
Dim stimer
Dim a 'Input
Dim b 'Dial

For i = 1 To Len(Text1.Text)
    a = Val(Mid(Text1.Text, i, 1))
    Label5.Caption = a
    Debug.Print Mid(Text1.Text, i, 1)
    If Asc(Mid(Text1.Text, i, 1)) >= 47 And Asc(Mid(Text1.Text, i, 1)) < 58 Then
        Select Case a
            Case 1
                b = 1
            Case 2
                b = 2
            Case 3
                b = 3
            Case 4
                b = 5
            Case 5
                b = 6
            Case 6
                b = 7
            Case 7
                b = 9
            Case 8
                b = 10
            Case 9
                b = 11
            Case 0
                b = 14
            Case Else
                GoTo delayit
        End Select
        cmdTone_Click (b)
    End If

delayit:
    stimer = Timer
    Do
    DoEvents
    Loop Until Timer - stimer > 0.2
Next i

End Sub

Private Sub cmdUSd_Click()
    BoxTone 1700, 2200, 0.1
    Silence 0.08
    BoxTone 1700, 2200, 0.1
    Silence 0.165
End Sub

Private Sub cmdUSn_Click()
    BoxTone 1700, 2200, 0.1
End Sub

Private Sub cmdUSq_Click()
    BoxTone 1700, 2200, 0.08
    Silence 0.08
    BoxTone 1700, 2200, 0.04
    Silence 0.04
    BoxTone 1700, 2200, 0.04
    Silence 0.04
    BoxTone 1700, 2200, 0.04
    Silence 0.04
    BoxTone 1700, 2200, 0.1
    Silence 0.06
End Sub

Private Sub cmdWhat_Click()
MsgBox "This is a DTMF tone generator." & vbNewLine & vbNewLine & _
       "How to use:" & vbNewLine & _
       "    1. Pick up phone (Handset)" & vbNewLine & _
       "    2. Hold microphone (end You talk into) against speaker" & vbNewLine & _
       "    3. Press Dial (if You have already entered number)," & vbNewLine & _
       "       Or click buttons on form" & vbNewLine & _
       "    4. You have now dialed..." & vbNewLine & vbNewLine & _
       "Easy, isn't it?" & vbNewLine _
       , vbOKOnly, frmTones.Caption

End Sub

Private Sub Form_Load()

    cboBoxes.AddItem "Red Box"
    cboBoxes.AddItem "Blue Box"
    cboBoxes.ListIndex = 0
    cboBoxOLD = 0
End Sub
