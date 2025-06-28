Attribute VB_Name = "Module7"
'***************************************************
'　FrequencyRatioの変換（DX21内部データ->表示データ)
'***************************************************
Function Conv_FR(FR)

    Select Case FR
        Case 0
            Conv_FR = "0.5"
        Case 1
            Conv_FR = "0.71"
        Case 2
            Conv_FR = "0.78"
        Case 3
            Conv_FR = "0.87"
        Case 4
            Conv_FR = "1.00"
        Case 5
            Conv_FR = "1.41"
        Case 6
            Conv_FR = "1.57"
        Case 7
            Conv_FR = "1.73"
        Case 8
            Conv_FR = "2.00"
        Case 9
            Conv_FR = "2.82"
        Case 10
            Conv_FR = "3.00"
        Case 11
            Conv_FR = "3.14"
        Case 12
            Conv_FR = "3.46"
        Case 13
            Conv_FR = "4.00"
        Case 14
            Conv_FR = "4.24"
        Case 15
            Conv_FR = "4.71"
        Case 16
            Conv_FR = "5.00"
        Case 17
            Conv_FR = "5.19"
        Case 18
            Conv_FR = "5.65"
        Case 19
            Conv_FR = "6.00"
        Case 20
            Conv_FR = "6.28"
        Case 21
            Conv_FR = "6.92"
        Case 22
            Conv_FR = "7.00"
        Case 23
            Conv_FR = "7.07"
        Case 24
            Conv_FR = "7.85"
        Case 25
            Conv_FR = "8.00"
        Case 26
            Conv_FR = "8.48"
        Case 27
            Conv_FR = "8.65"
        Case 28
            Conv_FR = "9.00"
        Case 29
            Conv_FR = "9.42"
        Case 30
            Conv_FR = "9.89"
        Case 31
            Conv_FR = "10.00"
        Case 32
            Conv_FR = "10.38"
        Case 33
            Conv_FR = "10.99"
        Case 34
            Conv_FR = "11.00"
        Case 35
            Conv_FR = "11.30"
        Case 36
            Conv_FR = "12.00"
        Case 37
            Conv_FR = "12.11"
        Case 38
            Conv_FR = "12.56"
        Case 39
            Conv_FR = "12.72"
        Case 40
            Conv_FR = "13.00"
        Case 41
            Conv_FR = "13.84"
        Case 42
            Conv_FR = "14.00"
        Case 43
            Conv_FR = "14.10"
        Case 44
            Conv_FR = "14.13"
        Case 45
            Conv_FR = "15.00"
        Case 46
            Conv_FR = "15.55"
        Case 47
            Conv_FR = "15.57"
        Case 48
            Conv_FR = "15.70"
        Case 49
            Conv_FR = "16.96"
        Case 50
            Conv_FR = "17.27"
        Case 51
            Conv_FR = "17.30"
        Case 52
            Conv_FR = "18.37"
        Case 53
            Conv_FR = "18.84"
        Case 54
            Conv_FR = "19.03"
        Case 55
            Conv_FR = "19.78"
        Case 56
            Conv_FR = "20.41"
        Case 57
            Conv_FR = "20.76"
        Case 58
            Conv_FR = "21.20"
        Case 59
            Conv_FR = "21.98"
        Case 60
            Conv_FR = "22.49"
        Case 61
            Conv_FR = "23.55"
        Case 62
            Conv_FR = "24.22"
        Case 63
            Conv_FR = "25.95"
        Case Else
            Conv_FR = ""
    End Select

End Function
'***************************************************
'　Transporseの変換（DX21内部データ->表示データ)
'***************************************************
Function Conv_TRS(TRS)
    Select Case TRS
        Case 0
            Conv_TRS = "C-2"
        Case 1
            Conv_TRS = "C#-2"
        Case 2
            Conv_TRS = "D-2"
        Case 3
            Conv_TRS = "D#-2"
        Case 4
            Conv_TRS = "E-2"
        Case 5
            Conv_TRS = "F-2"
        Case 6
            Conv_TRS = "F#-2"
        Case 7
            Conv_TRS = "G-2"
        Case 8
            Conv_TRS = "G#-2"
        Case 9
            Conv_TRS = "A-2"
        Case 10
            Conv_TRS = "A#-2"
        Case 11
            Conv_TRS = "B-2"
        Case 12
            Conv_TRS = "C-1"
        Case 13
            Conv_TRS = "C#-1"
        Case 14
            Conv_TRS = "D-1"
        Case 15
            Conv_TRS = "D#-1"
        Case 16
            Conv_TRS = "E-1"
        Case 17
            Conv_TRS = "F-1"
        Case 18
            Conv_TRS = "F#-1"
        Case 19
            Conv_TRS = "G-1"
        Case 20
            Conv_TRS = "G#-1"
        Case 21
            Conv_TRS = "A-1"
        Case 22
            Conv_TRS = "A#-1"
        Case 23
            Conv_TRS = "B-1"
        Case 24
            Conv_TRS = "C0"
        Case 25
            Conv_TRS = "C#0"
        Case 26
            Conv_TRS = "D0"
        Case 27
            Conv_TRS = "D#0"
        Case 28
            Conv_TRS = "E0"
        Case 29
            Conv_TRS = "F0"
        Case 30
            Conv_TRS = "F#0"
        Case 31
            Conv_TRS = "G0"
        Case 32
            Conv_TRS = "G#0"
        Case 33
            Conv_TRS = "A0"
        Case 34
            Conv_TRS = "A#0"
        Case 35
            Conv_TRS = "B0"
        Case 36
            Conv_TRS = "C1"
        Case 37
            Conv_TRS = "C#1"
        Case 38
            Conv_TRS = "D1"
        Case 39
            Conv_TRS = "D#1"
        Case 40
            Conv_TRS = "E1"
        Case 41
            Conv_TRS = "F1"
        Case 42
            Conv_TRS = "F#1"
        Case 43
            Conv_TRS = "G1"
        Case 44
            Conv_TRS = "G#1"
        Case 45
            Conv_TRS = "A1"
        Case 46
            Conv_TRS = "A#1"
        Case 47
            Conv_TRS = "B1"
        Case 48
            Conv_TRS = "C2"
        Case 49
            Conv_TRS = "C#2"
        Case 50
            Conv_TRS = "D2"
        Case 51
            Conv_TRS = "D#2"
        Case 52
            Conv_TRS = "E2"
        Case 53
            Conv_TRS = "F2"
        Case 54
            Conv_TRS = "F#2"
        Case 55
            Conv_TRS = "G2"
        Case 56
            Conv_TRS = "G#2"
        Case 57
            Conv_TRS = "A2"
        Case 58
            Conv_TRS = "A#2"
        Case 59
            Conv_TRS = "B2"
        Case 60
            Conv_TRS = "C3"
        Case 61
            Conv_TRS = "C#3"
        Case 62
            Conv_TRS = "D3"
        Case 63
            Conv_TRS = "D#3"
        Case 64
            Conv_TRS = "E3"
        Case 65
            Conv_TRS = "F3"
        Case 66
            Conv_TRS = "F#3"
        Case 67
            Conv_TRS = "G3"
        Case 68
            Conv_TRS = "G#3"
        Case 69
            Conv_TRS = "A3"
        Case 70
            Conv_TRS = "A#3"
        Case 71
            Conv_TRS = "B3"
        Case 72
            Conv_TRS = "C4"
        Case 73
            Conv_TRS = "C#4"
        Case 74
            Conv_TRS = "D4"
        Case 75
            Conv_TRS = "D#4"
        Case 76
            Conv_TRS = "E4"
        Case 77
            Conv_TRS = "F4"
        Case 78
            Conv_TRS = "F#4"
        Case 79
            Conv_TRS = "G4"
        Case 80
            Conv_TRS = "G#4"
        Case 81
            Conv_TRS = "A4"
        Case 82
            Conv_TRS = "A#4"
        Case 83
            Conv_TRS = "B4"
        Case 84
            Conv_TRS = "C5"
        Case 85
            Conv_TRS = "C#5"
        Case 86
            Conv_TRS = "D5"
        Case 87
            Conv_TRS = "D#5"
        Case 88
            Conv_TRS = "E5"
        Case 89
            Conv_TRS = "F5"
        Case 90
            Conv_TRS = "F#5"
        Case 91
            Conv_TRS = "G5"
        Case 92
            Conv_TRS = "G#5"
        Case 93
            Conv_TRS = "A5"
        Case 94
            Conv_TRS = "A#5"
        Case 95
            Conv_TRS = "B5"
        Case 96
            Conv_TRS = "C6"
        Case 97
            Conv_TRS = "C#6"
        Case 98
            Conv_TRS = "D6"
        Case 99
            Conv_TRS = "D#6"
        Case 100
            Conv_TRS = "E6"
        Case 101
            Conv_TRS = "F6"
        Case 102
            Conv_TRS = "F#6"
        Case 103
            Conv_TRS = "G6"
        Case 104
            Conv_TRS = "G#6"
        Case 105
            Conv_TRS = "A6"
        Case 106
            Conv_TRS = "A#6"
        Case 107
            Conv_TRS = "B6"
        Case 108
            Conv_TRS = "C7"
        Case 109
            Conv_TRS = "C#7"
        Case 110
            Conv_TRS = "D7"
        Case 111
            Conv_TRS = "D#7"
        Case 112
            Conv_TRS = "E7"
        Case 113
            Conv_TRS = "F7"
        Case 114
            Conv_TRS = "F#7"
        Case 115
            Conv_TRS = "G7"
        Case 116
            Conv_TRS = "G#7"
        Case 117
            Conv_TRS = "A7"
        Case 118
            Conv_TRS = "A#7"
        Case 119
            Conv_TRS = "B7"
        Case 120
            Conv_TRS = "C8"
        Case 121
            Conv_TRS = "C#8"
        Case 122
            Conv_TRS = "D8"
        Case 123
            Conv_TRS = "D#8"
        Case 124
            Conv_TRS = "E8"
        Case 125
            Conv_TRS = "F8"
        Case 126
            Conv_TRS = "F#8"
        Case 127
            Conv_TRS = "G8"
        Case Else
            Conv_TRS = ""
    End Select
    
End Function
'***************************************************
'　LFO Waveの変換（DX21内部データ->表示データ)
'***************************************************
Function LFOWave(lwv)

    Select Case lwv
        Case 0
            LFOWave = "SAW UP"
        Case 1
            LFOWave = "SQUARE"
        Case 2
            LFOWave = "TRIANGLE"
        Case 3
            LFOWave = "S/HOLD"
        Case Else
            LFOWave = ""
    End Select

End Function
'***************************************************
'　Play Modeの変換（DX21内部データ->表示データ)
'***************************************************
Function PlayMode(pm)

    Select Case pm
        Case 0
            PlayMode = "Poly"
        Case 1
            PlayMode = "Mono"
        Case Else
            PlayMode = ""
    End Select

End Function

'***************************************************
'　Portamento Modeの変換（DX21内部データ->表示データ)
'***************************************************
Function PortamentoMode(prtm)

    Select Case prtm
        Case 0
            PortamentoMode = "Full Time Porta"
        Case 1
            PortamentoMode = "Fingered Porta"
        Case Else
            PortamentoMode = ""
    End Select

End Function

'***************************************************
'　ON/OFの変換（DX21内部データ->表示データ)
'***************************************************
Function On_Off(of)

    Select Case of
        Case 0
            On_Off = "OFF"
        Case 1
            On_Off = "ON"
        Case Else
            On_Off = ""
    End Select

End Function

'***************************************************
'　DX7 LFO Waveへ変換（DX21->DX7表示データ)
'***************************************************
Function cnvDX7LFOWave(lwv)

    Select Case lwv
        Case "SAW UP"
            cnvDX7LFOWave = 2
        Case "SQUARE"
            cnvDX7LFOWave = 3
        Case "TRIANGLE"
            cnvDX7LFOWave = 0
        Case "S/HOLD"
            cnvDX7LFOWave = 5
        Case Else
            cnvDX7LFOWave = ""
    End Select
   
End Function

'***************************************************
'　DX7 LFO Waveの変換（DX7内部データ->表示データ)
'***************************************************
Function DX7LFOWave(lwv)

    Select Case lwv
        Case 0
            DX7LFOWave = "TR"
        Case 1
            DX7LFOWave = "SD"
        Case 2
            DX7LFOWave = "SU"
        Case 3
            DX7LFOWave = "SQ"
        Case 4
            DX7LFOWave = "SI"
        Case 5
            DX7LFOWave = "SH"
        Case Else
            DX7LFOWave = ""
    End Select

End Function

'***************************************************
'　Ratio/Fixの変換（DX7内部データ->表示データ)
'***************************************************
Function Ratio_Fixed(rf)

    Select Case rf
        Case 0
            Ratio_Fixed = "Ratio"
        Case 1
            Ratio_Fixed = "Fixed"
        Case Else
            Ratio_Fixed = ""
    End Select

End Function

'***************************************************
'　DX11 OP Waveの変換（DX11内部データ->表示データ)
'***************************************************
Function DX11_OP_Wave(opwv)

    Select Case opwv
        Case 0
            DX11_OP_Wave = "TR"
        Case 1
            DX11_OP_Wave = "SD"
        Case 2
            DX11_OP_Wave = "SU"
        Case 3
            DX11_OP_Wave = "SQ"
        Case 4
            DX11_OP_Wave = "SI"
        Case 5
            DX11_OP_Wave = "SH"
        Case Else
            DX11_OP_Wave = ""
    End Select

End Function

'***************************************************
'　FrequencyRatioの変換（DX21内部データ->OPMデータ)
'***************************************************
Function ConvFR_DX21toOPM(FR)

    Select Case FR
        Case 0
            ConvFR_DX21toOPM = "0"
        Case 1
            ConvFR_DX21toOPM = "0"
        Case 2
            ConvFR_DX21toOPM = "0"
        Case 3
            ConvFR_DX21toOPM = "0"
        Case 4
            ConvFR_DX21toOPM = "1"
        Case 5
            ConvFR_DX21toOPM = "1"
        Case 6
            ConvFR_DX21toOPM = "1"
        Case 7
            ConvFR_DX21toOPM = "1"
        Case 8
            ConvFR_DX21toOPM = "2"
        Case 9
            ConvFR_DX21toOPM = "2"
        Case 10
            ConvFR_DX21toOPM = "3"
        Case 11
            ConvFR_DX21toOPM = "3"
        Case 12
            ConvFR_DX21toOPM = "3"
        Case 13
            ConvFR_DX21toOPM = "4"
        Case 14
            ConvFR_DX21toOPM = "4"
        Case 15
            ConvFR_DX21toOPM = "4"
        Case 16
            ConvFR_DX21toOPM = "5"
        Case 17
            ConvFR_DX21toOPM = "5"
        Case 18
            ConvFR_DX21toOPM = "5"
        Case 19
            ConvFR_DX21toOPM = "6"
        Case 20
            ConvFR_DX21toOPM = "6"
        Case 21
            ConvFR_DX21toOPM = "6"
        Case 22
            ConvFR_DX21toOPM = "7"
        Case 23
            ConvFR_DX21toOPM = "7"
        Case 24
            ConvFR_DX21toOPM = "7"
        Case 25
            ConvFR_DX21toOPM = "8"
        Case 26
            ConvFR_DX21toOPM = "8"
        Case 27
            ConvFR_DX21toOPM = "8"
        Case 28
            ConvFR_DX21toOPM = "9"
        Case 29
            ConvFR_DX21toOPM = "9"
        Case 30
            ConvFR_DX21toOPM = "9"
        Case 31
            ConvFR_DX21toOPM = "10"
        Case 32
            ConvFR_DX21toOPM = "10"
        Case 33
            ConvFR_DX21toOPM = "10"
        Case 34
            ConvFR_DX21toOPM = "11"
        Case 35
            ConvFR_DX21toOPM = "11"
        Case 36
            ConvFR_DX21toOPM = "12"
        Case 37
            ConvFR_DX21toOPM = "12"
        Case 38
            ConvFR_DX21toOPM = "12"
        Case 39
            ConvFR_DX21toOPM = "12"
        Case 40
            ConvFR_DX21toOPM = "13"
        Case 41
            ConvFR_DX21toOPM = "13"
        Case 42
            ConvFR_DX21toOPM = "14"
        Case 43
            ConvFR_DX21toOPM = "14"
        Case 44
            ConvFR_DX21toOPM = "14"
        Case 45
            ConvFR_DX21toOPM = "15"
        Case 46
            ConvFR_DX21toOPM = "15"
        Case 47
            ConvFR_DX21toOPM = "15"
        Case 48
            ConvFR_DX21toOPMR = "15"
        Case Else
            ConvFR_DX21toOPM = "1"
    End Select

End Function

'***************************************************
'　D1Lの変換（DX21内部データ->OPMデータ)
'***************************************************
Function ConvD1L_DX21toOPM(D1L)

    ConvD1L_DX21toOPM = 15 - D1L

End Function

'***************************************************
'　OLの変換（DX21内部データ->OPMデータ)
'***************************************************
Function ConvOL_DX21toOPM(OL)

    ConvOL_DX21toOPM = 127 - Round((127 / 99) * OL, 0)

End Function

'***************************************************
'　DTの変換（DX21内部データ->OPMデータ)
'***************************************************
Function ConvDT_DX21toOPM(DT)

    ConvDT_DX21toOPM = DT + 3

End Function

'***************************************************
'　TLの変換（OPMデータ->DX21内部データ)
'***************************************************
Function ConvTL_OPMtoDX21(TL)

     ConvTL_OPMtoDX21 = Round((127 - TL) * 99 / 127, 0)

End Function

'***************************************************
'　Speedの変換（DX21内部データ->OPMデータ)
'***************************************************
Function ConvSP_DX21toOPM(SP)

    ConvSP_DX21toOPM = Round((255 / 99) * SP, 0)

End Function

'********************************************************
'　LFOのModuration Sensの変換（DX21内部データ->OPMデータ)
'********************************************************
Function ConvLFOs_DX21toOPM(sens)

    ConvLFOs_DX21toOPM = Round((127 / 99) * sens, 0)

End Function

'***************************************************
'　Algorismの変換（DX21内部データ->OPMデータ)
'***************************************************
Function ConvALG_DX21toOPM(ARG)

    ConvALG_DX21toOPM = ARG - 1
    
End Function

'***************************************************
'　Algorismの変換（DX21->DX7データ)
'***************************************************
Function ConvALG_DX21toDX7(ARG)
    
    Select Case ARG
    
        Case 1
            ConvALG_DX21toDX7 = 1
        Case 2
            ConvALG_DX21toDX7 = 14
        Case 3
            ConvALG_DX21toDX7 = 8
        Case 4
            ConvALG_DX21toDX7 = 7
        Case 5
            ConvALG_DX21toDX7 = 5
        Case 6
            ConvALG_DX21toDX7 = 22
        Case 7
            ConvALG_DX21toDX7 = 31
        Case 8
            ConvALG_DX21toDX7 = 32
    
    End Select
    
End Function

