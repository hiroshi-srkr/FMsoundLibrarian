Attribute VB_Name = "Module1"
'********************************************
'　Voiceデータをデータベースシートに表示
'********************************************
Sub Creat_VoiceList()

    Dim wsSource As String
    Dim wsSource1 As String
    Dim wsSource2 As String
    Dim wsSource3 As String
    Dim wsSource4 As String

    wsSource1 = "X68"
    wsSource2 = "PC88"
    wsSource3 = "PMD"
    wsSource4 = "FMLib"
    
    
    Call Write_VoiceList(wsSource1, 2, 1)
    Call Write_VoiceList(wsSource2, 70, 1)
    Call Write_VoiceList(wsSource3, 197, 1)
    Call Write_VoiceList(wsSource4, 453, 1)
    
End Sub

'********************************************
'　Voiceデータをデータベースシートに書き出し
'********************************************
Sub Write_VoiceList(wsSource, tr, tc)

    Dim wsTarget As String
    Dim LibName As String
    Dim VoiceName As String
    Dim ARG, FB As Integer
    Dim OP1_AR, OP1_D1R, OP1_D1L, OP1_D2R, OP1_RR, OP1_OL, OP1_KS, OP1_FR, OP1_DT, OP1_AMS, OP1_SN, OP1_SL, OP1_TL, OP1_ML, OP1_ODT As Integer
    Dim OP2_AR, OP2_D1R, OP2_D1L, OP2_D2R, OP2_RR, OP2_OL, OP2_KS, OP2_FR, OP2_DT, OP2_AMS, OP2_SN, OP2_SL, OP2_TL, OP2_ML, OP2_ODT As Integer
    Dim OP3_AR, OP3_D1R, OP3_D1L, OP3_D3R, OP3_RR, OP3_OL, OP3_KS, OP3_FR, OP3_DT, OP3_AMS, OP3_SN, OP3_SL, OP3_TL, OP3_ML, OP3_ODT As Integer
    Dim OP4_AR, OP4_D1R, OP4_D1L, OP4_D2R, OP4_RR, OP4_OL, OP4_KS, OP4_FR, OP4_DT, OP4_AMS, OP4_SN, OP4_SL, OP4_TL, OP4_ML, OP4_ODT As Integer
    Dim sr, sc As Integer
        
    wsTarget = "DX21_VoiceDATABASE"
        
    sr = 2
    sc = 12

    Do
        Sheets(wsSource).Activate
        LibName = wsSource
        VoiceName = Cells(sr, sc).Value
        ARG = Cells(sr, sc + 1).Value
        FB = Cells(sr, sc + 2).Value
        OP1_FR = Cells(sr + 5, sc + 1).Value
        OP1_DT = Cells(sr + 5, sc + 2).Value
        OP1_AR = Cells(sr + 5, sc + 3).Value
        OP1_D1R = Cells(sr + 5, sc + 4).Value
        OP1_D1L = Cells(sr + 5, sc + 5).Value
        OP1_D2R = Cells(sr + 5, sc + 6).Value
        OP1_RR = Cells(sr + 5, sc + 7).Value
        OP1_OL = Cells(sr + 5, sc + 8).Value
        OP1_KS = Cells(sr + 5, sc + 9).Value
        OP1_AMS = 0
        OP1_SN = 0
        OP1_SL = Cells(sr + 5, sc - 7).Value
        OP1_TL = Cells(sr + 5, sc - 6).Value
        OP1_ML = Cells(sr + 5, sc - 4).Value
        OP1_ODT = Cells(sr + 5, sc - 3).Value
    
        OP2_FR = Cells(sr + 4, sc + 1).Value
        OP2_DT = Cells(sr + 4, sc + 2).Value
        OP2_AR = Cells(sr + 4, sc + 3).Value
        OP2_D1R = Cells(sr + 4, sc + 4).Value
        OP2_D1L = Cells(sr + 4, sc + 5).Value
        OP2_D2R = Cells(sr + 4, sc + 6).Value
        OP2_RR = Cells(sr + 4, sc + 7).Value
        OP2_OL = Cells(sr + 4, sc + 8).Value
        OP2_KS = Cells(sr + 4, sc + 9).Value
        OP2_AMS = 0
        OP2_SN = 0
        OP2_SL = Cells(sr + 4, sc - 7).Value
        OP2_TL = Cells(sr + 4, sc - 6).Value
        OP2_ML = Cells(sr + 4, sc - 4).Value
        OP2_ODT = Cells(sr + 4, sc - 3).Value
    
        OP3_FR = Cells(sr + 3, sc + 1).Value
        OP3_DT = Cells(sr + 3, sc + 2).Value
        OP3_AR = Cells(sr + 3, sc + 3).Value
        OP3_D1R = Cells(sr + 3, sc + 4).Value
        OP3_D1L = Cells(sr + 3, sc + 5).Value
        OP3_D2R = Cells(sr + 3, sc + 6).Value
        OP3_RR = Cells(sr + 3, sc + 7).Value
        OP3_OL = Cells(sr + 3, sc + 8).Value
        OP3_KS = Cells(sr + 3, sc + 9).Value
        OP3_AMS = 0
        OP3_SN = 0
        OP3_SL = Cells(sr + 3, sc - 7).Value
        OP3_TL = Cells(sr + 3, sc - 6).Value
        OP3_ML = Cells(sr + 3, sc - 4).Value
        OP3_ODT = Cells(sr + 3, sc - 3).Value
    
        OP4_FR = Cells(sr + 2, sc + 1).Value
        OP4_DT = Cells(sr + 2, sc + 2).Value
        OP4_AR = Cells(sr + 2, sc + 3).Value
        OP4_D1R = Cells(sr + 2, sc + 4).Value
        OP4_D1L = Cells(sr + 2, sc + 5).Value
        OP4_D2R = Cells(sr + 2, sc + 6).Value
        OP4_RR = Cells(sr + 2, sc + 7).Value
        OP4_OL = Cells(sr + 2, sc + 8).Value
        OP4_KS = Cells(sr + 2, sc + 9).Value
        OP4_AMS = 0
        OP4_SN = 0
        OP4_SL = Cells(sr + 2, sc - 7).Value
        OP4_TL = Cells(sr + 2, sc - 6).Value
        OP4_ML = Cells(sr + 2, sc - 4).Value
        OP4_ODT = Cells(sr + 2, sc - 3).Value
    
        Sheets(wsTarget).Activate
        Cells(tr, tc).Value = LibName
        Cells(tr, tc + 1).Value = VoiceName
        Cells(tr, tc + 2).Value = ARG
        Cells(tr, tc + 3).Value = FB
    
        Cells(tr, tc + 4).Value = OP1_AR
        Cells(tr, tc + 5).Value = OP1_D1R
        Cells(tr, tc + 6).Value = OP1_D1L
        Cells(tr, tc + 7).Value = OP1_D2R
        Cells(tr, tc + 8).Value = OP1_RR
        Cells(tr, tc + 9).Value = OP1_OL
        Cells(tr, tc + 10).Value = OP1_KS
        Cells(tr, tc + 11).Value = OP1_FR
        Cells(tr, tc + 12).Value = OP1_DT
        Cells(tr, tc + 13).Value = OP1_AMS
        Cells(tr, tc + 14).Value = OP1_SN
    
        Cells(tr, tc + 15).Value = OP2_AR
        Cells(tr, tc + 16).Value = OP2_D1R
        Cells(tr, tc + 17).Value = OP2_D1L
        Cells(tr, tc + 18).Value = OP2_D2R
        Cells(tr, tc + 19).Value = OP2_RR
        Cells(tr, tc + 20).Value = OP2_OL
        Cells(tr, tc + 21).Value = OP2_KS
        Cells(tr, tc + 22).Value = OP2_FR
        Cells(tr, tc + 23).Value = OP2_DT
        Cells(tr, tc + 24).Value = OP2_AMS
        Cells(tr, tc + 25).Value = OP2_SN
    
        Cells(tr, tc + 26).Value = OP3_AR
        Cells(tr, tc + 27).Value = OP3_D1R
        Cells(tr, tc + 28).Value = OP3_D1L
        Cells(tr, tc + 29).Value = OP3_D2R
        Cells(tr, tc + 30).Value = OP3_RR
        Cells(tr, tc + 31).Value = OP3_OL
        Cells(tr, tc + 32).Value = OP3_KS
        Cells(tr, tc + 33).Value = OP3_FR
        Cells(tr, tc + 34).Value = OP3_DT
        Cells(tr, tc + 35).Value = OP3_AMS
        Cells(tr, tc + 36).Value = OP3_SN
    
        Cells(tr, tc + 37).Value = OP4_AR
        Cells(tr, tc + 38).Value = OP4_D1R
        Cells(tr, tc + 39).Value = OP4_D1L
        Cells(tr, tc + 40).Value = OP4_D2R
        Cells(tr, tc + 41).Value = OP4_RR
        Cells(tr, tc + 42).Value = OP4_OL
        Cells(tr, tc + 43).Value = OP4_KS
        Cells(tr, tc + 44).Value = OP4_FR
        Cells(tr, tc + 45).Value = OP4_DT
        Cells(tr, tc + 46).Value = OP4_AMS
        Cells(tr, tc + 47).Value = OP4_SN
        
        Cells(tr, tc + 48).Value = OP1_SL
        Cells(tr, tc + 49).Value = OP1_TL
        Cells(tr, tc + 50).Value = OP1_ML
        Cells(tr, tc + 51).Value = OP1_ODT
        Cells(tr, tc + 52).Value = OP2_SL
        Cells(tr, tc + 53).Value = OP2_TL
        Cells(tr, tc + 54).Value = OP2_ML
        Cells(tr, tc + 55).Value = OP2_ODT
        Cells(tr, tc + 56).Value = OP3_SL
        Cells(tr, tc + 57).Value = OP3_TL
        Cells(tr, tc + 58).Value = OP3_ML
        Cells(tr, tc + 59).Value = OP3_ODT
        Cells(tr, tc + 60).Value = OP4_SL
        Cells(tr, tc + 61).Value = OP4_TL
        Cells(tr, tc + 62).Value = OP4_ML
        Cells(tr, tc + 63).Value = OP4_ODT
        
        sr = sr + 8
        tr = tr + 1
        
    Loop Until VoiceName = ""
End Sub



