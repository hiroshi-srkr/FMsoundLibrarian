Attribute VB_Name = "Module4"

'********************************************
'
' 関数名：BinaryMain
'
' 機能：4OP用Sysexファイルの読み込み
'
' 呼び出し関数：ReadSysexFile
'
'********************************************
Sub BinaryMain()
    Dim strFilePath As String
    Dim strFileName As String
    Dim strFile As String
    '//ファイル名を生成
    
    Sheets("Menu").Select
    strFilePath = Cells(10, 5).Value
    strFileName = Cells(11, 5).Value

    If strFilePath = "" Then
        strFilePath = ThisWorkbook.Path
    End If
    
    If strFileName = "" Then
        MsgBox "ファイル名が指定されていません。"
        
    Else
        strFile = strFilePath & "\" & strFileName
        
        If Dir(strFile) <> "" Then
  
        '//バイナリファイルを読み込み（Sysexファイルを読み込み）
            Call ReadSysexFile(strFile)
            Sheets("SysexData").Select
            MsgBox "Sysexデータの読み込みが完了しました。"
            
        Else
            MsgBox strFile & vbCrLf & _
                    "が存在しません"
        End If
    End If

End Sub

'********************************************
'
' 関数名：ReadSysexFile
'
' 機能：4OP用Sysexファイルからデータを読み込み
' 　　　SysexDataシートに出力
'
' 引数：strfil：読み込むファイル名
'
'********************************************
Sub ReadSysexFile(ByVal strfil As String)
    Dim buff() As Byte
    Dim fp As Long
    Dim filesize As Long, NowLoc As Long
    Dim idx As Long, gyo As Long
    Dim strBinary As String
    Dim wsTALGet As String
    
    Dim VoiceName As String
    Dim ALG As Integer, FB As Integer
    Dim OP1_AR, OP1_D1R, OP1_D1L, OP1_D2R, OP1_RR, OP1_OL, OP1_KS, OP1_FR, OP1_DT, OP1_AMS, OP1_SN, OP1_SL, OP1_TL, OP1_ML, OP1_ODT, OP1_KL, OP1_EB
    Dim OP2_AR, OP2_D1R, OP2_D1L, OP2_D2R, OP2_RR, OP2_OL, OP2_KS, OP2_FR, OP2_DT, OP2_AMS, OP2_SN, OP2_SL, OP2_TL, OP2_ML, OP2_ODT, OP2_KL, OP2_EB
    Dim OP3_AR, OP3_D1R, OP3_D1L, OP3_D3R, OP3_RR, OP3_OL, OP3_KS, OP3_FR, OP3_DT, OP3_AMS, OP3_SN, OP3_SL, OP3_TL, OP3_ML, OP3_ODT, OP3_KL, OP3_EB
    Dim OP4_AR, OP4_D1R, OP4_D1L, OP4_D2R, OP4_RR, OP4_OL, OP4_KS, OP4_FR, OP4_DT, OP4_AMS, OP4_SN, OP4_SL, OP4_TL, OP4_ML, OP4_ODT, OP4_KL, OP4_EB
    Dim LFO_Speed, LFO_Delay, PMD, AMD, LFO_Sync, LFO_Wave, PMS, AMS, TRS, POLY_MONO, PBR, P_Mode, P_Time, FV, Sus_S, P_Switch, Chorus, MWP_Range, MWA_Range
    Dim BPM_Range, BAM_Range, BPB_Range, BEB_Range, PR1, PR2, PR3, PL1, PL2, PL3, OP_ONOFF
    
    Dim OP1_EGSFT, OP1_FIX, OP1_FIXRM, OP1_FIXRG, OP1_OSW, OP1_FINE
    Dim OP2_EGSFT, OP2_FIX, OP2_FIXRM, OP2_FIXRG, OP2_OSW, OP2_FINE
    Dim OP3_EGSFT, OP3_FIX, OP3_FIXRM, OP3_FIXRG, OP3_OSW, OP3_FINE
    Dim OP4_EGSFT, OP4_FIX, OP4_FIXRM, OP4_FIXRG, OP4_OSW, OP4_FINE
    Dim REV, FC_Pitch, FC_AMP, AT_Pitch, AT_AMP, AT_PBias, AT_EGBias
    
    Dim E_PRST, E_TIME, E_BAL, E_SEL, BAL, OUT_LV, ST_MIX, E_PRM1, E_PRM2, E_PRM3
    
    Dim Data_set, c, cc
    Dim sr As Long, tr As Long, tc As Long
    'Dim chksum2 As Long
    
    
    'データ出力用シート
    wsTALGet = "SysexData"
    
    Sheets(wsTALGet).Activate
    '    For Each cc In Range("A2:CG33", "CI2:DB33")
    '        If cc.Locked = False Then
    '            cc.Value = ""
    '        End If
    '    Next cc
     
    ActiveSheet.Unprotect
    Range("A2:CG33", "CI2:DB33").ClearContents
    Range("DD2:DJ33", "DL2:DU33").ClearContents
    Range("DW2:DZ33").ClearContents
    ActiveSheet.Protect
    
    '//FreeFile関数で使用可能なファイル番号を割り当て
    fp = FreeFile
    '//ファイルを開く
    Open strfil For Binary As #fp
    '//ファイルサイズ分の読み込み領域を確保して読み込む場合の実装例
    ReDim buff(FileLen(strfil))
    Get #fp, 1, buff
    '//実装例ここまで

    '//ファイルを閉じる
    Close (fp)
    
    '読み込みカウンター初期化（ヘッダーの読み飛ばし7バイト目から）
    sr = 6
    
    'シート書き込み用カウンター初期化
    tr = 2
    tc = 2
   
    'Sysexデータの4バイト目を確認し1Voiceデータか32Voiceバルクデータかを判別
    If buff(3) = 3 Then

        Data_set = 1
        
    ElseIf buff(3) = 4 Then

        Data_set = 32
        
    End If

    'chksum2 = 0

    For c = 1 To Data_set
        
        If Data_set = 1 Then
        
        '1Voiceデータの場合
            OP4_AR = Hex("&H" & buff(sr))
            OP4_D1R = Hex("&H" & buff(sr + 1))
            OP4_D2R = Hex("&H" & buff(sr + 2))
            OP4_RR = Hex("&H" & buff(sr + 3))
            OP4_D1L = Hex("&H" & buff(sr + 4))
            OP4_KL = Hex("&H" & buff(sr + 5))
            OP4_KS = Hex("&H" & buff(sr + 6))
            OP4_EB = Hex("&H" & buff(sr + 7))
            OP4_AMS = Hex("&H" & buff(sr + 8))
            OP4_SN = Hex("&H" & buff(sr + 9))
            OP4_OL = Hex("&H" & buff(sr + 10))
            OP4_FR = Hex("&H" & buff(sr + 11))
            OP4_DT = Hex("&H" & buff(sr + 12))
                    
            OP2_AR = Hex("&H" & buff(sr + 13))
            OP2_D1R = Hex("&H" & buff(sr + 14))
            OP2_D2R = Hex("&H" & buff(sr + 15))
            OP2_RR = Hex("&H" & buff(sr + 16))
            OP2_D1L = Hex("&H" & buff(sr + 17))
            OP2_KL = Hex("&H" & buff(sr + 18))
            OP2_KS = Hex("&H" & buff(sr + 19))
            OP2_EB = Hex("&H" & buff(sr + 20))
            OP2_AMS = Hex("&H" & buff(sr + 21))
            OP2_SN = Hex("&H" & buff(sr + 22))
            OP2_OL = Hex("&H" & buff(sr + 23))
            OP2_FR = Hex("&H" & buff(sr + 24))
            OP2_DT = Hex("&H" & buff(sr + 25))
            
            OP3_AR = Hex("&H" & buff(sr + 26))
            OP3_D1R = Hex("&H" & buff(sr + 27))
            OP3_D2R = Hex("&H" & buff(sr + 28))
            OP3_RR = Hex("&H" & buff(sr + 29))
            OP3_D1L = Hex("&H" & buff(sr + 30))
            OP3_KL = Hex("&H" & buff(sr + 31))
            OP3_KS = Hex("&H" & buff(sr + 32))
            OP3_EB = Hex("&H" & buff(sr + 33))
            OP3_AMS = Hex("&H" & buff(sr + 34))
            OP3_SN = Hex("&H" & buff(sr + 35))
            OP3_OL = Hex("&H" & buff(sr + 36))
            OP3_FR = Hex("&H" & buff(sr + 37))
            OP3_DT = Hex("&H" & buff(sr + 38))
            
            OP1_AR = Hex("&H" & buff(sr + 39))
            OP1_D1R = Hex("&H" & buff(sr + 40))
            OP1_D2R = Hex("&H" & buff(sr + 41))
            OP1_RR = Hex("&H" & buff(sr + 42))
            OP1_D1L = Hex("&H" & buff(sr + 43))
            OP1_KL = Hex("&H" & buff(sr + 44))
            OP1_KS = Hex("&H" & buff(sr + 45))
            OP1_EB = Hex("&H" & buff(sr + 46))
            OP1_AMS = Hex("&H" & buff(sr + 47))
            OP1_SN = Hex("&H" & buff(sr + 48))
            OP1_OL = Hex("&H" & buff(sr + 49))
            OP1_FR = Hex("&H" & buff(sr + 50))
            OP1_DT = Hex("&H" & buff(sr + 51))
            
            ALG = Hex("&H" & buff(sr + 52))
            FB = Hex("&H" & buff(sr + 53))
            LFO_Speed = Hex("&H" & buff(sr + 54))
            LFO_Delay = Hex("&H" & buff(sr + 55))
            PMD = Hex("&H" & buff(sr + 56))
            AMD = Hex("&H" & buff(sr + 57))
            LFO_Sync = Hex("&H" & buff(sr + 58))
            LFO_Wave = Hex("&H" & buff(sr + 59))
            PMS = Hex("&H" & buff(sr + 60))
            AMS = Hex("&H" & buff(sr + 61))
            TRS = Hex("&H" & buff(sr + 62))
            POLY_MONO = Hex("&H" & buff(sr + 63))
            PBR = Hex("&H" & buff(sr + 64))
            P_Mode = Hex("&H" & buff(sr + 65))
            P_Time = Hex("&H" & buff(sr + 66))
            FV = Hex("&H" & buff(sr + 67))
            Sus_S = Hex("&H" & buff(sr + 68))
            P_Switch = Hex("&H" & buff(sr + 69))
            Chorus = Hex("&H" & buff(sr + 70))
            MWP_Range = Hex("&H" & buff(sr + 71))
            MWA_Range = Hex("&H" & buff(sr + 72))
            BPM_Range = Hex("&H" & buff(sr + 73))
            BAM_Range = Hex("&H" & buff(sr + 74))
            BPB_Range = Hex("&H" & buff(sr + 75))
            BEB_Range = Hex("&H" & buff(sr + 76))
            
            VoiceName = Chr(Hex("&H" & buff(sr + 77))) & Chr(Hex("&H" & buff(sr + 78))) & Chr(Hex("&H" & buff(sr + 79))) & Chr(Hex("&H" & buff(sr + 80))) & Chr(Hex("&H" & buff(sr + 81))) & Chr(Hex("&H" & buff(sr + 82))) & Chr(Hex("&H" & buff(sr + 83))) & Chr(Hex("&H" & buff(sr + 84))) & Chr(Hex("&H" & buff(sr + 85))) & Chr(Hex("&H" & buff(sr + 86)))
            
            PR1 = Hex("&H" & buff(sr + 87))
            PR2 = Hex("&H" & buff(sr + 88))
            PR3 = Hex("&H" & buff(sr + 89))
            PL1 = Hex("&H" & buff(sr + 90))
            PL2 = Hex("&H" & buff(sr + 91))
            PL3 = Hex("&H" & buff(sr + 92))
            
            OP_ONOFF = Hex("&H" & buff(sr + 93))
            
            If Hex("&H" & buff(sr + 94)) <> 247 Then
            
                'YAMAHA V2(DX11)のオシレーターデータの取得

                OP4_FIX = Hex("&H" & buff(sr + 94))
                OP4_FIXRG = Hex("&H" & buff(sr + 95))
                OP4_FINE = Hex("&H" & buff(sr + 96))
                OP4_OSW = Hex("&H" & buff(sr + 97))
                OP4_EGSFT = Hex("&H" & buff(sr + 98))

                OP2_FIX = Hex("&H" & buff(sr + 99))
                OP2_FIXRG = Hex("&H" & buff(sr + 100))
                OP2_FINE = Hex("&H" & buff(sr + 101))
                OP2_OSW = Hex("&H" & buff(sr + 102))
                OP2_EGSFT = Hex("&H" & buff(sr + 103))

                OP3_FIX = Hex("&H" & buff(sr + 104))
                OP3_FIXRG = Hex("&H" & buff(sr + 105))
                OP3_FINE = Hex("&H" & buff(sr + 106))
                OP3_OSW = Hex("&H" & buff(sr + 107))
                OP3_EGSFT = Hex("&H" & buff(sr + 108))
            
                OP4_FIX = Hex("&H" & buff(sr + 109))
                OP4_FIXRG = Hex("&H" & buff(sr + 110))
                OP4_FINE = Hex("&H" & buff(sr + 111))
                OP4_OSW = Hex("&H" & buff(sr + 112))
                OP4_EGSFT = Hex("&H" & buff(sr + 113))
            
                'YAMAHA V2、V50のフットコントロール、アフタータッチ
                REV = Hex("&H" & buff(sr + 114))
                FC_Pitch = Hex("&H" & buff(sr + 115))
                FC_AMP = Hex("&H" & buff(sr + 116))
                AT_Pitch = Hex("&H" & buff(sr + 117))
                AT_AMP = Hex("&H" & buff(sr + 118))
                AT_PBias = Hex("&H" & buff(sr + 119))
                AT_EGBias = Hex("&H" & buff(sr + 120))
            
                OP4_FIXRM = Hex("&H" & buff(sr + 121))
                OP2_FIXRM = Hex("&H" & buff(sr + 122))
                OP3_FIXRM = Hex("&H" & buff(sr + 123))
                OP1_FIXRM = Hex("&H" & buff(sr + 124))
                LS_SIGN = Hex("&H" & buff(sr + 125))
            
                'YAMAHA V50のエフェクト
                E_PRST = Hex("&H" & buff(sr + 127))
                E_TIME = Hex("&H" & buff(sr + 128))
                E_BAL = Hex("&H" & buff(sr + 129))
                E_SEL = Hex("&H" & buff(sr + 130))
                BAL = Hex("&H" & buff(sr + 131))
                OUT_LV = Hex("&H" & buff(sr + 132))
                ST_MIX = Hex("&H" & buff(sr + 133))
                E_PRM1 = Hex("&H" & buff(sr + 134))
                E_PRM2 = Hex("&H" & buff(sr + 135))
                E_PRM3 = Hex("&H" & buff(sr + 136))
            
            Else
            
                'YAMAHA V2(DX11)のオシレーターデータ以降は0ゼロとする

                OP4_FIX = 0
                OP4_FIXRG = 0
                OP4_FINE = 0
                OP4_OSW = 0
                OP4_EGSFT = 0

                OP2_FIX = 0
                OP2_FIXRG = 0
                OP2_FINE = 0
                OP2_OSW = 0
                OP2_EGSFT = 0

                OP3_FIX = 0
                OP3_FIXRG = 0
                OP3_FINE = 0
                OP3_OSW = 0
                OP3_EGSFT = 0
            
                OP4_FIXR = 0
                OP4_FIXRG = 0
                OP4_FINE = 0
                OP4_OSW = 0
                OP4_EGSFT = 0
            
                'YAMAHA V2、V50のフットコントロール、アフタータッチ
                REV = 0
                FC_Pitch = 0
                FC_AMP = 0
                AT_Pitch = 0
                AT_AMP = 0
                AT_PBias = 0
                AT_EGBias = 0
            
                OP4_FIXRM = 0
                OP2_FIXRM = 0
                OP3_FIXRM = 0
                OP1_FIXRM = 0
                LS_SIGN = 0
            
                'YAMAHA V50のエフェクト
                E_PRST = 0
                E_TIME = 0
                E_BAL = 0
                E_SEL = 0
                BAL = 0
                OUT_LV = 0
                ST_MIX = 0
                E_PRM1 = 0
                E_PRM2 = 0
                E_PRM3 = 0
                
            End If
                    
        ElseIf Data_set = 32 Then
        
        '32Voiceバルクデータの場合
            OP4_AR = Hex("&H" & buff(sr))
            OP4_D1R = Hex("&H" & buff(sr + 1))
            OP4_D2R = Hex("&H" & buff(sr + 2))
            OP4_RR = Hex("&H" & buff(sr + 3))
            OP4_D1L = Hex("&H" & buff(sr + 4))
            OP4_KL = Hex("&H" & buff(sr + 5))
            
            ' Amplitude Modulation Enable/EG Bias Sensitivity/Key Volocity
            OP4_AMS = BinToDec(Left(DecToBin(Hex("&H" & buff(sr + 6)), 7), 1))
            OP4_EB = BinToDec(Mid(DecToBin(Hex("&H" & buff(sr + 6)), 7), 2, 3))
            OP4_SN = BinToDec(Right(DecToBin(Hex("&H" & buff(sr + 6)), 7), 3))
            
            OP4_OL = Hex("&H" & buff(sr + 7))
            OP4_FR = Hex("&H" & buff(sr + 8))
            
            'Keyboad Scaling Rate/Detune1
            OP4_KS = BinToDec(Left(DecToBin(Hex("&H" & buff(sr + 9)), 5), 2))
            OP4_DT = BinToDec(Right(DecToBin(Hex("&H" & buff(sr + 9)), 5), 3))
            
            OP2_AR = Hex("&H" & buff(sr + 10))
            OP2_D1R = Hex("&H" & buff(sr + 11))
            OP2_D2R = Hex("&H" & buff(sr + 12))
            OP2_RR = Hex("&H" & buff(sr + 13))
            OP2_D1L = Hex("&H" & buff(sr + 14))
            OP2_KL = Hex("&H" & buff(sr + 15))
            
            ' Amplitude Modulation Enable/EG Bias Sensitivity/Key Volocity
            OP2_AMS = BinToDec(Left(DecToBin(Hex("&H" & buff(sr + 16)), 7), 1))
            OP2_EB = BinToDec(Mid(DecToBin(Hex("&H" & buff(sr + 16)), 7), 2, 3))
            OP2_SN = BinToDec(Right(DecToBin(Hex("&H" & buff(sr + 16)), 7), 3))
            
            OP2_OL = Hex("&H" & buff(sr + 17))
            OP2_FR = Hex("&H" & buff(sr + 18))
            
            'Keyboad Scaling Rate/Detune1
            OP2_KS = BinToDec(Left(DecToBin(Hex("&H" & buff(sr + 19)), 5), 2))
            OP2_DT = BinToDec(Right(DecToBin(Hex("&H" & buff(sr + 19)), 5), 3))
            
            OP3_AR = Hex("&H" & buff(sr + 20))
            OP3_D1R = Hex("&H" & buff(sr + 21))
            OP3_D2R = Hex("&H" & buff(sr + 22))
            OP3_RR = Hex("&H" & buff(sr + 23))
            OP3_D1L = Hex("&H" & buff(sr + 24))
            OP3_KL = Hex("&H" & buff(sr + 25))
            
            ' Amplitude Modulation Enable/EG Bias Sensitivity/Key Volocity
            OP3_AMS = BinToDec(Left(DecToBin(Hex("&H" & buff(sr + 26)), 7), 1))
            OP3_EB = BinToDec(Mid(DecToBin(Hex("&H" & buff(sr + 26)), 7), 2, 3))
            OP3_SN = BinToDec(Right(DecToBin(Hex("&H" & buff(sr + 26)), 7), 3))
            
            OP3_OL = Hex("&H" & buff(sr + 27))
            OP3_FR = Hex("&H" & buff(sr + 28))
            
            'Keyboad Scaling Rate/Detune1
            OP3_KS = BinToDec(Left(DecToBin(Hex("&H" & buff(sr + 29)), 5), 2))
            OP3_DT = BinToDec(Right(DecToBin(Hex("&H" & buff(sr + 29)), 5), 3))
            
            OP1_AR = Hex("&H" & buff(sr + 30))
            OP1_D1R = Hex("&H" & buff(sr + 31))
            OP1_D2R = Hex("&H" & buff(sr + 32))
            OP1_RR = Hex("&H" & buff(sr + 33))
            OP1_D1L = Hex("&H" & buff(sr + 34))
            OP1_KL = Hex("&H" & buff(sr + 35))
            
            ' Amplitude Modulation Enable/EG Bias Sensitivity/Key Volocity
            OP1_AMS = BinToDec(Left(DecToBin(Hex("&H" & buff(sr + 36)), 7), 1))
            OP1_EB = BinToDec(Mid(DecToBin(Hex("&H" & buff(sr + 36)), 7), 2, 3))
            OP1_SN = BinToDec(Right(DecToBin(Hex("&H" & buff(sr + 36)), 7), 3))
            
            OP1_OL = Hex("&H" & buff(sr + 37))
            OP1_FR = Hex("&H" & buff(sr + 38))
            
            'Keyboad Scaling Rate/Detune1
            OP1_KS = BinToDec(Left(DecToBin(Hex("&H" & buff(sr + 39)), 5), 2))
            OP1_DT = BinToDec(Right(DecToBin(Hex("&H" & buff(sr + 39)), 5), 3))
            
            'LFO Sync/FeedBack/Algorithm
            LFO_Sync = BinToDec(Left(DecToBin(Hex("&H" & buff(sr + 40)), 7), 1))
            FB = BinToDec(Mid(DecToBin(Hex("&H" & buff(sr + 40)), 7), 2, 3))
            ALG = BinToDec(Right(DecToBin(Hex("&H" & buff(sr + 40)), 7), 3))
            
            LFO_Speed = Hex("&H" & buff(sr + 41))
            LFO_Delay = Hex("&H" & buff(sr + 42))
            PMD = Hex("&H" & buff(sr + 43))
            AMD = Hex("&H" & buff(sr + 44))
            
            'PMS/AMS/LFO Wave
            PMS = BinToDec(Left(DecToBin(Hex("&H" & buff(sr + 45)), 7), 3))
            AMS = BinToDec(Mid(DecToBin(Hex("&H" & buff(sr + 45)), 7), 4, 2))
            LFO_Wave = BinToDec(Right(DecToBin(Hex("&H" & buff(sr + 45)), 7), 2))

            TRS = Hex("&H" & buff(sr + 46))
            PBR = Hex("&H" & buff(sr + 47))

            'Chorus Switch/Play Mode/Sustain Foot Switch/Portament Foot Switch/Portament Mode
            Chorus = BinToDec(Left(DecToBin(Hex("&H" & buff(sr + 48)), 5), 1))
            POLY_MONO = BinToDec(Mid(DecToBin(Hex("&H" & buff(sr + 48)), 5), 2, 1))
            Sus_S = BinToDec(Mid(DecToBin(Hex("&H" & buff(sr + 48)), 5), 3, 1))
            P_Switch = BinToDec(Mid(DecToBin(Hex("&H" & buff(sr + 48)), 5), 4, 1))
            P_Mode = BinToDec(Right(DecToBin(Hex("&H" & buff(sr + 48)), 5), 1))
            
            P_Time = Hex("&H" & buff(sr + 49))
            FV = Hex("&H" & buff(sr + 50))

            MWP_Range = Hex("&H" & buff(sr + 51))
            MWA_Range = Hex("&H" & buff(sr + 52))
            BPM_Range = Hex("&H" & buff(sr + 53))
            BAM_Range = Hex("&H" & buff(sr + 54))
            BPB_Range = Hex("&H" & buff(sr + 55))
            BEB_Range = Hex("&H" & buff(sr + 56))
                    
            'VoiceNameをコードから文字列に変換
            VoiceName = Chr(Hex("&H" & buff(sr + 57))) & Chr(Hex("&H" & buff(sr + 58))) & Chr(Hex("&H" & buff(sr + 59))) & Chr(Hex("&H" & buff(sr + 60))) & Chr(Hex("&H" & buff(sr + 61))) & Chr(Hex("&H" & buff(sr + 62))) & Chr(Hex("&H" & buff(sr + 63))) & Chr(Hex("&H" & buff(sr + 64))) & Chr(Hex("&H" & buff(sr + 65))) & Chr(Hex("&H" & buff(sr + 66)))
                    
            PR1 = Hex("&H" & buff(sr + 67))
            PR2 = Hex("&H" & buff(sr + 68))
            PR3 = Hex("&H" & buff(sr + 69))
            PL1 = Hex("&H" & buff(sr + 70))
            PL2 = Hex("&H" & buff(sr + 71))
            PL3 = Hex("&H" & buff(sr + 72))
            
            'YAMAHA V2(DX11)のオシレーターデータの取得
            OP4_EGSFT = BinToDec(Left(DecToBin(Hex("&H" & buff(sr + 73)), 6), 2))
            OP4_FIX = BinToDec(Mid(DecToBin(Hex("&H" & buff(sr + 73)), 6), 3, 1))
            OP4_FIXRG = BinToDec(Right(DecToBin(Hex("&H" & buff(sr + 73)), 6), 3))
            OP4_OSW = BinToDec(Left(DecToBin(Hex("&H" & buff(sr + 74)), 7), 3))
            OP4_FINE = BinToDec(Right(DecToBin(Hex("&H" & buff(sr + 74)), 7), 4))
            
            OP2_EGSFT = BinToDec(Left(DecToBin(Hex("&H" & buff(sr + 75)), 6), 2))
            OP2_FIX = BinToDec(Mid(DecToBin(Hex("&H" & buff(sr + 75)), 6), 3, 1))
            OP2_FIXRG = BinToDec(Right(DecToBin(Hex("&H" & buff(sr + 75)), 6), 3))
            OP2_OSW = BinToDec(Left(DecToBin(Hex("&H" & buff(sr + 76)), 7), 3))
            OP2_FINE = BinToDec(Right(DecToBin(Hex("&H" & buff(sr + 76)), 7), 4))

            OP3_EGSFT = BinToDec(Left(DecToBin(Hex("&H" & buff(sr + 77)), 6), 2))
            OP3_FIX = BinToDec(Mid(DecToBin(Hex("&H" & buff(sr + 77)), 6), 3, 1))
            OP3_FIXRG = BinToDec(Right(DecToBin(Hex("&H" & buff(sr + 77)), 6), 3))
            OP3_OSW = BinToDec(Left(DecToBin(Hex("&H" & buff(sr + 78)), 7), 3))
            OP3_FINE = BinToDec(Right(DecToBin(Hex("&H" & buff(sr + 78)), 7), 4))
            
            OP1_EGSFT = BinToDec(Left(DecToBin(Hex("&H" & buff(sr + 79)), 6), 2))
            OP1_FIX = BinToDec(Mid(DecToBin(Hex("&H" & buff(sr + 79)), 6), 3, 1))
            OP1_FIXRG = BinToDec(Right(DecToBin(Hex("&H" & buff(sr + 79)), 6), 3))
            OP1_OSW = BinToDec(Left(DecToBin(Hex("&H" & buff(sr + 80)), 7), 3))
            OP1_FINE = BinToDec(Right(DecToBin(Hex("&H" & buff(sr + 80)), 7), 4))
            
            'YAMAHA V2、V50のフットコントロール、アフタータッチ
            REV = Hex("&H" & buff(sr + 81))
            FC_Pitch = Hex("&H" & buff(sr + 82))
            FC_AMP = Hex("&H" & buff(sr + 83))
            AT_Pitch = Hex("&H" & buff(sr + 84))
            AT_AMP = Hex("&H" & buff(sr + 85))
            AT_PBias = Hex("&H" & buff(sr + 86))
            AT_EGBias = Hex("&H" & buff(sr + 87))
            
            'YAMAHA V50のエフェクト
            E_PRST = Hex("&H" & buff(sr + 91))
            E_TIME = Hex("&H" & buff(sr + 92))
            E_BAL = Hex("&H" & buff(sr + 93))
            E_SEL = Hex("&H" & buff(sr + 94))
            BAL = Hex("&H" & buff(sr + 95))
            OUT_LV = Hex("&H" & buff(sr + 96))
            ST_MIX = Hex("&H" & buff(sr + 97))
            E_PRM1 = Hex("&H" & buff(sr + 98))
            E_PRM2 = Hex("&H" & buff(sr + 99))
            E_PRM3 = Hex("&H" & buff(sr + 100))
            
            OP1_FIXRM = "0"
            OP2_FIXRM = "0"
            OP3_FIXRM = "0"
            OP4_FIXRM = "0"
            
        End If
        
        'Voiceデータをシートへの書き込み
        Sheets(wsTALGet).Activate
        Cells(tr, tc).Value = VoiceName
        Cells(tr, tc + 1).Value = ALG + 1
        Cells(tr, tc + 2).Value = FB

        Cells(tr, tc + 3).Value = OP1_AR
        Cells(tr, tc + 4).Value = OP1_D1R
        Cells(tr, tc + 5).Value = OP1_D1L
        Cells(tr, tc + 6).Value = OP1_D2R
        Cells(tr, tc + 7).Value = OP1_RR
        Cells(tr, tc + 8).Value = OP1_OL
        Cells(tr, tc + 9).Value = OP1_KS
        Cells(tr, tc + 10).Value = OP1_FR
        Cells(tr, tc + 11).Value = OP1_DT - 3
        Cells(tr, tc + 12).Value = OP1_AMS
        Cells(tr, tc + 13).Value = OP1_SN
        Cells(tr, tc + 14).Value = OP1_KL
        Cells(tr, tc + 15).Value = OP1_EB

        Cells(tr, tc + 16).Value = OP2_AR
        Cells(tr, tc + 17).Value = OP2_D1R
        Cells(tr, tc + 18).Value = OP2_D1L
        Cells(tr, tc + 19).Value = OP2_D2R
        Cells(tr, tc + 20).Value = OP2_RR
        Cells(tr, tc + 21).Value = OP2_OL
        Cells(tr, tc + 22).Value = OP2_KS
        Cells(tr, tc + 23).Value = OP2_FR
        Cells(tr, tc + 24).Value = OP2_DT - 3
        Cells(tr, tc + 25).Value = OP2_AMS
        Cells(tr, tc + 26).Value = OP2_SN
        Cells(tr, tc + 27).Value = OP2_KL
        Cells(tr, tc + 28).Value = OP2_EB
        
        Cells(tr, tc + 29).Value = OP3_AR
        Cells(tr, tc + 30).Value = OP3_D1R
        Cells(tr, tc + 31).Value = OP3_D1L
        Cells(tr, tc + 32).Value = OP3_D2R
        Cells(tr, tc + 33).Value = OP3_RR
        Cells(tr, tc + 34).Value = OP3_OL
        Cells(tr, tc + 35).Value = OP3_KS
        Cells(tr, tc + 36).Value = OP3_FR
        Cells(tr, tc + 37).Value = OP3_DT - 3
        Cells(tr, tc + 38).Value = OP3_AMS
        Cells(tr, tc + 39).Value = OP3_SN
        Cells(tr, tc + 40).Value = OP3_KL
        Cells(tr, tc + 41).Value = OP3_EB
        
        Cells(tr, tc + 42).Value = OP4_AR
        Cells(tr, tc + 43).Value = OP4_D1R
        Cells(tr, tc + 44).Value = OP4_D1L
        Cells(tr, tc + 45).Value = OP4_D2R
        Cells(tr, tc + 46).Value = OP4_RR
        Cells(tr, tc + 47).Value = OP4_OL
        Cells(tr, tc + 48).Value = OP4_KS
        Cells(tr, tc + 49).Value = OP4_FR
        Cells(tr, tc + 50).Value = OP4_DT - 3
        Cells(tr, tc + 51).Value = OP4_AMS
        Cells(tr, tc + 52).Value = OP4_SN
        Cells(tr, tc + 53).Value = OP4_KL
        Cells(tr, tc + 54).Value = OP4_EB
                
        Cells(tr, tc + 55).Value = LFO_Speed
        Cells(tr, tc + 56).Value = LFO_Delay
        Cells(tr, tc + 57).Value = PMD
        Cells(tr, tc + 58).Value = AMD
        Cells(tr, tc + 59).Value = LFO_Sync
        Cells(tr, tc + 60).Value = LFO_Wave
        Cells(tr, tc + 61).Value = PMS
        Cells(tr, tc + 62).Value = AMS
        Cells(tr, tc + 63).Value = TRS
        Cells(tr, tc + 64).Value = POLY_MONO
        Cells(tr, tc + 65).Value = PBR
        Cells(tr, tc + 66).Value = P_Mode
        Cells(tr, tc + 67).Value = P_Time
        Cells(tr, tc + 68).Value = FV
        Cells(tr, tc + 69).Value = Sus_S
        Cells(tr, tc + 70).Value = P_Switch
        Cells(tr, tc + 71).Value = Chorus
        Cells(tr, tc + 72).Value = MWP_Range
        Cells(tr, tc + 73).Value = MWA_Range
        Cells(tr, tc + 74).Value = BPM_Range
        Cells(tr, tc + 75).Value = BAM_Range
        Cells(tr, tc + 76).Value = BPB_Range
        Cells(tr, tc + 77).Value = BEB_Range
        Cells(tr, tc + 78).Value = PR1
        Cells(tr, tc + 79).Value = PR2
        Cells(tr, tc + 80).Value = PR3
        Cells(tr, tc + 81).Value = PL1
        Cells(tr, tc + 82).Value = PL2
        Cells(tr, tc + 83).Value = PL3
        
        Cells(tr, tc + 85).Value = OP1_EGSFT
        Cells(tr, tc + 86).Value = OP1_FIX
        Cells(tr, tc + 87).Value = OP1_FIXRG
        Cells(tr, tc + 88).Value = OP1_OSW
        Cells(tr, tc + 89).Value = OP1_FINE
        
        Cells(tr, tc + 90).Value = OP2_EGSFT
        Cells(tr, tc + 91).Value = OP2_FIX
        Cells(tr, tc + 92).Value = OP2_FIXRG
        Cells(tr, tc + 93).Value = OP2_OSW
        Cells(tr, tc + 94).Value = OP2_FINE

        Cells(tr, tc + 95).Value = OP3_EGSFT
        Cells(tr, tc + 96).Value = OP3_FIX
        Cells(tr, tc + 97).Value = OP3_FIXRG
        Cells(tr, tc + 98).Value = OP3_OSW
        Cells(tr, tc + 99).Value = OP3_FINE

        Cells(tr, tc + 100).Value = OP4_EGSFT
        Cells(tr, tc + 101).Value = OP4_FIX
        Cells(tr, tc + 102).Value = OP4_FIXRG
        Cells(tr, tc + 103).Value = OP4_OSW
        Cells(tr, tc + 104).Value = OP4_FINE
        
        Cells(tr, tc + 106).Value = REV
        Cells(tr, tc + 107).Value = FC_Pitch
        Cells(tr, tc + 108).Value = FC_AMP
        Cells(tr, tc + 109).Value = AT_Pitch
        Cells(tr, tc + 110).Value = AT_AMP
        Cells(tr, tc + 111).Value = AT_PBias
        Cells(tr, tc + 112).Value = AT_EGBias
        
        Cells(tr, tc + 114).Value = E_PRST
        Cells(tr, tc + 115).Value = E_TIME
        Cells(tr, tc + 116).Value = E_BAL
        Cells(tr, tc + 117).Value = E_SEL
        Cells(tr, tc + 118).Value = BAL
        Cells(tr, tc + 119).Value = OUT_LV
        Cells(tr, tc + 120).Value = ST_MIX
        Cells(tr, tc + 121).Value = E_PRM1
        Cells(tr, tc + 122).Value = E_PRM2
        Cells(tr, tc + 123).Value = E_PRM3
        
        Cells(tr, tc + 125).Value = OP1_FIXRM
        Cells(tr, tc + 126).Value = OP2_FIXRM
        Cells(tr, tc + 127).Value = OP3_FIXRM
        Cells(tr, tc + 128).Value = OP4_FIXRM

        'カウンターの更新（32Voiceバルクデータの場合、次のVoiceデータまで読み飛ばすため128バイト先へ移動）
        tr = tr + 1
        sr = sr + 128
        
    Next c
    
    
'以下はデバッグ用（OutputDataから出力したSysexファイルのチェックサム確認用コード）
    
'    If buff(3) = 3 Then
'
'        If chksum2 = buff(99) Then
'            Debug.Print "OK"
'            Debug.Print chksum2
'            Debug.Print buff(99)
'        Else
'            Debug.Print "Error"
'            Debug.Print chksum2
'            Debug.Print buff(99)
'        End If
'
'    ElseIf buff(3) = 4 Then
'
'        If chksum2 = buff(4102) Then
'            Debug.Print "OK"
'            Debug.Print chksum2
'            Debug.Print buff(4102)
'        Else
'            Debug.Print "Error"
'            Debug.Print chksum2
'            Debug.Print buff(4102)
'        End If
'
'    End If
        
End Sub

 
'********************************************
'
' 関数名：BinaryMain_test
'
' 機能：バイナリファイルを読み込み
' 　　　BulkDataシートに出力
'
' 呼び出し関数：ReadBinaryFile
'
'********************************************
Sub BinaryMain_test()
    
    Dim strFilePath As String
    Dim strFileName As String
    Dim strFile As String
    '//ファイル名を生成
    
    Sheets("Menu").Select
    strFilePath = Cells(45, 5).Value
    strFileName = Cells(46, 5).Value

    If strFilePath = "" Then
        strFilePath = ThisWorkbook.Path
    End If
    
    If strFileName = "" Then
        MsgBox "ファイル名が指定されていません。"
        
    Else
        strFile = strFilePath & "\" & strFileName
        
        If Dir(strFile) <> "" Then
            
    '//全て(数式、文字列、書式、コメント、アウトライン)クリア
    
            Sheets("BulkData").Activate
            Cells.Select
            Selection.Clear
            Selection.Font.Name = "ＭＳ　ゴシック"
            Selection.Font.Size = 12
            Range("A1").Select
  
   
    '//バイナリファイルを読み込み（テストデータファイルを読み込み）
            Call ReadBinaryFile(strFile)
            
        Else
            MsgBox strFile & vbCrLf & _
                    "が存在しません"
        End If
    End If

End Sub

'********************************************
'
' 関数名：ReadBinaryFile
'
' 機能：バイナリファイルを読み取り
' 　　　BulkDataシートに出力する
'
' 引数：strfil：読み込むファイル名
'
'********************************************
Sub ReadBinaryFile(ByVal strfil As String)
    Dim buff() As Byte
    Dim fp As Long
    Dim filesize As Long, NowLoc As Long
    Dim idx As Long, gyo As Long
    Dim strBinary As String
    '//FreeFile関数で使用可能なファイル番号を割り当て
    fp = FreeFile
    '//ファイルを開く
    Open strfil For Binary As #fp
    '//ファイルサイズ分の読み込み領域を確保して読み込む場合の実装例
    'ReDim buff(FileLen(strfil))
    'Get #fp, 1, buff
    '//実装例ここまで
    '//ヘッダ情報をシートに出力
    
    Cells(1, 1) = "         00 01 02 03 04 05 06 07 08 09 0A 0B 0C 0D 0E 0F "
    Cells(2, 1) = "---------------------------------------------------------"
    '//ファイルの終端まで指定サイズ(最大16バイト)繰り返し読み込む
    gyo = 3
    Do While NowLoc < LOF(fp)
        '//最大16バイト分の領域を確保し初期化
        If (LOF(fp) - NowLoc) >= 16 Then
            '//残りのファイルサイズが16バイト以上のとき
            ReDim buff(15)
        Else
            '//最終読み込み時(497バイト～500バイト目)は残りのファイルサイズが16未満
            ReDim buff(LOF(fp) - NowLoc - 1)
        End If
        '//データを読み込み
        Get #fp, , buff
        '//表示用のアドレスを生成("00000000"とHexの戻り値を連結した文字列の右から8文字)
        strBinary = Right("00000000" & Hex(NowLoc), 8) & " "
        '//現在位置をを保持する(ループBreak判定用)
        NowLoc = Loc(fp)
        '//出力文字列を生成
        For idx = 0 To UBound(buff)
            strBinary = strBinary + Right("00" & Hex(buff(idx)), 2) + " "
        Next
        '//シートの1列目に結果を表示
        Cells(gyo, 1) = strBinary
        gyo = gyo + 1
    Loop
    '//ファイルを閉じる
    Close (fp)
End Sub


'********************************************
'
' 関数名：BinaryMainV50
'
' 機能：V50用Sysexファイルを読み込み
' 　　　SysexDataシートに出力する
'
' 呼び出し関数：ReadV50SysexFile
'
'********************************************
Sub BinaryMainV50()
    Dim strFilePath As String
    Dim strFileName As String
    Dim strFile As String
    '//ファイル名を生成
    
    Sheets("Menu").Select
    strFilePath = Cells(77, 5).Value
    strFileName = Cells(78, 5).Value

    If strFilePath = "" Then
        strFilePath = ThisWorkbook.Path
    End If
    
    If strFileName = "" Then
        MsgBox "ファイル名が指定されていません。"
        
    Else
        strFile = strFilePath & "\" & strFileName
        
        If Dir(strFile) <> "" Then
  
        '//バイナリファイルを読み込み（Sysexファイルを読み込み）
            Call ReadV50SysexFile(strFile)
            Sheets("SysexData").Select
            MsgBox "Sysexデータの読み込みが完了しました。"
            
        Else
            MsgBox strFile & vbCrLf & _
                    "が存在しません"
        End If
    End If

End Sub

'********************************************
'
' 関数名：ReadV50SysexFile
'
' 機能：V50用Sysexファイルを読み込み
' 　　　SysexDataシートに出力する
'
' 引数：strfil：読み込みファイル名
'
'********************************************
Sub ReadV50SysexFile(ByVal strfil As String)
    Dim buff() As Byte
    Dim fp As Long
    Dim filesize As Long, NowLoc As Long
    Dim idx As Long, gyo As Long
    Dim strBinary As String
    Dim wsTALGet As String
    
    Dim VoiceName As String
    Dim ALG As Integer, FB As Integer
    Dim OP1_AR, OP1_D1R, OP1_D1L, OP1_D2R, OP1_RR, OP1_OL, OP1_KS, OP1_FR, OP1_DT, OP1_AMS, OP1_SN, OP1_SL, OP1_TL, OP1_ML, OP1_ODT, OP1_KL, OP1_EB
    Dim OP2_AR, OP2_D1R, OP2_D1L, OP2_D2R, OP2_RR, OP2_OL, OP2_KS, OP2_FR, OP2_DT, OP2_AMS, OP2_SN, OP2_SL, OP2_TL, OP2_ML, OP2_ODT, OP2_KL, OP2_EB
    Dim OP3_AR, OP3_D1R, OP3_D1L, OP3_D3R, OP3_RR, OP3_OL, OP3_KS, OP3_FR, OP3_DT, OP3_AMS, OP3_SN, OP3_SL, OP3_TL, OP3_ML, OP3_ODT, OP3_KL, OP3_EB
    Dim OP4_AR, OP4_D1R, OP4_D1L, OP4_D2R, OP4_RR, OP4_OL, OP4_KS, OP4_FR, OP4_DT, OP4_AMS, OP4_SN, OP4_SL, OP4_TL, OP4_ML, OP4_ODT, OP4_KL, OP4_EB
    Dim LFO_Speed, LFO_Delay, PMD, AMD, LFO_Sync, LFO_Wave, PMS, AMS, TRS, POLY_MONO, PBR, P_Mode, P_Time, FV, Sus_S, P_Switch, Chorus, MWP_Range, MWA_Range
    Dim BPM_Range, BAM_Range, BPB_Range, BEB_Range, PR1, PR2, PR3, PL1, PL2, PL3, OP_ONOFF
    
    Dim OP1_EGSFT, OP1_FIX, OP1_FIXRM, OP1_FIXRG, OP1_OSW, OP1_FINE
    Dim OP2_EGSFT, OP2_FIX, OP2_FIXRM, OP2_FIXRG, OP2_OSW, OP2_FINE
    Dim OP3_EGSFT, OP3_FIX, OP3_FIXRM, OP3_FIXRG, OP3_OSW, OP3_FINE
    Dim OP4_EGSFT, OP4_FIX, OP4_FIXRM, OP4_FIXRG, OP4_OSW, OP4_FINE
    Dim REV, FC_Pitch, FC_AMP, AT_Pitch, AT_AMP, AT_PBias, AT_EGBias
    
    Dim E_PRST, E_TIME, E_BAL, E_SEL, BAL, OUT_LV, ST_MIX, E_PRM1, E_PRM2, E_PRM3
    
    Dim Data_set, c, cc
    Dim sr As Long, tr As Long, tc As Long
    'Dim chksum2 As Long
    
    
    'データ出力用シート
    wsTALGet = "SysexData"
    
    Sheets(wsTALGet).Activate
    '    For Each cc In Range("A2:CG33", "CI2:DB33")
    '        If cc.Locked = False Then
    '            cc.Value = ""
    '        End If
    '    Next cc
     
    ActiveSheet.Unprotect
    Range("A2:CG33", "CI2:DB33").ClearContents
    Range("DD2:DJ33", "DL2:DU33").ClearContents
    Range("DW2:DZ33").ClearContents
    ActiveSheet.Protect
    
    '//FreeFile関数で使用可能なファイル番号を割り当て
    fp = FreeFile
    '//ファイルを開く
    Open strfil For Binary As #fp
    '//ファイルサイズ分の読み込み領域を確保して読み込む場合の実装例
    ReDim buff(FileLen(strfil))
    Get #fp, 1, buff
    '//実装例ここまで

    '//ファイルを閉じる
    Close (fp)
    
    '読み込みカウンター初期化（ヘッダーを読み飛ばし14バイト目から読み込む）
    sr = 13
    
    'シート書き込み用カウンター初期化
    tr = 2
    tc = 2
    
    'データセットは32Voiceのみ（25Voice + InitVoice×7 = 32Voice）
        Data_set = 32


    'chksum2 = 0

    For c = 1 To Data_set

            OP4_AR = Hex("&H" & buff(sr))
            OP4_D1R = Hex("&H" & buff(sr + 1))
            OP4_D2R = Hex("&H" & buff(sr + 2))
            OP4_RR = Hex("&H" & buff(sr + 3))
            OP4_D1L = Hex("&H" & buff(sr + 4))
            OP4_KL = Hex("&H" & buff(sr + 5))
            
            ' Amplitude Modulation Enable/EG Bias Sensitivity/Key Volocity
            OP4_AMS = BinToDec(Left(DecToBin(Hex("&H" & buff(sr + 6)), 7), 1))
            OP4_EB = BinToDec(Mid(DecToBin(Hex("&H" & buff(sr + 6)), 7), 2, 3))
            OP4_SN = BinToDec(Right(DecToBin(Hex("&H" & buff(sr + 6)), 7), 3))
            
            OP4_OL = Hex("&H" & buff(sr + 7))
            OP4_FR = Hex("&H" & buff(sr + 8))
            
            'Keyboad Scaling Rate/Detune1
            OP4_KS = BinToDec(Left(DecToBin(Hex("&H" & buff(sr + 9)), 5), 2))
            OP4_DT = BinToDec(Right(DecToBin(Hex("&H" & buff(sr + 9)), 5), 3))
            
            OP2_AR = Hex("&H" & buff(sr + 10))
            OP2_D1R = Hex("&H" & buff(sr + 11))
            OP2_D2R = Hex("&H" & buff(sr + 12))
            OP2_RR = Hex("&H" & buff(sr + 13))
            OP2_D1L = Hex("&H" & buff(sr + 14))
            OP2_KL = Hex("&H" & buff(sr + 15))
            
            ' Amplitude Modulation Enable/EG Bias Sensitivity/Key Volocity
            OP2_AMS = BinToDec(Left(DecToBin(Hex("&H" & buff(sr + 16)), 7), 1))
            OP2_EB = BinToDec(Mid(DecToBin(Hex("&H" & buff(sr + 16)), 7), 2, 3))
            OP2_SN = BinToDec(Right(DecToBin(Hex("&H" & buff(sr + 16)), 7), 3))
            
            OP2_OL = Hex("&H" & buff(sr + 17))
            OP2_FR = Hex("&H" & buff(sr + 18))
            
            'Keyboad Scaling Rate/Detune1
            OP2_KS = BinToDec(Left(DecToBin(Hex("&H" & buff(sr + 19)), 5), 2))
            OP2_DT = BinToDec(Right(DecToBin(Hex("&H" & buff(sr + 19)), 5), 3))
            
            OP3_AR = Hex("&H" & buff(sr + 20))
            OP3_D1R = Hex("&H" & buff(sr + 21))
            OP3_D2R = Hex("&H" & buff(sr + 22))
            OP3_RR = Hex("&H" & buff(sr + 23))
            OP3_D1L = Hex("&H" & buff(sr + 24))
            OP3_KL = Hex("&H" & buff(sr + 25))
            
            ' Amplitude Modulation Enable/EG Bias Sensitivity/Key Volocity
            OP3_AMS = BinToDec(Left(DecToBin(Hex("&H" & buff(sr + 26)), 7), 1))
            OP3_EB = BinToDec(Mid(DecToBin(Hex("&H" & buff(sr + 26)), 7), 2, 3))
            OP3_SN = BinToDec(Right(DecToBin(Hex("&H" & buff(sr + 26)), 7), 3))
            
            OP3_OL = Hex("&H" & buff(sr + 27))
            OP3_FR = Hex("&H" & buff(sr + 28))
            
            'Keyboad Scaling Rate/Detune1
            OP3_KS = BinToDec(Left(DecToBin(Hex("&H" & buff(sr + 29)), 5), 2))
            OP3_DT = BinToDec(Right(DecToBin(Hex("&H" & buff(sr + 29)), 5), 3))
            
            OP1_AR = Hex("&H" & buff(sr + 30))
            OP1_D1R = Hex("&H" & buff(sr + 31))
            OP1_D2R = Hex("&H" & buff(sr + 32))
            OP1_RR = Hex("&H" & buff(sr + 33))
            OP1_D1L = Hex("&H" & buff(sr + 34))
            OP1_KL = Hex("&H" & buff(sr + 35))
            
            ' Amplitude Modulation Enable/EG Bias Sensitivity/Key Volocity
            OP1_AMS = BinToDec(Left(DecToBin(Hex("&H" & buff(sr + 36)), 7), 1))
            OP1_EB = BinToDec(Mid(DecToBin(Hex("&H" & buff(sr + 36)), 7), 2, 3))
            OP1_SN = BinToDec(Right(DecToBin(Hex("&H" & buff(sr + 36)), 7), 3))
            
            OP1_OL = Hex("&H" & buff(sr + 37))
            OP1_FR = Hex("&H" & buff(sr + 38))
            
            'Keyboad Scaling Rate/Detune1
            OP1_KS = BinToDec(Left(DecToBin(Hex("&H" & buff(sr + 39)), 5), 2))
            OP1_DT = BinToDec(Right(DecToBin(Hex("&H" & buff(sr + 39)), 5), 3))
            
            'LFO Sync/FeedBack/Algorithm
            LFO_Sync = BinToDec(Left(DecToBin(Hex("&H" & buff(sr + 40)), 7), 1))
            FB = BinToDec(Mid(DecToBin(Hex("&H" & buff(sr + 40)), 7), 2, 3))
            ALG = BinToDec(Right(DecToBin(Hex("&H" & buff(sr + 40)), 7), 3))
            
            LFO_Speed = Hex("&H" & buff(sr + 41))
            LFO_Delay = Hex("&H" & buff(sr + 42))
            PMD = Hex("&H" & buff(sr + 43))
            AMD = Hex("&H" & buff(sr + 44))
            
            'PMS/AMS/LFO Wave
            PMS = BinToDec(Left(DecToBin(Hex("&H" & buff(sr + 45)), 7), 3))
            AMS = BinToDec(Mid(DecToBin(Hex("&H" & buff(sr + 45)), 7), 4, 2))
            LFO_Wave = BinToDec(Right(DecToBin(Hex("&H" & buff(sr + 45)), 7), 2))

            TRS = Hex("&H" & buff(sr + 46))
            PBR = Hex("&H" & buff(sr + 47))

            'Chorus Switch/Play Mode/Sustain Foot Switch/Portament Foot Switch/Portament Mode
            Chorus = BinToDec(Left(DecToBin(Hex("&H" & buff(sr + 48)), 5), 1))
            POLY_MONO = BinToDec(Mid(DecToBin(Hex("&H" & buff(sr + 48)), 5), 2, 1))
            Sus_S = BinToDec(Mid(DecToBin(Hex("&H" & buff(sr + 48)), 5), 3, 1))
            P_Switch = BinToDec(Mid(DecToBin(Hex("&H" & buff(sr + 48)), 5), 4, 1))
            P_Mode = BinToDec(Right(DecToBin(Hex("&H" & buff(sr + 48)), 5), 1))
            
            P_Time = Hex("&H" & buff(sr + 49))
            FV = Hex("&H" & buff(sr + 50))

            MWP_Range = Hex("&H" & buff(sr + 51))
            MWA_Range = Hex("&H" & buff(sr + 52))
            BPM_Range = Hex("&H" & buff(sr + 53))
            BAM_Range = Hex("&H" & buff(sr + 54))
            BPB_Range = Hex("&H" & buff(sr + 55))
            BEB_Range = Hex("&H" & buff(sr + 56))
                    
            'VoiceNameをコードから文字列に変換
            VoiceName = Chr(Hex("&H" & buff(sr + 57))) & Chr(Hex("&H" & buff(sr + 58))) & Chr(Hex("&H" & buff(sr + 59))) & Chr(Hex("&H" & buff(sr + 60))) & Chr(Hex("&H" & buff(sr + 61))) & Chr(Hex("&H" & buff(sr + 62))) & Chr(Hex("&H" & buff(sr + 63))) & Chr(Hex("&H" & buff(sr + 64))) & Chr(Hex("&H" & buff(sr + 65))) & Chr(Hex("&H" & buff(sr + 66)))
                    
            PR1 = Hex("&H" & buff(sr + 67))
            PR2 = Hex("&H" & buff(sr + 68))
            PR3 = Hex("&H" & buff(sr + 69))
            PL1 = Hex("&H" & buff(sr + 70))
            PL2 = Hex("&H" & buff(sr + 71))
            PL3 = Hex("&H" & buff(sr + 72))
            
            'YAMAHA V2(DX11)のオシレーターデータの取得
            OP4_FIXRM = BinToDec(Left(DecToBin(Hex("&H" & buff(sr + 73)), 7), 1))
            OP4_EGSFT = BinToDec(Mid(DecToBin(Hex("&H" & buff(sr + 73)), 7), 2, 2))
            OP4_FIX = BinToDec(Mid(DecToBin(Hex("&H" & buff(sr + 73)), 7), 4, 1))
            OP4_FIXRG = BinToDec(Right(DecToBin(Hex("&H" & buff(sr + 73)), 7), 3))
            OP4_OSW = BinToDec(Left(DecToBin(Hex("&H" & buff(sr + 74)), 7), 3))
            OP4_FINE = BinToDec(Right(DecToBin(Hex("&H" & buff(sr + 74)), 7), 4))
            
            OP2_FIXRM = BinToDec(Left(DecToBin(Hex("&H" & buff(sr + 75)), 7), 1))
            OP2_EGSFT = BinToDec(Mid(DecToBin(Hex("&H" & buff(sr + 75)), 7), 2, 2))
            OP2_FIX = BinToDec(Mid(DecToBin(Hex("&H" & buff(sr + 75)), 7), 4, 1))
            OP2_FIXRG = BinToDec(Right(DecToBin(Hex("&H" & buff(sr + 75)), 7), 3))
            OP2_OSW = BinToDec(Left(DecToBin(Hex("&H" & buff(sr + 76)), 7), 3))
            OP2_FINE = BinToDec(Right(DecToBin(Hex("&H" & buff(sr + 76)), 7), 4))

            OP3_FIXRM = BinToDec(Left(DecToBin(Hex("&H" & buff(sr + 77)), 7), 1))
            OP3_EGSFT = BinToDec(Mid(DecToBin(Hex("&H" & buff(sr + 77)), 7), 2, 2))
            OP3_FIX = BinToDec(Mid(DecToBin(Hex("&H" & buff(sr + 77)), 7), 4, 1))
            OP3_FIXRG = BinToDec(Right(DecToBin(Hex("&H" & buff(sr + 77)), 7), 3))
            OP3_OSW = BinToDec(Left(DecToBin(Hex("&H" & buff(sr + 78)), 7), 3))
            OP3_FINE = BinToDec(Right(DecToBin(Hex("&H" & buff(sr + 78)), 7), 4))
            
            OP1_FIXRM = BinToDec(Left(DecToBin(Hex("&H" & buff(sr + 79)), 7), 1))
            OP1_EGSFT = BinToDec(Mid(DecToBin(Hex("&H" & buff(sr + 79)), 7), 2, 2))
            OP1_FIX = BinToDec(Mid(DecToBin(Hex("&H" & buff(sr + 79)), 7), 4, 1))
            OP1_FIXRG = BinToDec(Right(DecToBin(Hex("&H" & buff(sr + 79)), 7), 3))
            OP1_OSW = BinToDec(Left(DecToBin(Hex("&H" & buff(sr + 80)), 7), 3))
            OP1_FINE = BinToDec(Right(DecToBin(Hex("&H" & buff(sr + 80)), 7), 4))
            
            'YAMAHA V2、V50のフットコントロール、アフタータッチ
            REV = Hex("&H" & buff(sr + 81))
            FC_Pitch = Hex("&H" & buff(sr + 82))
            FC_AMP = Hex("&H" & buff(sr + 83))
            AT_Pitch = Hex("&H" & buff(sr + 84))
            AT_AMP = Hex("&H" & buff(sr + 85))
            AT_PBias = Hex("&H" & buff(sr + 86))
            AT_EGBias = Hex("&H" & buff(sr + 87))
            
            'YAMAHA V50のエフェクト
            E_PRST = Hex("&H" & buff(sr + 91))
            E_TIME = Hex("&H" & buff(sr + 92))
            E_BAL = Hex("&H" & buff(sr + 93))
            E_SEL = Hex("&H" & buff(sr + 94))
            BAL = Hex("&H" & buff(sr + 95))
            OUT_LV = Hex("&H" & buff(sr + 96))
            ST_MIX = Hex("&H" & buff(sr + 97))
            E_PRM1 = Hex("&H" & buff(sr + 98))
            E_PRM2 = Hex("&H" & buff(sr + 99))
            E_PRM3 = Hex("&H" & buff(sr + 100))
            
        
        'Voiceデータをシートへの書き込み
        Sheets(wsTALGet).Activate
        Cells(tr, tc).Value = VoiceName
        Cells(tr, tc + 1).Value = ALG + 1
        Cells(tr, tc + 2).Value = FB

        Cells(tr, tc + 3).Value = OP1_AR
        Cells(tr, tc + 4).Value = OP1_D1R
        Cells(tr, tc + 5).Value = OP1_D1L
        Cells(tr, tc + 6).Value = OP1_D2R
        Cells(tr, tc + 7).Value = OP1_RR
        Cells(tr, tc + 8).Value = OP1_OL
        Cells(tr, tc + 9).Value = OP1_KS
        Cells(tr, tc + 10).Value = OP1_FR
        Cells(tr, tc + 11).Value = OP1_DT - 3
        Cells(tr, tc + 12).Value = OP1_AMS
        Cells(tr, tc + 13).Value = OP1_SN
        Cells(tr, tc + 14).Value = OP1_KL
        Cells(tr, tc + 15).Value = OP1_EB

        Cells(tr, tc + 16).Value = OP2_AR
        Cells(tr, tc + 17).Value = OP2_D1R
        Cells(tr, tc + 18).Value = OP2_D1L
        Cells(tr, tc + 19).Value = OP2_D2R
        Cells(tr, tc + 20).Value = OP2_RR
        Cells(tr, tc + 21).Value = OP2_OL
        Cells(tr, tc + 22).Value = OP2_KS
        Cells(tr, tc + 23).Value = OP2_FR
        Cells(tr, tc + 24).Value = OP2_DT - 3
        Cells(tr, tc + 25).Value = OP2_AMS
        Cells(tr, tc + 26).Value = OP2_SN
        Cells(tr, tc + 27).Value = OP2_KL
        Cells(tr, tc + 28).Value = OP2_EB
        
        Cells(tr, tc + 29).Value = OP3_AR
        Cells(tr, tc + 30).Value = OP3_D1R
        Cells(tr, tc + 31).Value = OP3_D1L
        Cells(tr, tc + 32).Value = OP3_D2R
        Cells(tr, tc + 33).Value = OP3_RR
        Cells(tr, tc + 34).Value = OP3_OL
        Cells(tr, tc + 35).Value = OP3_KS
        Cells(tr, tc + 36).Value = OP3_FR
        Cells(tr, tc + 37).Value = OP3_DT - 3
        Cells(tr, tc + 38).Value = OP3_AMS
        Cells(tr, tc + 39).Value = OP3_SN
        Cells(tr, tc + 40).Value = OP3_KL
        Cells(tr, tc + 41).Value = OP3_EB
        
        Cells(tr, tc + 42).Value = OP4_AR
        Cells(tr, tc + 43).Value = OP4_D1R
        Cells(tr, tc + 44).Value = OP4_D1L
        Cells(tr, tc + 45).Value = OP4_D2R
        Cells(tr, tc + 46).Value = OP4_RR
        Cells(tr, tc + 47).Value = OP4_OL
        Cells(tr, tc + 48).Value = OP4_KS
        Cells(tr, tc + 49).Value = OP4_FR
        Cells(tr, tc + 50).Value = OP4_DT - 3
        Cells(tr, tc + 51).Value = OP4_AMS
        Cells(tr, tc + 52).Value = OP4_SN
        Cells(tr, tc + 53).Value = OP4_KL
        Cells(tr, tc + 54).Value = OP4_EB
                
        Cells(tr, tc + 55).Value = LFO_Speed
        Cells(tr, tc + 56).Value = LFO_Delay
        Cells(tr, tc + 57).Value = PMD
        Cells(tr, tc + 58).Value = AMD
        Cells(tr, tc + 59).Value = LFO_Sync
        Cells(tr, tc + 60).Value = LFO_Wave
        Cells(tr, tc + 61).Value = PMS
        Cells(tr, tc + 62).Value = AMS
        Cells(tr, tc + 63).Value = TRS
        Cells(tr, tc + 64).Value = POLY_MONO
        Cells(tr, tc + 65).Value = PBR
        Cells(tr, tc + 66).Value = P_Mode
        Cells(tr, tc + 67).Value = P_Time
        Cells(tr, tc + 68).Value = FV
        Cells(tr, tc + 69).Value = Sus_S
        Cells(tr, tc + 70).Value = P_Switch
        Cells(tr, tc + 71).Value = Chorus
        Cells(tr, tc + 72).Value = MWP_Range
        Cells(tr, tc + 73).Value = MWA_Range
        Cells(tr, tc + 74).Value = BPM_Range
        Cells(tr, tc + 75).Value = BAM_Range
        Cells(tr, tc + 76).Value = BPB_Range
        Cells(tr, tc + 77).Value = BEB_Range
        Cells(tr, tc + 78).Value = PR1
        Cells(tr, tc + 79).Value = PR2
        Cells(tr, tc + 80).Value = PR3
        Cells(tr, tc + 81).Value = PL1
        Cells(tr, tc + 82).Value = PL2
        Cells(tr, tc + 83).Value = PL3
        
        Cells(tr, tc + 85).Value = OP1_EGSFT
        Cells(tr, tc + 86).Value = OP1_FIX
        Cells(tr, tc + 87).Value = OP1_FIXRG
        Cells(tr, tc + 88).Value = OP1_OSW
        Cells(tr, tc + 89).Value = OP1_FINE
        
        Cells(tr, tc + 90).Value = OP2_EGSFT
        Cells(tr, tc + 91).Value = OP2_FIX
        Cells(tr, tc + 92).Value = OP2_FIXRG
        Cells(tr, tc + 93).Value = OP2_OSW
        Cells(tr, tc + 94).Value = OP2_FINE

        Cells(tr, tc + 95).Value = OP3_EGSFT
        Cells(tr, tc + 96).Value = OP3_FIX
        Cells(tr, tc + 97).Value = OP3_FIXRG
        Cells(tr, tc + 98).Value = OP3_OSW
        Cells(tr, tc + 99).Value = OP3_FINE

        Cells(tr, tc + 100).Value = OP4_EGSFT
        Cells(tr, tc + 101).Value = OP4_FIX
        Cells(tr, tc + 102).Value = OP4_FIXRG
        Cells(tr, tc + 103).Value = OP4_OSW
        Cells(tr, tc + 104).Value = OP4_FINE
        
        Cells(tr, tc + 106).Value = REV
        Cells(tr, tc + 107).Value = FC_Pitch
        Cells(tr, tc + 108).Value = FC_AMP
        Cells(tr, tc + 109).Value = AT_Pitch
        Cells(tr, tc + 110).Value = AT_AMP
        Cells(tr, tc + 111).Value = AT_PBias
        Cells(tr, tc + 112).Value = AT_EGBias
        
        Cells(tr, tc + 114).Value = E_PRST
        Cells(tr, tc + 115).Value = E_TIME
        Cells(tr, tc + 116).Value = E_BAL
        Cells(tr, tc + 117).Value = E_SEL
        Cells(tr, tc + 118).Value = BAL
        Cells(tr, tc + 119).Value = OUT_LV
        Cells(tr, tc + 120).Value = ST_MIX
        Cells(tr, tc + 121).Value = E_PRM1
        Cells(tr, tc + 122).Value = E_PRM2
        Cells(tr, tc + 123).Value = E_PRM3
        
        Cells(tr, tc + 125).Value = OP1_FIXRM
        Cells(tr, tc + 126).Value = OP2_FIXRM
        Cells(tr, tc + 127).Value = OP3_FIXRM
        Cells(tr, tc + 128).Value = OP4_FIXRM

        'カウンターの更新（32Voiceバルクデータの場合、次のVoiceデータまで読み飛ばすため128バイト先へ移動）
        tr = tr + 1
        sr = sr + 128
        
    Next c
    
    
'以下はデバッグ用（OutputDataから出力したSysexファイルのチェックサム確認用コード）
    
'    If buff(3) = 3 Then
'
'        If chksum2 = buff(99) Then
'            Debug.Print "OK"
'            Debug.Print chksum2
'            Debug.Print buff(99)
'        Else
'            Debug.Print "Error"
'            Debug.Print chksum2
'            Debug.Print buff(99)
'        End If
'
'    ElseIf buff(3) = 4 Then
'
'        If chksum2 = buff(4102) Then
'            Debug.Print "OK"
'            Debug.Print chksum2
'            Debug.Print buff(4102)
'        Else
'            Debug.Print "Error"
'            Debug.Print chksum2
'            Debug.Print buff(4102)
'        End If
'
'    End If
        
End Sub
