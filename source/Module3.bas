Attribute VB_Name = "Module3"

'********************************************
'
' 関数名：Create_DX21_SV_syx
'
' 機能：1VoiceデータをSysexファイルに出力
'
' 呼び出し関数：Write_DX21syx
'
'********************************************
Sub Create_DX21_SV_syx()

    Dim strFilePath As String
    Dim strFileName As String
    Dim strFile As String
    Dim Msg, Style, Title, Response
    
    Style = vbOKCancel
    Title = "エラー"
    
    '//ファイル名を生成
    
    Sheets("Menu").Select
    strFilePath = Cells(31, 5).Value
    strFileName = Cells(32, 5).Value

    If strFilePath = "" Then
        strFilePath = ThisWorkbook.Path
    End If
    
    If strFileName = "" Then
        MsgBox "ファイル名が指定されていません。"
        
    Else
        strFile = strFilePath & "\" & strFileName
                
        If Dir(strFile) = "" Then
        
            Call Write_DX21syx(1, strFile)
            Sheets("Menu").Select
            MsgBox "Sysexデータの書き出しが完了しました。"
            
        Else
        
            Msg = strFile & vbCrLf & _
                    "がすでに存在します。上書きしてもよろしいですか？"

            Response = MsgBox(Msg, Style, Title)
            
            If Response = vbOK Then
            
                Kill strFile
                Call Write_DX21syx(1, strFile)
                Sheets("Menu").Select
                MsgBox "Sysexデータの書き出しが完了しました。"
                
            ElseIf Response = vbCancel Then
            
                MsgBox "Sysexデータの書き出しをキャンセルしました。"
            
            Else
                
                MsgBox "Sysexデータの書き出しを中止しました。"
            
            End If
        End If
    End If
End Sub

'********************************************
'
' 関数名：Create_DX21_MV_syx
'
' 機能：32VoiceデータをSysexファイルに出力
'
' 呼び出し関数：Write_DX21syx
'
'********************************************
Sub Create_DX21_MV_syx()

    Dim strFilePath As String
    Dim strFileName As String
    Dim strFile As String
    Dim Msg, Style, Title, Response
    
    Style = vbOKCancel
    Title = "エラー"
    
    '//ファイル名を生成
    
    Sheets("Menu").Select
    strFilePath = Cells(38, 5).Value
    strFileName = Cells(39, 5).Value

    If strFilePath = "" Then
        strFilePath = ThisWorkbook.Path
    End If
    
    If strFileName = "" Then
        MsgBox "ファイル名が指定されていません。"
        
    Else
        strFile = strFilePath & "\" & strFileName
                
        If Dir(strFile) = "" Then
        
            Call Write_DX21syx(32, strFile)
            Sheets("Menu").Select
            MsgBox "Sysexデータの書き出しが完了しました。"
            
        Else
        
            Msg = strFile & vbCrLf & _
                    "がすでに存在します。上書きしてもよろしいですか？"

            Response = MsgBox(Msg, Style, Title)
            
            If Response = vbOK Then
            
                Kill strFile
                Call Write_DX21syx(32, strFile)
                Sheets("Menu").Select
                MsgBox "Sysexデータの書き出しが完了しました。"
                
            ElseIf Response = vbCancel Then
            
                MsgBox "Sysexデータの書き出しをキャンセルしました。"
            
            Else
                
                MsgBox "Sysexデータの書き出しを中止しました。"
            
            End If
        End If
    End If
End Sub

'********************************************
'
' 関数名：Write_DX21syx
'
' 機能：OPN/OPM→DX21に変換した音色データを
' 　　　OutputDataシートから読み込み
' 　　　DX21用Sysexファイルに書込む
'
' 引数：DataSet :1または32
' 　　　filename:Sysexファイル名
'
'********************************************
Sub Write_DX21syx(DataSet As Integer, filename As String)
    
    Dim wsSource As String
    Dim VoiceName As String
    Dim ALG As Long, FB As Long
    Dim OP1_AR, OP1_D1R, OP1_D1L, OP1_D2R, OP1_RR, OP1_OL, OP1_KS As Long, OP1_FR, OP1_DT As Long, OP1_AMS As Long, OP1_SN As Long, OP1_SL, OP1_TL, OP1_ML, OP1_ODT, OP1_KL, OP1_EB As Long
    Dim OP2_AR, OP2_D1R, OP2_D1L, OP2_D2R, OP2_RR, OP2_OL, OP2_KS As Long, OP2_FR, OP2_DT As Long, OP2_AMS As Long, OP2_SN As Long, OP2_SL, OP2_TL, OP2_ML, OP2_ODT, OP2_KL, OP2_EB As Long
    Dim OP3_AR, OP3_D1R, OP3_D1L, OP3_D3R, OP3_RR, OP3_OL, OP3_KS As Long, OP3_FR, OP3_DT As Long, OP3_AMS As Long, OP3_SN As Long, OP3_SL, OP3_TL, OP3_ML, OP3_ODT, OP3_KL, OP3_EB As Long
    Dim OP4_AR, OP4_D1R, OP4_D1L, OP4_D2R, OP4_RR, OP4_OL, OP4_KS As Long, OP4_FR, OP4_DT As Long, OP4_AMS As Long, OP4_SN As Long, OP4_SL, OP4_TL, OP4_ML, OP4_ODT, OP4_KL, OP4_EB As Long
    Dim LFO_Speed, LFO_Delay, PMD, AMD, LFO_Sync As Long, LFO_Wave As Long, PMS As Long, AMS As Long, TRS, POLY_MONO As Long, PBR, P_Mode As Long, P_Time, FV, Sus_S As Long, P_Switch As Long, Chorus As Long, MWP_Range As Long, MWA_Range As Long
    Dim BPM_Range, BAM_Range, BPB_Range, BEB_Range, PR1, PR2, PR3, PL1, PL2, PL3 As Long
        
    Dim data_str As String
    Dim sv_hdr, mv_hdr  As String
    Dim VoiceData_str() As String
    Dim VoiceData_byt() As Byte
    Dim chksum_data() As String
    Dim chksum_byte() As Byte
        
    Dim chksum As Long
        
    '1Voice用ヘッダー
    sv_hdr = "F0 43 00 03 00 5D"
    
    '32Voiceバルクデータ用ヘッダー
    mv_hdr = "F0 43 00 04 20 00"
    
    'チェックサム初期化
    chksum = 0
    
    'データストリング初期化
    data_str = ""
    
    'ソースデータシート読み込み用カウンターの初期化
    sr = 2
    sc = 2
    'データー生成用カウンターの初期化
    tr = 1
    tc = 1
    
    '元データシート
    wsSource = "OutputData"
    
    If DataSet = 1 Then

        data_str = sv_hdr
        
    ElseIf DataSet = 32 Then

        data_str = mv_hdr
        
    End If

    For c = 1 To DataSet
    
    'OutpuDataシートの値を読み込み
        Sheets(wsSource).Activate
        VoiceName = Cells(sr, sc).Value
        ALG = Cells(sr, sc + 1).Value
        FB = Cells(sr, sc + 2).Value

        OP1_AR = Cells(sr, sc + 3).Value
        OP1_D1R = Cells(sr, sc + 4).Value
        OP1_D1L = Cells(sr, sc + 5).Value
        OP1_D2R = Cells(sr, sc + 6).Value
        OP1_RR = Cells(sr, sc + 7).Value
        OP1_OL = Cells(sr, sc + 8).Value
        OP1_KS = Cells(sr, sc + 9).Value
        OP1_FR = Cells(sr, sc + 10).Value
        OP1_DT = Cells(sr, sc + 11).Value
        OP1_AMS = Cells(sr, sc + 12).Value
        OP1_SN = Cells(sr, sc + 13).Value

        OP2_AR = Cells(sr, sc + 14).Value
        OP2_D1R = Cells(sr, sc + 15).Value
        OP2_D1L = Cells(sr, sc + 16).Value
        OP2_D2R = Cells(sr, sc + 17).Value
        OP2_RR = Cells(sr, sc + 18).Value
        OP2_OL = Cells(sr, sc + 19).Value
        OP2_KS = Cells(sr, sc + 20).Value
        OP2_FR = Cells(sr, sc + 21).Value
        OP2_DT = Cells(sr, sc + 22).Value
        OP2_AMS = Cells(sr, sc + 23).Value
        OP2_SN = Cells(sr, sc + 24).Value

        OP3_AR = Cells(sr, sc + 25).Value
        OP3_D1R = Cells(sr, sc + 26).Value
        OP3_D1L = Cells(sr, sc + 27).Value
        OP3_D2R = Cells(sr, sc + 28).Value
        OP3_RR = Cells(sr, sc + 29).Value
        OP3_OL = Cells(sr, sc + 30).Value
        OP3_KS = Cells(sr, sc + 31).Value
        OP3_FR = Cells(sr, sc + 32).Value
        OP3_DT = Cells(sr, sc + 33).Value
        OP3_AMS = Cells(sr, sc + 34).Value
        OP3_SN = Cells(sr, sc + 35).Value
        
        OP4_AR = Cells(sr, sc + 36).Value
        OP4_D1R = Cells(sr, sc + 37).Value
        OP4_D1L = Cells(sr, sc + 38).Value
        OP4_D2R = Cells(sr, sc + 39).Value
        OP4_RR = Cells(sr, sc + 40).Value
        OP4_OL = Cells(sr, sc + 41).Value
        OP4_KS = Cells(sr, sc + 42).Value
        OP4_FR = Cells(sr, sc + 43).Value
        OP4_DT = Cells(sr, sc + 44).Value
        OP4_AMS = Cells(sr, sc + 45).Value
        OP4_SN = Cells(sr, sc + 46).Value
        
        OP1_SL = Cells(sr, sc + 47).Value
        OP1_TL = Cells(sr, sc + 48).Value
        OP1_ML = Cells(sr, sc + 49).Value
        OP1_ODT = Cells(sr, sc + 50).Value
        OP2_SL = Cells(sr, sc + 51).Value
        OP2_TL = Cells(sr, sc + 52).Value
        OP2_ML = Cells(sr, sc + 53).Value
        OP2_ODT = Cells(sr, sc + 54).Value
        OP3_SL = Cells(sr, sc + 55).Value
        OP3_TL = Cells(sr, sc + 56).Value
        OP3_ML = Cells(sr, sc + 57).Value
        OP3_ODT = Cells(sr, sc + 58).Value
        OP4_SL = Cells(sr, sc + 59).Value
        OP4_TL = Cells(sr, sc + 60).Value
        OP4_ML = Cells(sr, sc + 61).Value
        OP4_ODT = Cells(sr, sc + 62).Value
        
        OP1_FR = Conv_Freq_Ratio(OP1_FR)
        OP2_FR = Conv_Freq_Ratio(OP2_FR)
        OP3_FR = Conv_Freq_Ratio(OP3_FR)
        OP4_FR = Conv_Freq_Ratio(OP4_FR)
        
        OP1_KL = 0
        OP2_KL = 0
        OP3_KL = 0
        OP4_KL = 0
        
        OP1_EB = 0
        OP2_EB = 0
        OP3_EB = 0
        OP4_EB = 0
        
        LFO_Speed = 0
        LFO_Delay = 0
        PMD = 0
        AMD = 0
        LFO_Sync = 0
        LFO_Wave = 2
        PMS = 0
        AMS = 0
        TRS = 24
        POLY_MONO = 0
        PBR = 4
        P_Mode = 0
        P_Time = 0
        FV = 40
        Sus_S = 1
        P_Switch = 1
        Chorus = 1
        MWP_Range = 50
        MWA_Range = 0
        BPM_Range = 0
        BAM_Range = 0
        BPB_Range = 50
        BEB_Range = 0
                
        PR1 = 99
        PR2 = 99
        PR3 = 99
        PL1 = 50
        PL2 = 50
        PL3 = 50
        
        '1Voice用データ生成（1Voiceデータ93byteのうち、73byte目まで。VoiceName、PR、PLは共通）
        If DataSet = 1 Then

            data_str = data_str & " " & HEX2(OP4_AR)
            data_str = data_str & " " & HEX2(OP4_D1R)
            data_str = data_str & " " & HEX2(OP4_D2R)
            data_str = data_str & " " & HEX2(OP4_RR)
            data_str = data_str & " " & HEX2(OP4_D1L)
            data_str = data_str & " " & HEX2(OP4_KL)
            data_str = data_str & " " & HEX2(OP4_KS)
            data_str = data_str & " " & HEX2(OP4_EB)
            data_str = data_str & " " & HEX2(OP4_AMS)
            data_str = data_str & " " & HEX2(OP4_SN)
            data_str = data_str & " " & HEX2(OP4_OL)
            data_str = data_str & " " & HEX2(OP4_FR)
            data_str = data_str & " " & HEX2(OP4_DT + 3)

            data_str = data_str & " " & HEX2(OP2_AR)
            data_str = data_str & " " & HEX2(OP2_D1R)
            data_str = data_str & " " & HEX2(OP2_D2R)
            data_str = data_str & " " & HEX2(OP2_RR)
            data_str = data_str & " " & HEX2(OP2_D1L)
            data_str = data_str & " " & HEX2(OP2_KL)
            data_str = data_str & " " & HEX2(OP2_KS)
            data_str = data_str & " " & HEX2(OP2_EB)
            data_str = data_str & " " & HEX2(OP2_AMS)
            data_str = data_str & " " & HEX2(OP2_SN)
            data_str = data_str & " " & HEX2(OP2_OL)
            data_str = data_str & " " & HEX2(OP2_FR)
            data_str = data_str & " " & HEX2(OP2_DT + 3)

            data_str = data_str & " " & HEX2(OP3_AR)
            data_str = data_str & " " & HEX2(OP3_D1R)
            data_str = data_str & " " & HEX2(OP3_D2R)
            data_str = data_str & " " & HEX2(OP3_RR)
            data_str = data_str & " " & HEX2(OP3_D1L)
            data_str = data_str & " " & HEX2(OP3_KL)
            data_str = data_str & " " & HEX2(OP3_KS)
            data_str = data_str & " " & HEX2(OP3_EB)
            data_str = data_str & " " & HEX2(OP3_AMS)
            data_str = data_str & " " & HEX2(OP3_SN)
            data_str = data_str & " " & HEX2(OP3_OL)
            data_str = data_str & " " & HEX2(OP3_FR)
            data_str = data_str & " " & HEX2(OP3_DT + 3)

            data_str = data_str & " " & HEX2(OP1_AR)
            data_str = data_str & " " & HEX2(OP1_D1R)
            data_str = data_str & " " & HEX2(OP1_D2R)
            data_str = data_str & " " & HEX2(OP1_RR)
            data_str = data_str & " " & HEX2(OP1_D1L)
            data_str = data_str & " " & HEX2(OP1_KL)
            data_str = data_str & " " & HEX2(OP1_KS)
            data_str = data_str & " " & HEX2(OP1_EB)
            data_str = data_str & " " & HEX2(OP1_AMS)
            data_str = data_str & " " & HEX2(OP1_SN)
            data_str = data_str & " " & HEX2(OP1_OL)
            data_str = data_str & " " & HEX2(OP1_FR)
            data_str = data_str & " " & HEX2(OP1_DT + 3)

            data_str = data_str & " " & HEX2(ALG - 1)
            data_str = data_str & " " & HEX2(FB)
            data_str = data_str & " " & HEX2(LFO_Speed)
            data_str = data_str & " " & HEX2(LFO_Delay)
            data_str = data_str & " " & HEX2(PMD)
            data_str = data_str & " " & HEX2(AMD)
            data_str = data_str & " " & HEX2(LFO_Sync)
            data_str = data_str & " " & HEX2(LFO_Wave)
            data_str = data_str & " " & HEX2(PMS)
            data_str = data_str & " " & HEX2(AMS)
            data_str = data_str & " " & HEX2(TRS)
            data_str = data_str & " " & HEX2(POLY_MONO)
            data_str = data_str & " " & HEX2(PBR)
            data_str = data_str & " " & HEX2(P_Mode)
            data_str = data_str & " " & HEX2(P_Time)
            data_str = data_str & " " & HEX2(FV)
            data_str = data_str & " " & HEX2(Sus_S)
            data_str = data_str & " " & HEX2(P_Switch)
            data_str = data_str & " " & HEX2(Chorus)
            data_str = data_str & " " & HEX2(MWP_Range)
            data_str = data_str & " " & HEX2(MWA_Range)
            data_str = data_str & " " & HEX2(BPM_Range)
            data_str = data_str & " " & HEX2(BAM_Range)
            data_str = data_str & " " & HEX2(BPB_Range)
            data_str = data_str & " " & HEX2(BEB_Range)
        
        '32Voice用データ生成（128byteのうち57Byte目まで。VoiceName、PR、PLは共通。）
        ElseIf DataSet = 32 Then
            
            data_str = data_str & " " & HEX2(OP4_AR)
            data_str = data_str & " " & HEX2(OP4_D1R)
            data_str = data_str & " " & HEX2(OP4_D2R)
            data_str = data_str & " " & HEX2(OP4_RR)
            data_str = data_str & " " & HEX2(OP4_D1L)
            data_str = data_str & " " & HEX2(OP4_KL)
            
            ' Amplitude Modulation Enable/EG Bias Sensitivity/Key Volocity
            data_str = data_str & " " & HEX2(BinToDec(DecToBin(OP4_AMS, 1) & DecToBin(OP4_EB, 3) & DecToBin(OP4_SN, 3)))
            
            data_str = data_str & " " & HEX2(OP4_OL)
            data_str = data_str & " " & HEX2(OP4_FR)
                       
            'Keyboad Scaling Rate/Detune1
            data_str = data_str & " " & HEX2(BinToDec(DecToBin(OP4_KS, 2) & DecToBin((OP4_DT + 3), 3)))
                        
            data_str = data_str & " " & HEX2(OP2_AR)
            data_str = data_str & " " & HEX2(OP2_D1R)
            data_str = data_str & " " & HEX2(OP2_D2R)
            data_str = data_str & " " & HEX2(OP2_RR)
            data_str = data_str & " " & HEX2(OP2_D1L)
            data_str = data_str & " " & HEX2(OP2_KL)
            
            ' Amplitude Modulation Enable/EG Bias Sensitivity/Key Volocity
            data_str = data_str & " " & HEX2(BinToDec(DecToBin(OP2_AMS, 1) & DecToBin(OP2_EB, 3) & DecToBin(OP2_SN, 3)))
            
            data_str = data_str & " " & HEX2(OP2_OL)
            data_str = data_str & " " & HEX2(OP2_FR)
                       
            'Keyboad Scaling Rate/Detune1
            data_str = data_str & " " & HEX2(BinToDec(DecToBin(OP2_KS, 2) & DecToBin((OP2_DT + 3), 3)))
            
            data_str = data_str & " " & HEX2(OP3_AR)
            data_str = data_str & " " & HEX2(OP3_D1R)
            data_str = data_str & " " & HEX2(OP3_D2R)
            data_str = data_str & " " & HEX2(OP3_RR)
            data_str = data_str & " " & HEX2(OP3_D1L)
            data_str = data_str & " " & HEX2(OP3_KL)
            
            ' Amplitude Modulation Enable/EG Bias Sensitivity/Key Volocity
            data_str = data_str & " " & HEX2(BinToDec(DecToBin(OP3_AMS, 1) & DecToBin(OP3_EB, 3) & DecToBin(OP3_SN, 3)))
            
            data_str = data_str & " " & HEX2(OP3_OL)
            data_str = data_str & " " & HEX2(OP3_FR)
                       
            'Keyboad Scaling Rate/Detune1
            data_str = data_str & " " & HEX2(BinToDec(DecToBin(OP3_KS, 2) & DecToBin((OP3_DT + 3), 3)))
            
            data_str = data_str & " " & HEX2(OP1_AR)
            data_str = data_str & " " & HEX2(OP1_D1R)
            data_str = data_str & " " & HEX2(OP1_D2R)
            data_str = data_str & " " & HEX2(OP1_RR)
            data_str = data_str & " " & HEX2(OP1_D1L)
            data_str = data_str & " " & HEX2(OP1_KL)
            
            ' Amplitude Modulation Enable/EG Bias Sensitivity/Key Volocity
            data_str = data_str & " " & HEX2(BinToDec(DecToBin(OP1_AMS, 1) & DecToBin(OP1_EB, 3) & DecToBin(OP1_SN, 3)))
            
            data_str = data_str & " " & HEX2(OP1_OL)
            data_str = data_str & " " & HEX2(OP1_FR)
                       
            'Keyboad Scaling Rate/Detune1
            data_str = data_str & " " & HEX2(BinToDec(DecToBin(OP1_KS, 2) & DecToBin((OP1_DT + 3), 3)))
            
            'LFO Sync/FeedBack/Algorithm
            data_str = data_str & " " & HEX2(BinToDec(DecToBin(LFO_Sync, 1) & DecToBin(FB, 3) & DecToBin((ALG - 1), 3)))

            data_str = data_str & " " & HEX2(LFO_Speed)
            data_str = data_str & " " & HEX2(LFO_Delay)
            data_str = data_str & " " & HEX2(PMD)
            data_str = data_str & " " & HEX2(AMD)

            'PMS/AMS/LFO Wave
            data_str = data_str & " " & HEX2(BinToDec(DecToBin(PMS, 3) & DecToBin(AMS, 3) & DecToBin(LFO_Wave, 2)))
            
            data_str = data_str & " " & HEX2(TRS)
            data_str = data_str & " " & HEX2(PBR)
            
            'Chorus Switch/Play Mode/Sustain Foot Switch/Portament Foot Switch/Portament Mode
            data_str = data_str & " " & HEX2(BinToDec(DecToBin(Chorus, 1) & DecToBin(POLY_MONO, 1) & DecToBin(Sus_S, 1) & DecToBin(P_Switch, 1) & DecToBin(P_Mode, 1)))
            
            data_str = data_str & " " & HEX2(P_Time)
            data_str = data_str & " " & HEX2(FV)
            data_str = data_str & " " & HEX2(MWP_Range)
            data_str = data_str & " " & HEX2(MWA_Range)
            data_str = data_str & " " & HEX2(BPM_Range)
            data_str = data_str & " " & HEX2(BAM_Range)
            data_str = data_str & " " & HEX2(BPB_Range)
            data_str = data_str & " " & HEX2(BEB_Range)
            
        End If
        
        '1Voice、32Voice共通部分
        'VoiceNameの文字コードを16進数に変換しdata_strに付加（10バイト）
        data_str = data_str & " " & Conv_Text(Left(VoiceName, 10))
                
        'PR、PLデータ付加
        data_str = data_str & " " & HEX2(PR1)
        data_str = data_str & " " & HEX2(PR2)
        data_str = data_str & " " & HEX2(PR3)
        data_str = data_str & " " & HEX2(PL1)
        data_str = data_str & " " & HEX2(PL2)
        data_str = data_str & " " & HEX2(PL3)
        
        '32Voiceデータの場合、55バイトの0データを付加（32Voiceの1音色のデータは73Byte目まで）
        If DataSet = 32 Then
            data_str = data_str & " " & "00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00"
        End If
        
        'カウンターの更新
        sr = sr + 1
        tr = tr + 1
        
    Next c

    'チェックサムの計算（data_strをBiyte配列に変換して合計する）
    chksum_data = Split(data_str, " ")
    ReDim chksum_byte(UBound(chksum_data))
    
    For k = 6 To UBound(chksum_data)
        chksum_byte(k - 6) = "&h" & chksum_data(k)
    Next k
        
    For j = 0 To UBound(chksum_byte)
        chksum = chksum + Hex("&H" & chksum_byte(j))
    Next j
    
    '128で割った余りを128から引く
    chksum = 128 - (chksum Mod 128)
    
    '計算したチェックサムをdata_strの末尾に付加
    data_str = data_str & " " & HEX2(chksum)
    
    'Sysexデータの最終コードをdata_strの末尾に付加
    data_str = data_str & " " & "F7"
    
    'data_strのデータを配列データへ変換
    VoiceData_str = Split(data_str, " ")
    
    '文字列配列データをByte配列データへ変換
    ReDim VoiceData_byt(UBound(VoiceData_str))

    For i = 0 To UBound(VoiceData_str)
        VoiceData_byt(i) = "&h" & VoiceData_str(i)
    Next i
    
    'バイナリーファイルに書き込み
    fh = FreeFile
    Open filename For Binary Access Write As #fh
    Put #fh, , VoiceData_byt
    Close #fh
       
End Sub

