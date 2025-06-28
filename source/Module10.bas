Attribute VB_Name = "Module10"
'********************************************
' OPMデータの書出し
'********************************************
Sub Conv_OPM()

    Dim strFilePath As String
    Dim strFileName As String
    Dim strFile As String
    Dim Msg, Style, Title, Response
    
    Style = vbOKCancel
    Title = "エラー"
    
    '//ファイル名を生成
    
    Sheets("Menu").Select
    strFilePath = Cells(65, 5).Value
    strFileName = Cells(66, 5).Value

    If strFilePath = "" Then
        strFilePath = ThisWorkbook.Path
    End If
    
    If strFileName = "" Then
        MsgBox "ファイル名が指定されていません。"
        
    Else
        strFile = strFilePath & "\" & strFileName
                
        If Dir(strFile) = "" Then
        
            Call Make_OPM_Data(strFile)
            Sheets("Menu").Select
            MsgBox "OPMデータの書き出しが完了しました。"
            
        Else
        
            Msg = strFile & vbCrLf & _
                    "がすでに存在します。上書きしてもよろしいですか？"

            Response = MsgBox(Msg, Style, Title)
            
            If Response = vbOK Then
            
                Kill strFile
                Call Make_OPM_Data(strFile)
                Sheets("Menu").Select
                MsgBox "OPMデータの書き出しが完了しました。"
                
            ElseIf Response = vbCancel Then
            
                MsgBox "OPMデータの書き出しをキャンセルしました。"
            
            Else
                
                MsgBox "Sysexデータの書き出しを中止しました。"
            
            End If
        End If
    End If
    
End Sub

'*******************************
' OPMファイルへデータを書き込み
'*******************************
Sub Make_OPM_Data(filename As String)

    Dim Header1 As String
    Dim Header2 As String
    Dim Header3 As String
    Dim Header4 As String
    Dim Header5 As String
    Dim RowNum As String
    Dim RowLFO As String
    Dim RowCH As String
    Dim RowM1 As String
    Dim RowC1 As String
    Dim RowM2 As String
    Dim RowC2 As String
    
    Dim ff As Long

    Dim wsSource As String
    Dim VoiceName As String
    Dim VoiceNumber As String
    Dim ALG As Long, FB As Long
    Dim OP1_AR, OP1_D1R, OP1_D1L, OP1_D2R, OP1_RR, OP1_OL, OP1_KS As Long, OP1_FR, OP1_DT As Long, OP1_AMS As Long, OP1_SN As Long, OP1_SL, OP1_TL, OP1_ML, OP1_ODT, OP1_KL, OP1_EB As Long
    Dim OP2_AR, OP2_D1R, OP2_D1L, OP2_D2R, OP2_RR, OP2_OL, OP2_KS As Long, OP2_FR, OP2_DT As Long, OP2_AMS As Long, OP2_SN As Long, OP2_SL, OP2_TL, OP2_ML, OP2_ODT, OP2_KL, OP2_EB As Long
    Dim OP3_AR, OP3_D1R, OP3_D1L, OP3_D3R, OP3_RR, OP3_OL, OP3_KS As Long, OP3_FR, OP3_DT As Long, OP3_AMS As Long, OP3_SN As Long, OP3_SL, OP3_TL, OP3_ML, OP3_ODT, OP3_KL, OP3_EB As Long
    Dim OP4_AR, OP4_D1R, OP4_D1L, OP4_D2R, OP4_RR, OP4_OL, OP4_KS As Long, OP4_FR, OP4_DT As Long, OP4_AMS As Long, OP4_SN As Long, OP4_SL, OP4_TL, OP4_ML, OP4_ODT, OP4_KL, OP4_EB As Long
    Dim LFO_Speed, LFO_Delay, PMD, AMD, LFO_Sync As Long, LFO_Wave As Long, PMS As Long, AMS As Long, TRS, POLY_MONO As Long, PBR, P_Mode As Long, P_Time, FV, Sus_S As Long, P_Switch As Long, Chorus As Long, MWP_Range As Long, MWA_Range As Long
    Dim BPM_Range, BAM_Range, BPB_Range, BEB_Range, PR1, PR2, PR3, PL1, PL2, PL3 As Long
    
    Dim Data_set, c
    Dim sr As Long, sc As Long

    Header1 = "//MiOPMdrv sound bank Paramer Ver2002.04.22"
    Header2 = "//@:[Num] [Name]"
    Header3 = "//LFO: LFRQ AMD PMD WF NFRQ"
    Header4 = "//CH: PAN  FL CON AMS PMS SLOT NE"
    Header5 = "//[OPname]: AR D1R D2R  RR D1L  TL  KS MUL DT1 DT2 AMS-EN"
    
    '元データシート
    wsSource = "DXtoOPM_Output"
    
    'データ数カウント
    Sheets(wsSource).Activate
    DataSet = getLastRow(ActiveSheet, 2) - 1
    If DataSet > 128 Then
        DataSet = 128
    End If
    
    'ソースデータシート読み込み用カウンターの初期化
    sr = 2
    sc = 2
    
    ff = FileSystem.FreeFile()
    Open filename For Output As #ff
    
    Print #ff, Header1
    Print #ff, Header2
    Print #ff, Header3
    Print #ff, Header4
    Print #ff, Header5
    
    For c = 1 To DataSet
    
    'OutpuDataシートの値を読み込み
        Sheets(wsSource).Activate
        VoiceName = Cells(sr, sc).Value
        VoiceNumber = Cells(sr, sc - 1).Value
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

        OP2_AR = Cells(sr, sc + 16).Value
        OP2_D1R = Cells(sr, sc + 17).Value
        OP2_D1L = Cells(sr, sc + 18).Value
        OP2_D2R = Cells(sr, sc + 19).Value
        OP2_RR = Cells(sr, sc + 20).Value
        OP2_OL = Cells(sr, sc + 21).Value
        OP2_KS = Cells(sr, sc + 22).Value
        OP2_FR = Cells(sr, sc + 23).Value
        OP2_DT = Cells(sr, sc + 24).Value
        OP2_AMS = Cells(sr, sc + 25).Value
        OP2_SN = Cells(sr, sc + 26).Value

        OP3_AR = Cells(sr, sc + 29).Value
        OP3_D1R = Cells(sr, sc + 30).Value
        OP3_D1L = Cells(sr, sc + 31).Value
        OP3_D2R = Cells(sr, sc + 32).Value
        OP3_RR = Cells(sr, sc + 33).Value
        OP3_OL = Cells(sr, sc + 34).Value
        OP3_KS = Cells(sr, sc + 35).Value
        OP3_FR = Cells(sr, sc + 36).Value
        OP3_DT = Cells(sr, sc + 37).Value
        OP3_AMS = Cells(sr, sc + 38).Value
        OP3_SN = Cells(sr, sc + 39).Value
        
        OP4_AR = Cells(sr, sc + 42).Value
        OP4_D1R = Cells(sr, sc + 43).Value
        OP4_D1L = Cells(sr, sc + 44).Value
        OP4_D2R = Cells(sr, sc + 45).Value
        OP4_RR = Cells(sr, sc + 46).Value
        OP4_OL = Cells(sr, sc + 47).Value
        OP4_KS = Cells(sr, sc + 48).Value
        OP4_FR = Cells(sr, sc + 49).Value
        OP4_DT = Cells(sr, sc + 50).Value
        OP4_AMS = Cells(sr, sc + 51).Value
        OP4_SN = Cells(sr, sc + 52).Value
        
        LFO_Speed = Cells(sr, sc + 55).Value
        LFO_Delay = Cells(sr, sc + 56).Value
        PMD = Cells(sr, sc + 57).Value
        AMD = Cells(sr, sc + 58).Value
        LFO_Sync = Cells(sr, sc + 59).Value
        LFO_Wave = Cells(sr, sc + 60).Value
        PMS = Cells(sr, sc + 61).Value
        AMS = Cells(sr, sc + 62).Value
        
        RowNum = "@:" & VoiceNumber & " " & VoiceName
        RowLFO = "LFO:" & StringLen(ConvSP_DX21toOPM(LFO_Speed), 3) & StringLen(ConvLFOs_DX21toOPM(AMD), 4) _
            & StringLen(ConvLFOs_DX21toOPM(PMD), 4) & StringLen(LFO_Wave, 4) & "   0"
        RowCH = "CH: 64" & StringLen(FB, 4) & StringLen(ConvALG_DX21toOPM(ALG), 4) & StringLen(AMS, 4) & StringLen(PMS, 4) & " 120   0"
        RowM1 = "M1:" & StringLen(OP4_AR, 3) & StringLen(OP4_D1R, 4) & StringLen(OP4_D2R, 4) & StringLen(OP4_RR, 4) _
            & StringLen(ConvD1L_DX21toOPM(OP4_D1L), 4) & StringLen(ConvOL_DX21toOPM(OP4_OL), 4) & StringLen(OP4_KS, 4) _
            & StringLen(ConvFR_DX21toOPM(OP4_FR), 4) & StringLen(ConvDT_DX21toOPM(OP4_DT), 4) & "   0" & StringLen(OP4_AMS, 4)
        RowC1 = "C1:" & StringLen(OP3_AR, 3) & StringLen(OP3_D1R, 4) & StringLen(OP3_D2R, 4) & StringLen(OP3_RR, 4) _
            & StringLen(ConvD1L_DX21toOPM(OP3_D1L), 4) & StringLen(ConvOL_DX21toOPM(OP3_OL), 4) & StringLen(OP3_KS, 4) _
            & StringLen(ConvFR_DX21toOPM(OP3_FR), 4) & StringLen(ConvDT_DX21toOPM(OP3_DT), 4) & "   0" & StringLen(OP3_AMS, 4)
        RowM2 = "M2:" & StringLen(OP2_AR, 3) & StringLen(OP2_D1R, 4) & StringLen(OP2_D2R, 4) & StringLen(OP2_RR, 4) _
            & StringLen(ConvD1L_DX21toOPM(OP2_D1L), 4) & StringLen(ConvOL_DX21toOPM(OP2_OL), 4) & StringLen(OP2_KS, 4) _
            & StringLen(ConvFR_DX21toOPM(OP2_FR), 4) & StringLen(ConvDT_DX21toOPM(OP2_DT), 4) & "   0" & StringLen(OP2_AMS, 4)
        RowC2 = "C2:" & StringLen(OP1_AR, 3) & StringLen(OP1_D1R, 4) & StringLen(OP1_D2R, 4) & StringLen(OP1_RR, 4) _
            & StringLen(ConvD1L_DX21toOPM(OP1_D1L), 4) & StringLen(ConvOL_DX21toOPM(OP1_OL), 4) & StringLen(OP1_KS, 4) _
            & StringLen(ConvFR_DX21toOPM(OP1_FR), 4) & StringLen(ConvDT_DX21toOPM(OP1_DT), 4) & "   0" & StringLen(OP1_AMS, 4)
        
       Print #ff,
       Print #ff, RowNum
       Print #ff, RowLFO
       Print #ff, RowCH
       Print #ff, RowM1
       Print #ff, RowC1
       Print #ff, RowM2
       Print #ff, RowC2
       
       'カウンターの更新
        sr = sr + 1

    Next c

    Close #ff

End Sub


'********************************************
' OPMファイル読み込み メインプログラム
'********************************************
Sub OPM_Read_Main()
    Dim strFilePath As String
    Dim strFileName As String
    Dim strFile As String
    '//ファイル名を生成
    
    Sheets("Menu").Select
    strFilePath = Cells(57, 5).Value
    strFileName = Cells(58, 5).Value

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
' OPMファイルからデータ読み込み
'********************************************
Sub ReadSysexFile(ByVal strfil As String)
    Dim buff() As Byte
    Dim fp As Long
    Dim filesize As Long, NowLoc As Long
    Dim idx As Long, gyo As Long
    Dim strBinary As String
    Dim wsTALGet As String
    Dim fp As Long
    
    Dim VoiceName As String
    Dim VoiceNumber As String
    Dim CON As Long, FL As Long
    Dim OP1_AR, OP1_D1R, OP1_D1L, OP1_D2R, OP1_RR, OP1_TL, OP1_KS As Long, OP1_MUL, OP1_DT As Long, OP1_AMS As Long, OP1_DT2 As Long
    Dim OP2_AR, OP2_D1R, OP2_D1L, OP2_D2R, OP2_RR, OP2_TL, OP2_KS As Long, OP2_MUL, OP2_DT As Long, OP2_AMS As Long, OP2_DT2 As Long
    Dim OP3_AR, OP3_D1R, OP3_D1L, OP3_D3R, OP3_RR, OP3_TL, OP3_KS As Long, OP3_MUL, OP3_DT As Long, OP3_AMS As Long, OP3_DT2 As Long
    Dim OP4_AR, OP4_D1R, OP4_D1L, OP4_D2R, OP4_RR, OP4_TL, OP4_KS As Long, OP4_MUL, OP4_DT As Long, OP4_AMS As Long, OP4_DT2 As Long
    Dim LFRQ, PMD, AMD, LFO_Wave As Long, PMS As Long, AMS As Long
    Dim PAN, SLOT, NE
    
    Dim TextLine
    Dim Data_set, c
    Dim sr As Long, tr As Long, tc As Long
    'Dim chksum2 As Long
    
    
    'データ出力用シート
    wsTALGet = "OPM_DataBase"
    
    Sheets(wsTALGet).Activate
     
    'ActiveSheet.Unprotect
    'Range("A2:CG33", "CI2:DB33").ClearContents
    'ActiveSheet.Protect
    
    '//FreeFile関数で使用可能なファイル番号を割り当て
    fp = FileSystem.FreeFile
    '//ファイルを開く
    Open strfil For Input As #fp ' Open file.
    Do While Not EOF(1) ' Loop until end of file.
        Line Input #fp, TextLine
        Debug.Print TextLine ' Print to the Immediate window.
        
        
        
        
        
    Loop
    Close #fp ' Close file.
   
    '読み込みカウンター初期化（ヘッダーの読み飛ばし7バイト目から）
    sr = 6
    
    'シート書き込み用カウンター初期化
    tr = 2
    tc = 1
   
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
            
            'YAMAHA V2(DX11)のオシレーターデータの取得

            OP4_FIXR = Hex("&H" & buff(sr + 94))
            OP4_FIXRG = Hex("&H" & buff(sr + 95))
            OP4_FINE = Hex("&H" & buff(sr + 96))
            OP4_OSW Hex("&H" & buff(sr + 97))
            OP4_EGSFT = Hex("&H" & buff(sr + 98))

            OP2_FIXR = Hex("&H" & buff(sr + 99))
            OP2_FIXRG = Hex("&H" & buff(sr + 100))
            OP2_FINE = Hex("&H" & buff(sr + 101))
            OP2_OSW Hex("&H" & buff(sr + 102))
            OP2_EGSFT = Hex("&H" & buff(sr + 103))

            OP3_FIXR = Hex("&H" & buff(sr + 104))
            OP3_FIXRG = Hex("&H" & buff(sr + 105))
            OP3_FINE = Hex("&H" & buff(sr + 106))
            OP3_OSW Hex("&H" & buff(sr + 107))
            OP3_EGSFT = Hex("&H" & buff(sr + 108))
            
            OP4_FIXR = Hex("&H" & buff(sr + 109))
            OP4_FIXRG = Hex("&H" & buff(sr + 110))
            OP4_FINE = Hex("&H" & buff(sr + 111))
            OP4_OSW Hex("&H" & buff(sr + 112))
            OP4_EGSFT = Hex("&H" & buff(sr + 113))
            
            'YAMAHA V2、V50のフットコントロール、アフタータッチ
            REV = Hex("&H" & buff(sr + 114))
            FC_Pitch = Hex("&H" & buff(sr + 115))
            FC_AMP = Hex("&H" & buff(sr + 116))
            AT_Pitch = Hex("&H" & buff(sr + 117))
            AT_AMP = Hex("&H" & buff(sr + 118))
            AT_PBias = Hex("&H" & buff(sr + 119))
            AT_EGBias = Hex("&H" & buff(sr + 120))
            
            OP4_FIX_RM = Hex("&H" & buff(sr + 121))
            OP2_FIX_RM = Hex("&H" & buff(sr + 122))
            OP3_FIX_RM = Hex("&H" & buff(sr + 123))
            OP1_FIX_RM = Hex("&H" & buff(sr + 124))
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
            OP4_FIXR = BinToDec(Mid(DecToBin(Hex("&H" & buff(sr + 73)), 6), 3, 1))
            OP4_FIXRG = BinToDec(Right(DecToBin(Hex("&H" & buff(sr + 73)), 6), 3))
            OP4_OSW = BinToDec(Left(DecToBin(Hex("&H" & buff(sr + 74)), 7), 3))
            OP4_FINE = BinToDec(Right(DecToBin(Hex("&H" & buff(sr + 74)), 7), 4))
            
            OP2_EGSFT = BinToDec(Left(DecToBin(Hex("&H" & buff(sr + 75)), 6), 2))
            OP2_FIXR = BinToDec(Mid(DecToBin(Hex("&H" & buff(sr + 75)), 6), 3, 1))
            OP2_FIXRG = BinToDec(Right(DecToBin(Hex("&H" & buff(sr + 75)), 6), 3))
            OP2_OSW = BinToDec(Left(DecToBin(Hex("&H" & buff(sr + 76)), 7), 3))
            OP2_FINE = BinToDec(Right(DecToBin(Hex("&H" & buff(sr + 76)), 7), 4))

            OP3_EGSFT = BinToDec(Left(DecToBin(Hex("&H" & buff(sr + 77)), 6), 2))
            OP3_FIXR = BinToDec(Mid(DecToBin(Hex("&H" & buff(sr + 77)), 6), 3, 1))
            OP3_FIXRG = BinToDec(Right(DecToBin(Hex("&H" & buff(sr + 77)), 6), 3))
            OP3_OSW = BinToDec(Left(DecToBin(Hex("&H" & buff(sr + 78)), 7), 3))
            OP3_FINE = BinToDec(Right(DecToBin(Hex("&H" & buff(sr + 78)), 7), 4))
            
            OP1_EGSFT = BinToDec(Left(DecToBin(Hex("&H" & buff(sr + 79)), 6), 2))
            OP1_FIXR = BinToDec(Mid(DecToBin(Hex("&H" & buff(sr + 79)), 6), 3, 1))
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
        Cells(tr, tc + 86).Value = OP1_FIXR
        Cells(tr, tc + 87).Value = OP1_FIXRG
        Cells(tr, tc + 88).Value = OP1_OSW
        Cells(tr, tc + 89).Value = OP1_FINE
        
        Cells(tr, tc + 90).Value = OP2_EGSFT
        Cells(tr, tc + 91).Value = OP2_FIXR
        Cells(tr, tc + 92).Value = OP2_FIXRG
        Cells(tr, tc + 93).Value = OP2_OSW
        Cells(tr, tc + 94).Value = OP2_FINE

        Cells(tr, tc + 95).Value = OP3_EGSFT
        Cells(tr, tc + 96).Value = OP3_FIXR
        Cells(tr, tc + 97).Value = OP3_FIXRG
        Cells(tr, tc + 98).Value = OP3_OSW
        Cells(tr, tc + 99).Value = OP3_FINE

        Cells(tr, tc + 100).Value = OP4_EGSFT
        Cells(tr, tc + 101).Value = OP4_FIXR
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
