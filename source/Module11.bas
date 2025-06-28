Attribute VB_Name = "Module11"

'********************************************
' SY77用Sysexファイル読み込み メインプログラム
'********************************************
Sub BinaryMainSY77()
    Dim strFilePath As String
    Dim strFileName As String
    Dim strFile As String
    Dim objFileSys As Object
    Dim strLibName As String
    '//ファイル名を生成
    
    Sheets("MenuSY77").Select
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
  
        'ファイルシステムを扱うオブジェクトを作成
            Set objFileSys = CreateObject("Scripting.FileSystemObject")
            
        '//バイナリファイルを読み込み（Sysexファイルを読み込み）
            strLibName = objFileSys.GetBaseName(strFile)
            'Call ReadSysexFileSY77(strFile, strLibName)
            Call ReadSYBinaryFile(strFile)
            Sheets("SysexSY77Data").Select
            MsgBox "Sysexデータの読み込みが完了しました。"
            Set objFileSys = Nothing
        Else
            MsgBox strFile & vbCrLf & _
                    "が存在しません"
        End If
    End If

End Sub

'********************************************
' SY77用Sysexファイルからバイナリデータ読み込み
'********************************************
Sub ReadSysexFileSY77(ByVal strfil As String, Optional strLibName As String)
    Dim buff() As Byte
    Dim fp As Long
    Dim filesize As Long, NowLoc As Long
    Dim idx As Long, gyo As Long
    Dim strBinary As String
    Dim wsTALGet As String
    
    Dim VoiceName As String
    Dim ELMode, WPBR, ATPBR, PMASN, PMRNG, AMASN, AMRNG, FMASN, FMRNG, PNLASN, PNLRNG, COASN, CORNG
    Dim PNBASN, PNBRNG, EGBASN, EGBRNG, VVLASN, VVLLML, MCTUN, RNDP, PORM, POS, VVOL
    Dim ELVL, ELDT, ELNS, ENLL, ENLH, PANNM, MCTEN, OUTSEL0, OUTSEL1
    Dim ALGNUM, FPR1, FPR2, FPR3, FPRR1, FPL0, FPL1, FPL2, FPL3, FPRL1, FPEGR, FPRS, FVPSR
    Dim FLFSPD
     
    Dim OP1_EGR1, OP1_EGR2, OP1_EGR3, OP1_EGR4, OP1_EGL1, OP1_EGL2, OP1_EGL3, OP1_EGL4, OP1_KLS_BP, OP1_KLS_LD, OP1_KLS_RD, OP1_KLS_LC, OP1_KLS_RC, OP1_KRS, OP1_AMP, OP1_KVS, OP1_OL, OP1_OM, OP1_OSFC, OP1_OSFF, OP1_DT
    Dim OP2_EGR1, OP2_EGR2, OP2_EGR3, OP2_EGR4, OP2_EGL1, OP2_EGL2, OP2_EGL3, OP2_EGL4, OP2_KLS_BP, OP2_KLS_LD, OP2_KLS_RD, OP2_KLS_LC, OP2_KLS_RC, OP2_KRS, OP2_AMP, OP2_KVS, OP2_OL, OP2_OM, OP2_OSFC, OP2_OSFF, OP2_DT
    Dim OP3_EGR1, OP3_EGR2, OP3_EGR3, OP3_EGR4, OP3_EGL1, OP3_EGL2, OP3_EGL3, OP3_EGL4, OP3_KLS_BP, OP3_KLS_LD, OP3_KLS_RD, OP3_KLS_LC, OP3_KLS_RC, OP3_KRS, OP3_AMP, OP3_KVS, OP3_OL, OP3_OM, OP3_OSFC, OP3_OSFF, OP3_DT
    Dim OP4_EGR1, OP4_EGR2, OP4_EGR3, OP4_EGR4, OP4_EGL1, OP4_EGL2, OP4_EGL3, OP4_EGL4, OP4_KLS_BP, OP4_KLS_LD, OP4_KLS_RD, OP4_KLS_LC, OP4_KLS_RC, OP4_KRS, OP4_AMP, OP4_KVS, OP4_OL, OP4_OM, OP4_OSFC, OP4_OSFF, OP4_DT
    Dim OP5_EGR1, OP5_EGR2, OP5_EGR3, OP5_EGR4, OP5_EGL1, OP5_EGL2, OP5_EGL3, OP5_EGL4, OP5_KLS_BP, OP5_KLS_LD, OP5_KLS_RD, OP5_KLS_LC, OP5_KLS_RC, OP5_KRS, OP5_AMP, OP5_KVS, OP5_OL, OP5_OM, OP5_OSFC, OP5_OSFF, OP5_DT
    Dim OP6_EGR1, OP6_EGR2, OP6_EGR3, OP6_EGR4, OP6_EGL1, OP6_EGL2, OP6_EGL3, OP6_EGL4, OP6_KLS_BP, OP6_KLS_LD, OP6_KLS_RD, OP6_KLS_LC, OP6_KLS_RC, OP6_KRS, OP6_AMP, OP6_KVS, OP6_OL, OP6_OM, OP6_OSFC, OP6_OSFF, OP6_DT
    Dim PR1, PR2, PR3, PR4, PL1, PL2, PL3, PL4
    Dim OSC_Sync
    Dim LFO_Speed, LFO_Delay, PMD, AMD, LFO_Sync, LFO_Wave, PMS, TRS
    Dim OPR_S
    Dim POLY_MONO, PBR, PBS, P_Mode, P_Gliss, P_Time, MW_Range, MW_Assing, FC_Range, FC_Assigh, BC_Range, BC_Assign, AT_Range, AT_Assign
        
    Dim Data_set, c
    Dim sr As Long, tr As Long, tc As Long
    'Dim chksum2 As Long
        
    'データ出力用シート
    wsTALGet = "SysexSY77Data"
    
    Sheets(wsTALGet).Activate
    'ActiveSheet.Unprotect
    'Range("A2:ER33").ClearContents
    'ActiveSheet.Protect
        
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
    tc = 3
   
    'Sysexデータの4バイト目を確認し1Voiceデータか32Voiceバルクデータかを判別
    If buff(3) = 0 Then

        Data_set = 1
        
    ElseIf buff(3) = 9 Then

        Data_set = 32
        
    End If

    'chksum2 = 0

    For c = 1 To Data_set
        
        If Data_set = 1 Then
        
        '1Voiceデータの場合
            OP6_EGR1 = Hex("&H" & buff(sr))
            OP6_EGR2 = Hex("&H" & buff(sr + 1))
            OP6_EGR3 = Hex("&H" & buff(sr + 2))
            OP6_EGR4 = Hex("&H" & buff(sr + 3))
            OP6_EGL1 = Hex("&H" & buff(sr + 4))
            OP6_EGL2 = Hex("&H" & buff(sr + 5))
            OP6_EGL3 = Hex("&H" & buff(sr + 6))
            OP6_EGL4 = Hex("&H" & buff(sr + 7))
            OP6_KLS_BP = Hex("&H" & buff(sr + 8))
            OP6_KLS_LD = Hex("&H" & buff(sr + 9))
            OP6_KLS_RD = Hex("&H" & buff(sr + 10))
            OP6_KLS_LC = Hex("&H" & buff(sr + 11))
            OP6_KLS_RC = Hex("&H" & buff(sr + 12))
            OP6_KRS = Hex("&H" & buff(sr + 13))
            OP6_AMP = Hex("&H" & buff(sr + 14))
            OP6_KVS = Hex("&H" & buff(sr + 15))
            OP6_OL = Hex("&H" & buff(sr + 16))
            OP6_OM = Hex("&H" & buff(sr + 17))
            OP6_OSFC = Hex("&H" & buff(sr + 18))
            OP6_OSFF = Hex("&H" & buff(sr + 19))
            OP6_DT = Hex("&H" & buff(sr + 20))
            
            OP5_EGR1 = Hex("&H" & buff(sr + 21))
            OP5_EGR2 = Hex("&H" & buff(sr + 22))
            OP5_EGR3 = Hex("&H" & buff(sr + 23))
            OP5_EGR4 = Hex("&H" & buff(sr + 24))
            OP5_EGL1 = Hex("&H" & buff(sr + 25))
            OP5_EGL2 = Hex("&H" & buff(sr + 26))
            OP5_EGL3 = Hex("&H" & buff(sr + 27))
            OP5_EGL4 = Hex("&H" & buff(sr + 28))
            OP5_KLS_BP = Hex("&H" & buff(sr + 29))
            OP5_KLS_LD = Hex("&H" & buff(sr + 30))
            OP5_KLS_RD = Hex("&H" & buff(sr + 31))
            OP5_KLS_LC = Hex("&H" & buff(sr + 32))
            OP5_KLS_RC = Hex("&H" & buff(sr + 33))
            OP5_KRS = Hex("&H" & buff(sr + 34))
            OP5_AMP = Hex("&H" & buff(sr + 35))
            OP5_KVS = Hex("&H" & buff(sr + 36))
            OP5_OL = Hex("&H" & buff(sr + 37))
            OP5_OM = Hex("&H" & buff(sr + 38))
            OP5_OSFC = Hex("&H" & buff(sr + 39))
            OP5_OSFF = Hex("&H" & buff(sr + 40))
            OP5_DT = Hex("&H" & buff(sr + 41))
            
            OP4_EGR1 = Hex("&H" & buff(sr + 42))
            OP4_EGR2 = Hex("&H" & buff(sr + 43))
            OP4_EGR3 = Hex("&H" & buff(sr + 44))
            OP4_EGR4 = Hex("&H" & buff(sr + 45))
            OP4_EGL1 = Hex("&H" & buff(sr + 46))
            OP4_EGL2 = Hex("&H" & buff(sr + 47))
            OP4_EGL3 = Hex("&H" & buff(sr + 48))
            OP4_EGL4 = Hex("&H" & buff(sr + 49))
            OP4_KLS_BP = Hex("&H" & buff(sr + 50))
            OP4_KLS_LD = Hex("&H" & buff(sr + 51))
            OP4_KLS_RD = Hex("&H" & buff(sr + 52))
            OP4_KLS_LC = Hex("&H" & buff(sr + 53))
            OP4_KLS_RC = Hex("&H" & buff(sr + 54))
            OP4_KRS = Hex("&H" & buff(sr + 55))
            OP4_AMP = Hex("&H" & buff(sr + 56))
            OP4_KVS = Hex("&H" & buff(sr + 57))
            OP4_OL = Hex("&H" & buff(sr + 58))
            OP4_OM = Hex("&H" & buff(sr + 59))
            OP4_OSFC = Hex("&H" & buff(sr + 60))
            OP4_OSFF = Hex("&H" & buff(sr + 61))
            OP4_DT = Hex("&H" & buff(sr + 62))
            
            OP3_EGR1 = Hex("&H" & buff(sr + 63))
            OP3_EGR2 = Hex("&H" & buff(sr + 64))
            OP3_EGR3 = Hex("&H" & buff(sr + 65))
            OP3_EGR4 = Hex("&H" & buff(sr + 66))
            OP3_EGL1 = Hex("&H" & buff(sr + 67))
            OP3_EGL2 = Hex("&H" & buff(sr + 68))
            OP3_EGL3 = Hex("&H" & buff(sr + 69))
            OP3_EGL4 = Hex("&H" & buff(sr + 70))
            OP3_KLS_BP = Hex("&H" & buff(sr + 71))
            OP3_KLS_LD = Hex("&H" & buff(sr + 72))
            OP3_KLS_RD = Hex("&H" & buff(sr + 73))
            OP3_KLS_LC = Hex("&H" & buff(sr + 74))
            OP3_KLS_RC = Hex("&H" & buff(sr + 75))
            OP3_KRS = Hex("&H" & buff(sr + 76))
            OP3_AMP = Hex("&H" & buff(sr + 77))
            OP3_KVS = Hex("&H" & buff(sr + 78))
            OP3_OL = Hex("&H" & buff(sr + 79))
            OP3_OM = Hex("&H" & buff(sr + 80))
            OP3_OSFC = Hex("&H" & buff(sr + 81))
            OP3_OSFF = Hex("&H" & buff(sr + 82))
            OP3_DT = Hex("&H" & buff(sr + 83))
            
            OP2_EGR1 = Hex("&H" & buff(sr + 84))
            OP2_EGR2 = Hex("&H" & buff(sr + 85))
            OP2_EGR3 = Hex("&H" & buff(sr + 86))
            OP2_EGR4 = Hex("&H" & buff(sr + 87))
            OP2_EGL1 = Hex("&H" & buff(sr + 88))
            OP2_EGL2 = Hex("&H" & buff(sr + 89))
            OP2_EGL3 = Hex("&H" & buff(sr + 90))
            OP2_EGL4 = Hex("&H" & buff(sr + 91))
            OP2_KLS_BP = Hex("&H" & buff(sr + 92))
            OP2_KLS_LD = Hex("&H" & buff(sr + 93))
            OP2_KLS_RD = Hex("&H" & buff(sr + 94))
            OP2_KLS_LC = Hex("&H" & buff(sr + 95))
            OP2_KLS_RC = Hex("&H" & buff(sr + 96))
            OP2_KRS = Hex("&H" & buff(sr + 97))
            OP2_AMP = Hex("&H" & buff(sr + 98))
            OP2_KVS = Hex("&H" & buff(sr + 99))
            OP2_OL = Hex("&H" & buff(sr + 100))
            OP2_OM = Hex("&H" & buff(sr + 101))
            OP2_OSFC = Hex("&H" & buff(sr + 102))
            OP2_OSFF = Hex("&H" & buff(sr + 103))
            OP2_DT = Hex("&H" & buff(sr + 104))
            
            OP1_EGR1 = Hex("&H" & buff(sr + 105))
            OP1_EGR2 = Hex("&H" & buff(sr + 106))
            OP1_EGR3 = Hex("&H" & buff(sr + 107))
            OP1_EGR4 = Hex("&H" & buff(sr + 108))
            OP1_EGL1 = Hex("&H" & buff(sr + 109))
            OP1_EGL2 = Hex("&H" & buff(sr + 110))
            OP1_EGL3 = Hex("&H" & buff(sr + 111))
            OP1_EGL4 = Hex("&H" & buff(sr + 112))
            OP1_KLS_BP = Hex("&H" & buff(sr + 113))
            OP1_KLS_LD = Hex("&H" & buff(sr + 114))
            OP1_KLS_RD = Hex("&H" & buff(sr + 115))
            OP1_KLS_LC = Hex("&H" & buff(sr + 116))
            OP1_KLS_RC = Hex("&H" & buff(sr + 117))
            OP1_KRS = Hex("&H" & buff(sr + 118))
            OP1_AMP = Hex("&H" & buff(sr + 119))
            OP1_KVS = Hex("&H" & buff(sr + 120))
            OP1_OL = Hex("&H" & buff(sr + 121))
            OP1_OM = Hex("&H" & buff(sr + 122))
            OP1_OSFC = Hex("&H" & buff(sr + 123))
            OP1_OSFF = Hex("&H" & buff(sr + 124))
            OP1_DT = Hex("&H" & buff(sr + 125))
            
            PR1 = Hex("&H" & buff(sr + 126))
            PR2 = Hex("&H" & buff(sr + 127))
            PR3 = Hex("&H" & buff(sr + 128))
            PR4 = Hex("&H" & buff(sr + 129))
            PL1 = Hex("&H" & buff(sr + 130))
            PL2 = Hex("&H" & buff(sr + 131))
            PL3 = Hex("&H" & buff(sr + 132))
            PL4 = Hex("&H" & buff(sr + 133))
            
            ALG = Hex("&H" & buff(sr + 134))
            FB = Hex("&H" & buff(sr + 135))
            OSC_Sync = Hex("&H" & buff(sr + 136))
            LFO_Speed = Hex("&H" & buff(sr + 137))
            LFO_Delay = Hex("&H" & buff(sr + 138))
            PMD = Hex("&H" & buff(sr + 139))
            AMD = Hex("&H" & buff(sr + 140))
            LFO_Syn = Hex("&H" & buff(sr + 141))
            LFO_Wave = Hex("&H" & buff(sr + 142))
            PMS = Hex("&H" & buff(sr + 143))
            TRS = Hex("&H" & buff(sr + 144))
            
            'VoiceNameをコードから文字列に変換
            VoiceName = Chr(Hex("&H" & buff(sr + 145))) & Chr(Hex("&H" & buff(sr + 146))) & Chr(Hex("&H" & buff(sr + 147))) & Chr(Hex("&H" & buff(sr + 148))) & Chr(Hex("&H" & buff(sr + 149))) & Chr(Hex("&H" & buff(sr + 150))) & Chr(Hex("&H" & buff(sr + 151))) & Chr(Hex("&H" & buff(sr + 152))) & Chr(Hex("&H" & buff(sr + 153))) & Chr(Hex("&H" & buff(sr + 154)))
            
            OPR_S = Hex("&H" & buff(sr + 155))
                    
        ElseIf Data_set = 32 Then
        
        '32Voiceバルクデータの場合
            OP6_EGR1 = Hex("&H" & buff(sr))
            OP6_EGR2 = Hex("&H" & buff(sr + 1))
            OP6_EGR3 = Hex("&H" & buff(sr + 2))
            OP6_EGR4 = Hex("&H" & buff(sr + 3))
            OP6_EGL1 = Hex("&H" & buff(sr + 4))
            OP6_EGL2 = Hex("&H" & buff(sr + 5))
            OP6_EGL3 = Hex("&H" & buff(sr + 6))
            OP6_EGL4 = Hex("&H" & buff(sr + 7))
            OP6_KLS_BP = Hex("&H" & buff(sr + 8))
            OP6_KLS_LD = Hex("&H" & buff(sr + 9))
            OP6_KLS_RD = Hex("&H" & buff(sr + 10))
            
            'byte             bit #
            ' #     6   5   4   3   2   1   0   param A       range  param B       range
            '----  --- --- --- --- --- --- ---  ------------  -----  ------------  -----
            '11    0   0   0 |  RC   |   LC  | SCL LEFT CURVE 0-3   SCL RGHT CURVE 0-3
            OP6_KLS_LC = BinToDec(Left(DecToBin(Hex("&H" & buff(sr + 11)), 4), 2))
            OP6_KLS_RC = BinToDec(Right(DecToBin(Hex("&H" & buff(sr + 11)), 4), 2))
            
            '12  |      DET      |     RS    | OSC DETUNE     0-14  KBD RATE SCALE 0-7
            OP6_DT = BinToDec(Left(DecToBin(Hex("&H" & buff(sr + 12)), 7), 4))
            OP6_KRS = BinToDec(Right(DecToBin(Hex("&H" & buff(sr + 12)), 7), 3))
            
            '13    0   0 |    KVS    |  AMS  | KEY VEL SENS   0-7   AMP MOD SENS   0-3
            OP6_KVS = BinToDec(Left(DecToBin(Hex("&H" & buff(sr + 13)), 5), 3))
            OP6_AMP = BinToDec(Right(DecToBin(Hex("&H" & buff(sr + 13)), 5), 2))
            
            OP6_OL = Hex("&H" & buff(sr + 14))
            
            '15    0 |         FC        | M | FREQ COARSE    0-31  OSC MODE       0-1
            OP6_OSFC = BinToDec(Left(DecToBin(Hex("&H" & buff(sr + 15)), 6), 5))
            OP6_OM = BinToDec(Right(DecToBin(Hex("&H" & buff(sr + 15)), 6), 1))
            
            OP6_OSFF = Hex("&H" & buff(sr + 16))
            
            
            OP5_EGR1 = Hex("&H" & buff(sr + 17))
            OP5_EGR2 = Hex("&H" & buff(sr + 18))
            OP5_EGR3 = Hex("&H" & buff(sr + 19))
            OP5_EGR4 = Hex("&H" & buff(sr + 20))
            OP5_EGL1 = Hex("&H" & buff(sr + 21))
            OP5_EGL2 = Hex("&H" & buff(sr + 22))
            OP5_EGL3 = Hex("&H" & buff(sr + 23))
            OP5_EGL4 = Hex("&H" & buff(sr + 24))
            OP5_KLS_BP = Hex("&H" & buff(sr + 25))
            OP5_KLS_LD = Hex("&H" & buff(sr + 26))
            OP5_KLS_RD = Hex("&H" & buff(sr + 27))
            
            'byte             bit #
            ' #     6   5   4   3   2   1   0   param A       range  param B       range
            '----  --- --- --- --- --- --- ---  ------------  -----  ------------  -----
            '28    0   0   0 |  RC   |   LC  | SCL LEFT CURVE 0-3   SCL RGHT CURVE 0-3
            OP5_KLS_LC = BinToDec(Left(DecToBin(Hex("&H" & buff(sr + 28)), 4), 2))
            OP5_KLS_RC = BinToDec(Right(DecToBin(Hex("&H" & buff(sr + 28)), 4), 2))
            
            '29  |      DET      |     RS    | OSC DETUNE     0-14  KBD RATE SCALE 0-7
            OP5_DT = BinToDec(Left(DecToBin(Hex("&H" & buff(sr + 29)), 7), 4))
            OP5_KRS = BinToDec(Right(DecToBin(Hex("&H" & buff(sr + 29)), 7), 3))
            
            '30    0   0 |    KVS    |  AMS  | KEY VEL SENS   0-7   AMP MOD SENS   0-3
            OP5_KVS = BinToDec(Left(DecToBin(Hex("&H" & buff(sr + 30)), 5), 3))
            OP5_AMP = BinToDec(Right(DecToBin(Hex("&H" & buff(sr + 30)), 5), 2))
            
            OP5_OL = Hex("&H" & buff(sr + 31))
            
            '32    0 |         FC        | M | FREQ COARSE    0-31  OSC MODE       0-1
            OP5_OSFC = BinToDec(Left(DecToBin(Hex("&H" & buff(sr + 32)), 6), 5))
            OP5_OM = BinToDec(Right(DecToBin(Hex("&H" & buff(sr + 32)), 6), 1))
            
            OP5_OSFF = Hex("&H" & buff(sr + 33))
            
            OP4_EGR1 = Hex("&H" & buff(sr + 34))
            OP4_EGR2 = Hex("&H" & buff(sr + 35))
            OP4_EGR3 = Hex("&H" & buff(sr + 36))
            OP4_EGR4 = Hex("&H" & buff(sr + 37))
            OP4_EGL1 = Hex("&H" & buff(sr + 38))
            OP4_EGL2 = Hex("&H" & buff(sr + 39))
            OP4_EGL3 = Hex("&H" & buff(sr + 40))
            OP4_EGL4 = Hex("&H" & buff(sr + 41))
            OP4_KLS_BP = Hex("&H" & buff(sr + 42))
            OP4_KLS_LD = Hex("&H" & buff(sr + 43))
            OP4_KLS_RD = Hex("&H" & buff(sr + 44))
            
            'byte             bit #
            ' #     6   5   4   3   2   1   0   param A       range  param B       range
            '----  --- --- --- --- --- --- ---  ------------  -----  ------------  -----
            '45    0   0   0 |  RC   |   LC  | SCL LEFT CURVE 0-3   SCL RGHT CURVE 0-3
            OP4_KLS_LC = BinToDec(Left(DecToBin(Hex("&H" & buff(sr + 45)), 4), 2))
            OP4_KLS_RC = BinToDec(Right(DecToBin(Hex("&H" & buff(sr + 45)), 4), 2))
            
            '46  |      DET      |     RS    | OSC DETUNE     0-14  KBD RATE SCALE 0-7
            OP4_DT = BinToDec(Left(DecToBin(Hex("&H" & buff(sr + 46)), 7), 4))
            OP4_KRS = BinToDec(Right(DecToBin(Hex("&H" & buff(sr + 46)), 7), 3))
            
            '47    0   0 |    KVS    |  AMS  | KEY VEL SENS   0-7   AMP MOD SENS   0-3
            OP4_KVS = BinToDec(Left(DecToBin(Hex("&H" & buff(sr + 47)), 5), 3))
            OP4_AMP = BinToDec(Right(DecToBin(Hex("&H" & buff(sr + 47)), 5), 2))
            
            OP4_OL = Hex("&H" & buff(sr + 48))
            
            '49    0 |         FC        | M | FREQ COARSE    0-31  OSC MODE       0-1
            OP4_OSFC = BinToDec(Left(DecToBin(Hex("&H" & buff(sr + 49)), 6), 5))
            OP4_OM = BinToDec(Right(DecToBin(Hex("&H" & buff(sr + 49)), 6), 1))
            
            OP4_OSFF = Hex("&H" & buff(sr + 50))
            
            
            OP3_EGR1 = Hex("&H" & buff(sr + 51))
            OP3_EGR2 = Hex("&H" & buff(sr + 52))
            OP3_EGR3 = Hex("&H" & buff(sr + 53))
            OP3_EGR4 = Hex("&H" & buff(sr + 54))
            OP3_EGL1 = Hex("&H" & buff(sr + 55))
            OP3_EGL2 = Hex("&H" & buff(sr + 56))
            OP3_EGL3 = Hex("&H" & buff(sr + 57))
            OP3_EGL4 = Hex("&H" & buff(sr + 58))
            OP3_KLS_BP = Hex("&H" & buff(sr + 59))
            OP3_KLS_LD = Hex("&H" & buff(sr + 60))
            OP3_KLS_RD = Hex("&H" & buff(sr + 61))
            
            'byte             bit #
            ' #     6   5   4   3   2   1   0   param A       range  param B       range
            '----  --- --- --- --- --- --- ---  ------------  -----  ------------  -----
            '62    0   0   0 |  RC   |   LC  | SCL LEFT CURVE 0-3   SCL RGHT CURVE 0-3
            OP3_KLS_LC = BinToDec(Left(DecToBin(Hex("&H" & buff(sr + 62)), 4), 2))
            OP3_KLS_RC = BinToDec(Right(DecToBin(Hex("&H" & buff(sr + 62)), 4), 2))
            
            '63  |      DET      |     RS    | OSC DETUNE     0-14  KBD RATE SCALE 0-7
            OP3_DT = BinToDec(Left(DecToBin(Hex("&H" & buff(sr + 63)), 7), 4))
            OP3_KRS = BinToDec(Right(DecToBin(Hex("&H" & buff(sr + 63)), 7), 3))
            
            '64    0   0 |    KVS    |  AMS  | KEY VEL SENS   0-7   AMP MOD SENS   0-3
            OP3_KVS = BinToDec(Left(DecToBin(Hex("&H" & buff(sr + 64)), 5), 3))
            OP3_AMP = BinToDec(Right(DecToBin(Hex("&H" & buff(sr + 64)), 5), 2))
            
            OP3_OL = Hex("&H" & buff(sr + 65))
            
            '66    0 |         FC        | M | FREQ COARSE    0-31  OSC MODE       0-1
            OP3_OSFC = BinToDec(Left(DecToBin(Hex("&H" & buff(sr + 66)), 6), 5))
            OP3_OM = BinToDec(Right(DecToBin(Hex("&H" & buff(sr + 66)), 6), 1))
            
            OP3_OSFF = Hex("&H" & buff(sr + 67))
                       
            
            OP2_EGR1 = Hex("&H" & buff(sr + 68))
            OP2_EGR2 = Hex("&H" & buff(sr + 69))
            OP2_EGR3 = Hex("&H" & buff(sr + 70))
            OP2_EGR4 = Hex("&H" & buff(sr + 71))
            OP2_EGL1 = Hex("&H" & buff(sr + 72))
            OP2_EGL2 = Hex("&H" & buff(sr + 73))
            OP2_EGL3 = Hex("&H" & buff(sr + 74))
            OP2_EGL4 = Hex("&H" & buff(sr + 75))
            OP2_KLS_BP = Hex("&H" & buff(sr + 76))
            OP2_KLS_LD = Hex("&H" & buff(sr + 77))
            OP2_KLS_RD = Hex("&H" & buff(sr + 78))
            
            'byte             bit #
            ' #     6   5   4   3   2   1   0   param A       range  param B       range
            '----  --- --- --- --- --- --- ---  ------------  -----  ------------  -----
            '79    0   0   0 |  RC   |   LC  | SCL LEFT CURVE 0-3   SCL RGHT CURVE 0-3
            OP2_KLS_LC = BinToDec(Left(DecToBin(Hex("&H" & buff(sr + 79)), 4), 2))
            OP2_KLS_RC = BinToDec(Right(DecToBin(Hex("&H" & buff(sr + 79)), 4), 2))
            
            '80  |      DET      |     RS    | OSC DETUNE     0-14  KBD RATE SCALE 0-7
            OP2_DT = BinToDec(Left(DecToBin(Hex("&H" & buff(sr + 80)), 7), 4))
            OP2_KRS = BinToDec(Right(DecToBin(Hex("&H" & buff(sr + 80)), 7), 3))
            
            '81    0   0 |    KVS    |  AMS  | KEY VEL SENS   0-7   AMP MOD SENS   0-3
            OP2_KVS = BinToDec(Left(DecToBin(Hex("&H" & buff(sr + 81)), 5), 3))
            OP2_AMP = BinToDec(Right(DecToBin(Hex("&H" & buff(sr + 81)), 5), 2))
            
            OP2_OL = Hex("&H" & buff(sr + 82))
            
            '83    0 |         FC        | M | FREQ COARSE    0-31  OSC MODE       0-1
            OP2_OSFC = BinToDec(Left(DecToBin(Hex("&H" & buff(sr + 83)), 6), 5))
            OP2_OM = BinToDec(Right(DecToBin(Hex("&H" & buff(sr + 83)), 6), 1))
            
            OP2_OSFF = Hex("&H" & buff(sr + 84))
            
            
            OP1_EGR1 = Hex("&H" & buff(sr + 85))
            OP1_EGR2 = Hex("&H" & buff(sr + 86))
            OP1_EGR3 = Hex("&H" & buff(sr + 87))
            OP1_EGR4 = Hex("&H" & buff(sr + 88))
            OP1_EGL1 = Hex("&H" & buff(sr + 89))
            OP1_EGL2 = Hex("&H" & buff(sr + 90))
            OP1_EGL3 = Hex("&H" & buff(sr + 91))
            OP1_EGL4 = Hex("&H" & buff(sr + 92))
            OP1_KLS_BP = Hex("&H" & buff(sr + 93))
            OP1_KLS_LD = Hex("&H" & buff(sr + 94))
            OP1_KLS_RD = Hex("&H" & buff(sr + 95))
            
            'byte             bit #
            ' #     6   5   4   3   2   1   0   param A       range  param B       range
            '----  --- --- --- --- --- --- ---  ------------  -----  ------------  -----
            '96    0   0   0 |  RC   |   LC  | SCL LEFT CURVE 0-3   SCL RGHT CURVE 0-3
            OP1_KLS_LC = BinToDec(Left(DecToBin(Hex("&H" & buff(sr + 96)), 4), 2))
            OP1_KLS_RC = BinToDec(Right(DecToBin(Hex("&H" & buff(sr + 96)), 4), 2))
            
            '97  |      DET      |     RS    | OSC DETUNE     0-14  KBD RATE SCALE 0-7
            OP1_DT = BinToDec(Left(DecToBin(Hex("&H" & buff(sr + 97)), 7), 4))
            OP1_KRS = BinToDec(Right(DecToBin(Hex("&H" & buff(sr + 97)), 7), 3))
            
            '98    0   0 |    KVS    |  AMS  | KEY VEL SENS   0-7   AMP MOD SENS   0-3
            OP1_KVS = BinToDec(Left(DecToBin(Hex("&H" & buff(sr + 98)), 5), 3))
            OP1_AMP = BinToDec(Right(DecToBin(Hex("&H" & buff(sr + 98)), 5), 2))
            
            OP1_OL = Hex("&H" & buff(sr + 99))
            
            '100    0 |         FC        | M | FREQ COARSE    0-31  OSC MODE       0-1
            OP1_OSFC = BinToDec(Left(DecToBin(Hex("&H" & buff(sr + 100)), 6), 5))
            OP1_OM = BinToDec(Right(DecToBin(Hex("&H" & buff(sr + 100)), 6), 1))
            
            OP1_OSFF = Hex("&H" & buff(sr + 101))
            
            
            PR1 = Hex("&H" & buff(sr + 102))
            PR2 = Hex("&H" & buff(sr + 103))
            PR3 = Hex("&H" & buff(sr + 104))
            PR4 = Hex("&H" & buff(sr + 105))
            PL1 = Hex("&H" & buff(sr + 106))
            PL2 = Hex("&H" & buff(sr + 107))
            PL3 = Hex("&H" & buff(sr + 108))
            PL4 = Hex("&H" & buff(sr + 109))
            
            '110    0   0 |        ALG        | ALGORITHM     0-31
            ALG = BinToDec(Right(DecToBin(Hex("&H" & buff(sr + 110)), 5), 5))
            
            '111    0   0   0 |OKS|    FB     | OSC KEY SYNC  0-1    FEEDBACK      0-7
            OSC_Sync = BinToDec(Left(DecToBin(Hex("&H" & buff(sr + 111)), 4), 1))
            FB = BinToDec(Right(DecToBin(Hex("&H" & buff(sr + 111)), 4), 3))

            LFO_Speed = Hex("&H" & buff(sr + 112))
            LFO_Delay = Hex("&H" & buff(sr + 113))
            PMD = Hex("&H" & buff(sr + 114))
            AMD = Hex("&H" & buff(sr + 115))
            
            '116  |  LPMS |      LFW      |LKS| LF PT MOD SNS 0-7   WAVE 0-5,  SYNC 0-1
            PMS = BinToDec(Left(DecToBin(Hex("&H" & buff(sr + 116)), 7), 3))
            LFO_Wave = BinToDec(Mid(DecToBin(Hex("&H" & buff(sr + 116)), 7), 4, 3))
            LFO_Sync = BinToDec(Right(DecToBin(Hex("&H" & buff(sr + 116)), 7), 1))
            
            TRS = Hex("&H" & buff(sr + 117))
            
            'VoiceNameをコードから文字列に変換
            VoiceName = Chr(Hex("&H" & buff(sr + 118))) & Chr(Hex("&H" & buff(sr + 119))) & Chr(Hex("&H" & buff(sr + 120))) & Chr(Hex("&H" & buff(sr + 121))) & Chr(Hex("&H" & buff(sr + 122))) & Chr(Hex("&H" & buff(sr + 123))) & Chr(Hex("&H" & buff(sr + 124))) & Chr(Hex("&H" & buff(sr + 125))) & Chr(Hex("&H" & buff(sr + 126))) & Chr(Hex("&H" & buff(sr + 127)))
            
        End If
        
        'Voiceデータをシートへの書き込み
        Sheets(wsTALGet).Activate

        Cells(tr, tc).Value = VoiceName
        Cells(tr, tc - 1).Value = strLibName
        Cells(tr, tc + 1).Value = ALG + 1
        Cells(tr, tc + 2).Value = FB

        Cells(tr, tc + 3).Value = OP1_EGR1
        Cells(tr, tc + 4).Value = OP1_EGR2
        Cells(tr, tc + 5).Value = OP1_EGR3
        Cells(tr, tc + 6).Value = OP1_EGR4
        Cells(tr, tc + 7).Value = OP1_EGL1
        Cells(tr, tc + 8).Value = OP1_EGL2
        Cells(tr, tc + 9).Value = OP1_EGL3
        Cells(tr, tc + 10).Value = OP1_EGL4
        Cells(tr, tc + 11).Value = OP1_KLS_BP
        Cells(tr, tc + 12).Value = OP1_KLS_LD
        Cells(tr, tc + 13).Value = OP1_KLS_RD
        Cells(tr, tc + 14).Value = OP1_KLS_LC
        Cells(tr, tc + 15).Value = OP1_KLS_RC
        Cells(tr, tc + 16).Value = OP1_KRS
        Cells(tr, tc + 17).Value = OP1_AMP
        Cells(tr, tc + 18).Value = OP1_KVS
        Cells(tr, tc + 19).Value = OP1_OL
        Cells(tr, tc + 20).Value = OP1_OM
        Cells(tr, tc + 21).Value = OP1_OSFC
        Cells(tr, tc + 22).Value = OP1_OSFF
        Cells(tr, tc + 23).Value = OP1_DT - 7
        
        Cells(tr, tc + 24).Value = OP2_EGR1
        Cells(tr, tc + 25).Value = OP2_EGR2
        Cells(tr, tc + 26).Value = OP2_EGR3
        Cells(tr, tc + 27).Value = OP2_EGR4
        Cells(tr, tc + 28).Value = OP2_EGL1
        Cells(tr, tc + 29).Value = OP2_EGL2
        Cells(tr, tc + 30).Value = OP2_EGL3
        Cells(tr, tc + 31).Value = OP2_EGL4
        Cells(tr, tc + 32).Value = OP2_KLS_BP
        Cells(tr, tc + 33).Value = OP2_KLS_LD
        Cells(tr, tc + 34).Value = OP2_KLS_RD
        Cells(tr, tc + 35).Value = OP2_KLS_LC
        Cells(tr, tc + 36).Value = OP2_KLS_RC
        Cells(tr, tc + 37).Value = OP2_KRS
        Cells(tr, tc + 38).Value = OP2_AMP
        Cells(tr, tc + 39).Value = OP2_KVS
        Cells(tr, tc + 40).Value = OP2_OL
        Cells(tr, tc + 41).Value = OP2_OM
        Cells(tr, tc + 42).Value = OP2_OSFC
        Cells(tr, tc + 43).Value = OP2_OSFF
        Cells(tr, tc + 44).Value = OP2_DT - 7
        
        Cells(tr, tc + 45).Value = OP3_EGR1
        Cells(tr, tc + 46).Value = OP3_EGR2
        Cells(tr, tc + 47).Value = OP3_EGR3
        Cells(tr, tc + 48).Value = OP3_EGR4
        Cells(tr, tc + 49).Value = OP3_EGL1
        Cells(tr, tc + 50).Value = OP3_EGL2
        Cells(tr, tc + 51).Value = OP3_EGL3
        Cells(tr, tc + 52).Value = OP3_EGL4
        Cells(tr, tc + 53).Value = OP3_KLS_BP
        Cells(tr, tc + 54).Value = OP3_KLS_LD
        Cells(tr, tc + 55).Value = OP3_KLS_RD
        Cells(tr, tc + 56).Value = OP3_KLS_LC
        Cells(tr, tc + 57).Value = OP3_KLS_RC
        Cells(tr, tc + 58).Value = OP3_KRS
        Cells(tr, tc + 59).Value = OP3_AMP
        Cells(tr, tc + 60).Value = OP3_KVS
        Cells(tr, tc + 61).Value = OP3_OL
        Cells(tr, tc + 62).Value = OP3_OM
        Cells(tr, tc + 63).Value = OP3_OSFC
        Cells(tr, tc + 64).Value = OP3_OSFF
        Cells(tr, tc + 65).Value = OP3_DT - 7
        
        Cells(tr, tc + 66).Value = OP4_EGR1
        Cells(tr, tc + 67).Value = OP4_EGR2
        Cells(tr, tc + 68).Value = OP4_EGR3
        Cells(tr, tc + 69).Value = OP4_EGR4
        Cells(tr, tc + 70).Value = OP4_EGL1
        Cells(tr, tc + 71).Value = OP4_EGL2
        Cells(tr, tc + 72).Value = OP4_EGL3
        Cells(tr, tc + 73).Value = OP4_EGL4
        Cells(tr, tc + 74).Value = OP4_KLS_BP
        Cells(tr, tc + 75).Value = OP4_KLS_LD
        Cells(tr, tc + 76).Value = OP4_KLS_RD
        Cells(tr, tc + 77).Value = OP4_KLS_LC
        Cells(tr, tc + 78).Value = OP4_KLS_RC
        Cells(tr, tc + 79).Value = OP4_KRS
        Cells(tr, tc + 80).Value = OP4_AMP
        Cells(tr, tc + 81).Value = OP4_KVS
        Cells(tr, tc + 82).Value = OP4_OL
        Cells(tr, tc + 83).Value = OP4_OM
        Cells(tr, tc + 84).Value = OP4_OSFC
        Cells(tr, tc + 85).Value = OP4_OSFF
        Cells(tr, tc + 86).Value = OP4_DT - 7
        
        Cells(tr, tc + 87).Value = OP5_EGR1
        Cells(tr, tc + 88).Value = OP5_EGR2
        Cells(tr, tc + 89).Value = OP5_EGR3
        Cells(tr, tc + 90).Value = OP5_EGR4
        Cells(tr, tc + 91).Value = OP5_EGL1
        Cells(tr, tc + 92).Value = OP5_EGL2
        Cells(tr, tc + 93).Value = OP5_EGL3
        Cells(tr, tc + 94).Value = OP5_EGL4
        Cells(tr, tc + 95).Value = OP5_KLS_BP
        Cells(tr, tc + 96).Value = OP5_KLS_LD
        Cells(tr, tc + 97).Value = OP5_KLS_RD
        Cells(tr, tc + 98).Value = OP5_KLS_LC
        Cells(tr, tc + 99).Value = OP5_KLS_RC
        Cells(tr, tc + 100).Value = OP5_KRS
        Cells(tr, tc + 101).Value = OP5_AMP
        Cells(tr, tc + 102).Value = OP5_KVS
        Cells(tr, tc + 103).Value = OP5_OL
        Cells(tr, tc + 104).Value = OP5_OM
        Cells(tr, tc + 105).Value = OP5_OSFC
        Cells(tr, tc + 106).Value = OP5_OSFF
        Cells(tr, tc + 107).Value = OP5_DT - 7

        Cells(tr, tc + 108).Value = OP6_EGR1
        Cells(tr, tc + 109).Value = OP6_EGR2
        Cells(tr, tc + 110).Value = OP6_EGR3
        Cells(tr, tc + 111).Value = OP6_EGR4
        Cells(tr, tc + 112).Value = OP6_EGL1
        Cells(tr, tc + 113).Value = OP6_EGL2
        Cells(tr, tc + 114).Value = OP6_EGL3
        Cells(tr, tc + 115).Value = OP6_EGL4
        Cells(tr, tc + 116).Value = OP6_KLS_BP
        Cells(tr, tc + 117).Value = OP6_KLS_LD
        Cells(tr, tc + 118).Value = OP6_KLS_RD
        Cells(tr, tc + 119).Value = OP6_KLS_LC
        Cells(tr, tc + 120).Value = OP6_KLS_RC
        Cells(tr, tc + 121).Value = OP6_KRS
        Cells(tr, tc + 122).Value = OP6_AMP
        Cells(tr, tc + 123).Value = OP6_KVS
        Cells(tr, tc + 124).Value = OP6_OL
        Cells(tr, tc + 125).Value = OP6_OM
        Cells(tr, tc + 126).Value = OP6_OSFC
        Cells(tr, tc + 127).Value = OP6_OSFF
        Cells(tr, tc + 128).Value = OP6_DT - 7
        
        Cells(tr, tc + 129).Value = PR1
        Cells(tr, tc + 130).Value = PR2
        Cells(tr, tc + 131).Value = PR3
        Cells(tr, tc + 132).Value = PR4
        Cells(tr, tc + 133).Value = PL1
        Cells(tr, tc + 134).Value = PL2
        Cells(tr, tc + 135).Value = PL3
        Cells(tr, tc + 136).Value = PL4
        Cells(tr, tc + 137).Value = OSC_Sync
        Cells(tr, tc + 138).Value = LFO_Speed
        Cells(tr, tc + 139).Value = LFO_Delay
        Cells(tr, tc + 140).Value = PMD
        Cells(tr, tc + 141).Value = AMD
        Cells(tr, tc + 142).Value = LFO_Sync
        Cells(tr, tc + 143).Value = LFO_Wave
        Cells(tr, tc + 144).Value = PMS
        Cells(tr, tc + 145).Value = TRS
        Cells(tr, tc + 146).Value = OPR_S

        'カウンターの更新（32Voiceバルクデータの場合、128バイト先へ移動）
        tr = tr + 1
        sr = sr + 128
        
    Next c
    
    
'以下はデバッグ用（OutputDataDX7から出力したSysexファイルのチェックサム確認用コード）
    
'    If buff(3) = 0 Then
'
'        If chksum2 = buff(162) Then
'            Debug.Print "OK"
'            Debug.Print chksum2
'            Debug.Print buff(162)
'        Else
'            Debug.Print "Error"
'            Debug.Print chksum2
'            Debug.Print buff(162)
'        End If
'
'    ElseIf buff(3) = 9 Then
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
'Sysexファイルからバイナリデータを読み込み
'********************************************
Sub ReadSYBinaryFile(ByVal strfil As String)

    Dim buff() As Byte
    Dim fp As Long
    Dim filesize As Long, nowVal As Long
    Dim idx As Long, gyo As Long
    Dim strBinary As String
    
    'データ出力用シート
    wsTALGet = "SysexSY77Data"
    
    Sheets(wsTALGet).Activate
    ActiveSheet.Unprotect
    Range("D3:AOK66").ClearContents
    ActiveSheet.Protect
    
    
    
    '//FreeFile関数で使用可能なファイル番号を割り当て
    fp = FreeFile
    '//ファイルを開く
    Open strfil For Binary As #fp
    '//ファイルサイズ分の読み込み領域を確保して読み込む場合の実装例
    ReDim buff(FileLen(strfil))
    Get #fp, 1, buff
    '//実装例ここまで
    
    r = 3
    c = 4
    
    For i = 0 To LOF(fp)
        
        nowVal = Hex("&H" & buff(i))
        
        If nowVal = 247 Then
            Cells(r, c).Value = nowVal
            c = 4
            r = r + 1
        Else
            Cells(r, c).Value = nowVal
            c = c + 1
        End If
        
    Next i
    
    '//ファイルを閉じる
    Close (fp)
End Sub

'********************************************
' SY77 1Voiceバルクデータの変換
'********************************************
Sub Conv_SY77_1V_syx()

    Dim strFilePath As String
    Dim strFileName As String
    Dim strFile As String
    Dim Msg, Style, Title, Response
    
    Style = vbOKCancel
    Title = "エラー"
    
    '//ファイル名を生成
    
    Sheets("MenuSY77").Select
    strFilePath = Cells(17, 5).Value
    strFileName = Cells(18, 5).Value

    If strFilePath = "" Then
        strFilePath = ThisWorkbook.Path
    End If
    
    If strFileName = "" Then
        MsgBox "ファイル名が指定されていません。"
        
    Else
        strFile = strFilePath & "\" & strFileName
                
        If Dir(strFile) = "" Then
        
            Call Write_SY_syx(1, strFile)
            Sheets("MenuSY77").Select
            MsgBox "Sysexデータの書き出しが完了しました。"
            
        Else
        
            Msg = strFile & vbCrLf & _
                    "がすでに存在します。上書きしてもよろしいですか？"

            Response = MsgBox(Msg, Style, Title)
            
            If Response = vbOK Then
            
                Kill strFile
                Call Write_SY_syx(1, strFile)
                Sheets("MenuSY77").Select
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
' SY77 64Voiceバルクデータの変換
'********************************************
Sub Conv_SY77_64V_syx()

    Dim strFilePath As String
    Dim strFileName As String
    Dim strFile As String
    Dim Msg, Style, Title, Response
    
    Style = vbOKCancel
    Title = "エラー"
    
    '//ファイル名を生成
    
    Sheets("MenuSY77").Select
    strFilePath = Cells(24, 5).Value
    strFileName = Cells(25, 5).Value

    If strFilePath = "" Then
        strFilePath = ThisWorkbook.Path
    End If
    
    If strFileName = "" Then
        MsgBox "ファイル名が指定されていません。"
        
    Else
        strFile = strFilePath & "\" & strFileName
                
        If Dir(strFile) = "" Then
        
            Call Write_SY_syx(64, strFile)
            Sheets("MenuSY77").Select
            MsgBox "Sysexデータの書き出しが完了しました。"
            
        Else
        
            Msg = strFile & vbCrLf & _
                    "がすでに存在します。上書きしてもよろしいですか？"

            Response = MsgBox(Msg, Style, Title)
            
            If Response = vbOK Then
            
                Kill strFile
                Call Write_SY_syx(64, strFile)
                Sheets("MenuSY77").Select
                MsgBox "Sysexデータの書き出しが完了しました。"
                
            ElseIf Response = vbCancel Then
            
                MsgBox "Sysexデータの書き出しをキャンセルしました。"
            
            Else
                
                MsgBox "Sysexデータの書き出しを中止しました。"
            
            End If
        End If
    End If
    
End Sub

Sub Write_SY_syx(dSet As Integer, fName As String)

Dim V_Data_str As String
Dim VoiceData_str() As String
Dim VoiceData_byt() As Byte
Dim r As Long
Dim c As Long
Dim i As Long
Dim j As Integer
Dim nowVal As Long
Dim wsSource As String

V_Data_str = "0"
r = 3

wsSource = "SY77OutputData"
Sheets(wsSource).Activate

For j = 1 To dSet

    c = 5
    
    Do
        nowVal = Cells(r, c).Value
        V_Data_str = V_Data_str & " " & Hex(nowVal)
        c = c + 1
    Loop Until nowVal = 247

        r = r + 1
        
Next j
    
    'V_Data_strを分割して配列VoiceData_strに入れる
    VoiceData_str = Split(V_Data_str, " ")
    
    
    '配列VoiceData_strのデータをByteデータに変換するため配列VoiceData_bytに入れる
    ReDim VoiceData_byt(UBound(VoiceData_str) - 1)

    For i = 1 To UBound(VoiceData_str)
        VoiceData_byt(i - 1) = "&h" & VoiceData_str(i)
    Next i
    
    '1Voiceデータの場合30バイト目に7Fを書きこむ
    If dSet = 1 Then
        VoiceData_byt(31) = "&h" & "7F"
    End If
    
    'バイナリーファイルに書き込む
    fh = FreeFile
    Open fName For Binary Access Write As #fh
    Put #fh, , VoiceData_byt
    Close #fh

End Sub

Function EL_Mode(val As Integer) As String
    
    Select Case val
        Case 0
            EL_Mode = "1FM/Mono"
        Case 1
            EL_Mode = "2FM/Mono"
        Case 2
            EL_Mode = "1FM/Mono"
        Case 3
            EL_Mode = "1FM/Poly"
        Case 4
            EL_Mode = "2FM/Poly"
        Case 5
            EL_Mode = "1AWM/Poly"
        Case 6
            EL_Mode = "2AWM/Poly"
        Case 7
            EL_Mode = "4AWM/Poly"
        Case 8
            EL_Mode = "1FM_1AWM/Poly"
        Case 9
            EL_Mode = "2FM_2AWM/Poly"
        Case 10
            EL_Mode = "Drum_Set"
    End Select
    
End Function

