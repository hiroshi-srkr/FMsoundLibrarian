Attribute VB_Name = "Module12"

'********************************************
' DX7 1Voiceバルクデータの変換
'********************************************
Sub Conv_DX7_SV_syx()

    Dim strFilePath As String
    Dim strFileName As String
    Dim strFile As String
    Dim Msg, Style, Title, Response
    
    Style = vbOKCancel
    Title = "エラー"
    
    '//ファイル名を生成
    
    Sheets("MenuDX7").Select
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
        
            Call Write_DX7syx2(1, strFile)
            Sheets("MenuDX7").Select
            MsgBox "Sysexデータの書き出しが完了しました。"
            
        Else
        
            Msg = strFile & vbCrLf & _
                    "がすでに存在します。上書きしてもよろしいですか？"

            Response = MsgBox(Msg, Style, Title)
            
            If Response = vbOK Then
            
                Kill strFile
                Call Write_DX7syx2(1, strFile)
                Sheets("MenuDX7").Select
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
' DX7 32Voiceバルクデータの変換
'********************************************
Sub Conv_DX7_MV_syx()

    Dim strFilePath As String
    Dim strFileName As String
    Dim strFile As String
    Dim Msg, Style, Title, Response
    
    Style = vbOKCancel
    Title = "エラー"
    
    '//ファイル名を生成
    
    Sheets("MenuDX7").Select
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
        
            Call Write_DX7syx2(32, strFile)
            Sheets("MenuDX7").Select
            MsgBox "Sysexデータの書き出しが完了しました。"
            
        Else
        
            Msg = strFile & vbCrLf & _
                    "がすでに存在します。上書きしてもよろしいですか？"

            Response = MsgBox(Msg, Style, Title)
            
            If Response = vbOK Then
            
                Kill strFile
                Call Write_DX7syx2(32, strFile)
                Sheets("MenuDX7").Select
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
' Sysexファイルへバイナリデータを書き込み
'********************************************

Sub Write_DX7syx2(DataSet As Integer, filename As String)
    
    Dim wsSource As String
    Dim VoiceName As String
    Dim ALG As Integer, FB As Integer
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
        
    Dim data_str As String
    Dim sv_hdr, mv_hdr  As String
    Dim VoiceData_str() As String
    Dim VoiceData_byt() As Byte
    Dim chksum_data() As String
    Dim chksum_byte() As Byte
        
    Dim chksum As Long
        
    sv_hdr = "F0 43 00 00 00 9B"
    mv_hdr = "F0 43 00 09 20 00"
    chksum = 0
    data_str = ""
    
    sr = 2
    sc = 2
    tr = 2
    tc = 2
    
    wsSource = "DX7_OutputData"
    
    If DataSet = 1 Then

        data_str = sv_hdr
        
    ElseIf DataSet = 32 Then

        data_str = mv_hdr
        
    End If

    For c = 1 To DataSet
    
    'SysexDataシートからデータを取得
    
        Sheets(wsSource).Activate
        VoiceName = Cells(tr, tc).Value
        strLibName = Cells(tr, tc - 1).Value
        ALG = Cells(tr, tc + 1).Value
        FB = Cells(tr, tc + 2).Value

        OP1_EGR1 = Cells(tr, tc + 3).Value
        OP1_EGR2 = Cells(tr, tc + 4).Value
        OP1_EGR3 = Cells(tr, tc + 5).Value
        OP1_EGR4 = Cells(tr, tc + 6).Value
        OP1_EGL1 = Cells(tr, tc + 7).Value
        OP1_EGL2 = Cells(tr, tc + 8).Value
        OP1_EGL3 = Cells(tr, tc + 9).Value
        OP1_EGL4 = Cells(tr, tc + 10).Value
        OP1_KLS_BP = Cells(tr, tc + 11).Value
        OP1_KLS_LD = Cells(tr, tc + 12).Value
        OP1_KLS_RD = Cells(tr, tc + 13).Value
        OP1_KLS_LC = Cells(tr, tc + 14).Value
        OP1_KLS_RC = Cells(tr, tc + 15).Value
        OP1_KRS = Cells(tr, tc + 16).Value
        OP1_AMP = Cells(tr, tc + 17).Value
        OP1_KVS = Cells(tr, tc + 18).Value
        OP1_OL = Cells(tr, tc + 19).Value
        OP1_OM = Cells(tr, tc + 20).Value
        OP1_OSFC = Cells(tr, tc + 21).Value
        OP1_OSFF = Cells(tr, tc + 22).Value
        OP1_DT = Cells(tr, tc + 23).Value
        
        OP2_EGR1 = Cells(tr, tc + 24).Value
        OP2_EGR2 = Cells(tr, tc + 25).Value
        OP2_EGR3 = Cells(tr, tc + 26).Value
        OP2_EGR4 = Cells(tr, tc + 27).Value
        OP2_EGL1 = Cells(tr, tc + 28).Value
        OP2_EGL2 = Cells(tr, tc + 29).Value
        OP2_EGL3 = Cells(tr, tc + 30).Value
        OP2_EGL4 = Cells(tr, tc + 31).Value
        OP2_KLS_BP = Cells(tr, tc + 32).Value
        OP2_KLS_LD = Cells(tr, tc + 33).Value
        OP2_KLS_RD = Cells(tr, tc + 34).Value
        OP2_KLS_LC = Cells(tr, tc + 35).Value
        OP2_KLS_RC = Cells(tr, tc + 36).Value
        OP2_KRS = Cells(tr, tc + 37).Value
        OP2_AMP = Cells(tr, tc + 38).Value
        OP2_KVS = Cells(tr, tc + 39).Value
        OP2_OL = Cells(tr, tc + 40).Value
        OP2_OM = Cells(tr, tc + 41).Value
        OP2_OSFC = Cells(tr, tc + 42).Value
        OP2_OSFF = Cells(tr, tc + 43).Value
        OP2_DT = Cells(tr, tc + 44).Value
        
        OP3_EGR1 = Cells(tr, tc + 45).Value
        OP3_EGR2 = Cells(tr, tc + 46).Value
        OP3_EGR3 = Cells(tr, tc + 47).Value
        OP3_EGR4 = Cells(tr, tc + 48).Value
        OP3_EGL1 = Cells(tr, tc + 49).Value
        OP3_EGL2 = Cells(tr, tc + 50).Value
        OP3_EGL3 = Cells(tr, tc + 51).Value
        OP3_EGL4 = Cells(tr, tc + 52).Value
        OP3_KLS_BP = Cells(tr, tc + 53).Value
        OP3_KLS_LD = Cells(tr, tc + 54).Value
        OP3_KLS_RD = Cells(tr, tc + 55).Value
        OP3_KLS_LC = Cells(tr, tc + 56).Value
        OP3_KLS_RC = Cells(tr, tc + 57).Value
        OP3_KRS = Cells(tr, tc + 58).Value
        OP3_AMP = Cells(tr, tc + 59).Value
        OP3_KVS = Cells(tr, tc + 60).Value
        OP3_OL = Cells(tr, tc + 61).Value
        OP3_OM = Cells(tr, tc + 62).Value
        OP3_OSFC = Cells(tr, tc + 63).Value
        OP3_OSFF = Cells(tr, tc + 64).Value
        OP3_DT = Cells(tr, tc + 65).Value
        
        OP4_EGR1 = Cells(tr, tc + 66).Value
        OP4_EGR2 = Cells(tr, tc + 67).Value
        OP4_EGR3 = Cells(tr, tc + 68).Value
        OP4_EGR4 = Cells(tr, tc + 69).Value
        OP4_EGL1 = Cells(tr, tc + 70).Value
        OP4_EGL2 = Cells(tr, tc + 71).Value
        OP4_EGL3 = Cells(tr, tc + 72).Value
        OP4_EGL4 = Cells(tr, tc + 73).Value
        OP4_KLS_BP = Cells(tr, tc + 74).Value
        OP4_KLS_LD = Cells(tr, tc + 75).Value
        OP4_KLS_RD = Cells(tr, tc + 76).Value
        OP4_KLS_LC = Cells(tr, tc + 77).Value
        OP4_KLS_RC = Cells(tr, tc + 78).Value
        OP4_KRS = Cells(tr, tc + 79).Value
        OP4_AMP = Cells(tr, tc + 80).Value
        OP4_KVS = Cells(tr, tc + 81).Value
        OP4_OL = Cells(tr, tc + 82).Value
        OP4_OM = Cells(tr, tc + 83).Value
        OP4_OSFC = Cells(tr, tc + 84).Value
        OP4_OSFF = Cells(tr, tc + 85).Value
        OP4_DT = Cells(tr, tc + 86).Value
        
        OP5_EGR1 = Cells(tr, tc + 87).Value
        OP5_EGR2 = Cells(tr, tc + 88).Value
        OP5_EGR3 = Cells(tr, tc + 89).Value
        OP5_EGR4 = Cells(tr, tc + 90).Value
        OP5_EGL1 = Cells(tr, tc + 91).Value
        OP5_EGL2 = Cells(tr, tc + 92).Value
        OP5_EGL3 = Cells(tr, tc + 93).Value
        OP5_EGL4 = Cells(tr, tc + 94).Value
        OP5_KLS_BP = Cells(tr, tc + 95).Value
        OP5_KLS_LD = Cells(tr, tc + 96).Value
        OP5_KLS_RD = Cells(tr, tc + 97).Value
        OP5_KLS_LC = Cells(tr, tc + 98).Value
        OP5_KLS_RC = Cells(tr, tc + 99).Value
        OP5_KRS = Cells(tr, tc + 100).Value
        OP5_AMP = Cells(tr, tc + 101).Value
        OP5_KVS = Cells(tr, tc + 102).Value
        OP5_OL = Cells(tr, tc + 103).Value
        OP5_OM = Cells(tr, tc + 104).Value
        OP5_OSFC = Cells(tr, tc + 105).Value
        OP5_OSFF = Cells(tr, tc + 106).Value
        OP5_DT = Cells(tr, tc + 107).Value

        OP6_EGR1 = Cells(tr, tc + 108).Value
        OP6_EGR2 = Cells(tr, tc + 109).Value
        OP6_EGR3 = Cells(tr, tc + 110).Value
        OP6_EGR4 = Cells(tr, tc + 111).Value
        OP6_EGL1 = Cells(tr, tc + 112).Value
        OP6_EGL2 = Cells(tr, tc + 113).Value
        OP6_EGL3 = Cells(tr, tc + 114).Value
        OP6_EGL4 = Cells(tr, tc + 115).Value
        OP6_KLS_BP = Cells(tr, tc + 116).Value
        OP6_KLS_LD = Cells(tr, tc + 117).Value
        OP6_KLS_RD = Cells(tr, tc + 118).Value
        OP6_KLS_LC = Cells(tr, tc + 119).Value
        OP6_KLS_RC = Cells(tr, tc + 120).Value
        OP6_KRS = Cells(tr, tc + 121).Value
        OP6_AMP = Cells(tr, tc + 122).Value
        OP6_KVS = Cells(tr, tc + 123).Value
        OP6_OL = Cells(tr, tc + 124).Value
        OP6_OM = Cells(tr, tc + 125).Value
        OP6_OSFC = Cells(tr, tc + 126).Value
        OP6_OSFF = Cells(tr, tc + 127).Value
        OP6_DT = Cells(tr, tc + 128).Value
        
        PR1 = Cells(tr, tc + 129).Value
        PR2 = Cells(tr, tc + 130).Value
        PR3 = Cells(tr, tc + 131).Value
        PR4 = Cells(tr, tc + 132).Value
        PL1 = Cells(tr, tc + 133).Value
        PL2 = Cells(tr, tc + 134).Value
        PL3 = Cells(tr, tc + 135).Value
        PL4 = Cells(tr, tc + 136).Value
        OSC_Sync = Cells(tr, tc + 137).Value
        LFO_Speed = Cells(tr, tc + 138).Value
        LFO_Delay = Cells(tr, tc + 139).Value
        PMD = Cells(tr, tc + 140).Value
        AMD = Cells(tr, tc + 141).Value
        LFO_Sync = Cells(tr, tc + 142).Value
        LFO_Wave = Cells(tr, tc + 143).Value
        PMS = Cells(tr, tc + 144).Value
        TRS = Cells(tr, tc + 145).Value
        OPR_S = Cells(tr, tc + 146).Value
        
        If DataSet = 1 Then
        
        '1Voiceデータ　7バイト目から163バイト目まで
            data_str = data_str & " " & HEX2(OP6_EGR1)
            data_str = data_str & " " & HEX2(OP6_EGR2)
            data_str = data_str & " " & HEX2(OP6_EGR3)
            data_str = data_str & " " & HEX2(OP6_EGR4)
            data_str = data_str & " " & HEX2(OP6_EGL1)
            data_str = data_str & " " & HEX2(OP6_EGL2)
            data_str = data_str & " " & HEX2(OP6_EGL3)
            data_str = data_str & " " & HEX2(OP6_EGL4)
            data_str = data_str & " " & HEX2(OP6_KLS_BP)
            data_str = data_str & " " & HEX2(OP6_KLS_LD)
            data_str = data_str & " " & HEX2(OP6_KLS_RD)
            data_str = data_str & " " & HEX2(OP6_KLS_LC)
            data_str = data_str & " " & HEX2(OP6_KLS_RC)
            data_str = data_str & " " & HEX2(OP6_KRS)
            data_str = data_str & " " & HEX2(OP6_AMP)
            data_str = data_str & " " & HEX2(OP6_KVS)
            data_str = data_str & " " & HEX2(OP6_OL)
            data_str = data_str & " " & HEX2(OP6_OM)
            data_str = data_str & " " & HEX2(OP6_OSFC)
            data_str = data_str & " " & HEX2(OP6_OSFF)
            data_str = data_str & " " & HEX2(OP6_DT + 7)
            
            data_str = data_str & " " & HEX2(OP5_EGR1)
            data_str = data_str & " " & HEX2(OP5_EGR2)
            data_str = data_str & " " & HEX2(OP5_EGR3)
            data_str = data_str & " " & HEX2(OP5_EGR4)
            data_str = data_str & " " & HEX2(OP5_EGL1)
            data_str = data_str & " " & HEX2(OP5_EGL2)
            data_str = data_str & " " & HEX2(OP5_EGL3)
            data_str = data_str & " " & HEX2(OP5_EGL4)
            data_str = data_str & " " & HEX2(OP5_KLS_BP)
            data_str = data_str & " " & HEX2(OP5_KLS_LD)
            data_str = data_str & " " & HEX2(OP5_KLS_RD)
            data_str = data_str & " " & HEX2(OP5_KLS_LC)
            data_str = data_str & " " & HEX2(OP5_KLS_RC)
            data_str = data_str & " " & HEX2(OP5_KRS)
            data_str = data_str & " " & HEX2(OP5_AMP)
            data_str = data_str & " " & HEX2(OP5_KVS)
            data_str = data_str & " " & HEX2(OP5_OL)
            data_str = data_str & " " & HEX2(OP5_OM)
            data_str = data_str & " " & HEX2(OP5_OSFC)
            data_str = data_str & " " & HEX2(OP5_OSFF)
            data_str = data_str & " " & HEX2(OP5_DT + 7)

            data_str = data_str & " " & HEX2(OP4_EGR1)
            data_str = data_str & " " & HEX2(OP4_EGR2)
            data_str = data_str & " " & HEX2(OP4_EGR3)
            data_str = data_str & " " & HEX2(OP4_EGR4)
            data_str = data_str & " " & HEX2(OP4_EGL1)
            data_str = data_str & " " & HEX2(OP4_EGL2)
            data_str = data_str & " " & HEX2(OP4_EGL3)
            data_str = data_str & " " & HEX2(OP4_EGL4)
            data_str = data_str & " " & HEX2(OP4_KLS_BP)
            data_str = data_str & " " & HEX2(OP4_KLS_LD)
            data_str = data_str & " " & HEX2(OP4_KLS_RD)
            data_str = data_str & " " & HEX2(OP4_KLS_LC)
            data_str = data_str & " " & HEX2(OP4_KLS_RC)
            data_str = data_str & " " & HEX2(OP4_KRS)
            data_str = data_str & " " & HEX2(OP4_AMP)
            data_str = data_str & " " & HEX2(OP4_KVS)
            data_str = data_str & " " & HEX2(OP4_OL)
            data_str = data_str & " " & HEX2(OP4_OM)
            data_str = data_str & " " & HEX2(OP4_OSFC)
            data_str = data_str & " " & HEX2(OP4_OSFF)
            data_str = data_str & " " & HEX2(OP4_DT + 7)

            data_str = data_str & " " & HEX2(OP3_EGR1)
            data_str = data_str & " " & HEX2(OP3_EGR2)
            data_str = data_str & " " & HEX2(OP3_EGR3)
            data_str = data_str & " " & HEX2(OP3_EGR4)
            data_str = data_str & " " & HEX2(OP3_EGL1)
            data_str = data_str & " " & HEX2(OP3_EGL2)
            data_str = data_str & " " & HEX2(OP3_EGL3)
            data_str = data_str & " " & HEX2(OP3_EGL4)
            data_str = data_str & " " & HEX2(OP3_KLS_BP)
            data_str = data_str & " " & HEX2(OP3_KLS_LD)
            data_str = data_str & " " & HEX2(OP3_KLS_RD)
            data_str = data_str & " " & HEX2(OP3_KLS_LC)
            data_str = data_str & " " & HEX2(OP3_KLS_RC)
            data_str = data_str & " " & HEX2(OP3_KRS)
            data_str = data_str & " " & HEX2(OP3_AMP)
            data_str = data_str & " " & HEX2(OP3_KVS)
            data_str = data_str & " " & HEX2(OP3_OL)
            data_str = data_str & " " & HEX2(OP3_OM)
            data_str = data_str & " " & HEX2(OP3_OSFC)
            data_str = data_str & " " & HEX2(OP3_OSFF)
            data_str = data_str & " " & HEX2(OP3_DT + 7)

            data_str = data_str & " " & HEX2(OP2_EGR1)
            data_str = data_str & " " & HEX2(OP2_EGR2)
            data_str = data_str & " " & HEX2(OP2_EGR3)
            data_str = data_str & " " & HEX2(OP2_EGR4)
            data_str = data_str & " " & HEX2(OP2_EGL1)
            data_str = data_str & " " & HEX2(OP2_EGL2)
            data_str = data_str & " " & HEX2(OP2_EGL3)
            data_str = data_str & " " & HEX2(OP2_EGL4)
            data_str = data_str & " " & HEX2(OP2_KLS_BP)
            data_str = data_str & " " & HEX2(OP2_KLS_LD)
            data_str = data_str & " " & HEX2(OP2_KLS_RD)
            data_str = data_str & " " & HEX2(OP2_KLS_LC)
            data_str = data_str & " " & HEX2(OP2_KLS_RC)
            data_str = data_str & " " & HEX2(OP2_KRS)
            data_str = data_str & " " & HEX2(OP2_AMP)
            data_str = data_str & " " & HEX2(OP2_KVS)
            data_str = data_str & " " & HEX2(OP2_OL)
            data_str = data_str & " " & HEX2(OP2_OM)
            data_str = data_str & " " & HEX2(OP2_OSFC)
            data_str = data_str & " " & HEX2(OP2_OSFF)
            data_str = data_str & " " & HEX2(OP2_DT + 7)
            
            data_str = data_str & " " & HEX2(OP1_EGR1)
            data_str = data_str & " " & HEX2(OP1_EGR2)
            data_str = data_str & " " & HEX2(OP1_EGR3)
            data_str = data_str & " " & HEX2(OP1_EGR4)
            data_str = data_str & " " & HEX2(OP1_EGL1)
            data_str = data_str & " " & HEX2(OP1_EGL2)
            data_str = data_str & " " & HEX2(OP1_EGL3)
            data_str = data_str & " " & HEX2(OP1_EGL4)
            data_str = data_str & " " & HEX2(OP1_KLS_BP)
            data_str = data_str & " " & HEX2(OP1_KLS_LD)
            data_str = data_str & " " & HEX2(OP1_KLS_RD)
            data_str = data_str & " " & HEX2(OP1_KLS_LC)
            data_str = data_str & " " & HEX2(OP1_KLS_RC)
            data_str = data_str & " " & HEX2(OP1_KRS)
            data_str = data_str & " " & HEX2(OP1_AMP)
            data_str = data_str & " " & HEX2(OP1_KVS)
            data_str = data_str & " " & HEX2(OP1_OL)
            data_str = data_str & " " & HEX2(OP1_OM)
            data_str = data_str & " " & HEX2(OP1_OSFC)
            data_str = data_str & " " & HEX2(OP1_OSFF)
            data_str = data_str & " " & HEX2(OP1_DT + 7)
            
            'PR、PL（8バイト）
            data_str = data_str & " " & HEX2(PR1)
            data_str = data_str & " " & HEX2(PR2)
            data_str = data_str & " " & HEX2(PR3)
            data_str = data_str & " " & HEX2(PR4)
            data_str = data_str & " " & HEX2(PL1)
            data_str = data_str & " " & HEX2(PL2)
            data_str = data_str & " " & HEX2(PL3)
            data_str = data_str & " " & HEX2(PL4)

            data_str = data_str & " " & HEX2(ALG - 1)
            data_str = data_str & " " & HEX2(FB)
            data_str = data_str & " " & HEX2(OSC_Sync)
            data_str = data_str & " " & HEX2(LFO_Speed)
            data_str = data_str & " " & HEX2(LFO_Delay)
            data_str = data_str & " " & HEX2(PMD)
            data_str = data_str & " " & HEX2(AMD)
            data_str = data_str & " " & HEX2(LFO_Sync)
            data_str = data_str & " " & HEX2(LFO_Wave)
            data_str = data_str & " " & HEX2(PMS)
            data_str = data_str & " " & HEX2(TRS)
            
            'VoiceNameを付加（10バイト）
            data_str = data_str & " " & Conv_Text(Left(VoiceName, 10))
            
            data_str = data_str & " " & HEX2(OPR_S)
        
        ElseIf DataSet = 32 Then
        
        '32Voiceバルクデータ（1Voice 128バイト×32Voice=4096バイト）　1Voiceの1バイト目から57バイト目までのデータをdata_strに入れる
        
            'OP6
            data_str = data_str & " " & HEX2(OP6_EGR1)
            data_str = data_str & " " & HEX2(OP6_EGR2)
            data_str = data_str & " " & HEX2(OP6_EGR3)
            data_str = data_str & " " & HEX2(OP6_EGR4)
            data_str = data_str & " " & HEX2(OP6_EGL1)
            data_str = data_str & " " & HEX2(OP6_EGL2)
            data_str = data_str & " " & HEX2(OP6_EGL3)
            data_str = data_str & " " & HEX2(OP6_EGL4)
            data_str = data_str & " " & HEX2(OP6_KLS_BP)
            data_str = data_str & " " & HEX2(OP6_KLS_LD)
            data_str = data_str & " " & HEX2(OP6_KLS_RD)
            
            'byte             bit #
            ' #     6   5   4   3   2   1   0   param A       range  param B       range
            '----  --- --- --- --- --- --- ---  ------------  -----  ------------  -----
            '11    0   0   0 |  RC   |   LC  | SCL LEFT CURVE 0-3   SCL RGHT CURVE 0-3
            data_str = data_str & " " & HEX2(BinToDec(DecToBin(OP6_KLS_RC, 2) & DecToBin(OP6_KLS_LC, 2)))
            
            '12  |      DET      |     RS    | OSC DETUNE     0-14  KBD RATE SCALE 0-7
            data_str = data_str & " " & HEX2(BinToDec(DecToBin(OP6_DT + 7, 4) & DecToBin((OP6_KRS), 3)))

            '13    0   0 |    KVS    |  AMS  | KEY VEL SENS   0-7   AMP MOD SENS   0-3
            data_str = data_str & " " & HEX2(BinToDec(DecToBin(OP6_KVS, 3) & DecToBin((OP6_AMP), 2)))
            
            data_str = data_str & " " & HEX2(OP6_OL)
            
            '15    0 |         FC        | M | FREQ COARSE    0-31  OSC MODE       0-1
            data_str = data_str & " " & HEX2(BinToDec(DecToBin(OP6_OSFC, 5) & DecToBin((OP6_OM), 1)))
            
            data_str = data_str & " " & HEX2(OP6_OSFF)
            
            'OP5
            data_str = data_str & " " & HEX2(OP5_EGR1)
            data_str = data_str & " " & HEX2(OP5_EGR2)
            data_str = data_str & " " & HEX2(OP5_EGR3)
            data_str = data_str & " " & HEX2(OP5_EGR4)
            data_str = data_str & " " & HEX2(OP5_EGL1)
            data_str = data_str & " " & HEX2(OP5_EGL2)
            data_str = data_str & " " & HEX2(OP5_EGL3)
            data_str = data_str & " " & HEX2(OP5_EGL4)
            data_str = data_str & " " & HEX2(OP5_KLS_BP)
            data_str = data_str & " " & HEX2(OP5_KLS_LD)
            data_str = data_str & " " & HEX2(OP5_KLS_RD)
            
            'byte             bit #
            ' #     6   5   4   3   2   1   0   param A       range  param B       range
            '----  --- --- --- --- --- --- ---  ------------  -----  ------------  -----
            '28    0   0   0 |  RC   |   LC  | SCL LEFT CURVE 0-3   SCL RGHT CURVE 0-3
            data_str = data_str & " " & HEX2(BinToDec(DecToBin(OP5_KLS_RC, 2) & DecToBin(OP5_KLS_LC, 2)))
            
            '29  |      DET      |     RS    | OSC DETUNE     0-14  KBD RATE SCALE 0-7
            data_str = data_str & " " & HEX2(BinToDec(DecToBin(OP5_DT + 7, 4) & DecToBin((OP5_KRS), 3)))

            '30    0   0 |    KVS    |  AMS  | KEY VEL SENS   0-7   AMP MOD SENS   0-3
            data_str = data_str & " " & HEX2(BinToDec(DecToBin(OP5_KVS, 3) & DecToBin((OP5_AMP), 2)))
            
            data_str = data_str & " " & HEX2(OP5_OL)
            
            '32    0 |         FC        | M | FREQ COARSE    0-31  OSC MODE       0-1
            data_str = data_str & " " & HEX2(BinToDec(DecToBin(OP5_OSFC, 5) & DecToBin((OP5_OM), 1)))
            
            data_str = data_str & " " & HEX2(OP5_OSFF)

            'OP4
            data_str = data_str & " " & HEX2(OP4_EGR1)
            data_str = data_str & " " & HEX2(OP4_EGR2)
            data_str = data_str & " " & HEX2(OP4_EGR3)
            data_str = data_str & " " & HEX2(OP4_EGR4)
            data_str = data_str & " " & HEX2(OP4_EGL1)
            data_str = data_str & " " & HEX2(OP4_EGL2)
            data_str = data_str & " " & HEX2(OP4_EGL3)
            data_str = data_str & " " & HEX2(OP4_EGL4)
            data_str = data_str & " " & HEX2(OP4_KLS_BP)
            data_str = data_str & " " & HEX2(OP4_KLS_LD)
            data_str = data_str & " " & HEX2(OP4_KLS_RD)
            
            'byte             bit #
            ' #     6   5   4   3   2   1   0   param A       range  param B       range
            '----  --- --- --- --- --- --- ---  ------------  -----  ------------  -----
            '45    0   0   0 |  RC   |   LC  | SCL LEFT CURVE 0-3   SCL RGHT CURVE 0-3
            data_str = data_str & " " & HEX2(BinToDec(DecToBin(OP4_KLS_RC, 2) & DecToBin(OP4_KLS_LC, 2)))
            
            '46  |      DET      |     RS    | OSC DETUNE     0-14  KBD RATE SCALE 0-7
            data_str = data_str & " " & HEX2(BinToDec(DecToBin(OP4_DT + 7, 4) & DecToBin((OP4_KRS), 3)))

            '47    0   0 |    KVS    |  AMS  | KEY VEL SENS   0-7   AMP MOD SENS   0-3
            data_str = data_str & " " & HEX2(BinToDec(DecToBin(OP4_KVS, 3) & DecToBin((OP4_AMP), 2)))
            
            data_str = data_str & " " & HEX2(OP4_OL)
            
            '49    0 |         FC        | M | FREQ COARSE    0-31  OSC MODE       0-1
            data_str = data_str & " " & HEX2(BinToDec(DecToBin(OP4_OSFC, 5) & DecToBin((OP4_OM), 1)))
            
            data_str = data_str & " " & HEX2(OP4_OSFF)


            'OP3
            data_str = data_str & " " & HEX2(OP3_EGR1)
            data_str = data_str & " " & HEX2(OP3_EGR2)
            data_str = data_str & " " & HEX2(OP3_EGR3)
            data_str = data_str & " " & HEX2(OP3_EGR4)
            data_str = data_str & " " & HEX2(OP3_EGL1)
            data_str = data_str & " " & HEX2(OP3_EGL2)
            data_str = data_str & " " & HEX2(OP3_EGL3)
            data_str = data_str & " " & HEX2(OP3_EGL4)
            data_str = data_str & " " & HEX2(OP3_KLS_BP)
            data_str = data_str & " " & HEX2(OP3_KLS_LD)
            data_str = data_str & " " & HEX2(OP3_KLS_RD)
            
            'byte             bit #
            ' #     6   5   4   3   2   1   0   param A       range  param B       range
            '----  --- --- --- --- --- --- ---  ------------  -----  ------------  -----
            '62    0   0   0 |  RC   |   LC  | SCL LEFT CURVE 0-3   SCL RGHT CURVE 0-3
            data_str = data_str & " " & HEX2(BinToDec(DecToBin(OP3_KLS_RC, 2) & DecToBin(OP3_KLS_LC, 2)))
            
            '63  |      DET      |     RS    | OSC DETUNE     0-14  KBD RATE SCALE 0-7
            data_str = data_str & " " & HEX2(BinToDec(DecToBin(OP3_DT + 7, 4) & DecToBin((OP3_KRS), 3)))

            '64    0   0 |    KVS    |  AMS  | KEY VEL SENS   0-7   AMP MOD SENS   0-3
            data_str = data_str & " " & HEX2(BinToDec(DecToBin(OP3_KVS, 3) & DecToBin((OP3_AMP), 2)))
            
            data_str = data_str & " " & HEX2(OP3_OL)
            
            '66    0 |         FC        | M | FREQ COARSE    0-31  OSC MODE       0-1
            data_str = data_str & " " & HEX2(BinToDec(DecToBin(OP3_OSFC, 5) & DecToBin((OP3_OM), 1)))
            
            data_str = data_str & " " & HEX2(OP3_OSFF)

            'OP2
            data_str = data_str & " " & HEX2(OP2_EGR1)
            data_str = data_str & " " & HEX2(OP2_EGR2)
            data_str = data_str & " " & HEX2(OP2_EGR3)
            data_str = data_str & " " & HEX2(OP2_EGR4)
            data_str = data_str & " " & HEX2(OP2_EGL1)
            data_str = data_str & " " & HEX2(OP2_EGL2)
            data_str = data_str & " " & HEX2(OP2_EGL3)
            data_str = data_str & " " & HEX2(OP2_EGL4)
            data_str = data_str & " " & HEX2(OP2_KLS_BP)
            data_str = data_str & " " & HEX2(OP2_KLS_LD)
            data_str = data_str & " " & HEX2(OP2_KLS_RD)
            
            'byte             bit #
            ' #     6   5   4   3   2   1   0   param A       range  param B       range
            '----  --- --- --- --- --- --- ---  ------------  -----  ------------  -----
            '79    0   0   0 |  RC   |   LC  | SCL LEFT CURVE 0-3   SCL RGHT CURVE 0-3
            data_str = data_str & " " & HEX2(BinToDec(DecToBin(OP2_KLS_RC, 2) & DecToBin(OP2_KLS_LC, 2)))
            
            '80  |      DET      |     RS    | OSC DETUNE     0-14  KBD RATE SCALE 0-7
            data_str = data_str & " " & HEX2(BinToDec(DecToBin(OP2_DT + 7, 4) & DecToBin((OP2_KRS), 3)))

            '81    0   0 |    KVS    |  AMS  | KEY VEL SENS   0-7   AMP MOD SENS   0-3
            data_str = data_str & " " & HEX2(BinToDec(DecToBin(OP2_KVS, 3) & DecToBin((OP2_AMP), 2)))
            
            data_str = data_str & " " & HEX2(OP2_OL)
            
            '83    0 |         FC        | M | FREQ COARSE    0-31  OSC MODE       0-1
            data_str = data_str & " " & HEX2(BinToDec(DecToBin(OP2_OSFC, 5) & DecToBin((OP2_OM), 1)))
            
            data_str = data_str & " " & HEX2(OP2_OSFF)

            'OP1
            data_str = data_str & " " & HEX2(OP1_EGR1)
            data_str = data_str & " " & HEX2(OP1_EGR2)
            data_str = data_str & " " & HEX2(OP1_EGR3)
            data_str = data_str & " " & HEX2(OP1_EGR4)
            data_str = data_str & " " & HEX2(OP1_EGL1)
            data_str = data_str & " " & HEX2(OP1_EGL2)
            data_str = data_str & " " & HEX2(OP1_EGL3)
            data_str = data_str & " " & HEX2(OP1_EGL4)
            data_str = data_str & " " & HEX2(OP1_KLS_BP)
            data_str = data_str & " " & HEX2(OP1_KLS_LD)
            data_str = data_str & " " & HEX2(OP1_KLS_RD)
            
            'byte             bit #
            ' #     6   5   4   3   2   1   0   param A       range  param B       range
            '----  --- --- --- --- --- --- ---  ------------  -----  ------------  -----
            '96    0   0   0 |  RC   |   LC  | SCL LEFT CURVE 0-3   SCL RGHT CURVE 0-3
            data_str = data_str & " " & HEX2(BinToDec(DecToBin(OP1_KLS_RC, 2) & DecToBin(OP1_KLS_LC, 2)))
            
            '97  |      DET      |     RS    | OSC DETUNE     0-14  KBD RATE SCALE 0-7
            data_str = data_str & " " & HEX2(BinToDec(DecToBin(OP1_DT + 7, 4) & DecToBin((OP1_KRS), 3)))

            '98    0   0 |    KVS    |  AMS  | KEY VEL SENS   0-7   AMP MOD SENS   0-3
            data_str = data_str & " " & HEX2(BinToDec(DecToBin(OP1_KVS, 3) & DecToBin((OP1_AMP), 2)))
            
            data_str = data_str & " " & HEX2(OP1_OL)
            
            '100    0 |         FC        | M | FREQ COARSE    0-31  OSC MODE       0-1
            data_str = data_str & " " & HEX2(BinToDec(DecToBin(OP1_OSFC, 5) & DecToBin((OP1_OM), 1)))
            
            data_str = data_str & " " & HEX2(OP1_OSFF)

            'PR、PLを付加（8バイト）
            data_str = data_str & " " & HEX2(PR1)
            data_str = data_str & " " & HEX2(PR2)
            data_str = data_str & " " & HEX2(PR3)
            data_str = data_str & " " & HEX2(PR4)
            data_str = data_str & " " & HEX2(PL1)
            data_str = data_str & " " & HEX2(PL2)
            data_str = data_str & " " & HEX2(PL3)
            data_str = data_str & " " & HEX2(PL4)

            '110    0   0 |        ALG        | ALGORITHM     0-31
            data_str = data_str & " " & HEX2(BinToDec(DecToBin(ALG - 1, 5)))
            
            '111    0   0   0 |OKS|    FB     | OSC KEY SYNC  0-1    FEEDBACK      0-7
            data_str = data_str & " " & HEX2(BinToDec(DecToBin(OSC_Sync, 1) & DecToBin((FB), 3)))

            data_str = data_str & " " & HEX2(LFO_Speed)
            data_str = data_str & " " & HEX2(LFO_Delay)
            data_str = data_str & " " & HEX2(PMD)
            data_str = data_str & " " & HEX2(AMD)
            
            '116  |  LPMS |      LFW      |LKS| LF PT MOD SNS 0-7   WAVE 0-5,  SYNC 0-1
            data_str = data_str & " " & HEX2(BinToDec(DecToBin(PMS, 3) & DecToBin((LFO_Wave), 3)) & DecToBin((LFO_Sync), 1))
            
            data_str = data_str & " " & HEX2(TRS)
            
            'VoiceNameを付加（10バイト）
            data_str = data_str & " " & Conv_Text(Left(VoiceName, 10))
            
           
        End If
        
        'カウンターの更新
        sr = sr + 1
        tr = tr + 1
        
    Next c

    'チェックサムの計算
    chksum_data = Split(data_str, " ")
    ReDim chksum_byte(UBound(chksum_data))
    
    For k = 6 To UBound(chksum_data)
        chksum_byte(k - 6) = "&h" & chksum_data(k)
    Next k
        
    For j = 0 To UBound(chksum_byte)
        chksum = chksum + Hex("&H" & chksum_byte(j))
    Next j
    
    chksum = 128 - (chksum Mod 128)
    
    'チェックサムデータを付加
    data_str = data_str & " " & HEX2(chksum)
    
    'Sysex終了コードのF7を付加
    data_str = data_str & " " & "F7"
        
    'data_strを分割して配列VoiceData_strに入れる
    VoiceData_str = Split(data_str, " ")
    
    '配列VoiceData_strのデータをByteデータに変換するため配列VoiceData_bytに入れる
    ReDim VoiceData_byt(UBound(VoiceData_str))

    For i = 0 To UBound(VoiceData_str)
        VoiceData_byt(i) = "&h" & VoiceData_str(i)
    Next i
    
    'バイナリーファイルに書き込む
    fh = FreeFile
    Open filename For Binary Access Write As #fh
    Put #fh, , VoiceData_byt
    Close #fh
       
End Sub



