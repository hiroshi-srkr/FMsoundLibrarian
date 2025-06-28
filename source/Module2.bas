Attribute VB_Name = "Module2"

'***************************************************
'�@FrequencyRatio�̕ϊ��iOPM->DX21)
'***************************************************
Function Conv_Freq_Ratio(FR)

    Select Case FR
        Case 0.5
            Conv_Freq_Ratio = 0
        Case 1
            Conv_Freq_Ratio = 4
        Case 2
            Conv_Freq_Ratio = 8
        Case 3
            Conv_Freq_Ratio = 10
        Case 4
            Conv_Freq_Ratio = 13
        Case 5
            Conv_Freq_Ratio = 16
        Case 6
            Conv_Freq_Ratio = 19
        Case 7
            Conv_Freq_Ratio = 22
        Case 8
            Conv_Freq_Ratio = 25
        Case 9
            Conv_Freq_Ratio = 28
        Case 10
            Conv_Freq_Ratio = 31
        Case 11
            Conv_Freq_Ratio = 34
        Case 12
            Conv_Freq_Ratio = 36
        Case 13
            Conv_Freq_Ratio = 40
        Case 14
            Conv_Freq_Ratio = 42
        Case 15
            Conv_Freq_Ratio = 45
    End Select

End Function
'***************************************************
'�@16�i����2���\���ɕϊ��iHex�֐����p�j
'***************************************************
Function HEX2(n)
    HEX2 = Right("00" & Hex(n), 2)
End Function
'***************************************************
'�@�e�L�X�g��16�i���ɕϊ�
'***************************************************
Function Conv_Text(s)

Dim i, j As Integer
Dim char As String
Dim code As String

code = ""
' 1 �������擾����
For i = 1 To Len(s)
    char = Mid(s, i, 1)
    If i = 1 Then
        code = code & HEX2(Asc(char))
    Else
        code = code & " " & HEX2(Asc(char))
    End If
Next

'  10�����ȉ��̏ꍇ10�����ɂȂ�܂ŃX�y�[�X�Ŗ��߂�
If Len(s) < 10 Then
    For i = 1 To 10 - Len(s)
        code = code & " " & HEX2(Asc(" "))
    Next
End If

Conv_Text = code

End Function
'***************************************************
'�@�e�L�X�g��16�i���ɕϊ���16�i���̍��v�����߂�
'***************************************************
Function Conv_Text2(s)

Dim i As Integer
Dim char As String
Dim code As Integer

code = 0
' 1 �������擾����
For i = 1 To Len(s)
    char = Mid(s, i, 1)
    code = code + Asc(char)
Next

' 10�����ȉ��̏ꍇ10�����ɂȂ�܂ŃX�y�[�X�Ŗ��߂�
If Len(s) < 10 Then
    For i = 1 To 10 - Len(s)
        code = code + Asc(" ")
    Next
End If

Conv_Text2 = code

End Function
