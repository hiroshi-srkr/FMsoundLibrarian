Attribute VB_Name = "Module5"

'********************************************
' 10進数から2進数へ変換する関数
'********************************************
Function DecToBin(Number, digit As Integer) As String

Dim myYard As Long
Dim myNumber As Long
Dim myExpo As Long

    myNumber = Number
    While 2 ^ myYard <= myNumber
        myYard = myYard + 1
    Wend
    
    For myExpo = myYard - 1 To 0 Step -1
        If myNumber >= 2 ^ myExpo Then
            DecToBin = DecToBin & "1"
            myNumber = myNumber - 2 ^ myExpo
        Else
            DecToBin = DecToBin & "0"
        End If
    Next
    
    DecToBin = Right("00000000" & DecToBin, digit)

End Function

'********************************************
' 2進数から10進数へ変換する関数
'********************************************
Function BinToDec(Binary As String) As Long

Dim myLen As Integer
Dim i As Integer

    myLen = Len(Binary)
    For i = 1 To myLen
        If Mid(Binary, i, 1) = "1" Then
            BinToDec = BinToDec + 2 ^ (myLen - i)
        End If
    Next

End Function

'********************************************
' 10進数から2進数へ変換する関数サンプル
'********************************************
Function DecToBin_sample(Number) As String

Dim binCnt As Long
Dim DeciNum As Long 'Decimal Number
Dim i As Long

    DeciNum = Number
    
    '累乗確認
    binCnt = 0
    While 2 ^ binCnt <= DeciNum ' <=(以下の場合)
        binCnt = binCnt + 1
    Wend
    
    '2進数作成
    For i = binCnt - 1 To 0 Step -1
        If DeciNum >= 2 ^ i Then
            DecToBin = ConvToBin & "1"
            DeciNum = DeciNum - 2 ^ i
        Else
            DecToBin = ConvToBin & "0"
        End If
    Next

End Function

'***********************************************
' 文字列の前を指定の長さまでスペースで埋める関数
'***********************************************
Function StringLen(str, sLen) As String
    Dim c As Long
    c = sLen - Len(str)
    If c < 0 Then
        StringLen = ""
    Else
        StringLen = space(c) & str
    End If
End Function

'******************************
' 指定の列の最終行を求める関数
'******************************
Function getLastRow(WS As Worksheet, Optional CheckCol As Long = 1)
 With WS
  getLastRow = 0
  
  If Not Intersect(.UsedRange, .Columns(CheckCol)) Is Nothing Then
   Dim LastRow As Long
   LastRow = .UsedRange.Row + .UsedRange.Rows.Count - 1
   
   If LastRow > 1 Then
    Dim buf As Variant
    buf = .Range(.Cells(1, CheckCol), .Cells(LastRow, CheckCol)).Value
    
    Dim c As Long
    For c = UBound(buf, 1) To 1 Step -1
     If Not IsEmpty(buf(c, 1)) Then
      getLastRow = c
      Exit Function
     End If
    Next
   
   ElseIf Not IsEmpty(.Cells(1, CheckCol).Value) Then
     getLastRow = 1
   End If
  
  End If
 End With
End Function
