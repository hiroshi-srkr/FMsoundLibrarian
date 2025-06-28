Attribute VB_Name = "Module9"
Sub Test()
    cnt = 0
    Call SearchDir("C:\temp")
End Sub

Sub test2()
    Dim v As Variant
    v = GetFileList("C:\temp")
End Sub


Sub SearchDir(Path As String)
    Dim buf As String, f As Object
    Dim cnt As Long
    
    buf = Dir(Path & "\*.*")
    Do While buf <> ""
        cnt = cnt + 1
        Cells(cnt, 1) = buf
        buf = Dir()
    Loop
    With CreateObject("Scripting.FileSystemObject")
        For Each f In .GetFolder(Path).SubFolders
            Debug.Print f.Path
            Call SearchDir(f.Path)
        Next f
    End With
End Sub

Function GetFileList(ByVal argDir As String) As String()
  Dim i As Long
  Dim aryDir() As String
  Dim aryFile() As String
  Dim strName As String

  ReDim aryDir(i)
  aryDir(i) = argDir '引数のフォルダを配列の先頭に入れる
 
  'まずは、指定フォルダ以下の全サブフォルダを取得し、配列aryDirに入れます。
  i = 0
  Do
    strName = Dir(aryDir(i) & "\", vbDirectory)
    Do While strName <> ""
      If GetAttr(aryDir(i) & "\" & strName) And vbDirectory Then
        If strName <> "." And strName <> ".." Then
          ReDim Preserve aryDir(UBound(aryDir) + 1)
          aryDir(UBound(aryDir)) = aryDir(i) & "\" & strName
        End If
      End If
      strName = Dir()
    Loop
    i = i + 1
    If i > UBound(aryDir) Then Exit Do
  Loop
 
  '配列aryDirの全フォルダについて、ファイルを取得し、配列aryFileに入れます。
  ReDim aryFile(0)
  For i = 0 To UBound(aryDir)
    strName = Dir(aryDir(i) & "\", vbNormal + vbHidden + vbReadOnly + vbSystem)
    Do While strName <> ""
      If aryFile(0) <> "" Then
        ReDim Preserve aryFile(UBound(aryFile) + 1)
      End If
      aryFile(UBound(aryFile)) = aryDir(i) & "\" & strName
      '実行結果が分かりやすいように、テスト的にセルに書き出す場合
      'Cells(UBound(aryFile) + 1, 1) = aryFile(UBound(aryFile))
      'Debug.Print aryFile(UBound(aryFile))
      Debug.Print strName
      strName = Dir()
    Loop
  Next
 
  GetFileList = aryFile
End Function

'**********************************************************
' フォルダー選択後、そのフォルダーの全Sysexファイル読み込み
'**********************************************************
Sub AllDataReadDX7()

Dim filename    As String
Dim IsBookOpen  As Boolean
Dim OpenBook    As Workbook
Dim myFolder As Variant
Dim strFilePath As String
Dim strFileName As String
Dim strFile As String
Dim objFileSys As Object
Dim strLibName As String
    
With Application.FileDialog(msoFileDialogFolderPicker)

    If .Show <> 0 Then
    
        myFolder = .SelectedItems(1)
    

        With CreateObject("WScript.Shell")
    
            .CurrentDirectory = myFolder
    
        End With

        filename = Dir("*.syx")
     
        Do While filename <> ""
            'MsgBox Filename

            '//ファイル名を生成
    
            strFilePath = myFolder
            strFileName = filename

            If strFilePath = "" Then
                strFilePath = ThisWorkbook.Path
            End If
            
            If strFileName = "" Then
                MsgBox "ファイル名が指定されていません。"
        
            Else
                'ファイルシステムを扱うオブジェクトを作成
                Set objFileSys = CreateObject("Scripting.FileSystemObject")
                
                strFile = strFilePath & "\" & strFileName
        
                'If Dir(strFile) <> "" Then
                
                '//バイナリファイルを読み込み（Sysexファイルを読み込み）
                    strLibName = objFileSys.GetBaseName(strFile)
                    Call ReadSysexFileDX7(strFile, strLibName)
                    Sheets("SysexDX7Data").Select
                    MsgBox "Sysexデータの読み込みが完了しました。"
                    Set objFileSys = Nothing
                'Else
                    'MsgBox strFile & vbCrLf & _
                            "が存在しません"
                'End If
            End If
    
            filename = Dir()
    
        Loop
                
    Else
    
        MsgBox "Sysexファイルが見つかりません"
                
    End If

End With

End Sub
