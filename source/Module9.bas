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
  aryDir(i) = argDir '�����̃t�H���_��z��̐擪�ɓ����
 
  '�܂��́A�w��t�H���_�ȉ��̑S�T�u�t�H���_���擾���A�z��aryDir�ɓ���܂��B
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
 
  '�z��aryDir�̑S�t�H���_�ɂ��āA�t�@�C�����擾���A�z��aryFile�ɓ���܂��B
  ReDim aryFile(0)
  For i = 0 To UBound(aryDir)
    strName = Dir(aryDir(i) & "\", vbNormal + vbHidden + vbReadOnly + vbSystem)
    Do While strName <> ""
      If aryFile(0) <> "" Then
        ReDim Preserve aryFile(UBound(aryFile) + 1)
      End If
      aryFile(UBound(aryFile)) = aryDir(i) & "\" & strName
      '���s���ʂ�������₷���悤�ɁA�e�X�g�I�ɃZ���ɏ����o���ꍇ
      'Cells(UBound(aryFile) + 1, 1) = aryFile(UBound(aryFile))
      'Debug.Print aryFile(UBound(aryFile))
      Debug.Print strName
      strName = Dir()
    Loop
  Next
 
  GetFileList = aryFile
End Function

'**********************************************************
' �t�H���_�[�I����A���̃t�H���_�[�̑SSysex�t�@�C���ǂݍ���
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

            '//�t�@�C�����𐶐�
    
            strFilePath = myFolder
            strFileName = filename

            If strFilePath = "" Then
                strFilePath = ThisWorkbook.Path
            End If
            
            If strFileName = "" Then
                MsgBox "�t�@�C�������w�肳��Ă��܂���B"
        
            Else
                '�t�@�C���V�X�e���������I�u�W�F�N�g���쐬
                Set objFileSys = CreateObject("Scripting.FileSystemObject")
                
                strFile = strFilePath & "\" & strFileName
        
                'If Dir(strFile) <> "" Then
                
                '//�o�C�i���t�@�C����ǂݍ��݁iSysex�t�@�C����ǂݍ��݁j
                    strLibName = objFileSys.GetBaseName(strFile)
                    Call ReadSysexFileDX7(strFile, strLibName)
                    Sheets("SysexDX7Data").Select
                    MsgBox "Sysex�f�[�^�̓ǂݍ��݂��������܂����B"
                    Set objFileSys = Nothing
                'Else
                    'MsgBox strFile & vbCrLf & _
                            "�����݂��܂���"
                'End If
            End If
    
            filename = Dir()
    
        Loop
                
    Else
    
        MsgBox "Sysex�t�@�C����������܂���"
                
    End If

End With

End Sub
