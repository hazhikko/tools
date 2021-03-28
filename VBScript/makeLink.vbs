' Option Explicit

Const MSG_ERR_01 = "������2�w�肵�Ă��������F�Ώۃt�@�C���̃p�X �쐬��̃p�X"
Const MSG_ERR_02 = "���݂��Ȃ��p�X���w�肳��Ă��܂�"
Const MSG_ERR_03 = "�쐬��ɂ̓t�H���_���w�肵�Ă�������"
Const MSG_ERR_04 = "�p�X�͓����h���C�u���̂��̂��w�肵�Ă�������"

Const MSG_01 = "�Ώۃt�@�C����"
Const MSG_02 = "�쐬���"

Const EXPLORER_PATH = "%windir%\explorer.exe"

' ����������Ȃ���Ώ������I������
If CheckArguments() Then
    Dim targetFile   :targetFile   = WScript.Arguments(0)
    Dim directory   :directory   = WScript.Arguments(1)
Else
    WScript.Echo MSG_ERR_01
    WScript.Quit
End If

' �쐬��̃p�X���t�H���_�łȂ���Ώ������I������
If CheckFileType(directory) = -1 Then
    WScript.Echo MSG_02 & MSG_ERR_02
    WScript.Quit
ElseIf CheckFileType(directory) = 1 Then
    WScript.Echo MSG_ERR_03
    WScript.Quit
End If

' �Ώۃt�@�C���̃p�X�����݂��Ȃ���Ώ������I������
If CheckFileType(targetFile) = -1 Then
    WScript.Echo  MSG_01 & MSG_ERR_02
    WScript.Quit
End If

' 2�̃p�X�������h���C�u��ɂȂ���Ώ������I������
If CheckDrive(targetFile, directory) = False Then
    WScript.Echo MSG_ERR_04
    WScript.Quit
End If

' �V���[�g�J�b�g���쐬����
Call MakeShortcut(targetFile, directory)


' �����̃`�F�b�N
' return {bool} True / False
Function CheckArguments()
    Dim result
    If Wscript.Arguments.Count <= 1 Then
        result = False
    Else
        result = True
    End If
    CheckArguments = result
End Function

Sub MakeShortcut(targetFile, directory)
    ' �쐬����΃p�X�ɂ��Ă���
    directory = ConvertToFullPath(directory)

    ' �V���[�g�J�b�g���̑��΃p�X���擾
    Dim targetRelativePath: targetRelativePath = ConvertToRelativePath(directory, targetFile)
    ' �t�@�C�����擾
    Set fso = createObject("Scripting.FileSystemObject")
    Dim targetFileName: targetFileName = fso.GetFileName(targetFile)
    Set fso = Nothing

    ' �V���[�g�J�b�g�쐬
    Set ws = CreateObject("WScript.Shell")
    Set shortcut = ws.CreateShortcut(directory & "\" & targetFileName & "-�V���[�g�J�b�g.lnk")
    With shortcut
        .TargetPath = EXPLORER_PATH
        .Arguments = """" & targetRelativePath & """"
        .Save
    End With
    WScript.Echo "�쐬���܂����F" & directory & "\" & targetFileName & "-�V���[�g�J�b�g.lnk"
End Sub

' ���΃p�X�ւ̕ϊ�
' param {string} base ��ƂȂ�p�X
' param {string} target ���΃p�X�ɕϊ�����p�X
' return {string} �ϊ������p�X�i�ϊ��ł��Ȃ��ꍇ�͋󕶎��j
Function ConvertToRelativePath(base, target)
    Dim baseFull: baseFull = ConvertToFullPath(base)
    Dim targetFull: targetFull = ConvertToFullPath(target)

    ' �p�X��z��ɕϊ����ĊK�w���r����
    Dim baseArray: baseArray = Split(baseFull, "\")
    Dim targetArray: targetArray = Split(targetFull, "\")

    ' �Z�����̔z������[�v�̊�ɂ���
    Dim cnt
    If UBound(baseArray) < UBound(targetArray) Then
        cnt = UBound(baseArray)
    Else
        cnt = UBound(targetArray)
    End If

    ' 2�̃p�X��擪����`�F�b�N
    ' �قȂ�f�B���N�g���������Ă������瑊�΃p�X���쐬����
    Dim relativePath: relativePath = "..\"
    For i = 0 To cnt
        If baseArray(i) <> targetArray(i) Then
            For j = i To UBound(baseArray) -1
                relativePath =  relativePath & "..\"
            Next
            Dim targetArray_()
            ReDim targetArray_(UBound(targetArray) - i)
            Dim k: k = 0
            For j = i To UBound(targetArray)
                targetArray_(k) = targetArray(j)
                k = k + 1
            Next
            relativePath = relativePath & join(targetArray_, "\")
            Exit For
        End If
    Next
    ConvertToRelativePath = relativePath
End Function

' ��΃p�X�ւ̕ϊ�
' param {string} path �ϊ�����p�X
' return {string} �ϊ������p�X
Function ConvertToFullPath(path)
    Set fso = createObject("Scripting.FileSystemObject")
    path = fso.GetAbsolutePathName(path)
    Set fso = Nothing
    ConvertToFullPath = path
End Function

' �t�H���_/�t�@�C���𔻒肷��
' param {string} path ����Ώۂ̃p�X
' return {int} 0:�t�H���_ 1:�t�@�C�� -1:���݂��Ȃ��p�X
Function CheckFileType(path)
    Dim fileType
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(path) Then
        fileType = 0
    ElseIf fso.FileExists(path) Then
        fileType = 1
    Else
        fileType = -1
    End If

    Set fso = Nothing
    Set oArgs = Nothing

    CheckFileType = fileType
End Function

' 2�̃p�X�������h���C�u�ɑ��݂��邩�`�F�b�N����
' param {string} path1 ����Ώۂ̃p�X1
' param {string} path2 ����Ώۂ̃p�X2
' return {bool} True / False
Function CheckDrive(path1, path2)
    Dim path1Full: path1Full = ConvertToFullPath(path1)
    Dim path1Drive: path1Drive = Left(path1Full, InStr(path1Full, ":"))
    Dim pathFull: pathFull = ConvertToFullPath(path2)
    Dim path2Drive: path2Drive = Left(pathFull, InStr(pathFull, ":"))

    If StrComp(path1Drive, path2Drive) = 0 Then
        CheckDrive = True
    Else
        CheckDrive = False
    End If
End Function