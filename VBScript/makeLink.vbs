' Option Explicit

Const MSG_ERR_01 = "引数は2つ指定してください：対象ファイルのパス 作成先のパス"
Const MSG_ERR_02 = "存在しないパスが指定されています"
Const MSG_ERR_03 = "作成先にはフォルダを指定してください"
Const MSG_ERR_04 = "パスは同じドライブ内のものを指定してください"

Const MSG_01 = "対象ファイルに"
Const MSG_02 = "作成先に"

Const EXPLORER_PATH = "%windir%\explorer.exe"

' 引数が足りなければ処理を終了する
If CheckArguments() Then
    Dim targetFile   :targetFile   = WScript.Arguments(0)
    Dim directory   :directory   = WScript.Arguments(1)
Else
    WScript.Echo MSG_ERR_01
    WScript.Quit
End If

' 作成先のパスがフォルダでなければ処理を終了する
If CheckFileType(directory) = -1 Then
    WScript.Echo MSG_02 & MSG_ERR_02
    WScript.Quit
ElseIf CheckFileType(directory) = 1 Then
    WScript.Echo MSG_ERR_03
    WScript.Quit
End If

' 対象ファイルのパスが存在しなければ処理を終了する
If CheckFileType(targetFile) = -1 Then
    WScript.Echo  MSG_01 & MSG_ERR_02
    WScript.Quit
End If

' 2つのパスが同じドライブ上になければ処理を終了する
If CheckDrive(targetFile, directory) = False Then
    WScript.Echo MSG_ERR_04
    WScript.Quit
End If

' ショートカットを作成する
Call MakeShortcut(targetFile, directory)


' 引数のチェック
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
    ' 作成先を絶対パスにしておく
    directory = ConvertToFullPath(directory)

    ' ショートカット元の相対パスを取得
    Dim targetRelativePath: targetRelativePath = ConvertToRelativePath(directory, targetFile)
    ' ファイル名取得
    Set fso = createObject("Scripting.FileSystemObject")
    Dim targetFileName: targetFileName = fso.GetFileName(targetFile)
    Set fso = Nothing

    ' ショートカット作成
    Set ws = CreateObject("WScript.Shell")
    Set shortcut = ws.CreateShortcut(directory & "\" & targetFileName & "-ショートカット.lnk")
    With shortcut
        .TargetPath = EXPLORER_PATH
        .Arguments = """" & targetRelativePath & """"
        .Save
    End With
    WScript.Echo "作成しました：" & directory & "\" & targetFileName & "-ショートカット.lnk"
End Sub

' 相対パスへの変換
' param {string} base 基準となるパス
' param {string} target 相対パスに変換するパス
' return {string} 変換したパス（変換できない場合は空文字）
Function ConvertToRelativePath(base, target)
    Dim baseFull: baseFull = ConvertToFullPath(base)
    Dim targetFull: targetFull = ConvertToFullPath(target)

    ' パスを配列に変換して階層を比較する
    Dim baseArray: baseArray = Split(baseFull, "\")
    Dim targetArray: targetArray = Split(targetFull, "\")

    ' 短い方の配列をループの基準にする
    Dim cnt
    If UBound(baseArray) < UBound(targetArray) Then
        cnt = UBound(baseArray)
    Else
        cnt = UBound(targetArray)
    End If

    ' 2つのパスを先頭からチェック
    ' 異なるディレクトリを見つけてそこから相対パスを作成する
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

' 絶対パスへの変換
' param {string} path 変換するパス
' return {string} 変換したパス
Function ConvertToFullPath(path)
    Set fso = createObject("Scripting.FileSystemObject")
    path = fso.GetAbsolutePathName(path)
    Set fso = Nothing
    ConvertToFullPath = path
End Function

' フォルダ/ファイルを判定する
' param {string} path 判定対象のパス
' return {int} 0:フォルダ 1:ファイル -1:存在しないパス
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

' 2つのパスが同じドライブに存在するかチェックする
' param {string} path1 判定対象のパス1
' param {string} path2 判定対象のパス2
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