On Error Resume Next

Dim installPath
Dim addInName
Dim addInFileName
Dim objExcel
Dim objAddin

'// アドイン情報を設定
addInName = "稲妻線"
addInFileName = "稲妻線.xlam"

If MsgBox(addInName & " アドインをインストールしますか？", vbYesNo + vbQuestion,addInName) = vbNo Then
  WScript.Quit
End If

Set objWshShell = CreateObject("WScript.Shell")
Set objFileSys = CreateObject("Scripting.FileSystemObject")

'// インストール先パスの作成
'// (ex)C:\Users\[User]\AppData\Roaming\Microsoft\AddIns\[addInFileName]
installPath = objWshShell.SpecialFolders("Appdata") & "\Microsoft\Addins\" & addInFileName

'// ファイルコピー(上書き)
objFileSys.CopyFile addInFileName, installPath, True

Set objWshShell = Nothing
Set objFileSys = Nothing

'// アドイン登録
Set objExcel = CreateObject("Excel.Application")
objExcel.Workbooks.Add
Set objAddin = objExcel.AddIns.Add(installPath, True)
objAddin.Installed = True

If Err.Number = 0 Then
   MsgBox "アドインは正常にインストールされました。", vbInformation,addInName
Else
   MsgBox "エラーが発生しました。" & vbCrLf & "実行環境を確認してください。", vbExclamation,addInName
End If

'// Excel 終了
objExcel.Quit

Set objAddin = Nothing
Set objExcel = Nothing

