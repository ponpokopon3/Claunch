On Error Resume Next

Dim installPath
Dim addInName
Dim addInFileName
Dim objExcel
Dim objAddin

'// アドイン情報を設定
addInName = "稲妻線"
addInFileName = "稲妻線.xlam"

If MsgBox(addInName & " アドインをアンインストールしますか？", vbYesNo + vbQuestion,addInName) = vbNo Then
  WScript.Quit
End If

'// アドイン登録解除
Set objExcel = CreateObject("Excel.Application")
objExcel.Workbooks.Add
For i = 1 To objExcel.AddIns.Count
  Set objAddin = objExcel.AddIns.Item(i)
  If objAddin.Name = addInFileName Then
    objAddin.Installed = False
  End If
Next

'// Excel 終了
objExcel.Quit

Set objAddin = Nothing
Set objExcel = Nothing

Set objWshShell = CreateObject("WScript.Shell")
Set objFileSys = CreateObject("Scripting.FileSystemObject")

'// インストール先パスの作成
'// (ex)C:\Users\[User]\AppData\Roaming\Microsoft\AddIns\[addInFileName]
installPath = objWshShell.SpecialFolders("Appdata") & "\Microsoft\Addins\" & addInFileName

'// ファイル削除
If objFileSys.FileExists(installPath) = True Then
  objFileSys.DeleteFile installPath, True
Else
  MsgBox "アドインファイルはインストールされていません。" & Chr(10) & "処理を終了します。", vbExclamation,addInName
  Set objWshShell = Nothing
  Set objFileSys = Nothing
  WScript.Quit
End If


If Err.Number = 0 Then
   MsgBox "アドインは正常にアンインストールされました。", vbInformation,addInName
Else
   MsgBox "エラーが発生しました。" & vbCrLf & "実行環境を確認してください。", vbExclamation,addInName
End If

Set objWshShell = Nothing
Set objFileSys = Nothing

