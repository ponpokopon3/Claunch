On Error Resume Next

Dim installPath
Dim addInName
Dim addInFileName
Dim objExcel
Dim objAddin

'// �A�h�C������ݒ�
addInName = "��Ȑ�"
addInFileName = "��Ȑ�.xlam"

If MsgBox(addInName & " �A�h�C�����A���C���X�g�[�����܂����H", vbYesNo + vbQuestion,addInName) = vbNo Then
  WScript.Quit
End If

'// �A�h�C���o�^����
Set objExcel = CreateObject("Excel.Application")
objExcel.Workbooks.Add
For i = 1 To objExcel.AddIns.Count
  Set objAddin = objExcel.AddIns.Item(i)
  If objAddin.Name = addInFileName Then
    objAddin.Installed = False
  End If
Next

'// Excel �I��
objExcel.Quit

Set objAddin = Nothing
Set objExcel = Nothing

Set objWshShell = CreateObject("WScript.Shell")
Set objFileSys = CreateObject("Scripting.FileSystemObject")

'// �C���X�g�[����p�X�̍쐬
'// (ex)C:\Users\[User]\AppData\Roaming\Microsoft\AddIns\[addInFileName]
installPath = objWshShell.SpecialFolders("Appdata") & "\Microsoft\Addins\" & addInFileName

'// �t�@�C���폜
If objFileSys.FileExists(installPath) = True Then
  objFileSys.DeleteFile installPath, True
Else
  MsgBox "�A�h�C���t�@�C���̓C���X�g�[������Ă��܂���B" & Chr(10) & "�������I�����܂��B", vbExclamation,addInName
  Set objWshShell = Nothing
  Set objFileSys = Nothing
  WScript.Quit
End If


If Err.Number = 0 Then
   MsgBox "�A�h�C���͐���ɃA���C���X�g�[������܂����B", vbInformation,addInName
Else
   MsgBox "�G���[���������܂����B" & vbCrLf & "���s�����m�F���Ă��������B", vbExclamation,addInName
End If

Set objWshShell = Nothing
Set objFileSys = Nothing

