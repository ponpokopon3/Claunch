On Error Resume Next

Dim installPath
Dim addInName
Dim addInFileName
Dim objExcel
Dim objAddin

'// �A�h�C������ݒ�
addInName = "��Ȑ�"
addInFileName = "��Ȑ�.xlam"

If MsgBox(addInName & " �A�h�C�����C���X�g�[�����܂����H", vbYesNo + vbQuestion,addInName) = vbNo Then
  WScript.Quit
End If

Set objWshShell = CreateObject("WScript.Shell")
Set objFileSys = CreateObject("Scripting.FileSystemObject")

'// �C���X�g�[����p�X�̍쐬
'// (ex)C:\Users\[User]\AppData\Roaming\Microsoft\AddIns\[addInFileName]
installPath = objWshShell.SpecialFolders("Appdata") & "\Microsoft\Addins\" & addInFileName

'// �t�@�C���R�s�[(�㏑��)
objFileSys.CopyFile addInFileName, installPath, True

Set objWshShell = Nothing
Set objFileSys = Nothing

'// �A�h�C���o�^
Set objExcel = CreateObject("Excel.Application")
objExcel.Workbooks.Add
Set objAddin = objExcel.AddIns.Add(installPath, True)
objAddin.Installed = True

If Err.Number = 0 Then
   MsgBox "�A�h�C���͐���ɃC���X�g�[������܂����B", vbInformation,addInName
Else
   MsgBox "�G���[���������܂����B" & vbCrLf & "���s�����m�F���Ă��������B", vbExclamation,addInName
End If

'// Excel �I��
objExcel.Quit

Set objAddin = Nothing
Set objExcel = Nothing

