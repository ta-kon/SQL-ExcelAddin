' �A�ԑ��Y�A���C���X�g�[���X�N���v�g
' 
' �Q�l�T�C�g
' ���� SE �̂Ԃ₫
' VBScript �� Excel �ɃA�h�C���������ŃC���X�g�[��/�A���C���X�g�[��������@
' http://fnya.cocolog-nifty.com/blog/2014/03/vbscript-excel-.html
On Error Resume Next

Dim installPath 
Dim addInName 
Dim addInFileName 
Dim objExcel 
Dim objAddin

'�A�h�C������ݒ� 
addInName = "�N�G�����Y" 
addInFileName = "QueryTaroAddin.xlam"

If (MsgBox(addInName & " �A�h�C�����A���C���X�g�[�����܂����H", vbYesNo + vbQuestion) = vbNo) Then 
  WScript.Quit 
End If

'Excel �C���X�^���X�� 
Set objExcel = CreateObject("Excel.Application") 
objExcel.Workbooks.Add

'�A�h�C���o�^���� 
For i = 1 To objExcel.Addins.Count 
  Set objAddin = objExcel.Addins.item(i) 
  If objAddin.Name = addInFileName Then 
    objAddin.Installed = False 
  End If 
Next

'Excel �I�� 
objExcel.Quit

Set objAddin = Nothing 
Set objExcel = Nothing

Set objWshShell = CreateObject("WScript.Shell") 
Set objFileSys = CreateObject("Scripting.FileSystemObject")

'�C���X�g�[����p�X�̍쐬 
'(ex)C:\Users\[User]\AppData\Roaming\Microsoft\AddIns\[addInFileName] 
installPath = objWshShell.SpecialFolders("Appdata") & "\Microsoft\Addins\" & addInFileName

'�t�@�C���폜 
If (objFileSys.FileExists(installPath)) Then 
  objFileSys.DeleteFile installPath , True 
Else 
  MsgBox "�A�h�C���t�@�C�������݂��܂���B", vbExclamation 
End If

Set objWshShell = Nothing 
Set objFileSys = Nothing

If Err.Number = 0 Then 
   MsgBox "�A�h�C���̃A���C���X�g�[�����I�����܂����B", vbInformation 
Else 
   MsgBox "�G���[���������܂����B" & vbCrLF & "���s�����m�F���Ă��������B", vbExclamation 
End If