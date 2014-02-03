Option Explicit
'##### �T�v ##################
'    hosts�t�@�C���̐ݒ��z�z���邽�߂̃X�N���v�g�ł��B
'    �z�z����Settings�z��ɐݒ肵�������e���L�q���܂��B
'    �g�p�҂̓X�N���v�g�����s���邾���Ŕz�z���ꂽ�ݒ��hosts�t�@�C���ɒǉ����邱�Ƃ��ł��܂��B
'    �ēx���s���邱�ƂŁASettings�z��̓��e��hosts�t�@�C������폜���邱�Ƃ��ł��܂��B
'#############################

'##### hosts settings ########
Dim Settings: Settings = Array( _
	 "127.0.0.1 localhost.example.jp" _
	,"127.0.0.1 localhost.example.com" _
)
'#############################

'##### �Ǘ��Ҍ����Ŏ��s ######
Dim WMI, OS, Value, Shell
Do While WScript.Arguments.Count = 0 And WScript.Version >= 5.7
    '##### WScript5.7 �܂��� Vista �ȏォ���`�F�b�N
    Set WMI = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\.\root\cimv2")
    Set OS = WMI.ExecQuery("SELECT *FROM Win32_OperatingSystem")
    For Each Value In OS
        If left(Value.Version, 3) < 6.0 Then Exit Do
    Next

    
    Set Shell = CreateObject("Shell.Application")
    Shell.ShellExecute "wscript.exe", """" & WScript.ScriptFullName & """ uac", "", "runas"

    WScript.Quit
Loop
'#############################

If Msgbox("�ȉ��̓��e�̐ݒ���s���܂��B��낵���ł���?" & String(2, vbCrLf) & Join(Settings, vbCrLf), vbOKCancel, WScript.ScriptName) = vbCancel Then WScript.Quit

'##### ���C�����������s ######
Const ForReading   = 1
Const ForWriting   = 2
Const ForAppending = 8
Const FileName = "C:\Windows\System32\drivers\etc\hosts"

IsObject New Main

Class Main
	Private FSO
	
	Private Sub Class_Initialize
		Set FSO  = CreateObject("Scripting.FileSystemObject")
		Write CreateNewHosts
		Msgbox ReadAll, VbMsgBoxSetForeground, FileName
	End Sub
	
	Private Function CreateNewHosts
		Dim OldHosts: OldHosts = ReadAll
		Dim NewHosts
		Dim Line
		Dim OldLine

		For Each OldLine In Split(OldHosts, vbCrLf)
			If InStr(1, Join(Settings), OldLine) = 0 Then
				NewHosts = NewHosts & OldLine & vbCrLf '����̐ݒ�ǉ��Ɩ��֌W�Ȋ����s�͂��̂܂�
			End If
		Next
		
		For Each Line In Settings
			If InStr(1, OldHosts, Line) = 0 Then
				NewHosts = NewHosts & Line & vbCrLf '�����s�ɂȂ��ݒ�Ȃ�ǉ�
			ElseIf Msgbox("[" & Line & "] �͊��ɑ��݂��܂��B" & vbCrLf & " �폜���܂���?", vbYesNo, Line) = vbNo Then
				NewHosts = NewHosts & Line & vbCrLf
			End If
		Next
		
		CreateNewHosts = NewHosts
	End Function

	Private Sub Write(Str)
		Dim File: Set File = FSO.OpenTextFile(FileName, ForWriting)
		File.Write Str
	End Sub
	
	Private Function ReadAll()
		Dim File: Set File = FSO.OpenTextFile(FileName, ForReading)
		ReadAll = File.ReadAll
	End Function
End Class
