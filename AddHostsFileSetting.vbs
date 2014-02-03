Option Explicit
'##### 概要 ##################
'    hostsファイルの設定を配布するためのスクリプトです。
'    配布時にSettings配列に設定したい内容を記述します。
'    使用者はスクリプトを実行するだけで配布された設定をhostsファイルに追加することができます。
'    再度実行することで、Settings配列の内容をhostsファイルから削除することができます。
'#############################

'##### hosts settings ########
Dim Settings: Settings = Array( _
	 "127.0.0.1 localhost.example.jp" _
	,"127.0.0.1 localhost.example.com" _
)
'#############################

'##### 管理者権限で実行 ######
Dim WMI, OS, Value, Shell
Do While WScript.Arguments.Count = 0 And WScript.Version >= 5.7
    '##### WScript5.7 または Vista 以上かをチェック
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

If Msgbox("以下の内容の設定を行います。よろしいですか?" & String(2, vbCrLf) & Join(Settings, vbCrLf), vbOKCancel, WScript.ScriptName) = vbCancel Then WScript.Quit

'##### メイン処理を実行 ######
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
				NewHosts = NewHosts & OldLine & vbCrLf '今回の設定追加と無関係な既存行はそのまま
			End If
		Next
		
		For Each Line In Settings
			If InStr(1, OldHosts, Line) = 0 Then
				NewHosts = NewHosts & Line & vbCrLf '既存行にない設定なら追加
			ElseIf Msgbox("[" & Line & "] は既に存在します。" & vbCrLf & " 削除しますか?", vbYesNo, Line) = vbNo Then
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
