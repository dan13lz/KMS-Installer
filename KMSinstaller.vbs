' ==== Script Information Header ====
' script name:   KMS Installer
' version:       1.0
' date:          08.09.22
' autor:         Daniel Zalivakhin
' description:   ������ ������� KMS �����, ���������� Windows � Office ����� ������ kms.aamajor.local

' ==== Script Main Logic ====
' ��������� ������ ��������� ������
'On Error Resume Next

' �������� �������� �������� � �������� �������
Set oShell = CreateObject("wscript.shell")
Set oFSO = CreateObject("Scripting.Filesystemobject")

' ����������� ����� ��������� �����
sProgramFiles = oShell.ExpandEnvironmentStrings("%ProgramFiles%")
sUserProfileDir = oShell.ExpandEnvironmentStrings("D:\Temp")            '%temp%

' �������� ������� ������ ��������
sLogFileName = sUserProfileDir & "\KMSTemp_" 
Set oLogFile = oFSO.CreateTextFile(sLogFileName & ".log",true)
oLogFile.WriteLine "========== Script KMS Install started =========="

oLogFile.WriteLine Date & " " & Time & " "





Set objShellApp = CreateObject("Shell.Application")
Set objFolder = objShellApp.NameSpace("D:\")
Set objFolderItems = objFolder.Items()
objFolderItems.Filter 64 + 128, "*.txt"
For Each objItem In objFolderItems
    strList = strList & objItem.Name & vbNewLine
Next
WScript.Echo "txt file Count:*.txt: " & objFolderItems.Count _
    & vbNewLine & vbNewLine & strList
    oLogFile.WriteLine Date & " " & Time & " " & strList
WScript.Quit 0













' �������� ����� �������
oLogFile.WriteLine vbCrLf & "======== Script KMS Install is finished ========"
oLogFile.Close