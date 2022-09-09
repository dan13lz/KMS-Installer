' ==== Script Information Header ====
' script name:   KMS_Installer
' version:       1.0
' date:          08.09.22
' autor:         Daniel Zalivakhin
' description:   The script removes KMS files, activates Windows and Office, set KMS server - kms.aamajor.local

' ==== Script Main Logic ====
' Enable manual error handler/Включить ручной обработчик ошибок
'On Error Resume Next

' Creating shell and file system objects/Создание объектов оболочки и файловой системы
Set oShell = CreateObject("wscript.shell")
Set oFSO = CreateObject("Scripting.Filesystemobject")

' Defining the paths of service folders/Определение путей служебных папок
sProgramFiles = oShell.ExpandEnvironmentStrings("%ProgramFiles%")
sUserProfileDir = oShell.ExpandEnvironmentStrings("D:\Temp")            '%temp%

' Creating a script log/Создание журнала работы сценария
sLogFileName = sUserProfileDir & "\KMSTemp_" 
Set oLogFile = oFSO.CreateTextFile(sLogFileName & ".log",true)
oLogFile.WriteLine "========== Script 'KMS_Install' started =========="

' Поиск файлов на удаление
Set objShellApp = CreateObject("Shell.Application")
Set objFolder = objShellApp.NameSpace("D:\")
Set objFolderItems = objFolder.Items()
objFolderItems.Filter 64 + 128, "*.txt"
For Each objItem In objFolderItems
    strList = strList & objItem.Name & vbNewLine
    oLogFile.WriteLine Date & " " & Time & " - File found: " & objItem.Path
    WScript.Echo Date & " " & Time & " - File found: " & objItem.Path
Next
WScript.Echo Date & " " & Time & " - Files number: " & objFolderItems.Count 
oLogFile.WriteLine Date & " " & Time & " - Files number: " 











' Closing the log file/Закрытие файла журнала
oLogFile.WriteLine vbCrLf & "======== Script 'KMS_Install' is finished ========"
WScript.Quit 0
oLogFile.Close
gFile.Close
