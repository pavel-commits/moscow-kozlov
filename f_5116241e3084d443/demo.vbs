Set FSO = CreateObject("Scripting.FileSystemObject")
Set F = FSO.GetFile(Wscript.ScriptFullName)
 
Set WshShell = WScript.CreateObject("WScript.Shell") 
DesktopPath = WshShell.SpecialFolders("Desktop") 
Lnk_Title = "\demo@mining_20_bot.lnk" 
Set Shortcut = WshShell.CreateShortcut(DesktopPath&Lnk_Title) 
 
Shortcut.TargetPath = WshShell.ExpandEnvironmentStrings(FSO.GetParentFolderName(F) + "\demo.bat") 
Shortcut.WorkingDirectory = WshShell.ExpandEnvironmentStrings(FSO.GetParentFolderName(F)) 
Shortcut.IconLocation = FSO.GetParentFolderName(F) + "\demo.ico"
Shortcut.WindowStyle = 1 
 
Shortcut.Save 

Dim Res,Text,Title  ' Объявляем переменные

Text="Програма  установлена. Открыть её можно с ярлыка на рабочем столе! Инструкция по настройке в боте. Пароль у @btc_faerm_pro. На тест выделен всего 1 час!"

Title="Готово!"

' Выводим диалоговое окно на экран

Res=MsgBox(Text,vbOkCancel+vbInformation+vbDefaultButton2,Title)

' Определяем, какая из кнопок была нажата в диалоговом окне

If Res=vbOk Then

 MsgBox "Приятного использования. Желаю удачи! Результаты будут в папке, с которой ты устанавливал софт!"

Else

 MsgBox "Брат, ты зачем кнопку отмена нажал?"

End If
