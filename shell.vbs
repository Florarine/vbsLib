'attributes
Shell.CurrentDirectory
Shell.Environment[(location)]
'   "system"
'   "user"
'   "process"
'   "volatile"
Shell.SpecialFolders[(folder)]
'   "AllUsersDesktop"
'   "AllUsersStartMenu"
'   "AllUsersPrograms"
'   "AllUsersStartup"
'   "Desktop"
'   "Favorites"
'   "Fonts"
'   "MyDocumentes"
'   "NetHood"
'   "PrintHood"
'   "Programs"
'   "Recent"
'   "SentTo"
'   "StartMenu"
'   "Startup"
'   "Templates"

'method
Shell.AppActivate title
Shell.CreateShortcut(strPathname)
Shell.Exec(strCommand)
Shell.ExpandEnvironmentString(strString)
Shell.LogEvent Type, Message[, Target]
'   value   Event
'   0       success
'   1       error
'   2       warning
'   4       information
'   8       audit success
'   16      audit failure
Shell.Popup(strText, [nWait], [strTitle], [nType])
Shell.RegDelete strName
Shell.RegRead(strName)
Shell.RegWrite strName, value[, strType]
Shell.Run(strCommand, [intWindowStyle], [bWaitOnReturn])
Shell.SendKeys(string)

WshScriptExec.ExitCode
WshScriptExec.ProcessID
WshScriptExec.Status
WshScriptExec.StdErr
WshScriptExec.StdIn
WshScriptExec.StdOut
WshScriptExec.Terminate

'WshShortcut/WshUrlShortcut
WshShortcut.Arguments
WshShortcut.Description
WshShortcut.FullName
WshShortcut.Hotkey
WshShortcut.IconLocation = "filename, iconnumber"
WshShortcut.RelativePath
WshShortcut.TargetPath
WshShortcut.intWindowStyle
WshShortcut.WorkingDirectory
WshShortcut.Save

Environment.Item(name)
Environment.Length
Environment.Count
Environment.Remove strName
