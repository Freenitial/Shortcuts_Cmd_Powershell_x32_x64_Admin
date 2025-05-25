@echo off & setlocal EnableDelayedExpansion & cd /d %~dp0

REM Author      : Léo Gillet - Freenitial on GitHub
REM Version     : 1.0
REM Tested on   : Windows XP / 10 / 11

REM Description :
REM - Script to create CMD and PowerShell shortcuts as x32/x64/Admin contexts
REM - Shortcuts will be dropped where you want, define the "Drop_Shortcuts_On" variable.
REM - CMD and PowerShell will set their current directory at opening, where you want, define the "Console_CD" variable.

:: __________________________________________________
::|                   CONFIGURATION                  |
::|__________________________________________________|

REM Careful on windows XP, "Desktop" must be written as your Windows language (if you want to write Desktop in 'Drop_Shortcuts_On' or 'Console_CD')
REM To define desktop with greater compatibility, you can replace "Drop_Shortcuts_On" and "Console_CD" by these three lines :
:: for /f "tokens=2*" %%a in ('reg query "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" /v Desktop') do set "DesktopPath=%%b"
:: set "Drop_Shortcuts_On=%DesktopPath%"
:: set "Console_CD=%DesktopPath%"

REM Where to place the created shortcuts (%~dp0 mean in the same folder where you put this script)
set "Drop_Shortcuts_On=%cd%"

REM Where you want to set CD (current directory) at CMD/PowerShell opening.
REM If 260+ characters in shortcut, you won't be able to edit path in properties after creation.
REM IF YOU DON'T CARE ABOUT EDITING THE SHORTCUTS AFTER CREATION, YOU CAN IGNORE THIS LIMITATION.
REM Actually the longest shortcut command (not including Console_CD) is 209 characters.
REM The "Console_CD" path below must not exceed 50 characters to keep all the sortcuts editables.
REM Double %% of environment variables in Console_CD to keep them unexpanded in final shortcuts (shorter).
set "Console_CD=%%UserProfile%%\Desktop"

::__________________________________________________

call set "Test_Drop_Shortcuts_On=!Drop_Shortcuts_On!"
call set "Test_Console_CD=!Console_CD!"
if not exist !Test_Drop_Shortcuts_On! echo Shortcuts cannot be dropped in '!Test_Drop_Shortcuts_On!' because this path was not found. & pause & exit /b 2
if not exist !Test_Console_CD! echo CMD/PowerShell cannot set location in '!Test_Console_CD!' because this path was not found. & pause & exit /b 2

:: Store paths in registry temporarily - preserving accents encoding
reg add "HKCU\Software\TempShortcutCreator" /v "ConsolePath" /d "%Console_CD%" /f >nul
reg add "HKCU\Software\TempShortcutCreator" /v "DropPath" /d "%Drop_Shortcuts_On%" /f >nul

:: Create temporary VBScript
set "vbs=%temp%\CreateShortcut.vbs"
(
    echo Option Explicit
    echo Dim WshShell, FSO, dropPath, consolePath, consolePathRaw
    echo Set WshShell = CreateObject("WScript.Shell"^)
    echo Set FSO = CreateObject("Scripting.FileSystemObject"^)
    echo.
    echo ' Read paths from registry
    echo On Error Resume Next
    echo consolePathRaw = WshShell.RegRead("HKCU\Software\TempShortcutCreator\ConsolePath"^)
    echo dropPath = WshShell.RegRead("HKCU\Software\TempShortcutCreator\DropPath"^)
    echo On Error Goto 0
    echo.
    echo ' Keep raw console path for arguments with variables
    echo ' Expand drop path for file creation
    echo dropPath = WshShell.ExpandEnvironmentStrings(dropPath^)
    echo.
    echo Sub CreateLink(name, target, args, icon^)
    echo     On Error Resume Next
    echo     Dim lnk, fullPath, finalArgs
    echo     fullPath = dropPath ^& "\" ^& name ^& ".lnk"
    echo     Set lnk = WshShell.CreateShortcut(fullPath^)
    echo     lnk.TargetPath = target
    echo     ' Replace CONSOLEPATH with raw path keeping variables
    echo     finalArgs = Replace(args, "CONSOLEPATH", consolePathRaw^)
    echo     lnk.Arguments = finalArgs
    echo     lnk.WorkingDirectory = consolePathRaw
    echo     lnk.IconLocation = icon
    echo     lnk.Save
    echo     If Err.Number = 0 Then
    echo         WScript.Echo "Created: " ^& name
    echo     Else
    echo         WScript.Echo "Error saving " ^& name ^& ": " ^& Err.Description
    echo         Err.Clear
    echo     End If
    echo     On Error Goto 0
    echo End Sub
    echo.
    echo WScript.Echo "Creating shortcuts..."
    echo WScript.Echo ""
    echo.
) > "%vbs%"

:: Normal shortcuts
echo CreateLink "CMD", "%%ComSpec%%", "", "%ComSpec%, 0" >> "%vbs%"
echo CreateLink "CMD - Admin", "%%WinDir%%\System32\WindowsPowerShell\v1.0\powershell", "-Nologo -NoProfile -Ex Bypass -Command SAPS %%comspec%% '/k cd/d ""CONSOLEPATH""' -Verb RunAs", "%WinDir%\System32\SHELL32.dll, 24" >> "%vbs%"
echo CreateLink "PowerShell", "%%ComSpec%%", "/c start %%WinDir%%\System32\WindowsPowerShell\v1.0\powershell -Nologo -NoProfile -NoExit -Ex Bypass cd 'CONSOLEPATH'", "%%WinDir%%\System32\WindowsPowerShell\v1.0\powershell.exe, 0" >> "%vbs%"
echo CreateLink "PowerShell - Admin", "%%WinDir%%\System32\WindowsPowerShell\v1.0\powershell", "-Nologo -NoProfile -Ex Bypass -Command SAPS %%WinDir%%\System32\WindowsPowerShell\v1.0\powershell '-Nologo -NoProfile -Ex Bypass -NoExit cd '""'CONSOLEPATH'""'' -Verb RunAs", "%WinDir%\System32\WindowsPowerShell\v1.0\powershell.exe, 1" >> "%vbs%"
:: 32 bits forced shortcuts
if exist "%WinDir%\SysWOW64" (
    echo CreateLink "CMD x32", "%%ComSpec%%", "/c start %%WinDir%%\SysWOW64\cmd /k cd/d ""CONSOLEPATH""", "%WinDir%\SysWOW64\cmd.exe, 0" >> "%vbs%"
    echo CreateLink "CMD x32 - Admin", "%%WinDir%%\SysWOW64\WindowsPowerShell\v1.0\powershell", "-Nologo -NoProfile -Ex Bypass -Command SAPS %%WinDir%%\SysWOW64\cmd '/k cd/d ""CONSOLEPATH""' -Verb RunAs", "%WinDir%\System32\SHELL32.dll, 24" >> "%vbs%"
    echo CreateLink "PowerShell x32", "%%ComSpec%%", "/c start %%WinDir%%\SysWOW64\WindowsPowerShell\v1.0\powershell -Nologo -NoProfile -NoExit -Ex Bypass cd 'CONSOLEPATH'", "%%WinDir%%\SysWOW64\WindowsPowerShell\v1.0\powershell.exe, 0" >> "%vbs%"
    echo CreateLink "PowerShell x32 - Admin", "%%WinDir%%\SysWOW64\WindowsPowerShell\v1.0\powershell", "-Nologo -NoProfile -Ex Bypass -Command SAPS %%WinDir%%\SysWOW64\WindowsPowerShell\v1.0\powershell '-Nologo -NoProfile -Ex Bypass -NoExit cd '""'CONSOLEPATH'""'' -Verb RunAs", "%WinDir%\SysWOW64\WindowsPowerShell\v1.0\powershell.exe, 1" >> "%vbs%"
)

:: Execute and cleanup
echo.
cscript //nologo "%vbs%"
del "%vbs%" 2>nul
reg delete "HKCU\Software\TempShortcutCreator" /f >nul 2>&1

echo.
echo Process completed.
echo.
pause