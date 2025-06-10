[Icons]
Name: {group}\PiLfIuS!; Filename: {app}\PiLfIuS!.exe; WorkingDir: {app}; IconFilename: {app}\PiLfIuS!.exe; IconIndex: 0; Components: ; Tasks: " startmenu"
Name: {group}\What's New; Filename: {app}\whatsnew.txt; Components: ; Tasks: " startmenu"
Name: {group}\Readme; Filename: {app}\readme.txt; Components: ; Tasks: " startmenu"
Name: {commondesktop}\PiLfIuS!; Filename: {app}\PiLfIuS!.exe; WorkingDir: {app}; IconFilename: {app}\PiLfIuS!.exe; IconIndex: 0; Tasks: " desktopicon"
[Setup]
OutputDir=C:\Diego\Developments\PiLfIuS!\Application\installer
AppCopyright=Copyright © 2007 Diego Wasser
AppName=PiLfIuS!
AppVerName=PiLfIuS! 0.9
LicenseFile=C:\Diego\Developments\PiLfIuS!\Application\license.txt
DefaultDirName={pf}\PiLfIuS
DefaultGroupName=PiLfIuS!
WizardImageBackColor=clWhite
WizardImageFile=C:\Diego\Developments\PiLfIuS!\Application\installimgINNO.bmp
WizardSmallImageFile=C:\Diego\Developments\PiLfIuS!\Application\installimgINNOsmall.bmp
VersionInfoCopyright=Copyright © 2007 Diego Wasser
AppMutex=PiLfIuS
SetupIconFile=C:\Diego\Developments\PiLfIuS!\Application\pilfius_installicon.ico
UninstallFilesDir={app}\uninstall
OutputBaseFilename=pilfius_nonstable
InfoBeforeFile=C:\Diego\Developments\PiLfIuS!\Application\install_notice.rtf
[Files]
Source: ..\readme.txt; DestDir: {app}
Source: ..\whatsnew.txt; DestDir: {app}
Source: ..\PiLfIuS!.exe; DestDir: {app}
Source: ..\comdlg32.dll; DestDir: {sys}; Flags: regserver uninsneveruninstall onlyifdoesntexist allowunsafefiles noregerror
Source: ..\Comdlg32.ocx; DestDir: {sys}; Flags: regserver uninsneveruninstall onlyifdoesntexist allowunsafefiles noregerror; Tasks: 
[Tasks]
Name: desktopicon; Description: Create desktop icon
Name: startmenu; Description: Create folder in start menu
