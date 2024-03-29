; Script generated by the Inno Setup Script Wizard.
; SEE THE DOCUMENTATION FOR DETAILS ON CREATING INNO SETUP SCRIPT FILES!

#define MyAppName "Journal Entries Converter"
#define MyAppVersion "1.0.0.1"
#define MyAppPublisher "PGC IT Department"
#define MyAppURL "http://pgc-itdept.org"
#define MyAppExeName "JEconverter.exe"

[Setup]
; NOTE: The value of AppId uniquely identifies this application.
; Do not use the same AppId value in installers for other applications.
; (To generate a new GUID, click Tools | Generate GUID inside the IDE.)
AppId={{4F198B07-D38A-4A6C-9644-1A51D09A1BEE}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
;AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={pf}\cdt-S0ft\Journal Entries Converter
DefaultGroupName=cdt-S0ft\Journal Entries Converter
OutputDir=D:\CaDeAtH\Documents\DevInstaller\JEc
OutputBaseFilename=jeconv1001
SetupIconFile=D:\CaDeAtH\Documents\GitHub\SAP\JEconverter\JEconverter\Ncrow-Mega-Pack-1-Any-DVD-Converter.ico
Compression=lzma
SolidCompression=yes

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
Source: "D:\CaDeAtH\Documents\GitHub\SAP\JEconverter\JEconverter\bin\Release\JEconverter.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "D:\CaDeAtH\Documents\GitHub\SAP\JEconverter\JEconverter\bin\Release\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
; NOTE: Don't use "Flags: ignoreversion" on any shared system files

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"
Name: "{commondesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

