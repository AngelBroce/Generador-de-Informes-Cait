; Script de Inno Setup para CAIT Informes (ACTUALIZACIÓN UNIVERSAL)
; Este script utiliza el AppId exacto de la versión instalada para asegurar la sobrescritura.

[Setup]
; AppId original corregido (con doble llave de cierre para Inno Setup)
AppId={{4D174F38-C2F9-4D81-AB4F-8F4C9A1EC9CE}}
AppName=CAIT Informes
AppVersion=2.0
AppPublisher=CAIT Panamá
DefaultDirName={autopf}\CAIT Informes
DefaultGroupName=CAIT Informes
AllowNoIcons=yes
OutputDir=dist
OutputBaseFilename=Instalador_CAIT_Oficial
SetupIconFile=..\logo-apli-removebg-preview.ico
Compression=lzma
SolidCompression=yes
WizardStyle=modern
DisableProgramGroupPage=yes
PrivilegesRequired=admin
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible

[Languages]
Name: "spanish"; MessagesFile: "compiler:Languages\Spanish.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
; Archivos del ejecutable generados por PyInstaller (carpeta dist/CAIT_Informes)
Source: "..\dist\CAIT_Informes\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\CAIT Informes"; Filename: "{app}\CAIT_Informes.exe"; IconFilename: "{app}\logo-apli-removebg-preview.ico"
Name: "{autodesktop}\CAIT Informes"; Filename: "{app}\CAIT_Informes.exe"; Tasks: desktopicon; IconFilename: "{app}\logo-apli-removebg-preview.ico"

[Run]
Filename: "{app}\CAIT_Informes.exe"; Description: "{cm:LaunchProgram,CAIT Informes}"; Flags: nowait postinstall skipifsilent
