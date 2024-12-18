#define MyAppName "Cutting Optimizer Pro"
#define MyAppVersion "1.2.0"
#define MyAppPublisher "Mehdi Harzallah"
#define MyAppURL "https://github.com/opestro"
#define MyAppExeName "Cutting-Optimizer-Pro.exe"

[Setup]
; NOTE: The value of AppId uniquely identifies this application.
AppId={{ElHamdulilah}}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
OutputDir=installer
OutputBaseFilename=CuttingOptimizerPro_Setup
Compression=lzma
SolidCompression=yes
; Require admin rights to install
PrivilegesRequired=admin
; Set minimum Windows version
MinVersion=6.1
; Architecture settings
ArchitecturesAllowed=x64
ArchitecturesInstallIn64BitMode=x64

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"
Name: "french"; MessagesFile: "compiler:Languages\French.isl"
Name: "arabic"; MessagesFile: "compiler:Languages\Arabic.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
; Main executable and dependencies
Source: "dist\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs

; Fonts
Source: "fonts\*"; DestDir: "{fonts}"; FontInstall: "Noto Sans Arabic"; Flags: onlyifdoesntexist uninsneveruninstall

; Icons
Source: "icons\*"; DestDir: "{app}\icons"; Flags: ignoreversion recursesubdirs

; Configuration files
Source: "translations.json"; DestDir: "{app}"; Flags: ignoreversion
Source: "version.json"; DestDir: "{app}"; Flags: ignoreversion

[Dirs]
; Create AppData directories during installation
Name: "{userappdata}\{#MyAppName}"
Name: "{userappdata}\{#MyAppName}\logs"
Name: "{userappdata}\{#MyAppName}\output"

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

[Code]
// Check if .NET Framework 4.7.2 or later is installed
function IsDotNetDetected(): boolean;
var
    success: boolean;
    release: cardinal;
begin
    success := RegQueryDWordValue(HKLM, 'SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full', 'Release', release);
    result := success and (release >= 461808); // .NET 4.7.2 or later
end;

// Initialize setup
function InitializeSetup(): Boolean;
begin
    if not IsDotNetDetected() then begin
        MsgBox('This application requires .NET Framework 4.7.2 or later.'#13#13
               'Please install it and run this setup again.', 
               mbInformation, MB_OK);
        result := false;
    end else
        result := true;
end;

// Create application data folder during installation
procedure CurStepChanged(CurStep: TSetupStep);
var
    AppDataPath: string;
begin
    if CurStep = ssPostInstall then begin
        AppDataPath := ExpandConstant('{userappdata}\{#MyAppName}');
        if not DirExists(AppDataPath) then
            CreateDir(AppDataPath);
            
        // Create logs directory
        if not DirExists(AppDataPath + '\logs') then
            CreateDir(AppDataPath + '\logs');
    end;
end; 