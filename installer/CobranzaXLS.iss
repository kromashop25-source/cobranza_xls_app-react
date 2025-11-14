#define MyAppName       "Cobranza XLS"
#define MyAppVersion    "1.0.0"
#define MyAppPublisher  "Medileser OI"
#define MyAppExeName    "CobranzaTray.exe"

[Setup]
AppId={{C1C2E6B5-5B0B-4F0E-8D5D-9A8E1F7C7F11}}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}

; PER-USUARIO (sin UAC): instala en AppData del usuario actual
PrivilegesRequired=lowest
DefaultDirName={userappdata}\{#MyAppName}
DisableDirPage=yes

DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputBaseFilename=CobranzaXLS-Setup
Compression=lzma
SolidCompression=yes
ArchitecturesInstallIn64BitMode=x64
SetupLogging=yes
WizardStyle=modern
; Opcional: tu icono del instalador si quieres
; SetupIconFile=backend\app\static\cobranza.ico

[Languages]
Name: "spanish"; MessagesFile: "compiler:Languages\Spanish.isl"

[Tasks]
Name: "desktopicon"; Description: "Crear icono en el Escritorio"; Flags: unchecked
Name: "autostart";   Description: "Iniciar Cobranza XLS al abrir sesión de Windows"; Flags: checkedonce

[Files]
; Copia TODO lo que generó PyInstaller
Source: "dist\CobranzaTray\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
; Menú Inicio
Name: "{group}\Cobranza XLS";  Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"
; Escritorio (opcional)
Name: "{commondesktop}\Cobranza XLS"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon; WorkingDir: "{app}"
; AutoStart (carpeta Inicio del usuario)
Name: "{userstartup}\Cobranza XLS"; Filename: "{app}\{#MyAppExeName}"; Tasks: autostart; WorkingDir: "{app}"

[Run]
; Corre la app al terminar la instalación
Filename: "{app}\{#MyAppExeName}"; Description: "Ejecutar Cobranza XLS"; Flags: nowait postinstall skipifsilent

[UninstallRun]
; Cierra la app si está abierta al desinstalar
Filename: "{cmd}"; Parameters: "/C taskkill /IM {#MyAppExeName} /F"; Flags: runhidden

[Code]
procedure CurStepChanged(CurStep: TSetupStep);
var ResultCode: Integer;
begin
  if CurStep = ssInstall then
  begin
    { Cierra instancias previas antes de copiar }
    ShellExec('open', ExpandConstant('{cmd}'), '/C taskkill /IM {#MyAppExeName} /F', '',
      SW_HIDE, ewWaitUntilTerminated, ResultCode);
  end;
end;
