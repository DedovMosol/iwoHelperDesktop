; Установщик iwo Helper Desktop (Inno Setup 6).
; Ставит приложение + вшитый Ghostscript (сжатие PDF «как в Acrobat» из коробки).
; По умолчанию — для текущего пользователя без прав администратора (%LOCALAPPDATA%);
; в диалоге можно выбрать установку «для всех» (Program Files, потребует админа).
; Сборка: tools\make_installer.ps1 (stage GS -> ISCC -> подпись). Версия передаётся
; через /DAppVersion; при отсутствии берётся из версии exe.

#ifndef AppVersion
  #define AppVersion GetFileVersion("..\dist\iwoHelperDesktop.exe")
#endif
#define AppName "iwo Helper Desktop"
#define AppExe "iwoHelperDesktop.exe"
#define Publisher "iwo"

[Setup]
AppId={{8F3A1B62-9D4E-4C7A-B0E5-2A6F1C93D7E4}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#Publisher}
DefaultDirName={autopf}\{#AppName}
DefaultGroupName={#AppName}
DisableProgramGroupPage=yes
; По умолчанию — без админа (per-user); пользователь может выбрать «для всех».
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
; Показать страницу приветствия — на ней явно предупреждаем про установку
; только для текущего пользователя (см. [Messages] WelcomeLabel2).
DisableWelcomePage=no
; Приложение и вшитый Ghostscript — 64-битные (x64compatible включает и ARM64
; с эмуляцией x64; x64compatible — рекомендуемый идентификатор в Inno Setup 6.3+/7).
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
OutputDir=..\dist
OutputBaseFilename=iwoHelperDesktop-setup-{#AppVersion}
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
UninstallDisplayIcon={app}\{#AppExe}
SetupIconFile=..\build\app.ico
; Фирменные картинки мастера вместо стандартных (генерируются tools\make_wizard_images.ps1).
WizardImageFile=wizard.bmp
WizardSmallImageFile=wizard_small.bmp
LicenseFile=license_installer.txt

[Languages]
Name: "ru"; MessagesFile: "compiler:Languages\Russian.isl"

[Messages]
; Явное предупреждение про режим установки на странице приветствия.
ru.WelcomeLabel2=Будет установлено приложение «iwo Helper Desktop» {#AppVersion}.%n%nВНИМАНИЕ: по умолчанию программа устанавливается ТОЛЬКО для текущего пользователя (права администратора не нужны). Чтобы установить для всех пользователей этого компьютера, выберите соответствующий вариант в начале установки.%n%nРекомендуется закрыть остальные приложения перед продолжением.

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
Source: "..\dist\{#AppExe}"; DestDir: "{app}"; Flags: ignoreversion
; Ghostscript (подготовлен tools\stage_gs.ps1 в installer\gs\).
Source: "gs\bin\gsdll64.dll"; DestDir: "{app}\gs\bin"; Flags: ignoreversion
Source: "gs\bin\gswin64c.exe"; DestDir: "{app}\gs\bin"; Flags: ignoreversion
Source: "gs\lib\*"; DestDir: "{app}\gs\lib"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "gs\Resource\*"; DestDir: "{app}\gs\Resource"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "gs\iccprofiles\*"; DestDir: "{app}\gs\iccprofiles"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "gs\LICENSE"; DestDir: "{app}\gs"; DestName: "LICENSE.txt"; Flags: ignoreversion skipifsourcedoesntexist

[Icons]
Name: "{autoprograms}\{#AppName}"; Filename: "{app}\{#AppExe}"
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExe}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#AppExe}"; Description: "{cm:LaunchProgram,{#AppName}}"; Flags: nowait postinstall skipifsilent
