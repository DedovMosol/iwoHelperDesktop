; Установщик iwo Helper Desktop (Inno Setup 6).
; Ставит приложение + вшитый Ghostscript (сжатие PDF «как в Acrobat» из коробки).
; По умолчанию — для текущего пользователя без прав администратора (%LOCALAPPDATA%);
; в диалоге можно выбрать установку «для всех» (Program Files, потребует админа).
; Сборка: tools\make_installer.ps1 [-Arch x86] (stage GS -> ISCC -> подпись). Версия
; передаётся через /DAppVersion; при отсутствии берётся из версии exe своей разрядности.

#define AppName "iwo Helper Desktop"
#define AppExe "iwoHelperDesktop.exe"
#define Publisher "Dodonov Andrey (DedovMosol)"
#define AppUrl "https://github.com/DedovMosol/iwoHelperDesktop"

; Разрядность пакета: /DArch=x86 собирает 32-битный установщик — exe из dist\x86,
; 32-битный Ghostscript из installer\gs32, суффикс -x86 в имени файла. По умолчанию x64.
#ifndef Arch
  #define Arch "x64"
#endif
#if Arch == "x86"
  #define DistDir "..\dist\x86"
  #define GsDir "gs32"
  #define GsDll "gsdll32.dll"
  #define GsExe "gswin32c.exe"
  #define SetupSuffix "-x86"
#elif Arch == "x64"
  #define DistDir "..\dist"
  #define GsDir "gs"
  #define GsDll "gsdll64.dll"
  #define GsExe "gswin64c.exe"
  #define SetupSuffix ""
#else
  #error Unknown Arch - use x64 or x86
#endif
#ifndef AppVersion
  #define AppVersion GetFileVersion(DistDir + "\" + AppExe)
#endif

[Setup]
AppId={{8F3A1B62-9D4E-4C7A-B0E5-2A6F1C93D7E4}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#Publisher}
AppPublisherURL={#AppUrl}
AppSupportURL={#AppUrl}
; Версия и описание в РЕСУРСАХ самого Setup.exe (свойства файла в проводнике):
; без этих директив Inno оставляет 0.0.0.0 и пустое описание.
VersionInfoVersion={#AppVersion}
VersionInfoProductVersion={#AppVersion}
VersionInfoProductName={#AppName}
VersionInfoDescription={#AppName} Setup
VersionInfoCompany={#Publisher}
VersionInfoCopyright=© 2026 {#Publisher}. MIT License.
DefaultDirName={autopf}\{#AppName}
DefaultGroupName={#AppName}
DisableProgramGroupPage=yes
; По умолчанию — без админа (per-user); пользователь может выбрать «для всех».
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
; ВСЕГДА спрашивать режим (для текущего пользователя / для всех) и путь установки —
; даже при повторной установке. По умолчанию Inno при обновлении их скрывает:
;   UsePreviousPrivileges=no  -> возвращает вопрос режима (иначе берёт прошлый);
;   DisableDirPage=no         -> всегда показывать выбор папки (дефолт auto прячет);
;   UsePreviousAppDir=yes     -> при этом прошлый путь подставляется как значение по умолчанию.
UsePreviousPrivileges=no
DisableDirPage=no
UsePreviousAppDir=yes
; Показать страницу приветствия — на ней явно предупреждаем про установку
; только для текущего пользователя (см. [Messages] WelcomeLabel2).
DisableWelcomePage=no
; Минимальная ОС — Windows 8.1 (NT 6.3): раньше неё нет Windows.Data.Pdf (миниатюры).
; .NET Framework 4.8 проверяется в [Code] (в Windows 10 1903+ уже встроен).
MinVersion=6.3
#if Arch == "x64"
; 64-битный пакет — только на 64-битные Windows (x64compatible включает и ARM64
; с эмуляцией x64; x64compatible — рекомендуемый идентификатор в Inno Setup 6.3+/7).
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
#endif
; 32-битный пакет без директив архитектуры: работает всюду, ставится в 32-битном режиме.
OutputDir=..\dist
OutputBaseFilename=iwoHelperDesktop-setup-{#AppVersion}{#SetupSuffix}
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
Source: "{#DistDir}\{#AppExe}"; DestDir: "{app}"; Flags: ignoreversion
; Ghostscript своей разрядности (подготовлен tools\stage_gs.ps1 в installer\{#GsDir}\).
; Приложение ищет в {app}\gs\bin оба имени (gswin64c/gswin32c), путь установки общий.
Source: "{#GsDir}\bin\{#GsDll}"; DestDir: "{app}\gs\bin"; Flags: ignoreversion
Source: "{#GsDir}\bin\{#GsExe}"; DestDir: "{app}\gs\bin"; Flags: ignoreversion
Source: "{#GsDir}\lib\*"; DestDir: "{app}\gs\lib"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "{#GsDir}\Resource\*"; DestDir: "{app}\gs\Resource"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "{#GsDir}\iccprofiles\*"; DestDir: "{app}\gs\iccprofiles"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "{#GsDir}\LICENSE"; DestDir: "{app}\gs"; DestName: "LICENSE.txt"; Flags: ignoreversion skipifsourcedoesntexist

[Icons]
Name: "{autoprograms}\{#AppName}"; Filename: "{app}\{#AppExe}"
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExe}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#AppExe}"; Description: "{cm:LaunchProgram,{#AppName}}"; Flags: nowait postinstall skipifsilent

[Code]
// Приложению нужен .NET Framework 4.8: в Windows 10 1903+ он встроен, на Windows 8.1
// ставится один раз. Release >= 528040 означает 4.8+ (документированные значения
// Microsoft). Читаем 64-битную ветку на 64-битных ОС: ключ NDP пишется именно туда,
// а 32-битный установщик по умолчанию видел бы WOW6432Node.
function InitializeSetup(): Boolean;
var
  Root: Integer;
  Release: Cardinal;
  ErrCode: Integer;
begin
  Result := True;
  if IsWin64 then Root := HKLM64 else Root := HKLM;
  if not (RegQueryDWordValue(Root, 'SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full',
      'Release', Release) and (Release >= 528040)) then
  begin
    if MsgBox('Для работы «iwo Helper Desktop» нужен .NET Framework 4.8.' + #13#10 +
        'В Windows 10 (1903+) и Windows 11 он уже встроен, на Windows 8.1 его нужно установить один раз.' + #13#10#13#10 +
        'Открыть страницу загрузки .NET Framework 4.8?', mbConfirmation, MB_YESNO) = IDYES then
      ShellExecAsOriginalUser('open',
        'https://dotnet.microsoft.com/download/dotnet-framework/net48', '', '',
        SW_SHOWNORMAL, ewNoWait, ErrCode);
    Result := False;
  end;
end;
