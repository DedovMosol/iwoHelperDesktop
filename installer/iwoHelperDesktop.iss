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
; Язык выбирается СВОИМ стартовым экраном с флагами (Великобритания / Россия — см. [Code]
; InitializeWizard/PromptLanguageByFlags), а не стандартным дропдауном Inno: он не умеет
; флаги. Дефолтный дропдаун выключен; язык мастера авто-определяется по системе, а при
; выборе другого флага setup перезапускается с /LANG=. Выбранный язык становится и языком
; приложения по умолчанию (см. SeedLanguage) — англоязычному не выскакивает русский.
ShowLanguageDialog=no
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
; Английский (Default.isl) и русский. Язык выбирается в начале установки (ShowLanguageDialog)
; и задаёт ЯЗЫК ПРИЛОЖЕНИЯ по умолчанию — [Code] сидит его в settings.txt. По умолчанию
; предлагается язык системы (LanguageDetectionMethod=uilanguage); англ. — первый в списке,
; поэтому нераспознанная локаль получает английский мастер, а не русский.
Name: "en"; MessagesFile: "compiler:Default.isl"
Name: "ru"; MessagesFile: "compiler:Languages\Russian.isl"

[Messages]
; Явное предупреждение про режим установки на странице приветствия (на языке мастера).
ru.WelcomeLabel2=Будет установлено приложение «iwo Helper Desktop» {#AppVersion}.%n%nВНИМАНИЕ: по умолчанию программа устанавливается ТОЛЬКО для текущего пользователя (права администратора не нужны). Чтобы установить для всех пользователей этого компьютера, выберите соответствующий вариант в начале установки.%n%nРекомендуется закрыть остальные приложения перед продолжением.
en.WelcomeLabel2=This will install iwo Helper Desktop {#AppVersion}.%n%nNOTE: by default the app is installed for the CURRENT USER ONLY (no administrator rights required). To install it for all users of this computer, choose that option at the start of setup.%n%nIt is recommended to close all other applications before continuing.

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
; Флаги для стартового выбора языка (извлекаются во временную папку, не устанавливаются).
Source: "flag_en.bmp"; Flags: dontcopy
Source: "flag_ru.bmp"; Flags: dontcopy

[Icons]
Name: "{autoprograms}\{#AppName}"; Filename: "{app}\{#AppExe}"
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExe}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#AppExe}"; Description: "{cm:LaunchProgram,{#AppName}}"; Flags: nowait postinstall skipifsilent

[Code]
// ExitProcess — закрыть текущий экземпляр после перезапуска в выбранном языке (Inno фиксирует
// язык мастера до его создания; сменить на лету нельзя — только перезапуск).
procedure ExitProcess(uExitCode: Cardinal);
  external 'ExitProcess@kernel32.dll stdcall';

// Есть ли среди аргументов командной строки переключатель с данным префиксом (например
// «/SILENT» или «/LANG=»). Чтобы не показывать выбор языка при тихой установке и в уже
// перезапущенном экземпляре (там передан /LANG=).
function HasSwitch(const Prefix: String): Boolean;
var
  I: Integer;
begin
  Result := False;
  for I := 1 to ParamCount do
    if CompareText(Copy(ParamStr(I), 1, Length(Prefix)), Prefix) = 0 then
    begin
      Result := True;
      Exit;
    end;
end;

// Стартовый экран выбора языка флагами (Великобритания / Россия). Возвращает 'en' или 'ru';
// по умолчанию (окно закрыли) — текущий язык. Свой, потому что дропдаун Inno флаги не умеет.
function PromptLanguageByFlags(): String;
var
  Form: TSetupForm;
  Title: TNewStaticText;
  EnImg, RuImg: TBitmapImage;
  EnBtn, RuBtn: TNewButton;
  cw, colL, colR, flagY, btnY: Integer;
begin
  Result := ActiveLanguage();
  ExtractTemporaryFile('flag_en.bmp');
  ExtractTemporaryFile('flag_ru.bmp');
  { Inno 7: CreateCustomForm(ClientWidth, ClientHeight, KeepSizeX, KeepSizeY). Без вызова
    FlipAndCenterIfNeeded форма сама центрируется на экране — то, что нужно до мастера. }
  Form := CreateCustomForm(ScaleX(400), ScaleY(210), False, True);
  try
    Form.Caption := 'iwo Helper Desktop';
    cw := Form.ClientWidth;
    colL := cw div 4;          { центр левой колонки — English }
    colR := (cw * 3) div 4;    { центр правой колонки — Русский }
    flagY := ScaleY(70);
    btnY := ScaleY(132);

    Title := TNewStaticText.Create(Form);
    Title.Parent := Form;
    Title.AutoSize := True;
    Title.Font.Size := 11;
    Title.Caption := 'Choose language  /  Выберите язык';
    Title.Left := (cw - Title.Width) div 2;
    Title.Top := ScaleY(22);

    EnImg := TBitmapImage.Create(Form);
    EnImg.Parent := Form;
    EnImg.Stretch := True;
    EnImg.Bitmap.LoadFromFile(ExpandConstant('{tmp}\flag_en.bmp'));
    EnImg.SetBounds(colL - ScaleX(33), flagY, ScaleX(66), ScaleY(44));

    RuImg := TBitmapImage.Create(Form);
    RuImg.Parent := Form;
    RuImg.Stretch := True;
    RuImg.Bitmap.LoadFromFile(ExpandConstant('{tmp}\flag_ru.bmp'));
    RuImg.SetBounds(colR - ScaleX(33), flagY, ScaleX(66), ScaleY(44));

    EnBtn := TNewButton.Create(Form);
    EnBtn.Parent := Form;
    EnBtn.SetBounds(colL - ScaleX(48), btnY, ScaleX(96), ScaleY(30));
    EnBtn.Caption := 'English';
    EnBtn.ModalResult := 1;

    RuBtn := TNewButton.Create(Form);
    RuBtn.Parent := Form;
    RuBtn.SetBounds(colR - ScaleX(48), btnY, ScaleX(96), ScaleY(30));
    RuBtn.Caption := 'Русский';
    RuBtn.ModalResult := 2;

    if ActiveLanguage() = 'en' then EnBtn.Default := True else RuBtn.Default := True;

    case Form.ShowModal() of
      1: Result := 'en';
      2: Result := 'ru';
    end;
  finally
    Form.Free();
  end;
end;

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

// Показать флаг-пикер языка ПЕРВЫМ делом в мастере (это первое окно после стандартного диалога
// режима установки Inno — раньше него кастомную форму показать нельзя: в InitializeSetup ещё нет
// цикла сообщений, модальная форма не держится). Если выбран НЕ системный язык — перезапускаем
// setup с /LANG= И тем же режимом установки (/ALLUSERS|/CURRENTUSER, чтобы диалог режима не
// повторялся): весь мастер и приложение (см. SeedLanguage) получат выбранный язык. Пропуск —
// тихая установка и уже перезапущенный экземпляр.
procedure InitializeWizard();
var
  Chosen, Mode: String;
  ErrCode: Integer;
begin
  if HasSwitch('/SILENT') or HasSwitch('/VERYSILENT') or HasSwitch('/LANG=') then
    Exit;
  try
    Chosen := PromptLanguageByFlags();
  except
    Exit; // сбой показа не роняет установку — остаёмся на языке системы
  end;
  if Chosen <> ActiveLanguage() then
  begin
    if IsAdminInstallMode() then Mode := ' /ALLUSERS' else Mode := ' /CURRENTUSER';
    if Exec(ExpandConstant('{srcexe}'), '/LANG=' + Chosen + Mode, '', SW_SHOW, ewNoWait, ErrCode) then
      ExitProcess(0); // закрыть текущий экземпляр; перезапущенный идёт в выбранном языке
  end;
end;

// Записать язык приложения по умолчанию из выбранного языка мастера в settings.txt
// (%APPDATA%\iwo Helper Desktop). Приложение читает эту строку при старте — так язык
// установки становится языком интерфейса. Пишем ТОЛЬКО при ПЕРВОЙ установке (файла ещё
// нет): единственную ASCII-строку language=xx. Перезаписывать существующий файл нельзя —
// Inno читает/пишет его в системной кодировке (ANSI), а .NET пишет UTF-8, поэтому
// read-modify-write испортил бы НЕ-ASCII содержимое (кириллические пути в lastInputFolder).
// Если файл уже есть — язык задан приложением или прошлой установкой, не трогаем.
procedure SeedLanguage();
var
  Dir, FilePath, Code: String;
  Lines: TArrayOfString;
begin
  Dir := ExpandConstant('{userappdata}\iwo Helper Desktop');
  FilePath := Dir + '\settings.txt';
  if FileExists(FilePath) then
    Exit; // настройки уже есть — язык задан, кодировку не рискуем
  if ActiveLanguage() = 'en' then Code := 'en' else Code := 'ru';
  ForceDirectories(Dir);
  SetArrayLength(Lines, 1);
  Lines[0] := 'language=' + Code; // единственная строка, чистый ASCII
  SaveStringsToFile(FilePath, Lines, False);
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
    SeedLanguage(); // язык установки → язык приложения по умолчанию
end;
