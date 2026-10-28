; AlwaysIME インストーラー（Inno Setup 6.3 以上。7.x でもコンパイルできる書き方にしてある）
; ビルド: installer\build.ps1（dotnet publish のあと、このスクリプトを ISCC でコンパイルします）
; 個別に実行する場合: ISCC /DMyAppVersion=0.6.8 /DPublishDir=publish AlwaysIME.iss
;
; 方針
;  ・既定は、現在のユーザーだけに入れます（管理者権限は不要。%LOCALAPPDATA%\Programs\AlwaysIME）。
;    全ユーザー向け（Program Files）は、画面の選択、または /ALLUSERS で選べます。
;  ・設定とログは、アプリが %AppData%\AlwaysIME に保存します。インストール先には置きません。
;    更新でも、アンインストールでも、設定は消しません。
;  ・更新やアンインストールの前に、起動中のAlwaysIMEを「AlwaysIME.exe --exit」で正常に終了させます。
;    強制終了をすると、MS-IMEのスペース・句読点・キー割り当てが元に戻らないことがあるためです。

#define MyAppName "AlwaysIME"
#define MyAppExe "AlwaysIME.exe"
#define MyAppUrl "https://github.com/jarupxx/AlwaysIME"
#ifndef MyAppVersion
  #define MyAppVersion "0.0.0"
#endif
#ifndef PublishDir
  #define PublishDir "..\AlwaysIME\bin\Release\net10.0-windows10.0.22000.0\win-x64\publish"
#endif

[Setup]
; AppId は、バージョンが変わっても変えないこと（変えると、別のアプリとして扱われます）
AppId={{3C391D6A-BAFD-4E22-8B2C-C7C7E02E0459}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher=jarupxx
AppPublisherURL={#MyAppUrl}
AppSupportURL={#MyAppUrl}/issues
AppUpdatesURL={#MyAppUrl}/releases
AppCopyright=Copyright (C) jarupxx 2024-2026
DefaultDirName={autopf}\{#MyAppName}
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=commandline dialog
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
MinVersion=10.0
UninstallDisplayIcon={app}\{#MyAppExe}
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
OutputDir=Output
OutputBaseFilename=AlwaysIME-{#MyAppVersion}-x64-setup
SetupMutex=AlwaysIME_Setup
; 起動中のアプリは、PrepareToInstall で --exit を使って終了させます（Restart Manager による強制終了はしません）
CloseApplications=no
RestartApplications=no

[Languages]
Name: "japanese"; MessagesFile: "compiler:Languages\Japanese.isl"

[Files]
; 設定ファイル（AlwaysIME.dll.config）と AlwaysIME.portable は、含めません。
; 含めると、ポータブル動作になったり、利用者の設定を上書きしたりします。
Source: "{#PublishDir}\*"; DestDir: "{app}"; Excludes: "*.pdb,AlwaysIME.portable,AlwaysIME.dll.config"; Flags: ignoreversion recursesubdirs createallsubdirs

[InstallDelete]
; 以前のバージョンで同梱していて、いまは無いファイルを残さない
Type: filesandordirs; Name: "{app}\Resources"

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExe}"

[Run]
; 対話式のインストールの最後に、起動するかを選べる
Filename: "{app}\{#MyAppExe}"; Description: "{cm:LaunchProgram,{#MyAppName}}"; Flags: nowait postinstall skipifsilent runasoriginaluser
; サイレントの更新では、更新前に起動していた場合だけ、起動し直す
Filename: "{app}\{#MyAppExe}"; Flags: nowait skipifnotsilent runasoriginaluser; Check: WasRunning

[UninstallRun]
; 終了 → 自動起動の解除、の順。管理者権限で登録された自動起動は解除できないため、利用者が自分で解除する
Filename: "{app}\{#MyAppExe}"; Parameters: "--exit"; RunOnceId: "ExitApp"; Flags: runhidden waituntilterminated
Filename: "{app}\{#MyAppExe}"; Parameters: "--remove-autostart"; RunOnceId: "RemoveAutoStart"; Flags: runhidden waituntilterminated

[Code]
var
  WasRunningBeforeInstall: Boolean;

function WasRunning: Boolean;
begin
  Result := WasRunningBeforeInstall;
end;

// .NET 10 デスクトップ ランタイム（x64）が入っているか。見つからなくても、インストールは止めない
// （wingetでは依存パッケージとして入る。ほかの方法で入れた場合の登録先は、これと異なることがある）
function IsDesktopRuntime10Installed: Boolean;
var
  Names: TArrayOfString;
  I: Integer;
begin
  Result := False;
  if RegGetValueNames(HKLM64, 'SOFTWARE\dotnet\Setup\InstalledVersions\x64\sharedfx\Microsoft.WindowsDesktop.App', Names) then
  begin
    for I := 0 to GetArrayLength(Names) - 1 do
    begin
      if Copy(Names[I], 1, 3) = '10.' then
      begin
        Result := True;
        Exit;
      end;
    end;
  end;
end;

function InitializeSetup: Boolean;
begin
  Result := True;
  if not IsDesktopRuntime10Installed then
  begin
    Log('.NET 10 デスクトップ ランタイム（x64）が見つかりません');
    if not WizardSilent then
    begin
      Result := MsgBox(
        '.NET 10 デスクトップ ランタイム（x64）が見つかりません。' + #13#10 +
        'AlwaysIMEの実行には、これが必要です。' + #13#10 + #13#10 +
        'https://dotnet.microsoft.com/ja-jp/download/dotnet/10.0' + #13#10 + #13#10 +
        'このままインストールを続けますか？',
        mbConfirmation, MB_YESNO) = IDYES;
    end;
  end;
end;

// ファイルをコピーする前に、起動中のAlwaysIMEを正常に終了させる
function PrepareToInstall(var NeedsRestart: Boolean): String;
var
  Exe: String;
  ResultCode: Integer;
begin
  Result := '';
  WasRunningBeforeInstall := CheckForMutexes('AlwaysIME');
  Exe := ExpandConstant('{app}\{#MyAppExe}');
  if WasRunningBeforeInstall and FileExists(Exe) then
  begin
    if not Exec(Exe, '--exit', '', SW_HIDE, ewWaitUntilTerminated, ResultCode) then
    begin
      Log('--exit を実行できませんでした: ' + SysErrorMessage(ResultCode));
      Result := 'AlwaysIMEを終了できませんでした。トレイアイコンの「常駐の終了」で終了してから、もう一度実行してください。';
    end
    else if ResultCode <> 0 then
    begin
      Log('--exit の終了コード: ' + IntToStr(ResultCode));
      Result := 'AlwaysIMEを終了できませんでした。トレイアイコンの「常駐の終了」で終了してから、もう一度実行してください。';
    end;
  end;
end;
