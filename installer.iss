#define AppName "PPT 병합기"
#define AppVersion "1.0"
#define AppExeName "PPT병합기.exe"

[Setup]
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher=zionp
DefaultDirName={autopf}\PPT병합기
DefaultGroupName={#AppName}
OutputDir=dist
OutputBaseFilename=PPT병합기_Setup_v{#AppVersion}
Compression=lzma2
SolidCompression=yes
PrivilegesRequired=admin
WizardStyle=modern
#if FileExists("icon.ico")
SetupIconFile=icon.ico
#endif
UninstallDisplayIcon={app}\{#AppExeName}
ArchitecturesInstallIn64BitMode=x64compatible

[Languages]
Name: "korean"; MessagesFile: "compiler:Languages\Korean.isl"

[Files]
; 앱 실행 파일
Source: "dist\{#AppExeName}"; DestDir: "{app}"; Flags: ignoreversion

; LibreOffice 설치 파일 (LibreOffice 미설치 시에만 추출)
Source: "deps\LibreOffice_Win_x86-64.msi"; DestDir: "{tmp}"; \
  Flags: deleteafterinstall; Check: NeedsLibreOffice

[Icons]
Name: "{group}\{#AppName}"; Filename: "{app}\{#AppExeName}"
Name: "{commondesktop}\{#AppName}"; Filename: "{app}\{#AppExeName}"; \
  Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "바탕 화면에 바로 가기 만들기"; GroupDescription: "추가 작업:"

[Run]
; LibreOffice 설치 (미설치 시에만)
Filename: "msiexec.exe"; \
  Parameters: "/i ""{tmp}\LibreOffice_Win_x86-64.msi"" /quiet /norestart"; \
  Check: NeedsLibreOffice; \
  StatusMsg: "LibreOffice 설치 중... (잠시 기다려 주세요)"; \
  Flags: waituntilterminated

; 설치 완료 후 앱 실행 옵션
Filename: "{app}\{#AppExeName}"; Description: "PPT 병합기 실행"; \
  Flags: nowait postinstall skipifsilent

[Code]
{ LibreOffice 설치 여부 확인 }
function NeedsLibreOffice: Boolean;
var
  Keys: TArrayOfString;
  I: Integer;
  DisplayName: string;
  SofficeExe: string;
begin
  Result := True;

  { soffice.exe 경로 직접 확인 }
  if FileExists('C:\Program Files\LibreOffice\program\soffice.exe') or
     FileExists('C:\Program Files (x86)\LibreOffice\program\soffice.exe') then
  begin
    Result := False;
    Exit;
  end;

  { 언인스톨 레지스트리에서 LibreOffice 검색 (64비트) }
  if RegGetSubkeyNames(HKLM,
    'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall', Keys) then
  begin
    for I := 0 to GetArrayLength(Keys) - 1 do
    begin
      if RegQueryStringValue(HKLM,
        'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\' + Keys[I],
        'DisplayName', DisplayName) then
      begin
        if Pos('LibreOffice', DisplayName) > 0 then
        begin
          Result := False;
          Exit;
        end;
      end;
    end;
  end;

  { 언인스톨 레지스트리에서 LibreOffice 검색 (32비트 호환) }
  if Result and RegGetSubkeyNames(HKLM,
    'SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall', Keys) then
  begin
    for I := 0 to GetArrayLength(Keys) - 1 do
    begin
      if RegQueryStringValue(HKLM,
        'SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\' + Keys[I],
        'DisplayName', DisplayName) then
      begin
        if Pos('LibreOffice', DisplayName) > 0 then
        begin
          Result := False;
          Exit;
        end;
      end;
    end;
  end;
end;

procedure InitializeWizard;
begin
  if NeedsLibreOffice then
    MsgBox(
      'LibreOffice가 설치되어 있지 않아 함께 설치합니다.' + #13#10 +
      '(.ppt 파일 변환에 필요합니다. .pptx만 사용한다면 없어도 됩니다.)',
      mbInformation, MB_OK);
end;
