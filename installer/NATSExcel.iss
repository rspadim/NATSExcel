; NATSExcel installer + uninstaller (Inno Setup 6).
;
; Build:  ISCC /DAppVersion=1.0.0 .\installer\NATSExcel.iss
; CI builds it on every run and attaches NATSExcelSetup-<version>.exe to a
; v* tag release (see .github/workflows/build.yml).
;
; Flow (in this order, as the wizard runs):
;   1. asks whether to install a NATS server (native download or Docker) and
;      the connection (server URL + optional token);
;   2. INSTALLS THE SERVER (if chosen);
;   3. INSTALLS NATSExcel (copies the files, keeps the matching XLL, writes
;      NATSExcel.json) and registers the XLL so Excel loads it on start;
;   4. TESTS the connection (visible OK/fail) and, if asked, opens a sample
;      workbook.
;
; NATS runs natively on Windows (a single nats-server.exe), so Docker is an
; option, not a requirement. The add-in installs per-user; only registering the
; optional Windows service (nats-service.ps1) needs admin.
;
; The uninstaller (Start Menu "Uninstall NATSExcel" or Apps & features) removes
; the files, the Excel auto-load entry and the COM ProgId; it asks whether to
; keep your NATSExcel.json and offers to remove a server it created.
;
; Silent mode: NATSExcelSetup.exe /VERYSILENT /DIR=... /URL=... /TOKEN=...
;   /SERVER=<none|native|docker> /MODE=<simple|advanced> /OPENEXCEL=<0|1>

#define AppName "NATSExcel"
#define AppPublisher "rspadim"
#define AppUrl "https://github.com/rspadim/NATSExcel"
#ifndef AppVersion
  #define AppVersion "dev"
#endif
#ifndef BuildDir
  #define BuildDir "..\bin\Release\net48\publish"
#endif
#ifndef RepoRoot
  #define RepoRoot ".."
#endif

[Setup]
AppId={{C41A7E52-9D3B-4F2E-8A17-2B6C0E4D9F51}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#AppPublisher}
AppPublisherURL={#AppUrl}
AppSupportURL={#AppUrl}/issues
DefaultDirName=C:\NATSExcel
DisableDirPage=no
; Per-user install (no admin): the add-in registers its own COM ProgId under HKCU.
PrivilegesRequired=lowest
DefaultGroupName={#AppName}
AllowNoIcons=yes
OutputDir={#RepoRoot}\installer\output
OutputBaseFilename=NATSExcelSetup-{#AppVersion}
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
UninstallDisplayName={#AppName} {#AppVersion}
UninstallDisplayIcon={app}\NATSExcel64-packed.xll

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"
Name: "spanish"; MessagesFile: "compiler:Languages\Spanish.isl"
Name: "brazilianportuguese"; MessagesFile: "compiler:Languages\BrazilianPortuguese.isl"

; All user-visible text. Order: english (default) / spanish / brazilianportuguese.
[CustomMessages]
english.ServerPageTitle=NATS server
english.ServerPageSubtitle=Do you want to install a NATS server now?
english.ServerPageDesc=NATSExcel needs a NATS server. NATS runs natively on Windows (one small nats-server.exe) and JetStream is the same binary with -js, so no Docker is required. Docker is offered as a convenience. The server is installed first, then NATSExcel, and the connection is tested at the end.
english.ServerNone=No - I already have a server (or will set it up later)
english.ServerNative=Yes - native Windows (download nats-server.exe, recommended)
english.ServerDocker=Yes - Docker (nats:latest, JetStream on)
english.ConnPageTitle=NATS connection
english.ConnPageSubtitle=Where should NATSExcel connect?
english.ConnPageDesc=Press Next for the default (nats://localhost:4222). For a cluster, separate the URLs with commas - one is enough, the client discovers the rest. You can change this later by editing NATSExcel.json.
english.ConnUrl=Server URL(s) (comma-separated for a cluster):
english.ConnToken=Token (optional):
english.ConnTokenHint=A token is a single shared secret. For user/password, NKey, JWT or a creds file, edit NATSExcel.json after installing.
english.OptionsPageTitle=Extra checks
english.OptionsPageSubtitle=Almost done
english.OptionsPageDesc=The installer tests the connection when it finishes. Optionally open a sample Excel workbook showing it working.
english.OptionsOpenExcel=Open a sample NATSExcel workbook at the end
english.OptionsService=Register NATS as a Windows service (needs admin; skipped if not elevated)
english.TaskConfigDesc=Create NATSExcel.json in the install folder (connection settings)
english.TaskNlogDesc=Create NLog.config (write a log file next to the add-in)
english.TaskAutoLoadDesc=Register the add-in so Excel loads it automatically (recommended)
english.GroupConfig=Configuration:
english.GroupExcel=Excel:
english.IconUninstall=Uninstall {#AppName}
english.UrlRequired=Please enter a NATS server URL (for example nats://localhost:4222).
english.ServerFailed=Could not set up %1 automatically (exit %2).%nYou can install a server later and point NATSExcel.json at it.
english.ConfigWriteFailed=Could not write %1. Create it manually from the sample (NATSExcel.json).
english.TestFailed=Installed, but the NATS server at %1 did not answer.%nCheck that the server is running and adjust NATSExcel.json if it is on another URL.
english.TestOk=All set! NATS answered at %1.%n%nTo use it: open Excel. If the add-in does not load automatically, go to File > Options > Add-ins > Go... > Browse... and pick the .xll in %2.
english.FinishedText=NATSExcel is installed in:%n%2%n%nIf the add-in does not load automatically, enable it in Excel:%nFile > Options > Add-ins > Go... > Browse... and pick NATSExcel64-packed.xll (or NATSExcel-packed.xll for 32-bit Office).%n%nThen try =RTD("NatsRtd",,"SUB","test.>") in a cell.
english.UninstServersTitle=Servers
english.UninstServersSubtitle=Remove a NATS server too?
english.UninstServersDesc=These NATS servers were found on this machine. Tick the ones you want to remove with NATSExcel.%nBy default NOTHING is removed.
english.UninstServerFound=%1  [found]
english.UninstKeepConfig=Keep your NATSExcel.json settings?
english.ServerNameNative=NATS (native nats-server.exe)
english.ServerNameDocker=NATS (Docker)
english.ServerNameGeneric=the server
english.SummaryPageTitle=Your settings
english.SummaryPageSubtitle=Review the connection that will be written to NATSExcel.json
english.SummaryIntro=These values will be saved to NATSExcel.json in the install folder.
english.SummaryLabel=Connection
english.SummaryAuth=Authentication
english.SummaryToken=Token
english.SummaryNone=(none)
english.SummaryServer=Server setup

spanish.ServerPageTitle=Servidor NATS
spanish.ServerPageSubtitle=¿Desea instalar ahora un servidor NATS?
spanish.ServerPageDesc=NATSExcel necesita un servidor NATS. NATS se ejecuta de forma nativa en Windows (un nats-server.exe) y JetStream es el mismo binario con -js, por lo que no se necesita Docker. Docker se ofrece como comodidad. Primero se instala el servidor, luego NATSExcel, y la conexión se prueba al final.
spanish.ServerNone=No - ya tengo un servidor (o lo instalaré después)
spanish.ServerNative=Sí - Windows nativo (descargar nats-server.exe, recomendado)
spanish.ServerDocker=Sí - Docker (nats:latest, JetStream activado)
spanish.ConnPageTitle=Conexión a NATS
spanish.ConnPageSubtitle=¿Dónde debe conectarse NATSExcel?
spanish.ConnPageDesc=Pulse Siguiente para el valor predeterminado (nats://localhost:4222). Para un clúster, separe las URL con comas: con una basta, el cliente descubre el resto. Puede cambiarlo luego editando NATSExcel.json.
spanish.ConnUrl=URL del servidor (separadas por comas para un clúster):
spanish.ConnToken=Token (opcional):
spanish.ConnTokenHint=Un token es un único secreto compartido. Para usuario/contraseña, NKey, JWT o un archivo creds, edite NATSExcel.json después de instalar.
spanish.OptionsPageTitle=Comprobaciones
spanish.OptionsPageSubtitle=Casi listo
spanish.OptionsPageDesc=El instalador prueba la conexión al terminar. Opcionalmente abre un libro de Excel de ejemplo.
spanish.OptionsOpenExcel=Abrir un libro de NATSExcel de ejemplo al final
spanish.OptionsService=Registrar NATS como servicio de Windows (requiere admin; se omite si no hay elevación)
spanish.TaskConfigDesc=Crear NATSExcel.json en la carpeta de instalación
spanish.TaskNlogDesc=Crear NLog.config (registro junto al complemento)
spanish.TaskAutoLoadDesc=Registrar el complemento para que Excel lo cargue automáticamente (recomendado)
spanish.GroupConfig=Configuración:
spanish.GroupExcel=Excel:
spanish.IconUninstall=Desinstalar {#AppName}
spanish.UrlRequired=Introduzca una URL de NATS (por ejemplo nats://localhost:4222).
spanish.ServerFailed=No se pudo configurar %1 automáticamente (salida %2).
spanish.ConfigWriteFailed=No se pudo escribir %1. Créelo manualmente desde el ejemplo (NATSExcel.json).
spanish.TestFailed=Instalado, pero el NATS en %1 no respondió. Compruebe que el servidor esté en ejecución y ajuste NATSExcel.json si está en otra URL.
spanish.TestOk=¡Todo listo! NATS respondió en %1.%n%nPara usarlo: abra Excel. Si el complemento no se carga solo, vaya a Archivo > Opciones > Complementos > Ir... > Examinar... y elija el .xll en %2.
spanish.FinishedText=NATSExcel está instalado en:%n%2%n%nSi el complemento no se carga solo, actívelo en Excel:%nArchivo > Opciones > Complementos > Ir... > Examinar... y elija NATSExcel64-packed.xll (o NATSExcel-packed.xll para Office de 32 bits).%n%nLuego pruebe =RTD("NatsRtd",,"SUB","test.>") en una celda.
spanish.UninstServersTitle=Servidores
spanish.UninstServersSubtitle=¿Eliminar también un servidor NATS?
spanish.UninstServersDesc=Se encontraron estos servidores NATS en esta máquina. Marque los que desee eliminar.%nDe forma predeterminada NO se elimina nada.
spanish.UninstServerFound=%1  [encontrado]
spanish.UninstKeepConfig=¿Conservar su configuración NATSExcel.json?
spanish.ServerNameNative=NATS (nats-server.exe nativo)
spanish.ServerNameDocker=NATS (Docker)
spanish.ServerNameGeneric=el servidor
spanish.SummaryPageTitle=Sus ajustes
spanish.SummaryPageSubtitle=Revise la conexión que se guardará en NATSExcel.json
spanish.SummaryIntro=Estos valores se guardarán en NATSExcel.json en la carpeta de instalación.
spanish.SummaryLabel=Conexión
spanish.SummaryAuth=Autenticación
spanish.SummaryToken=Token
spanish.SummaryNone=(ninguna)
spanish.SummaryServer=Servidor

brazilianportuguese.ServerPageTitle=Servidor NATS
brazilianportuguese.ServerPageSubtitle=Deseja instalar agora um servidor NATS?
brazilianportuguese.ServerPageDesc=O NATSExcel precisa de um servidor NATS. O NATS roda nativo no Windows (um nats-server.exe) e o JetStream é o mesmo binário com -js, então não precisa de Docker. O Docker é oferecido como conveniência. O servidor é instalado primeiro, depois o NATSExcel, e a conexão é testada no final.
brazilianportuguese.ServerNone=Não - eu já tenho um servidor (ou vou instalar depois)
brazilianportuguese.ServerNative=Sim - Windows nativo (baixar nats-server.exe, recomendado)
brazilianportuguese.ServerDocker=Sim - Docker (nats:latest, JetStream ligado)
brazilianportuguese.ConnPageTitle=Conexão com o NATS
brazilianportuguese.ConnPageSubtitle=Onde o NATSExcel deve se conectar?
brazilianportuguese.ConnPageDesc=Clique em Avançar para o padrão (nats://localhost:4222). Para um cluster, separe as URLs por vírgula - uma basta, o cliente descobre o resto. Você pode mudar depois editando o NATSExcel.json.
brazilianportuguese.ConnUrl=URL do servidor (separadas por vírgula para um cluster):
brazilianportuguese.ConnToken=Token (opcional):
brazilianportuguese.ConnTokenHint=Um token é um único segredo compartilhado. Para usuário/senha, NKey, JWT ou arquivo creds, edite o NATSExcel.json depois de instalar.
brazilianportuguese.OptionsPageTitle=Verificações
brazilianportuguese.OptionsPageSubtitle=Quase pronto
brazilianportuguese.OptionsPageDesc=O instalador testa a conexão ao terminar. Opcionalmente abre uma planilha de exemplo.
brazilianportuguese.OptionsOpenExcel=Abrir uma planilha de exemplo do NATSExcel no final
brazilianportuguese.OptionsService=Registrar o NATS como serviço do Windows (precisa de admin; ignorado se não elevado)
brazilianportuguese.TaskConfigDesc=Criar NATSExcel.json na pasta de instalação
brazilianportuguese.TaskNlogDesc=Criar NLog.config (log junto ao add-in)
brazilianportuguese.TaskAutoLoadDesc=Registrar o add-in para o Excel carregá-lo automaticamente (recomendado)
brazilianportuguese.GroupConfig=Configuração:
brazilianportuguese.GroupExcel=Excel:
brazilianportuguese.IconUninstall=Desinstalar {#AppName}
brazilianportuguese.UrlRequired=Informe uma URL do NATS (por exemplo nats://localhost:4222).
brazilianportuguese.ServerFailed=Não foi possível configurar %1 automaticamente (saída %2).
brazilianportuguese.ConfigWriteFailed=Não foi possível gravar %1. Crie manualmente a partir do exemplo (NATSExcel.json).
brazilianportuguese.TestFailed=Instalado, mas o NATS em %1 não respondeu. Confira se o servidor está rodando e ajuste o NATSExcel.json se estiver em outra URL.
brazilianportuguese.TestOk=Tudo certo! O NATS respondeu em %1.%n%nPara usar: abra o Excel. Se o add-in não carregar sozinho, vá em Arquivo > Opções > Suplementos > Ir... > Procurar... e escolha o .xll em %2.
brazilianportuguese.FinishedText=O NATSExcel foi instalado em:%n%2%n%nSe o add-in não carregar sozinho, habilite no Excel:%nArquivo > Opções > Suplementos > Ir... > Procurar... e escolha NATSExcel64-packed.xll (ou NATSExcel-packed.xll para Office 32 bits).%n%nDepois teste =RTD("NatsRtd",,"SUB","test.>") em uma célula.
brazilianportuguese.UninstServersTitle=Servidores
brazilianportuguese.UninstServersSubtitle=Remover também um servidor NATS?
brazilianportuguese.UninstServersDesc=Estes servidores NATS foram encontrados na máquina. Marque os que deseja remover.%nPor padrão NADA é removido.
brazilianportuguese.UninstServerFound=%1  [encontrado]
brazilianportuguese.UninstKeepConfig=Manter suas configurações do NATSExcel.json?
brazilianportuguese.ServerNameNative=NATS (nats-server.exe nativo)
brazilianportuguese.ServerNameDocker=NATS (Docker)
brazilianportuguese.ServerNameGeneric=o servidor
brazilianportuguese.SummaryPageTitle=Suas configuracoes
brazilianportuguese.SummaryPageSubtitle=Confira a conexao que sera gravada no NATSExcel.json
brazilianportuguese.SummaryIntro=Estes valores serao salvos no NATSExcel.json na pasta de instalacao.
brazilianportuguese.SummaryLabel=Conexao
brazilianportuguese.SummaryAuth=Autenticacao
brazilianportuguese.SummaryToken=Token
brazilianportuguese.SummaryNone=(nenhuma)
brazilianportuguese.SummaryServer=Servidor

[Tasks]
Name: "addinconfig"; Description: "{cm:TaskConfigDesc}"; GroupDescription: "{cm:GroupConfig}"
Name: "nlogconfig";  Description: "{cm:TaskNlogDesc}"; GroupDescription: "{cm:GroupConfig}"
Name: "excelautoload"; Description: "{cm:TaskAutoLoadDesc}"; GroupDescription: "{cm:GroupExcel}"

[Icons]
Name: "{group}\{cm:IconUninstall}"; Filename: "{uninstallexe}"

[Files]
Source: "{#BuildDir}\NATSExcel.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#BuildDir}\NATSExcel64-packed.xll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#BuildDir}\NATSExcel-packed.xll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#RepoRoot}\NLog.config"; DestDir: "{app}"; Flags: onlyifdoesntexist; Tasks: nlogconfig
Source: "{#RepoRoot}\installer\install-server.ps1"; DestDir: "{app}\scripts"; Flags: ignoreversion
Source: "{#RepoRoot}\installer\nats-service.ps1"; DestDir: "{app}\scripts"; Flags: ignoreversion
Source: "{#RepoRoot}\installer\test-connection.ps1"; DestDir: "{app}\scripts"; Flags: ignoreversion
Source: "{#RepoRoot}\installer\verify-connection.ps1"; DestDir: "{app}\scripts"; Flags: ignoreversion
Source: "{#RepoRoot}\installer\detect-servers.ps1"; DestDir: "{app}\scripts"; Flags: ignoreversion
Source: "{#RepoRoot}\installer\uninstall-server.ps1"; DestDir: "{app}\scripts"; Flags: ignoreversion
Source: "{#RepoRoot}\installer\validate-config.ps1"; DestDir: "{app}\scripts"; Flags: ignoreversion
Source: "{#RepoRoot}\NATSExcel.schema.json"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#RepoRoot}\installer\Validate-Config.cmd"; DestDir: "{app}"; Flags: ignoreversion
; A temporary copy, so the server can be installed BEFORE the files are extracted.
Source: "{#RepoRoot}\installer\install-server.ps1"; DestDir: "{tmp}"; Flags: dontcopy
Source: "{#RepoRoot}\installer\nats-service.ps1"; DestDir: "{tmp}"; Flags: dontcopy

[UninstallDelete]
Type: files; Name: "{app}\NATSExcel.log"
Type: filesandordirs; Name: "{app}\scripts"

[Code]
var
  Is64BitOffice: Boolean;
  ConnPage: TInputQueryWizardPage;
  ServerPage: TInputOptionWizardPage;
  OptionsPage: TInputOptionWizardPage;
  SummaryPage: TOutputMsgMemoWizardPage;
  ServerInstallCode: Integer;
  TestOk: Boolean;
  InstallCompleted: Boolean;
  UninstallServersPage: TInputOptionWizardPage;
  ServerKinds: TArrayOfString;
  ServerLabels: TArrayOfString;
  ServerPresent: TArrayOfBoolean;

// ---- Office bitness / versions --------------------------------------------
function Detect64BitOffice(): Boolean;
var
  Names: TArrayOfString;
  I: Integer;
  V, P: String;
begin
  Result := True;
  P := 'SOFTWARE\Microsoft\Office\ClickToRun\Configuration';
  if RegQueryStringValue(HKLM, P, 'Platform', V) then
    Result := (CompareText(V, 'x64') = 0)
  else if RegQueryStringValue(HKCU, P, 'Platform', V) then
    Result := (CompareText(V, 'x64') = 0)
  else if RegGetSubkeyNames(HKCU, 'SOFTWARE\Microsoft\Office', Names) then
  begin
    for I := 0 to GetArrayLength(Names) - 1 do
      if RegQueryStringValue(HKCU,
           'SOFTWARE\Microsoft\Office\' + Names[I] + '\Common\General', 'Bitness', V) then
      begin
        Result := (CompareText(V, 'x64') = 0);
        Exit;
      end;
  end;
end;

function OfficeExcelKeys(): TArrayOfString;
var
  Names: TArrayOfString;
  I: Integer;
  Acc: TArrayOfString;
begin
  SetArrayLength(Acc, 0);
  if RegGetSubkeyNames(HKCU, 'SOFTWARE\Microsoft\Office', Names) then
    for I := 0 to GetArrayLength(Names) - 1 do
      if (Length(Names[I]) > 0) and (Names[I][1] >= '0') and (Names[I][1] <= '9') then
      begin
        SetArrayLength(Acc, GetArrayLength(Acc) + 1);
        Acc[GetArrayLength(Acc) - 1] := Names[I];
      end;
  Result := Acc;
end;

// ---- Register the XLL as an Excel auto-open add-in -------------------------
function RegisterExcelAutoLoad(const XllPath: String): Boolean;
var
  Keys: TArrayOfString;
  I, N: Integer;
  BaseKey, ValName, Cur: String;
  Found: Boolean;
begin
  Result := False;
  Keys := OfficeExcelKeys();
  for I := 0 to GetArrayLength(Keys) - 1 do
  begin
    BaseKey := 'SOFTWARE\Microsoft\Office\' + Keys[I] + '\Excel\Options';
    Found := False;
    if RegQueryStringValue(HKCU, BaseKey, 'OPEN', Cur) and (CompareText(Cur, XllPath) = 0) then
      Found := True;
    N := 1;
    while (not Found) and (N < 100) do
    begin
      ValName := 'OPEN' + IntToStr(N);
      if not RegQueryStringValue(HKCU, BaseKey, ValName, Cur) then
      begin
        if RegWriteStringValue(HKCU, BaseKey, ValName, XllPath) then Result := True;
        Found := True;
      end
      else if CompareText(Cur, XllPath) = 0 then
        Found := True;
      N := N + 1;
    end;
    if Found and (not Result) then Result := True;
  end;
end;

procedure UnregisterExcelAutoLoad();
var
  Keys: TArrayOfString;
  I, N: Integer;
  BaseKey, ValName, Cur: String;
begin
  Keys := OfficeExcelKeys();
  for I := 0 to GetArrayLength(Keys) - 1 do
  begin
    BaseKey := 'SOFTWARE\Microsoft\Office\' + Keys[I] + '\Excel\Options';
    if RegQueryStringValue(HKCU, BaseKey, 'OPEN', Cur) and (Pos('NATSExcel', Cur) > 0) then
      RegDeleteValue(HKCU, BaseKey, 'OPEN');
    N := 1;
    while N < 100 do
    begin
      ValName := 'OPEN' + IntToStr(N);
      if RegQueryStringValue(HKCU, BaseKey, ValName, Cur) then
      begin
        if Pos('NATSExcel', Cur) > 0 then RegDeleteValue(HKCU, BaseKey, ValName);
      end;
      N := N + 1;
    end;
  end;
end;

// ---- Wizard pages ---------------------------------------------------------
function ParamOr(Param, Default: String): String;
begin
  Result := ExpandConstant('{param:' + Param + '|' + Default + '}');
end;

procedure InitializeWizard();
begin
  Is64BitOffice := Detect64BitOffice();

  ServerPage := CreateInputOptionPage(wpSelectTasks,
    ExpandConstant('{cm:ServerPageTitle}'),
    ExpandConstant('{cm:ServerPageSubtitle}'),
    ExpandConstant('{cm:ServerPageDesc}'),
    True, False);
  ServerPage.Add(ExpandConstant('{cm:ServerNone}'));
  ServerPage.Add(ExpandConstant('{cm:ServerNative}'));
  ServerPage.Add(ExpandConstant('{cm:ServerDocker}'));
  ServerPage.SelectedValueIndex := 0;

  ConnPage := CreateInputQueryPage(ServerPage.ID,
    ExpandConstant('{cm:ConnPageTitle}'),
    ExpandConstant('{cm:ConnPageSubtitle}'),
    ExpandConstant('{cm:ConnPageDesc}'));
  ConnPage.Add(ExpandConstant('{cm:ConnUrl}'), False);
  ConnPage.Add(ExpandConstant('{cm:ConnToken}'), True);
  ConnPage.Values[0] := ParamOr('URL', 'nats://localhost:4222');
  ConnPage.Values[1] := ParamOr('TOKEN', '');

  OptionsPage := CreateInputOptionPage(ConnPage.ID,
    ExpandConstant('{cm:OptionsPageTitle}'),
    ExpandConstant('{cm:OptionsPageSubtitle}'),
    ExpandConstant('{cm:OptionsPageDesc}'),
    False, False);
  OptionsPage.Add(ExpandConstant('{cm:OptionsOpenExcel}'));
  OptionsPage.Add(ExpandConstant('{cm:OptionsService}'));
  OptionsPage.Values[0] := (ParamOr('OPENEXCEL', '0') = '1');
  OptionsPage.Values[1] := False;

  SummaryPage := CreateOutputMsgMemoPage(OptionsPage.ID,
    ExpandConstant('{cm:SummaryPageTitle}'),
    ExpandConstant('{cm:SummaryPageSubtitle}'),
    ExpandConstant('{cm:SummaryIntro}'), '');
end;

function ShouldSkipPage(PageID: Integer): Boolean;
begin
  if WizardSilent() then
    Result := (PageID = ServerPage.ID) or (PageID = ConnPage.ID) or (PageID = OptionsPage.ID)
  else
    Result := False;
end;

function NextButtonClick(CurPageID: Integer): Boolean;
begin
  Result := True;
  if (CurPageID = ConnPage.ID) and (Trim(ConnPage.Values[0]) = '') then
  begin
    MsgBox(ExpandConstant('{cm:UrlRequired}'), mbError, MB_OK);
    Result := False;
  end;
end;

// ---- Values ---------------------------------------------------------------
function Url(): String;
begin
  if (not WizardSilent()) and (ConnPage <> nil) then
    Result := Trim(ConnPage.Values[0])
  else
    Result := ParamOr('URL', 'nats://localhost:4222');
end;

function Token(): String;
begin
  if (not WizardSilent()) and (ConnPage <> nil) then
    Result := Trim(ConnPage.Values[1])
  else
    Result := ParamOr('TOKEN', '');
end;

// First host:port of the URL list, for the end-of-install test.
function FirstHostPort(): String;
var u: String; p: Integer;
begin
  u := Url();
  p := Pos(',', u);
  if p > 0 then u := Copy(u, 1, p - 1);
  u := Trim(u);
  StringChangeEx(u, 'nats://', '', True);
  Result := u;
end;

function ServerPort(): String;
var hp: String; p: Integer;
begin
  hp := FirstHostPort();
  p := Pos(':', hp);
  if p > 0 then Result := Copy(hp, p + 1, Length(hp)) else Result := '4222';
end;

function ServerKind(): String;
var idx: Integer;
begin
  if WizardSilent() then begin Result := ParamOr('SERVER', 'none'); Exit; end;
  if ServerPage = nil then begin Result := 'none'; Exit; end;
  idx := ServerPage.SelectedValueIndex;
  case idx of
    1: Result := 'native';
    2: Result := 'docker';
  else
    Result := 'none';
  end;
end;

function OpenExcelChecked(): Boolean;
begin
  if WizardSilent() then
    Result := (ParamOr('OPENEXCEL', '0') = '1')
  else
    Result := (OptionsPage <> nil) and OptionsPage.Values[0];
end;

function ServiceChecked(): Boolean;
begin
  if WizardSilent() then
    Result := False
  else
    Result := (OptionsPage <> nil) and OptionsPage.Values[1];
end;

// The generated config: minimal, with the code's defaults spelled out.
// Comments are accepted, so the template stays self-documenting.
function ConfigJson(): String;
var U, T: String;
begin
  U := Url();
  T := Token();
  Result :=
    '{' + #13#10 +
    '  // NATSExcel.json - comments (// and /* */) are allowed.' + #13#10 +
    '  "Default": {' + #13#10 +
    '    "cluster": "local",' + #13#10 +
    '    "timeout": 1000,' + #13#10 +
    '    "auth": {' + #13#10;
  if T <> '' then
    Result := Result + '      "token": "' + T + '"' + #13#10
  else
    Result := Result + '      "token": null' + #13#10;
  Result := Result +
    '    },' + #13#10 +
    '    "reconnect": { "maxAttempts": 60, "waitMs": 2000, "jitterMs": 100 }' + #13#10 +
    '  },' + #13#10 +
    '  "RTD": {' + #13#10 +
    '    "NatsUpdateRateMs": 1000,' + #13#10 +
    '    "ExcelUpdateRateMs": 100,' + #13#10 +
    '    "MessageCounterThreshold": 10000,' + #13#10 +
    '    "NatsUpdateStyle": "Automatic"' + #13#10 +
    '  },' + #13#10 +
    '  "UDF": { "timeout": 1000 },' + #13#10 +
    '  "Servers": {' + #13#10 +
    '    "srv1": "' + U + '"' + #13#10 +
    '  },' + #13#10 +
    '  "Clusters": {' + #13#10 +
    '    "local": { "urls": ["srv1"] }' + #13#10 +
    '  },' + #13#10 +
    '  "JetStream": { "enabled": true, "defaultTimeoutMs": 5000 },' + #13#10 +
    '  "AllowAdminWrites": false,' + #13#10 +
    '  "SkipRepeatedMessages": true,' + #13#10 +
    '  "CoalesceRealtimeUpdates": true,' + #13#10 +
    '  "ConflationMs": 200,' + #13#10 +
    '  "PublishDedupCacheSize": 10000,' + #13#10 +
    '  "SyncWrite": "fireforget",' + #13#10 +
    '  "AsyncWrites": true,' + #13#10 +
    '  "UpdateCheck": true' + #13#10 +
    '}' + #13#10;
end;

// ---- Run PowerShell helpers ----------------------------------------------
function RunScript(ScriptName, Args: String): Integer;
var
  Script, Params: String;
  Code: Integer;
begin
  Script := ExpandConstant('{app}\scripts\') + ScriptName;
  Params := '-NoProfile -ExecutionPolicy Bypass -File "' + Script + '" ' + Args;
  if Exec('powershell.exe', Params, ExpandConstant('{app}'), SW_SHOW, ewWaitUntilTerminated, Code) then
    Result := Code
  else
    Result := -1;
end;

function ServerName(Kind: String): String;
begin
  if Kind = 'native' then Result := ExpandConstant('{cm:ServerNameNative}')
  else if Kind = 'docker' then Result := ExpandConstant('{cm:ServerNameDocker}')
  else Result := ExpandConstant('{cm:ServerNameGeneric}');
end;

// ---- Steps: server (before files) -> files -> config + test ---------------
function PrepareToInstall(var NeedsRestart: Boolean): String;
var
  Kind, Script, Params, ServerDir: String;
  Code: Integer;
begin
  Result := '';
  Kind := ServerKind();
  if Kind = 'none' then Exit;

  ExtractTemporaryFile('install-server.ps1');
  Script := ExpandConstant('{tmp}\install-server.ps1');
  ServerDir := ExpandConstant('{app}\server');
  Params := '-NoProfile -ExecutionPolicy Bypass -File "' + Script + '" -Kind ' + Kind +
            ' -Port ' + ServerPort() + ' -Dir "' + ServerDir + '"';
  if Exec('powershell.exe', Params, ExpandConstant('{tmp}'), SW_SHOW, ewWaitUntilTerminated, Code) then
    ServerInstallCode := Code
  else
    ServerInstallCode := -1;

  if (ServerInstallCode <> 0) and (not WizardSilent()) then
    MsgBox(FmtMessage(ExpandConstant('{cm:ServerFailed}'), [ServerName(Kind), IntToStr(ServerInstallCode)]), mbInformation, MB_OK);
end;

procedure CurStepChanged(CurStep: TSetupStep);
var
  ConfigPath, Lines, SvcArgs: String;
  Code: Integer;
begin
  if CurStep = ssPostInstall then
  begin
    if Is64BitOffice then
      DeleteFile(ExpandConstant('{app}\NATSExcel-packed.xll'))
    else
      DeleteFile(ExpandConstant('{app}\NATSExcel64-packed.xll'));

    if WizardIsTaskSelected('addinconfig') then
    begin
      ConfigPath := ExpandConstant('{app}\NATSExcel.json');
      if not FileExists(ConfigPath) then
      begin
        Lines := ConfigJson();
        if not SaveStringToFile(ConfigPath, Lines, False) then
          MsgBox(FmtMessage(ExpandConstant('{cm:ConfigWriteFailed}'), [ConfigPath]),
                 mbInformation, MB_OK);
      end;
    end;

    if WizardIsTaskSelected('excelautoload') then
    begin
      if Is64BitOffice then
        RegisterExcelAutoLoad(ExpandConstant('{app}\NATSExcel64-packed.xll'))
      else
        RegisterExcelAutoLoad(ExpandConstant('{app}\NATSExcel-packed.xll'));
    end;

    // Optional Windows service (native install only). Needs admin; if the
    // wizard is not elevated the helper exits 2 and nothing is registered.
    if ServiceChecked() and (ServerKind() = 'native') then
    begin
      ExtractTemporaryFile('nats-service.ps1');
      SvcArgs := '-NoProfile -ExecutionPolicy Bypass -File "' + ExpandConstant('{tmp}\nats-service.ps1') +
                 '" -Action Install -Dir "' + ExpandConstant('{app}\server') + '"';
      Exec('powershell.exe', SvcArgs, ExpandConstant('{app}'), SW_SHOW, ewWaitUntilTerminated, Code);
    end;

    InstallCompleted := True;
  end;
end;

// ---- Finish: TEST the configuration, then optionally open a sample sheet ---
procedure DeinitializeSetup();
var
  HostPort, Args, Msg: String;
begin
  if not InstallCompleted then Exit;

  HostPort := FirstHostPort();
  Args := '-ServerHost "' + HostPort + '"';
  if Token() <> '' then Args := Args + ' -Token "' + Token() + '"';
  TestOk := (RunScript('test-connection.ps1', Args) = 0);

  if not TestOk then
    Msg := FmtMessage(ExpandConstant('{cm:TestFailed}'), [HostPort])
  else
    Msg := FmtMessage(ExpandConstant('{cm:TestOk}'), [HostPort, ExpandConstant('{app}')]);

  if not WizardSilent() then
    MsgBox(Msg, mbInformation, MB_OK);

  if OpenExcelChecked() then
    RunScript('verify-connection.ps1',
      '-AddInDir "' + ExpandConstant('{app}') + '" -ServerHost "' + HostPort + '" -OpenExcel -Quiet');
end;

procedure CurPageChanged(CurPageID: Integer);
var
  Lines, AuthText: String;
begin
  if (CurPageID = SummaryPage.ID) and (not WizardSilent()) then
  begin
    if Token() <> '' then AuthText := ExpandConstant('{cm:SummaryToken}')
    else AuthText := ExpandConstant('{cm:SummaryNone}');
    Lines :=
      ExpandConstant('{cm:SummaryLabel}') + ':  ' + Url() + #13#10 +
      ExpandConstant('{cm:SummaryAuth}') + ':  ' + AuthText + #13#10 +
      ExpandConstant('{cm:SummaryServer}') + ':  ' + ServerKind();
    SummaryPage.RichEditViewer.Lines.Text := Lines;
  end;

  if (CurPageID = wpFinished) and (not WizardSilent()) then
    WizardForm.FinishedLabel.Caption :=
      FmtMessage(ExpandConstant('{cm:FinishedText}'), ['', ExpandConstant('{app}')]);
end;

// ---- Uninstall: detect servers, let the user pick what to remove ----------
function DetectServers(): Integer;
var
  Script, Params, TmpFile, Line: String;
  Code, I, Count, Sep: Integer;
  Lines: TArrayOfString;
  KindS, LblS, Present: String;
begin
  SetArrayLength(ServerKinds, 0);
  SetArrayLength(ServerLabels, 0);
  SetArrayLength(ServerPresent, 0);

  TmpFile := ExpandConstant('{tmp}\natsexcel-servers.txt');
  Script := ExpandConstant('{app}\scripts\detect-servers.ps1');
  Params := '-NoProfile -ExecutionPolicy Bypass -File "' + Script + '" > "' + TmpFile + '" 2>nul';
  if not Exec('powershell.exe', Params, ExpandConstant('{app}'), SW_HIDE, ewWaitUntilTerminated, Code) then
  begin
    Result := 0;
    Exit;
  end;

  if not LoadStringsFromFile(TmpFile, Lines) then
  begin
    Result := 0;
    Exit;
  end;

  Count := 0;
  for I := 0 to GetArrayLength(Lines) - 1 do
  begin
    Line := Trim(Lines[I]);
    if Line <> '' then
    begin
      Sep := Pos('|', Line);
      if Sep > 0 then
      begin
        KindS := Copy(Line, 1, Sep - 1);
        Line := Copy(Line, Sep + 1, Length(Line));
        Sep := Pos('|', Line);
        if Sep > 0 then
        begin
          LblS := Copy(Line, 1, Sep - 1);
          Present := Trim(Copy(Line, Sep + 1, Length(Line)));

          SetArrayLength(ServerKinds, Count + 1);
          SetArrayLength(ServerLabels, Count + 1);
          SetArrayLength(ServerPresent, Count + 1);
          ServerKinds[Count] := KindS;
          ServerLabels[Count] := LblS;
          ServerPresent[Count] := (Present = '1');
          Count := Count + 1;
        end;
      end;
    end;
  end;
  Result := Count;
end;

function InitializeUninstall(): Boolean;
var
  I, Present: Integer;
begin
  Result := True;
  UninstallServersPage := nil;
  try
    DetectServers();
    Present := 0;
    for I := 0 to GetArrayLength(ServerPresent) - 1 do
      if ServerPresent[I] then Present := Present + 1;
    if Present > 0 then
    begin
      UninstallServersPage := CreateInputOptionPage(1,
        ExpandConstant('{cm:UninstServersTitle}'),
        ExpandConstant('{cm:UninstServersSubtitle}'),
        ExpandConstant('{cm:UninstServersDesc}'),
        False, True);
      for I := 0 to GetArrayLength(ServerKinds) - 1 do
        UninstallServersPage.Add(FmtMessage(ExpandConstant('{cm:UninstServerFound}'), [ServerLabels[I]]));
    end;
  except
    UninstallServersPage := nil;
  end;
end;

procedure RunUninstallServers();
var
  I: Integer;
begin
  if UninstallServersPage = nil then Exit;
  for I := 0 to GetArrayLength(ServerKinds) - 1 do
    if UninstallServersPage.Values[I] then
      RunScript('uninstall-server.ps1', '-Kind ' + ServerKinds[I] + ' -Dir "' + ExpandConstant('{app}\server') + '"');
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
  if CurUninstallStep = usUninstall then
  begin
    UnregisterExcelAutoLoad();
    // The add-in registers this ProgId under HKCU on load; remove it.
    RegDeleteKeyIncludingSubkeys(HKCU, 'Software\Classes\NatsRtd');
  end
  else if CurUninstallStep = usPostUninstall then
  begin
    RunUninstallServers();

    if FileExists(ExpandConstant('{app}\NATSExcel.json')) then
      if MsgBox(ExpandConstant('{cm:UninstKeepConfig}'), mbConfirmation, MB_YESNO) = IDNO then
        DeleteFile(ExpandConstant('{app}\NATSExcel.json'));
  end;
end;
