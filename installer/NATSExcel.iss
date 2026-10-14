; NATSExcel installer + uninstaller (Inno Setup 6).
;
; Build:  ISCC /DAppVersion=1.0.0 .\installer\NATSExcel.iss
; CI builds it on every run and attaches NATSExcelSetup-<version>.exe to a
; v* tag release (see .github/workflows/build.yml).
;
; Flow (in this order, as the wizard runs):
;   1. asks whether to install a NATS server (native download or Docker);
;   2. INSTALLS THE SERVER (if chosen);
;   3. asks for the connection (URL) and the credential style (none, token,
;      user/password, NKey seed, JWT + seed, creds file or mTLS);
;   4. INSTALLS NATSExcel (copies the files, keeps the matching XLL, writes
;      NATSExcel.json) and registers the XLL so Excel loads it on start;
;   5. TESTS the connection (visible OK/fail) and, if asked, opens a sample
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
; Silent mode: NATSExcelSetup.exe /VERYSILENT /DIR=... /URL=... /SERVER=...
;   /AUTH=<none|token|userpass|nkey|jwt|creds|mtls> /TOKEN=... /USER=...
;   /PASSWORD=... /NKEYSEED=... /JWT=... /CREDS=... /TLSCERT=... /TLSCERTPASSWORD=...
;   /OPENEXCEL=<0|1>

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
; The PE's FileVersion: without these the installer's ProductVersion is right (it
; comes from AppVersion) but Explorer's Details tab shows an EMPTY "File version"
; and inventory/AV tooling that keys on FileVersion reads 0.0.0.0. Inno pads the
; three-part value to 0.1.2.0. VersionInfoVersion and VersionInfoProductVersion must
; be NUMERIC - Inno rejects "dev" and a "-rc1" suffix - so they come from their own
; define: the build passes the tag's number when it looks like one and 0.0.0.0
; otherwise (AppVersion, which may be "dev", still drives ProductVersion below).
#ifndef AppVersionInfo
#define AppVersionInfo "0.0.0.0"
#endif
VersionInfoVersion={#AppVersionInfo}
VersionInfoProductVersion={#AppVersionInfo}
VersionInfoProductName={#AppName}
VersionInfoDescription={#AppName} Setup
VersionInfoCompany={#AppPublisher}
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
; The wizard is localized in 11 languages. The first (english) is the default.
; The per-language text lives in [CustomMessages] below; every language carries
; the SAME key set, so a missing translation is a hard failure of the key-count
; check, never a silent English string.
;
; The Excel-facing surface is deliberately NOT localized: the UDF names,
; markers and error text are English on purpose so a formula's return value
; does not change with the user's locale (see AGENTS.md golden rule 1 and the
; README "Languages" note).
Name: "english"; MessagesFile: "compiler:Default.isl"
Name: "chinesesimplified"; MessagesFile: "languages\ChineseSimplified.isl"
Name: "chinesetraditional"; MessagesFile: "languages\ChineseTraditional.isl"
Name: "portuguese"; MessagesFile: "compiler:Languages\Portuguese.isl"
Name: "spanish"; MessagesFile: "compiler:Languages\Spanish.isl"
Name: "russian"; MessagesFile: "compiler:Languages\Russian.isl"
Name: "french"; MessagesFile: "compiler:Languages\French.isl"
Name: "korean"; MessagesFile: "compiler:Languages\Korean.isl"
Name: "german"; MessagesFile: "compiler:Languages\German.isl"
Name: "japanese"; MessagesFile: "compiler:Languages\Japanese.isl"
Name: "brazilianportuguese"; MessagesFile: "compiler:Languages\BrazilianPortuguese.isl"

; All user-visible text. 11 languages, every one with the SAME 63 keys
; (english is the source of truth; installer\languages\messages.ps1 reports a
; missing or extra key for a tag).
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
english.AuthPageTitle=Authentication
english.AuthPageSubtitle=How should NATSExcel authenticate?
english.AuthPageDesc=Pick the credential style this server expects. Token is the most common for a simple setup; user/password, NKey, JWT + seed, a creds file and mTLS are all supported. You can change this later by editing NATSExcel.json.
english.AuthNone=No authentication (the server admits anonymous connections)
english.AuthToken=Token (a single shared secret)
english.AuthUserPass=User and password
english.AuthNKey=NKey (the .nk seed file signs the server nonce)
english.AuthJwt=JWT + NKey seed (operator mode / decentralized auth)
english.AuthCreds=Credentials file (.creds: a JWT + seed in one file)
english.AuthMtls=mTLS (a client certificate is the credential; implies TLS)
english.CredsPageTitle=Credentials
english.CredsPageSubtitle=Type the credentials for the style you chose
english.CredsPageDesc=Only the fields for the selected style are shown. Values are written to NATSExcel.json in the install folder; keep that file out of version control.
english.CredsUser=User:
english.CredsPassword=Password:
english.CredsToken=Token:
english.CredsNKeySeed=NKey seed (the text, or a path to the .nk file):
english.CredsJwt=User JWT:
english.CredsCredsFile=Credentials file (.creds) path:
english.CredsTlsCert=Client certificate path (.pfx / .pem):
english.CredsTlsCertPassword=Client certificate password (optional):
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
english.AuthRequired=Please fill in the fields for the authentication style you selected.
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
spanish.AuthPageTitle=Autenticación
spanish.AuthPageSubtitle=¿Cómo debe autenticarse NATSExcel?
spanish.AuthPageDesc=Elija el estilo de credencial que espera este servidor. El token es lo más común en una instalación simple; se admiten usuario/contraseña, NKey, JWT + semilla, un archivo creds y mTLS. Puede cambiarlo luego editando NATSExcel.json.
spanish.AuthNone=Sin autenticación (el servidor admite conexiones anónimas)
spanish.AuthToken=Token (un único secreto compartido)
spanish.AuthUserPass=Usuario y contraseña
spanish.AuthNKey=NKey (el archivo semilla .nk firma el nonce del servidor)
spanish.AuthJwt=JWT + semilla NKey (modo operador / autenticación descentralizada)
spanish.AuthCreds=Archivo de credenciales (.creds: JWT + semilla en un archivo)
spanish.AuthMtls=mTLS (un certificado de cliente es la credencial; implica TLS)
spanish.CredsPageTitle=Credenciales
spanish.CredsPageSubtitle=Escriba las credenciales del estilo elegido
spanish.CredsPageDesc=Solo se muestran los campos del estilo seleccionado. Los valores se escriben en NATSExcel.json en la carpeta de instalación; mantenga ese archivo fuera del control de versiones.
spanish.CredsUser=Usuario:
spanish.CredsPassword=Contraseña:
spanish.CredsToken=Token:
spanish.CredsNKeySeed=Semilla NKey (el texto, o una ruta al archivo .nk):
spanish.CredsJwt=JWT del usuario:
spanish.CredsCredsFile=Ruta del archivo de credenciales (.creds):
spanish.CredsTlsCert=Ruta del certificado de cliente (.pfx / .pem):
spanish.CredsTlsCertPassword=Contraseña del certificado de cliente (opcional):
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
spanish.AuthRequired=Complete los campos del estilo de autenticación que eligió.
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
brazilianportuguese.AuthPageTitle=Autenticação
brazilianportuguese.AuthPageSubtitle=Como o NATSExcel deve se autenticar?
brazilianportuguese.AuthPageDesc=Escolha o estilo de credencial que este servidor espera. Token é o mais comum num setup simples; usuário/senha, NKey, JWT + seed, arquivo creds e mTLS também são suportados. Você pode mudar depois editando o NATSExcel.json.
brazilianportuguese.AuthNone=Sem autenticação (o servidor aceita conexões anônimas)
brazilianportuguese.AuthToken=Token (um único segredo compartilhado)
brazilianportuguese.AuthUserPass=Usuário e senha
brazilianportuguese.AuthNKey=NKey (o arquivo seed .nk assina o nonce do servidor)
brazilianportuguese.AuthJwt=JWT + seed NKey (modo operador / autenticação descentralizada)
brazilianportuguese.AuthCreds=Arquivo de credenciais (.creds: JWT + seed num só arquivo)
brazilianportuguese.AuthMtls=mTLS (um certificado de cliente é a credencial; implica TLS)
brazilianportuguese.CredsPageTitle=Credenciais
brazilianportuguese.CredsPageSubtitle=Digite as credenciais do estilo escolhido
brazilianportuguese.CredsPageDesc=Só aparecem os campos do estilo selecionado. Os valores vão para o NATSExcel.json na pasta de instalação; mantenha esse arquivo fora do controle de versão.
brazilianportuguese.CredsUser=Usuário:
brazilianportuguese.CredsPassword=Senha:
brazilianportuguese.CredsToken=Token:
brazilianportuguese.CredsNKeySeed=Seed NKey (o texto, ou um caminho para o arquivo .nk):
brazilianportuguese.CredsJwt=JWT do usuário:
brazilianportuguese.CredsCredsFile=Caminho do arquivo de credenciais (.creds):
brazilianportuguese.CredsTlsCert=Caminho do certificado de cliente (.pfx / .pem):
brazilianportuguese.CredsTlsCertPassword=Senha do certificado de cliente (opcional):
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
brazilianportuguese.AuthRequired=Preencha os campos do estilo de autenticação que você escolheu.
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
brazilianportuguese.SummaryServer=Servidor

german.ServerPageTitle=NATS-Server
german.ServerPageSubtitle=Möchten Sie jetzt einen NATS-Server installieren?
german.ServerPageDesc=NATSExcel benötigt einen NATS-Server. NATS läuft nativ unter Windows (eine kleine nats-server.exe) und JetStream ist dieselbe Binärdatei mit -js, daher ist Docker nicht nötig. Docker wird nur als Komfort angeboten. Zuerst wird der Server installiert, dann NATSExcel, und die Verbindung wird am Ende getestet.
german.ServerNone=Nein - ich habe bereits einen Server (oder richte später einen ein)
german.ServerNative=Ja - natives Windows (nats-server.exe herunterladen, empfohlen)
german.ServerDocker=Ja - Docker (nats:latest, JetStream an)
german.ConnPageTitle=NATS-Verbindung
german.ConnPageSubtitle=Womit soll NATSExcel sich verbinden?
german.ConnPageDesc=Klicken Sie auf Weiter für die Standardeinstellung (nats://localhost:4222). Trennen Sie bei einem Cluster die URLs durch Kommas - eine genügt, der Client findet die übrigen. Sie können dies später in NATSExcel.json ändern.
german.ConnUrl=Server-URL(s) (kommagetrennt für einen Cluster):
german.AuthPageTitle=Authentifizierung
german.AuthPageSubtitle=Wie soll sich NATSExcel authentifizieren?
german.AuthPageDesc=Wählen Sie den Credential-Stil, den dieser Server erwartet. Token ist bei einer einfachen Einrichtung am häufigsten; Benutzer/Passwort, NKey, JWT + Seed, eine Creds-Datei und mTLS werden ebenfalls unterstützt. Sie können dies später in NATSExcel.json ändern.
german.AuthNone=Keine Authentifizierung (der Server erlaubt anonyme Verbindungen)
german.AuthToken=Token (ein einzelnes gemeinsames Geheimnis)
german.AuthUserPass=Benutzer und Passwort
german.AuthNKey=NKey (die .nk-Seed-Datei signiert die Server-Nonce)
german.AuthJwt=JWT + NKey-Seed (Operator-Modus / dezentrale Authentifizierung)
german.AuthCreds=Credentials-Datei (.creds: ein JWT + Seed in einer Datei)
german.AuthMtls=mTLS (ein Client-Zertifikat ist das Credential; impliziert TLS)
german.CredsPageTitle=Credentials
german.CredsPageSubtitle=Geben Sie die Credentials für den gewählten Stil ein
german.CredsPageDesc=Es werden nur die Felder des ausgewählten Stils angezeigt. Die Werte werden in NATSExcel.json im Installationsordner geschrieben; halten Sie diese Datei aus der Versionsverwaltung heraus.
german.CredsUser=Benutzer:
german.CredsPassword=Passwort:
german.CredsToken=Token:
german.CredsNKeySeed=NKey-Seed (der Text oder ein Pfad zur .nk-Datei):
german.CredsJwt=Benutzer-JWT:
german.CredsCredsFile=Pfad zur Credentials-Datei (.creds):
german.CredsTlsCert=Pfad zum Client-Zertifikat (.pfx / .pem):
german.CredsTlsCertPassword=Passwort des Client-Zertifikats (optional):
german.OptionsPageTitle=Zusätzliche Prüfungen
german.OptionsPageSubtitle=Fast fertig
german.OptionsPageDesc=Der Installer testet die Verbindung am Ende. Optional wird eine Beispiel-Arbeitsmappe in Excel geöffnet, die es zeigt.
german.OptionsOpenExcel=Am Ende eine Beispiel-Arbeitsmappe von NATSExcel öffnen
german.OptionsService=NATS als Windows-Dienst registrieren (erfordert Admin; wird ohne Erhöhung übersprungen)
german.TaskConfigDesc=NATSExcel.json im Installationsordner erstellen (Verbindungseinstellungen)
german.TaskNlogDesc=NLog.config erstellen (eine Protokolldatei neben dem Add-in schreiben)
german.TaskAutoLoadDesc=Das Add-in registrieren, damit Excel es automatisch lädt (empfohlen)
german.GroupConfig=Konfiguration:
german.GroupExcel=Excel:
german.IconUninstall={#AppName} deinstallieren
german.UrlRequired=Bitte geben Sie eine NATS-Server-URL ein (zum Beispiel nats://localhost:4222).
german.AuthRequired=Bitte füllen Sie die Felder für den gewählten Authentifizierungsstil aus.
german.ServerFailed=%1 konnte nicht automatisch eingerichtet werden (Exit %2).%nSie können später einen Server installieren und NATSExcel.json darauf verweisen.
german.ConfigWriteFailed=%1 konnte nicht geschrieben werden. Erstellen Sie sie manuell anhand des Beispiels (NATSExcel.json).
german.TestFailed=Installiert, aber der NATS-Server unter %1 hat nicht geantwortet.%nPrüfen Sie, ob der Server läuft, und passen Sie NATSExcel.json an, falls er unter einer anderen URL läuft.
german.TestOk=Alles bereit! NATS hat unter %1 geantwortet.%n%nSo verwenden Sie es: Öffnen Sie Excel. Falls das Add-in nicht automatisch geladen wird, gehen Sie zu Datei > Optionen > Add-Ins > Gehe zu... > Durchsuchen... und wählen Sie die .xll in %2.
german.FinishedText=NATSExcel ist installiert in:%n%2%n%nFalls das Add-in nicht automatisch geladen wird, aktivieren Sie es in Excel:%nDatei > Optionen > Add-Ins > Gehe zu... > Durchsuchen... und wählen Sie NATSExcel64-packed.xll (oder NATSExcel-packed.xll für 32-Bit-Office).%n%nProbieren Sie dann =RTD("NatsRtd",,"SUB","test.>") in einer Zelle.
german.UninstServersTitle=Server
german.UninstServersSubtitle=Auch einen NATS-Server entfernen?
german.UninstServersDesc=Diese NATS-Server wurden auf diesem Computer gefunden. Markieren Sie die, die Sie mit NATSExcel entfernen möchten.%nStandardmäßig wird NICHTS entfernt.
german.UninstServerFound=%1  [gefunden]
german.UninstKeepConfig=Ihre NATSExcel.json-Einstellungen behalten?
german.ServerNameNative=NATS (natives nats-server.exe)
german.ServerNameDocker=NATS (Docker)
german.ServerNameGeneric=der Server
german.SummaryPageTitle=Ihre Einstellungen
german.SummaryPageSubtitle=Prüfen Sie die Verbindung, die in NATSExcel.json geschrieben wird
german.SummaryIntro=Diese Werte werden in NATSExcel.json im Installationsordner gespeichert.
german.SummaryLabel=Verbindung
german.SummaryAuth=Authentifizierung
german.SummaryServer=Server-Einrichtung

french.ServerPageTitle=Serveur NATS
french.ServerPageSubtitle=Voulez-vous installer un serveur NATS maintenant ?
french.ServerPageDesc=NATSExcel a besoin d'un serveur NATS. NATS s'exécute nativement sous Windows (un petit nats-server.exe) et JetStream est le même binaire avec -js, donc Docker n'est pas nécessaire. Docker est proposé par commodité. Le serveur est installé d'abord, puis NATSExcel, et la connexion est testée à la fin.
french.ServerNone=Non - j'ai déjà un serveur (ou j'en installerai un plus tard)
french.ServerNative=Oui - Windows natif (télécharger nats-server.exe, recommandé)
french.ServerDocker=Oui - Docker (nats:latest, JetStream activé)
french.ConnPageTitle=Connexion NATS
french.ConnPageSubtitle=Où NATSExcel doit-il se connecter ?
french.ConnPageDesc=Cliquez sur Suivant pour la valeur par défaut (nats://localhost:4222). Pour un cluster, séparez les URL par des virgules - une suffit, le client découvre les autres. Vous pouvez modifier cela plus tard en éditant NATSExcel.json.
french.ConnUrl=URL du serveur (séparées par des virgules pour un cluster) :
french.AuthPageTitle=Authentification
french.AuthPageSubtitle=Comment NATSExcel doit-il s'authentifier ?
french.AuthPageDesc=Choisissez le style d'identifiants attendu par ce serveur. Le token est le plus courant pour une installation simple ; utilisateur/mot de passe, NKey, JWT + seed, un fichier creds et mTLS sont tous pris en charge. Vous pouvez modifier cela plus tard en éditant NATSExcel.json.
french.AuthNone=Aucune authentification (le serveur accepte les connexions anonymes)
french.AuthToken=Token (un unique secret partagé)
french.AuthUserPass=Utilisateur et mot de passe
french.AuthNKey=NKey (le fichier seed .nk signe le nonce du serveur)
french.AuthJwt=JWT + seed NKey (mode opérateur / authentification décentralisée)
french.AuthCreds=Fichier d'identifiants (.creds : un JWT + seed dans un seul fichier)
french.AuthMtls=mTLS (un certificat client est l'identifiant ; implique TLS)
french.CredsPageTitle=Identifiants
french.CredsPageSubtitle=Saisissez les identifiants du style choisi
french.CredsPageDesc=Seuls les champs du style sélectionné sont affichés. Les valeurs sont écrites dans NATSExcel.json dans le dossier d'installation ; gardez ce fichier hors du contrôle de version.
french.CredsUser=Utilisateur :
french.CredsPassword=Mot de passe :
french.CredsToken=Token :
french.CredsNKeySeed=Seed NKey (le texte, ou un chemin vers le fichier .nk) :
french.CredsJwt=JWT utilisateur :
french.CredsCredsFile=Chemin du fichier d'identifiants (.creds) :
french.CredsTlsCert=Chemin du certificat client (.pfx / .pem) :
french.CredsTlsCertPassword=Mot de passe du certificat client (facultatif) :
french.OptionsPageTitle=Vérifications supplémentaires
french.OptionsPageSubtitle=Presque terminé
french.OptionsPageDesc=L'installateur teste la connexion à la fin. Il peut ouvrir un classeur Excel d'exemple le montrant en fonctionnement.
french.OptionsOpenExcel=Ouvrir un classeur NATSExcel d'exemple à la fin
french.OptionsService=Enregistrer NATS comme service Windows (nécessite admin ; ignoré si non élevé)
french.TaskConfigDesc=Créer NATSExcel.json dans le dossier d'installation (paramètres de connexion)
french.TaskNlogDesc=Créer NLog.config (écrire un fichier journal à côté de l'add-in)
french.TaskAutoLoadDesc=Enregistrer l'add-in pour qu'Excel le charge automatiquement (recommandé)
french.GroupConfig=Configuration :
french.GroupExcel=Excel :
french.IconUninstall=Désinstaller {#AppName}
french.UrlRequired=Veuillez saisir une URL de serveur NATS (par exemple nats://localhost:4222).
french.AuthRequired=Veuillez remplir les champs du style d'authentification sélectionné.
french.ServerFailed=Impossible de configurer %1 automatiquement (sortie %2).%nVous pouvez installer un serveur plus tard et pointer NATSExcel.json vers celui-ci.
french.ConfigWriteFailed=Impossible d'écrire %1. Créez-le manuellement à partir de l'exemple (NATSExcel.json).
french.TestFailed=Installé, mais le serveur NATS à %1 n'a pas répondu.%nVérifiez que le serveur est en cours d'exécution et ajustez NATSExcel.json s'il est sur une autre URL.
french.TestOk=Tout est prêt ! NATS a répondu à %1.%n%nPour l'utiliser : ouvrez Excel. Si l'add-in ne se charge pas automatiquement, allez dans Fichier > Options > Compléments > Atteindre... > Parcourir... et choisissez le .xll dans %2.
french.FinishedText=NATSExcel est installé dans :%n%2%n%nSi l'add-in ne se charge pas automatiquement, activez-le dans Excel :%nFichier > Options > Compléments > Atteindre... > Parcourir... et choisissez NATSExcel64-packed.xll (ou NATSExcel-packed.xll pour Office 32 bits).%n%nEssayez ensuite =RTD("NatsRtd",,"SUB","test.>") dans une cellule.
french.UninstServersTitle=Serveurs
french.UninstServersSubtitle=Supprimer aussi un serveur NATS ?
french.UninstServersDesc=Ces serveurs NATS ont été trouvés sur cette machine. Cochez ceux que vous voulez supprimer avec NATSExcel.%nPar défaut, RIEN n'est supprimé.
french.UninstServerFound=%1  [trouvé]
french.UninstKeepConfig=Conserver vos paramètres NATSExcel.json ?
french.ServerNameNative=NATS (nats-server.exe natif)
french.ServerNameDocker=NATS (Docker)
french.ServerNameGeneric=le serveur
french.SummaryPageTitle=Vos paramètres
french.SummaryPageSubtitle=Vérifiez la connexion qui sera écrite dans NATSExcel.json
french.SummaryIntro=Ces valeurs seront enregistrées dans NATSExcel.json dans le dossier d'installation.
french.SummaryLabel=Connexion
french.SummaryAuth=Authentification
french.SummaryServer=Installation du serveur

portuguese.ServerPageTitle=Servidor NATS
portuguese.ServerPageSubtitle=Pretende instalar agora um servidor NATS?
portuguese.ServerPageDesc=O NATSExcel precisa de um servidor NATS. O NATS é executado nativamente no Windows (um pequeno nats-server.exe) e o JetStream é o mesmo binário com -js, pelo que não é necessário Docker. O Docker é disponibilizado por conveniência. O servidor é instalado primeiro, depois o NATSExcel, e a ligação é testada no final.
portuguese.ServerNone=Não - já tenho um servidor (ou vou configurá-lo mais tarde)
portuguese.ServerNative=Sim - Windows nativo (transferir nats-server.exe, recomendado)
portuguese.ServerDocker=Sim - Docker (nats:latest, JetStream ativado)
portuguese.ConnPageTitle=Ligação NATS
portuguese.ConnPageSubtitle=Onde deve o NATSExcel ligar-se?
portuguese.ConnPageDesc=Prima Seguinte para a predefinição (nats://localhost:4222). Para um cluster, separe os URLs por vírgulas - um basta, o cliente descobre os restantes. Pode alterar isto mais tarde editando o NATSExcel.json.
portuguese.ConnUrl=URL(s) do servidor (separados por vírgulas para um cluster):
portuguese.AuthPageTitle=Autenticação
portuguese.AuthPageSubtitle=Como deve o NATSExcel autenticar-se?
portuguese.AuthPageDesc=Escolha o estilo de credenciais que este servidor espera. O token é o mais comum numa configuração simples; utilizador/palavra-passe, NKey, JWT + seed, um ficheiro creds e mTLS são todos suportados. Pode alterar isto mais tarde editando o NATSExcel.json.
portuguese.AuthNone=Sem autenticação (o servidor aceita ligações anónimas)
portuguese.AuthToken=Token (um único segredo partilhado)
portuguese.AuthUserPass=Utilizador e palavra-passe
portuguese.AuthNKey=NKey (o ficheiro seed .nk assina o nonce do servidor)
portuguese.AuthJwt=JWT + seed NKey (modo operador / autenticação descentralizada)
portuguese.AuthCreds=Ficheiro de credenciais (.creds: um JWT + seed num só ficheiro)
portuguese.AuthMtls=mTLS (um certificado de cliente é a credencial; implica TLS)
portuguese.CredsPageTitle=Credenciais
portuguese.CredsPageSubtitle=Introduza as credenciais do estilo escolhido
portuguese.CredsPageDesc=Só são mostrados os campos do estilo selecionado. Os valores são escritos no NATSExcel.json na pasta de instalação; mantenha esse ficheiro fora do controlo de versões.
portuguese.CredsUser=Utilizador:
portuguese.CredsPassword=Palavra-passe:
portuguese.CredsToken=Token:
portuguese.CredsNKeySeed=Seed NKey (o texto, ou um caminho para o ficheiro .nk):
portuguese.CredsJwt=JWT do utilizador:
portuguese.CredsCredsFile=Caminho do ficheiro de credenciais (.creds):
portuguese.CredsTlsCert=Caminho do certificado de cliente (.pfx / .pem):
portuguese.CredsTlsCertPassword=Palavra-passe do certificado de cliente (opcional):
portuguese.OptionsPageTitle=Verificações adicionais
portuguese.OptionsPageSubtitle=Quase pronto
portuguese.OptionsPageDesc=O instalador testa a ligação ao terminar. Opcionalmente abre um livro de Excel de exemplo a mostrar o funcionamento.
portuguese.OptionsOpenExcel=Abrir um livro NATSExcel de exemplo no final
portuguese.OptionsService=Registar o NATS como serviço do Windows (requer admin; ignorado se não elevado)
portuguese.TaskConfigDesc=Criar o NATSExcel.json na pasta de instalação (definições de ligação)
portuguese.TaskNlogDesc=Criar o NLog.config (escrever um ficheiro de registo junto ao add-in)
portuguese.TaskAutoLoadDesc=Registar o add-in para que o Excel o carregue automaticamente (recomendado)
portuguese.GroupConfig=Configuração:
portuguese.GroupExcel=Excel:
portuguese.IconUninstall=Desinstalar {#AppName}
portuguese.UrlRequired=Introduza um URL de servidor NATS (por exemplo nats://localhost:4222).
portuguese.AuthRequired=Preencha os campos do estilo de autenticação que selecionou.
portuguese.ServerFailed=Não foi possível configurar %1 automaticamente (saída %2).%nPode instalar um servidor mais tarde e apontar o NATSExcel.json para ele.
portuguese.ConfigWriteFailed=Não foi possível escrever %1. Crie-o manualmente a partir do exemplo (NATSExcel.json).
portuguese.TestFailed=Instalado, mas o servidor NATS em %1 não respondeu.%nVerifique se o servidor está em execução e ajuste o NATSExcel.json se estiver noutro URL.
portuguese.TestOk=Tudo pronto! O NATS respondeu em %1.%n%nPara usar: abra o Excel. Se o add-in não carregar automaticamente, vá a Ficheiro > Opções > Suplementos > Ir... > Procurar... e escolha o .xll em %2.
portuguese.FinishedText=O NATSExcel está instalado em:%n%2%n%nSe o add-in não carregar automaticamente, ative-o no Excel:%nFicheiro > Opções > Suplementos > Ir... > Procurar... e escolha NATSExcel64-packed.xll (ou NATSExcel-packed.xll para Office de 32 bits).%n%nDepois experimente =RTD("NatsRtd",,"SUB","test.>") numa célula.
portuguese.UninstServersTitle=Servidores
portuguese.UninstServersSubtitle=Remover também um servidor NATS?
portuguese.UninstServersDesc=Estes servidores NATS foram encontrados nesta máquina. Assinale os que pretende remover com o NATSExcel.%nPor predefinição NADA é removido.
portuguese.UninstServerFound=%1  [encontrado]
portuguese.UninstKeepConfig=Manter as suas definições do NATSExcel.json?
portuguese.ServerNameNative=NATS (nats-server.exe nativo)
portuguese.ServerNameDocker=NATS (Docker)
portuguese.ServerNameGeneric=o servidor
portuguese.SummaryPageTitle=As suas definições
portuguese.SummaryPageSubtitle=Reveja a ligação que será escrita no NATSExcel.json
portuguese.SummaryIntro=Estes valores serão guardados no NATSExcel.json na pasta de instalação.
portuguese.SummaryLabel=Ligação
portuguese.SummaryAuth=Autenticação
portuguese.SummaryServer=Configuração do servidor

russian.ServerPageTitle=Сервер NATS
russian.ServerPageSubtitle=Установить сервер NATS сейчас?
russian.ServerPageDesc=NATSExcel требует сервер NATS. NATS работает в Windows изначально (небольшой nats-server.exe), а JetStream — это тот же бинарник с -js, поэтому Docker не нужен. Docker предлагается для удобства. Сначала устанавливается сервер, затем NATSExcel, а подключение проверяется в конце.
russian.ServerNone=Нет — у меня уже есть сервер (или настрою позже)
russian.ServerNative=Да — изначально в Windows (загрузить nats-server.exe, рекомендуется)
russian.ServerDocker=Да — Docker (nats:latest, JetStream включён)
russian.ConnPageTitle=Подключение к NATS
russian.ConnPageSubtitle=Куда должен подключаться NATSExcel?
russian.ConnPageDesc=Нажмите «Далее» для значения по умолчанию (nats://localhost:4222). Для кластера разделите URL запятыми — одного достаточно, клиент найдёт остальные. Это можно изменить позже, отредактировав NATSExcel.json.
russian.ConnUrl=URL сервера (через запятую для кластера):
russian.AuthPageTitle=Аутентификация
russian.AuthPageSubtitle=Как NATSExcel должен проходить аутентификацию?
russian.AuthPageDesc=Выберите стиль учётных данных, который ожидает этот сервер. Token чаще всего используется в простой настройке; также поддерживаются пользователь/пароль, NKey, JWT + seed, файл creds и mTLS. Это можно изменить позже, отредактировав NATSExcel.json.
russian.AuthNone=Без аутентификации (сервер допускает анонимные подключения)
russian.AuthToken=Token (один общий секрет)
russian.AuthUserPass=Пользователь и пароль
russian.AuthNKey=NKey (файл seed .nk подписывает nonce сервера)
russian.AuthJwt=JWT + seed NKey (режим оператора / децентрализованная аутентификация)
russian.AuthCreds=Файл учётных данных (.creds: JWT + seed в одном файле)
russian.AuthMtls=mTLS (учётные данные — это клиентский сертификат; подразумевает TLS)
russian.CredsPageTitle=Учётные данные
russian.CredsPageSubtitle=Введите учётные данные для выбранного стиля
russian.CredsPageDesc=Показываются только поля выбранного стиля. Значения записываются в NATSExcel.json в папке установки; не добавляйте этот файл в систему контроля версий.
russian.CredsUser=Пользователь:
russian.CredsPassword=Пароль:
russian.CredsToken=Token:
russian.CredsNKeySeed=Seed NKey (текст или путь к файлу .nk):
russian.CredsJwt=JWT пользователя:
russian.CredsCredsFile=Путь к файлу учётных данных (.creds):
russian.CredsTlsCert=Путь к клиентскому сертификату (.pfx / .pem):
russian.CredsTlsCertPassword=Пароль клиентского сертификата (необязательно):
russian.OptionsPageTitle=Дополнительные проверки
russian.OptionsPageSubtitle=Почти готово
russian.OptionsPageDesc=Установщик проверяет подключение после завершения. При желании открывается пример книги Excel с демонстрацией работы.
russian.OptionsOpenExcel=Открыть пример книги NATSExcel в конце
russian.OptionsService=Зарегистрировать NATS как службу Windows (нужны права администратора; пропускается без повышения прав)
russian.TaskConfigDesc=Создать NATSExcel.json в папке установки (параметры подключения)
russian.TaskNlogDesc=Создать NLog.config (записывать файл журнала рядом с add-in)
russian.TaskAutoLoadDesc=Зарегистрировать add-in, чтобы Excel загружал его автоматически (рекомендуется)
russian.GroupConfig=Конфигурация:
russian.GroupExcel=Excel:
russian.IconUninstall=Удалить {#AppName}
russian.UrlRequired=Введите URL сервера NATS (например, nats://localhost:4222).
russian.AuthRequired=Заполните поля для выбранного стиля аутентификации.
russian.ServerFailed=Не удалось настроить %1 автоматически (код выхода %2).%nВы можете установить сервер позже и указать на него в NATSExcel.json.
russian.ConfigWriteFailed=Не удалось записать %1. Создайте его вручную по образцу (NATSExcel.json).
russian.TestFailed=Установлено, но сервер NATS по адресу %1 не ответил.%nУбедитесь, что сервер запущен, и измените NATSExcel.json, если он находится по другому URL.
russian.TestOk=Всё готово! NATS ответил по адресу %1.%n%nЧтобы использовать: откройте Excel. Если add-in не загружается автоматически, перейдите в Файл > Параметры > Надстройки > Перейти... > Обзор... и выберите .xll в %2.
russian.FinishedText=NATSExcel установлен в:%n%2%n%nЕсли add-in не загружается автоматически, включите его в Excel:%nФайл > Параметры > Надстройки > Перейти... > Обзор... и выберите NATSExcel64-packed.xll (или NATSExcel-packed.xll для 32-разрядного Office).%n%nЗатем попробуйте =RTD("NatsRtd",,"SUB","test.>") в ячейке.
russian.UninstServersTitle=Серверы
russian.UninstServersSubtitle=Также удалить сервер NATS?
russian.UninstServersDesc=На этом компьютере найдены следующие серверы NATS. Отметьте те, которые хотите удалить вместе с NATSExcel.%nПо умолчанию НИЧЕГО не удаляется.
russian.UninstServerFound=%1  [найдено]
russian.UninstKeepConfig=Сохранить ваши настройки NATSExcel.json?
russian.ServerNameNative=NATS (изначальный nats-server.exe)
russian.ServerNameDocker=NATS (Docker)
russian.ServerNameGeneric=сервер
russian.SummaryPageTitle=Ваши параметры
russian.SummaryPageSubtitle=Проверьте подключение, которое будет записано в NATSExcel.json
russian.SummaryIntro=Эти значения будут сохранены в NATSExcel.json в папке установки.
russian.SummaryLabel=Подключение
russian.SummaryAuth=Аутентификация
russian.SummaryServer=Настройка сервера

korean.ServerPageTitle=NATS 서버
korean.ServerPageSubtitle=지금 NATS 서버를 설치하시겠습니까?
korean.ServerPageDesc=NATSExcel에는 NATS 서버가 필요합니다. NATS는 Windows에서 기본 실행되며(작은 nats-server.exe 하나) JetStream은 -js를 붙인 같은 바이너리이므로 Docker가 필요하지 않습니다. Docker는 편의를 위해 제공됩니다. 서버를 먼저 설치하고 그다음 NATSExcel을 설치하며 연결은 마지막에 테스트합니다.
korean.ServerNone=아니요 - 이미 서버가 있습니다(또는 나중에 설정합니다)
korean.ServerNative=예 - Windows 기본(nats-server.exe 다운로드, 권장)
korean.ServerDocker=예 - Docker(nats:latest, JetStream 사용)
korean.ConnPageTitle=NATS 연결
korean.ConnPageSubtitle=NATSExcel이 어디에 연결해야 합니까?
korean.ConnPageDesc=기본값(nats://localhost:4222)을 사용하려면 다음을 누르십시오. 클러스터의 경우 URL을 쉼표로 구분하십시오 - 하나면 충분하며 클라이언트가 나머지를 찾습니다. 나중에 NATSExcel.json을 편집하여 변경할 수 있습니다.
korean.ConnUrl=서버 URL(클러스터의 경우 쉼표로 구분):
korean.AuthPageTitle=인증
korean.AuthPageSubtitle=NATSExcel은 어떻게 인증해야 합니까?
korean.AuthPageDesc=이 서버가 기대하는 자격 증명 방식을 선택하십시오. 간단한 설정에서는 token이 가장 흔하며, 사용자/비밀번호, NKey, JWT + seed, creds 파일, mTLS도 모두 지원됩니다. 나중에 NATSExcel.json을 편집하여 변경할 수 있습니다.
korean.AuthNone=인증 없음(서버가 익명 연결을 허용함)
korean.AuthToken=Token(단일 공유 비밀)
korean.AuthUserPass=사용자 및 비밀번호
korean.AuthNKey=NKey(.nk seed 파일이 서버 nonce에 서명함)
korean.AuthJwt=JWT + NKey seed(운영자 모드 / 분산 인증)
korean.AuthCreds=자격 증명 파일(.creds: JWT + seed가 한 파일에)
korean.AuthMtls=mTLS(클라이언트 인증서가 자격 증명이며 TLS를 암시함)
korean.CredsPageTitle=자격 증명
korean.CredsPageSubtitle=선택한 방식의 자격 증명을 입력하십시오
korean.CredsPageDesc=선택한 방식의 필드만 표시됩니다. 값은 설치 폴더의 NATSExcel.json에 기록되며, 이 파일은 버전 관리에서 제외하십시오.
korean.CredsUser=사용자:
korean.CredsPassword=비밀번호:
korean.CredsToken=Token:
korean.CredsNKeySeed=NKey seed(텍스트 또는 .nk 파일 경로):
korean.CredsJwt=사용자 JWT:
korean.CredsCredsFile=자격 증명 파일(.creds) 경로:
korean.CredsTlsCert=클라이언트 인증서 경로(.pfx / .pem):
korean.CredsTlsCertPassword=클라이언트 인증서 비밀번호(선택):
korean.OptionsPageTitle=추가 확인
korean.OptionsPageSubtitle=거의 완료됨
korean.OptionsPageDesc=설치 프로그램이 완료될 때 연결을 테스트합니다. 원하면 작동을 보여주는 예제 Excel 통합 문서를 엽니다.
korean.OptionsOpenExcel=마지막에 예제 NATSExcel 통합 문서 열기
korean.OptionsService=NATS를 Windows 서비스로 등록(관리자 필요, 권한이 없으면 건너뜀)
korean.TaskConfigDesc=설치 폴더에 NATSExcel.json 만들기(연결 설정)
korean.TaskNlogDesc=NLog.config 만들기(add-in 옆에 로그 파일 기록)
korean.TaskAutoLoadDesc=Excel이 자동으로 로드하도록 add-in 등록(권장)
korean.GroupConfig=구성:
korean.GroupExcel=Excel:
korean.IconUninstall={#AppName} 제거
korean.UrlRequired=NATS 서버 URL을 입력하십시오(예: nats://localhost:4222).
korean.AuthRequired=선택한 인증 방식의 필드를 채우십시오.
korean.ServerFailed=%1을(를) 자동으로 설정할 수 없습니다(종료 %2).%n나중에 서버를 설치하고 NATSExcel.json이 이를 가리키도록 할 수 있습니다.
korean.ConfigWriteFailed=%1을(를) 쓸 수 없습니다. 예제(NATSExcel.json)에서 직접 만드십시오.
korean.TestFailed=설치되었지만 %1의 NATS 서버가 응답하지 않았습니다.%n서버가 실행 중인지 확인하고 다른 URL이라면 NATSExcel.json을 조정하십시오.
korean.TestOk=모두 완료! NATS가 %1에서 응답했습니다.%n%n사용 방법: Excel을 여십시오. add-in이 자동으로 로드되지 않으면 파일 > 옵션 > 추가 기능 > 이동... > 찾아보기...로 가서 %2의 .xll을 선택하십시오.
korean.FinishedText=NATSExcel이 설치되었습니다:%n%2%n%nadd-in이 자동으로 로드되지 않으면 Excel에서 활성화하십시오:%n파일 > 옵션 > 추가 기능 > 이동... > 찾아보기...에서 NATSExcel64-packed.xll(32비트 Office는 NATSExcel-packed.xll)을 선택하십시오.%n%n그런 다음 셀에 =RTD("NatsRtd",,"SUB","test.>")를 시도하십시오.
korean.UninstServersTitle=서버
korean.UninstServersSubtitle=NATS 서버도 제거하시겠습니까?
korean.UninstServersDesc=이 컴퓨터에서 다음 NATS 서버를 찾았습니다. NATSExcel과 함께 제거할 항목을 선택하십시오.%n기본적으로 아무것도 제거되지 않습니다.
korean.UninstServerFound=%1  [발견됨]
korean.UninstKeepConfig=NATSExcel.json 설정을 유지하시겠습니까?
korean.ServerNameNative=NATS(기본 nats-server.exe)
korean.ServerNameDocker=NATS(Docker)
korean.ServerNameGeneric=서버
korean.SummaryPageTitle=설정
korean.SummaryPageSubtitle=NATSExcel.json에 기록될 연결을 검토하십시오
korean.SummaryIntro=이 값은 설치 폴더의 NATSExcel.json에 저장됩니다.
korean.SummaryLabel=연결
korean.SummaryAuth=인증
korean.SummaryServer=서버 설정

japanese.ServerPageTitle=NATS サーバー
japanese.ServerPageSubtitle=今すぐ NATS サーバーをインストールしますか?
japanese.ServerPageDesc=NATSExcel には NATS サーバーが必要です。NATS は Windows でネイティブに動作し(小さな nats-server.exe 1 つ)、JetStream は -js を付けた同じバイナリなので Docker は不要です。Docker は便宜上提供されています。先にサーバーをインストールし、次に NATSExcel をインストールし、最後に接続をテストします。
japanese.ServerNone=いいえ - 既にサーバーがあります(または後で設定します)
japanese.ServerNative=はい - Windows ネイティブ(nats-server.exe をダウンロード、推奨)
japanese.ServerDocker=はい - Docker(nats:latest、JetStream 有効)
japanese.ConnPageTitle=NATS 接続
japanese.ConnPageSubtitle=NATSExcel はどこに接続しますか?
japanese.ConnPageDesc=既定値(nats://localhost:4222)を使うには [次へ] を押します。クラスターの場合は URL をカンマで区切ります - 1 つで十分で、クライアントが残りを検出します。後で NATSExcel.json を編集して変更できます。
japanese.ConnUrl=サーバー URL(クラスターの場合はカンマ区切り):
japanese.AuthPageTitle=認証
japanese.AuthPageSubtitle=NATSExcel はどのように認証しますか?
japanese.AuthPageDesc=このサーバーが期待する資格情報の方式を選びます。単純な構成では token が最も一般的で、ユーザー/パスワード、NKey、JWT + seed、creds ファイル、mTLS もすべてサポートされます。後で NATSExcel.json を編集して変更できます。
japanese.AuthNone=認証なし(サーバーが匿名接続を許可します)
japanese.AuthToken=Token(単一の共有シークレット)
japanese.AuthUserPass=ユーザーとパスワード
japanese.AuthNKey=NKey(.nk seed ファイルがサーバーの nonce に署名します)
japanese.AuthJwt=JWT + NKey seed(オペレーターモード / 分散認証)
japanese.AuthCreds=資格情報ファイル(.creds: JWT + seed を 1 ファイルに)
japanese.AuthMtls=mTLS(クライアント証明書が資格情報。TLS を意味します)
japanese.CredsPageTitle=資格情報
japanese.CredsPageSubtitle=選択した方式の資格情報を入力します
japanese.CredsPageDesc=選択した方式のフィールドだけが表示されます。値はインストール フォルダーの NATSExcel.json に書き込まれます。このファイルはバージョン管理に含めないでください。
japanese.CredsUser=ユーザー:
japanese.CredsPassword=パスワード:
japanese.CredsToken=Token:
japanese.CredsNKeySeed=NKey seed(テキスト、または .nk ファイルへのパス):
japanese.CredsJwt=ユーザー JWT:
japanese.CredsCredsFile=資格情報ファイル(.creds)のパス:
japanese.CredsTlsCert=クライアント証明書のパス(.pfx / .pem):
japanese.CredsTlsCertPassword=クライアント証明書のパスワード(省略可):
japanese.OptionsPageTitle=追加の確認
japanese.OptionsPageSubtitle=ほぼ完了
japanese.OptionsPageDesc=インストーラーは終了時に接続をテストします。必要に応じて、動作を示すサンプル Excel ブックを開きます。
japanese.OptionsOpenExcel=最後にサンプルの NATSExcel ブックを開く
japanese.OptionsService=NATS を Windows サービスとして登録(管理者が必要。昇格していない場合はスキップ)
japanese.TaskConfigDesc=インストール フォルダーに NATSExcel.json を作成(接続設定)
japanese.TaskNlogDesc=NLog.config を作成(add-in の隣にログ ファイルを書き込む)
japanese.TaskAutoLoadDesc=Excel が自動で読み込むように add-in を登録(推奨)
japanese.GroupConfig=構成:
japanese.GroupExcel=Excel:
japanese.IconUninstall={#AppName} をアンインストール
japanese.UrlRequired=NATS サーバー URL を入力してください(例: nats://localhost:4222)。
japanese.AuthRequired=選択した認証方式のフィールドを入力してください。
japanese.ServerFailed=%1 を自動的に設定できませんでした(終了コード %2)。%n後でサーバーをインストールし、NATSExcel.json をそちらに向けることができます。
japanese.ConfigWriteFailed=%1 を書き込めませんでした。サンプル(NATSExcel.json)から手動で作成してください。
japanese.TestFailed=インストールしましたが、%1 の NATS サーバーが応答しませんでした。%nサーバーが実行中か確認し、別の URL にある場合は NATSExcel.json を調整してください。
japanese.TestOk=準備完了! NATS が %1 で応答しました。%n%n使用方法: Excel を開きます。add-in が自動で読み込まれない場合は、ファイル > オプション > アドイン > 設定... > 参照... に進み、%2 の .xll を選びます。
japanese.FinishedText=NATSExcel は次にインストールされています:%n%2%n%nadd-in が自動で読み込まれない場合は、Excel で有効にします:%nファイル > オプション > アドイン > 設定... > 参照... に進み、NATSExcel64-packed.xll(32 ビット Office の場合は NATSExcel-packed.xll)を選びます。%n%n次に、セルで =RTD("NatsRtd",,"SUB","test.>") を試します。
japanese.UninstServersTitle=サーバー
japanese.UninstServersSubtitle=NATS サーバーも削除しますか?
japanese.UninstServersDesc=このマシンで次の NATS サーバーが見つかりました。NATSExcel と一緒に削除するものを選んでください。%n既定では何も削除されません。
japanese.UninstServerFound=%1  [検出]
japanese.UninstKeepConfig=NATSExcel.json の設定を保持しますか?
japanese.ServerNameNative=NATS(ネイティブ nats-server.exe)
japanese.ServerNameDocker=NATS(Docker)
japanese.ServerNameGeneric=サーバー
japanese.SummaryPageTitle=設定
japanese.SummaryPageSubtitle=NATSExcel.json に書き込まれる接続を確認してください
japanese.SummaryIntro=これらの値はインストール フォルダーの NATSExcel.json に保存されます。
japanese.SummaryLabel=接続
japanese.SummaryAuth=認証
japanese.SummaryServer=サーバー セットアップ

chinesesimplified.ServerPageTitle=NATS 服务器
chinesesimplified.ServerPageSubtitle=现在安装 NATS 服务器吗?
chinesesimplified.ServerPageDesc=NATSExcel 需要一个 NATS 服务器。NATS 在 Windows 上原生运行(一个小的 nats-server.exe), JetStream 就是同一个二进制文件加上 -js, 因此不需要 Docker。Docker 只是作为便利选项提供。先安装服务器, 再安装 NATSExcel, 最后测试连接。
chinesesimplified.ServerNone=否 - 我已经有服务器(或稍后自行设置)
chinesesimplified.ServerNative=是 - Windows 原生(下载 nats-server.exe, 推荐)
chinesesimplified.ServerDocker=是 - Docker(nats:latest, 启用 JetStream)
chinesesimplified.ConnPageTitle=NATS 连接
chinesesimplified.ConnPageSubtitle=NATSExcel 应连接到何处?
chinesesimplified.ConnPageDesc=按 [下一步] 使用默认值(nats://localhost:4222)。对于集群, 请用逗号分隔 URL - 一个即可, 客户端会发现其余的。你可以稍后编辑 NATSExcel.json 来更改。
chinesesimplified.ConnUrl=服务器 URL(集群请用逗号分隔):
chinesesimplified.AuthPageTitle=身份验证
chinesesimplified.AuthPageSubtitle=NATSExcel 应如何身份验证?
chinesesimplified.AuthPageDesc=选择此服务器所需的凭据方式。简单设置中 token 最常见; 用户/密码、NKey、JWT + seed、creds 文件和 mTLS 都受支持。你可以稍后编辑 NATSExcel.json 来更改。
chinesesimplified.AuthNone=无身份验证(服务器接受匿名连接)
chinesesimplified.AuthToken=Token(单个共享密钥)
chinesesimplified.AuthUserPass=用户名和密码
chinesesimplified.AuthNKey=NKey(.nk seed 文件对服务器 nonce 签名)
chinesesimplified.AuthJwt=JWT + NKey seed(操作员模式 / 去中心化身份验证)
chinesesimplified.AuthCreds=凭据文件(.creds: 一个文件中包含 JWT + seed)
chinesesimplified.AuthMtls=mTLS(客户端证书即凭据; 隐含 TLS)
chinesesimplified.CredsPageTitle=凭据
chinesesimplified.CredsPageSubtitle=输入所选方式的凭据
chinesesimplified.CredsPageDesc=仅显示所选方式的字段。值会写入安装文件夹中的 NATSExcel.json; 请勿将该文件纳入版本控制。
chinesesimplified.CredsUser=用户名:
chinesesimplified.CredsPassword=密码:
chinesesimplified.CredsToken=Token:
chinesesimplified.CredsNKeySeed=NKey seed(文本, 或 .nk 文件的路径):
chinesesimplified.CredsJwt=用户 JWT:
chinesesimplified.CredsCredsFile=凭据文件(.creds)路径:
chinesesimplified.CredsTlsCert=客户端证书路径(.pfx / .pem):
chinesesimplified.CredsTlsCertPassword=客户端证书密码(可选):
chinesesimplified.OptionsPageTitle=额外检查
chinesesimplified.OptionsPageSubtitle=即将完成
chinesesimplified.OptionsPageDesc=安装程序在结束时测试连接。可选地打开一个演示其工作的示例 Excel 工作簿。
chinesesimplified.OptionsOpenExcel=结束时打开示例 NATSExcel 工作簿
chinesesimplified.OptionsService=将 NATS 注册为 Windows 服务(需要管理员; 未提升权限时跳过)
chinesesimplified.TaskConfigDesc=在安装文件夹中创建 NATSExcel.json(连接设置)
chinesesimplified.TaskNlogDesc=创建 NLog.config(在 add-in 旁边写入日志文件)
chinesesimplified.TaskAutoLoadDesc=注册 add-in, 让 Excel 自动加载(推荐)
chinesesimplified.GroupConfig=配置:
chinesesimplified.GroupExcel=Excel:
chinesesimplified.IconUninstall=卸载 {#AppName}
chinesesimplified.UrlRequired=请输入 NATS 服务器 URL(例如 nats://localhost:4222)。
chinesesimplified.AuthRequired=请填写所选身份验证方式的字段。
chinesesimplified.ServerFailed=无法自动设置 %1(退出 %2)。%n你可以稍后安装服务器并将 NATSExcel.json 指向它。
chinesesimplified.ConfigWriteFailed=无法写入 %1。请参照示例(NATSExcel.json)手动创建。
chinesesimplified.TestFailed=已安装, 但 %1 处的 NATS 服务器没有响应。%n请检查服务器是否运行, 如果它在其他 URL 上, 请调整 NATSExcel.json。
chinesesimplified.TestOk=全部就绪! NATS 在 %1 处作出了响应。%n%n使用方法: 打开 Excel。如果 add-in 没有自动加载, 请转到 文件 > 选项 > 加载项 > 转到... > 浏览..., 选择 %2 中的 .xll。
chinesesimplified.FinishedText=NATSExcel 已安装到:%n%2%n%n如果 add-in 没有自动加载, 请在 Excel 中启用它:%n文件 > 选项 > 加载项 > 转到... > 浏览..., 选择 NATSExcel64-packed.xll(32 位 Office 请选 NATSExcel-packed.xll)。%n%n然后在单元格中尝试 =RTD("NatsRtd",,"SUB","test.>")。
chinesesimplified.UninstServersTitle=服务器
chinesesimplified.UninstServersSubtitle=也删除 NATS 服务器?
chinesesimplified.UninstServersDesc=在此机器上找到了这些 NATS 服务器。勾选要与 NATSExcel 一起删除的项。%n默认不删除任何内容。
chinesesimplified.UninstServerFound=%1  [已找到]
chinesesimplified.UninstKeepConfig=保留你的 NATSExcel.json 设置?
chinesesimplified.ServerNameNative=NATS(原生 nats-server.exe)
chinesesimplified.ServerNameDocker=NATS(Docker)
chinesesimplified.ServerNameGeneric=服务器
chinesesimplified.SummaryPageTitle=你的设置
chinesesimplified.SummaryPageSubtitle=查看将写入 NATSExcel.json 的连接
chinesesimplified.SummaryIntro=这些值将保存到安装文件夹中的 NATSExcel.json。
chinesesimplified.SummaryLabel=连接
chinesesimplified.SummaryAuth=身份验证
chinesesimplified.SummaryServer=服务器设置

chinesetraditional.ServerPageTitle=NATS 伺服器
chinesetraditional.ServerPageSubtitle=現在安裝 NATS 伺服器嗎?
chinesetraditional.ServerPageDesc=NATSExcel 需要一個 NATS 伺服器。NATS 在 Windows 上原生執行(一個小型 nats-server.exe), JetStream 是加上 -js 的同一個二進位檔, 因此不需要 Docker。Docker 僅作為便利選項提供。先安裝伺服器, 再安裝 NATSExcel, 最後測試連線。
chinesetraditional.ServerNone=否 - 我已經有伺服器(或稍後自行設定)
chinesetraditional.ServerNative=是 - Windows 原生(下載 nats-server.exe, 建議)
chinesetraditional.ServerDocker=是 - Docker(nats:latest, 啟用 JetStream)
chinesetraditional.ConnPageTitle=NATS 連線
chinesetraditional.ConnPageSubtitle=NATSExcel 應連線到哪裡?
chinesetraditional.ConnPageDesc=按 [下一步] 使用預設值(nats://localhost:4222)。若為叢集, 請以逗號分隔 URL - 一個即可, 用戶端會探索其餘的。你可以稍後編輯 NATSExcel.json 來變更。
chinesetraditional.ConnUrl=伺服器 URL(叢集請以逗號分隔):
chinesetraditional.AuthPageTitle=驗證
chinesetraditional.AuthPageSubtitle=NATSExcel 應如何驗證?
chinesetraditional.AuthPageDesc=選擇此伺服器所需的認證方式。簡易設定中 token 最常見; 使用者/密碼、NKey、JWT + seed、creds 檔案與 mTLS 皆受支援。你可以稍後編輯 NATSExcel.json 來變更。
chinesetraditional.AuthNone=無驗證(伺服器接受匿名連線)
chinesetraditional.AuthToken=Token(單一共用密鑰)
chinesetraditional.AuthUserPass=使用者名稱與密碼
chinesetraditional.AuthNKey=NKey(.nk seed 檔案會簽署伺服器 nonce)
chinesetraditional.AuthJwt=JWT + NKey seed(操作員模式 / 去中心化驗證)
chinesetraditional.AuthCreds=認證檔案(.creds: JWT + seed 在同一個檔案中)
chinesetraditional.AuthMtls=mTLS(用戶端憑證即為認證; 隱含 TLS)
chinesetraditional.CredsPageTitle=認證
chinesetraditional.CredsPageSubtitle=輸入所選方式的認證
chinesetraditional.CredsPageDesc=僅顯示所選方式的欄位。值會寫入安裝資料夾中的 NATSExcel.json; 請勿將該檔案納入版本控制。
chinesetraditional.CredsUser=使用者名稱:
chinesetraditional.CredsPassword=密碼:
chinesetraditional.CredsToken=Token:
chinesetraditional.CredsNKeySeed=NKey seed(文字, 或 .nk 檔案的路徑):
chinesetraditional.CredsJwt=使用者 JWT:
chinesetraditional.CredsCredsFile=認證檔案(.creds)路徑:
chinesetraditional.CredsTlsCert=用戶端憑證路徑(.pfx / .pem):
chinesetraditional.CredsTlsCertPassword=用戶端憑證密碼(選填):
chinesetraditional.OptionsPageTitle=其他檢查
chinesetraditional.OptionsPageSubtitle=即將完成
chinesetraditional.OptionsPageDesc=安裝程式會在結束時測試連線。可選擇開啟示範其運作的範例 Excel 活頁簿。
chinesetraditional.OptionsOpenExcel=結束時開啟範例 NATSExcel 活頁簿
chinesetraditional.OptionsService=將 NATS 註冊為 Windows 服務(需要管理員; 未提升權限時略過)
chinesetraditional.TaskConfigDesc=在安裝資料夾中建立 NATSExcel.json(連線設定)
chinesetraditional.TaskNlogDesc=建立 NLog.config(在 add-in 旁寫入記錄檔)
chinesetraditional.TaskAutoLoadDesc=註冊 add-in, 讓 Excel 自動載入(建議)
chinesetraditional.GroupConfig=設定:
chinesetraditional.GroupExcel=Excel:
chinesetraditional.IconUninstall=解除安裝 {#AppName}
chinesetraditional.UrlRequired=請輸入 NATS 伺服器 URL(例如 nats://localhost:4222)。
chinesetraditional.AuthRequired=請填寫所選驗證方式的欄位。
chinesetraditional.ServerFailed=無法自動設定 %1(結束碼 %2)。%n你可以稍後安裝伺服器, 並將 NATSExcel.json 指向它。
chinesetraditional.ConfigWriteFailed=無法寫入 %1。請依範例(NATSExcel.json)手動建立。
chinesetraditional.TestFailed=已安裝, 但 %1 的 NATS 伺服器沒有回應。%n請檢查伺服器是否正在執行, 若在其他 URL, 請調整 NATSExcel.json。
chinesetraditional.TestOk=全部就緒! NATS 已在 %1 回應。%n%n使用方法: 開啟 Excel。若 add-in 沒有自動載入, 請前往 檔案 > 選項 > 增益集 > 前往... > 瀏覽..., 選擇 %2 中的 .xll。
chinesetraditional.FinishedText=NATSExcel 已安裝於:%n%2%n%n若 add-in 沒有自動載入, 請在 Excel 中啟用它:%n檔案 > 選項 > 增益集 > 前往... > 瀏覽..., 選擇 NATSExcel64-packed.xll(32 位元 Office 請選 NATSExcel-packed.xll)。%n%n接著在儲存格中試試 =RTD("NatsRtd",,"SUB","test.>")。
chinesetraditional.UninstServersTitle=伺服器
chinesetraditional.UninstServersSubtitle=也要移除 NATS 伺服器嗎?
chinesetraditional.UninstServersDesc=在此電腦上找到這些 NATS 伺服器。勾選要與 NATSExcel 一併移除的項目。%n預設不會移除任何項目。
chinesetraditional.UninstServerFound=%1  [已找到]
chinesetraditional.UninstKeepConfig=保留你的 NATSExcel.json 設定?
chinesetraditional.ServerNameNative=NATS(原生 nats-server.exe)
chinesetraditional.ServerNameDocker=NATS(Docker)
chinesetraditional.ServerNameGeneric=伺服器
chinesetraditional.SummaryPageTitle=你的設定
chinesetraditional.SummaryPageSubtitle=檢閱將寫入 NATSExcel.json 的連線
chinesetraditional.SummaryIntro=這些值將儲存到安裝資料夾中的 NATSExcel.json。
chinesetraditional.SummaryLabel=連線
chinesetraditional.SummaryAuth=驗證
chinesetraditional.SummaryServer=伺服器設定

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
const
  // Credential style indexes (AuthPage radio order).
  AuthIdxNone = 0;
  AuthIdxToken = 1;
  AuthIdxUserPass = 2;
  AuthIdxNKey = 3;
  AuthIdxJwt = 4;
  AuthIdxCreds = 5;
  AuthIdxMtls = 6;

var
  Is64BitOffice: Boolean;
  ConnPage: TInputQueryWizardPage;
  AuthPage: TInputOptionWizardPage;
  CredsPage: TInputQueryWizardPage;
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
  BaseKey, ValName, Cur, Xll64, Xll32: String;
begin
  // Remove ONLY the entry that points at THIS installation's XLL. A blanket
  // "contains NATSExcel" match would also strip a second install (a parallel
  // folder, or a lab copy) from the user's Excel add-ins list.
  Xll64 := ExpandConstant('{app}\NATSExcel64-packed.xll');
  Xll32 := ExpandConstant('{app}\NATSExcel-packed.xll');
  Keys := OfficeExcelKeys();
  for I := 0 to GetArrayLength(Keys) - 1 do
  begin
    BaseKey := 'SOFTWARE\Microsoft\Office\' + Keys[I] + '\Excel\Options';
    if RegQueryStringValue(HKCU, BaseKey, 'OPEN', Cur) and
       ((CompareText(Cur, Xll64) = 0) or (CompareText(Cur, Xll32) = 0)) then
      RegDeleteValue(HKCU, BaseKey, 'OPEN');
    N := 1;
    while N < 100 do
    begin
      ValName := 'OPEN' + IntToStr(N);
      if RegQueryStringValue(HKCU, BaseKey, ValName, Cur) then
      begin
        if (CompareText(Cur, Xll64) = 0) or (CompareText(Cur, Xll32) = 0) then
          RegDeleteValue(HKCU, BaseKey, ValName);
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

// Forward declarations (Pascal requires a definition before use).
function AuthStyleFromParam(): Integer; forward;
function AuthStyle(): Integer; forward;
function Cred(Field: Integer): String; forward;

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
  ConnPage.Values[0] := ParamOr('URL', 'nats://localhost:4222');

  AuthPage := CreateInputOptionPage(ConnPage.ID,
    ExpandConstant('{cm:AuthPageTitle}'),
    ExpandConstant('{cm:AuthPageSubtitle}'),
    ExpandConstant('{cm:AuthPageDesc}'),
    True, False);
  AuthPage.Add(ExpandConstant('{cm:AuthNone}'));
  AuthPage.Add(ExpandConstant('{cm:AuthToken}'));
  AuthPage.Add(ExpandConstant('{cm:AuthUserPass}'));
  AuthPage.Add(ExpandConstant('{cm:AuthNKey}'));
  AuthPage.Add(ExpandConstant('{cm:AuthJwt}'));
  AuthPage.Add(ExpandConstant('{cm:AuthCreds}'));
  AuthPage.Add(ExpandConstant('{cm:AuthMtls}'));

  CredsPage := CreateInputQueryPage(AuthPage.ID,
    ExpandConstant('{cm:CredsPageTitle}'),
    ExpandConstant('{cm:CredsPageSubtitle}'),
    ExpandConstant('{cm:CredsPageDesc}'));
  CredsPage.Add(ExpandConstant('{cm:CredsUser}'), False);              // 0
  CredsPage.Add(ExpandConstant('{cm:CredsPassword}'), True);           // 1
  CredsPage.Add(ExpandConstant('{cm:CredsToken}'), True);              // 2
  CredsPage.Add(ExpandConstant('{cm:CredsNKeySeed}'), False);          // 3
  CredsPage.Add(ExpandConstant('{cm:CredsJwt}'), True);                // 4
  CredsPage.Add(ExpandConstant('{cm:CredsCredsFile}'), False);         // 5
  CredsPage.Add(ExpandConstant('{cm:CredsTlsCert}'), False);           // 6
  CredsPage.Add(ExpandConstant('{cm:CredsTlsCertPassword}'), True);    // 7

  // Seed from silent parameters, then default to no auth.
  CredsPage.Values[0] := ParamOr('USER', '');
  CredsPage.Values[1] := ParamOr('PASSWORD', '');
  CredsPage.Values[2] := ParamOr('TOKEN', '');
  CredsPage.Values[3] := ParamOr('NKEYSEED', '');
  CredsPage.Values[4] := ParamOr('JWT', '');
  CredsPage.Values[5] := ParamOr('CREDS', '');
  CredsPage.Values[6] := ParamOr('TLSCERT', '');
  CredsPage.Values[7] := ParamOr('TLSCERTPASSWORD', '');
  AuthPage.SelectedValueIndex := AuthStyleFromParam();

  OptionsPage := CreateInputOptionPage(CredsPage.ID,
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

function AuthStyleFromParam(): Integer;
var s: String;
begin
  s := Lowercase(ParamOr('AUTH', 'none'));
  if s = 'token' then Result := AuthIdxToken
  else if (s = 'userpass') or (s = 'user') then Result := AuthIdxUserPass
  else if s = 'nkey' then Result := AuthIdxNKey
  else if s = 'jwt' then Result := AuthIdxJwt
  else if s = 'creds' then Result := AuthIdxCreds
  else if s = 'mtls' then Result := AuthIdxMtls
  else Result := AuthIdxNone;
end;

// Which credential fields are meaningful for the selected style.
function FieldVisible(Style, Field: Integer): Boolean;
begin
  Result := False;
  case Style of
    AuthIdxToken:    Result := (Field = 2);
    AuthIdxUserPass: Result := (Field = 0) or (Field = 1);
    AuthIdxNKey:     Result := (Field = 3);
    AuthIdxJwt:      Result := (Field = 3) or (Field = 4);
    AuthIdxCreds:    Result := (Field = 5);
    AuthIdxMtls:     Result := (Field = 6) or (Field = 7);
  end;
end;

procedure UpdateCredsFields();
var
  I, Style: Integer;
  Vis: Boolean;
begin
  if CredsPage = nil then Exit;
  Style := AuthPage.SelectedValueIndex;
  for I := 0 to 7 do
  begin
    Vis := FieldVisible(Style, I);
    CredsPage.Edits[I].Visible := Vis;
    CredsPage.PromptLabels[I].Visible := Vis;
  end;
end;

function AuthStyle(): Integer;
begin
  if WizardSilent() then Result := AuthStyleFromParam()
  else if AuthPage = nil then Result := AuthIdxNone
  else Result := AuthPage.SelectedValueIndex;
end;

function ShouldSkipPage(PageID: Integer): Boolean;
begin
  if WizardSilent() then
    Result := (PageID = ServerPage.ID) or (PageID = ConnPage.ID) or
              (PageID = AuthPage.ID) or (PageID = CredsPage.ID) or (PageID = OptionsPage.ID)
  else
    // The credentials page is only useful when a style was chosen.
    Result := (PageID = CredsPage.ID) and (AuthStyle() = AuthIdxNone);
end;

function FieldsFilled(): Boolean;
var s: Integer;
begin
  s := AuthStyle();
  Result := True;
  case s of
    AuthIdxUserPass: Result := (Trim(CredsPage.Values[0]) <> '');
    AuthIdxNKey:     Result := (Trim(CredsPage.Values[3]) <> '');
    AuthIdxJwt:      Result := (Trim(CredsPage.Values[3]) <> '') and (Trim(CredsPage.Values[4]) <> '');
    AuthIdxCreds:    Result := (Trim(CredsPage.Values[5]) <> '');
    AuthIdxMtls:     Result := (Trim(CredsPage.Values[6]) <> '');
  end;
end;

function NextButtonClick(CurPageID: Integer): Boolean;
begin
  Result := True;
  if (CurPageID = ConnPage.ID) and (Trim(ConnPage.Values[0]) = '') then
  begin
    MsgBox(ExpandConstant('{cm:UrlRequired}'), mbError, MB_OK);
    Result := False;
  end
  else if (CurPageID = CredsPage.ID) and (not FieldsFilled()) then
  begin
    MsgBox(ExpandConstant('{cm:AuthRequired}'), mbError, MB_OK);
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

// Credential value by field index (seeded from params when silent).
function Cred(Field: Integer): String;
begin
  if (not WizardSilent()) and (CredsPage <> nil) then
    Result := Trim(CredsPage.Values[Field])
  else
    case Field of
      0: Result := ParamOr('USER', '');
      1: Result := ParamOr('PASSWORD', '');
      2: Result := ParamOr('TOKEN', '');
      3: Result := ParamOr('NKEYSEED', '');
      4: Result := ParamOr('JWT', '');
      5: Result := ParamOr('CREDS', '');
      6: Result := ParamOr('TLSCERT', '');
      7: Result := ParamOr('TLSCERTPASSWORD', '');
    else
      Result := '';
    end;
end;

function AuthStyleName(): String;
begin
  case AuthStyle() of
    AuthIdxToken:    Result := ExpandConstant('{cm:AuthToken}');
    AuthIdxUserPass: Result := ExpandConstant('{cm:AuthUserPass}');
    AuthIdxNKey:     Result := ExpandConstant('{cm:AuthNKey}');
    AuthIdxJwt:      Result := ExpandConstant('{cm:AuthJwt}');
    AuthIdxCreds:    Result := ExpandConstant('{cm:AuthCreds}');
    AuthIdxMtls:     Result := ExpandConstant('{cm:AuthMtls}');
  else
    Result := ExpandConstant('{cm:AuthNone}');
  end;
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
  // Silent mode: honour /SERVICE=1 so an automated install can register the
  // service. It still needs elevation - the helper exits 2 and nothing is
  // registered when the process is not elevated, which is the documented
  // behavior (the add-in install never depends on it).
  if WizardSilent() then
    Result := (ParamOr('SERVICE', '0') = '1')
  else
    Result := (OptionsPage <> nil) and OptionsPage.Values[1];
end;

// ---- JSON helpers ---------------------------------------------------------
// Escape a value for a JSON string: backslash first (a Windows path is the
// common case), then the quote.
function JsEscape(const S: String): String;
begin
  Result := S;
  StringChangeEx(Result, '\', '\\', True);
  StringChangeEx(Result, '"', '\"', True);
end;

// The lines inside the "auth" object for the selected style. Empty for no auth.
function AuthJsonBody(): String;
var L: String;
begin
  L := '';
  case AuthStyle() of
    AuthIdxToken:
      L := '      "token": "' + JsEscape(Cred(2)) + '"';
    AuthIdxUserPass:
      begin
        L := '      "user": "' + JsEscape(Cred(0)) + '"';
        L := L + ',' + #13#10 + '      "password": "' + JsEscape(Cred(1)) + '"';
      end;
    AuthIdxNKey:
      L := '      "nkeySeed": "' + JsEscape(Cred(3)) + '"';
    AuthIdxJwt:
      begin
        L := '      "jwt": "' + JsEscape(Cred(4)) + '"';
        L := L + ',' + #13#10 + '      "nkeySeed": "' + JsEscape(Cred(3)) + '"';
      end;
    AuthIdxCreds:
      L := '      "credsFile": "' + JsEscape(Cred(5)) + '"';
    AuthIdxMtls:
      begin
        L := '      "tls": true,' + #13#10 + '      "tlsCert": "' + JsEscape(Cred(6)) + '"';
        if Cred(7) <> '' then
          L := L + ',' + #13#10 + '      "tlsCertPassword": "' + JsEscape(Cred(7)) + '"';
      end;
  end;
  Result := L;
end;

// The generated config: the same shape the add-in ships, with the chosen URL
// and credential. Comments are accepted, so the template stays self-documenting.
function ConfigJson(): String;
var Auth: String;
begin
  Auth := AuthJsonBody();
  Result :=
    '{' + #13#10 +
    '  // NATSExcel.json - comments (// and /* */) are allowed.' + #13#10 +
    '  "Default": "local",' + #13#10 +
    '  "Servers": {' + #13#10 +
    '    "srv1": "' + JsEscape(Url()) + '"' + #13#10 +
    '  },' + #13#10 +
    '  "Clusters": {' + #13#10 +
    '    "local": { "urls": ["srv1"] }' + #13#10 +
    '  },' + #13#10;
  if Auth = '' then
    Result := Result + '  "Auth": {},' + #13#10
  else
    Result := Result + '  "Auth": {' + #13#10 + Auth + #13#10 + '  },' + #13#10;
  Result := Result +
    '  "Reconnect": { "maxAttempts": 60, "waitMs": 2000, "jitterMs": 100 },' + #13#10 +
    '  "RTD": {' + #13#10 +
    '    "NatsUpdateRateMs": 1000,' + #13#10 +
    '    "ExcelUpdateRateMs": 100,' + #13#10 +
    '    "MessageCounterThreshold": 10000,' + #13#10 +
    '    "NatsUpdateStyle": "Automatic"' + #13#10 +
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
  // The connection test can only use a token; other styles are verified by the
  // add-in itself (the test script speaks the protocol, not the client libs).
  if AuthStyle() = AuthIdxToken then
    Args := Args + ' -Token "' + Cred(2) + '"';
  TestOk := (RunScript('test-connection.ps1', Args) = 0);

  if not TestOk then
    Msg := FmtMessage(ExpandConstant('{cm:TestFailed}'), [HostPort])
  else
    Msg := FmtMessage(ExpandConstant('{cm:TestOk}'), [HostPort, ExpandConstant('{app}')]);

  if not WizardSilent() then
    MsgBox(Msg, mbInformation, MB_OK);

  if OpenExcelChecked() then
  begin
    Args := '-AddInDir "' + ExpandConstant('{app}') + '" -ServerHost "' + HostPort + '" -OpenExcel -Quiet';
    if AuthStyle() = AuthIdxToken then
      Args := Args + ' -Token "' + Cred(2) + '"';
    RunScript('verify-connection.ps1', Args);
  end;
end;

procedure CurPageChanged(CurPageID: Integer);
var
  Lines: String;
begin
  if CurPageID = CredsPage.ID then
    UpdateCredsFields();

  if (CurPageID = SummaryPage.ID) and (not WizardSilent()) then
  begin
    Lines :=
      ExpandConstant('{cm:SummaryLabel}') + ':  ' + Url() + #13#10 +
      ExpandConstant('{cm:SummaryAuth}') + ':  ' + AuthStyleName() + #13#10 +
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

// The COM ProgId `NatsRtd` is per-USER, not per-install: if another NATSExcel
// copy (a parallel folder, a lab install) still has its XLL registered in
// Excel, deleting the key would break its =RTD cells. Only remove it when no
// other OPEN* entry points at a NATSExcel XLL.
function AnotherNatsExcelInstallRemains(): Boolean;
var
  Keys: TArrayOfString;
  I, N: Integer;
  BaseKey, ValName, Cur: String;
begin
  Result := False;
  Keys := OfficeExcelKeys();
  for I := 0 to GetArrayLength(Keys) - 1 do
  begin
    BaseKey := 'SOFTWARE\Microsoft\Office\' + Keys[I] + '\Excel\Options';
    if RegQueryStringValue(HKCU, BaseKey, 'OPEN', Cur) then
      if Pos('NATSExcel', Cur) > 0 then begin Result := True; Exit; end;
    N := 1;
    while N < 100 do
    begin
      ValName := 'OPEN' + IntToStr(N);
      if RegQueryStringValue(HKCU, BaseKey, ValName, Cur) then
        if Pos('NATSExcel', Cur) > 0 then begin Result := True; Exit; end;
      N := N + 1;
    end;
  end;
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
  if CurUninstallStep = usUninstall then
  begin
    UnregisterExcelAutoLoad();
    // Remove the ProgId only when this was the last NATSExcel install (the
    // auto-load unregister above already dropped THIS install's entries).
    if not AnotherNatsExcelInstallRemains() then
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
