---
title: NATSExcel
description: Add-in do Excel (XLL) que traz o NATS para dentro do Excel via RTD e funções de planilha
---

**Add-in do Excel (XLL) que traz o [NATS](https://nats.io/) para dentro do Excel**
por meio de RTD (Real-Time Data) e UDF (funções definidas pelo usuário): você digita
uma fórmula, o dado do NATS chega na célula e — nas funções de streaming — fica
ao vivo.

Cobre o **NATS Core** (pub/sub, request/reply, grupos de fila, headers) e o
**JetStream** (streams, consumidores, replay, Key/Value Store, Object Store).
Construído com Excel-DNA sobre o .NET Framework 4.8, usando NATS.Net, NLog e
Newtonsoft.Json.

Os binários para Windows são **assinados digitalmente**; a assinatura é oferecida
gratuitamente a projetos de código aberto pela
[SignPath Foundation](https://signpath.org/).

[![Licença: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://github.com/rspadim/NATSExcel/blob/main/LICENSE)
[![Última versão](https://img.shields.io/github/v/release/rspadim/NATSExcel?sort=semver)](https://github.com/rspadim/NATSExcel/releases/latest)


Requisitos
----------

- **Windows** com Microsoft Office (Excel), 64 bits ou 32 bits.
- **.NET Framework 4.8**, que já vem no Windows 10 e no Windows 11.
- Um **servidor NATS** para conectar. O instalador consegue subir um nativamente
  no Windows; Docker é opcional e nunca obrigatório.


O que ele faz
-------------

| Área | Cobertura |
| --- | --- |
| NATS Core | publish, request/reply, último valor por assunto, unsubscribe, grupos de fila, headers |
| JetStream | publish com ack, get/fetch/replay, streams e consumidores (informações + administração) |
| Key/Value | get/put/create/update/delete/keys/history/watch |
| Object Store | put/get/info/list/watch/delete |
| Status | versão, caminho do config, número de conexões, informações do servidor, RTT, topologia do cluster |
| Células ao vivo | tópicos de `=RTD("NatsRtd", ...)`: `SUB`, `KV`, `KVGET`, `KVKEYS`, `REQ`, `JS`, `MSG`, `OBJ`, `RTDCLUSTER` |


Download
--------

Baixe a versão mais recente na **[página de releases](https://github.com/rspadim/NATSExcel/releases/latest)**:

| Arquivo | Para que serve |
| --- | --- |
| `NATSExcelSetup-<versão>.exe` | Instalador de um clique (por usuário, nativo primeiro; assistente em 11 idiomas) |
| `NATSExcel64-packed.xll` | Add-in para Office 64 bits |
| `NATSExcel-packed.xll` | Add-in para Office 32 bits |
| `NATSExcel.dll` | Assembly gerenciado do add-in (vai ao lado do XLL) |
| `NLog.config` | Configuração de log |
| `NATSExcel.json` | Configuração inicial |

Os releases são assinados digitalmente; o serviço de assinatura é fornecido pela
[SignPath Foundation](https://signpath.org/), que emite certificados para projetos
de código aberto.


Instalação
----------

1. Execute `NATSExcelSetup-<versão>.exe`. Ele copia o add-in, detecta a
   arquitetura do seu Office, escreve um `NATSExcel.json` inicial e (se você
   quiser) sobe um servidor NATS para você.
2. Abra o Excel. O add-in se registra sozinho, então o XLL carrega automaticamente.
   Se não carregar, habilite à mão em **Arquivo → Opções → Suplementos →
   Gerenciar: Suplementos do Excel → Procurar…**.
3. Conecte: edite o `NATSExcel.json` (ou use o assistente do instalador). O padrão
   é `nats://localhost:4222`.

Ainda não tem servidor NATS? O instalador consegue rodar um nativamente no Windows
(um único `nats-server.exe`, com JetStream incluído). Docker é opcional, nunca
obrigatório.


Veja funcionando
----------------

```excel
=NatsUDFVersion()                                        ' o add-in responde
=NatsUDFPublish("orders.new","hello")                    ' publish no Core
=NatsUDFSubjectLatest("orders.new")                      ' lê o último valor
=NatsUDFKVPut("config","rate","1.23")                    ' escrita no Key/Value
=RTD("NatsRtd",,"SUB","orders.>")                        ' uma célula ao vivo
```


Documentação
------------

- **Catálogo de funções (congelado):** [docs/PUBLIC-SURFACE.md](https://github.com/rspadim/NATSExcel/blob/main/docs/PUBLIC-SURFACE.md)
- **Arquitetura e design:** [DESIGN.md](https://github.com/rspadim/NATSExcel/blob/main/DESIGN.md)
- **Guia de primeiros passos:** [Readme.md](https://github.com/rspadim/NATSExcel/blob/main/Readme.md)
- **Plano de desenvolvimento:** [PLAN.md](https://github.com/rspadim/NATSExcel/blob/main/PLAN.md)
- **Testes:** [test/README.md](https://github.com/rspadim/NATSExcel/blob/main/test/README.md)


Privacidade
-----------

O NATSExcel não coleta nenhum dado do usuário. A única requisição que ele faz por
conta própria é uma chamada em segundo plano à API de releases do GitHub, para
verificar se existe uma versão mais nova; o resultado fica em cache e falhas
(offline, limite de requisições) são ignoradas. Ele grava um arquivo de log local
(`NATSExcel.log`) ao lado dos próprios arquivos. O tráfego NATS vai apenas para o(s)
servidor(es) que você configurar. A checagem de atualização pode ser desligada com
`UpdateCheck: false` no `NATSExcel.json`.


Suporte e contribuição
----------------------

Dúvidas, relatos de bug e pull requests são bem-vindos em
**[github.com/rspadim/NATSExcel](https://github.com/rspadim/NATSExcel/issues)**.
Ao abrir um bug, inclua a versão do add-in (`=NatsUDFVersion()`), a arquitetura do
seu Office e se o servidor tinha JetStream ativado.


Licença
-------

O NATSExcel é distribuído sob a **Licença MIT** — veja o
[LICENSE](https://github.com/rspadim/NATSExcel/blob/main/LICENSE).
O cliente NATS.Net embarcado é Apache-2.0.

> O NATSExcel é um projeto independente e não é afiliado nem endossado pelo projeto
> NATS ou pela Synadia.
