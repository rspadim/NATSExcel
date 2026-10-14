---
title: NATSExcel
description: Complemento de Excel (XLL) que lleva NATS a Excel mediante RTD y funciones de hoja de cálculo
---

<p align="center">
  <picture>
    <source media="(prefers-color-scheme: dark)" srcset="../assets/logo-dark.svg">
    <img src="../assets/logo.svg" alt="NATSExcel" width="420">
  </picture>
</p>

**Complemento de Excel (XLL) que lleva [NATS](https://nats.io/) a Excel** mediante
RTD (Real-Time Data) y UDF (funciones definidas por el usuario): escribes una
fórmula, el dato de NATS llega a la celda y —en las funciones de streaming— se
mantiene en vivo.

Cubre **NATS Core** (pub/sub, request/reply, grupos de cola, headers) y
**JetStream** (streams, consumidores, replay, Key/Value Store, Object Store).
Construido con Excel-DNA sobre .NET Framework 4.8, usando NATS.Net, NLog y
Newtonsoft.Json.

Los binarios para Windows están **firmados digitalmente**; la firma la ofrece
gratuitamente a proyectos de código abierto la
[SignPath Foundation](https://signpath.org/).

[![Licencia: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://github.com/rspadim/NATSExcel/blob/main/LICENSE)
[![Última versión](https://img.shields.io/github/v/release/rspadim/NATSExcel?sort=semver)](https://github.com/rspadim/NATSExcel/releases/latest)

## Requisitos

- **Windows** con Microsoft Office (Excel), 64 bits o 32 bits.
- **.NET Framework 4.8**, que ya viene en Windows 10 y Windows 11.
- Un **servidor NATS** al que conectarse. El instalador puede levantar uno de
  forma nativa en Windows; Docker es opcional y nunca obligatorio.

## Qué hace

| Área | Cobertura |
| --- | --- |
| NATS Core | publish, request/reply, último valor por asunto, unsubscribe, grupos de cola, headers |
| JetStream | publish con ack, get/fetch/replay, streams y consumidores (información + administración) |
| Key/Value | get/put/create/update/delete/keys/history/watch |
| Object Store | put/get/info/list/watch/delete |
| Estado | versión, ruta de configuración, número de conexiones, información del servidor, RTT, topología del clúster |
| Celdas en vivo | temas de `=RTD("NatsRtd", ...)`: `SUB`, `KV`, `KVGET`, `KVKEYS`, `REQ`, `JS`, `MSG`, `OBJ`, `RTDCLUSTER` |

## Descarga

Descarga la versión más reciente en la
**[página de releases](https://github.com/rspadim/NATSExcel/releases/latest)**:

| Archivo | Para qué sirve |
| --- | --- |
| `NATSExcelSetup-<versión>.exe` | Instalador de un clic (por usuario, nativo primero; asistente en 11 idiomas) |
| `NATSExcel64-packed.xll` | Complemento para Office de 64 bits |
| `NATSExcel-packed.xll` | Complemento para Office de 32 bits |
| `NATSExcel.dll` | Ensamblado gestionado del complemento (va junto al XLL) |
| `NLog.config` | Configuración de registro (log) |
| `NATSExcel.json` | Configuración inicial |

Los releases están firmados digitalmente; el servicio de firma lo proporciona la
[SignPath Foundation](https://signpath.org/), que emite certificados a proyectos de
código abierto.

## Instalación

1. Ejecuta `NATSExcelSetup-<versión>.exe`. Copia el complemento, detecta la
   arquitectura de tu Office, escribe un `NATSExcel.json` inicial y (si quieres)
   levanta un servidor NATS por ti.
2. Abre Excel. El complemento se registra solo, así que el XLL se carga
   automáticamente. Si no se carga, habilítalo a mano en **Archivo → Opciones →
   Complementos → Administrar: Complementos de Excel → Examinar…**.
3. Conecta: edita el `NATSExcel.json` (o usa el asistente del instalador). El
   valor por defecto es `nats://localhost:4222`.

¿Aún no tienes servidor NATS? El instalador puede ejecutar uno de forma nativa en
Windows (un único `nats-server.exe`, con JetStream incluido). Docker es opcional,
nunca obligatorio.

## Verlo funcionar

```excel
=NatsUDFVersion()                                        ' el complemento responde
=NatsUDFPublish("orders.new","hello")                    ' publish en Core
=NatsUDFSubjectLatest("orders.new")                      ' lee el último valor
=NatsUDFKVPut("config","rate","1.23")                    ' escritura en Key/Value
=RTD("NatsRtd",,"SUB","orders.>")                        ' una celda en vivo
```

## Documentación

- **Catálogo de funciones (congelado):** [PUBLIC-SURFACE.md](../PUBLIC-SURFACE.md)
- **Arquitectura y diseño:** [DESIGN.md](https://github.com/rspadim/NATSExcel/blob/main/DESIGN.md)
- **Guía de primeros pasos:** [Readme.md](https://github.com/rspadim/NATSExcel/blob/main/Readme.md)
- **Plan de desarrollo:** [PLAN.md](https://github.com/rspadim/NATSExcel/blob/main/PLAN.md)
- **Pruebas:** [test/README.md](https://github.com/rspadim/NATSExcel/blob/main/test/README.md)

## Privacidad

NATSExcel no recopila ningún dato del usuario. La única petición que hace por su
cuenta es una llamada en segundo plano a la API de releases de GitHub, para
comprobar si existe una versión más nueva; el resultado se guarda en caché y los
fallos (sin conexión, límite de peticiones) se ignoran. Escribe un archivo de log
local (`NATSExcel.log`) junto a sus propios archivos. El tráfico NATS va solo al
servidor o servidores que configures. La comprobación de actualización puede
desactivarse con `UpdateCheck: false` en `NATSExcel.json`.

## Soporte y contribuciones

Las dudas, los informes de errores y los pull requests son bienvenidos en
**[github.com/rspadim/NATSExcel](https://github.com/rspadim/NATSExcel/issues)**.
Al abrir un error, incluye la versión del complemento (`=NatsUDFVersion()`), la
arquitectura de tu Office y si el servidor tenía JetStream activado.

## Traducciones

Este sitio está en portugués (Brasil), inglés y español. **Las traducciones son
muy bienvenidas** — es una de las formas más fáciles de ayudar al proyecto, aunque
no sepas C#:

1. Abre un _issue_ en
   **[github.com/rspadim/NATSExcel/issues](https://github.com/rspadim/NATSExcel/issues)**
   diciendo qué idioma quieres añadir, para que nadie haga el mismo trabajo dos
   veces.
2. Abre un _pull request_ con una carpeta `docs/<idioma>/` nueva (por ejemplo
   `docs/de/`), copiando `docs/en/index.md` y traduciendo el texto.
3. Añade el `locale` correspondiente en la lista `plugins: i18n: languages:` del
   `mkdocs.yml` (dos o tres líneas).

El **catálogo de funciones** y el resto de la documentación técnica **no** hace
falta traducirlos: son la referencia en inglés del complemento.

Otras formas de ayudar: reportar errores con la versión y la arquitectura de tu
Office, mejorar la documentación, probar en distintas versiones de Excel y
proponer ejemplos de uso.

## Licencia

NATSExcel se distribuye bajo la **Licencia MIT** — consulta el
[LICENSE](https://github.com/rspadim/NATSExcel/blob/main/LICENSE).
El cliente NATS.Net incluido es Apache-2.0.
