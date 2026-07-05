# Reporte de Rendimiento V3.6 — win1x-report-check

Herramienta PowerShell para generar diagnósticos integrales de equipos Windows: consumo de CPU/RAM, temperatura, disco, red, eventos y seguridad. Incluye benchmark de carga con estrés de CPU, pruebas secuenciales + 4K en disco, evaluación de RAM y conectividad básica.

[![License](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![CI](https://github.com/tu-usuario/win1x-report-check/actions/workflows/ci.yml/badge.svg)](https://github.com/tu-usuario/win1x-report-check/actions/workflows/ci.yml)

## Tabla de Contenidos

- [Características](#características)
- [Stack](#stack)
- [Arquitectura](#arquitectura)
- [Requisitos](#requisitos)
- [Instalación](#instalación)
- [Uso](#uso)
- [Tests](#tests)
- [Configuración](#configuración)
- [CI](#ci)
- [Salida](#salida)
- [Limitaciones / Roadmap](#limitaciones--roadmap)
- [Licencia](#licencia)

## Características

- Benchmark de CPU, RAM y disco (secuencial + 4K)
- Monitoreo de temperatura con LibreHardwareMonitor
- Lectura S.M.A.R.T. con smartmontools (opcional)
- Análisis de fragmentación y chkdsk
- Reporte HTML con Bootstrap 5.3.3 + Chart.js
- Exportación a JSON y TXT
- Parámetros configurables para duración y carga

## Stack

- PowerShell 5.1+, Pester 5.x (tests)
- Bootstrap 5.3.3, Chart.js (reporte HTML)

## Arquitectura

```
win1x-report-check/
├── Reporte-Rendimiento-V3.6-Temperatura-Eventos.ps1  # Script principal
├── template.html             # Plantilla HTML para reporte
├── modules/                  # Módulos PowerShell
├── tests/                    # Tests Pester
├── .github/workflows/ci.yml
└── README.md
```

## Requisitos

- Windows 10/11 o Windows Server 2016+
- PowerShell ejecutado **como Administrador**
- LibreHardwareMonitor (opcional, se auto-instala con `-InstalarLibreHardwareMonitor`)
- smartmontools (opcional, con `-UseSmartCtl`)

## Instalación

```powershell
git clone https://github.com/tu-usuario/win1x-report-check.git
cd win1x-report-check
```

## Uso

```powershell
# Ejecución básica
.\Reporte-Rendimiento-V3.6-Temperatura-Eventos.ps1 -Modo Antes -Cliente ACME -OrdenTrabajo 12345 -Tecnico "Gabriela"

# Con sensores de temperatura
.\Reporte-Rendimiento-V3.6-Temperatura-Eventos.ps1 -Modo Antes -UseLibreHardwareMonitor -InstalarLibreHardwareMonitor

# Con S.M.A.R.T.
.\Reporte-Rendimiento-V3.6-Temperatura-Eventos.ps1 -Modo Antes -UseSmartCtl
```

## Tests

```powershell
Import-Module Pester
Invoke-Pester ./tests
```

## Configuración

| Parámetro                    | Descripción                                      |
|------------------------------|--------------------------------------------------|
| `-DuracionReposoSeg`         | Duración de la fase de reposo (default: 60)      |
| `-DuracionPruebaSeg`         | Duración de la prueba de estrés (default: 60)    |
| `-MaxCpuWorkers`             | Workers de CPU (default: 4)                      |
| `-DiskTestFileMB`            | Tamaño del archivo de prueba de disco (default: 1000) |
| `-UseLibreHardwareMonitor`   | Habilitar sensores de temperatura                |
| `-UseSmartCtl`               | Habilitar lectura S.M.A.R.T.                     |
| `-InstalarLibreHardwareMonitor` | Descarga e instala la DLL automáticamente    |

## CI

GitHub Actions ejecuta tests Pester en Windows latest en cada push y PR.

## Salida

- `reporte.html` — versión visual con mosaicos, resumen de fases y gráfica Chart.js
- `reporte.json` — todos los datos capturados en formato estructurado
- `resumen.txt` — resumen textual para revisión rápida
- CSVs de temperaturas, procesos, discos y eventos

## Limitaciones / Roadmap

- [ ] Integración con WMI/CIM para métricas adicionales
- [ ] Programación de diagnósticos periódicos
- [ ] Subida automática de reportes a servidor
- [ ] Tests de integración con Pester + PSScriptAnalyzer

## Licencia

MIT
