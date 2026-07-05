# Reporte de Rendimiento V3.6

Herramienta PowerShell para generar diagnósticos integrales de equipos Windows: consumo de CPU/RAM, temperatura, disco, red, eventos y seguridad. Incluye benchmark de carga con estrés de CPU, pruebas secuenciales + 4K en disco, evaluación de RAM y conectividad básica.

## Stack

- PowerShell 5.1+
- Pester 5.x (tests)
- Bootstrap 5.3.3 + Chart.js (reporte HTML)

## Requisitos

- Windows 10/11 o Windows Server 2016+
- PowerShell ejecutado **como Administrador** para datos completos de sensores/SMART
- LibreHardwareMonitor (opcional, se auto-instala con `-InstalarLibreHardwareMonitor`)
- smartmontools (opcional, con `-UseSmartCtl`)

## Uso rápido

```powershell
.\Reporte-Rendimiento-V3.6-Temperatura-Eventos.ps1 -Modo Antes -Cliente ACME -OrdenTrabajo 12345 -Tecnico Gabriela
```

## Tests

```powershell
Import-Module Pester
Invoke-Pester ./tests
```

## Opciones relevantes

| Parámetro | Descripción |
|---|---|
| `-DuracionReposoSeg` / `-DuracionPruebaSeg` | Ajustan duración de cada fase |
| `-MaxCpuWorkers` / `-DiskTestFileMB` | Controlan carga de benchmark |
| `-UseLibreHardwareMonitor` / `-UseSmartCtl` | Habilitan sensores/S.M.A.R.T. |
| `-AnalizarFragmentacion` / `-IncluirChkdskScan` | Análisis avanzados de almacenamiento |
| `-InstalarLibreHardwareMonitor` | Descarga e instala la DLL automáticamente |

## Salida generada

- `reporte.html` — versión visual con mosaicos, resumen de fases y gráfica Chart.js
- `reporte.json` — todos los datos capturados
- `resumen.txt` — resumen textual
- CSVs de temperaturas, procesos, discos y eventos

## Licencia

MIT
