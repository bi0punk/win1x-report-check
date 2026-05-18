# Reporte de Rendimiento V3.6

Herramienta PowerShell para generar diagnósticos integrales de equipos Windows: consumo de CPU/RAM, temperatura, disco, red, eventos y seguridad. Incluye benchmark de carga con carga de CPU, pruebas secuenciales + 4K en disco, evaluación de RAM y conectividad básica.

## Uso rápido

1. Ejecuta `Reporte-Rendimiento-V3.6-Temperatura-Eventos.ps1` desde PowerShell con parámetros mínimos:
   ```powershell
   .\Reporte-Rendimiento-V3.6-Temperatura-Eventos.ps1 -Modo Antes -Cliente ACME -OrdenTrabajo 12345 -Tecnico Gabriela
   ```
2. Paramétralo según necesites: duración de reposo/prueba, intervalos, herramientas externas (`-UseLibreHardwareMonitor`, `-UseSmartCtl`, etc.).
3. Al finalizar crea carpeta con HTML interactivo, JSON, CSVs, TXT, recomendaciones y resumen de hallazgos.

## Salida generada

- `reporte.html`: versión visual con mosaicos, resumen de fases, tabla de red y gráfica Chart.js.
- `reporte.json`: todos los datos capturados.
- `resumen.txt`: resumen textual.
- CSVs adicionales para temperaturas, procesos, discos y eventos.

## Opciones relevantes

- `-DuracionReposoSeg` / `-DuracionPruebaSeg`: ajustan cuánto tiempo se monitorea cada fase.
- `-MaxCpuWorkers` / `-DiskTestFileMB`: controlan cuánta carga crea cada benchmark.
- `-UseLibreHardwareMonitor` / `-UseSmartCtl`: habilitan lecturas de sensores/S.M.A.R.T.`
- `-AnalizarFragmentacion`, `-IncluirChkdskScan`: ejecutan análisis avanzados de almacenamiento.

## Buenas prácticas

- Corre en PowerShell elevado para obtener datos de sensores/SMART más completos.
- Ejecuta `-InstalarLibreHardwareMonitor` si la DLL no está disponible.
- Revisa `reporte.html` con cualquier navegador moderno para aprovechar las gráficas.





.\Reporte-Rendimiento-V3.6-Temperatura-Eventos.ps1 -Modo Antes -Cliente "ClienteX" -OrdenTrabajo "OT-001" -Tecnico "Juan" -DuracionReposoSeg 30 -DuracionPruebaSeg 60