param(
    [Parameter(Mandatory=$true)]
    [ValidateSet("Antes","Despues")]
    [string]$Modo,

    [string]$Cliente = "SinCliente",
    [string]$OrdenTrabajo = "SinOT",
    [string]$Tecnico = "Tecnico",

    [int]$DuracionReposoSeg = 60,
    [int]$DuracionPruebaSeg = 120,
    [int]$IntervaloMuestreoSeg = 5,

    [string]$RutaSalida = "",

    [switch]$AbrirReporte,

    [switch]$UseLibreHardwareMonitor,
    [switch]$InstalarLibreHardwareMonitor,
    [string]$LibreHardwareMonitorPath = "",

    [switch]$UseSmartCtl,
    [string]$SmartCtlPath = "smartctl.exe",

    [int]$MaxCpuWorkers = 0,
    [int]$DiskTestFileMB = 256,

    [switch]$AnalizarFragmentacion,
    [switch]$IncluirChkdskScan
)

$ErrorActionPreference = "Continue"

function Write-Info {
    param([string]$Message)
    Write-Host "[INFO] $Message" -ForegroundColor Cyan
}

function Write-Warn2 {
    param([string]$Message)
    Write-Host "[WARN] $Message" -ForegroundColor Yellow
}

function Safe-Name {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return "SinDato" }
    $safe = $Text -replace '[\\/:*?"<>|]', '_'
    $safe = $safe -replace '\s+', '_'
    return $safe.Trim('_')
}

function Html-Encode {
    param([object]$Value)
    if ($null -eq $Value) { return "" }
    return [System.Net.WebUtility]::HtmlEncode([string]$Value)
}

function Is-Admin {
    try {
        $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object Security.Principal.WindowsPrincipal($identity)
        return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    } catch {
        return $false
    }
}


function Get-DefaultWritableBase {
    $candidates = @()

    if ($env:ProgramData) {
        $candidates += (Join-Path $env:ProgramData "ReporteRendimiento")
    }

    if ($env:USERPROFILE) {
        $candidates += (Join-Path $env:USERPROFILE "Documents\ReporteRendimiento")
        $candidates += (Join-Path $env:USERPROFILE "Desktop\ReporteRendimiento")
    }

    if ($env:TEMP) {
        $candidates += (Join-Path $env:TEMP "ReporteRendimiento")
    }

    foreach ($base in $candidates) {
        try {
            if (!(Test-Path $base)) {
                New-Item -ItemType Directory -Path $base -Force | Out-Null
            }

            $testFile = Join-Path $base ("write_test_" + [guid]::NewGuid().ToString() + ".tmp")
            "ok" | Out-File -FilePath $testFile -Encoding ASCII -Force
            Remove-Item $testFile -Force -ErrorAction SilentlyContinue

            return $base
        } catch {}
    }

    return (Join-Path $env:TEMP "ReporteRendimiento")
}

function Initialize-DefaultPaths {
    param(
        [string]$InputRutaSalida,
        [string]$InputLibreHardwareMonitorPath
    )

    $base = Get-DefaultWritableBase

    $resolvedRuta = $InputRutaSalida
    if ([string]::IsNullOrWhiteSpace($resolvedRuta)) {
        $resolvedRuta = Join-Path $base "Reportes-Rendimiento"
    }

    $toolsDir = Join-Path $base "Tools"
    $lhmDir = Join-Path $toolsDir "LibreHardwareMonitor"

    $resolvedLhm = $InputLibreHardwareMonitorPath
    if ([string]::IsNullOrWhiteSpace($resolvedLhm)) {
        $resolvedLhm = Join-Path $lhmDir "LibreHardwareMonitorLib.dll"
    }

    try {
        if (!(Test-Path $resolvedRuta)) {
            New-Item -ItemType Directory -Path $resolvedRuta -Force | Out-Null
        }
    } catch {}

    try {
        if (!(Test-Path $toolsDir)) {
            New-Item -ItemType Directory -Path $toolsDir -Force | Out-Null
        }
    } catch {}

    return New-Object PSObject -Property @{
        BaseDir = $base
        RutaSalida = $resolvedRuta
        ToolsDir = $toolsDir
        LibreHardwareMonitorDir = $lhmDir
        LibreHardwareMonitorPath = $resolvedLhm
    }
}

function New-RunFolder {
    param([string]$Root,[string]$Cliente,[string]$OrdenTrabajo,[string]$Modo)
    $clienteSafe = Safe-Name $Cliente
    $otSafe = Safe-Name $OrdenTrabajo
    $hostSafe = Safe-Name $env:COMPUTERNAME
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $base = Join-Path $Root ($clienteSafe + "_" + $otSafe + "_" + $hostSafe)
    $run = Join-Path $base ($timestamp + "_" + $Modo.ToUpper())
    New-Item -ItemType Directory -Path $run -Force | Out-Null
    return $run
}

function Get-SystemInfo {
    Write-Info "Recolectando datos del sistema"
    $cs = $null
    $os = $null
    $bios = $null
    $cpu = $null
    $gpuList = @()
    $ramList = @()

    try { $cs = Get-CimInstance Win32_ComputerSystem } catch {}
    try { $os = Get-CimInstance Win32_OperatingSystem } catch {}
    try { $bios = Get-CimInstance Win32_BIOS } catch {}
    try { $cpu = Get-CimInstance Win32_Processor | Select-Object -First 1 } catch {}
    try { $gpuList = @(Get-CimInstance Win32_VideoController) } catch {}
    try { $ramList = @(Get-CimInstance Win32_PhysicalMemory) } catch {}

    $ramTotalGB = $null
    if ($cs -ne $null) {
        if ($cs.TotalPhysicalMemory -ne $null) {
            $ramTotalGB = [math]::Round($cs.TotalPhysicalMemory / 1GB, 2)
        }
    }

    $ramModules = @()
    foreach ($m in $ramList) {
        $part = ""
        if ($m.PartNumber -ne $null) { $part = ([string]$m.PartNumber).Trim() }
        $capGB = $null
        if ($m.Capacity -ne $null) { $capGB = [math]::Round($m.Capacity / 1GB, 2) }

        $ramModules += New-Object PSObject -Property @{
            Banco = $m.BankLabel
            Marca = $m.Manufacturer
            Parte = $part
            CapacidadGB = $capGB
            SpeedMHz = $m.Speed
        }
    }

    $gpus = @()
    foreach ($g in $gpuList) {
        $gRam = $null
        if ($g.AdapterRAM -ne $null) { $gRam = [math]::Round($g.AdapterRAM / 1GB, 2) }
        $gpus += New-Object PSObject -Property @{
            Nombre = $g.Name
            Driver = $g.DriverVersion
            RAMGB = $gRam
        }
    }

    $fabricante = $null
    $modelo = $null
    $serial = $null
    $windows = $null
    $version = $null
    $build = $null
    $arquitectura = $null
    $ultimoArranque = $null
    $cpuName = $null
    $cpuCores = $null
    $cpuThreads = $null
    $cpuMaxMHz = $null

    if ($cs -ne $null) {
        $fabricante = $cs.Manufacturer
        $modelo = $cs.Model
    }

    if ($bios -ne $null) { $serial = $bios.SerialNumber }

    if ($os -ne $null) {
        $windows = $os.Caption
        $version = $os.Version
        $build = $os.BuildNumber
        $arquitectura = $os.OSArchitecture
        $ultimoArranque = $os.LastBootUpTime
    }

    if ($cpu -ne $null) {
        $cpuName = $cpu.Name
        $cpuCores = $cpu.NumberOfCores
        $cpuThreads = $cpu.NumberOfLogicalProcessors
        $cpuMaxMHz = $cpu.MaxClockSpeed
    }

    return New-Object PSObject -Property @{
        Hostname = $env:COMPUTERNAME
        Usuario = "$env:USERDOMAIN\$env:USERNAME"
        EsAdministrador = Is-Admin
        Fabricante = $fabricante
        Modelo = $modelo
        Serial = $serial
        Windows = $windows
        Version = $version
        Build = $build
        Arquitectura = $arquitectura
        UltimoArranque = $ultimoArranque
        CPU = $cpuName
        CPUCores = $cpuCores
        CPUThreads = $cpuThreads
        CPUMaxMHz = $cpuMaxMHz
        RAMTotalGB = $ramTotalGB
        RAMModulos = $ramModules
        GPU = $gpus
    }
}

function Add-LhmTemperatureSensors {
    param($Hardware, [ref]$Results)

    try { $Hardware.Update() } catch {}

    foreach ($sub in $Hardware.SubHardware) {
        try { $sub.Update() } catch {}
    }

    foreach ($sensor in $Hardware.Sensors) {
        try {
            if (($sensor.SensorType.ToString() -eq "Temperature") -and ($sensor.Value -ne $null)) {
                $Results.Value += New-Object PSObject -Property @{
                    Sensor = "$($Hardware.Name) - $($sensor.Name)"
                    Tipo = $Hardware.HardwareType.ToString()
                    TemperaturaC = [math]::Round([double]$sensor.Value, 1)
                    Fuente = "LibreHardwareMonitor_DLL"
                }
            }
        } catch {}
    }

    foreach ($sub in $Hardware.SubHardware) {
        foreach ($sensor in $sub.Sensors) {
            try {
                if (($sensor.SensorType.ToString() -eq "Temperature") -and ($sensor.Value -ne $null)) {
                    $Results.Value += New-Object PSObject -Property @{
                        Sensor = "$($sub.Name) - $($sensor.Name)"
                        Tipo = $sub.HardwareType.ToString()
                        TemperaturaC = [math]::Round([double]$sensor.Value, 1)
                        Fuente = "LibreHardwareMonitor_DLL"
                    }
                }
            } catch {}
        }
    }
}

function Get-LhmLoaderExceptionText {
    param($Exception)

    $msgs = @()
    try {
        if ($Exception.LoaderExceptions -ne $null) {
            foreach ($le in $Exception.LoaderExceptions) {
                if ($le -ne $null) { $msgs += $le.Message }
            }
        }
    } catch {}

    try {
        if ($Exception.InnerException -ne $null) {
            if ($Exception.InnerException.LoaderExceptions -ne $null) {
                foreach ($le in $Exception.InnerException.LoaderExceptions) {
                    if ($le -ne $null) { $msgs += $le.Message }
                }
            }
        }
    } catch {}

    if ($msgs.Count -eq 0) { return $Exception.Message }
    return ($msgs -join " | ")
}

function Import-LhmDependencies {
    param([string]$DllPath)

    try {
        $dir = Split-Path -Parent $DllPath
        if ([string]::IsNullOrWhiteSpace($dir)) { return }
        if (!(Test-Path $dir)) { return }

        try { Get-ChildItem -Path $dir -Recurse -ErrorAction SilentlyContinue | Unblock-File -ErrorAction SilentlyContinue } catch {}

        $dlls = @(Get-ChildItem -Path $dir -Filter "*.dll" -ErrorAction SilentlyContinue)

        # Cargar primero dependencias comunes y luego LibreHardwareMonitorLib.
        foreach ($d in $dlls) {
            if ($d.Name -ne "LibreHardwareMonitorLib.dll") {
                try { [System.Reflection.Assembly]::LoadFrom($d.FullName) | Out-Null } catch {}
            }
        }

        foreach ($d in $dlls) {
            if ($d.Name -eq "LibreHardwareMonitorLib.dll") {
                try { [System.Reflection.Assembly]::LoadFrom($d.FullName) | Out-Null } catch {}
            }
        }
    } catch {}
}

function Get-TemperatureFromLhm {
    param([string]$DllPath)
    $results = @()

    if (!(Test-Path $DllPath)) {
        $results += New-Object PSObject -Property @{
            Sensor = "LibreHardwareMonitor"
            Tipo = "Info"
            TemperaturaC = $null
            Fuente = "DLL no encontrada: $DllPath"
        }
        return $results
    }

    $computer = $null
    try {
        Import-LhmDependencies -DllPath $DllPath
        Add-Type -Path $DllPath -ErrorAction Stop

        $computer = New-Object LibreHardwareMonitor.Hardware.Computer
        $computer.IsCpuEnabled = $true
        $computer.IsGpuEnabled = $true
        $computer.IsStorageEnabled = $true
        $computer.IsMotherboardEnabled = $true
        $computer.IsMemoryEnabled = $true
        $computer.IsControllerEnabled = $true
        $computer.IsNetworkEnabled = $false
        $computer.Open()

        foreach ($hw in $computer.Hardware) {
            Add-LhmTemperatureSensors -Hardware $hw -Results ([ref]$results)
        }

        if ($computer -ne $null) { $computer.Close() }
    } catch {
        try { if ($computer -ne $null) { $computer.Close() } } catch {}
        $loaderText = Get-LhmLoaderExceptionText -Exception $_.Exception
        $results += New-Object PSObject -Property @{
            Sensor = "LibreHardwareMonitor"
            Tipo = "Error"
            TemperaturaC = $null
            Fuente = "Error cargando DLL/dependencias: $loaderText"
        }
    }

    return $results
}


function Resolve-LhmDllPath {
    param([string]$RequestedPath)

    $candidates = @()

    if (![string]::IsNullOrWhiteSpace($RequestedPath)) {
        $candidates += $RequestedPath
    }

    try {
        $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
        if (![string]::IsNullOrWhiteSpace($scriptDir)) {
            $candidates += (Join-Path $scriptDir "LibreHardwareMonitorLib.dll")
            $candidates += (Join-Path $scriptDir "LibreHardwareMonitor\LibreHardwareMonitorLib.dll")
            $candidates += (Join-Path $scriptDir "Tools\LibreHardwareMonitor\LibreHardwareMonitorLib.dll")
        }
    } catch {}

    try {
        $defaultBase = Get-DefaultWritableBase
        $candidates += (Join-Path $defaultBase "Tools\LibreHardwareMonitor\LibreHardwareMonitorLib.dll")
    } catch {}

    if ($env:ProgramData) {
        $candidates += (Join-Path $env:ProgramData "ReporteRendimiento\Tools\LibreHardwareMonitor\LibreHardwareMonitorLib.dll")
    }

    if ($env:USERPROFILE) {
        $candidates += (Join-Path $env:USERPROFILE "Documents\ReporteRendimiento\Tools\LibreHardwareMonitor\LibreHardwareMonitorLib.dll")
        $candidates += (Join-Path $env:USERPROFILE "Desktop\ReporteRendimiento\Tools\LibreHardwareMonitor\LibreHardwareMonitorLib.dll")
    }

    $candidates += "C:\Program Files\LibreHardwareMonitor\LibreHardwareMonitorLib.dll"
    $candidates += "C:\Program Files (x86)\LibreHardwareMonitor\LibreHardwareMonitorLib.dll"
    $candidates += "C:\ReporteRendimiento\Tools\LibreHardwareMonitor\LibreHardwareMonitorLib.dll"

    foreach ($c in $candidates) {
        if (![string]::IsNullOrWhiteSpace($c)) {
            if (Test-Path $c) {
                return $c
            }
        }
    }

    return $RequestedPath
}


function Install-LibreHardwareMonitor {
    param(
        [string]$TargetDir = ""
    )

    Write-Info "LibreHardwareMonitor no encontrado. Intentando instalar desde GitHub Releases."

    if ([string]::IsNullOrWhiteSpace($TargetDir)) {
        $defaultBase = Get-DefaultWritableBase
        $TargetDir = Join-Path $defaultBase "Tools\LibreHardwareMonitor"
    }

    $result = New-Object PSObject -Property @{
        Installed = $false
        DllPath = $null
        TargetDir = $TargetDir
        Error = $null
        SourceUrl = $null
    }

    try {
        if (!(Test-Path $TargetDir)) {
            New-Item -ItemType Directory -Path $TargetDir -Force | Out-Null
        }

        try {
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        } catch {}

        $tmpRoot = Join-Path $env:TEMP ("LibreHardwareMonitor_" + [guid]::NewGuid().ToString())
        New-Item -ItemType Directory -Path $tmpRoot -Force | Out-Null

        $zipPath = Join-Path $tmpRoot "LibreHardwareMonitor.zip"
        $assetUrl = $null

        try {
            $api = "https://api.github.com/repos/LibreHardwareMonitor/LibreHardwareMonitor/releases/latest"
            $release = Invoke-RestMethod -Uri $api -UseBasicParsing -ErrorAction Stop
            foreach ($asset in $release.assets) {
                $name = [string]$asset.name
                if (($name -match "LibreHardwareMonitor") -and ($name -match "\.zip$")) {
                    $assetUrl = [string]$asset.browser_download_url
                    break
                }
            }
        } catch {
            $result.Error = "No se pudo consultar GitHub API: $($_.Exception.Message)"
        }

        if ([string]::IsNullOrWhiteSpace($assetUrl)) {
            # Fallback conocido. Si cambia el nombre en releases futuras, el GitHub API anterior suele resolverlo.
            $assetUrl = "https://github.com/LibreHardwareMonitor/LibreHardwareMonitor/releases/latest/download/LibreHardwareMonitor-net472.zip"
        }

        $result.SourceUrl = $assetUrl
        Write-Info "Descargando LibreHardwareMonitor: $assetUrl"
        Invoke-WebRequest -Uri $assetUrl -OutFile $zipPath -UseBasicParsing -ErrorAction Stop

        $extractDir = Join-Path $tmpRoot "extract"
        New-Item -ItemType Directory -Path $extractDir -Force | Out-Null
        Expand-Archive -Path $zipPath -DestinationPath $extractDir -Force

        $dll = Get-ChildItem -Path $extractDir -Filter "LibreHardwareMonitorLib.dll" -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($dll -eq $null) {
            throw "No se encontro LibreHardwareMonitorLib.dll dentro del ZIP descargado."
        }

        $sourceDir = $dll.Directory.FullName
        Copy-Item -Path (Join-Path $sourceDir "*") -Destination $TargetDir -Recurse -Force

        try {
            Get-ChildItem -Path $TargetDir -Recurse -ErrorAction SilentlyContinue | Unblock-File -ErrorAction SilentlyContinue
        } catch {}

        $targetDll = Join-Path $TargetDir "LibreHardwareMonitorLib.dll"
        if (!(Test-Path $targetDll)) {
            $targetDll = (Get-ChildItem -Path $TargetDir -Filter "LibreHardwareMonitorLib.dll" -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1).FullName
        }

        if ([string]::IsNullOrWhiteSpace($targetDll)) {
            throw "La instalacion terminó, pero no se encontro la DLL final."
        }

        $result.Installed = $true
        $result.DllPath = $targetDll
        Write-Info "LibreHardwareMonitor instalado en: $TargetDir"
        Write-Info "DLL detectada: $targetDll"

        try { Remove-Item $tmpRoot -Recurse -Force -ErrorAction SilentlyContinue } catch {}
        return $result
    } catch {
        $result.Error = $_.Exception.Message
        Write-Warn2 "No se pudo instalar LibreHardwareMonitor: $($result.Error)"
        return $result
    }
}

function Ensure-LibreHardwareMonitor {
    param([string]$RequestedDllPath)

    $resolved = Resolve-LhmDllPath -RequestedPath $RequestedDllPath
    if (![string]::IsNullOrWhiteSpace($resolved)) {
        if (Test-Path $resolved) {
            return $resolved
        }
    }

    $defaultBase = Get-DefaultWritableBase
    $targetDir = Join-Path $defaultBase "Tools\LibreHardwareMonitor"
    if (![string]::IsNullOrWhiteSpace($RequestedDllPath)) {
        try {
            $parent = Split-Path -Parent $RequestedDllPath
            if (![string]::IsNullOrWhiteSpace($parent)) { $targetDir = $parent }
        } catch {}
    }

    $install = Install-LibreHardwareMonitor -TargetDir $targetDir
    if ($install.Installed -eq $true) {
        return $install.DllPath
    }

    return $RequestedDllPath
}

function Get-TemperatureFromHardwareMonitorWmi {
    $results = @()

    $namespaces = @(
        "root\LibreHardwareMonitor",
        "root\OpenHardwareMonitor"
    )

    foreach ($ns in $namespaces) {
        try {
            $sensors = @(Get-CimInstance -Namespace $ns -ClassName Sensor -ErrorAction Stop)

            foreach ($sensor in $sensors) {
                $sensorType = $null
                $value = $null
                $name = $null
                $identifier = $null
                $parent = $null

                try { $sensorType = [string]$sensor.SensorType } catch {}
                try { $value = $sensor.Value } catch {}
                try { $name = [string]$sensor.Name } catch {}
                try { $identifier = [string]$sensor.Identifier } catch {}
                try { $parent = [string]$sensor.Parent } catch {}

                if (($sensorType -eq "Temperature") -and ($value -ne $null)) {
                    $sensorName = $name
                    if ([string]::IsNullOrWhiteSpace($sensorName)) { $sensorName = $identifier }
                    if (![string]::IsNullOrWhiteSpace($parent)) { $sensorName = $parent + " - " + $sensorName }

                    $results += New-Object PSObject -Property @{
                        Sensor = $sensorName
                        Tipo = "HardwareMonitorWMI"
                        TemperaturaC = [math]::Round([double]$value, 1)
                        Fuente = $ns
                    }
                }
            }
        } catch {}
    }

    return $results
}

function Get-TemperatureFromWmi {
    $temps = @()

    try {
        $thermal = @(Get-CimInstance -Namespace root/wmi -ClassName MSAcpi_ThermalZoneTemperature -ErrorAction Stop)
        foreach ($t in $thermal) {
            if ($t.CurrentTemperature -ne $null) {
                $c = [math]::Round(($t.CurrentTemperature / 10) - 273.15, 1)
                if (($c -gt -20) -and ($c -lt 130)) {
                    $temps += New-Object PSObject -Property @{
                        Sensor = $t.InstanceName
                        Tipo = "ACPI"
                        TemperaturaC = $c
                        Fuente = "WMI_ACPI"
                    }
                }
            }
        }
    } catch {}

    return $temps
}


function Get-TemperatureFromNvidiaSmi {
    $results = @()
    $paths = @()

    $cmd = Get-Command "nvidia-smi.exe" -ErrorAction SilentlyContinue
    if ($cmd -ne $null) { $paths += $cmd.Source }

    $paths += "C:\Program Files\NVIDIA Corporation\NVSMI\nvidia-smi.exe"
    $paths += "C:\Windows\System32\nvidia-smi.exe"

    $exe = $null
    foreach ($p in $paths) {
        if (($p -ne $null) -and (Test-Path $p)) { $exe = $p; break }
    }

    if ($exe -eq $null) { return $results }

    try {
        $out = & $exe --query-gpu=name,temperature.gpu --format=csv,noheader,nounits 2>$null
        foreach ($line in $out) {
            if ($line -match "^(.+),\s*([0-9]+)") {
                $name = $Matches[1].Trim()
                $temp = [int]$Matches[2]
                $results += New-Object PSObject -Property @{
                    Sensor = "NVIDIA GPU - $name"
                    Tipo = "GPU"
                    TemperaturaC = $temp
                    Fuente = "nvidia-smi"
                }
            }
        }
    } catch {}

    return $results
}

function Get-TemperatureFromThermalZoneCounters {
    $results = @()

    try {
        $zones = @(Get-CimInstance -ClassName Win32_PerfFormattedData_Counters_ThermalZoneInformation -ErrorAction Stop)
        foreach ($z in $zones) {
            $temp = $null
            if ($z.HighPrecisionTemperature -ne $null) {
                # Contador en décimas Kelvin en varias versiones de Windows.
                $temp = [math]::Round(($z.HighPrecisionTemperature / 10) - 273.15, 1)
            } elseif ($z.Temperature -ne $null) {
                $temp = [math]::Round($z.Temperature, 1)
            }

            if (($temp -ne $null) -and ($temp -gt -20) -and ($temp -lt 130)) {
                $results += New-Object PSObject -Property @{
                    Sensor = "ThermalZone - $($z.Name)"
                    Tipo = "ThermalZone"
                    TemperaturaC = $temp
                    Fuente = "Win32_PerfFormattedData_Counters_ThermalZoneInformation"
                }
            }
        }
    } catch {}

    try {
        $probes = @(Get-CimInstance Win32_TemperatureProbe -ErrorAction Stop)
        foreach ($p in $probes) {
            if ($p.CurrentReading -ne $null) {
                $temp = [double]$p.CurrentReading
                if (($temp -gt -20) -and ($temp -lt 130)) {
                    $results += New-Object PSObject -Property @{
                        Sensor = "TemperatureProbe - $($p.Name)"
                        Tipo = "TemperatureProbe"
                        TemperaturaC = [math]::Round($temp, 1)
                        Fuente = "Win32_TemperatureProbe"
                    }
                }
            }
        }
    } catch {}

    return $results
}

function Get-TemperatureInfo {
    param([bool]$Advanced,[string]$DllPath)

    $temps = @()

    # 1) LibreHardwareMonitor/OpenHardwareMonitor por WMI si el programa está corriendo y expone sensores.
    $temps += @(Get-TemperatureFromHardwareMonitorWmi)

    # 2) nvidia-smi si hay GPU NVIDIA. En tu reporte aparece NVIDIA Share, por eso este método ayuda en notebooks con GPU dedicada.
    $temps += @(Get-TemperatureFromNvidiaSmi)

    # 3) Contadores ThermalZone / Win32_TemperatureProbe si Windows los expone.
    $temps += @(Get-TemperatureFromThermalZoneCounters)

    # 4) LibreHardwareMonitor por DLL. Requiere LibreHardwareMonitorLib.dll y dependencias en la misma carpeta.
    if ($Advanced -eq $true) {
        $resolvedDll = Resolve-LhmDllPath -RequestedPath $DllPath
        $temps += @(Get-TemperatureFromLhm -DllPath $resolvedDll)
    }

    # 5) ACPI nativo de Windows. Muchos equipos no exponen CPU real por aquí.
    $temps += @(Get-TemperatureFromWmi)

    # Filtrar duplicados simples y temperaturas inválidas.
    $clean = @()
    $seen = @{}
    foreach ($t in $temps) {
        $key = ([string]$t.Fuente) + "|" + ([string]$t.Sensor) + "|" + ([string]$t.TemperaturaC)
        if (!$seen.ContainsKey($key)) {
            $seen[$key] = $true
            $clean += $t
        }
    }

    # Si existen temperaturas reales, no contaminar resumen con errores informativos.
    $hasRealTemp = $false
    foreach ($t in $clean) {
        if ($t.TemperaturaC -ne $null) { $hasRealTemp = $true }
    }

    if ($hasRealTemp -eq $true) {
        $onlyGood = @()
        foreach ($t in $clean) {
            if ($t.TemperaturaC -ne $null) { $onlyGood += $t }
        }
        $clean = $onlyGood
    }

    if ($clean.Count -eq 0) {
        $clean += New-Object PSObject -Property @{
            Sensor = "No disponible"
            Tipo = "No disponible"
            TemperaturaC = $null
            Fuente = "Windows no expone sensores. Ejecute como Administrador, use -InstalarLibreHardwareMonitor, o ejecute LibreHardwareMonitor/OpenHardwareMonitor como Administrador."
        }
    }

    return $clean
}

function Get-MaxTemperature {
    param($Temps)
    $vals = @()
    foreach ($t in $Temps) {
        if ($t.TemperaturaC -ne $null) { $vals += [double]$t.TemperaturaC }
    }
    if ($vals.Count -gt 0) {
        return [math]::Round(($vals | Measure-Object -Maximum).Maximum, 1)
    }
    return $null
}

function Get-AvgTemperature {
    param($Samples)
    $vals = @()
    foreach ($s in $Samples) {
        if ($s.TempMaxC -ne $null) { $vals += [double]$s.TempMaxC }
    }
    if ($vals.Count -gt 0) {
        return [math]::Round(($vals | Measure-Object -Average).Average, 1)
    }
    return $null
}

function Get-CurrentMetrics {
    param([bool]$AdvancedTemp,[string]$LhmDll)

    $cpuValue = $null
    $ramUsedPct = $null
    $ramFreeGB = $null
    $diskQueue = $null
    $diskTime = $null

    try {
        $cpuPerf = Get-CimInstance Win32_PerfFormattedData_PerfOS_Processor -Filter "Name='_Total'" -ErrorAction Stop
        $cpuValue = [double]$cpuPerf.PercentProcessorTime
    } catch {
        try {
            $counter = Get-Counter '\Processor(_Total)\% Processor Time' -ErrorAction Stop
            $cpuValue = [math]::Round($counter.CounterSamples[0].CookedValue, 2)
        } catch {}
    }

    try {
        $os = Get-CimInstance Win32_OperatingSystem
        $totalKB = [double]$os.TotalVisibleMemorySize
        $freeKB = [double]$os.FreePhysicalMemory
        if ($totalKB -gt 0) {
            $ramUsedPct = [math]::Round((($totalKB - $freeKB) / $totalKB) * 100, 2)
            $ramFreeGB = [math]::Round(($freeKB * 1KB) / 1GB, 2)
        }
    } catch {}

    try {
        $disk = Get-CimInstance Win32_PerfFormattedData_PerfDisk_LogicalDisk | Where-Object { $_.Name -eq "_Total" } | Select-Object -First 1
        if ($disk -ne $null) {
            $diskQueue = $disk.AvgDiskQueueLength
            $diskTime = $disk.PercentDiskTime
        }
    } catch {}

    $temps = @(Get-TemperatureInfo -Advanced $AdvancedTemp -DllPath $LhmDll)
    $tempMax = Get-MaxTemperature -Temps $temps

    return New-Object PSObject -Property @{
        Fecha = Get-Date
        CPUPercent = $cpuValue
        RAMUsedPercent = $ramUsedPct
        RAMFreeGB = $ramFreeGB
        DiskQueue = $diskQueue
        DiskTimePercent = $diskTime
        TempMaxC = $tempMax
        Temperatures = $temps
    }
}

function Measure-Phase {
    param(
        [string]$Phase,
        [int]$Seconds,
        [int]$Interval,
        [bool]$AdvancedTemp,
        [string]$LhmDll
    )

    Write-Info "Fase $Phase durante $Seconds segundos. Muestreo cada $Interval segundos."

    $samples = @()
    $end = (Get-Date).AddSeconds($Seconds)

    while ((Get-Date) -lt $end) {
        $sample = Get-CurrentMetrics -AdvancedTemp $AdvancedTemp -LhmDll $LhmDll
        $sample | Add-Member -MemberType NoteProperty -Name Fase -Value $Phase -Force
        $samples += $sample

        $cpuText = $sample.CPUPercent
        $ramText = $sample.RAMUsedPercent
        $tempText = $sample.TempMaxC
        Write-Host ("[{0}] CPU={1}% RAM={2}% TempMax={3}C" -f $Phase,$cpuText,$ramText,$tempText)

        Start-Sleep -Seconds $Interval
    }

    $cpuVals = @()
    $ramVals = @()
    $tempVals = @()

    foreach ($s in $samples) {
        if ($s.CPUPercent -ne $null) { $cpuVals += [double]$s.CPUPercent }
        if ($s.RAMUsedPercent -ne $null) { $ramVals += [double]$s.RAMUsedPercent }
        if ($s.TempMaxC -ne $null) { $tempVals += [double]$s.TempMaxC }
    }

    $cpuAvg = $null
    $cpuMax = $null
    $ramAvg = $null
    $ramMax = $null
    $tempAvg = $null
    $tempMax = $null

    if ($cpuVals.Count -gt 0) {
        $cpuAvg = [math]::Round(($cpuVals | Measure-Object -Average).Average, 2)
        $cpuMax = [math]::Round(($cpuVals | Measure-Object -Maximum).Maximum, 2)
    }

    if ($ramVals.Count -gt 0) {
        $ramAvg = [math]::Round(($ramVals | Measure-Object -Average).Average, 2)
        $ramMax = [math]::Round(($ramVals | Measure-Object -Maximum).Maximum, 2)
    }

    if ($tempVals.Count -gt 0) {
        $tempAvg = [math]::Round(($tempVals | Measure-Object -Average).Average, 1)
        $tempMax = [math]::Round(($tempVals | Measure-Object -Maximum).Maximum, 1)
    }

    return New-Object PSObject -Property @{
        Fase = $Phase
        DuracionSegundos = $Seconds
        IntervaloSegundos = $Interval
        Samples = $samples
        CPUAveragePct = $cpuAvg
        CPUMaxPct = $cpuMax
        RAMAveragePct = $ramAvg
        RAMMaxPct = $ramMax
        TempAverageC = $tempAvg
        TempMaxC = $tempMax
    }
}

function Start-CpuStressJobs {
    param([int]$Workers,[int]$Seconds)

    $jobs = @()
    if ($Workers -lt 1) { return $jobs }

    Write-Info "Iniciando carga CPU controlada: $Workers workers por $Seconds segundos"

    for ($i = 1; $i -le $Workers; $i++) {
        $job = Start-Job -ScriptBlock {
            param($Duration)
            $end = (Get-Date).AddSeconds($Duration)
            $x = 0.0
            $iterations = 0
            while ((Get-Date) -lt $end) {
                for ($j = 1; $j -le 50000; $j++) {
                    $x += [Math]::Sqrt($j) * [Math]::Sin($j)
                }
                $iterations++
            }
            return $iterations
        } -ArgumentList $Seconds
        $jobs += $job
    }

    return $jobs
}

function Start-DiskStressJob {
    param([string]$Folder,[int]$Seconds,[int]$FileMB)

    Write-Info "Iniciando prueba de disco controlada: archivo $FileMB MB por $Seconds segundos"

    $job = Start-Job -ScriptBlock {
        param($Folder,$Duration,$FileMB)

        $file = Join-Path $Folder "disk_stress_test.bin"
        $chunkMB = 4
        $buffer = New-Object byte[] ($chunkMB * 1MB)
        (New-Object Random).NextBytes($buffer)

        $end = (Get-Date).AddSeconds($Duration)
        $writeMB = 0
        $readMB = 0
        $cycles = 0

        try {
            while ((Get-Date) -lt $end) {
                $fs = [IO.File]::Open($file, [IO.FileMode]::Create, [IO.FileAccess]::Write)
                for ($i = 0; $i -lt ($FileMB / $chunkMB); $i++) {
                    $fs.Write($buffer, 0, $buffer.Length)
                    $writeMB += $chunkMB
                }
                $fs.Flush()
                $fs.Close()

                $readBuffer = New-Object byte[] ($chunkMB * 1MB)
                $fsr = [IO.File]::Open($file, [IO.FileMode]::Open, [IO.FileAccess]::Read)
                while ($fsr.Read($readBuffer, 0, $readBuffer.Length) -gt 0) {
                    $readMB += $chunkMB
                }
                $fsr.Close()
                $cycles++
            }
        } catch {
        } finally {
            Remove-Item $file -Force -ErrorAction SilentlyContinue
        }

        return New-Object PSObject -Property @{
            WriteMB = $writeMB
            ReadMB = $readMB
            Cycles = $cycles
        }
    } -ArgumentList $Folder,$Seconds,$FileMB

    return $job
}

function Invoke-SeriousBenchmark {
    param(
        [string]$Folder,
        [int]$Duration,
        [int]$Interval,
        [bool]$AdvancedTemp,
        [string]$LhmDll,
        [int]$MaxWorkers,
        [int]$DiskFileMB
    )

    $threads = 1
    try {
        $cpu = Get-CimInstance Win32_Processor | Select-Object -First 1
        if ($cpu.NumberOfLogicalProcessors -ne $null) { $threads = [int]$cpu.NumberOfLogicalProcessors }
    } catch {}

    $workers = $MaxWorkers
    if ($workers -le 0) {
        $workers = $threads
        if ($workers -gt 8) { $workers = 8 }
    }

    if ($workers -lt 1) { $workers = 1 }

    $cpuJobs = @(Start-CpuStressJobs -Workers $workers -Seconds $Duration)
    $diskJob = Start-DiskStressJob -Folder $Folder -Seconds $Duration -FileMB $DiskFileMB

    $phase = Measure-Phase -Phase "Carga" -Seconds $Duration -Interval $Interval -AdvancedTemp $AdvancedTemp -LhmDll $LhmDll

    Write-Info "Esperando cierre de trabajos de carga"

    foreach ($j in $cpuJobs) {
        Wait-Job $j -Timeout 15 | Out-Null
    }

    Wait-Job $diskJob -Timeout 15 | Out-Null

    $cpuIterations = 0
    foreach ($j in $cpuJobs) {
        try {
            $r = Receive-Job $j -ErrorAction SilentlyContinue
            foreach ($item in $r) {
                if ($item -ne $null) { $cpuIterations += [int]$item }
            }
        } catch {}
        Remove-Job $j -Force -ErrorAction SilentlyContinue
    }

    $diskResult = $null
    try { $diskResult = Receive-Job $diskJob -ErrorAction SilentlyContinue } catch {}
    Remove-Job $diskJob -Force -ErrorAction SilentlyContinue

    $writeMB = 0
    $readMB = 0
    $cycles = 0

    if ($diskResult -ne $null) {
        foreach ($dr in $diskResult) {
            if ($dr.WriteMB -ne $null) { $writeMB += [double]$dr.WriteMB }
            if ($dr.ReadMB -ne $null) { $readMB += [double]$dr.ReadMB }
            if ($dr.Cycles -ne $null) { $cycles += [int]$dr.Cycles }
        }
    }

    $writeMBs = $null
    $readMBs = $null
    if ($Duration -gt 0) {
        $writeMBs = [math]::Round($writeMB / $Duration, 2)
        $readMBs = [math]::Round($readMB / $Duration, 2)
    }

    return New-Object PSObject -Property @{
        Phase = $phase
        CpuWorkers = $workers
        CpuIterations = $cpuIterations
        DiskFileMB = $DiskFileMB
        DiskWriteTotalMB = $writeMB
        DiskReadTotalMB = $readMB
        DiskCycles = $cycles
        DiskWriteMBs = $writeMBs
        DiskReadMBs = $readMBs
    }
}

function Get-DiskInfo {
    Write-Info "Recolectando discos"
    $volumes = @()
    $physical = @()
    $perf = @()

    try {
        $logicalDisks = @(Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3")
        foreach ($d in $logicalDisks) {
            $sizeGB = 0
            $freeGB = 0
            $freePct = 0
            $usedPct = 0

            if ($d.Size -ne $null) {
                $sizeGB = [math]::Round($d.Size / 1GB, 2)
                $usedPct = [math]::Round((($d.Size - $d.FreeSpace) / $d.Size) * 100, 2)
            }

            if ($d.FreeSpace -ne $null) { $freeGB = [math]::Round($d.FreeSpace / 1GB, 2) }

            if (($d.Size -ne $null) -and ($d.FreeSpace -ne $null) -and ($d.Size -gt 0)) {
                $freePct = [math]::Round(($d.FreeSpace / $d.Size) * 100, 2)
            }

            $volumes += New-Object PSObject -Property @{
                Unidad = $d.DeviceID
                Nombre = $d.VolumeName
                Sistema = $d.FileSystem
                TamanoGB = $sizeGB
                LibreGB = $freeGB
                LibrePct = $freePct
                UsadoPct = $usedPct
            }
        }
    } catch {}

    try {
        $diskDrives = @(Get-CimInstance Win32_DiskDrive)
        foreach ($d in $diskDrives) {
            $sizeGB = $null
            if ($d.Size -ne $null) { $sizeGB = [math]::Round($d.Size / 1GB, 2) }

            $physical += New-Object PSObject -Property @{
                Modelo = $d.Model
                Interface = $d.InterfaceType
                Serial = $d.SerialNumber
                TamanoGB = $sizeGB
                MediaType = $d.MediaType
                Status = $d.Status
                HealthStatus = $null
                PhysicalMediaType = $null
            }
        }
    } catch {}

    try {
        $physicalDisks = @(Get-PhysicalDisk)
        foreach ($p in $physicalDisks) {
            $sizeGB = $null
            if ($p.Size -ne $null) { $sizeGB = [math]::Round($p.Size / 1GB, 2) }

            $op = $null
            if ($p.OperationalStatus -ne $null) { $op = $p.OperationalStatus -join "," }

            $physical += New-Object PSObject -Property @{
                Modelo = $p.FriendlyName
                Interface = $p.BusType
                Serial = $p.SerialNumber
                TamanoGB = $sizeGB
                MediaType = $p.MediaType
                Status = $op
                HealthStatus = $p.HealthStatus
                PhysicalMediaType = $p.MediaType
            }
        }
    } catch {}

    try {
        $perfDisks = @(Get-CimInstance Win32_PerfFormattedData_PerfDisk_LogicalDisk | Where-Object { $_.Name -ne "_Total" })
        foreach ($p in $perfDisks) {
            $readMs = $null
            $writeMs = $null
            if ($p.AvgDisksecPerRead -ne $null) { $readMs = [math]::Round($p.AvgDisksecPerRead * 1000, 2) }
            if ($p.AvgDisksecPerWrite -ne $null) { $writeMs = [math]::Round($p.AvgDisksecPerWrite * 1000, 2) }

            $perf += New-Object PSObject -Property @{
                Unidad = $p.Name
                DiskReadsPerSec = $p.DiskReadsPerSec
                DiskWritesPerSec = $p.DiskWritesPerSec
                AvgDiskQueueLength = $p.AvgDiskQueueLength
                PercentDiskTime = $p.PercentDiskTime
                AvgDiskSecPerRead_ms = $readMs
                AvgDiskSecPerWrite_ms = $writeMs
            }
        }
    } catch {}

    return New-Object PSObject -Property @{
        Volumes = $volumes
        PhysicalDisk = $physical
        DiskPerf = $perf
    }
}

function Get-ProcessesInfo {
    Write-Info "Recolectando procesos"
    $items = @()

    try {
        foreach ($p in Get-Process) {
            $path = $null
            $startTime = $null
            $cpuSeconds = 0
            try { $path = $p.Path } catch {}
            try { $startTime = $p.StartTime } catch {}
            if ($p.CPU -ne $null) { $cpuSeconds = [math]::Round($p.CPU, 2) }

            $items += New-Object PSObject -Property @{
                Name = $p.ProcessName
                Id = $p.Id
                CPUSeconds = $cpuSeconds
                RAMMB = [math]::Round($p.WorkingSet64 / 1MB, 2)
                Path = $path
                StartTime = $startTime
            }
        }
    } catch {}

    return New-Object PSObject -Property @{
        TotalProcesos = $items.Count
        TopCPU = @($items | Sort-Object CPUSeconds -Descending | Select-Object -First 20)
        TopRAM = @($items | Sort-Object RAMMB -Descending | Select-Object -First 20)
    }
}

function Get-SmartCtlInfo {
    param([string]$SmartCtlPath)
    Write-Info "Intentando recolectar SMART con smartctl"
    $items = @()
    $cmd = Get-Command $SmartCtlPath -ErrorAction SilentlyContinue

    if (($cmd -eq $null) -and (!(Test-Path $SmartCtlPath))) {
        $items += New-Object PSObject -Property @{
            Device = "smartctl"
            Available = $false
            Health = $null
            TemperatureC = $null
            Raw = "smartctl no encontrado"
        }
        return $items
    }

    try {
        $scan = & $SmartCtlPath --scan-open 2>$null
        foreach ($line in $scan) {
            if ($line -match "^(?<dev>\S+)") {
                $dev = $Matches.dev
                $raw = (& $SmartCtlPath -a $dev 2>&1) -join "`n"
                $health = $null
                if ($raw -match "SMART overall-health self-assessment test result:\s*(.+)") {
                    $health = $Matches[1].Trim()
                } elseif ($raw -match "SMART Health Status:\s*(.+)") {
                    $health = $Matches[1].Trim()
                }

                $temp = $null
                if ($raw -match "Temperature_Celsius\s+\S+\s+\S+\s+\S+\s+\S+\s+\S+\s+\S+\s+\S+\s+\S+\s+\S+\s+(?<t>\d+)") {
                    $temp = [int]$Matches.t
                } elseif ($raw -match "Current Drive Temperature:\s*(?<t>\d+)") {
                    $temp = [int]$Matches.t
                } elseif ($raw -match "Temperature:\s*(?<t>\d+)\s+Celsius") {
                    $temp = [int]$Matches.t
                }

                $items += New-Object PSObject -Property @{
                    Device = $dev
                    Available = $true
                    Health = $health
                    TemperatureC = $temp
                    Raw = $raw
                }
            }
        }
    } catch {
        $items += New-Object PSObject -Property @{
            Device = "smartctl"
            Available = $false
            Health = $null
            TemperatureC = $null
            Raw = $_.Exception.Message
        }
    }

    if ($items.Count -eq 0) {
        $items += New-Object PSObject -Property @{
            Device = "smartctl"
            Available = $false
            Health = $null
            TemperatureC = $null
            Raw = "smartctl no devolvio discos"
        }
    }

    return $items
}


function Get-TrimStatus {
    Write-Info "Recolectando estado TRIM / DisableDeleteNotify"

    $items = @()

    try {
        $raw = fsutil behavior query DisableDeleteNotify 2>$null
        foreach ($line in $raw) {
            $enabled = $null
            $name = "DisableDeleteNotify"

            if ($line -match "NTFS\s+DisableDeleteNotify\s*=\s*(\d+)") {
                $name = "NTFS"
                if ([int]$Matches[1] -eq 0) { $enabled = $true } else { $enabled = $false }
            } elseif ($line -match "ReFS\s+DisableDeleteNotify\s*=\s*(\d+)") {
                $name = "ReFS"
                if ([int]$Matches[1] -eq 0) { $enabled = $true } else { $enabled = $false }
            } elseif ($line -match "DisableDeleteNotify\s*=\s*(\d+)") {
                if ([int]$Matches[1] -eq 0) { $enabled = $true } else { $enabled = $false }
            }

            if ($line -match "DisableDeleteNotify") {
                $items += New-Object PSObject -Property @{
                    FileSystem = $name
                    TrimEnabled = $enabled
                    Raw = $line
                }
            }
        }
    } catch {
        $items += New-Object PSObject -Property @{
            FileSystem = "N/D"
            TrimEnabled = $null
            Raw = $_.Exception.Message
        }
    }

    if ($items.Count -eq 0) {
        $items += New-Object PSObject -Property @{
            FileSystem = "N/D"
            TrimEnabled = $null
            Raw = "No disponible"
        }
    }

    return $items
}

function Get-StorageReliabilityInfo {
    Write-Info "Recolectando Storage Reliability Counters"

    $items = @()

    try {
        $counters = @(Get-PhysicalDisk | Get-StorageReliabilityCounter -ErrorAction Stop)
        foreach ($c in $counters) {
            $items += New-Object PSObject -Property @{
                DeviceId = $c.DeviceId
                TemperatureC = $c.Temperature
                TemperatureMaxC = $c.TemperatureMax
                Wear = $c.Wear
                PowerOnHours = $c.PowerOnHours
                ReadErrorsTotal = $c.ReadErrorsTotal
                WriteErrorsTotal = $c.WriteErrorsTotal
                ReadLatencyMax = $c.ReadLatencyMax
                WriteLatencyMax = $c.WriteLatencyMax
                FlushLatencyMax = $c.FlushLatencyMax
                LoadUnloadCycleCount = $c.LoadUnloadCycleCount
                StartStopCycleCount = $c.StartStopCycleCount
            }
        }
    } catch {
        $items += New-Object PSObject -Property @{
            DeviceId = "N/D"
            TemperatureC = $null
            TemperatureMaxC = $null
            Wear = $null
            PowerOnHours = $null
            ReadErrorsTotal = $null
            WriteErrorsTotal = $null
            ReadLatencyMax = $null
            WriteLatencyMax = $null
            FlushLatencyMax = $null
            LoadUnloadCycleCount = $null
            StartStopCycleCount = $null
            Error = $_.Exception.Message
        }
    }

    return $items
}

function Get-WmiSmartBasic {
    Write-Info "Recolectando SMART basico por WMI"

    $items = @()

    try {
        $status = @(Get-CimInstance -Namespace root\wmi -ClassName MSStorageDriver_FailurePredictStatus -ErrorAction Stop)
        foreach ($s in $status) {
            $items += New-Object PSObject -Property @{
                InstanceName = $s.InstanceName
                PredictFailure = $s.PredictFailure
                Reason = $s.Reason
            }
        }
    } catch {
        $items += New-Object PSObject -Property @{
            InstanceName = "N/D"
            PredictFailure = $null
            Reason = $null
            Error = $_.Exception.Message
        }
    }

    return $items
}

function Get-DiskTemperatureFromSmartCtl {
    param([string]$SmartCtlPath)

    $items = @()
    $cmd = Get-Command $SmartCtlPath -ErrorAction SilentlyContinue

    if (($cmd -eq $null) -and (!(Test-Path $SmartCtlPath))) {
        return $items
    }

    try {
        $scan = & $SmartCtlPath --scan-open 2>$null

        foreach ($line in $scan) {
            if ($line -match "^(?<dev>\S+)") {
                $dev = $Matches.dev
                $raw = (& $SmartCtlPath -a $dev 2>&1) -join "`n"

                $health = $null
                if ($raw -match "SMART overall-health self-assessment test result:\s*(.+)") {
                    $health = $Matches[1].Trim()
                } elseif ($raw -match "SMART Health Status:\s*(.+)") {
                    $health = $Matches[1].Trim()
                }

                $temp = $null
                if ($raw -match "Temperature_Celsius\s+\S+\s+\S+\s+\S+\s+\S+\s+\S+\s+\S+\s+\S+\s+\S+\s+\S+\s+(?<t>\d+)") {
                    $temp = [int]$Matches.t
                } elseif ($raw -match "Current Drive Temperature:\s*(?<t>\d+)") {
                    $temp = [int]$Matches.t
                } elseif ($raw -match "Temperature:\s*(?<t>\d+)\s+Celsius") {
                    $temp = [int]$Matches.t
                }

                $items += New-Object PSObject -Property @{
                    Device = $dev
                    Health = $health
                    TemperatureC = $temp
                    Raw = $raw
                }
            }
        }
    } catch {}

    return $items
}

function Get-DefragAnalysis {
    param([bool]$Enabled)

    Write-Info "Recolectando fragmentacion / optimizacion de volumenes"

    $items = @()

    try {
        foreach ($v in @(Get-Volume | Where-Object { $_.DriveLetter -ne $null })) {
            $drive = ([string]$v.DriveLetter) + ":"
            $mediaType = $null
            $needsDefrag = $null
            $fragmentationPct = $null
            $raw = "No ejecutado. Use -AnalizarFragmentacion para ejecutar defrag /A."

            try {
                $part = Get-Partition -DriveLetter $v.DriveLetter -ErrorAction Stop | Select-Object -First 1
                $disk = Get-Disk -Number $part.DiskNumber -ErrorAction Stop
                $mediaType = $disk.MediaType
            } catch {}

            if ($Enabled -eq $true) {
                try {
                    $rawLines = defrag $drive /A /U /V 2>&1
                    $raw = ($rawLines -join "`n")

                    if ($raw -match "Total fragmented space\s*=\s*([0-9]+)%") {
                        $fragmentationPct = [int]$Matches[1]
                    } elseif ($raw -match "Espacio fragmentado total\s*=\s*([0-9]+)%") {
                        $fragmentationPct = [int]$Matches[1]
                    } elseif ($raw -match "fragmentation\s*:\s*([0-9]+)%") {
                        $fragmentationPct = [int]$Matches[1]
                    }

                    if ($raw -match "You do not need to defragment|No es necesario desfragmentar|no necesita desfragmentar") {
                        $needsDefrag = $false
                    } elseif ($raw -match "should be defragmented|debe desfragmentar|se recomienda") {
                        $needsDefrag = $true
                    }
                } catch {
                    $raw = $_.Exception.Message
                }
            }

            $items += New-Object PSObject -Property @{
                Drive = $drive
                FileSystem = $v.FileSystem
                DriveType = $v.DriveType
                HealthStatus = $v.HealthStatus
                SizeGB = if ($v.Size -ne $null) { [math]::Round($v.Size / 1GB, 2) } else { $null }
                SizeRemainingGB = if ($v.SizeRemaining -ne $null) { [math]::Round($v.SizeRemaining / 1GB, 2) } else { $null }
                MediaType = $mediaType
                FragmentationPercent = $fragmentationPct
                NeedsDefrag = $needsDefrag
                Raw = $raw
            }
        }
    } catch {
        $items += New-Object PSObject -Property @{
            Drive = "N/D"
            FileSystem = $null
            DriveType = $null
            HealthStatus = $null
            SizeGB = $null
            SizeRemainingGB = $null
            MediaType = $null
            FragmentationPercent = $null
            NeedsDefrag = $null
            Raw = $_.Exception.Message
        }
    }

    return $items
}

function Get-ChkdskScanInfo {
    param([bool]$Enabled)

    Write-Info "Recolectando estado CHKDSK scan"

    $items = @()

    try {
        foreach ($v in @(Get-Volume | Where-Object { $_.DriveLetter -ne $null })) {
            $drive = ([string]$v.DriveLetter) + ":"
            $raw = "No ejecutado. Use -IncluirChkdskScan para ejecutar chkdsk /scan."
            $status = "No ejecutado"

            if ($Enabled -eq $true) {
                try {
                    $rawLines = chkdsk $drive /scan 2>&1
                    $raw = ($rawLines -join "`n")
                    if ($raw -match "Windows has scanned the file system and found no problems|no se encontraron problemas|No further action is required") {
                        $status = "Sin problemas detectados"
                    } elseif ($raw -match "found problems|encontr.*problemas|corrupt") {
                        $status = "Problemas detectados"
                    } else {
                        $status = "Revisar salida"
                    }
                } catch {
                    $raw = $_.Exception.Message
                    $status = "Error"
                }
            }

            $items += New-Object PSObject -Property @{
                Drive = $drive
                Status = $status
                Raw = $raw
            }
        }
    } catch {
        $items += New-Object PSObject -Property @{
            Drive = "N/D"
            Status = "Error"
            Raw = $_.Exception.Message
        }
    }

    return $items
}

function Get-BitLockerInfoSafe {
    Write-Info "Recolectando BitLocker por volumen"

    $items = @()

    try {
        foreach ($b in Get-BitLockerVolume -ErrorAction Stop) {
            $items += New-Object PSObject -Property @{
                MountPoint = $b.MountPoint
                VolumeStatus = $b.VolumeStatus
                ProtectionStatus = $b.ProtectionStatus
                EncryptionMethod = $b.EncryptionMethod
                EncryptionPercentage = $b.EncryptionPercentage
                LockStatus = $b.LockStatus
            }
        }
    } catch {
        $items += New-Object PSObject -Property @{
            MountPoint = "N/D"
            VolumeStatus = $null
            ProtectionStatus = $null
            EncryptionMethod = $null
            EncryptionPercentage = $null
            LockStatus = $null
            Error = $_.Exception.Message
        }
    }

    return $items
}

function Get-DiskAdvancedInfo {
    param(
        [bool]$RunFragmentation,
        [bool]$RunChkdskScan,
        [bool]$UseSmartCtl,
        [string]$SmartCtlPath
    )

    Write-Info "Recolectando integraciones avanzadas de disco"

    $storageReliability = @(Get-StorageReliabilityInfo)
    $wmiSmart = @(Get-WmiSmartBasic)
    $trim = @(Get-TrimStatus)
    $fragmentation = @(Get-DefragAnalysis -Enabled $RunFragmentation)
    $chkdsk = @(Get-ChkdskScanInfo -Enabled $RunChkdskScan)
    $bitlocker = @(Get-BitLockerInfoSafe)
    $smartCtlTemp = @()

    if ($UseSmartCtl -eq $true) {
        $smartCtlTemp = @(Get-DiskTemperatureFromSmartCtl -SmartCtlPath $SmartCtlPath)
    }

    $maxDiskTemp = $null
    $temps = @()

    foreach ($r in $storageReliability) {
        if ($r.TemperatureC -ne $null) { $temps += [double]$r.TemperatureC }
        if ($r.TemperatureMaxC -ne $null) { $temps += [double]$r.TemperatureMaxC }
    }

    foreach ($s in $smartCtlTemp) {
        if ($s.TemperatureC -ne $null) { $temps += [double]$s.TemperatureC }
    }

    if ($temps.Count -gt 0) {
        $maxDiskTemp = [math]::Round(($temps | Measure-Object -Maximum).Maximum, 1)
    }

    return New-Object PSObject -Property @{
        StorageReliability = $storageReliability
        WmiSmart = $wmiSmart
        Trim = $trim
        Fragmentation = $fragmentation
        Chkdsk = $chkdsk
        BitLocker = $bitlocker
        SmartCtlDiskTemperature = $smartCtlTemp
        MaxDiskTemperatureC = $maxDiskTemp
    }
}

function Get-StartupPrograms {
    Write-Info "Recolectando programas de inicio"
    $items = @()
    try {
        $startup = @(Get-CimInstance Win32_StartupCommand)
        foreach ($s in $startup) {
            $items += New-Object PSObject -Property @{
                Nombre = $s.Name
                Comando = $s.Command
                Ubicacion = $s.Location
                Usuario = $s.User
            }
        }
    } catch {}
    return $items
}

function Get-ServicesInfo {
    Write-Info "Recolectando servicios"
    $services = @()

    try {
        foreach ($svc in Get-Service) {
            $startType = $null
            try { $startType = $svc.StartType.ToString() } catch {}

            $services += New-Object PSObject -Property @{
                Nombre = $svc.Name
                DisplayName = $svc.DisplayName
                Estado = $svc.Status.ToString()
                TipoInicio = $startType
            }
        }
    } catch {}

    $running = @($services | Where-Object { $_.Estado -eq "Running" }).Count
    $autoRunning = @($services | Where-Object { ($_.TipoInicio -eq "Automatic") -and ($_.Estado -eq "Running") }).Count
    $autoStopped = @($services | Where-Object { ($_.TipoInicio -eq "Automatic") -and ($_.Estado -ne "Running") }).Count

    return New-Object PSObject -Property @{
        TotalServicios = $services.Count
        ServiciosRunning = $running
        AutomaticosEjecutando = $autoRunning
        AutomaticosDetenidos = $autoStopped
        ListaServicios = $services
    }
}

function Get-CriticalEvents {
    param([int]$Days = 7)
    Write-Info "Recolectando eventos relevantes de los ultimos $Days dias"
    $events = @()
    $start = (Get-Date).AddDays(-$Days)
    $providers = @("Microsoft-Windows-Kernel-Power","Disk","Ntfs","Microsoft-Windows-WHEA-Logger","Service Control Manager","Microsoft-Windows-WindowsUpdateClient","Display")

    foreach ($provider in $providers) {
        try {
            $ev = @(Get-WinEvent -FilterHashtable @{
                LogName = "System"
                ProviderName = $provider
                StartTime = $start
            } -ErrorAction Stop | Select-Object -First 80)

            foreach ($e in $ev) {
                $include = $false
                if (($e.LevelDisplayName -eq "Critical") -or ($e.LevelDisplayName -eq "Error") -or ($e.LevelDisplayName -eq "Warning")) { $include = $true }
                if ($e.Id -in @(41,7,51,55,129,153,18,19)) { $include = $true }

                if ($include -eq $true) {
                    $msg = ""
                    if ($e.Message -ne $null) {
                        $msg = (($e.Message -replace "`r|`n", " ") -replace "\s+", " ")
                        if ($msg.Length -gt 600) { $msg = $msg.Substring(0, 600) }
                    }

                    $events += New-Object PSObject -Property @{
                        Fecha = $e.TimeCreated
                        Nivel = $e.LevelDisplayName
                        Provider = $e.ProviderName
                        Id = $e.Id
                        Mensaje = $msg
                    }
                }
            }
        } catch {}
    }

    $events = @($events | Sort-Object Fecha -Descending)

    return New-Object PSObject -Property @{
        DiasAnalizados = $Days
        TotalEventos = $events.Count
        Eventos = $events
    }
}

function Get-SecurityInfo {
    Write-Info "Recolectando seguridad basica"
    $def = New-Object PSObject -Property @{ Nota = "No disponible" }
    $fw = @()
    $bit = @()
    $admins = @()

    try {
        $d = Get-MpComputerStatus -ErrorAction Stop
        $def = New-Object PSObject -Property @{
            AntivirusEnabled = $d.AntivirusEnabled
            RealTimeProtection = $d.RealTimeProtectionEnabled
            AMServiceEnabled = $d.AMServiceEnabled
            QuickScanAge = $d.QuickScanAge
            FullScanAge = $d.FullScanAge
        }
    } catch {}

    try {
        foreach ($profile in Get-NetFirewallProfile) {
            $fw += New-Object PSObject -Property @{
                Perfil = $profile.Name
                Enabled = $profile.Enabled
            }
        }
    } catch {}

    try {
        foreach ($b in Get-BitLockerVolume -ErrorAction Stop) {
            $bit += New-Object PSObject -Property @{
                MountPoint = $b.MountPoint
                ProtectionStatus = $b.ProtectionStatus
                VolumeStatus = $b.VolumeStatus
                EncryptionMethod = $b.EncryptionMethod
            }
        }
    } catch {}

    try {
        foreach ($a in Get-LocalGroupMember -Group "Administrators" -ErrorAction Stop) {
            $admins += New-Object PSObject -Property @{
                Nombre = $a.Name
                Tipo = $a.ObjectClass
            }
        }
    } catch {
        try {
            foreach ($a in Get-LocalGroupMember -Group "Administradores" -ErrorAction Stop) {
                $admins += New-Object PSObject -Property @{
                    Nombre = $a.Name
                    Tipo = $a.ObjectClass
                }
            }
        } catch {}
    }

    return New-Object PSObject -Property @{
        Defender = $def
        Firewall = $fw
        BitLocker = $bit
        Admins = $admins
    }
}

function Add-Finding {
    param([ref]$Score,[ref]$Findings,[int]$Penalty,[string]$Area,[string]$Severity,[string]$Message)
    $Score.Value = $Score.Value - $Penalty
    $arr = $Findings.Value
    $arr += New-Object PSObject -Property @{
        Area = $Area
        Severidad = $Severity
        Penalidad = $Penalty
        Mensaje = $Message
    }
    $Findings.Value = $arr
}

function Get-Score {
    param($Reposo,$Benchmark,$Disks,$DiskAdvanced,$TempsFinal,$Smart,$Startup,$Events,$Security)

    $score = 100
    $findings = @()

    if ($Reposo.CPUAveragePct -ne $null) {
        if ($Reposo.CPUAveragePct -gt 40) {
            Add-Finding ([ref]$score) ([ref]$findings) 10 "CPU Reposo" "Alto" "CPU alta en reposo: $($Reposo.CPUAveragePct)%."
        } elseif ($Reposo.CPUAveragePct -gt 20) {
            Add-Finding ([ref]$score) ([ref]$findings) 5 "CPU Reposo" "Medio" "CPU moderada en reposo: $($Reposo.CPUAveragePct)%."
        }
    }

    if ($Reposo.RAMAveragePct -ne $null) {
        if ($Reposo.RAMAveragePct -gt 90) {
            Add-Finding ([ref]$score) ([ref]$findings) 15 "RAM" "Critico" "RAM critica en reposo: $($Reposo.RAMAveragePct)%."
        } elseif ($Reposo.RAMAveragePct -gt 80) {
            Add-Finding ([ref]$score) ([ref]$findings) 10 "RAM" "Alto" "RAM alta en reposo: $($Reposo.RAMAveragePct)%."
        } elseif ($Reposo.RAMAveragePct -gt 65) {
            Add-Finding ([ref]$score) ([ref]$findings) 5 "RAM" "Medio" "RAM moderada en reposo: $($Reposo.RAMAveragePct)%."
        }
    }

    if ($Benchmark.Phase.TempMaxC -ne $null) {
        if ($Benchmark.Phase.TempMaxC -gt 95) {
            Add-Finding ([ref]$score) ([ref]$findings) 20 "Temperatura Carga" "Critico" "Temperatura maxima critica durante prueba: $($Benchmark.Phase.TempMaxC) C."
        } elseif ($Benchmark.Phase.TempMaxC -gt 85) {
            Add-Finding ([ref]$score) ([ref]$findings) 15 "Temperatura Carga" "Alto" "Temperatura maxima alta durante prueba: $($Benchmark.Phase.TempMaxC) C."
        } elseif ($Benchmark.Phase.TempMaxC -gt 75) {
            Add-Finding ([ref]$score) ([ref]$findings) 7 "Temperatura Carga" "Medio" "Temperatura maxima moderada durante prueba: $($Benchmark.Phase.TempMaxC) C."
        }
    }

    if (($Reposo.TempAverageC -ne $null) -and ($Benchmark.Phase.TempMaxC -ne $null)) {
        $deltaTemp = [math]::Round($Benchmark.Phase.TempMaxC - $Reposo.TempAverageC, 1)
        if ($deltaTemp -gt 40) {
            Add-Finding ([ref]$score) ([ref]$findings) 10 "Temperatura Delta" "Alto" "Aumento termico fuerte: +$deltaTemp C entre reposo y carga."
        } elseif ($deltaTemp -gt 30) {
            Add-Finding ([ref]$score) ([ref]$findings) 5 "Temperatura Delta" "Medio" "Aumento termico moderado: +$deltaTemp C entre reposo y carga."
        }
    }

    foreach ($v in $Disks.Volumes) {
        if ($v.LibrePct -lt 10) {
            Add-Finding ([ref]$score) ([ref]$findings) 10 "Disco" "Critico" "Unidad $($v.Unidad) con espacio critico: $($v.LibrePct)%."
        } elseif ($v.LibrePct -lt 15) {
            Add-Finding ([ref]$score) ([ref]$findings) 6 "Disco" "Alto" "Unidad $($v.Unidad) con poco espacio: $($v.LibrePct)%."
        } elseif ($v.LibrePct -lt 25) {
            Add-Finding ([ref]$score) ([ref]$findings) 3 "Disco" "Medio" "Unidad $($v.Unidad) con espacio bajo: $($v.LibrePct)%."
        }
    }

    foreach ($d in $Disks.PhysicalDisk) {
        if (($d.HealthStatus -ne $null) -and ($d.HealthStatus -ne "Healthy")) {
            Add-Finding ([ref]$score) ([ref]$findings) 15 "Disco" "Critico" "Disco $($d.Modelo) salud no saludable: $($d.HealthStatus)."
        }
        if (($d.PhysicalMediaType -eq "HDD") -or ($d.MediaType -like "*Hard*")) {
            Add-Finding ([ref]$score) ([ref]$findings) 5 "Disco" "Medio" "Disco mecanico detectado: $($d.Modelo)."
        }
    }

    if ($Benchmark.DiskWriteMBs -ne $null) {
        if ($Benchmark.DiskWriteMBs -lt 25) {
            Add-Finding ([ref]$score) ([ref]$findings) 10 "Benchmark Disco" "Alto" "Escritura promedio muy baja durante prueba: $($Benchmark.DiskWriteMBs) MB/s."
        } elseif ($Benchmark.DiskWriteMBs -lt 70) {
            Add-Finding ([ref]$score) ([ref]$findings) 5 "Benchmark Disco" "Medio" "Escritura promedio baja/moderada durante prueba: $($Benchmark.DiskWriteMBs) MB/s."
        }
    }

    foreach ($s in $Smart) {
        if (($s.Available -eq $true) -and ($s.Health -ne $null) -and ($s.Health -notmatch "PASSED|OK")) {
            Add-Finding ([ref]$score) ([ref]$findings) 20 "SMART" "Critico" "SMART no saludable en $($s.Device): $($s.Health)."
        }
        if ($s.TemperatureC -gt 60) {
            Add-Finding ([ref]$score) ([ref]$findings) 10 "SMART" "Alto" "Temperatura de disco alta: $($s.TemperatureC) C."
        }
    }


    if ($DiskAdvanced -ne $null) {
        foreach ($w in $DiskAdvanced.WmiSmart) {
            if ($w.PredictFailure -eq $true) {
                Add-Finding ([ref]$score) ([ref]$findings) 25 "SMART WMI" "Critico" "Windows predice falla SMART: $($w.InstanceName). Reason=$($w.Reason)"
            }
        }

        foreach ($r in $DiskAdvanced.StorageReliability) {
            if ($r.ReadErrorsTotal -gt 0) {
                Add-Finding ([ref]$score) ([ref]$findings) 8 "Disco Errores" "Alto" "Errores de lectura reportados por Storage Reliability: $($r.ReadErrorsTotal)."
            }

            if ($r.WriteErrorsTotal -gt 0) {
                Add-Finding ([ref]$score) ([ref]$findings) 8 "Disco Errores" "Alto" "Errores de escritura reportados por Storage Reliability: $($r.WriteErrorsTotal)."
            }

            if ($r.Wear -ne $null) {
                if ($r.Wear -gt 80) {
                    Add-Finding ([ref]$score) ([ref]$findings) 15 "Desgaste SSD" "Alto" "Desgaste SSD alto reportado: $($r.Wear)."
                } elseif ($r.Wear -gt 60) {
                    Add-Finding ([ref]$score) ([ref]$findings) 8 "Desgaste SSD" "Medio" "Desgaste SSD moderado reportado: $($r.Wear)."
                }
            }

            if ($r.TemperatureC -gt 60) {
                Add-Finding ([ref]$score) ([ref]$findings) 10 "Temperatura Disco" "Alto" "Temperatura de disco alta: $($r.TemperatureC) C."
            } elseif ($r.TemperatureC -gt 50) {
                Add-Finding ([ref]$score) ([ref]$findings) 5 "Temperatura Disco" "Medio" "Temperatura de disco moderada: $($r.TemperatureC) C."
            }
        }

        foreach ($f in $DiskAdvanced.Fragmentation) {
            if ($f.FragmentationPercent -ne $null) {
                if (($f.MediaType -notmatch "SSD|Unspecified") -and ($f.FragmentationPercent -gt 20)) {
                    Add-Finding ([ref]$score) ([ref]$findings) 6 "Fragmentacion" "Medio" "Fragmentacion alta en $($f.Drive): $($f.FragmentationPercent)%."
                }
            }
        }

        foreach ($c in $DiskAdvanced.Chkdsk) {
            if ($c.Status -eq "Problemas detectados") {
                Add-Finding ([ref]$score) ([ref]$findings) 15 "CHKDSK" "Critico" "CHKDSK reporta problemas en $($c.Drive)."
            }
        }

        foreach ($t in $DiskAdvanced.Trim) {
            if ($t.TrimEnabled -eq $false) {
                Add-Finding ([ref]$score) ([ref]$findings) 5 "TRIM" "Medio" "TRIM parece desactivado para $($t.FileSystem)."
            }
        }
    }

    if ($Startup.Count -gt 35) {
        Add-Finding ([ref]$score) ([ref]$findings) 10 "Inicio" "Alto" "Exceso de programas de inicio: $($Startup.Count)."
    } elseif ($Startup.Count -gt 20) {
        Add-Finding ([ref]$score) ([ref]$findings) 6 "Inicio" "Medio" "Muchos programas de inicio: $($Startup.Count)."
    }

    if ($Events.TotalEventos -gt 20) {
        Add-Finding ([ref]$score) ([ref]$findings) 15 "Eventos" "Critico" "Muchos eventos relevantes: $($Events.TotalEventos)."
    } elseif ($Events.TotalEventos -gt 10) {
        Add-Finding ([ref]$score) ([ref]$findings) 10 "Eventos" "Alto" "Eventos elevados: $($Events.TotalEventos)."
    } elseif ($Events.TotalEventos -gt 3) {
        Add-Finding ([ref]$score) ([ref]$findings) 5 "Eventos" "Medio" "Algunos eventos relevantes: $($Events.TotalEventos)."
    }

    try {
        if ($Security.Defender.RealTimeProtection -eq $false) {
            Add-Finding ([ref]$score) ([ref]$findings) 5 "Seguridad" "Medio" "Defender sin proteccion en tiempo real."
        }
    } catch {}

    if ($score -lt 0) { $score = 0 }

    $estado = "Critico"
    if ($score -ge 90) { $estado = "Excelente" }
    elseif ($score -ge 75) { $estado = "Bueno" }
    elseif ($score -ge 60) { $estado = "Regular" }
    elseif ($score -ge 40) { $estado = "Malo" }

    return New-Object PSObject -Property @{
        Score = $score
        Estado = $estado
        Hallazgos = $findings
    }
}

function Get-Recommendations {
    param($Reposo,$Benchmark,$Disks,$Smart,$Startup,$Events)
    $recs = @()

    if ($Reposo.CPUAveragePct -gt 30) { $recs += "Revisar procesos residentes: CPU en reposo elevada ($($Reposo.CPUAveragePct)%)." }
    if ($Reposo.RAMAveragePct -gt 80) { $recs += "Revisar programas de inicio o ampliar RAM: uso en reposo $($Reposo.RAMAveragePct)%." }

    if ($Benchmark.Phase.TempMaxC -gt 85) {
        $recs += "Revisar sistema termico: limpieza interna, ventilador, pasta termica y obstrucciones. Temperatura maxima en carga: $($Benchmark.Phase.TempMaxC) C."
    } elseif ($Benchmark.Phase.TempMaxC -gt 75) {
        $recs += "Temperatura en carga moderada. Recomendada limpieza fisica preventiva."
    }

    foreach ($v in $Disks.Volumes) {
        if ($v.LibrePct -lt 15) { $recs += "Liberar espacio en $($v.Unidad): libre $($v.LibrePct)%." }
    }

    foreach ($d in $Disks.PhysicalDisk) {
        if (($d.PhysicalMediaType -eq "HDD") -or ($d.MediaType -like "*Hard*")) { $recs += "Evaluar migracion a SSD: disco mecanico detectado." }
        if (($d.HealthStatus -ne $null) -and ($d.HealthStatus -ne "Healthy")) { $recs += "Respaldar informacion y revisar disco: Health=$($d.HealthStatus)." }
    }

    foreach ($s in $Smart) {
        if (($s.Available -eq $true) -and ($s.Health -ne $null) -and ($s.Health -notmatch "PASSED|OK")) {
            $recs += "Respaldar y reemplazar/revisar disco: SMART $($s.Device) = $($s.Health)."
        }
    }

    if ($Startup.Count -gt 20) { $recs += "Reducir programas de inicio: detectados $($Startup.Count)." }
    if ($Events.TotalEventos -gt 10) { $recs += "Revisar visor de eventos: $($Events.TotalEventos) eventos relevantes recientes." }

    if ($recs.Count -eq 0) { $recs += "No se detectan problemas graves automaticos. Mantener limpieza, actualizaciones y respaldo periodico." }
    return $recs
}

function Export-Tables {
    param($Folder, $Report)

    try { $Report.Reposo.Samples | Select-Object Fase,Fecha,CPUPercent,RAMUsedPercent,RAMFreeGB,DiskQueue,DiskTimePercent,TempMaxC | Export-Csv (Join-Path $Folder "muestras_reposo.csv") -NoTypeInformation -Encoding UTF8 } catch {}
    try { $Report.Benchmark.Phase.Samples | Select-Object Fase,Fecha,CPUPercent,RAMUsedPercent,RAMFreeGB,DiskQueue,DiskTimePercent,TempMaxC | Export-Csv (Join-Path $Folder "muestras_carga.csv") -NoTypeInformation -Encoding UTF8 } catch {}
    try { $Report.Processes.TopCPU | Export-Csv (Join-Path $Folder "procesos_top_cpu.csv") -NoTypeInformation -Encoding UTF8 } catch {}
    try { $Report.Processes.TopRAM | Export-Csv (Join-Path $Folder "procesos_top_ram.csv") -NoTypeInformation -Encoding UTF8 } catch {}
    try { $Report.Disks.Volumes | Export-Csv (Join-Path $Folder "volumenes.csv") -NoTypeInformation -Encoding UTF8 } catch {}
    try { $Report.Disks.PhysicalDisk | Export-Csv (Join-Path $Folder "discos_fisicos.csv") -NoTypeInformation -Encoding UTF8 } catch {}
    try { $Report.Disks.DiskPerf | Export-Csv (Join-Path $Folder "disco_perf.csv") -NoTypeInformation -Encoding UTF8 } catch {}
    try { $Report.DiskAdvanced.StorageReliability | Export-Csv (Join-Path $Folder "disco_storage_reliability.csv") -NoTypeInformation -Encoding UTF8 } catch {}
    try { $Report.DiskAdvanced.WmiSmart | Export-Csv (Join-Path $Folder "disco_wmi_smart.csv") -NoTypeInformation -Encoding UTF8 } catch {}
    try { $Report.DiskAdvanced.Trim | Export-Csv (Join-Path $Folder "disco_trim.csv") -NoTypeInformation -Encoding UTF8 } catch {}
    try { $Report.DiskAdvanced.Fragmentation | Export-Csv (Join-Path $Folder "disco_fragmentacion.csv") -NoTypeInformation -Encoding UTF8 } catch {}
    try { $Report.DiskAdvanced.Chkdsk | Export-Csv (Join-Path $Folder "disco_chkdsk_scan.csv") -NoTypeInformation -Encoding UTF8 } catch {}
    try { $Report.DiskAdvanced.BitLocker | Export-Csv (Join-Path $Folder "disco_bitlocker.csv") -NoTypeInformation -Encoding UTF8 } catch {}
    try { $Report.DiskAdvanced.SmartCtlDiskTemperature | Export-Csv (Join-Path $Folder "disco_temperatura_smartctl.csv") -NoTypeInformation -Encoding UTF8 } catch {}
    try { $Report.TemperaturesFinal | Export-Csv (Join-Path $Folder "temperaturas_finales.csv") -NoTypeInformation -Encoding UTF8 } catch {}
    try { $Report.SmartCtl | Select-Object Device,Available,Health,TemperatureC | Export-Csv (Join-Path $Folder "smartctl_resumen.csv") -NoTypeInformation -Encoding UTF8 } catch {}
    try { $Report.StartupPrograms | Export-Csv (Join-Path $Folder "programas_inicio.csv") -NoTypeInformation -Encoding UTF8 } catch {}
    try { $Report.Events.Eventos | Export-Csv (Join-Path $Folder "eventos.csv") -NoTypeInformation -Encoding UTF8 } catch {}
    try { $Report.Events.Eventos | Group-Object Provider,Nivel | ForEach-Object { New-Object PSObject -Property @{ Grupo = $_.Name; Cantidad = $_.Count } } | Export-Csv (Join-Path $Folder "eventos_resumen.csv") -NoTypeInformation -Encoding UTF8 } catch {}
    try { $Report.Services.ListaServicios | Export-Csv (Join-Path $Folder "servicios.csv") -NoTypeInformation -Encoding UTF8 } catch {}
}

function New-SummaryText {
    param($Report)
    $lines = @()
    $lines += "REPORTE DE RENDIMIENTO V3.6 - SERVICIO TECNICO"
    $lines += "================================================"
    $lines += "Modo: $($Report.Metadata.Modo)"
    $lines += "Cliente: $($Report.Metadata.Cliente)"
    $lines += "OT: $($Report.Metadata.OrdenTrabajo)"
    $lines += "Tecnico: $($Report.Metadata.Tecnico)"
    $lines += "Fecha: $($Report.Metadata.Fecha)"
    $lines += "Equipo: $($Report.System.Fabricante) $($Report.System.Modelo)"
    $lines += "Hostname: $($Report.System.Hostname)"
    $lines += "Serial: $($Report.System.Serial)"
    $lines += "Windows: $($Report.System.Windows) $($Report.System.Version) Build $($Report.System.Build)"
    $lines += "CPU: $($Report.System.CPU)"
    $lines += "RAM: $($Report.System.RAMTotalGB) GB"
    $lines += ""
    $lines += "SCORE: $($Report.Score.Score)/100 - $($Report.Score.Estado)"
    $lines += ""
    $lines += "REPOSO:"
    $lines += "- Duracion: $($Report.Reposo.DuracionSegundos) s"
    $lines += "- CPU promedio: $($Report.Reposo.CPUAveragePct)%"
    $lines += "- CPU maximo: $($Report.Reposo.CPUMaxPct)%"
    $lines += "- RAM promedio: $($Report.Reposo.RAMAveragePct)%"
    $lines += "- Temp promedio: $($Report.Reposo.TempAverageC) C"
    $lines += "- Temp maxima: $($Report.Reposo.TempMaxC) C"
    $lines += ""
    $lines += "CARGA:"
    $lines += "- Duracion: $($Report.Benchmark.Phase.DuracionSegundos) s"
    $lines += "- CPU workers: $($Report.Benchmark.CpuWorkers)"
    $lines += "- CPU promedio: $($Report.Benchmark.Phase.CPUAveragePct)%"
    $lines += "- CPU maximo: $($Report.Benchmark.Phase.CPUMaxPct)%"
    $lines += "- RAM promedio: $($Report.Benchmark.Phase.RAMAveragePct)%"
    $lines += "- Temp promedio: $($Report.Benchmark.Phase.TempAverageC) C"
    $lines += "- Temp maxima: $($Report.Benchmark.Phase.TempMaxC) C"
    $lines += "- Disco escritura promedio: $($Report.Benchmark.DiskWriteMBs) MB/s"
    $lines += "- Disco lectura promedio: $($Report.Benchmark.DiskReadMBs) MB/s"
    $lines += ""
    $lines += "DISCO AVANZADO:"
    $lines += "- Temperatura maxima disco: $($Report.DiskAdvanced.MaxDiskTemperatureC) C"
    foreach ($r in $Report.DiskAdvanced.StorageReliability) {
        $lines += "- Reliability Device=$($r.DeviceId) Temp=$($r.TemperatureC)C Wear=$($r.Wear) PowerOnHours=$($r.PowerOnHours) ReadErr=$($r.ReadErrorsTotal) WriteErr=$($r.WriteErrorsTotal)"
    }
    foreach ($f in $Report.DiskAdvanced.Fragmentation) {
        $lines += "- Fragmentacion $($f.Drive): $($f.FragmentationPercent)% Media=$($f.MediaType) NeedsDefrag=$($f.NeedsDefrag)"
    }
    $lines += ""
    $lines += "EVENTOS RELEVANTES:"
    $lines += "- Total: $($Report.Events.TotalEventos) en $($Report.Events.DiasAnalizados) dias"
    foreach ($e in ($Report.Events.Eventos | Select-Object -First 10)) {
        $lines += "- $($e.Fecha) [$($e.Nivel)] $($e.Provider) ID=$($e.Id): $($e.Mensaje)"
    }
    $lines += ""
    $lines += "Hallazgos:"
    foreach ($h in $Report.Score.Hallazgos) { $lines += "- [$($h.Severidad)] $($h.Area): $($h.Mensaje)" }
    $lines += ""
    $lines += "Recomendaciones:"
    foreach ($r in $Report.Recommendations) { $lines += "- $r" }
    return $lines -join "`r`n"
}

function New-HtmlReport {
    param($Report)

    $scoreClass = "bad"
    if ($Report.Score.Score -ge 90) { $scoreClass = "good" }
    elseif ($Report.Score.Score -ge 75) { $scoreClass = "ok" }
    elseif ($Report.Score.Score -ge 60) { $scoreClass = "warn" }


    $tempReposoMaxText = "N/D"
    $tempReposoAvgText = "N/D"
    $tempCargaMaxText = "N/D"
    $tempCargaAvgText = "N/D"

    if ($Report.Reposo.TempMaxC -ne $null) { $tempReposoMaxText = [string]$Report.Reposo.TempMaxC }
    if ($Report.Reposo.TempAverageC -ne $null) { $tempReposoAvgText = [string]$Report.Reposo.TempAverageC }
    if ($Report.Benchmark.Phase.TempMaxC -ne $null) { $tempCargaMaxText = [string]$Report.Benchmark.Phase.TempMaxC }
    if ($Report.Benchmark.Phase.TempAverageC -ne $null) { $tempCargaAvgText = [string]$Report.Benchmark.Phase.TempAverageC }

    $html = ""
    $html += "<!DOCTYPE html><html lang='es'><head><meta charset='UTF-8'>"
    $html += "<title>Reporte Rendimiento V3.6 - $(Html-Encode $Report.System.Hostname)</title>"
    $html += "<style>"
    $html += "body{font-family:Segoe UI,Arial,sans-serif;background:#f4f6f8;color:#222;margin:0;padding:24px}"
    $html += ".container{max-width:1200px;margin:auto}.header{background:#111827;color:white;padding:24px;border-radius:14px;margin-bottom:20px}"
    $html += ".grid{display:grid;grid-template-columns:repeat(4,1fr);gap:14px}.card{background:white;padding:18px;border-radius:14px;box-shadow:0 2px 12px rgba(0,0,0,.08);margin-bottom:16px}"
    $html += ".metric{font-size:28px;font-weight:700}.label{color:#6b7280;font-size:13px}.good{color:#047857}.ok{color:#2563eb}.warn{color:#b45309}.bad{color:#b91c1c}"
    $html += "table{width:100%;border-collapse:collapse;margin-top:8px}th{background:#e5e7eb;text-align:left;padding:8px;font-size:13px}td{border-bottom:1px solid #e5e7eb;padding:8px;font-size:13px;vertical-align:top}"
    $html += ".badge{display:inline-block;padding:4px 8px;border-radius:999px;background:#e5e7eb;font-size:12px}.finding-Critico{color:#991b1b;font-weight:700}.finding-Alto{color:#b45309;font-weight:700}"
    $html += "@media print{body{background:white}.card{box-shadow:none;border:1px solid #ddd}}"
    $html += "</style></head><body><div class='container'>"

    $html += "<div class='header'><h1>Reporte de Rendimiento V3.6 - Servicio Tecnico</h1>"
    $html += "<p>Modo: <strong>$(Html-Encode $Report.Metadata.Modo)</strong> | Cliente: <strong>$(Html-Encode $Report.Metadata.Cliente)</strong> | OT: <strong>$(Html-Encode $Report.Metadata.OrdenTrabajo)</strong></p>"
    $html += "<p>Tecnico: $(Html-Encode $Report.Metadata.Tecnico) | Fecha: $($Report.Metadata.Fecha)</p></div>"

    $html += "<div class='grid'>"
    $html += "<div class='card'><div class='label'>Score</div><div class='metric $scoreClass'>$($Report.Score.Score)/100</div><div class='badge'>$($Report.Score.Estado)</div></div>"
    $html += "<div class='card'><div class='label'>Temp max reposo</div><div class='metric'>$tempReposoMaxText C</div><div class='label'>Prom: $tempReposoAvgText C</div></div>"
    $html += "<div class='card'><div class='label'>Temp max carga</div><div class='metric'>$tempCargaMaxText C</div><div class='label'>Prom: $tempCargaAvgText C</div></div>"
    $html += "<div class='card'><div class='label'>Duracion prueba</div><div class='metric'>$($Report.Metadata.DuracionTotalSeg)s</div><div class='label'>Reposo + carga</div></div>"
    $html += "</div>"

    $html += "<div class='card'><h2>Resumen reposo vs carga</h2><table>"
    $html += "<tr><th>Metrica</th><th>Reposo</th><th>Carga</th></tr>"
    $html += "<tr><td>Duracion</td><td>$($Report.Reposo.DuracionSegundos) s</td><td>$($Report.Benchmark.Phase.DuracionSegundos) s</td></tr>"
    $html += "<tr><td>CPU promedio</td><td>$($Report.Reposo.CPUAveragePct)%</td><td>$($Report.Benchmark.Phase.CPUAveragePct)%</td></tr>"
    $html += "<tr><td>CPU maximo</td><td>$($Report.Reposo.CPUMaxPct)%</td><td>$($Report.Benchmark.Phase.CPUMaxPct)%</td></tr>"
    $html += "<tr><td>RAM promedio</td><td>$($Report.Reposo.RAMAveragePct)%</td><td>$($Report.Benchmark.Phase.RAMAveragePct)%</td></tr>"
    $html += "<tr><td>Temperatura promedio</td><td>$($Report.Reposo.TempAverageC) C</td><td>$($Report.Benchmark.Phase.TempAverageC) C</td></tr>"
    $html += "<tr><td>Temperatura maxima</td><td>$($Report.Reposo.TempMaxC) C</td><td>$($Report.Benchmark.Phase.TempMaxC) C</td></tr>"
    $html += "<tr><td>Disco escritura</td><td>-</td><td>$($Report.Benchmark.DiskWriteMBs) MB/s</td></tr>"
    $html += "<tr><td>Disco lectura</td><td>-</td><td>$($Report.Benchmark.DiskReadMBs) MB/s</td></tr>"
    $html += "</table></div>"

    $html += "<div class='card'><h2>Ficha del equipo</h2><table>"
    $html += "<tr><th>Campo</th><th>Valor</th></tr>"
    $html += "<tr><td>Hostname</td><td>$(Html-Encode $Report.System.Hostname)</td></tr>"
    $html += "<tr><td>Usuario</td><td>$(Html-Encode $Report.System.Usuario)</td></tr>"
    $html += "<tr><td>Administrador</td><td>$($Report.System.EsAdministrador)</td></tr>"
    $html += "<tr><td>Fabricante / Modelo</td><td>$(Html-Encode $Report.System.Fabricante) $(Html-Encode $Report.System.Modelo)</td></tr>"
    $html += "<tr><td>Serial</td><td>$(Html-Encode $Report.System.Serial)</td></tr>"
    $html += "<tr><td>Windows</td><td>$(Html-Encode $Report.System.Windows) $($Report.System.Version) Build $($Report.System.Build)</td></tr>"
    $html += "<tr><td>CPU</td><td>$(Html-Encode $Report.System.CPU) / Cores $($Report.System.CPUCores) / Hilos $($Report.System.CPUThreads)</td></tr>"
    $html += "<tr><td>RAM</td><td>$($Report.System.RAMTotalGB) GB</td></tr>"
    $html += "</table></div>"

    $html += "<div class='card'><h2>Sensores temperatura finales</h2><table><tr><th>Sensor</th><th>Tipo</th><th>Temperatura C</th><th>Fuente</th></tr>"
    foreach ($t in $Report.TemperaturesFinal) {
        $html += "<tr><td>$(Html-Encode $t.Sensor)</td><td>$(Html-Encode $t.Tipo)</td><td>$($t.TemperaturaC)</td><td>$(Html-Encode $t.Fuente)</td></tr>"
    }
    $html += "</table></div>"

    $html += "<div class='card'><h2>Discos y volumenes</h2><table><tr><th>Unidad</th><th>Nombre</th><th>FS</th><th>Tamano GB</th><th>Libre GB</th><th>Libre %</th><th>Usado %</th></tr>"
    foreach ($v in $Report.Disks.Volumes) {
        $html += "<tr><td>$($v.Unidad)</td><td>$(Html-Encode $v.Nombre)</td><td>$($v.Sistema)</td><td>$($v.TamanoGB)</td><td>$($v.LibreGB)</td><td>$($v.LibrePct)</td><td>$($v.UsadoPct)</td></tr>"
    }
    $html += "</table></div>"


    $html += "<div class='card'><h2>Disco avanzado</h2>"
    $html += "<p><strong>Temperatura maxima disco:</strong> $($Report.DiskAdvanced.MaxDiskTemperatureC) C</p>"

    $html += "<h3>Storage Reliability</h3><table><tr><th>Device</th><th>Temp C</th><th>Temp Max C</th><th>Wear</th><th>PowerOnHours</th><th>ReadErr</th><th>WriteErr</th><th>ReadLatencyMax</th><th>WriteLatencyMax</th></tr>"
    foreach ($r in $Report.DiskAdvanced.StorageReliability) {
        $html += "<tr><td>$($r.DeviceId)</td><td>$($r.TemperatureC)</td><td>$($r.TemperatureMaxC)</td><td>$($r.Wear)</td><td>$($r.PowerOnHours)</td><td>$($r.ReadErrorsTotal)</td><td>$($r.WriteErrorsTotal)</td><td>$($r.ReadLatencyMax)</td><td>$($r.WriteLatencyMax)</td></tr>"
    }
    $html += "</table>"

    $html += "<h3>SMART WMI</h3><table><tr><th>Instance</th><th>PredictFailure</th><th>Reason</th></tr>"
    foreach ($w in $Report.DiskAdvanced.WmiSmart) {
        $html += "<tr><td>$(Html-Encode $w.InstanceName)</td><td>$($w.PredictFailure)</td><td>$($w.Reason)</td></tr>"
    }
    $html += "</table>"

    $html += "<h3>TRIM</h3><table><tr><th>FileSystem</th><th>TrimEnabled</th><th>Raw</th></tr>"
    foreach ($t in $Report.DiskAdvanced.Trim) {
        $html += "<tr><td>$($t.FileSystem)</td><td>$($t.TrimEnabled)</td><td>$(Html-Encode $t.Raw)</td></tr>"
    }
    $html += "</table>"

    $html += "<h3>Fragmentacion / Optimizacion</h3><table><tr><th>Drive</th><th>FS</th><th>Media</th><th>Health</th><th>Frag %</th><th>NeedsDefrag</th></tr>"
    foreach ($f in $Report.DiskAdvanced.Fragmentation) {
        $html += "<tr><td>$($f.Drive)</td><td>$($f.FileSystem)</td><td>$($f.MediaType)</td><td>$($f.HealthStatus)</td><td>$($f.FragmentationPercent)</td><td>$($f.NeedsDefrag)</td></tr>"
    }
    $html += "</table>"

    $html += "<h3>CHKDSK scan</h3><table><tr><th>Drive</th><th>Status</th></tr>"
    foreach ($c in $Report.DiskAdvanced.Chkdsk) {
        $html += "<tr><td>$($c.Drive)</td><td>$($c.Status)</td></tr>"
    }
    $html += "</table>"

    $html += "<h3>BitLocker</h3><table><tr><th>MountPoint</th><th>Status</th><th>Protection</th><th>Method</th><th>%</th></tr>"
    foreach ($b in $Report.DiskAdvanced.BitLocker) {
        $html += "<tr><td>$($b.MountPoint)</td><td>$($b.VolumeStatus)</td><td>$($b.ProtectionStatus)</td><td>$($b.EncryptionMethod)</td><td>$($b.EncryptionPercentage)</td></tr>"
    }
    $html += "</table></div>"

    $html += "<div class='card'><h2>Hallazgos</h2><table><tr><th>Severidad</th><th>Area</th><th>Penalidad</th><th>Mensaje</th></tr>"
    foreach ($h in $Report.Score.Hallazgos) {
        $html += "<tr><td class='finding-$($h.Severidad)'>$($h.Severidad)</td><td>$($h.Area)</td><td>-$($h.Penalidad)</td><td>$(Html-Encode $h.Mensaje)</td></tr>"
    }
    if ($Report.Score.Hallazgos.Count -eq 0) { $html += "<tr><td colspan='4'>Sin hallazgos negativos relevantes.</td></tr>" }
    $html += "</table></div>"

    $html += "<div class='card'><h2>Recomendaciones</h2><ul>"
    foreach ($r in $Report.Recommendations) { $html += "<li>$(Html-Encode $r)</li>" }
    $html += "</ul></div>"

    $html += "<div class='card'><h2>Top procesos por CPU</h2><table><tr><th>Proceso</th><th>PID</th><th>CPU s</th><th>RAM MB</th><th>Ruta</th></tr>"
    foreach ($p in $Report.Processes.TopCPU) {
        $html += "<tr><td>$(Html-Encode $p.Name)</td><td>$($p.Id)</td><td>$($p.CPUSeconds)</td><td>$($p.RAMMB)</td><td>$(Html-Encode $p.Path)</td></tr>"
    }
    $html += "</table></div>"

    $html += "<div class='card'><h2>Eventos relevantes detectados</h2>"
    $html += "<p>Total eventos relevantes: <strong>$($Report.Events.TotalEventos)</strong> en los ultimos $($Report.Events.DiasAnalizados) dias.</p>"
    $html += "<h3>Resumen por origen y nivel</h3><table><tr><th>Provider / Nivel</th><th>Cantidad</th></tr>"
    try {
        $groups = @($Report.Events.Eventos | Group-Object Provider,Nivel | Sort-Object Count -Descending)
        foreach ($g in $groups) {
            $html += "<tr><td>$(Html-Encode $g.Name)</td><td>$($g.Count)</td></tr>"
        }
    } catch {}
    $html += "</table>"
    $html += "<h3>Detalle ultimos eventos</h3><table><tr><th>Fecha</th><th>Nivel</th><th>Provider</th><th>ID</th><th>Mensaje</th></tr>"
    foreach ($e in ($Report.Events.Eventos | Select-Object -First 80)) {
        $html += "<tr><td>$($e.Fecha)</td><td>$($e.Nivel)</td><td>$(Html-Encode $e.Provider)</td><td>$($e.Id)</td><td>$(Html-Encode $e.Mensaje)</td></tr>"
    }
    if ($Report.Events.Eventos.Count -eq 0) {
        $html += "<tr><td colspan='5'>Sin eventos relevantes detectados.</td></tr>"
    }
    $html += "</table></div>"

    $html += "<div class='card'><h2>Muestras de temperatura y carga</h2><table><tr><th>Fase</th><th>Fecha</th><th>CPU %</th><th>RAM %</th><th>Temp Max C</th><th>Disk Time %</th></tr>"
    $allSamples = @()
    $allSamples += $Report.Reposo.Samples
    $allSamples += $Report.Benchmark.Phase.Samples
    foreach ($s in $allSamples) {
        $html += "<tr><td>$($s.Fase)</td><td>$($s.Fecha)</td><td>$($s.CPUPercent)</td><td>$($s.RAMUsedPercent)</td><td>$($s.TempMaxC)</td><td>$($s.DiskTimePercent)</td></tr>"
    }
    $html += "</table></div>"

    $html += "</div></body></html>"
    return $html
}

Write-Host "================================================" -ForegroundColor Cyan
Write-Host "REPORTE DE RENDIMIENTO V3.6 - SERVICIO TECNICO" -ForegroundColor Cyan
Write-Host "================================================" -ForegroundColor Cyan

if ($DuracionReposoSeg -lt 30) {
    Write-Warn2 "DuracionReposoSeg muy baja. Se ajusta a 30 segundos."
    $DuracionReposoSeg = 30
}

if ($DuracionPruebaSeg -lt 60) {
    Write-Warn2 "DuracionPruebaSeg muy baja. Se ajusta a 60 segundos."
    $DuracionPruebaSeg = 60
}

if ($IntervaloMuestreoSeg -lt 2) {
    Write-Warn2 "IntervaloMuestreoSeg muy bajo. Se ajusta a 2 segundos."
    $IntervaloMuestreoSeg = 2
}


$paths = Initialize-DefaultPaths -InputRutaSalida $RutaSalida -InputLibreHardwareMonitorPath $LibreHardwareMonitorPath
$RutaSalida = $paths.RutaSalida
$LibreHardwareMonitorPath = $paths.LibreHardwareMonitorPath

Write-Info "Ruta base local: $($paths.BaseDir)"
Write-Info "Ruta de reportes: $RutaSalida"
Write-Info "Ruta LibreHardwareMonitor DLL: $LibreHardwareMonitorPath"

if (!(Test-Path $RutaSalida)) {
    New-Item -Path $RutaSalida -ItemType Directory -Force | Out-Null
}

$runFolder = New-RunFolder -Root $RutaSalida -Cliente $Cliente -OrdenTrabajo $OrdenTrabajo -Modo $Modo
Write-Info "Carpeta de reporte: $runFolder"

$advancedTemp = $false
if ($UseLibreHardwareMonitor.IsPresent) { $advancedTemp = $true }
if ($InstalarLibreHardwareMonitor.IsPresent) { $advancedTemp = $true }

if ($advancedTemp -eq $true) {
    $resolvedBeforeInstall = Resolve-LhmDllPath -RequestedPath $LibreHardwareMonitorPath
    if (!(Test-Path $resolvedBeforeInstall)) {
        if (($UseLibreHardwareMonitor.IsPresent) -or ($InstalarLibreHardwareMonitor.IsPresent)) {
            $LibreHardwareMonitorPath = Ensure-LibreHardwareMonitor -RequestedDllPath $LibreHardwareMonitorPath
        }
    } else {
        $LibreHardwareMonitorPath = $resolvedBeforeInstall
    }
}

$system = Get-SystemInfo
$reposo = Measure-Phase -Phase "Reposo" -Seconds $DuracionReposoSeg -Interval $IntervaloMuestreoSeg -AdvancedTemp $advancedTemp -LhmDll $LibreHardwareMonitorPath
$benchmark = Invoke-SeriousBenchmark -Folder $runFolder -Duration $DuracionPruebaSeg -Interval $IntervaloMuestreoSeg -AdvancedTemp $advancedTemp -LhmDll $LibreHardwareMonitorPath -MaxWorkers $MaxCpuWorkers -DiskFileMB $DiskTestFileMB

Write-Info "Tomando muestra termica final"
$tempsFinal = @(Get-TemperatureInfo -Advanced $advancedTemp -DllPath $LibreHardwareMonitorPath)

$processes = Get-ProcessesInfo
$disks = Get-DiskInfo
$diskAdvanced = Get-DiskAdvancedInfo -RunFragmentation $AnalizarFragmentacion.IsPresent -RunChkdskScan $IncluirChkdskScan.IsPresent -UseSmartCtl $UseSmartCtl.IsPresent -SmartCtlPath $SmartCtlPath

if ($UseSmartCtl.IsPresent) {
    $smart = @(Get-SmartCtlInfo -SmartCtlPath $SmartCtlPath)
} else {
    $smart = @()
    $smart += New-Object PSObject -Property @{
        Device = "smartctl"
        Available = $false
        Health = $null
        TemperatureC = $null
        Raw = "No solicitado"
    }
}

$startup = Get-StartupPrograms
$services = Get-ServicesInfo
$events = Get-CriticalEvents -Days 7
$security = Get-SecurityInfo

$score = Get-Score -Reposo $reposo -Benchmark $benchmark -Disks $disks -DiskAdvanced $diskAdvanced -TempsFinal $tempsFinal -Smart $smart -Startup $startup -Events $events -Security $security
$recommendations = Get-Recommendations -Reposo $reposo -Benchmark $benchmark -Disks $disks -Smart $smart -Startup $startup -Events $events

$metadata = New-Object PSObject -Property @{
    Modo = $Modo
    Cliente = $Cliente
    OrdenTrabajo = $OrdenTrabajo
    Tecnico = $Tecnico
    Fecha = Get-Date
    RutaReporte = $runFolder
    ScriptVersion = "3.6-thermal-events-disk-integrations-ps5-safe"
    DuracionReposoSeg = $DuracionReposoSeg
    DuracionPruebaSeg = $DuracionPruebaSeg
    DuracionTotalSeg = ($DuracionReposoSeg + $DuracionPruebaSeg)
    IntervaloMuestreoSeg = $IntervaloMuestreoSeg
    UseLibreHardwareMonitor = $UseLibreHardwareMonitor.IsPresent
    InstalarLibreHardwareMonitor = $InstalarLibreHardwareMonitor.IsPresent
    LibreHardwareMonitorPath = $LibreHardwareMonitorPath
    UseSmartCtl = $UseSmartCtl.IsPresent
    SmartCtlPath = $SmartCtlPath
}

$report = New-Object PSObject -Property @{
    Metadata = $metadata
    System = $system
    Reposo = $reposo
    Benchmark = $benchmark
    Processes = $processes
    Disks = $disks
    DiskAdvanced = $diskAdvanced
    TemperaturesFinal = $tempsFinal
    SmartCtl = $smart
    StartupPrograms = $startup
    Services = $services
    Events = $events
    Security = $security
    Score = $score
    Recommendations = $recommendations
}

$jsonPath = Join-Path $runFolder "reporte.json"
$htmlPath = Join-Path $runFolder "reporte.html"
$txtPath = Join-Path $runFolder "resumen.txt"
$readmePath = Join-Path $runFolder "leer_primero.txt"

$report | ConvertTo-Json -Depth 14 | Out-File -FilePath $jsonPath -Encoding UTF8
Export-Tables -Folder $runFolder -Report $report
New-HtmlReport -Report $report | Out-File -FilePath $htmlPath -Encoding UTF8
New-SummaryText -Report $report | Out-File -FilePath $txtPath -Encoding UTF8

$readmeLines = @()
$readmeLines += "Reporte de rendimiento generado localmente en este equipo."
$readmeLines += "No modifica configuraciones del sistema."
$readmeLines += "Crea carpetas locales para reportes y herramientas auxiliares."
$readmeLines += "Ruta del reporte: $runFolder"
$readmeLines += "HTML: $htmlPath"
$readmeLines += "JSON: $jsonPath"
$readmeLines += "TXT: $txtPath"
$readmeLines += "Si la temperatura aparece N/D, el equipo no expuso sensores por los metodos disponibles o falta ejecutar como administrador."
$readmeLines += "Archivos de disco avanzado: disco_storage_reliability.csv, disco_wmi_smart.csv, disco_trim.csv, disco_fragmentacion.csv, disco_chkdsk_scan.csv, disco_bitlocker.csv."
$readmeLines += "Eventos: eventos.csv contiene detalle; eventos_resumen.csv agrupa por Provider/Nivel."
$readmeLines += "La fragmentacion detallada solo se ejecuta con -AnalizarFragmentacion. CHKDSK scan solo se ejecuta con -IncluirChkdskScan."
$readmeLines | Out-File -FilePath $readmePath -Encoding UTF8

Write-Host ""
Write-Host "================================================" -ForegroundColor Green
Write-Host "REPORTE GENERADO" -ForegroundColor Green
Write-Host "================================================" -ForegroundColor Green
Write-Host "REPORTE_GENERADO=$runFolder" -ForegroundColor Green
Write-Host "SCORE=$($score.Score)"
Write-Host "ESTADO=$($score.Estado)"
Write-Host "TEMP_REPOSO_MAX=$($reposo.TempMaxC)"
Write-Host "TEMP_CARGA_MAX=$($benchmark.Phase.TempMaxC)"
Write-Host "HTML=$htmlPath"
Write-Host "JSON=$jsonPath"
Write-Host "TXT=$txtPath"
Write-Host ""

if ($AbrirReporte.IsPresent) {
    Start-Process $htmlPath
}
