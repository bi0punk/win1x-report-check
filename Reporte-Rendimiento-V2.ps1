param(
    [Parameter(Mandatory=$true)]
    [ValidateSet("Antes", "Despues")]
    [string]$Modo,

    [string]$Cliente = "SinCliente",
    [string]$OrdenTrabajo = "SinOT",
    [string]$Tecnico = "ServicioTecnico",

    [int]$DuracionMuestreo = 30,

    [string]$RutaSalida = "C:\SYSTEC\Reportes-Rendimiento",

    [switch]$Benchmark,

    [switch]$UseLibreHardwareMonitor,

    [string]$LibreHardwareMonitorPath = "C:\SYSTEC\Tools\LibreHardwareMonitor\LibreHardwareMonitorLib.dll",

    [switch]$UseSmartCtl,

    [string]$SmartCtlPath = "smartctl.exe"
)

$ErrorActionPreference = "Continue"

function Write-Info {
    param([string]$m)
    Write-Host "[INFO] $m" -ForegroundColor Cyan
}

function Write-Warn {
    param([string]$m)
    Write-Host "[WARN] $m" -ForegroundColor Yellow
}

function Write-Err {
    param([string]$m)
    Write-Host "[ERROR] $m" -ForegroundColor Red
}

function ConvertTo-SafeName {
    param([string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return "SinDato"
    }

    return (($Text -replace '[\\/:*?"<>|]', '_') -replace '\s+', '_').Trim('_')
}

function HtmlEncode {
    param([object]$Value)

    if ($null -eq $Value) {
        return ""
    }

    return [System.Net.WebUtility]::HtmlEncode([string]$Value)
}

function Get-IsAdmin {
    try {
        $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object Security.Principal.WindowsPrincipal($identity)
        return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    } catch {
        return $false
    }
}

function New-ReportFolders {
    param(
        [string]$Root,
        [string]$Cliente,
        [string]$OT,
        [string]$Modo
    )

    $hostSafe = ConvertTo-SafeName $env:COMPUTERNAME
    $clienteSafe = ConvertTo-SafeName $Cliente
    $otSafe = ConvertTo-SafeName $OT

    $base = Join-Path $Root "${clienteSafe}_${otSafe}_${hostSafe}"
    $run = Join-Path $base "$((Get-Date).ToString('yyyyMMdd_HHmmss'))_$($Modo.ToUpper())"

    New-Item -Path $run -ItemType Directory -Force | Out-Null

    return $run
}

function Get-SystemInfoSafe {
    Write-Info "Recolectando ficha del sistema"

    $computer = $null
    $os = $null
    $bios = $null
    $cpu = $null
    $gpu = @()
    $ram = @()

    try { $computer = Get-CimInstance Win32_ComputerSystem } catch {}
    try { $os = Get-CimInstance Win32_OperatingSystem } catch {}
    try { $bios = Get-CimInstance Win32_BIOS } catch {}
    try { $cpu = Get-CimInstance Win32_Processor | Select-Object -First 1 } catch {}
    try { $gpu = @(Get-CimInstance Win32_VideoController) } catch {}
    try { $ram = @(Get-CimInstance Win32_PhysicalMemory) } catch {}

    [PSCustomObject]@{
        Hostname        = $env:COMPUTERNAME
        Usuario         = "$env:USERDOMAIN\$env:USERNAME"
        EsAdministrador = Get-IsAdmin
        Fabricante      = $computer.Manufacturer
        Modelo          = $computer.Model
        Serial          = $bios.SerialNumber
        Windows         = $os.Caption
        Version         = $os.Version
        Build           = $os.BuildNumber
        Arquitectura    = $os.OSArchitecture
        UltimoArranque  = $os.LastBootUpTime
        CPU             = $cpu.Name
        CPUCores        = $cpu.NumberOfCores
        CPUThreads      = $cpu.NumberOfLogicalProcessors
        CPUMaxMHz       = $cpu.MaxClockSpeed
        RAMTotalGB      = if ($computer.TotalPhysicalMemory) { [math]::Round($computer.TotalPhysicalMemory / 1GB, 2) } else { $null }
        RAMModulos      = @($ram | ForEach-Object {
            [PSCustomObject]@{
                Banco       = $_.BankLabel
                Marca       = $_.Manufacturer
                Parte       = ($_.PartNumber -as [string]).Trim()
                CapacidadGB = [math]::Round($_.Capacity / 1GB, 2)
                SpeedMHz    = $_.Speed
            }
        })
        GPU             = @($gpu | ForEach-Object {
            [PSCustomObject]@{
                Nombre = $_.Name
                Driver = $_.DriverVersion
                RAMGB  = if ($_.AdapterRAM) { [math]::Round($_.AdapterRAM / 1GB, 2) } else { $null }
            }
        })
    }
}

function Get-CpuMemorySample {
    param([int]$Seconds)

    Write-Info "Muestreando CPU/RAM durante $Seconds segundos"

    $cpuSamples = @()
    $ramSamples = @()

    for ($i = 1; $i -le $Seconds; $i++) {
        try {
            $cpu = Get-CimInstance Win32_PerfFormattedData_PerfOS_Processor -Filter "Name='_Total'" -ErrorAction Stop
            $cpuSamples += [double]$cpu.PercentProcessorTime
        } catch {
            try {
                $counter = Get-Counter '\Processor(_Total)\% Processor Time' -ErrorAction Stop
                $cpuSamples += [math]::Round($counter.CounterSamples[0].CookedValue, 2)
            } catch {}
        }

        try {
            $os = Get-CimInstance Win32_OperatingSystem
            $totalKB = [double]$os.TotalVisibleMemorySize
            $freeKB = [double]$os.FreePhysicalMemory
            $usedPct = (($totalKB - $freeKB) / $totalKB) * 100

            $ramSamples += [PSCustomObject]@{
                UsedPercent = [math]::Round($usedPct, 2)
                FreeGB      = [math]::Round(($freeKB * 1KB) / 1GB, 2)
                TotalGB     = [math]::Round(($totalKB * 1KB) / 1GB, 2)
            }
        } catch {}

        Start-Sleep -Seconds 1
    }

    [PSCustomObject]@{
        DuracionSegundos = $Seconds
        CPUAveragePct    = if ($cpuSamples.Count) { [math]::Round(($cpuSamples | Measure-Object -Average).Average, 2) } else { $null }
        CPUMaxPct        = if ($cpuSamples.Count) { [math]::Round(($cpuSamples | Measure-Object -Maximum).Maximum, 2) } else { $null }
        CPUMinPct        = if ($cpuSamples.Count) { [math]::Round(($cpuSamples | Measure-Object -Minimum).Minimum, 2) } else { $null }
        RAMAveragePct    = if ($ramSamples.Count) { [math]::Round(($ramSamples.UsedPercent | Measure-Object -Average).Average, 2) } else { $null }
        RAMMaxPct        = if ($ramSamples.Count) { [math]::Round(($ramSamples.UsedPercent | Measure-Object -Maximum).Maximum, 2) } else { $null }
        RAMFreeAvgGB     = if ($ramSamples.Count) { [math]::Round(($ramSamples.FreeGB | Measure-Object -Average).Average, 2) } else { $null }
        CPUSamples       = $cpuSamples
        RAMSamples       = $ramSamples
    }
}

function Get-TopProcessesSafe {
    Write-Info "Recolectando procesos"

    try {
        $items = @(Get-Process | ForEach-Object {
            [PSCustomObject]@{
                Name       = $_.ProcessName
                Id         = $_.Id
                CPUSeconds = if ($_.CPU) { [math]::Round($_.CPU, 2) } else { 0 }
                RAMMB      = [math]::Round($_.WorkingSet64 / 1MB, 2)
                Path       = try { $_.Path } catch { $null }
                StartTime  = try { $_.StartTime } catch { $null }
            }
        })

        [PSCustomObject]@{
            TotalProcesos = $items.Count
            TopCPU        = @($items | Sort-Object CPUSeconds -Descending | Select-Object -First 20)
            TopRAM        = @($items | Sort-Object RAMMB -Descending | Select-Object -First 20)
        }
    } catch {
        [PSCustomObject]@{
            TotalProcesos = $null
            TopCPU        = @()
            TopRAM        = @()
        }
    }
}

function Get-DiskInfoSafe {
    Write-Info "Recolectando discos"

    $volumes = @()
    $physical = @()
    $perf = @()

    try {
        $volumes = @(Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3" | ForEach-Object {
            [PSCustomObject]@{
                Unidad   = $_.DeviceID
                Nombre   = $_.VolumeName
                Sistema  = $_.FileSystem
                TamanoGB = if ($_.Size) { [math]::Round($_.Size / 1GB, 2) } else { 0 }
                LibreGB  = if ($_.FreeSpace) { [math]::Round($_.FreeSpace / 1GB, 2) } else { 0 }
                LibrePct = if ($_.Size) { [math]::Round(($_.FreeSpace / $_.Size) * 100, 2) } else { 0 }
                UsadoPct = if ($_.Size) { [math]::Round((($_.Size - $_.FreeSpace) / $_.Size) * 100, 2) } else { 0 }
            }
        })
    } catch {}

    try {
        $physical = @(Get-CimInstance Win32_DiskDrive | ForEach-Object {
            [PSCustomObject]@{
                Modelo    = $_.Model
                Interface = $_.InterfaceType
                Serial    = $_.SerialNumber
                TamanoGB  = if ($_.Size) { [math]::Round($_.Size / 1GB, 2) } else { $null }
                MediaType = $_.MediaType
                Status    = $_.Status
            }
        })
    } catch {}

    try {
        foreach ($p in @(Get-PhysicalDisk)) {
            $physical += [PSCustomObject]@{
                Modelo            = $p.FriendlyName
                Interface         = $p.BusType
                Serial            = $p.SerialNumber
                TamanoGB          = if ($p.Size) { [math]::Round($p.Size / 1GB, 2) } else { $null }
                MediaType         = $p.MediaType
                Status            = $p.OperationalStatus -join ","
                HealthStatus      = $p.HealthStatus
                PhysicalMediaType = $p.MediaType
            }
        }
    } catch {}

    try {
        $perf = @(Get-CimInstance Win32_PerfFormattedData_PerfDisk_LogicalDisk |
            Where-Object { $_.Name -ne "_Total" } |
            ForEach-Object {
                [PSCustomObject]@{
                    Unidad                = $_.Name
                    DiskReadsPerSec       = $_.DiskReadsPerSec
                    DiskWritesPerSec      = $_.DiskWritesPerSec
                    AvgDiskQueueLength    = $_.AvgDiskQueueLength
                    PercentDiskTime       = $_.PercentDiskTime
                    AvgDiskSecPerRead_ms  = [math]::Round($_.AvgDisksecPerRead * 1000, 2)
                    AvgDiskSecPerWrite_ms = [math]::Round($_.AvgDisksecPerWrite * 1000, 2)
                }
            })
    } catch {}

    [PSCustomObject]@{
        Volumes      = $volumes
        PhysicalDisk = $physical
        DiskPerf     = $perf
    }
}

function Get-TemperatureFromLibreHardwareMonitor {
    param([string]$DllPath)

    $results = @()

    if (!(Test-Path $DllPath)) {
        return @([PSCustomObject]@{
            Sensor       = "LibreHardwareMonitor"
            Tipo         = "Info"
            TemperaturaC = $null
            Fuente       = "DLL no encontrada: $DllPath"
        })
    }

    try {
        Add-Type -Path $DllPath -ErrorAction Stop

        $computer = New-Object LibreHardwareMonitor.Hardware.Computer
        $computer.IsCpuEnabled = $true
        $computer.IsGpuEnabled = $true
        $computer.IsStorageEnabled = $true
        $computer.IsMotherboardEnabled = $true
        $computer.IsMemoryEnabled = $true
        $computer.Open()

        foreach ($hw in $computer.Hardware) {
            $hw.Update()

            foreach ($sub in $hw.SubHardware) {
                $sub.Update()
            }

            foreach ($sensor in $hw.Sensors) {
                if ($sensor.SensorType.ToString() -eq "Temperature" -and $null -ne $sensor.Value) {
                    $results += [PSCustomObject]@{
                        Sensor       = "$($hw.Name) - $($sensor.Name)"
                        Tipo         = $hw.HardwareType.ToString()
                        TemperaturaC = [math]::Round([double]$sensor.Value, 1)
                        Fuente       = "LibreHardwareMonitor"
                    }
                }
            }

            foreach ($sub in $hw.SubHardware) {
                foreach ($sensor in $sub.Sensors) {
                    if ($sensor.SensorType.ToString() -eq "Temperature" -and $null -ne $sensor.Value) {
                        $results += [PSCustomObject]@{
                            Sensor       = "$($sub.Name) - $($sensor.Name)"
                            Tipo         = $sub.HardwareType.ToString()
                            TemperaturaC = [math]::Round([double]$sensor.Value, 1)
                            Fuente       = "LibreHardwareMonitor"
                        }
                    }
                }
            }
        }

        $computer.Close()
    } catch {
        $results += [PSCustomObject]@{
            Sensor       = "LibreHardwareMonitor"
            Tipo         = "Error"
            TemperaturaC = $null
            Fuente       = "Error cargando DLL: $($_.Exception.Message)"
        }
    }

    return $results
}

function Get-TemperatureFromWmi {
    $temps = @()

    try {
        $thermal = @(Get-CimInstance -Namespace root/wmi -ClassName MSAcpi_ThermalZoneTemperature -ErrorAction Stop)

        foreach ($t in $thermal) {
            if ($t.CurrentTemperature) {
                $c = [math]::Round(($t.CurrentTemperature / 10) - 273.15, 1)

                if ($c -gt -20 -and $c -lt 130) {
                    $temps += [PSCustomObject]@{
                        Sensor       = $t.InstanceName
                        Tipo         = "ACPI"
                        TemperaturaC = $c
                        Fuente       = "WMI_ACPI"
                    }
                }
            }
        }
    } catch {}

    return $temps
}

function Get-TemperatureInfoSafe {
    param(
        [bool]$Advanced,
        [string]$LhmDll
    )

    Write-Info "Recolectando temperaturas"

    $temps = @()

    if ($Advanced) {
        $temps += @(Get-TemperatureFromLibreHardwareMonitor -DllPath $LhmDll)
    }

    $temps += @(Get-TemperatureFromWmi)

    if ($temps.Count -eq 0) {
        $temps += [PSCustomObject]@{
            Sensor       = "No disponible"
            Tipo         = "No disponible"
            TemperaturaC = $null
            Fuente       = "Sin sensores expuestos por WMI. Use LibreHardwareMonitorLib.dll para modo avanzado."
        }
    }

    return $temps
}

function Get-SmartCtlInfo {
    param([string]$SmartCtlPath)

    Write-Info "Intentando recolectar SMART con smartctl"

    $items = @()

    $cmd = Get-Command $SmartCtlPath -ErrorAction SilentlyContinue

    if (!$cmd -and !(Test-Path $SmartCtlPath)) {
        return @([PSCustomObject]@{
            Device       = "smartctl"
            Available    = $false
            Health       = $null
            TemperatureC = $null
            Raw          = "smartctl no encontrado"
        })
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

                $items += [PSCustomObject]@{
                    Device       = $dev
                    Available    = $true
                    Health       = $health
                    TemperatureC = $temp
                    Raw          = $raw
                }
            }
        }
    } catch {
        $items += [PSCustomObject]@{
            Device       = "smartctl"
            Available    = $false
            Health       = $null
            TemperatureC = $null
            Raw          = $_.Exception.Message
        }
    }

    return $items
}

function Get-StartupProgramsSafe {
    Write-Info "Recolectando programas de inicio"

    try {
        @(Get-CimInstance Win32_StartupCommand | ForEach-Object {
            [PSCustomObject]@{
                Nombre    = $_.Name
                Comando   = $_.Command
                Ubicacion = $_.Location
                Usuario   = $_.User
            }
        })
    } catch {
        @()
    }
}

function Get-ServiceSummarySafe {
    Write-Info "Recolectando servicios"

    try {
        $services = @(Get-Service | ForEach-Object {
            [PSCustomObject]@{
                Nombre      = $_.Name
                DisplayName = $_.DisplayName
                Estado      = $_.Status.ToString()
                TipoInicio  = try { $_.StartType.ToString() } catch { $null }
            }
        })

        [PSCustomObject]@{
            TotalServicios        = $services.Count
            ServiciosRunning      = @($services | Where-Object Estado -eq "Running").Count
            AutomaticosEjecutando = @($services | Where-Object { $_.TipoInicio -eq "Automatic" -and $_.Estado -eq "Running" }).Count
            AutomaticosDetenidos  = @($services | Where-Object { $_.TipoInicio -eq "Automatic" -and $_.Estado -ne "Running" }).Count
            ListaServicios        = $services
        }
    } catch {
        [PSCustomObject]@{
            TotalServicios        = $null
            ServiciosRunning      = $null
            AutomaticosEjecutando = $null
            AutomaticosDetenidos  = $null
            ListaServicios        = @()
        }
    }
}

function Get-CriticalEventsSafe {
    param([int]$Days = 7)

    Write-Info "Recolectando eventos relevantes últimos $Days días"

    $start = (Get-Date).AddDays(-$Days)

    $providers = @(
        "Microsoft-Windows-Kernel-Power",
        "Disk",
        "Ntfs",
        "Microsoft-Windows-WHEA-Logger",
        "Service Control Manager",
        "Microsoft-Windows-WindowsUpdateClient",
        "Display"
    )

    $events = @()

    foreach ($provider in $providers) {
        try {
            $ev = @(Get-WinEvent -FilterHashtable @{
                LogName      = "System"
                ProviderName = $provider
                StartTime    = $start
            } -ErrorAction Stop | Select-Object -First 80)

            foreach ($e in $ev) {
                if ($e.LevelDisplayName -in @("Critical", "Error", "Warning") -or $e.Id -in @(41, 7, 51, 55, 129, 153, 18, 19)) {
                    $msg = (($e.Message -replace "`r|`n", " ") -replace "\s+", " ")

                    if ($msg.Length -gt 600) {
                        $msg = $msg.Substring(0, 600)
                    }

                    $events += [PSCustomObject]@{
                        Fecha    = $e.TimeCreated
                        Nivel    = $e.LevelDisplayName
                        Provider = $e.ProviderName
                        Id       = $e.Id
                        Mensaje  = $msg
                    }
                }
            }
        } catch {}
    }

    [PSCustomObject]@{
        DiasAnalizados = $Days
        TotalEventos   = $events.Count
        Eventos        = @($events | Sort-Object Fecha -Descending)
    }
}

function Get-SecurityBasicSafe {
    Write-Info "Recolectando seguridad básica"

    $def = $null
    $fw = @()
    $bit = @()
    $admins = @()

    try {
        $d = Get-MpComputerStatus -ErrorAction Stop

        $def = [PSCustomObject]@{
            AntivirusEnabled  = $d.AntivirusEnabled
            RealTimeProtection = $d.RealTimeProtectionEnabled
            AMServiceEnabled   = $d.AMServiceEnabled
            QuickScanAge       = $d.QuickScanAge
            FullScanAge        = $d.FullScanAge
        }
    } catch {
        $def = [PSCustomObject]@{
            Nota = "No disponible"
        }
    }

    try {
        $fw = @(Get-NetFirewallProfile | ForEach-Object {
            [PSCustomObject]@{
                Perfil  = $_.Name
                Enabled = $_.Enabled
            }
        })
    } catch {}

    try {
        $bit = @(Get-BitLockerVolume -ErrorAction Stop | ForEach-Object {
            [PSCustomObject]@{
                MountPoint       = $_.MountPoint
                ProtectionStatus = $_.ProtectionStatus
                VolumeStatus     = $_.VolumeStatus
                EncryptionMethod = $_.EncryptionMethod
            }
        })
    } catch {}

    try {
        $admins = @(Get-LocalGroupMember -Group "Administrators" -ErrorAction Stop | ForEach-Object {
            [PSCustomObject]@{
                Nombre = $_.Name
                Tipo   = $_.ObjectClass
            }
        })
    } catch {
        try {
            $admins = @(Get-LocalGroupMember -Group "Administradores" -ErrorAction Stop | ForEach-Object {
                [PSCustomObject]@{
                    Nombre = $_.Name
                    Tipo   = $_.ObjectClass
                }
            })
        } catch {}
    }

    [PSCustomObject]@{
        Defender  = $def
        Firewall  = $fw
        BitLocker = $bit
        Admins    = $admins
    }
}

function Invoke-LightBenchmark {
    param([string]$TempFolder)

    Write-Info "Ejecutando benchmark liviano"

    $benchFile = Join-Path $TempFolder "bench_temp.bin"
    $cpuScore = $null
    $diskWrite = $null
    $diskRead = $null

    try {
        $sw = [Diagnostics.Stopwatch]::StartNew()
        $r = 0.0

        for ($i = 1; $i -le 2500000; $i++) {
            $r += [Math]::Sqrt($i) * [Math]::Sin($i)
        }

        $sw.Stop()

        $cpuScore = [PSCustomObject]@{
            Operaciones = 2500000
            Segundos    = [math]::Round($sw.Elapsed.TotalSeconds, 3)
            Resultado   = [math]::Round($r, 3)
        }
    } catch {}

    try {
        $sizeMB = 128
        $chunkMB = 4

        $buffer = New-Object byte[] ($chunkMB * 1MB)
        (New-Object Random).NextBytes($buffer)

        $swW = [Diagnostics.Stopwatch]::StartNew()
        $fs = [IO.File]::Open($benchFile, [IO.FileMode]::Create, [IO.FileAccess]::Write)

        for ($i = 0; $i -lt ($sizeMB / $chunkMB); $i++) {
            $fs.Write($buffer, 0, $buffer.Length)
        }

        $fs.Flush()
        $fs.Close()
        $swW.Stop()

        $diskWrite = [math]::Round($sizeMB / $swW.Elapsed.TotalSeconds, 2)

        $readBuffer = New-Object byte[] ($chunkMB * 1MB)

        $swR = [Diagnostics.Stopwatch]::StartNew()
        $fsr = [IO.File]::Open($benchFile, [IO.FileMode]::Open, [IO.FileAccess]::Read)

        while ($fsr.Read($readBuffer, 0, $readBuffer.Length) -gt 0) {}

        $fsr.Close()
        $swR.Stop()

        $diskRead = [math]::Round($sizeMB / $swR.Elapsed.TotalSeconds, 2)
    } catch {}
    finally {
        Remove-Item $benchFile -Force -ErrorAction SilentlyContinue
    }

    [PSCustomObject]@{
        CPU          = $cpuScore
        DiskWriteMBs = $diskWrite
        DiskReadMBs  = $diskRead
        FileSizeMB    = 128
    }
}

function Get-PerformanceScore {
    param(
        $Sample,
        $Disks,
        $Temps,
        $Smart,
        $Startup,
        $Events,
        $Security,
        $BenchmarkResult
    )

    $score = 100
    $findings = @()

    function Add-Penalty {
        param(
            [int]$p,
            [string]$area,
            [string]$msg,
            [string]$sev
        )

        $script:score -= $p

        $script:findings += [PSCustomObject]@{
            Area      = $area
            Severidad = $sev
            Penalidad = $p
            Mensaje   = $msg
        }
    }

    if ($Sample.CPUAveragePct -ne $null) {
        if ($Sample.CPUAveragePct -gt 70) {
            Add-Penalty 15 "CPU" "CPU promedio muy alta: $($Sample.CPUAveragePct)%." "Critico"
        } elseif ($Sample.CPUAveragePct -gt 40) {
            Add-Penalty 10 "CPU" "CPU promedio alta: $($Sample.CPUAveragePct)%." "Alto"
        } elseif ($Sample.CPUAveragePct -gt 20) {
            Add-Penalty 5 "CPU" "CPU promedio moderada: $($Sample.CPUAveragePct)%." "Medio"
        }
    } else {
        Add-Penalty 3 "CPU" "No se pudo medir CPU." "Medio"
    }

    if ($Sample.RAMAveragePct -ne $null) {
        if ($Sample.RAMAveragePct -gt 90) {
            Add-Penalty 15 "RAM" "RAM crítica: $($Sample.RAMAveragePct)%." "Critico"
        } elseif ($Sample.RAMAveragePct -gt 80) {
            Add-Penalty 10 "RAM" "RAM alta: $($Sample.RAMAveragePct)%." "Alto"
        } elseif ($Sample.RAMAveragePct -gt 60) {
            Add-Penalty 5 "RAM" "RAM moderada: $($Sample.RAMAveragePct)%." "Medio"
        }
    }

    foreach ($v in $Disks.Volumes) {
        if ($v.LibrePct -lt 10) {
            Add-Penalty 10 "Disco" "Unidad $($v.Unidad) con espacio crítico: $($v.LibrePct)%." "Critico"
        } elseif ($v.LibrePct -lt 15) {
            Add-Penalty 6 "Disco" "Unidad $($v.Unidad) con poco espacio: $($v.LibrePct)%." "Alto"
        } elseif ($v.LibrePct -lt 25) {
            Add-Penalty 3 "Disco" "Unidad $($v.Unidad) con espacio bajo: $($v.LibrePct)%." "Medio"
        }
    }

    foreach ($d in $Disks.PhysicalDisk) {
        if ($d.HealthStatus -and $d.HealthStatus -ne "Healthy") {
            Add-Penalty 15 "Disco" "Disco $($d.Modelo) salud no saludable: $($d.HealthStatus)." "Critico"
        }

        if ($d.Status -and $d.Status -notmatch "OK|Healthy|Online") {
            Add-Penalty 8 "Disco" "Disco $($d.Modelo) estado: $($d.Status)." "Alto"
        }

        if ($d.PhysicalMediaType -eq "HDD" -or $d.MediaType -like "*Hard*") {
            Add-Penalty 5 "Disco" "Disco mecánico detectado: $($d.Modelo)." "Medio"
        }
    }

    foreach ($s in $Smart) {
        if ($s.Available -eq $true -and $s.Health -and $s.Health -notmatch "PASSED|OK") {
            Add-Penalty 20 "SMART" "SMART no saludable en $($s.Device): $($s.Health)." "Critico"
        }

        if ($s.TemperatureC -gt 60) {
            Add-Penalty 10 "SMART" "Temperatura de disco alta en $($s.Device): $($s.TemperatureC) C." "Alto"
        } elseif ($s.TemperatureC -gt 50) {
            Add-Penalty 5 "SMART" "Temperatura de disco moderada en $($s.Device): $($s.TemperatureC) C." "Medio"
        }
    }

    foreach ($t in $Temps) {
        if ($t.TemperaturaC -ne $null) {
            if ($t.TemperaturaC -gt 90) {
                Add-Penalty 15 "Temperatura" "Temperatura crítica en $($t.Sensor): $($t.TemperaturaC) C." "Critico"
            } elseif ($t.TemperaturaC -gt 80) {
                Add-Penalty 10 "Temperatura" "Temperatura alta en $($t.Sensor): $($t.TemperaturaC) C." "Alto"
            } elseif ($t.TemperaturaC -gt 70) {
                Add-Penalty 5 "Temperatura" "Temperatura moderada en $($t.Sensor): $($t.TemperaturaC) C." "Medio"
            }
        }
    }

    if ($Startup.Count -gt 35) {
        Add-Penalty 10 "Inicio" "Exceso de programas de inicio: $($Startup.Count)." "Alto"
    } elseif ($Startup.Count -gt 20) {
        Add-Penalty 6 "Inicio" "Muchos programas de inicio: $($Startup.Count)." "Medio"
    } elseif ($Startup.Count -gt 10) {
        Add-Penalty 3 "Inicio" "Programas de inicio moderados: $($Startup.Count)." "Bajo"
    }

    if ($Events.TotalEventos -gt 20) {
        Add-Penalty 15 "Eventos" "Muchos eventos relevantes: $($Events.TotalEventos)." "Critico"
    } elseif ($Events.TotalEventos -gt 10) {
        Add-Penalty 10 "Eventos" "Eventos elevados: $($Events.TotalEventos)." "Alto"
    } elseif ($Events.TotalEventos -gt 3) {
        Add-Penalty 5 "Eventos" "Algunos eventos relevantes: $($Events.TotalEventos)." "Medio"
    }

    try {
        if ($Security.Defender.RealTimeProtection -eq $false) {
            Add-Penalty 5 "Seguridad" "Defender sin protección en tiempo real." "Medio"
        }
    } catch {}

    if ($BenchmarkResult -and $BenchmarkResult.DiskWriteMBs) {
        if ($BenchmarkResult.DiskWriteMBs -lt 30) {
            Add-Penalty 8 "Benchmark" "Escritura muy baja: $($BenchmarkResult.DiskWriteMBs) MB/s." "Alto"
        } elseif ($BenchmarkResult.DiskWriteMBs -lt 80) {
            Add-Penalty 4 "Benchmark" "Escritura baja/moderada: $($BenchmarkResult.DiskWriteMBs) MB/s." "Medio"
        }
    }

    if ($score -lt 0) {
        $score = 0
    }

    $estado = if ($score -ge 90) {
        "Excelente"
    } elseif ($score -ge 75) {
        "Bueno"
    } elseif ($score -ge 60) {
        "Regular"
    } elseif ($score -ge 40) {
        "Malo"
    } else {
        "Critico"
    }

    [PSCustomObject]@{
        Score     = $score
        Estado    = $estado
        Hallazgos = $findings
    }
}

function New-Recommendations {
    param(
        $Sample,
        $Disks,
        $Temps,
        $Smart,
        $Startup,
        $Events
    )

    $recs = @()

    if ($Sample.CPUAveragePct -gt 40) {
        $recs += "Revisar procesos residentes: CPU promedio elevada ($($Sample.CPUAveragePct)%)."
    }

    if ($Sample.RAMAveragePct -gt 80) {
        $recs += "Revisar programas de inicio o ampliar RAM: uso promedio $($Sample.RAMAveragePct)%."
    }

    foreach ($v in $Disks.Volumes) {
        if ($v.LibrePct -lt 15) {
            $recs += "Liberar espacio en $($v.Unidad): libre $($v.LibrePct)%."
        }
    }

    foreach ($d in $Disks.PhysicalDisk) {
        if ($d.PhysicalMediaType -eq "HDD" -or $d.MediaType -like "*Hard*") {
            $recs += "Evaluar migración a SSD: disco mecánico detectado."
        }

        if ($d.HealthStatus -and $d.HealthStatus -ne "Healthy") {
            $recs += "Respaldar información y revisar disco: Health=$($d.HealthStatus)."
        }
    }

    foreach ($t in $Temps) {
        if ($t.TemperaturaC -gt 80) {
            $recs += "Revisar limpieza física, ventilación y pasta térmica: $($t.Sensor) $($t.TemperaturaC) C."
        }
    }

    foreach ($s in $Smart) {
        if ($s.Available -and $s.Health -and $s.Health -notmatch "PASSED|OK") {
            $recs += "Respaldar y reemplazar/revisar disco: SMART $($s.Device) = $($s.Health)."
        }
    }

    if ($Startup.Count -gt 20) {
        $recs += "Reducir programas de inicio: detectados $($Startup.Count)."
    }

    if ($Events.TotalEventos -gt 10) {
        $recs += "Revisar visor de eventos: $($Events.TotalEventos) eventos relevantes recientes."
    }

    if (!$recs.Count) {
        $recs += "No se detectan problemas graves automáticos. Mantener limpieza, actualizaciones y respaldo periódico."
    }

    return $recs
}

function Export-CsvTables {
    param(
        $Folder,
        $Report
    )

    try { $Report.Processes.TopCPU | Export-Csv (Join-Path $Folder "procesos_top_cpu.csv") -NoTypeInformation -Encoding UTF8 } catch {}
    try { $Report.Processes.TopRAM | Export-Csv (Join-Path $Folder "procesos_top_ram.csv") -NoTypeInformation -Encoding UTF8 } catch {}
    try { $Report.Disks.Volumes | Export-Csv (Join-Path $Folder "volumenes.csv") -NoTypeInformation -Encoding UTF8 } catch {}
    try { $Report.Disks.PhysicalDisk | Export-Csv (Join-Path $Folder "discos_fisicos.csv") -NoTypeInformation -Encoding UTF8 } catch {}
    try { $Report.Disks.DiskPerf | Export-Csv (Join-Path $Folder "disco_perf.csv") -NoTypeInformation -Encoding UTF8 } catch {}
    try { $Report.Temperatures | Export-Csv (Join-Path $Folder "temperaturas.csv") -NoTypeInformation -Encoding UTF8 } catch {}
    try { $Report.SmartCtl | Select-Object Device,Available,Health,TemperatureC | Export-Csv (Join-Path $Folder "smartctl_resumen.csv") -NoTypeInformation -Encoding UTF8 } catch {}
    try { $Report.StartupPrograms | Export-Csv (Join-Path $Folder "programas_inicio.csv") -NoTypeInformation -Encoding UTF8 } catch {}
    try { $Report.Events.Eventos | Export-Csv (Join-Path $Folder "eventos.csv") -NoTypeInformation -Encoding UTF8 } catch {}
    try { $Report.Services.ListaServicios | Export-Csv (Join-Path $Folder "servicios.csv") -NoTypeInformation -Encoding UTF8 } catch {}
}

function New-TextSummary {
    param($Report)

    $lines = @()

    $lines += "REPORTE DE RENDIMIENTO V2 - SERVICIO TECNICO"
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
    $lines += "CPU promedio: $($Report.Sample.CPUAveragePct)%"
    $lines += "RAM usada promedio: $($Report.Sample.RAMAveragePct)%"
    $lines += "RAM libre promedio: $($Report.Sample.RAMFreeAvgGB) GB"
    $lines += "Procesos: $($Report.Processes.TotalProcesos)"
    $lines += "Inicio: $($Report.StartupPrograms.Count)"
    $lines += "Eventos relevantes: $($Report.Events.TotalEventos)"
    $lines += ""
    $lines += "Temperaturas:"

    foreach ($t in $Report.Temperatures) {
        $lines += "- $($t.Sensor): $($t.TemperaturaC) C [$($t.Fuente)]"
    }

    $lines += ""
    $lines += "Volumenes:"

    foreach ($v in $Report.Disks.Volumes) {
        $lines += "- $($v.Unidad): $($v.TamanoGB) GB / Libre $($v.LibreGB) GB ($($v.LibrePct)%)"
    }

    $lines += ""
    $lines += "Hallazgos:"

    foreach ($h in $Report.Score.Hallazgos) {
        $lines += "- [$($h.Severidad)] $($h.Area): $($h.Mensaje)"
    }

    $lines += ""
    $lines += "Recomendaciones:"

    foreach ($r in $Report.Recommendations) {
        $lines += "- $r"
    }

    return $lines -join "`r`n"
}

function New-HtmlReport {
    param($Report)

    $scoreClass = if ($Report.Score.Score -ge 90) {
        "good"
    } elseif ($Report.Score.Score -ge 75) {
        "ok"
    } elseif ($Report.Score.Score -ge 60) {
        "warn"
    } else {
        "bad"
    }

    $html = @"
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<title>Reporte Rendimiento V2 - $($Report.System.Hostname)</title>
<style>
body{font-family:Segoe UI,Arial,sans-serif;background:#f4f6f8;color:#222;margin:0;padding:24px}
.container{max-width:1200px;margin:auto}
.header{background:#111827;color:white;padding:24px;border-radius:14px;margin-bottom:20px}
.grid{display:grid;grid-template-columns:repeat(4,1fr);gap:14px}
.card{background:white;padding:18px;border-radius:14px;box-shadow:0 2px 12px rgba(0,0,0,.08);margin-bottom:16px}
.metric{font-size:28px;font-weight:700}
.label{color:#6b7280;font-size:13px}
.good{color:#047857}
.ok{color:#2563eb}
.warn{color:#b45309}
.bad{color:#b91c1c}
table{width:100%;border-collapse:collapse;margin-top:8px}
th{background:#e5e7eb;text-align:left;padding:8px;font-size:13px}
td{border-bottom:1px solid #e5e7eb;padding:8px;font-size:13px;vertical-align:top}
.badge{display:inline-block;padding:4px 8px;border-radius:999px;background:#e5e7eb;font-size:12px}
.finding-Critico{color:#991b1b;font-weight:700}
.finding-Alto{color:#b45309;font-weight:700}
@media print{body{background:white}.card{box-shadow:none;border:1px solid #ddd}}
</style>
</head>
<body>
<div class="container">

<div class="header">
<h1>Reporte de Rendimiento V2 - Servicio Técnico</h1>
<p>Modo: <strong>$($Report.Metadata.Modo)</strong> | Cliente: <strong>$(HtmlEncode $Report.Metadata.Cliente)</strong> | OT: <strong>$(HtmlEncode $Report.Metadata.OrdenTrabajo)</strong></p>
<p>Técnico: $(HtmlEncode $Report.Metadata.Tecnico) | Fecha: $($Report.Metadata.Fecha)</p>
</div>

<div class="grid">
<div class="card"><div class="label">Score</div><div class="metric $scoreClass">$($Report.Score.Score)/100</div><div class="badge">$($Report.Score.Estado)</div></div>
<div class="card"><div class="label">CPU promedio</div><div class="metric">$($Report.Sample.CPUAveragePct)%</div><div class="label">Max: $($Report.Sample.CPUMaxPct)%</div></div>
<div class="card"><div class="label">RAM usada</div><div class="metric">$($Report.Sample.RAMAveragePct)%</div><div class="label">Libre: $($Report.Sample.RAMFreeAvgGB) GB</div></div>
<div class="card"><div class="label">Eventos relevantes</div><div class="metric">$($Report.Events.TotalEventos)</div><div class="label">Últimos $($Report.Events.DiasAnalizados) días</div></div>
</div>

<div class="card">
<h2>Ficha del equipo</h2>
<table>
<tr><th>Campo</th><th>Valor</th></tr>
<tr><td>Hostname</td><td>$(HtmlEncode $Report.System.Hostname)</td></tr>
<tr><td>Usuario</td><td>$(HtmlEncode $Report.System.Usuario)</td></tr>
<tr><td>Administrador</td><td>$($Report.System.EsAdministrador)</td></tr>
<tr><td>Fabricante / Modelo</td><td>$(HtmlEncode $Report.System.Fabricante) $(HtmlEncode $Report.System.Modelo)</td></tr>
<tr><td>Serial</td><td>$(HtmlEncode $Report.System.Serial)</td></tr>
<tr><td>Windows</td><td>$(HtmlEncode $Report.System.Windows) $($Report.System.Version) Build $($Report.System.Build)</td></tr>
<tr><td>CPU</td><td>$(HtmlEncode $Report.System.CPU) / Cores $($Report.System.CPUCores) / Hilos $($Report.System.CPUThreads)</td></tr>
<tr><td>RAM</td><td>$($Report.System.RAMTotalGB) GB</td></tr>
<tr><td>Último arranque</td><td>$($Report.System.UltimoArranque)</td></tr>
</table>
</div>

<div class="card">
<h2>Temperaturas</h2>
<table>
<tr><th>Sensor</th><th>Tipo</th><th>Temperatura C</th><th>Fuente</th></tr>
"@

    foreach ($t in $Report.Temperatures) {
        $html += "<tr><td>$(HtmlEncode $t.Sensor)</td><td>$(HtmlEncode $t.Tipo)</td><td>$($t.TemperaturaC)</td><td>$(HtmlEncode $t.Fuente)</td></tr>"
    }

    $html += "</table></div>"

    $html += '<div class="card"><h2>SMART / smartctl</h2><table><tr><th>Device</th><th>Disponible</th><th>Health</th><th>Temp C</th></tr>'

    foreach ($s in $Report.SmartCtl) {
        $html += "<tr><td>$(HtmlEncode $s.Device)</td><td>$($s.Available)</td><td>$(HtmlEncode $s.Health)</td><td>$($s.TemperatureC)</td></tr>"
    }

    $html += "</table></div>"

    $html += '<div class="card"><h2>Discos y volúmenes</h2><table><tr><th>Unidad</th><th>Nombre</th><th>FS</th><th>Tamaño GB</th><th>Libre GB</th><th>Libre %</th><th>Usado %</th></tr>'

    foreach ($v in $Report.Disks.Volumes) {
        $html += "<tr><td>$($v.Unidad)</td><td>$(HtmlEncode $v.Nombre)</td><td>$($v.Sistema)</td><td>$($v.TamanoGB)</td><td>$($v.LibreGB)</td><td>$($v.LibrePct)</td><td>$($v.UsadoPct)</td></tr>"
    }

    $html += "</table></div>"

    $html += '<div class="card"><h2>Hallazgos</h2><table><tr><th>Severidad</th><th>Área</th><th>Penalidad</th><th>Mensaje</th></tr>'

    foreach ($h in $Report.Score.Hallazgos) {
        $html += "<tr><td class='finding-$($h.Severidad)'>$($h.Severidad)</td><td>$($h.Area)</td><td>-$($h.Penalidad)</td><td>$(HtmlEncode $h.Mensaje)</td></tr>"
    }

    if (!$Report.Score.Hallazgos.Count) {
        $html += "<tr><td colspan='4'>Sin hallazgos negativos relevantes.</td></tr>"
    }

    $html += "</table></div>"

    $html += '<div class="card"><h2>Recomendaciones</h2><ul>'

    foreach ($r in $Report.Recommendations) {
        $html += "<li>$(HtmlEncode $r)</li>"
    }

    $html += "</ul></div>"

    $html += '<div class="card"><h2>Top procesos por CPU</h2><table><tr><th>Proceso</th><th>PID</th><th>CPU s</th><th>RAM MB</th><th>Ruta</th></tr>'

    foreach ($p in $Report.Processes.TopCPU) {
        $html += "<tr><td>$(HtmlEncode $p.Name)</td><td>$($p.Id)</td><td>$($p.CPUSeconds)</td><td>$($p.RAMMB)</td><td>$(HtmlEncode $p.Path)</td></tr>"
    }

    $html += "</table></div>"

    $html += '<div class="card"><h2>Top procesos por RAM</h2><table><tr><th>Proceso</th><th>PID</th><th>RAM MB</th><th>CPU s</th><th>Ruta</th></tr>'

    foreach ($p in $Report.Processes.TopRAM) {
        $html += "<tr><td>$(HtmlEncode $p.Name)</td><td>$($p.Id)</td><td>$($p.RAMMB)</td><td>$($p.CPUSeconds)</td><td>$(HtmlEncode $p.Path)</td></tr>"
    }

    $html += "</table></div>"

    $html += '<div class="card"><h2>Programas de inicio</h2><table><tr><th>Nombre</th><th>Ubicación</th><th>Usuario</th><th>Comando</th></tr>'

    foreach ($s in $Report.StartupPrograms) {
        $html += "<tr><td>$(HtmlEncode $s.Nombre)</td><td>$(HtmlEncode $s.Ubicacion)</td><td>$(HtmlEncode $s.Usuario)</td><td>$(HtmlEncode $s.Comando)</td></tr>"
    }

    $html += "</table></div>"

    if ($Report.Benchmark) {
        $html += '<div class="card"><h2>Benchmark liviano</h2><table><tr><th>Prueba</th><th>Resultado</th></tr>'
        $html += "<tr><td>CPU segundos</td><td>$($Report.Benchmark.CPU.Segundos)</td></tr>"
        $html += "<tr><td>Disco escritura</td><td>$($Report.Benchmark.DiskWriteMBs) MB/s</td></tr>"
        $html += "<tr><td>Disco lectura</td><td>$($Report.Benchmark.DiskReadMBs) MB/s</td></tr>"
        $html += "</table></div>"
    }

    $html += '<div class="card"><h2>Eventos recientes</h2><table><tr><th>Fecha</th><th>Nivel</th><th>Provider</th><th>ID</th><th>Mensaje</th></tr>'

    foreach ($e in ($Report.Events.Eventos | Select-Object -First 40)) {
        $html += "<tr><td>$($e.Fecha)</td><td>$($e.Nivel)</td><td>$(HtmlEncode $e.Provider)</td><td>$($e.Id)</td><td>$(HtmlEncode $e.Mensaje)</td></tr>"
    }

    $html += "</table></div>"
    $html += "</div></body></html>"

    return $html
}

Write-Host "================================================" -ForegroundColor Cyan
Write-Host "REPORTE DE RENDIMIENTO V2 - SERVICIO TECNICO" -ForegroundColor Cyan
Write-Host "================================================" -ForegroundColor Cyan

if ($DuracionMuestreo -lt 5) {
    $DuracionMuestreo = 5
}

if (!(Test-Path $RutaSalida)) {
    New-Item -Path $RutaSalida -ItemType Directory -Force | Out-Null
}

$runFolder = New-ReportFolders -Root $RutaSalida -Cliente $Cliente -OT $OrdenTrabajo -Modo $Modo

Write-Info "Carpeta de reporte: $runFolder"

$system = Get-SystemInfoSafe
$sample = Get-CpuMemorySample -Seconds $DuracionMuestreo
$processes = Get-TopProcessesSafe
$disks = Get-DiskInfoSafe
$temps = Get-TemperatureInfoSafe -Advanced:$UseLibreHardwareMonitor.IsPresent -LhmDll $LibreHardwareMonitorPath

if ($UseSmartCtl) {
    $smart = @(Get-SmartCtlInfo -SmartCtlPath $SmartCtlPath)
} else {
    $smart = @([PSCustomObject]@{
        Device       = "smartctl"
        Available    = $false
        Health       = $null
        TemperatureC = $null
        Raw          = "No solicitado"
    })
}

$startup = Get-StartupProgramsSafe
$services = Get-ServiceSummarySafe
$events = Get-CriticalEventsSafe -Days 7
$security = Get-SecurityBasicSafe

if ($Benchmark) {
    $bench = Invoke-LightBenchmark -TempFolder $runFolder
} else {
    $bench = $null
}

$score = Get-PerformanceScore `
    -Sample $sample `
    -Disks $disks `
    -Temps $temps `
    -Smart $smart `
    -Startup $startup `
    -Events $events `
    -Security $security `
    -BenchmarkResult $bench

$recommendations = New-Recommendations `
    -Sample $sample `
    -Disks $disks `
    -Temps $temps `
    -Smart $smart `
    -Startup $startup `
    -Events $events

$report = [PSCustomObject]@{
    Metadata = [PSCustomObject]@{
        Modo                     = $Modo
        Cliente                  = $Cliente
        OrdenTrabajo             = $OrdenTrabajo
        Tecnico                  = $Tecnico
        Fecha                    = Get-Date
        RutaReporte              = $runFolder
        ScriptVersion            = "2.0-standalone"
        UseLibreHardwareMonitor  = $UseLibreHardwareMonitor.IsPresent
        LibreHardwareMonitorPath = $LibreHardwareMonitorPath
        UseSmartCtl              = $UseSmartCtl.IsPresent
        SmartCtlPath             = $SmartCtlPath
    }
    System          = $system
    Sample          = $sample
    Processes       = $processes
    Disks           = $disks
    Temperatures    = $temps
    SmartCtl        = $smart
    StartupPrograms = $startup
    Services        = $services
    Events          = $events
    Security        = $security
    Benchmark       = $bench
    Score           = $score
    Recommendations = $recommendations
}

$jsonPath = Join-Path $runFolder "reporte.json"
$htmlPath = Join-Path $runFolder "reporte.html"
$txtPath = Join-Path $runFolder "resumen.txt"

$report | ConvertTo-Json -Depth 12 | Out-File $jsonPath -Encoding UTF8
Export-CsvTables -Folder $runFolder -Report $report
New-HtmlReport -Report $report | Out-File $htmlPath -Encoding UTF8
New-TextSummary -Report $report | Out-File $txtPath -Encoding UTF8

Write-Host ""
Write-Host "================================================" -ForegroundColor Green
Write-Host "REPORTE GENERADO" -ForegroundColor Green
Write-Host "================================================" -ForegroundColor Green
Write-Host "REPORTE_GENERADO=$runFolder" -ForegroundColor Green
Write-Host "SCORE=$($score.Score)"
Write-Host "ESTADO=$($score.Estado)"
Write-Host "HTML=$htmlPath"
Write-Host "JSON=$jsonPath"
Write-Host "TXT=$txtPath"
Write-Host ""