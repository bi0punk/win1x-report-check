function Write-Log {
    param([string]$Message)
    if ($script:LogPath) {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        try { Add-Content -Path $script:LogPath -Value "$timestamp $Message" -ErrorAction SilentlyContinue } catch { Write-Warning "Error: $_" }
    }
}

function Write-Info {
    param([string]$Message)
    Write-Host "[INFO] $Message" -ForegroundColor Cyan
    Write-Log "[INFO] $Message"
}

function Write-Warn2 {
    param([string]$Message)
    Write-Host "[WARN] $Message" -ForegroundColor Yellow
    Write-Log "[WARN] $Message"
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

function Get-NormalizedCpuScore {
    param([double]$OpsPerSecond, [int]$Workers)
    if ($OpsPerSecond -eq $null -or $Workers -lt 1) { return $null }
    $baselinePerWorker = 1800000
    $target = $baselinePerWorker * $Workers
    if ($target -le 0) { return $null }
    $score = (($OpsPerSecond / $target) * 100)
    $score = [math]::Max(0, [math]::Min(150, [math]::Round($score, 2)))
    if ($score -lt 5) { $score = 5 }
    return $score
}

function Get-NormalizedRamScore {
    param([double]$WriteMBs, [double]$ReadMBs)
    $avg = $null
    if (($WriteMBs -ne $null) -and ($ReadMBs -ne $null)) { $avg = ($WriteMBs + $ReadMBs) / 2 }
    elseif ($WriteMBs -ne $null) { $avg = $WriteMBs }
    elseif ($ReadMBs -ne $null) { $avg = $ReadMBs }
    if ($avg -eq $null) { return $null }
    $target = 18000
    if ($target -le 0) { return $null }
    $score = (($avg / $target) * 100)
    $score = [math]::Max(0, [math]::Min(150, [math]::Round($score, 2)))
    if ($score -lt 5) { $score = 5 }
    return $score
}

function Get-NormalizedDiskIOPSScore {
    param([double]$ReadIOPS, [double]$WriteIOPS)
    $avg = $null
    if (($ReadIOPS -ne $null) -and ($WriteIOPS -ne $null)) { $avg = ($ReadIOPS + $WriteIOPS) / 2 }
    elseif ($ReadIOPS -ne $null) { $avg = $ReadIOPS }
    elseif ($WriteIOPS -ne $null) { $avg = $WriteIOPS }
    if ($avg -eq $null) { return $null }
    $target = 2000
    if ($target -le 0) { return $null }
    $score = (($avg / $target) * 100)
    $score = [math]::Max(0, [math]::Min(150, [math]::Round($score, 2)))
    if ($score -lt 5) { $score = 5 }
    return $score
}

function Add-Finding {
    param([int]$Penalty, [string]$Area, [string]$Severity, [string]$Message)
    return [PSCustomObject]@{
        Area = $Area
        Severidad = $Severity
        Penalidad = $Penalty
        Mensaje = $Message
    }
}

function Get-Score {
    param($Reposo, $Benchmark, $Disks, $DiskAdvanced, $TempsFinal, $Smart, $Startup, $Events, $Security, $RamBenchmark, $Network)
    $score = 100
    $findings = @()

    if ($Reposo.CPUAveragePct -ne $null) {
        if ($Reposo.CPUAveragePct -gt 40) {
            $score -= 10
            $findings += Add-Finding -Penalty 10 -Area "CPU Reposo" -Severity "Alto" -Message "CPU alta en reposo: $($Reposo.CPUAveragePct)%."
        } elseif ($Reposo.CPUAveragePct -gt 20) {
            $score -= 5
            $findings += Add-Finding -Penalty 5 -Area "CPU Reposo" -Severity "Medio" -Message "CPU moderada en reposo: $($Reposo.CPUAveragePct)%."
        }
    }

    if ($Reposo.RAMAveragePct -ne $null) {
        if ($Reposo.RAMAveragePct -gt 90) {
            $score -= 15
            $findings += Add-Finding -Penalty 15 -Area "RAM" -Severity "Critico" -Message "RAM critica en reposo: $($Reposo.RAMAveragePct)%."
        } elseif ($Reposo.RAMAveragePct -gt 80) {
            $score -= 10
            $findings += Add-Finding -Penalty 10 -Area "RAM" -Severity "Alto" -Message "RAM alta en reposo: $($Reposo.RAMAveragePct)%."
        } elseif ($Reposo.RAMAveragePct -gt 65) {
            $score -= 5
            $findings += Add-Finding -Penalty 5 -Area "RAM" -Severity "Medio" -Message "RAM moderada en reposo: $($Reposo.RAMAveragePct)%."
        }
    }

    if ($Benchmark.Phase.TempMaxC -ne $null) {
        if ($Benchmark.Phase.TempMaxC -gt 95) {
            $score -= 20
            $findings += Add-Finding -Penalty 20 -Area "Temperatura Carga" -Severity "Critico" -Message "Temperatura maxima critica durante prueba: $($Benchmark.Phase.TempMaxC) C."
        } elseif ($Benchmark.Phase.TempMaxC -gt 85) {
            $score -= 15
            $findings += Add-Finding -Penalty 15 -Area "Temperatura Carga" -Severity "Alto" -Message "Temperatura maxima alta durante prueba: $($Benchmark.Phase.TempMaxC) C."
        } elseif ($Benchmark.Phase.TempMaxC -gt 75) {
            $score -= 7
            $findings += Add-Finding -Penalty 7 -Area "Temperatura Carga" -Severity "Medio" -Message "Temperatura maxima moderada durante prueba: $($Benchmark.Phase.TempMaxC) C."
        }
    }

    if ($Benchmark.CpuScore -ne $null) {
        if ($Benchmark.CpuScore -lt 40) {
            $score -= 12
            $findings += Add-Finding -Penalty 12 -Area "Benchmark CPU" -Severity "Critico" -Message "Puntaje CPU muy bajo: $($Benchmark.CpuScore)/100 (Ops/s $($Benchmark.CpuOpsPerSecond))."
        } elseif ($Benchmark.CpuScore -lt 60) {
            $score -= 6
            $findings += Add-Finding -Penalty 6 -Area "Benchmark CPU" -Severity "Medio" -Message "Puntaje CPU moderado: $($Benchmark.CpuScore)/100."
        }
    }

    if (($Reposo.TempAverageC -ne $null) -and ($Benchmark.Phase.TempMaxC -ne $null)) {
        $deltaTemp = [math]::Round($Benchmark.Phase.TempMaxC - $Reposo.TempAverageC, 1)
        if ($deltaTemp -gt 40) {
            $score -= 10
            $findings += Add-Finding -Penalty 10 -Area "Temperatura Delta" -Severity "Alto" -Message "Aumento termico fuerte: +$deltaTemp C entre reposo y carga."
        } elseif ($deltaTemp -gt 30) {
            $score -= 5
            $findings += Add-Finding -Penalty 5 -Area "Temperatura Delta" -Severity "Medio" -Message "Aumento termico moderado: +$deltaTemp C entre reposo y carga."
        }
    }

    if ($RamBenchmark -ne $null -and $RamBenchmark.RamScore -ne $null) {
        if ($RamBenchmark.RamScore -lt 40) {
            $score -= 8
            $findings += Add-Finding -Penalty 8 -Area "Benchmark RAM" -Severity "Alto" -Message "Rendimiento RAM bajo: $($RamBenchmark.RamScore)/100 (aprox. $($RamBenchmark.AvgMBs) MB/s)."
        } elseif ($RamBenchmark.RamScore -lt 60) {
            $score -= 4
            $findings += Add-Finding -Penalty 4 -Area "Benchmark RAM" -Severity "Medio" -Message "Rendimiento RAM moderado: $($RamBenchmark.RamScore)/100."
        }
    }

    foreach ($v in $Disks.Volumes) {
        if ($v.LibrePct -lt 10) {
            $score -= 10
            $findings += Add-Finding -Penalty 10 -Area "Disco" -Severity "Critico" -Message "Unidad $($v.Unidad) con espacio critico: $($v.LibrePct)%."
        } elseif ($v.LibrePct -lt 15) {
            $score -= 6
            $findings += Add-Finding -Penalty 6 -Area "Disco" -Severity "Alto" -Message "Unidad $($v.Unidad) con poco espacio: $($v.LibrePct)%."
        } elseif ($v.LibrePct -lt 25) {
            $score -= 3
            $findings += Add-Finding -Penalty 3 -Area "Disco" -Severity "Medio" -Message "Unidad $($v.Unidad) con espacio bajo: $($v.LibrePct)%."
        }
    }

    foreach ($d in $Disks.PhysicalDisk) {
        if (($d.HealthStatus -ne $null) -and ($d.HealthStatus -ne "Healthy")) {
            $score -= 15
            $findings += Add-Finding -Penalty 15 -Area "Disco" -Severity "Critico" -Message "Disco $($d.Modelo) salud no saludable: $($d.HealthStatus)."
        }
        if (($d.PhysicalMediaType -eq "HDD") -or ($d.MediaType -like "*Hard*")) {
            $score -= 5
            $findings += Add-Finding -Penalty 5 -Area "Disco" -Severity "Medio" -Message "Disco mecanico detectado: $($d.Modelo)."
        }
    }

    if ($Benchmark.DiskWriteMBs -ne $null) {
        if ($Benchmark.DiskWriteMBs -lt 25) {
            $score -= 10
            $findings += Add-Finding -Penalty 10 -Area "Benchmark Disco" -Severity "Alto" -Message "Escritura promedio muy baja durante prueba: $($Benchmark.DiskWriteMBs) MB/s."
        } elseif ($Benchmark.DiskWriteMBs -lt 70) {
            $score -= 5
            $findings += Add-Finding -Penalty 5 -Area "Benchmark Disco" -Severity "Medio" -Message "Escritura promedio baja/moderada durante prueba: $($Benchmark.DiskWriteMBs) MB/s."
        }
    }

    if ($Benchmark.Disk4K -ne $null) {
        if ($Benchmark.Disk4K.IOScore -lt 45) {
            $score -= 8
            $findings += Add-Finding -Penalty 8 -Area "Benchmark Disco" -Severity "Alto" -Message "IOPS 4K bajo: Read $($Benchmark.Disk4K.ReadIOPS) / Write $($Benchmark.Disk4K.WriteIOPS)."
        } elseif ($Benchmark.Disk4K.IOScore -lt 70) {
            $score -= 4
            $findings += Add-Finding -Penalty 4 -Area "Benchmark Disco" -Severity "Medio" -Message "IOPS 4K moderado: $($Benchmark.Disk4K.IOScore)/100."
        }
    }

    foreach ($s in $Smart) {
        if (($s.Available -eq $true) -and ($s.Health -ne $null) -and ($s.Health -notmatch "PASSED|OK")) {
            $score -= 20
            $findings += Add-Finding -Penalty 20 -Area "SMART" -Severity "Critico" -Message "SMART no saludable en $($s.Device): $($s.Health)."
        }
        if ($s.TemperatureC -gt 60) {
            $score -= 10
            $findings += Add-Finding -Penalty 10 -Area "SMART" -Severity "Alto" -Message "Temperatura de disco alta: $($s.TemperatureC) C."
        }
    }

    if ($DiskAdvanced -ne $null) {
        foreach ($w in $DiskAdvanced.WmiSmart) {
            if ($w.PredictFailure -eq $true) {
                $score -= 25
                $findings += Add-Finding -Penalty 25 -Area "SMART WMI" -Severity "Critico" -Message "Windows predice falla SMART: $($w.InstanceName). Reason=$($w.Reason)"
            }
        }

        foreach ($r in $DiskAdvanced.StorageReliability) {
            if ($r.ReadErrorsTotal -gt 0) {
                $score -= 8
                $findings += Add-Finding -Penalty 8 -Area "Disco Errores" -Severity "Alto" -Message "Errores de lectura reportados por Storage Reliability: $($r.ReadErrorsTotal)."
            }
            if ($r.WriteErrorsTotal -gt 0) {
                $score -= 8
                $findings += Add-Finding -Penalty 8 -Area "Disco Errores" -Severity "Alto" -Message "Errores de escritura reportados por Storage Reliability: $($r.WriteErrorsTotal)."
            }
            if ($r.Wear -ne $null) {
                if ($r.Wear -gt 80) {
                    $score -= 15
                    $findings += Add-Finding -Penalty 15 -Area "Desgaste SSD" -Severity "Alto" -Message "Desgaste SSD alto reportado: $($r.Wear)."
                } elseif ($r.Wear -gt 60) {
                    $score -= 8
                    $findings += Add-Finding -Penalty 8 -Area "Desgaste SSD" -Severity "Medio" -Message "Desgaste SSD moderado reportado: $($r.Wear)."
                }
            }
            if ($r.TemperatureC -gt 60) {
                $score -= 10
                $findings += Add-Finding -Penalty 10 -Area "Temperatura Disco" -Severity "Alto" -Message "Temperatura de disco alta: $($r.TemperatureC) C."
            } elseif ($r.TemperatureC -gt 50) {
                $score -= 5
                $findings += Add-Finding -Penalty 5 -Area "Temperatura Disco" -Severity "Medio" -Message "Temperatura de disco moderada: $($r.TemperatureC) C."
            }
        }

        foreach ($f in $DiskAdvanced.Fragmentation) {
            if ($f.FragmentationPercent -ne $null) {
                if (($f.MediaType -notmatch "SSD|Unspecified") -and ($f.FragmentationPercent -gt 20)) {
                    $score -= 6
                    $findings += Add-Finding -Penalty 6 -Area "Fragmentacion" -Severity "Medio" -Message "Fragmentacion alta en $($f.Drive): $($f.FragmentationPercent)%."
                }
            }
        }

        foreach ($c in $DiskAdvanced.Chkdsk) {
            if ($c.Status -eq "Problemas detectados") {
                $score -= 15
                $findings += Add-Finding -Penalty 15 -Area "CHKDSK" -Severity "Critico" -Message "CHKDSK reporta problemas en $($c.Drive)."
            }
        }

        foreach ($t in $DiskAdvanced.Trim) {
            if ($t.TrimEnabled -eq $false) {
                $score -= 5
                $findings += Add-Finding -Penalty 5 -Area "TRIM" -Severity "Medio" -Message "TRIM parece desactivado para $($t.FileSystem)."
            }
        }
    }

    if ($Startup.Count -gt 35) {
        $score -= 10
        $findings += Add-Finding -Penalty 10 -Area "Inicio" -Severity "Alto" -Message "Exceso de programas de inicio: $($Startup.Count)."
    } elseif ($Startup.Count -gt 20) {
        $score -= 6
        $findings += Add-Finding -Penalty 6 -Area "Inicio" -Severity "Medio" -Message "Muchos programas de inicio: $($Startup.Count)."
    }

    if ($Events.TotalEventos -gt 20) {
        $score -= 15
        $findings += Add-Finding -Penalty 15 -Area "Eventos" -Severity "Critico" -Message "Muchos eventos relevantes: $($Events.TotalEventos)."
    } elseif ($Events.TotalEventos -gt 10) {
        $score -= 10
        $findings += Add-Finding -Penalty 10 -Area "Eventos" -Severity "Alto" -Message "Eventos elevados: $($Events.TotalEventos)."
    } elseif ($Events.TotalEventos -gt 3) {
        $score -= 5
        $findings += Add-Finding -Penalty 5 -Area "Eventos" -Severity "Medio" -Message "Algunos eventos relevantes: $($Events.TotalEventos)."
    }

    if ($null -ne $Security.Defender -and $Security.Defender.RealTimeProtection -eq $false) {
        $score -= 5
        $findings += Add-Finding -Penalty 5 -Area "Seguridad" -Severity "Medio" -Message "Defender sin proteccion en tiempo real."
    }

    if ($Network -ne $null) {
        foreach ($n in $Network) {
            if ($n.SuccessRatePct -lt 60) {
                $score -= 6
                $findings += Add-Finding -Penalty 6 -Area "Red" -Severity "Medio" -Message "Pings a $($n.Target) con tasas bajas: $($n.SuccessRatePct)%."
            }
            if ($n.AvgLatencyMs -ne $null -and $n.AvgLatencyMs -gt 120) {
                $score -= 5
                $findings += Add-Finding -Penalty 5 -Area "Red" -Severity "Medio" -Message "Latencia elevada a $($n.Target): $($n.AvgLatencyMs) ms."
            }
        }
    }

    if ($score -lt 0) { $score = 0 }
    $estado = "Critico"
    if ($score -ge 90) { $estado = "Excelente" }
    elseif ($score -ge 75) { $estado = "Bueno" }
    elseif ($score -ge 60) { $estado = "Regular" }
    elseif ($score -ge 40) { $estado = "Malo" }

    return [PSCustomObject]@{
        Score = $score
        Estado = $estado
        Hallazgos = $findings
    }
}

function Get-Recommendations {
    param($Reposo, $Benchmark, $Disks, $Smart, $Startup, $Events, $RamBenchmark, $Network)
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
    if ($RamBenchmark -ne $null) {
        if ($RamBenchmark.AvgMBs -ne $null -and $RamBenchmark.AvgMBs -lt 6000) {
            $recs += "Revisar configuracion de memoria: throughput promedio $($RamBenchmark.AvgMBs) MB/s."
        }
    }
    if ($Benchmark.Disk4K -ne $null) {
        if ($Benchmark.Disk4K.ReadIOPS -lt 800 -or $Benchmark.Disk4K.WriteIOPS -lt 600) {
            $recs += "Revisar salud del disco: IOPS 4K bajos (R $($Benchmark.Disk4K.ReadIOPS), W $($Benchmark.Disk4K.WriteIOPS))."
        }
    }
    if ($Startup.Count -gt 20) { $recs += "Reducir programas de inicio: detectados $($Startup.Count)." }
    if ($Events.TotalEventos -gt 10) { $recs += "Revisar visor de eventos: $($Events.TotalEventos) eventos relevantes recientes." }
    if ($Network -ne $null) {
        foreach ($n in $Network) {
            if ($n.SuccessRatePct -lt 60) { $recs += "Verificar conectividad a $($n.Target): fallan los pings ($($n.SuccessRatePct)% exitosos)." }
            if ($n.AvgLatencyMs -ne $null -and $n.AvgLatencyMs -gt 120) { $recs += "Reducir latencia a $($n.Target): promedio $($n.AvgLatencyMs) ms." }
        }
    }
    if ($recs.Count -eq 0) { $recs += "No se detectan problemas graves automaticos. Mantener limpieza, actualizaciones y respaldo periodico." }
    return $recs
}

Export-ModuleMember -Function @(
    'Write-Log','Write-Info','Write-Warn2','Safe-Name','Html-Encode','Is-Admin',
    'Get-NormalizedCpuScore','Get-NormalizedRamScore','Get-NormalizedDiskIOPSScore',
    'Add-Finding','Get-Score','Get-Recommendations'
)
