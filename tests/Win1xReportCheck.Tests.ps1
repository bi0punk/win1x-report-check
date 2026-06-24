BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\modules\Win1xUtilities.psm1'
    Remove-Module Win1xUtilities -ErrorAction SilentlyContinue
    Import-Module $modulePath -Force
}

Describe 'Safe-Name' {
    It 'returns SinDato for null or whitespace' {
        Safe-Name '' | Should -Be 'SinDato'
        Safe-Name '   ' | Should -Be 'SinDato'
        Safe-Name $null | Should -Be 'SinDato'
    }

    It 'replaces invalid path characters with underscores' {
        Safe-Name 'test:file?' | Should -Be 'test_file_'
    }

    It 'replaces whitespace with underscores' {
        Safe-Name 'my folder' | Should -Be 'my_folder'
    }

    It 'trims trailing and leading underscores' {
        Safe-Name '_hello_' | Should -Be 'hello'
    }

    It 'returns normal text unchanged' {
        Safe-Name 'HelloWorld123' | Should -Be 'HelloWorld123'
    }
}

Describe 'Html-Encode' {
    It 'returns empty string for null' {
        Html-Encode $null | Should -Be ''
    }

    It 'encodes HTML special characters' {
        Html-Encode '<test>' | Should -Be '&lt;test&gt;'
        Html-Encode '"quoted"' | Should -Be '&quot;quoted&quot;'
        Html-Encode '&' | Should -Be '&amp;'
    }

    It 'returns normal text as-is' {
        Html-Encode 'Hello World' | Should -Be 'Hello World'
    }

    It 'converts non-string objects to string' {
        Html-Encode 42 | Should -Be '42'
    }
}

Describe 'Is-Admin' {
    It 'returns a boolean value' {
        $result = Is-Admin
        $result -is [bool] | Should -Be $true
    }
}

Describe 'Get-NormalizedCpuScore' {
    It 'returns null for invalid inputs' {
        Get-NormalizedCpuScore -OpsPerSecond $null -Workers 2 | Should -Be $null
        Get-NormalizedCpuScore -OpsPerSecond 1000000 -Workers 0 | Should -Be $null
    }

    It 'returns baseline score for matching performance' {
        $score = Get-NormalizedCpuScore -OpsPerSecond 3600000 -Workers 2
        $score | Should -Be 100
    }

    It 'clamps to minimum of 5' {
        $score = Get-NormalizedCpuScore -OpsPerSecond 100 -Workers 8
        $score | Should -BeGreaterOrEqual 5
    }

    It 'clamps to maximum of 150' {
        $score = Get-NormalizedCpuScore -OpsPerSecond 100000000 -Workers 1
        $score | Should -BeLessOrEqual 150
    }

    It 'returns proportional score' {
        $score = Get-NormalizedCpuScore -OpsPerSecond 1800000 -Workers 2
        $score | Should -Be 50
    }
}

Describe 'Get-NormalizedRamScore' {
    It 'returns null for null inputs' {
        Get-NormalizedRamScore -WriteMBs $null -ReadMBs $null | Should -Be $null
    }

    It 'scores single input correctly' {
        $score = Get-NormalizedRamScore -WriteMBs 9000 -ReadMBs $null
        $score | Should -Be 50
    }

    It 'uses average of both inputs' {
        $score = Get-NormalizedRamScore -WriteMBs 9000 -ReadMBs 27000
        $score | Should -Be 100
    }

    It 'clamps to minimum of 5' {
        $score = Get-NormalizedRamScore -WriteMBs 10 -ReadMBs 10
        $score | Should -BeGreaterOrEqual 5
    }

    It 'clamps to maximum of 150' {
        $score = Get-NormalizedRamScore -WriteMBs 50000 -ReadMBs 50000
        $score | Should -BeLessOrEqual 150
    }
}

Describe 'Get-NormalizedDiskIOPSScore' {
    It 'returns null for null inputs' {
        Get-NormalizedDiskIOPSScore -ReadIOPS $null -WriteIOPS $null | Should -Be $null
    }

    It 'scores average of both inputs' {
        $score = Get-NormalizedDiskIOPSScore -ReadIOPS 2000 -WriteIOPS 2000
        $score | Should -Be 100
    }

    It 'clamps to minimum of 5' {
        $score = Get-NormalizedDiskIOPSScore -ReadIOPS 1 -WriteIOPS 1
        $score | Should -BeGreaterOrEqual 5
    }

    It 'clamps to maximum of 150' {
        $score = Get-NormalizedDiskIOPSScore -ReadIOPS 10000 -WriteIOPS 10000
        $score | Should -BeLessOrEqual 150
    }
}

Describe 'Add-Finding' {
    It 'creates a finding object with correct properties' {
        $f = Add-Finding -Penalty 10 -Area 'CPU' -Severity 'Alto' -Message 'Test'
        $f.Area | Should -Be 'CPU'
        $f.Severidad | Should -Be 'Alto'
        $f.Penalidad | Should -Be 10
        $f.Mensaje | Should -Be 'Test'
    }
}

Describe 'Get-Score' {
    It 'returns perfect score for ideal conditions' {
        $reposo = [PSCustomObject]@{ CPUAveragePct = 5; RAMAveragePct = 30; TempAverageC = 35 }
        $benchmark = [PSCustomObject]@{
            Phase = [PSCustomObject]@{ TempMaxC = 55 }
            CpuScore = 90; CpuOpsPerSecond = 5000000
            DiskWriteMBs = 500; Disk4K = $null
        }
        $disks = [PSCustomObject]@{ Volumes = @(); PhysicalDisk = @() }
        $diskAdvanced = $null
        $smart = @()
        $startup = @()
        $events = [PSCustomObject]@{ TotalEventos = 0 }
        $security = [PSCustomObject]@{ Defender = [PSCustomObject]@{ RealTimeProtection = $true } }
        $ramBenchmark = [PSCustomObject]@{ RamScore = 90; AvgMBs = 20000 }
        $network = @()

        $result = Get-Score -Reposo $reposo -Benchmark $benchmark -Disks $disks -DiskAdvanced $diskAdvanced -Smart $smart -Startup $startup -Events $events -Security $security -RamBenchmark $ramBenchmark -Network $network
        $result.Score | Should -Be 100
        $result.Estado | Should -Be 'Excelente'
    }

    It 'penalizes high CPU in idle' {
        $reposo = [PSCustomObject]@{ CPUAveragePct = 50; RAMAveragePct = 30; TempAverageC = 35 }
        $benchmark = [PSCustomObject]@{ Phase = [PSCustomObject]@{ TempMaxC = 55 }; CpuScore = 90; CpuOpsPerSecond = 5000000; DiskWriteMBs = 500; Disk4K = $null }
        $disks = [PSCustomObject]@{ Volumes = @(); PhysicalDisk = @() }
        $rest = @{ DiskAdvanced = $null; Smart = @(); Startup = @(); Events = [PSCustomObject]@{ TotalEventos = 0 }; Security = [PSCustomObject]@{ Defender = [PSCustomObject]@{ RealTimeProtection = $true } }; RamBenchmark = $null; Network = @() }

        $result = Get-Score -Reposo $reposo -Benchmark $benchmark -Disks $disks @rest
        $result.Score | Should -Be 90
        $result.Hallazgos.Count | Should -BeGreaterThan 0
    }

    It 'penalizes high temperature under load' {
        $reposo = [PSCustomObject]@{ CPUAveragePct = 5; RAMAveragePct = 30; TempAverageC = 35 }
        $benchmark = [PSCustomObject]@{ Phase = [PSCustomObject]@{ TempMaxC = 97 }; CpuScore = 90; CpuOpsPerSecond = 5000000; DiskWriteMBs = 500; Disk4K = $null }
        $disks = [PSCustomObject]@{ Volumes = @(); PhysicalDisk = @() }
        $rest = @{ DiskAdvanced = $null; Smart = @(); Startup = @(); Events = [PSCustomObject]@{ TotalEventos = 0 }; Security = [PSCustomObject]@{ Defender = [PSCustomObject]@{ RealTimeProtection = $true } }; RamBenchmark = $null; Network = @() }

        $result = Get-Score -Reposo $reposo -Benchmark $benchmark -Disks $disks @rest
        $result.Score | Should -Be 80
    }

    It 'penalizes low disk space' {
        $reposo = [PSCustomObject]@{ CPUAveragePct = 5; RAMAveragePct = 30; TempAverageC = 35 }
        $benchmark = [PSCustomObject]@{ Phase = [PSCustomObject]@{ TempMaxC = 55 }; CpuScore = 90; CpuOpsPerSecond = 5000000; DiskWriteMBs = 500; Disk4K = $null }
        $disks = [PSCustomObject]@{
            Volumes = @([PSCustomObject]@{ Unidad = 'C:'; LibrePct = 8 })
            PhysicalDisk = @()
        }
        $rest = @{ DiskAdvanced = $null; Smart = @(); Startup = @(); Events = [PSCustomObject]@{ TotalEventos = 0 }; Security = [PSCustomObject]@{ Defender = [PSCustomObject]@{ RealTimeProtection = $true } }; RamBenchmark = $null; Network = @() }

        $result = Get-Score -Reposo $reposo -Benchmark $benchmark -Disks $disks @rest
        $result.Score | Should -Be 90
    }

    It 'returns Critico when score is low' {
        $reposo = [PSCustomObject]@{ CPUAveragePct = 5; RAMAveragePct = 95; TempAverageC = 35 }
        $benchmark = [PSCustomObject]@{ Phase = [PSCustomObject]@{ TempMaxC = 97 }; CpuScore = 20; CpuOpsPerSecond = 1000; DiskWriteMBs = 10; Disk4K = [PSCustomObject]@{ IOScore = 20; ReadIOPS = 100; WriteIOPS = 50 } }
        $disks = [PSCustomObject]@{
            Volumes = @([PSCustomObject]@{ Unidad = 'C:'; LibrePct = 5 })
            PhysicalDisk = @([PSCustomObject]@{ HealthStatus = 'Unhealthy'; Modelo = 'DiskX'; PhysicalMediaType = 'HDD' })
        }
        $events = [PSCustomObject]@{ TotalEventos = 30 }
        $rest = @{ DiskAdvanced = $null; Smart = @(); Startup = @(1..30); Security = [PSCustomObject]@{ Defender = $null }; RamBenchmark = $null; Network = $null }

        $result = Get-Score -Reposo $reposo -Benchmark $benchmark -Disks $disks -Events $events @rest
        $result.Estado | Should -Be 'Critico'
    }

    It 'handles null benchmark phase' {
        $reposo = [PSCustomObject]@{ CPUAveragePct = $null; RAMAveragePct = $null; TempAverageC = $null }
        $benchmark = [PSCustomObject]@{ Phase = $null; CpuScore = $null; CpuOpsPerSecond = $null; DiskWriteMBs = $null; Disk4K = $null }
        $disks = [PSCustomObject]@{ Volumes = @(); PhysicalDisk = @() }
        $rest = @{ DiskAdvanced = $null; Smart = @(); Startup = @(); Events = [PSCustomObject]@{ TotalEventos = 0 }; Security = [PSCustomObject]@{ Defender = $null }; RamBenchmark = $null; Network = $null }

        $result = Get-Score -Reposo $reposo -Benchmark $benchmark -Disks $disks @rest
        $result.Score | Should -Be 92
    }

    It 'score never goes below 0' {
        $reposo = [PSCustomObject]@{ CPUAveragePct = 100; RAMAveragePct = 100; TempAverageC = 100 }
        $benchmark = [PSCustomObject]@{ Phase = [PSCustomObject]@{ TempMaxC = 100 }; CpuScore = 0; CpuOpsPerSecond = 0; DiskWriteMBs = 0; Disk4K = $null }
        $disks = [PSCustomObject]@{ Volumes = @(); PhysicalDisk = @() }
        $rest = @{ DiskAdvanced = $null; Smart = @(); Startup = @(); Events = [PSCustomObject]@{ TotalEventos = 0 }; Security = [PSCustomObject]@{ Defender = $null }; RamBenchmark = $null; Network = $null }

        $result = Get-Score -Reposo $reposo -Benchmark $benchmark -Disks $disks @rest
        $result.Score | Should -BeGreaterOrEqual 0
    }
}

Describe 'Get-Recommendations' {
    It 'returns generic recommendation when no issues found' {
        $reposo = [PSCustomObject]@{ CPUAveragePct = 5; RAMAveragePct = 30 }
        $benchmark = [PSCustomObject]@{ Phase = [PSCustomObject]@{ TempMaxC = 60 }; Disk4K = $null }
        $disks = [PSCustomObject]@{ Volumes = @(); PhysicalDisk = @() }
        $recs = Get-Recommendations -Reposo $reposo -Benchmark $benchmark -Disks $disks -Smart @() -Startup @() -Events ([PSCustomObject]@{ TotalEventos = 0 }) -RamBenchmark $null -Network $null
        $recs.Count | Should -Be 1
        $recs[0] | Should -Match 'No se detectan problemas'
    }

    It 'recommends checking CPU processes when idle is high' {
        $reposo = [PSCustomObject]@{ CPUAveragePct = 50; RAMAveragePct = 30 }
        $benchmark = [PSCustomObject]@{ Phase = [PSCustomObject]@{ TempMaxC = 60 }; Disk4K = $null }
        $disks = [PSCustomObject]@{ Volumes = @(); PhysicalDisk = @() }
        $recs = Get-Recommendations -Reposo $reposo -Benchmark $benchmark -Disks $disks -Smart @() -Startup @() -Events ([PSCustomObject]@{ TotalEventos = 0 }) -RamBenchmark $null -Network $null
        $recs | Should -ContainMatch 'CPU en reposo elevada'
    }

    It 'recommends RAM review when usage is high' {
        $reposo = [PSCustomObject]@{ CPUAveragePct = 5; RAMAveragePct = 85 }
        $benchmark = [PSCustomObject]@{ Phase = [PSCustomObject]@{ TempMaxC = 60 }; Disk4K = $null }
        $disks = [PSCustomObject]@{ Volumes = @(); PhysicalDisk = @() }
        $recs = Get-Recommendations -Reposo $reposo -Benchmark $benchmark -Disks $disks -Smart @() -Startup @() -Events ([PSCustomObject]@{ TotalEventos = 0 }) -RamBenchmark $null -Network $null
        $recs | Should -ContainMatch 'uso en reposo'
    }
}
