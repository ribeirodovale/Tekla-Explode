param(
    [Parameter(Mandatory = $true)]
    [string]$WorkspaceFolder
)

$ErrorActionPreference = "Stop"

$mutexName = "Global\EZVistaExplodidaBuildRun"
$mutex = New-Object System.Threading.Mutex($false, $mutexName)
$lockTaken = $false

try {
    $lockTaken = $mutex.WaitOne(0)
    if (-not $lockTaken) {
        Write-Host "Task ja esta em execucao. Ignorando disparo duplicado."
        exit 0
    }

    $processName = "WindowsFormsApp1"
    $runningProcesses = Get-Process -Name $processName -ErrorAction SilentlyContinue
    if ($runningProcesses) {
        $runningProcesses | Stop-Process -Force -ErrorAction SilentlyContinue

        $deadline = (Get-Date).AddSeconds(10)
        while (Get-Process -Name $processName -ErrorAction SilentlyContinue) {
            if ((Get-Date) -gt $deadline) {
                throw "Nao foi possivel fechar o processo $processName antes do build."
            }

            Start-Sleep -Milliseconds 200
        }
    }

    $slnCandidates = @(
        (Join-Path $WorkspaceFolder "WindowsFormsApp1\\WindowsFormsApp1.slnx"),
        (Join-Path $WorkspaceFolder "WindowsFormsApp1.slnx")
    )

    $slnPath = $slnCandidates | Where-Object { Test-Path $_ } | Select-Object -First 1
    if (-not $slnPath) {
        throw "Arquivo .slnx nao encontrado no workspace: $WorkspaceFolder"
    }

    dotnet build $slnPath -c Debug /p:RunPostBuildEvent=Never
    if ($LASTEXITCODE -ne 0) {
        exit $LASTEXITCODE
    }

    $exeCandidates = @(
        (Join-Path $WorkspaceFolder "WindowsFormsApp1\\WindowsFormsApp1\\bin\\Debug\\WindowsFormsApp1.exe"),
        (Join-Path $WorkspaceFolder "WindowsFormsApp1\\bin\\Debug\\WindowsFormsApp1.exe")
    )

    $exePath = $exeCandidates | Where-Object { Test-Path $_ } | Select-Object -First 1
    if (-not $exePath) {
        throw "Exe Debug nao encontrado."
    }

    Start-Process -FilePath $exePath
    Write-Host "Build Debug concluido e app iniciado: $exePath"
}
finally {
    if ($lockTaken) {
        $mutex.ReleaseMutex() | Out-Null
    }

    $mutex.Dispose()
}
