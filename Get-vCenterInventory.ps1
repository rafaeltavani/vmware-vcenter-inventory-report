<#
.SYNOPSIS
    Gera um inventário de um ambiente VMware vCenter utilizando VMware PowerCLI.

.DESCRIPTION
    Coleta informações de hosts ESXi, máquinas virtuais, datastores, redes e snapshots.
    Os dados são exportados em arquivos CSV e, se o módulo ImportExcel estiver instalado,
    também em um único arquivo XLSX.

    Este script foi preparado para uso público em portfólio:
    - Não contém credenciais.
    - Não contém IPs, nomes de servidores ou domínios corporativos.
    - Solicita credenciais em tempo de execução.
    - Recebe o endereço do vCenter por parâmetro.

.PARAMETER VCenterServer
    Nome DNS ou endereço IP do VMware vCenter.

.PARAMETER OutputPath
    Diretório onde os relatórios serão gravados.
    Padrão: uma pasta "Relatorios-vCenter" dentro do diretório atual.

.PARAMETER IgnoreInvalidCertificate
    Quando informado, configura o PowerCLI para ignorar certificados inválidos
    durante a sessão atual.

.EXAMPLE
    .\Get-vCenterInventory.ps1 -VCenterServer "vcenter.exemplo.local"

.EXAMPLE
    .\Get-vCenterInventory.ps1 `
        -VCenterServer "192.0.2.10" `
        -OutputPath "C:\Relatorios" `
        -IgnoreInvalidCertificate

.NOTES
    Requisitos:
    - PowerShell 7+ ou Windows PowerShell 5.1
    - VMware.PowerCLI
    - Opcional: ImportExcel para geração do XLSX

    Instalação do PowerCLI:
        Install-Module VMware.PowerCLI -Scope CurrentUser

    Instalação opcional do ImportExcel:
        Install-Module ImportExcel -Scope CurrentUser
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$VCenterServer,

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$OutputPath = (Join-Path (Get-Location) "Relatorios-vCenter"),

    [Parameter(Mandatory = $false)]
    [switch]$IgnoreInvalidCertificate
)

$ErrorActionPreference = "Stop"

function Write-Step {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    Write-Host ""
    Write-Host "==> $Message"
}

function Convert-ToGB {
    param(
        [Parameter(Mandatory = $false)]
        [double]$Value
    )

    if ($null -eq $Value) {
        return 0
    }

    return [math]::Round($Value, 2)
}

try {
    Write-Step "Validando módulos"

    if (-not (Get-Module -ListAvailable -Name VMware.PowerCLI)) {
        throw "O módulo VMware.PowerCLI não está instalado. Execute: Install-Module VMware.PowerCLI -Scope CurrentUser"
    }

    Import-Module VMware.PowerCLI -ErrorAction Stop

    if ($IgnoreInvalidCertificate) {
        Set-PowerCLIConfiguration `
            -InvalidCertificateAction Ignore `
            -Scope Session `
            -Confirm:$false | Out-Null
    }

    if (-not (Test-Path -Path $OutputPath)) {
        New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
    }

    $timestamp = Get-Date -Format "yyyy-MM-dd_HHmm"
    $reportFolder = Join-Path $OutputPath "Inventario-vCenter_$timestamp"
    New-Item -ItemType Directory -Path $reportFolder -Force | Out-Null

    Write-Step "Solicitando credenciais"
    $credential = Get-Credential -Message "Informe as credenciais para acessar o vCenter $VCenterServer"

    Write-Step "Conectando ao vCenter $VCenterServer"
    Connect-VIServer `
        -Server $VCenterServer `
        -Credential $credential `
        -ErrorAction Stop | Out-Null

    Write-Step "Coletando hosts ESXi"
    $hosts = Get-VMHost | Sort-Object Name | ForEach-Object {
        [pscustomobject]@{
            HostName          = $_.Name
            ConnectionState   = $_.ConnectionState
            PowerState        = $_.PowerState
            Manufacturer      = $_.Manufacturer
            Model             = $_.Model
            ESXiVersion       = $_.Version
            ESXiBuild         = $_.Build
            CpuPackages       = $_.NumCpu
            CpuCores          = $_.NumCpu
            CpuUsageMhz       = [math]::Round($_.CpuUsageMhz, 2)
            CpuTotalMhz       = [math]::Round($_.CpuTotalMhz, 2)
            MemoryUsageGB     = Convert-ToGB ($_.MemoryUsageGB)
            MemoryTotalGB     = Convert-ToGB ($_.MemoryTotalGB)
            Parent            = $_.Parent.Name
        }
    }

    Write-Step "Coletando máquinas virtuais"
    $vms = Get-VM | Sort-Object Name | ForEach-Object {
        $vm = $_

        $guestIp = $null
        if ($vm.Guest -and $vm.Guest.IPAddress) {
            $guestIp = ($vm.Guest.IPAddress | Where-Object { $_ -and $_ -notmatch ":" } | Select-Object -First 1)
        }

        $networks = @()
        try {
            $networks = Get-NetworkAdapter -VM $vm -ErrorAction Stop |
                Select-Object -ExpandProperty NetworkName -Unique
        } catch {
            $networks = @()
        }

        $datastores = @()
        try {
            $datastores = Get-Datastore -VM $vm -ErrorAction Stop |
                Select-Object -ExpandProperty Name -Unique
        } catch {
            $datastores = @()
        }

        [pscustomobject]@{
            VMName            = $vm.Name
            PowerState        = $vm.PowerState
            NumCpu            = $vm.NumCpu
            MemoryGB          = Convert-ToGB ($vm.MemoryGB)
            UsedSpaceGB       = Convert-ToGB ($vm.UsedSpaceGB)
            ProvisionedGB     = Convert-ToGB ($vm.ProvisionedSpaceGB)
            GuestOS           = $vm.Guest.OSFullName
            IPAddress         = $guestIp
            VMHost            = $vm.VMHost.Name
            Folder            = $vm.Folder.Name
            Datastores        = ($datastores -join ", ")
            Networks          = ($networks -join ", ")
            ToolsStatus       = $vm.ExtensionData.Guest.ToolsStatus
            ToolsVersion      = $vm.ExtensionData.Guest.ToolsVersion
        }
    }

    Write-Step "Coletando discos virtuais"
    $virtualDisks = Get-VM | Sort-Object Name | ForEach-Object {
        $vm = $_

        Get-HardDisk -VM $vm | ForEach-Object {
            [pscustomobject]@{
                VMName        = $vm.Name
                DiskName      = $_.Name
                CapacityGB    = Convert-ToGB ($_.CapacityGB)
                StorageFormat = $_.StorageFormat
                Persistence   = $_.Persistence
                Filename      = $_.Filename
            }
        }
    }

    Write-Step "Coletando datastores"
    $datastores = Get-Datastore | Sort-Object Name | ForEach-Object {
        [pscustomobject]@{
            DatastoreName = $_.Name
            Type          = $_.Type
            CapacityGB    = Convert-ToGB ($_.CapacityGB)
            FreeSpaceGB   = Convert-ToGB ($_.FreeSpaceGB)
            UsedSpaceGB   = [math]::Round(($_.CapacityGB - $_.FreeSpaceGB), 2)
            PercentFree   = if ($_.CapacityGB -gt 0) {
                [math]::Round(($_.FreeSpaceGB / $_.CapacityGB) * 100, 2)
            } else {
                0
            }
        }
    }

    Write-Step "Coletando redes"
    $networks = Get-VirtualPortGroup | Sort-Object Name | ForEach-Object {
        [pscustomobject]@{
            PortGroupName = $_.Name
            VirtualSwitch = $_.VirtualSwitchName
            VLANId        = $_.VLanId
            VMHost        = $_.VMHost.Name
        }
    }

    Write-Step "Coletando snapshots"
    $snapshots = Get-VM | Sort-Object Name | ForEach-Object {
        $vm = $_

        Get-Snapshot -VM $vm -ErrorAction SilentlyContinue | ForEach-Object {
            [pscustomobject]@{
                VMName      = $vm.Name
                Snapshot    = $_.Name
                Description = $_.Description
                Created     = $_.Created
                SizeGB      = Convert-ToGB ($_.SizeGB)
            }
        }
    }

    Write-Step "Gerando resumo executivo"
    $summary = [pscustomobject]@{
        VCenter              = $VCenterServer
        GeneratedAt          = Get-Date
        TotalHosts           = @($hosts).Count
        TotalVMs             = @($vms).Count
        PoweredOnVMs         = @($vms | Where-Object PowerState -eq "PoweredOn").Count
        PoweredOffVMs        = @($vms | Where-Object PowerState -eq "PoweredOff").Count
        TotalDatastores      = @($datastores).Count
        TotalSnapshots       = @($snapshots).Count
        TotalProvisionedGB   = [math]::Round((($vms | Measure-Object ProvisionedGB -Sum).Sum), 2)
        TotalVMUsedSpaceGB   = [math]::Round((($vms | Measure-Object UsedSpaceGB -Sum).Sum), 2)
    }

    Write-Step "Exportando arquivos CSV"
    $summary      | Export-Csv -Path (Join-Path $reportFolder "Resumo.csv")       -NoTypeInformation -Encoding UTF8
    $hosts        | Export-Csv -Path (Join-Path $reportFolder "Hosts.csv")        -NoTypeInformation -Encoding UTF8
    $vms          | Export-Csv -Path (Join-Path $reportFolder "VMs.csv")          -NoTypeInformation -Encoding UTF8
    $virtualDisks | Export-Csv -Path (Join-Path $reportFolder "Discos.csv")       -NoTypeInformation -Encoding UTF8
    $datastores   | Export-Csv -Path (Join-Path $reportFolder "Datastores.csv")   -NoTypeInformation -Encoding UTF8
    $networks     | Export-Csv -Path (Join-Path $reportFolder "Redes.csv")        -NoTypeInformation -Encoding UTF8
    $snapshots    | Export-Csv -Path (Join-Path $reportFolder "Snapshots.csv")    -NoTypeInformation -Encoding UTF8

    $excelAvailable = Get-Module -ListAvailable -Name ImportExcel

    if ($excelAvailable) {
        Write-Step "Gerando arquivo Excel"

        Import-Module ImportExcel -ErrorAction Stop

        $excelFile = Join-Path $reportFolder "Inventario-vCenter_$timestamp.xlsx"

        $summary      | Export-Excel -Path $excelFile -WorksheetName "Resumo"      -AutoSize -FreezeTopRow
        $hosts        | Export-Excel -Path $excelFile -WorksheetName "Hosts"       -AutoSize -FreezeTopRow
        $vms          | Export-Excel -Path $excelFile -WorksheetName "VMs"         -AutoSize -FreezeTopRow
        $virtualDisks | Export-Excel -Path $excelFile -WorksheetName "Discos"      -AutoSize -FreezeTopRow
        $datastores   | Export-Excel -Path $excelFile -WorksheetName "Datastores"  -AutoSize -FreezeTopRow
        $networks     | Export-Excel -Path $excelFile -WorksheetName "Redes"       -AutoSize -FreezeTopRow
        $snapshots    | Export-Excel -Path $excelFile -WorksheetName "Snapshots"   -AutoSize -FreezeTopRow

        Write-Host ""
        Write-Host "Relatório Excel gerado: $excelFile"
    }
    else {
        Write-Host ""
        Write-Host "Módulo ImportExcel não encontrado."
        Write-Host "Os relatórios CSV foram gerados normalmente."
        Write-Host "Para gerar XLSX, instale: Install-Module ImportExcel -Scope CurrentUser"
    }

    Write-Host ""
    Write-Host "Inventário concluído com sucesso."
    Write-Host "Diretório dos relatórios: $reportFolder"
}
catch {
    Write-Error "Falha durante a execução do inventário: $($_.Exception.Message)"
    exit 1
}
finally {
    $connectedServers = $global:DefaultVIServers

    if ($connectedServers) {
        Write-Step "Desconectando do vCenter"
        Disconnect-VIServer -Server $connectedServers -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
    }
}
