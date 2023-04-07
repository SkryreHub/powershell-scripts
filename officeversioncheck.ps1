function Get-OfficeVersion {
    $officePath = 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'
    if (Test-Path $officePath) {
        $version = (Get-ItemProperty -Path $officePath).ClientVersionToReport
        if ($version) {
            return "Microsoft 365 ($version)"
        }
    }
    return "Unknown Office Version"
}

$officeApps = @(
    @{
        Name = "Word"
        ExeName = "winword.exe"
    },
    @{
        Name = "Excel"
        ExeName = "excel.exe"
    },
    @{
        Name = "PowerPoint"
        ExeName = "powerpnt.exe"
    },
    @{
        Name = "Outlook"
        ExeName = "outlook.exe"
    },
    @{
        Name = "OneNote"
        ExeName = "onenote.exe"
    },
    @{
        Name = "Access"
        ExeName = "msaccess.exe"
    },
    @{
        Name = "Publisher"
        ExeName = "mspub.exe"
    }
)

$programFileDirs = @($env:ProgramFiles, ${env:ProgramFiles(x86)})
$office365Paths = @("Microsoft Office\root\Office16", "Microsoft Office\Office16")

$installedOfficeApps = @()

foreach ($app in $officeApps) {
    foreach ($programFileDir in $programFileDirs) {
        foreach ($office365Path in $office365Paths) {
            $partialPath = Join-Path -Path $programFileDir -ChildPath $office365Path
            $path = Join-Path -Path $partialPath -ChildPath $app.ExeName
            $appPath = Get-ChildItem -Path $path -ErrorAction SilentlyContinue
            if ($appPath) {
                $version = Get-OfficeVersion
                $installedOfficeApps += @{
                    Name = $app.Name
                    Version = $version
                }
                break
            }
        }
    }
}

if ($installedOfficeApps.Count -gt 0) {
    Write-Output "Microsoft Office applications installed on this system:"
    $installedOfficeApps | Format-Table -AutoSize
} else {
    Write-Output "No Microsoft Office applications found on this system."
}