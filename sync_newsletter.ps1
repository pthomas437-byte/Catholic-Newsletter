# Pulls the latest Catholic Intelligence Newsletter issue from GitHub and
# copies the newest .docx into the iCloud Drive folder so it syncs to other devices.

$RepoDir = "C:\Users\pthom\Documents\Catholic-Newsletter"
$DestDir = "C:\Users\pthom\iCloudDrive\Catholic Newsletter"
$Git = "C:\Program Files\Git\cmd\git.exe"

Set-Location $RepoDir
& $Git pull

if (-not (Test-Path $DestDir)) {
    New-Item -ItemType Directory -Path $DestDir -Force | Out-Null
}

$Latest = Get-ChildItem -Path (Join-Path $RepoDir "issues") -Filter "Catholic_Intelligence_Newsletter_*.docx" |
    Sort-Object Name -Descending | Select-Object -First 1

if ($Latest) {
    Copy-Item -Path $Latest.FullName -Destination $DestDir -Force
    Write-Output "Synced $($Latest.Name) to $DestDir"
} else {
    Write-Output "No newsletter .docx found in $RepoDir\issues"
}
