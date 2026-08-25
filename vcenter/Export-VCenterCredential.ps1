<#
.SYNOPSIS
    Prompt for vCenter credentials and save them as CLIXML for reboot_uag.py.

.DESCRIPTION
    Uses Get-Credential and Export-Clixml. The password is encrypted with Windows
    DPAPI for the current user on this machine only.

    Type the username as DOMAIN\username or username@domain.com.
    The Python script converts DOMAIN\username to username@<UPN_SUFFIX>.

.EXAMPLE
    .\Export-VCenterCredential.ps1
    .\Export-VCenterCredential.ps1 -Path .\vcenter_credential.xml
#>
[CmdletBinding()]
param(
    [Parameter()]
    [string]$Path = (Join-Path $PSScriptRoot "vcenter_credential.xml")
)

$cred = Get-Credential -Message "vCenter login (DOMAIN\user is stored as-is; Python converts to user@domain.com)"
if (-not $cred) {
    Write-Error "No credential was provided."
    exit 1
}

$directory = Split-Path -Parent $Path
if ($directory -and -not (Test-Path $directory)) {
    New-Item -ItemType Directory -Path $directory | Out-Null
}

$cred | Export-Clixml -Path $Path
Write-Host "Saved credential for '$($cred.UserName)' to $Path"
Write-Host "Re-run this script if you change Windows users or move the file to another computer."
