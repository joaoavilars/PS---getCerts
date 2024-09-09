$CurrentDate = Get-Date
$WarningThreshold = 366

$Certs = Get-Website | Where-Object {$_.serverAutoStart -eq $True} | Get-WebBinding -Protocol https | Select-Object -ExpandProperty certificateHash

$MyCerts = Get-ChildItem Cert:\LocalMachine\My | Select-Object NotAfter, Subject, Thumbprint

$CertAlreadyPrinted = $false

foreach ($Cert in $Certs) {
    $CertImLookingFor = $MyCerts | Where-Object {$_.Thumbprint -eq $Cert}
    [datetime]$CertExp = $CertImLookingFor.NotAfter
    $TimeSpan = New-TimeSpan -Start $CurrentDate -End $CertExp

    if (!($CertImLookingFor.Subject -like "*.local") -and $TimeSpan.Days -le $WarningThreshold -and $TimeSpan.Days -ge 0 -and !$CertAlreadyPrinted) {
        $DaysRemaining = $TimeSpan.Days
        Write-Output $DaysRemaining
        $CertAlreadyPrinted = $true
    }
}