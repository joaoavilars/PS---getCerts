param(
    [string]$OutputFile
)

$CurrentDate = Get-Date
$WarningThreshold = 366
$results = @()

$MyCerts = Get-ChildItem Cert:\CurrentUser\My | Where-Object { $_.Issuer -ne $_.Subject }

foreach ($Cert in $MyCerts) {
    [datetime]$CertExp = $Cert.NotAfter
    $TimeSpan = New-TimeSpan -Start $CurrentDate -End $CertExp

    if ($TimeSpan.Days -le $WarningThreshold -and $TimeSpan.Days -ge 0) {
        $DaysRemaining = $TimeSpan.Days
        $ExpirationDate = $CertExp.ToShortDateString()
        $Subject = $Cert.Subject
        $Handle = $Cert.Handle  # Substitua 'Handle' pela propriedade correta se necessário

        # Extrair o nome da empresa (CN)
        $CN = if ($Subject -match 'CN=([^,]+)') { $matches[1] } else { "Unknown" }

        # Criar um objeto PS personalizado
        $result = [PSCustomObject]@{
            enterprise    = $CN
            expire_date   = $ExpirationDate
            days          = $DaysRemaining
            handle        = $Handle  # Adiciona o dado do 'Handle'
        }

        # Adicionar o objeto à lista de resultados
        $results += $result
    }
}

# Converter o array de objetos para JSON e salvar no arquivo de saída
$results | ConvertTo-Json | Out-File -FilePath $OutputFile -Encoding utf8
