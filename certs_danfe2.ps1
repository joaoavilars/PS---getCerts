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
        $Subject = $Cert.Subject -replace '^CN=(.*?),.*', '$1'
        
        # Extrair o CNPJ do campo DnsNameList
        $DnsNameList = $Cert.DnsNameList
        $CNPJ = $null
        foreach ($DnsName in $Cert.DnsNameList) {
            if ($DnsName -match '.*:(\d{14})$') {
                $CNPJ = [int64]$matches[1]  # Convertendo o CNPJ para número inteiro de 64 bits
                break  # Parar o loop quando o CNPJ for encontrado
            }
        }

        # Criar um objeto PS personalizado
        $result = [PSCustomObject]@{
            enterprise    = "${Subject}:${CNPJ}"
            expire_date   = $ExpirationDate
            days          = $DaysRemaining
            cnpj          = $CNPJ
        }

        # Adicionar o objeto à lista de resultados
        $results += $result
    }
}

# Converter o array de objetos para JSON e salvar no arquivo de saída
$results | ConvertTo-Json | Out-File -FilePath $OutputFile -Encoding utf8
