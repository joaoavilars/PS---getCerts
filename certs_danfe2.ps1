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

        # Exibir mensagem de depuração com o nome da empresa
        Write-Host "Empresa encontrada: $Subject"

        # Extrair todas as instâncias de OU no campo Subject
        $OUs = [regex]::Matches($Cert.Subject, 'OU=([^,]+)')

        # Exibir todos os OUs encontrados
        Write-Host "OUs encontrados: $($OUs | ForEach-Object { $_.Groups[1].Value })"

        $CNPJ = $null
        foreach ($OU in $OUs) {
            $OUValue = $OU.Groups[1].Value
            
            # Verificar se o valor encontrado é um CNPJ válido (14 dígitos)
            if ($OUValue -match '^\d{14}$') {
                $CNPJ = $OUValue
                Write-Host "CNPJ válido encontrado: $CNPJ"
                break  # Interromper a busca ao encontrar o CNPJ
            }
        }

        if ($CNPJ -eq $null) {
            Write-Host "Nenhum CNPJ válido encontrado para: $Subject"
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
