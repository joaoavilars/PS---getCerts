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

        # Extrair o nome da empresa do campo CN
        $Subject = $Cert.GetNameInfo([System.Security.Cryptography.X509Certificates.X509NameType]::SimpleName, $false)
        
        # Remover números e dois pontos do nome da empresa
        $SubjectClean = $Subject -replace '\d+', '' -replace '\s+', ' ' -replace '^\s+|\s+$', '' -replace ':+$', ''
        
        # Inicializar o CNPJ como null
        $CNPJ = $null

        # Tentar extrair o CNPJ da extensão específica do certificado
        $CNPJExtension = $Cert.Extensions | Where-Object { $_.Oid.Value -eq "2.16.76.1.3.3" }
        if ($CNPJExtension) {
            $CNPJ = $CNPJExtension.Format($false) -replace '[^\d]', ''
            Write-Host "CNPJ encontrado na extensão: $CNPJ para $SubjectClean"
        }

        # Se não encontrar na extensão, procurar no campo Subject inteiro o primeiro número de 14 dígitos
        if (-not $CNPJ) {
            if ($Cert.Subject -match '(\d{14})') {
                $CNPJ = $Matches[1]
                Write-Host "CNPJ encontrado no Subject: $CNPJ para $SubjectClean"
            }
        }

        # Se ainda não encontrar, exibir mensagem de aviso
        if (-not $CNPJ) {
            Write-Host "Nenhum CNPJ válido encontrado para: $SubjectClean"
        }

        # Criar um objeto PS personalizado
        $result = [PSCustomObject]@{
            enterprise    = $SubjectClean
            expire_date   = $ExpirationDate
            days          = $DaysRemaining
            cnpj          = $CNPJ
        }

        # Adicionar o objeto à lista de resultados
        $results += $result
    }
}

# Remover entradas duplicadas com base no CNPJ
$results = $results | Where-Object { $_.cnpj } | Group-Object cnpj | ForEach-Object { $_.Group | Select-Object -First 1 }

# Converter o array de objetos para JSON e salvar no arquivo de saída
$results | ConvertTo-Json -Depth 5 | Out-File -FilePath $OutputFile -Encoding utf8
