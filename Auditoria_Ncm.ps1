Import-Module ImportExcel
#Instalar: Install-Module ImportExcel -Scope CurrentUser
# Configurações de caminho
$caminhoArquivo = "C:\caminho\nomedoarquivo.xlsx"
$caminhoSaida = "C:\caminho\Relatorio_Detalhado_NCM.xlsx"

$colCodigo    = "Código"
$colDescricao = "Descrição"
$colNCM       = "Ncm"

if (-not (Test-Path $caminhoArquivo)) {
    Write-Host "⚠️ Arquivo não encontrado: $caminhoArquivo" -ForegroundColor Yellow
    return
}

$dados = Import-Excel $caminhoArquivo
$listaFinal = @()

Write-Host "--- Iniciando Auditoria Detalhada ---" -ForegroundColor Cyan

foreach ($linha in $dados) {
    # Captura os dados originais das linhas
    $codigoSis = $linha.$colCodigo
    $descSis   = $linha.$colDescricao
    $ncmBruto  = $linha.$colNCM
    
    $ncmLimpo = $ncmBruto -replace '[^0-9]', ''
    
    if ([string]::IsNullOrWhiteSpace($ncmLimpo)) { continue }

    # Objeto que mantém os dados originais e adiciona a resposta da API
    $resultado = [PSCustomObject]@{
        Codigo_Sistema    = $codigoSis
        Descricao_Sistema = $descSis
        NCM_Informado     = $ncmBruto
        NCM_Novo_A_Informar = ''
        NCM_Status        = "❌ INVÁLIDO/VENCIDO"
        Descricao_Oficial = "Não encontrado na base da Receita"
    }

    try {
        $url = "https://brasilapi.com.br/api/ncm/v1/$ncmLimpo"
        $api = Invoke-RestMethod -Uri $url -Method Get -ErrorAction Stop
        
        $resultado.NCM_Status        = "✅ ATIVO"
        $resultado.Descricao_Oficial = $api.descricao
        
        Write-Host "[OK] Item $codigoSis verificado." -ForegroundColor Green
    }
    catch {
        Write-Host "[!] Item $codigoSis NCM inválido." -ForegroundColor Red
    }

    $listaFinal += $resultado
    Start-Sleep -Milliseconds 300
}

# Salva o relatório completo
$listaFinal | Export-Excel -Path $caminhoSaida -AutoSize -Title "Auditoria de Cadastro NCM"

Write-Host "`n✅ Relatório Detalhado gerado: $caminhoSaida" -ForegroundColor Blue -BackgroundColor White