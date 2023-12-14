# conversor.ps1 – converte arquivos .xlsb em .xlsx enquanto remove macros
# - pega um argumento da linha de comando (aceita caracteres curinga)
# – abre cada arquivo correspondente no Excel
# - desativa o salvamento automático, remove as macros
# – salva usando o mesmo nome + .xlsx
# - fecha o Excel
# Manipulação de erros:
# - se o parâmetro estiver faltando exibe ajuda
# - se o arquivo não terminar em .xlsb, ignora-o
# - se o arquivo não existir exibe informações

param ($in)

$processed = 0
Write-Host [ conversor ] Iniciando...

# Checka os parametros passados
if ($null -ne $in){
    try {
        # checa os parametros dos arquivos
        $all = Get-ChildItem $in -ErrorAction Stop
        # processa todos os arquivos encontrados
        foreach ($s_file in $all) 
        {
            # continua se encontrar os arquivos
            if ($s_file -like "*.xlsb")
            {
                $f_input = $s_file.FullName
                $loc = (Get-Item $s_file).DirectoryName
                $f_name = (Get-Item $s_file.FullName).BaseName + ".xlsx"
                $f_output = Join-Path $loc $f_name
                # monta os existentes
                if (Test-Path $s_file -PathType Leaf)
                {
                    Write-Host [ $f_name ] Convertendo...
                    $xlApp = New-Object -Com Excel.Application      # cria um objeto
                    $xlApp.Visible = $false                         # Previne caso uma janela de auda abra
                    $xlApp.DisplayAlerts = $false                   # Desabilita os alertas
                    $xlApp.Workbooks.Open($f_input) | Out-Null      # abre o arquivo    
                    # remove todas as macros
                    foreach ($v in $xlApp.ActiveWorkbook.VBProject.VBComponents) 
                    { 
                        if ($v.Type -eq 1) 
                        { 
                            $xlApp.ActiveWorkbook.VBProject.VBComponents.Remove($v)  
                        } 
                    }     
                    $xlApp.ActiveWorkbook.SaveAs($f_output, 51) | Out-Null 
                    $xlApp.Quit()
                    Write-Host [ $f_name ] Done.
                    $processed = $processed + 1                     # Acrescenta um numero nos arquivos
                } 
                else 
                {
                    Write-Host [Skipping] $f_input Não existe
                }
            } 
            else 
            {
                Write-Host [($s_file.Name)] NAO e .xlsb
            }
        }
    }
    catch [System.Management.Automation.ItemNotFoundException]
    {
        Write-Host [Skipping] Nenhum arquivo encontrado
    }
    # Mostra o numero de arquivos processados
    Write-Host [ conversor ] Finished: $processed Arquivos processados.
} 
# nenhum parametro passado
else 
{
    Write-Host [ conversor ] Error: Parametro perdido / Usage: conversor [filename.xlsb]
}