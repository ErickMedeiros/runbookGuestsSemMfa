# Implantação

## Configuração dos módulos

Você pode importar manualmente os módulos necessários no Azure Automation executando o script abaixo em uma sessão PowerShell com acesso à sua assinatura.

> Substitua os valores de exemplo pelos nomes do seu ambiente antes de executar.

```powershell
# --- Parâmetros ---
$resourceGroupName     = "<SEU_RESOURCE_GROUP>"
$automationAccountName = "<SUA_AUTOMATION_ACCOUNT>"
$runtimeVersion        = "7.2"   # ou "5.1"

# Conecta (use Connect-AzAccount -Identity se rodar de dentro de outro runbook)
# Connect-AzAccount

# Função auxiliar para importar e aguardar conclusão
function Import-AAModule {
    param($Name, $Wait = $true)

    $uri = "https://www.powershellgallery.com/api/v2/package/$Name"
    Write-Output "Importando $Name ..."

    New-AzAutomationModule `
        -ResourceGroupName     $resourceGroupName `
        -AutomationAccountName $automationAccountName `
        -Name                  $Name `
        -ContentLinkUri        $uri `
        -RuntimeVersion        $runtimeVersion | Out-Null

    if ($Wait) {
        do {
            Start-Sleep -Seconds 20
            $m = Get-AzAutomationModule `
                    -ResourceGroupName     $resourceGroupName `
                    -AutomationAccountName $automationAccountName `
                    -Name                  $Name `
                    -RuntimeVersion        $runtimeVersion
            Write-Output "  $Name -> $($m.ProvisioningState)"
        } while ($m.ProvisioningState -notin @("Succeeded","Failed"))

        if ($m.ProvisioningState -eq "Failed") {
            throw "Falha ao importar $Name"
        }
    }
}

# 1) Dependência base — ESPERAR terminar antes de seguir
Import-AAModule -Name "Microsoft.Graph.Authentication" -Wait $true

# 2-5) Módulos que dependem do Authentication (podem ir em sequência)
Import-AAModule -Name "Microsoft.Graph.Users"
Import-AAModule -Name "Microsoft.Graph.Groups"
Import-AAModule -Name "Microsoft.Graph.Reports"
Import-AAModule -Name "Microsoft.Graph.Users.Actions"

# 6) Opcional — independente dos demais
Import-AAModule -Name "ImportExcel"

Write-Output "Concluído."
```
