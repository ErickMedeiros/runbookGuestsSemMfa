# Runbook-GuestsSemMfa

Runbook PowerShell para Azure Automation que **detecta, bloqueia e reporta convidados (Guest) sem MFA** em um grupo específico do Microsoft Entra ID. A solução gera relatório em **Excel ou CSV** e envia o resultado por e-mail via **Microsoft Graph**.

[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue.svg)](https://docs.microsoft.com/en-us/powershell/)
[![Azure Automation](https://img.shields.io/badge/Azure-Automation-0089D6.svg)](https://azure.microsoft.com/en-us/services/automation/)
[![Microsoft Graph](https://img.shields.io/badge/Microsoft-Graph-742774.svg)](https://docs.microsoft.com/en-us/graph/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](#-licença)
[![Version](https://img.shields.io/badge/Version-2.1.0-orange.svg)](#-changelog)
[![Author](https://img.shields.io/badge/Author-Erick%20Medeiros%20%7C%20MVP%20Azure-blueviolet.svg)](#-autor)

---

## 📋 Sumário

- [Visão geral](#-visão-geral)
- [Principais funcionalidades](#-principais-funcionalidades)
- [Arquitetura](#-arquitetura)
- [Pré-requisitos](#-pré-requisitos)
- [Permissões necessárias](#-permissões-necessárias)
- [Módulos PowerShell](#-módulos-powershell)
- [Parâmetros](#-parâmetros)
- [Como funciona](#-como-funciona)
- [Instalação](#-instalação)
- [Uso](#-uso)
- [Output esperado](#-output-esperado)
- [Troubleshooting](#-troubleshooting)
- [Boas práticas](#-boas-práticas)
- [Changelog](#-changelog)
- [Contribuindo](#-contribuindo)
- [Licença](#-licença)
- [Autor](#-autor)

---

## 🎯 Visão geral

Este projeto automatiza a governança de identidades externas no Microsoft Entra ID, garantindo que usuários convidados sem MFA registrado sejam identificados e tratados de forma controlada. O runbook foi pensado para execução em **Azure Automation**, com autenticação por **Managed Identity**, integração com **Microsoft Graph** e emissão de relatório por e-mail.

## ✨ Principais funcionalidades

- ✅ Consulta os membros de um grupo específico do Entra ID
- ✅ Suporta grupos **atribuídos** e **dinâmicos**
- ✅ Filtra apenas usuários **Guest** ativos e fora da janela de tolerância configurada
- ✅ Valida métodos de autenticação por usuário via Graph
- ✅ Bloqueia convidados sem MFA com `accountEnabled = false`
- ✅ Executa em modo **DryRun** para simulação segura
- ✅ Gera relatório em **XLSX** com `ImportExcel` ou **CSV** como fallback
- ✅ Envia e-mail com anexo usando Microsoft Graph
- ✅ Aceita múltiplos destinatários em `To` e `Cc`

---

## 🏗 Arquitetura

```text
Azure Automation Runbook
        |
        |-- Managed Identity
        |
        +--> Microsoft Graph
              |-- GroupMember.Read.All
              |-- User.ReadWrite.All
              |-- UserAuthenticationMethod.Read.All
              |-- Mail.Send
```

---

## ✅ Pré-requisitos

| Recurso | Descrição |
|---|---|
| Azure Automation Account | Com identidade gerenciada habilitada |
| Microsoft Entra ID | Tenant com grupo contendo convidados |
| Permissões Graph | Application permissions concedidas à Managed Identity |
| Módulos PowerShell | Microsoft.Graph.* e opcionalmente ImportExcel |
| Caixa de e-mail | Mailbox válida para envio do relatório |

---

## 🔐 Permissões necessárias

A Managed Identity da Automation Account precisa, no mínimo, das seguintes permissões no Microsoft Graph:

- `GroupMember.Read.All`
- `User.ReadWrite.All`
- `UserAuthenticationMethod.Read.All`
- `Mail.Send`
- `AuditLog.Read.All` *(opcional/reserva operacional)*

---

## 📦 Módulos PowerShell

Instale os módulos abaixo na Automation Account, respeitando a ordem:

1. `Microsoft.Graph.Authentication`
2. `Microsoft.Graph.Users`
3. `Microsoft.Graph.Groups`
4. `Microsoft.Graph.Reports`
5. `Microsoft.Graph.Users.Actions`
6. `ImportExcel` *(opcional)*

> Runtime recomendado: **PowerShell 7.2** ou **5.1**, conforme compatibilidade do ambiente.

### Instalação manual dos módulos

Além da importação pela galeria do Azure Automation, você pode executar manualmente o script abaixo para instalar os módulos:

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

Import-AAModule -Name "Microsoft.Graph.Authentication" -Wait $true
Import-AAModule -Name "Microsoft.Graph.Users"
Import-AAModule -Name "Microsoft.Graph.Groups"
Import-AAModule -Name "Microsoft.Graph.Reports"
Import-AAModule -Name "Microsoft.Graph.Users.Actions"
Import-AAModule -Name "ImportExcel"

Write-Output "Concluído."
```

> Para instruções detalhadas de implantação, consulte **[IMPLANTACAO.md](./IMPLANTACAO.md)**.

---

## ⚙️ Parâmetros

| Parâmetro | Tipo | Padrão | Descrição |
|---|---|---|---|
| `HoursWithoutMfa` | `int` | `24` | Janela mínima em horas para considerar o convidado |
| `SkipPending` | `bool` | `$false` | Ignora usuários com convite pendente |
| `GuestGroupId` | `string` | obrigatório | Object ID do grupo que será avaliado |
| `SenderUpn` | `string` | obrigatório | Remetente do e-mail |
| `To` | `string` | obrigatório | Lista de destinatários |
| `Cc` | `string` | opcional | Lista de destinatários em cópia |
| `Subject` | `string` | padrão interno | Assunto do e-mail |
| `DryRun` | `bool` | `$false` | Simula sem bloquear usuários |

---

## 🔄 Como funciona

1. Autentica no Microsoft Graph com Managed Identity
2. Lista os membros do grupo informado em `GuestGroupId`
3. Filtra apenas convidados elegíveis para análise
4. Consulta os métodos de autenticação de cada usuário
5. Identifica quem não possui MFA efetivamente registrado
6. Executa bloqueio ou simulação, conforme `DryRun`
7. Gera relatório em CSV ou XLSX
8. Envia relatório por e-mail

---

## 📥 Instalação

Resumo do processo:

1. Criar a Automation Account com Managed Identity
2. Instalar os módulos PowerShell necessários
3. Conceder permissões Graph à identidade gerenciada
4. Criar ou importar o runbook `Runbook-GuestsSemMfa.ps1`
5. Publicar o runbook
6. Configurar parâmetros padrão
7. Agendar a execução

Para o passo a passo detalhado, consulte **[IMPLANTACAO.md](./IMPLANTACAO.md)**.

---

## 🚀 Uso

### Teste manual

Exemplo de parâmetros para DryRun:

```text
HoursWithoutMfa : 1
SkipPending     : False
GuestGroupId    : <Object ID do grupo>
SenderUpn       : noreply@seudominio.com
To              : secops@seudominio.com
DryRun          : True
```

### Execução em produção

Recomendação: agendamento diário, com `DryRun=False` e `SkipPending=True`.

---

## 📊 Output esperado

O runbook registra informações como:

- início e fim da execução
- quantidade de usuários analisados
- usuários sem MFA
- ação tomada: simulado ou bloqueado
- caminho do relatório gerado
- status de envio do e-mail

---

## 🩺 Troubleshooting

| Erro | Causa provável | Solução |
|---|---|---|
| `Connect-MgGraph not recognized` | Módulos Graph ausentes | Instalar módulos necessários |
| `403 Forbidden` | Permissões insuficientes | Revisar app roles da Managed Identity |
| `400 Bad Request` no sendMail | Payload JSON inválido | Validar estrutura do corpo e serialização |
| Relatório apenas CSV | `ImportExcel` ausente | Instalar módulo opcional |

---

## ✨ Boas práticas

- Validar primeiro com `DryRun=True`
- Agendar fora do horário comercial
- Monitorar jobs com falha no Azure Automation
- Revisar permissões da Managed Identity periodicamente
- Manter governança complementar com Access Reviews e Conditional Access

---

## 📜 Changelog

### v2.1.0 — 2026-06-10

#### 📘 Documentação

- Atualizado o `README.md` com estrutura mais objetiva e alinhada ao estado atual do projeto
- Adicionada instrução de **instalação manual dos módulos** diretamente no README
- Padronizados exemplos com placeholders para ambiente (`<SEU_RESOURCE_GROUP>` e `<SUA_AUTOMATION_ACCOUNT>`)
- Ajustado o direcionamento para o guia de implantação

### v2.0 — 2026-06-06

- Refatoração estrutural do runbook
- Substituição do filtro hardcoded por grupo do Entra ID
- Migração de `[switch]` para `[bool]`
- Inclusão de `GuestGroupId`
- Melhorias no fluxo de envio de e-mail e validação de MFA

### v1.0

- Versão inicial do projeto

---

## 🤝 Contribuindo

1. Fork este repositório
2. Crie uma branch para sua alteração
3. Faça commit das mudanças
4. Envie para o seu fork
5. Abra um Pull Request

---

## 📄 Licença

Distribuído sob a licença **MIT**.

---

## 👤 Autor

**Erick Medeiros**  
Microsoft MVP — Azure
