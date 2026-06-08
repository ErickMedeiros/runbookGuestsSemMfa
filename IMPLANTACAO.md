# 🚀 Guia de Implantação Passo a Passo

**Runbook-GuestsSemMfa** — Implementação completa do zero até produção.

> **Autor:** Erick Medeiros — *Microsoft MVP Azure*
> **Versão:** 2.0 · **Data:** 06 de junho de 2026
> **Duração estimada:** 60-90 minutos (excluindo propagação de permissões: +30 min)

---

## 📑 Índice

1. [Preparação do ambiente](#1-preparação-do-ambiente)
2. [Criar Automation Account](#2-criar-automation-account)
3. [Habilitar Managed Identity](#3-habilitar-managed-identity)
4. [Instalar módulos PowerShell](#4-instalar-módulos-powershell)
5. [Conceder permissões Graph](#5-conceder-permissões-graph)
6. [Criar o grupo de Guests](#6-criar-o-grupo-de-guests)
7. [Criar e publicar o Runbook](#7-criar-e-publicar-o-runbook)
8. [Teste em modo DryRun](#8-teste-em-modo-dryrun)
9. [Configurar Schedule](#9-configurar-schedule)
10. [Configurar Alertas](#10-configurar-alertas)
11. [Validação final](#11-validação-final)
12. [Checklist de produção](#12-checklist-de-produção)

---

## 1. Preparação do ambiente

### Pré-requisitos

- [ ] Subscription Azure ativa
- [ ] Resource Group criado (ex: `rg-governance`)
- [ ] Permissão **Owner** ou **Contributor** na subscription
- [ ] Permissão **Global Administrator** ou **Privileged Role Administrator** no Entra ID
- [ ] Mailbox Exchange Online disponível para envio (usuário ou shared)
- [ ] Lista de destinatários do relatório definida

### Decisões prévias

| Item | Sua decisão |
|---|---|
| Nome da Automation Account | `aa-disablegueusers` |
| Resource Group | `rg-governance` |
| Região | `East US` (ou conforme política) |
| Mailbox remetente | `noreply@seudominio.com` |
| Destinatários | `secops@seudominio.com; gestao@seudominio.com` |
| Janela inicial sem MFA | `24` horas |
| Frequência do schedule | Diária às 02:00 BRT |

---

## 2. Criar Automation Account

### Via Portal Azure

1. Acessar [portal.azure.com](https://portal.azure.com)
2. **Create a resource → Automation**
3. Preencher:
   - **Subscription:** sua subscription
   - **Resource group:** `rg-governance`
   - **Name:** `aa-disablegueusers`
   - **Region:** conforme política
4. Aba **Advanced:**
   - **Managed identities:** ✅ **System assigned**
5. **Review + Create → Create**
6. Aguardar conclusão do deploy (~2 min)

### Via Azure CLI (alternativa)

```bash
az automation account create \
    --name aa-disablegueusers \
    --resource-group rg-governance \
    --location eastus \
    --sku Basic \
    --assign-identity
```

---

## 3. Habilitar Managed Identity

Se não habilitou no passo 2:

1. **Automation Account → Account Settings → Identity**
2. Aba **System assigned**
3. **Status:** `On`
4. **Save**
5. **Copie o Object (principal) ID** — você vai precisar dele depois

> 💡 Guarde o `Object ID` em local seguro. Exemplo: `6cf94988-7df3-44bd-a4bf-6f9c2ab066d5`

---

## 4. Instalar módulos PowerShell

### Caminho no portal

**Automation Account → Shared Resources → Modules → + Add a module**

### Origem

- **Source:** `PowerShell Gallery`
- **Runtime version:** `5.1`

### Módulos (instalar nesta ordem)

| Ordem | Módulo | Aguardar status |
|---|---|---|
| 1 | `Microsoft.Graph.Authentication` | ✅ Available |
| 2 | `Microsoft.Graph.Users` | ✅ Available |
| 3 | `Microsoft.Graph.Groups` | ✅ Available |
| 4 | `Microsoft.Graph.Reports` | ✅ Available |
| 5 | `Microsoft.Graph.Users.Actions` | ✅ Available |
| 6 *(opcional)* | `ImportExcel` | ✅ Available |

> ⚠️ **Crítico:** aguarde cada módulo ficar `Available` antes de instalar o próximo. A ordem importa por causa das dependências internas do Graph SDK.
> 
> Tempo total: ~10-15 minutos.

### Validação

Após instalar todos, em **Modules**, confirme:
- Status `Available` em todos
- Nenhum com erro de importação

---

## 5. Conceder permissões Graph

A Managed Identity precisa de permissões **Application** no Microsoft Graph para operar.

### Via Cloud Shell (recomendado)

1. Abrir [shell.azure.com](https://shell.azure.com)
2. Selecionar **PowerShell**
3. Executar:

```powershell
# Ajustar para o nome real da sua Automation Account
$miObjectId = (Get-AzADServicePrincipal -DisplayName "aa-disablegueusers").Id
Write-Output "Managed Identity Object ID: $miObjectId"

# Obter o Service Principal do Microsoft Graph
$graphSP = Get-AzADServicePrincipal -Filter "displayName eq 'Microsoft Graph'"

# Permissões necessárias
$permissions = @(
    "GroupMember.Read.All",
    "AuditLog.Read.All",
    "Mail.Send",
    "User.ReadWrite.All",
    "UserAuthenticationMethod.Read.All"
)

foreach ($permName in $permissions) {
    $role = $graphSP.AppRole | Where-Object { $_.Value -eq $permName }
    if ($role) {
        try {
            New-AzADServicePrincipalAppRoleAssignment `
                -ServicePrincipalId $miObjectId `
                -ResourceId $graphSP.Id `
                -AppRoleId $role.Id
            Write-Output "OK: $permName"
        } catch {
            Write-Output "Já existe ou falhou: $permName - $($_.Exception.Message)"
        }
    } else {
        Write-Output "Role não encontrada: $permName"
    }
}
```

### Validação

Saída esperada:
```
OK: GroupMember.Read.All
OK: AuditLog.Read.All
OK: Mail.Send
OK: User.ReadWrite.All
OK: UserAuthenticationMethod.Read.All
```

Ou no portal: **Enterprise Applications → All applications → Filter: Application Type = Managed Identities → aa-disablegueusers → Permissions** — devem aparecer as 5 permissões.

> ⏰ **Aguarde 15-60 minutos** após a concessão. Tokens da Managed Identity são cacheados e novas permissões só entram em vigor após o refresh do token.

---

## 6. Criar o grupo de Guests

Crie um grupo no Entra ID com os usuários Guest que serão monitorados.

### Opção A — Grupo dinâmico (recomendado)

1. **Entra ID → Groups → + New group**
2. **Group type:** `Security`
3. **Group name:** `GR-UsersGuest`
4. **Membership type:** `Dynamic User`
5. **Dynamic query:**
   ```
   (user.userType -eq "Guest")
   ```
   Ou para filtrar por domínio:
   ```
   (user.userType -eq "Guest") and (user.userPrincipalName -contains "fornecedor.com")
   ```
6. **Save → Create**
7. **Copie o Object ID** do grupo

### Opção B — Grupo atribuído

1. **Entra ID → Groups → + New group**
2. **Membership type:** `Assigned`
3. Adicionar manualmente os Guests
4. **Copie o Object ID** do grupo

> 💡 O runbook suporta os dois tipos sem alteração.

---

## 7. Criar e publicar o Runbook

### 7.1. Criar o Runbook

1. **Automation Account → Process Automation → Runbooks → + Create a runbook**
2. Preencher:
   - **Name:** `RumGuestDisable`
   - **Type:** `PowerShell`
   - **Runtime version:** `5.1`
3. **Create**

### 7.2. Colar o conteúdo

1. No editor que abrir, **selecionar tudo (Ctrl+A) → Delete**
2. Colar o conteúdo completo de `Runbook-GuestsSemMfa.ps1`
3. **Ajustar os defaults** no bloco `param()`:

```powershell
[string] $GuestGroupId = "<COLE-O-OBJECT-ID-DO-GRUPO-AQUI>"
[string] $SenderUpn    = "noreply@seudominio.com"
[string] $To           = "secops@seudominio.com; gestao@seudominio.com"
```

### 7.3. Salvar e publicar

1. **Save** (Ctrl+S)
2. **Publish** ← **passo crítico!**

> ⚠️ Sem **Publish**, o runbook não pode ser agendado ou chamado externamente. Apenas o Test Pane usa o draft.

---

## 8. Teste em modo DryRun

### 8.1. Abrir o Test Pane

**Runbook → Edit → Test pane**

### 8.2. Configurar parâmetros do teste

| Parâmetro | Valor de teste |
|---|---|
| `HoursWithoutMfa` | `1` (para detectar Guests recém-criados) |
| `SkipPending` | `False` |
| `GuestGroupId` | *(default já configurado)* |
| `SenderUpn` | *(default já configurado)* |
| `To` | *(default já configurado)* |
| `Cc` | *(em branco)* |
| `Subject` | *(default)* |
| `DryRun` | **`True`** ← **importante** |

### 8.3. Executar

**Start**

### 8.4. Output esperado

```
=== Início do runbook ... ===
Janela sem MFA: 1 h | SkipPending: False | DryRun: True
Conectando ao Microsoft Graph com Managed Identity...
Data/hora de corte: ...
Listando convidados (Guest) do tenant...
Buscando membros do grupo '...'...
Total de convidados considerados após filtro: N
Verificando métodos de autenticação de cada convidado...
  SEM MFA: usuario1@...
  COM MFA: usuario2@... (2 método(s))
Candidatos a bloqueio (sem MFA e dentro da janela): X
Total processado: X | Bloqueados/Simulados: X
Relatório gerado (CSV): ...
E-mail enviado a partir de '...'.
=== Fim do runbook ... ===
```

### 8.5. Validar o e-mail

- [ ] E-mail chegou nos destinatários?
- [ ] Anexo CSV está presente?
- [ ] Conteúdo do CSV bate com o output do runbook?
- [ ] Todos com `Action = Simulado` (devido ao DryRun=True)?

> ✅ Se tudo OK, prosseguir. Se houver erro, consultar o [Troubleshooting](./README.md#-troubleshooting).

---

## 9. Configurar Schedule

### 9.1. Criar o agendamento

1. **Automation Account → Shared Resources → Schedules → + Add a schedule**
2. Preencher:
   - **Name:** `Daily-02h-BRT`
   - **Description:** `Execução diária do bloqueio de Guests sem MFA`
   - **Starts:** próxima ocorrência das 02:00 (UTC-3 = 05:00 UTC)
   - **Time zone:** `(UTC-03:00) Brasilia`
   - **Recurrence:** `Recurring`
   - **Recur every:** `1 Day`
   - **Set expiration:** `No`
3. **Create**

### 9.2. Linkar ao Runbook

1. **Runbook `RumGuestDisable` → Schedules → + Add a schedule**
2. **Link a schedule to your runbook:** selecionar `Daily-02h-BRT`
3. **Configure parameters and run settings:**

   | Parâmetro | Valor de produção |
   |---|---|
   | `HoursWithoutMfa` | `24` |
   | `SkipPending` | `True` |
   | `GuestGroupId` | *(default já configurado)* |
   | `SenderUpn` | *(default já configurado)* |
   | `To` | *(default já configurado)* |
   | `Subject` | `[PROD] Relatório - Guests bloqueados sem MFA` |
   | `DryRun` | **`False`** ← **bloqueio real** |

4. **OK → Save**

---

## 10. Configurar Alertas

Criar alerta para notificar falhas no runbook.

### 10.1. Criar Action Group (se não existir)

1. **Monitor → Alerts → Action groups → + Create**
2. **Name:** `ag-runbook-failures`
3. **Notifications:** adicionar e-mail da equipe SecOps
4. **Create**

### 10.2. Criar regra de alerta

1. **Automation Account → Monitoring → Alerts → + Create → Alert rule**
2. **Scope:** automation account já selecionado
3. **Condition:**
   - **Signal name:** `Total Jobs`
   - **Dimension — Status:** `Failed`
   - **Operator:** `Greater than`
   - **Threshold:** `0`
   - **Aggregation granularity:** `5 minutes`
   - **Frequency:** `5 minutes`
4. **Actions:** selecionar `ag-runbook-failures`
5. **Details:**
   - **Severity:** `2 - Warning`
   - **Alert rule name:** `Runbook-GuestsSemMfa-Failed`
6. **Create**

---

## 11. Validação final

### Lista de verificação

- [ ] Runbook executou com sucesso no Test Pane (DryRun=True)
- [ ] E-mail com relatório foi recebido
- [ ] Relatório contém os dados esperados
- [ ] Schedule criado e linkado com `DryRun=False`
- [ ] Alerta configurado
- [ ] Permissões da MI verificadas no portal
- [ ] Documentação interna atualizada

### Primeira execução agendada

Aguardar a primeira execução pelo schedule (próxima 02:00 BRT) e validar:

1. **Automation Account → Jobs** — status `Completed`
2. Conferir o **e-mail recebido** com o relatório
3. Verificar no **Entra ID** se os Guests listados estão `Account status: Disabled`
4. **Audit logs** no Entra ID mostram a alteração feita pela MI

---

## 12. Checklist de produção

### Operacional

- [ ] Schedule rodando diariamente sem falhas
- [ ] E-mails sendo recebidos pela equipe
- [ ] Logs de jobs sendo arquivados (retenção padrão: 30 dias)
- [ ] Alertas testados (forçar uma falha intencional para validar)

### Segurança

- [ ] Managed Identity tem **apenas** as permissões mínimas necessárias
- [ ] Object ID do grupo `GR-UsersGuest` documentado
- [ ] Mailbox remetente protegida por MFA / Conditional Access
- [ ] Permissões da MI revisadas trimestralmente

### Governança

- [ ] Procedimento de **revogação de bloqueio** documentado (caso de falso positivo)
- [ ] Processo de **exceção** para Guests que precisam estar isentos
- [ ] Revisão periódica do filtro do grupo dinâmico
- [ ] Integração com processo de **Access Reviews** do Entra ID

---

## 🆘 Suporte

Em caso de problemas durante a implantação, consulte:

- **[Troubleshooting](./README.md#-troubleshooting)** no README
- **Automation Account → Jobs** — logs detalhados de cada execução
- **Entra ID → Audit logs** — confirmação das alterações feitas pela MI
- **Enterprise Applications → Managed Identity → Sign-ins** — verificar autenticações

---

## 📝 Histórico de alterações

### v2.0 — 2026-06-06

**Autor:** Erick Medeiros — MVP Microsoft Azure

#### Mudanças no script base

| Categoria | Item |
|---|---|
| 🔄 **Breaking** | Filtro de Guests alterado de sufixo de UPN para membership de grupo Entra ID |
| 🔄 **Breaking** | Parâmetros `$SkipPending` e `$DryRun` migrados de `[switch]` para `[bool]` |
| ✨ **Adicionado** | Parâmetro obrigatório `$GuestGroupId` (suporta grupos dinâmicos e atribuídos) |
| ✨ **Adicionado** | Parametrização de remetente e destinatários (sem hardcode) |
| ✨ **Adicionado** | Módulo `Microsoft.Graph.Groups` na lista de dependências |
| ✨ **Adicionado** | Permissão `UserAuthenticationMethod.Read.All` na Managed Identity |
| ✨ **Adicionado** | Output detalhado SEM MFA / COM MFA por usuário |
| 🔧 **Alterado** | Verificação de MFA via `Invoke-MgGraphRequest` per-user (substitui `Get-MgReportAuthenticationMethodUserRegistrationDetail`) |
| 🔧 **Alterado** | Envio de e-mail via `Invoke-MgGraphRequest` + `ConvertTo-Json -Depth 20` (substitui `Send-MgUserMail`) |
| 🐛 **Corrigido** | Serialização JSON do payload de e-mail (`StartObject` error) |
| 🐛 **Corrigido** | Falsos negativos na detecção de MFA |
| 🐛 **Corrigido** | Incompatibilidade `[switch]` com Test Pane do Azure Automation |

#### Documentação

- ➕ Criado `README.md` com documentação técnica completa
- ➕ Criado `IMPLANTACAO.md` (este documento) com guia passo a passo
- ➕ Header do script com SYNOPSIS, PARAMETER, EXAMPLE e CHANGELOG completo

### v1.0 — Modelo base

- Versão inicial do script (modelo de referência interna)
- Filtro por sufixo de UPN, MFA via relatório agregado, e-mail via SDK cmdlet

---

---

## 👤 Autor

**Erick Medeiros**
*Microsoft MVP — Azure*

Especialista em administração Azure, arquitetura Microsoft 365 e automação de infraestrutura cloud com foco em segurança e governança de identidade.

---

*Última atualização: 06 de junho de 2026 · Versão 2.0*
