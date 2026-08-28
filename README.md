# Dashboard de Inventário Intune

Sincroniza o inventário de dispositivos gerenciados pelo Microsoft Intune com uma Lista do SharePoint. A rotina é executada por PowerShell e pode ser agendada pelo Agendador de Tarefas do Windows; não requer Power Automate Premium.

> **Segurança:** nunca inclua IDs reais, valores de segredo, tokens ou o arquivo de configuração neste repositório. O `.gitignore` foi configurado para evitar o versionamento desses dados.

## O que a sincronização registra

Cada item da lista representa um dispositivo do Intune. A rotina cria equipamentos novos e atualiza os existentes com base no identificador único `IntuneDeviceId`.

| Campo na lista | Origem no Intune |
| --- | --- |
| Nome do dispositivo | `deviceName` |
| IntuneDeviceId | `id` |
| Fabricante / Modelo / NumeroSerie | `manufacturer`, `model`, `serialNumber` |
| UsuarioPrincipal / UsuarioNome | `userPrincipalName`, `userDisplayName` |
| SistemaOperacional / VersaoSO | `operatingSystem`, `osVersion` |
| Conformidade | `complianceState` |
| Propriedade | `managedDeviceOwnerType` |
| UltimaSincronizacao | `lastSyncDateTime` |
| UltimaAtualizacao | horário da execução da rotina |

## Pré-requisitos

- Conta com acesso ao Microsoft Entra ID, Intune e ao site do SharePoint que hospedará a lista.
- Licença ativa do Microsoft Intune no tenant.
- Um computador Windows que permaneça ligado e conectado à internet no horário do agendamento.
- Permissão para criar uma tarefa no Agendador de Tarefas do Windows.

## 1. Criar o aplicativo no Microsoft Entra ID

1. Acesse o [Microsoft Entra admin center](https://entra.microsoft.com).
2. Abra **Identidade** → **Aplicativos** → **Registros de aplicativos**.
3. Clique em **+ Novo registro**.
4. Preencha:
   - **Nome:** `Dashboard Inventário Intune`.
   - **Tipos de conta com suporte:** **Contas somente neste diretório organizacional**.
   - **URI de redirecionamento:** deixe em branco.
5. Clique em **Registrar**.
6. Copie e guarde, em local seguro:
   - **ID do aplicativo (cliente)**;
   - **ID do diretório (locatário)**.

## 2. Conceder permissões ao aplicativo

No aplicativo criado, abra **Permissões de API** → **+ Adicionar uma permissão** → **Microsoft Graph** → **Permissões de aplicativo**.

Adicione as duas permissões abaixo e, depois, clique em **Conceder consentimento do administrador** para a organização:

| Permissão | Finalidade |
| --- | --- |
| `DeviceManagementManagedDevices.Read.All` | Ler os dispositivos gerenciados pelo Intune. |
| `Sites.ReadWrite.All` | Criar e atualizar os itens da Lista do SharePoint. |

As duas permissões devem exibir o status **Granted/Concedido**. `Sites.ReadWrite.All` permite escrita em sites do SharePoint do tenant; para uma implantação com privilégio mínimo, uma evolução futura pode usar `Sites.Selected` com acesso restrito ao site da TI.

> A permissão padrão `User.Read` não é necessária para esta automação baseada em aplicativo e pode ser removida.

## 3. Criar o segredo do aplicativo

1. Abra **Certificates & secrets** → **Client secrets** → **+ New client secret**.
2. Use a descrição `Integração Dashboard Intune` e validade de 12 meses.
3. Clique em **Add**.
4. Copie imediatamente o campo **Value**.

Use o **Value**, e não o **Secret ID**. O valor só é exibido uma vez. Não o envie por e-mail, chat, planilha ou GitHub.

## 4. Criar a Lista do SharePoint

No site do SharePoint da TI, crie uma **Lista em branco** com o nome que desejar, por exemplo `Inventário Dispositivos INTUNE`.

### Colunas necessárias

Mantenha a coluna padrão `Title`, podendo renomear seu nome visível para **Nome do dispositivo**. O nome interno continuará sendo `Title`, que é usado pelo script.

Crie as colunas abaixo. Para as colunas novas, prefira nomes sem espaços e sem acentos.

| Nome | Tipo | Observação |
| --- | --- | --- |
| `IntuneDeviceId` | Uma linha de texto | Marque como obrigatório e com valores exclusivos. |
| `Fabricante` | Uma linha de texto | |
| `Modelo` | Uma linha de texto | |
| `NumeroSerie` | Uma linha de texto | |
| `UsuarioPrincipal` | Uma linha de texto | |
| `UsuarioNome` | Uma linha de texto | |
| `SistemaOperacional` | Uma linha de texto | |
| `VersaoSO` | Uma linha de texto | |
| `Conformidade` | Uma linha de texto | |
| `Propriedade` | Uma linha de texto | |
| `UltimaSincronizacao` | Data e hora | |
| `UltimaAtualizacao` | Data e hora | |

Copie a URL completa da lista. Exemplo:

```text
https://empresa.sharepoint.com/sites/TI/Lists/Inventario%20Dispositivos/AllItems.aspx
```

Também anote a URL do site, sem o trecho `/Lists/...`:

```text
https://empresa.sharepoint.com/sites/TI
```

## 5. Configurar a automação localmente

Abra o PowerShell na pasta deste projeto e execute:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\Configure-IntuneDashboard.ps1
```

Informe, quando solicitado:

- ID do diretório (locatário);
- ID do aplicativo (cliente);
- URL do site do SharePoint;
- URL completa da Lista do SharePoint;
- **valor** do segredo do aplicativo.

O script salva a configuração em:

```text
%LOCALAPPDATA%\IntuneDashboard\config.json
```

O segredo é armazenado criptografado pelo Windows (DPAPI). Apenas o mesmo usuário, na mesma máquina, consegue usá-lo para executar a sincronização.

## 6. Testar a sincronização manual

Execute:

```powershell
.\Sync-IntuneDashboard.ps1
```

O resultado esperado é semelhante a:

```text
Sincronizacao concluida: 10 criado(s), 90 atualizado(s), 100 dispositivo(s) processado(s).
```

Atualize a Lista do SharePoint e confira se os equipamentos e seus campos foram preenchidos. O script envia JSON em UTF-8 para preservar nomes com acentos.

## 7. Agendar a atualização automática

1. Abra o **Agendador de Tarefas** do Windows.
2. Clique em **Criar Tarefa**.
3. Na aba **Geral**:
   - Nome: `Sincronizar Inventario Intune`;
   - selecione **Executar somente quando o usuário estiver conectado**. Isso é necessário para acessar o segredo protegido pelo DPAPI.
4. Na aba **Disparadores**, crie um agendamento **Semanal**, de segunda a sexta-feira, às **08:00**.
5. Na aba **Ações**, crie uma ação **Iniciar um programa**:
   - **Programa/script:** `powershell.exe`;
   - **Adicionar argumentos** (ajuste o caminho conforme sua pasta local):

```text
-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "C:\caminho\para\o\projeto\Sync-IntuneDashboard.ps1"
```

6. Na aba **Configurações**, marque **Executar a tarefa o mais cedo possível após a perda de um início agendado**.
7. Salve e use **Executar** na tarefa para um teste final. O resultado esperado é `0x0`.

> O computador deve estar ligado e o usuário deve estar conectado às 08:00. Para maior disponibilidade, migre futuramente a rotina para uma VM ou servidor gerenciado.

