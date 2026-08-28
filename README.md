# Dashboard de Inventário Intune

## Scripts

- `Configure-IntuneDashboard.ps1`: solicita os identificadores do Entra, a URL do site do SharePoint e o segredo do aplicativo. O segredo é guardado criptografado com DPAPI no perfil Windows do usuário que executa a configuração.
- `Sync-IntuneDashboard.ps1`: lê os dispositivos gerenciados do Intune e cria ou atualiza seus registros na lista **Inventário de Dispositivos**.

Execute primeiro a configuração e, em seguida, a sincronização manual. O agendamento no Windows só deve ser criado depois que a execução manual for validada.
