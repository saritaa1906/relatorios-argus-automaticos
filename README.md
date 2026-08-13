# Relatórios Argus automáticos

Este projeto roda diariamente no GitHub Actions. Ele consulta a API Argus, identifica carteiras iniciadas desde 01/08/2026 que completaram 14 dias, baixa todas as ligações em duas janelas de 7 dias, gera o Excel e envia ao Google Drive.

## Segurança

- O repositório deve ser **privado**.
- Tokens, credenciais, bases e relatórios nunca entram no Git.
- Os dados pessoais existem somente na memória/disco temporário do runner e são apagados após o envio.

## Secrets necessários

Cadastre em **Settings → Secrets and variables → Actions**:

- `ARGUS_TOKEN`
- `GOOGLE_CLIENT_ID`
- `GOOGLE_CLIENT_SECRET`
- `GOOGLE_REFRESH_TOKEN`
- `GOOGLE_DRIVE_FOLDER_ID`

## Execução

- Automática: todos os dias às 09:00 de São Paulo.
- Manual: **Actions → Relatórios Argus após 14 dias → Run workflow**.
- Teste sem upload: marque `dry_run` na execução manual.

O Google Drive é usado também para evitar duplicidade: se o arquivo final já existir na pasta, o lote não é processado novamente.
