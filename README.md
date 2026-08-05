# Dashboard de Fluxo de Caixa — atualização automática diária

Este repositório gera e publica sozinho, todo dia, o dashboard de fluxo de caixa
a partir da planilha Excel que fica no OneDrive. Ninguém precisa rodar nada
manualmente — a única pessoa que continua tendo trabalho manual é quem já
atualiza a planilha (isso não muda).

## Como funciona

1. Todo dia (horário configurável), o GitHub Actions "acorda" sozinho.
2. Ele baixa a planilha mais recente do link do OneDrive.
3. Recalcula os KPIs, a projeção diária e a agenda de recebimentos/pagamentos.
4. Regenera `index.html` e publica no GitHub Pages.
5. Se nada mudou na planilha, ele não faz commit — sem ruído no histórico.

## Passo a passo para colocar no ar (uma vez só)

### 1. Confirme a permissão do link do OneDrive

O link precisa permitir acesso **sem login**, com download liberado:
- Abra a planilha no OneDrive → **Compartilhar** → verifique se a opção é
  **"Qualquer pessoa com o link"** (não "Pessoas específicas").
- Certifique-se de que a opção **"Permitir edição"** pode estar desligada
  (queremos só leitura), mas o link deve permitir visualizar/baixar o arquivo.

### 2. Crie o repositório no GitHub

1. Acesse [github.com/new](https://github.com/new).
2. Nome sugerido: `fluxo-caixa-dashboard`. Visibilidade: **Public** (necessário
   para o GitHub Pages gratuito publicar um site acessível por link — vocês já
   confirmaram que um link público sem senha é aceitável para este conteúdo).
3. Crie vazio (sem README, sem .gitignore).

### 3. Suba os arquivos deste pacote

Na pasta deste projeto (que você recebeu comigo), rode:

```bash
cd dashboard-repo
git init
git add .
git commit -m "Setup inicial do dashboard automático"
git branch -M main
git remote add origin https://github.com/SEU-USUARIO/fluxo-caixa-dashboard.git
git push -u origin main
```

(Se preferir, pode simplesmente arrastar os arquivos pela interface web do
GitHub em vez de usar linha de comando — funciona igual.)

### 4. Cadastre o link do OneDrive como "Secret"

1. No repositório no GitHub: **Settings → Secrets and variables → Actions →
   New repository secret**.
2. Nome: `ONEDRIVE_SHARE_URL`
3. Valor: cole o link de compartilhamento completo da planilha (o mesmo que
   você me passou).
4. Salvar. Isso mantém o link fora do código-fonte público.

### 5. Ative o GitHub Pages

1. **Settings → Pages**.
2. Em "Source", escolha **Deploy from a branch**.
3. Branch: `main`, pasta: `/ (root)`.
4. Salvar. Em alguns minutos o GitHub mostra o link público (algo como
   `https://SEU-USUARIO.github.io/fluxo-caixa-dashboard/`).

### 6. Rode o workflow uma vez manualmente para testar

1. Aba **Actions** do repositório → workflow **"Atualizar dashboard
   diariamente"** → **Run workflow**.
2. Acompanhe o log. Se tudo certo, ele baixa a planilha, gera o `index.html` e
   faz commit.
3. Abra o link do GitHub Pages e confira se o dashboard aparece atualizado.

Se o passo de download falhar (erro HTTP 401/403 no log), o link do OneDrive
provavelmente não está com a permissão "Qualquer pessoa com o link" — revise o
passo 1. Se mesmo assim não funcionar (comum em contas corporativas com
políticas de bloqueio), me avise: nesse caso o caminho alternativo é
autenticação via Microsoft Graph API (mais burocrático, exige registrar um app
no Azure AD), que também posso montar.

## Depois de configurado

Você não precisa mexer em mais nada. Todo dia às 09h (horário de Brasília) o
dashboard se atualiza sozinho. Para mudar o horário, edite a linha `cron` em
`.github/workflows/update-dashboard.yml` (o formato é `minuto hora * * *`, em
UTC — Brasília é UTC-3, então 12:00 UTC = 09:00 em Brasília, sem horário de
verão atualmente).

## Arquivos deste pacote

- `scripts/update_dashboard.py` — lê a planilha e gera o `index.html`.
- `scripts/chart.umd.min.js` — Chart.js embutido (o dashboard funciona
  offline, sem depender de CDN).
- `.github/workflows/update-dashboard.yml` — agenda a execução diária.
- `requirements.txt` — dependência Python (`openpyxl`).
- `index.html` — versão gerada de exemplo (será sobrescrita automaticamente).
