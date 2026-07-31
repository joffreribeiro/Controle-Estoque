# estoque-google-bridge — ponte de sincronização com o Google Agenda

Resolve o problema de precisar refazer login no Google a cada ~1h/sessão: em
vez do navegador falar direto com o Google (fluxo que nunca recebe um
`refresh_token`), este Worker guarda o `refresh_token` com segurança e entrega
`access_token`s novos sob demanda para o Controle de Estoque.

Uma vez configurado, você só faz login no Google **uma vez**. Depois disso o
Worker renova sozinho, para sempre (até você clicar "Desconectar" ou revogar o
acesso na sua conta Google).

Quem autoriza o navegador a falar com o Worker é o próprio login do Estoque
(Firebase Auth, projeto "estoquefi") — o Worker verifica esse token direto
contra a chave pública do Google, sem precisar de nenhuma chave colada
manualmente. Ou seja: **qualquer navegador ou dispositivo onde você já esteja
logado no Estoque libera a Agenda automaticamente**, sem passo extra nenhum.

## Pré-requisitos

- Conta gratuita na [Cloudflare](https://dash.cloudflare.com/sign-up) (sem cartão).
- Node.js instalado (para rodar `npx wrangler`).
- O OAuth Client já existente do Google Cloud Console (o mesmo `CLIENT_ID` que
  já estava em `google-calendar-sync.js`) — vamos precisar também da **Client
  Secret** dele, que hoje não é usada.

## Passo 1 — Pegar a Client Secret e cadastrar o redirect URI no Google

1. Abra o [Google Cloud Console → Credenciais](https://console.cloud.google.com/apis/credentials).
2. Clique no OAuth Client "Aplicativo da Web" já usado pelo app (o Client ID
   está no topo do `google-calendar-sync.js` antigo).
3. Copie o **Client secret** (guarde para o Passo 3).
4. Em **URIs de redirecionamento autorizados**, adicione (você só vai saber a
   URL exata depois do Passo 2 — pode voltar aqui pra completar):
   `https://estoque-google-bridge.<seu-subdominio>.workers.dev/callback`
5. Salve.

## Passo 2 — Criar e configurar o Worker

Nesta pasta (`cloudflare-worker/`), rode:

```bash
npx wrangler login
```

Isso abre o navegador para logar na Cloudflare (gratuito, sem cartão).

Crie o KV namespace que guarda o token:

```bash
npx wrangler kv namespace create TOKENS
```

O comando devolve algo como:

```
[[kv_namespaces]]
binding = "TOKENS"
id = "abcd1234..."
```

Copie esse `id` e cole em `wrangler.toml`, no lugar de
`SUBSTITUA_PELO_ID_DO_KV`.

## Passo 3 — Segredos (nunca vão para o git)

```bash
npx wrangler secret put GOOGLE_CLIENT_ID
# cole o Client ID (o mesmo que já estava no app)

npx wrangler secret put GOOGLE_CLIENT_SECRET
# cole a Client Secret copiada no Passo 1
```

Não existe um terceiro segredo pra colar em lugar nenhum — quem autoriza o
navegador é o seu próprio login no Estoque (ver explicação no topo deste
README).

## Passo 4 — Deploy

```bash
npx wrangler deploy
```

Ao final, o Wrangler mostra a URL pública, algo como:

```
https://estoque-google-bridge.SEUSUBDOMINIO.workers.dev
```

1. Volte ao **Passo 1**, item 4, e confirme que essa URL + `/callback` está
   cadastrada como redirect URI no Google Cloud Console.
2. Cole essa mesma URL (sem `/callback`, só a base) em `WORKER_URL`, no topo
   de `google-calendar-sync.js`.

## Passo 5 — Conectar pelo app

1. Abra o Controle de Estoque (logado normalmente) → Relacionamento → Calendário.
2. Clique **Conectar Google Agenda**.
3. Abre uma janela de consentimento do Google — aprove.
4. Pronto. A partir daqui, qualquer navegador ou dispositivo onde você fizer
   login no Estoque reconecta sozinho, sem pedir login do Google de novo.

## Como verificar que está funcionando

O `/token` exige o login do Estoque (`Authorization: Bearer <idToken>`), então
não dá pra testar com `curl` direto sem gerar esse token. O jeito mais simples
é abrir o Console do navegador na página do app já logado e rodar:

```js
firebase.auth().currentUser.getIdToken().then(t =>
  fetch('https://estoque-google-bridge.SEUSUBDOMINIO.workers.dev/token', { headers: { Authorization: 'Bearer ' + t } })
    .then(r => r.json()).then(console.log)
)
```

- Antes de conectar pelo app: `{"conectado":false}` (HTTP 404).
- Depois de conectar: `{"conectado":true,"access_token":"...","expires_in":3599}`.

## Se precisar desconectar/trocar de conta Google

Use o botão "Desconectar" no app (chama `/revoke`, que apaga o refresh_token
guardado e revoga no Google). Depois é só clicar "Conectar" de novo para
autorizar outra conta.

## Sobre o custo

Plano gratuito da Cloudflare Workers: 100 mil requisições/dia, sem cartão de
crédito. Uso normal deste app (algumas sincronizações por dia) fica muito
abaixo disso — custo esperado: **R$0**.
