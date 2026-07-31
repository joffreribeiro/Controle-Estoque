/**
 * estoque-google-bridge — Cloudflare Worker
 *
 * Guarda o refresh_token do Google Calendar (obtido uma única vez, via fluxo
 * OAuth "Authorization Code") e devolve access_tokens novos sob demanda, para
 * que o Controle de Estoque sincronize com o Google Agenda sem pedir login a
 * cada ~1h ou a cada sessão de navegador — o problema do fluxo antigo (Google
 * Identity Services / token implícito), que nunca recebe refresh_token porque
 * é pensado para rodar só no navegador.
 *
 * Depois de configurar (ver README.md deste diretório), a URL pública deste
 * Worker vai em `WORKER_URL` no topo de google-calendar-sync.js.
 *
 * Segredos (definir com `wrangler secret put NOME`, nunca commitados):
 *   GOOGLE_CLIENT_ID     - mesmo Client ID já usado no app (OAuth "Aplicativo Web")
 *   GOOGLE_CLIENT_SECRET - Client Secret desse mesmo OAuth Client
 *   SYNC_SECRET          - string aleatória só sua; é o que autoriza o
 *                          navegador a chamar /authorize, /token e /revoke
 *                          (sem ela, ninguém que só veja o código-fonte
 *                          público do app consegue puxar tokens da sua agenda)
 *
 * KV namespace "TOKENS" (bind no wrangler.toml):
 *   refresh_token    - o refresh_token do Google (chave fixa; uso single-user)
 *   pending:<state>  - nonce de curta duração do fluxo /authorize -> /callback,
 *                       impede que alguém sem o SYNC_SECRET grude o refresh_token
 *                       DA CONTA DELE no seu slot só completando o consentimento
 *                       do Google por fora
 */

const SCOPE = 'https://www.googleapis.com/auth/calendar.events https://www.googleapis.com/auth/calendar.readonly';
const GOOGLE_AUTH_URL = 'https://accounts.google.com/o/oauth2/v2/auth';
const GOOGLE_TOKEN_URL = 'https://oauth2.googleapis.com/token';
const GOOGLE_REVOKE_URL = 'https://oauth2.googleapis.com/revoke';
const PENDING_TTL_SEGUNDOS = 600; // 10 min para completar o consentimento no popup

function cors(resp) {
    resp.headers.set('Access-Control-Allow-Origin', '*');
    resp.headers.set('Access-Control-Allow-Headers', 'Content-Type, X-Sync-Secret');
    return resp;
}

function json(dados, status) {
    return cors(new Response(JSON.stringify(dados), {
        status: status || 200,
        headers: { 'Content-Type': 'application/json' }
    }));
}

function textoAleatorio(tamanho) {
    const bytes = new Uint8Array(tamanho);
    crypto.getRandomValues(bytes);
    return [...bytes].map(b => b.toString(16).padStart(2, '0')).join('');
}

/** Aceita o segredo tanto por header (chamadas via fetch) quanto por query (navegação direta do popup). */
function autorizado(request, env) {
    const url = new URL(request.url);
    const enviado = url.searchParams.get('secret') || request.headers.get('X-Sync-Secret');
    return !!enviado && !!env.SYNC_SECRET && enviado === env.SYNC_SECRET;
}

function redirectUriDoWorker(request) {
    return new URL(request.url).origin + '/callback';
}

async function handleAuthorize(request, env) {
    if (!autorizado(request, env)) return json({ erro: 'não autorizado' }, 401);

    const state = textoAleatorio(16);
    await env.TOKENS.put('pending:' + state, '1', { expirationTtl: PENDING_TTL_SEGUNDOS });

    const params = new URLSearchParams({
        client_id: env.GOOGLE_CLIENT_ID,
        redirect_uri: redirectUriDoWorker(request),
        response_type: 'code',
        scope: SCOPE,
        access_type: 'offline',
        prompt: 'consent', // força o Google a reemitir refresh_token mesmo se você já tinha aprovado antes
        include_granted_scopes: 'true',
        state
    });

    return Response.redirect(GOOGLE_AUTH_URL + '?' + params.toString(), 302);
}

function paginaFechar(mensagem, sucesso) {
    const html = '<!doctype html><html><body style="font-family:sans-serif;padding:24px;color:#333">' +
        '<p>' + mensagem + '</p>' +
        '<script>' +
        'try { window.opener && window.opener.postMessage({ tipo: "google-bridge", sucesso: ' + (sucesso ? 'true' : 'false') + ' }, "*"); } catch (e) {}' +
        'setTimeout(function () { window.close(); }, 900);' +
        '</script>' +
        '</body></html>';
    return new Response(html, { headers: { 'Content-Type': 'text/html; charset=utf-8' } });
}

async function handleCallback(request, env) {
    const url = new URL(request.url);
    const code = url.searchParams.get('code');
    const state = url.searchParams.get('state');
    const erro = url.searchParams.get('error');

    if (erro) return paginaFechar('Autorização recusada pelo Google: ' + erro, false);
    if (!code || !state) return paginaFechar('Retorno inválido do Google (faltou code ou state).', false);

    const pendingKey = 'pending:' + state;
    const pendente = await env.TOKENS.get(pendingKey);
    if (!pendente) return paginaFechar('Sessão de autorização expirada ou inválida (state desconhecido). Tente conectar de novo.', false);
    await env.TOKENS.delete(pendingKey);

    const corpo = new URLSearchParams({
        code,
        client_id: env.GOOGLE_CLIENT_ID,
        client_secret: env.GOOGLE_CLIENT_SECRET,
        redirect_uri: redirectUriDoWorker(request),
        grant_type: 'authorization_code'
    });

    const resp = await fetch(GOOGLE_TOKEN_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: corpo.toString()
    });
    const dados = await resp.json();

    if (!resp.ok || !dados.refresh_token) {
        // O Google só devolve refresh_token na primeira aprovação (ou quando o
        // consentimento é revogado e refeito — por isso prompt=consent acima).
        return paginaFechar(
            'Não recebi um refresh_token do Google. Detalhe: ' + (dados.error_description || dados.error || 'desconhecido') +
            '. Se você já tinha aprovado este app antes, revogue o acesso em myaccount.google.com/permissions e tente de novo.',
            false
        );
    }

    await env.TOKENS.put('refresh_token', dados.refresh_token);
    return paginaFechar('Google Agenda conectado! Esta janela fecha sozinha.', true);
}

async function handleToken(request, env) {
    if (!autorizado(request, env)) return json({ erro: 'não autorizado' }, 401);

    const refreshToken = await env.TOKENS.get('refresh_token');
    if (!refreshToken) return json({ conectado: false }, 404);

    const corpo = new URLSearchParams({
        refresh_token: refreshToken,
        client_id: env.GOOGLE_CLIENT_ID,
        client_secret: env.GOOGLE_CLIENT_SECRET,
        grant_type: 'refresh_token'
    });

    const resp = await fetch(GOOGLE_TOKEN_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: corpo.toString()
    });
    const dados = await resp.json();

    if (!resp.ok) {
        // refresh_token revogado/expirado no lado do Google — precisa reconectar do zero.
        return json({ conectado: false, erro: dados.error_description || dados.error }, 401);
    }

    return json({
        conectado: true,
        access_token: dados.access_token,
        expires_in: dados.expires_in || 3500
    });
}

async function handleRevoke(request, env) {
    if (!autorizado(request, env)) return json({ erro: 'não autorizado' }, 401);

    const refreshToken = await env.TOKENS.get('refresh_token');
    if (refreshToken) {
        try {
            await fetch(GOOGLE_REVOKE_URL + '?token=' + encodeURIComponent(refreshToken), { method: 'POST' });
        } catch (e) { /* best-effort — mesmo se falhar, apaga do KV abaixo */ }
        await env.TOKENS.delete('refresh_token');
    }
    return json({ ok: true });
}

export default {
    async fetch(request, env) {
        if (request.method === 'OPTIONS') return cors(new Response(null, { status: 204 }));

        const url = new URL(request.url);
        try {
            if (url.pathname === '/authorize') return await handleAuthorize(request, env);
            if (url.pathname === '/callback') return await handleCallback(request, env);
            if (url.pathname === '/token') return await handleToken(request, env);
            if (url.pathname === '/revoke' && request.method === 'POST') return await handleRevoke(request, env);
            return json({ erro: 'rota não encontrada' }, 404);
        } catch (e) {
            return json({ erro: 'erro interno: ' + (e && e.message ? e.message : String(e)) }, 500);
        }
    }
};
