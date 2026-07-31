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
 * Autorização: em vez de uma chave colada manualmente em cada navegador, este
 * Worker aceita o próprio token de login do Firebase que o Estoque já usa
 * (projeto "estoquefi") — verificado aqui via JWKS público do Google, sem
 * precisar de Firebase Admin SDK nem de nenhum segredo novo por navegador.
 * Qualquer navegador/dispositivo já logado no Estoque libera a Agenda sem
 * nenhum passo extra.
 *
 * Depois de configurar (ver README.md deste diretório), a URL pública deste
 * Worker vai em `WORKER_URL` no topo de google-calendar-sync.js.
 *
 * Segredos (definir com `wrangler secret put NOME`, nunca commitados):
 *   GOOGLE_CLIENT_ID     - mesmo Client ID já usado no app (OAuth "Aplicativo Web")
 *   GOOGLE_CLIENT_SECRET - Client Secret desse mesmo OAuth Client
 *
 * KV namespace "TOKENS" (bind no wrangler.toml):
 *   refresh_token   - o refresh_token do Google (chave fixa; uso single-user)
 *   pending:<state> - nonce de curta duração do fluxo /authorize -> /callback
 *   ticket:<ticket> - nonce de uso único que autoriza abrir o /authorize
 *                     (emitido só depois de validar o login do Firebase,
 *                     porque /authorize é uma navegação de página inteira e
 *                     não dá pra anexar um header Authorization nela)
 */

const SCOPE = 'https://www.googleapis.com/auth/calendar.events https://www.googleapis.com/auth/calendar.readonly';
const GOOGLE_AUTH_URL = 'https://accounts.google.com/o/oauth2/v2/auth';
const GOOGLE_TOKEN_URL = 'https://oauth2.googleapis.com/token';
const GOOGLE_REVOKE_URL = 'https://oauth2.googleapis.com/revoke';
const PENDING_TTL_SEGUNDOS = 600; // 10 min para completar o consentimento no popup
const TICKET_TTL_SEGUNDOS = 120; // 2 min — só o tempo de abrir o popup

const FIREBASE_PROJECT_ID = 'estoquefi';
const FIREBASE_JWKS_URL = 'https://www.googleapis.com/service_accounts/v1/jwk/securetoken@system.gserviceaccount.com';

let _jwksCache = null; // { porKid: {...}, expiraEm: timestamp } — vive enquanto o isolate do Worker ficar quente

function cors(resp) {
    resp.headers.set('Access-Control-Allow-Origin', '*');
    resp.headers.set('Access-Control-Allow-Headers', 'Content-Type, Authorization');
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

function base64UrlParaBytes(str) {
    str = str.replace(/-/g, '+').replace(/_/g, '/');
    while (str.length % 4) str += '=';
    const bin = atob(str);
    const bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    return bytes;
}

function base64UrlParaJson(str) {
    return JSON.parse(new TextDecoder().decode(base64UrlParaBytes(str)));
}

async function buscarJwksFirebase() {
    if (_jwksCache && Date.now() < _jwksCache.expiraEm) return _jwksCache.porKid;
    const resp = await fetch(FIREBASE_JWKS_URL);
    const dados = await resp.json();
    const porKid = {};
    (dados.keys || []).forEach(function (k) { porKid[k.kid] = k; });
    _jwksCache = { porKid, expiraEm: Date.now() + 3600 * 1000 };
    return porKid;
}

/**
 * Verifica um ID token do Firebase Auth (o mesmo que o Estoque já emite ao
 * fazer login) sem depender do Firebase Admin SDK: confere assinatura RS256
 * contra a JWKS pública do Google, e os campos padrão (iss/aud/exp).
 * Devolve { uid, email } se válido, ou null.
 */
async function verificarFirebaseIdToken(idToken) {
    if (!idToken) return null;
    const partes = idToken.split('.');
    if (partes.length !== 3) return null;

    let header, payload;
    try {
        header = base64UrlParaJson(partes[0]);
        payload = base64UrlParaJson(partes[1]);
    } catch (e) { return null; }

    if (payload.iss !== 'https://securetoken.google.com/' + FIREBASE_PROJECT_ID) return null;
    if (payload.aud !== FIREBASE_PROJECT_ID) return null;
    if (!payload.exp || Date.now() >= payload.exp * 1000) return null;
    if (!payload.sub) return null;

    const jwks = await buscarJwksFirebase();
    const jwk = jwks[header.kid];
    if (!jwk) return null;

    let chave;
    try {
        chave = await crypto.subtle.importKey('jwk', jwk, { name: 'RSASSA-PKCS1-v1_5', hash: 'SHA-256' }, false, ['verify']);
    } catch (e) { return null; }

    const dadosAssinados = new TextEncoder().encode(partes[0] + '.' + partes[1]);
    const assinatura = base64UrlParaBytes(partes[2]);

    const valido = await crypto.subtle.verify('RSASSA-PKCS1-v1_5', chave, assinatura, dadosAssinados);
    if (!valido) return null;

    return { uid: payload.sub, email: payload.email || null };
}

/** Extrai e valida o "Authorization: Bearer <idToken>" da requisição. */
async function usuarioAutenticado(request) {
    const auth = request.headers.get('Authorization') || '';
    const m = auth.match(/^Bearer\s+(.+)$/);
    if (!m) return null;
    return verificarFirebaseIdToken(m[1]);
}

function redirectUriDoWorker(request) {
    return new URL(request.url).origin + '/callback';
}

/**
 * Chamada via fetch() (com Authorization: Bearer <idToken do Firebase>).
 * Devolve um "ticket" de uso único, porque o próximo passo (/authorize) é uma
 * navegação de página inteira (window.open) — não dá pra mandar header nela.
 */
async function handleIniciarAutorizacao(request, env) {
    const usuario = await usuarioAutenticado(request);
    if (!usuario) return json({ erro: 'não autenticado' }, 401);

    const ticket = textoAleatorio(16);
    await env.TOKENS.put('ticket:' + ticket, usuario.uid, { expirationTtl: TICKET_TTL_SEGUNDOS });
    return json({ ticket });
}

async function handleAuthorize(request, env) {
    const url = new URL(request.url);
    const ticket = url.searchParams.get('ticket');
    if (!ticket) return json({ erro: 'ticket ausente' }, 401);

    const ticketKey = 'ticket:' + ticket;
    const dono = await env.TOKENS.get(ticketKey);
    if (!dono) return json({ erro: 'ticket inválido ou expirado — clique em Conectar de novo' }, 401);
    await env.TOKENS.delete(ticketKey);

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
    const usuario = await usuarioAutenticado(request);
    if (!usuario) return json({ erro: 'não autenticado' }, 401);

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
    const usuario = await usuarioAutenticado(request);
    if (!usuario) return json({ erro: 'não autenticado' }, 401);

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
            if (url.pathname === '/iniciar-autorizacao' && request.method === 'POST') return await handleIniciarAutorizacao(request, env);
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
