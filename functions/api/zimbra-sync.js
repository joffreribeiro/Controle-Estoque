/**
 * functions/api/zimbra-sync.js - Cloudflare Pages Function
 *
 * Proxy que fala com o servidor Zimbra via CalDAV (server-to-server, sem
 * restrição de CORS) e devolve os eventos como JSON simples para o front-end.
 * Não guarda nada no Firestore aqui — quem grava é o navegador (crm-zimbra-import.js),
 * usando a sessão já autenticada do usuário. Este endpoint é só o "tradutor".
 *
 * Requer 3 variáveis de ambiente (Cloudflare Pages → Settings → Environment
 * variables and secrets), configuradas como "Secret":
 *   - ZIMBRA_CALDAV_URL  ex: https://mail.suaempresa.com/dav/usuario@dominio.com/Calendar/
 *   - ZIMBRA_USER        ex: usuario@dominio.com
 *   - ZIMBRA_PASSWORD    senha (ou senha de aplicativo) do Zimbra
 *
 * Limitações conhecidas (v1):
 *   - Eventos recorrentes (RRULE) aparecem só na data original, sem expandir
 *     as repetições.
 *   - Importação é sob demanda (botão "Importar do Zimbra"), não há polling
 *     automático em background nesta versão.
 */

const JANELA_DIAS_PASSADO = 7;
const JANELA_DIAS_FUTURO = 90;

export async function onRequestGet(context) {
    const env = context.env || {};
    const url = env.ZIMBRA_CALDAV_URL;
    const user = env.ZIMBRA_USER;
    const pass = env.ZIMBRA_PASSWORD;

    if (!url || !user || !pass) {
        return json({ erro: 'Configure as variáveis ZIMBRA_CALDAV_URL, ZIMBRA_USER e ZIMBRA_PASSWORD no Cloudflare Pages (Settings → Environment variables and secrets).' }, 500);
    }

    const agora = new Date();
    const inicio = new Date(agora.getTime() - JANELA_DIAS_PASSADO * 86400000);
    const fim = new Date(agora.getTime() + JANELA_DIAS_FUTURO * 86400000);

    let resp;
    try {
        resp = await fetch(url, {
            method: 'REPORT',
            headers: {
                'Authorization': 'Basic ' + btoa(user + ':' + pass),
                'Content-Type': 'application/xml; charset=utf-8',
                'Depth': '1'
            },
            body: buildCalendarQueryXml(inicio, fim)
        });
    } catch (e) {
        return json({ erro: 'Falha ao conectar ao servidor Zimbra: ' + e.message }, 502);
    }

    if (!resp.ok) {
        const texto = await resp.text();
        return json({ erro: 'Zimbra respondeu ' + resp.status + ': ' + texto.slice(0, 400) }, 502);
    }

    const xml = await resp.text();
    const blocos = extractCalendarDataBlocks(xml);
    const eventos = [];
    blocos.forEach((bloco) => {
        const ev = parseIcsEvent(bloco);
        if (ev) eventos.push(ev);
    });

    return json({ eventos, total: eventos.length });
}

function json(obj, status) {
    return new Response(JSON.stringify(obj), {
        status: status || 200,
        headers: { 'Content-Type': 'application/json; charset=utf-8' }
    });
}

// ── Requisição CalDAV (RFC 4791 calendar-query) ──

function pad(n) { return String(n).padStart(2, '0'); }

function toIcsUtc(d) {
    return d.getUTCFullYear() + pad(d.getUTCMonth() + 1) + pad(d.getUTCDate()) + 'T' +
        pad(d.getUTCHours()) + pad(d.getUTCMinutes()) + pad(d.getUTCSeconds()) + 'Z';
}

function buildCalendarQueryXml(inicio, fim) {
    return '<?xml version="1.0" encoding="utf-8" ?>' +
        '<C:calendar-query xmlns:D="DAV:" xmlns:C="urn:ietf:params:xml:ns:caldav">' +
            '<D:prop><D:getetag/><C:calendar-data/></D:prop>' +
            '<C:filter><C:comp-filter name="VCALENDAR"><C:comp-filter name="VEVENT">' +
                '<C:time-range start="' + toIcsUtc(inicio) + '" end="' + toIcsUtc(fim) + '"/>' +
            '</C:comp-filter></C:comp-filter></C:filter>' +
        '</C:calendar-query>';
}

// ── Extração do XML multistatus (resposta CalDAV) ──

function extractCalendarDataBlocks(xml) {
    const blocos = [];
    const re = /<[^>]*calendar-data[^>]*>([\s\S]*?)<\/[^>]*calendar-data>/gi;
    let m;
    while ((m = re.exec(xml))) blocos.push(unescapeXml(m[1]));
    return blocos;
}

function unescapeXml(s) {
    return String(s)
        .replace(/&lt;/g, '<').replace(/&gt;/g, '>')
        .replace(/&quot;/g, '"').replace(/&apos;/g, "'")
        .replace(/&amp;/g, '&');
}

// ── Parser mínimo de iCalendar (VEVENT) ──

function unfold(ics) {
    return ics.replace(/\r\n[ \t]/g, '').replace(/\n[ \t]/g, '');
}

function unescapeIcsText(s) {
    return String(s || '').replace(/\\n/gi, '\n').replace(/\\,/g, ',').replace(/\\;/g, ';').replace(/\\\\/g, '\\');
}

function acharLinha(linhas, nome) {
    const alvo = nome.toUpperCase();
    for (let i = 0; i < linhas.length; i++) {
        const upper = linhas[i].toUpperCase();
        if (upper.indexOf(alvo + ':') === 0 || upper.indexOf(alvo + ';') === 0) return linhas[i];
    }
    return null;
}

function valorDaLinha(linha) {
    const idx = linha.indexOf(':');
    return idx === -1 ? '' : linha.slice(idx + 1).trim();
}

function parseIcsDateTime(linha) {
    const params = linha.slice(0, linha.indexOf(':'));
    const valor = valorDaLinha(linha);

    if (/VALUE=DATE\b/i.test(params) || /^\d{8}$/.test(valor)) {
        return { allDay: true, data: valor.slice(0, 4) + '-' + valor.slice(4, 6) + '-' + valor.slice(6, 8) };
    }

    const m = valor.match(/^(\d{4})(\d{2})(\d{2})T(\d{2})(\d{2})(\d{2})(Z)?$/);
    if (!m) return null;

    if (m[7] === 'Z') {
        const utcIso = m[1] + '-' + m[2] + '-' + m[3] + 'T' + m[4] + ':' + m[5] + ':' + m[6] + 'Z';
        const d = new Date(utcIso);
        const local = d.toLocaleString('sv-SE', { timeZone: 'America/Sao_Paulo' }); // "2026-08-01 11:00:00"
        const partes = local.split(' ');
        return { allDay: false, data: partes[0], hora: (partes[1] || '00:00:00').slice(0, 5) };
    }
    // TZID explícito (ou sem timezone) — assume já em horário local
    return { allDay: false, data: m[1] + '-' + m[2] + '-' + m[3], hora: m[4] + ':' + m[5] };
}

function parseIcsEvent(blocoIcs) {
    const bloco = blocoIcs.match(/BEGIN:VEVENT([\s\S]*?)END:VEVENT/i);
    if (!bloco) return null;

    const linhas = unfold(bloco[1]).split(/\r\n|\n|\r/).map((l) => l.trim()).filter(Boolean);

    const statusLinha = acharLinha(linhas, 'STATUS');
    if (statusLinha && /CANCELLED/i.test(statusLinha)) return null;

    const uidLinha = acharLinha(linhas, 'UID');
    const dtStartLinha = acharLinha(linhas, 'DTSTART');
    if (!uidLinha || !dtStartLinha) return null;

    const ini = parseIcsDateTime(dtStartLinha);
    if (!ini) return null;

    const dtEndLinha = acharLinha(linhas, 'DTEND');
    const fimP = dtEndLinha ? parseIcsDateTime(dtEndLinha) : null;

    const summaryLinha = acharLinha(linhas, 'SUMMARY');
    const descLinha = acharLinha(linhas, 'DESCRIPTION');

    return {
        uid: valorDaLinha(uidLinha),
        assunto: summaryLinha ? unescapeIcsText(valorDaLinha(summaryLinha)) : '(sem assunto)',
        descricao: descLinha ? unescapeIcsText(valorDaLinha(descLinha)) : '',
        data: ini.data,
        horaInicio: ini.allDay ? '' : (ini.hora || ''),
        horaFim: (fimP && !fimP.allDay) ? (fimP.hora || '') : ''
    };
}
