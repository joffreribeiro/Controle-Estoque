/**
 * google-calendar-sync.js - Sincronização bidirecional entre as atividades do
 * Relacionamento (CRM) e o Google Calendar.
 *
 * Requer:
 *  - A biblioteca "Google Identity Services" carregada antes deste arquivo
 *    (<script src="https://accounts.google.com/gsi/client" async defer></script>).
 *  - Um Client ID OAuth 2.0 ("Aplicativo da Web") do Google Cloud Console,
 *    com a Calendar API habilitada e este domínio (e localhost, em dev)
 *    cadastrado em "Authorized JavaScript origins". Cole o Client ID abaixo.
 *
 * Não guarda nenhum segredo: o Client ID é público (identifica o app, não
 * autentica nada sozinho) — a restrição de segurança vem da origem cadastrada.
 * O token de acesso obtido fica só em memória (nunca é persistido) e expira
 * em ~1h; a cada sessão do navegador é preciso conectar de novo.
 *
 * Política de sincronização:
 *  CRM → Google:
 *   - atividade criada/editada (não concluída, com data) → cria/atualiza evento no Google;
 *   - atividade marcada concluída → evento é removido do Google (mantém a agenda limpa);
 *   - atividade excluída no CRM → evento é removido do Google.
 *  Google → CRM:
 *   - ao conectar e ao clicar "Sincronizar tudo", eventos criados direto no
 *     Google Agenda (que ainda não têm uma atividade correspondente no CRM)
 *     são importados como novas atividades (sem negócio vinculado — quem
 *     importou pode depois associar a um negócio pela tela normal);
 *   - eventos que o próprio CRM criou não são reimportados (identificados
 *     pela marca extendedProperties.private.crmAtividadeId);
 *   - depois de importado uma vez, o evento vira uma atividade normal do CRM
 *     e passa a seguir a política CRM → Google acima (fonte da verdade passa
 *     a ser o CRM a partir daí, para evitar loops/edições divergentes).
 */
(function () {
    'use strict';

    // ── CONFIGURAÇÃO — troque pelo Client ID gerado no Google Cloud Console ──
    var CLIENT_ID = '339770116384-ng8nr2da6lla6sgk0enti1vd02b8j14q.apps.googleusercontent.com';
    var SCOPE = 'https://www.googleapis.com/auth/calendar.events https://www.googleapis.com/auth/calendar.readonly';
    var TIMEZONE = 'America/Sao_Paulo';
    var EVENTS_URL = 'https://www.googleapis.com/calendar/v3/calendars/primary/events';
    // colorId dos eventos criados pelo CRM, para diferenciar visualmente na Google Agenda
    // dos eventos criados manualmente. Paleta oficial do Google Calendar (1-11):
    // 1 Lavanda, 2 Sálvia, 3 Uva, 4 Flamingo, 5 Banana, 6 Tangerina,
    // 7 Pavão, 8 Grafite, 9 Mirtilo, 10 Manjericão, 11 Tomate.
    var EVENT_COLOR_ID = '3'; // Uva (roxo) — combina com a cor do workspace Relacionamento no app

    var _tokenClient = null;
    var _accessToken = null;
    var _tokenExpiraEm = 0;
    var _listeners = [];

    function configurado() {
        return !!CLIENT_ID && CLIENT_ID.indexOf('COLOQUE_SEU_CLIENT_ID') === -1;
    }

    function conectado() {
        return !!_accessToken && Date.now() < _tokenExpiraEm;
    }

    function notificar() {
        _listeners.forEach(function (fn) { try { fn(); } catch (_) {} });
    }

    function aoMudarStatus(fn) {
        if (typeof fn === 'function') _listeners.push(fn);
    }

    function bibliotecaCarregada() {
        return !!(window.google && google.accounts && google.accounts.oauth2);
    }

    function garantirTokenClient() {
        if (_tokenClient) return _tokenClient;
        if (!bibliotecaCarregada()) return null;
        _tokenClient = google.accounts.oauth2.initTokenClient({
            client_id: CLIENT_ID,
            scope: SCOPE,
            callback: function (resp) {
                if (resp && resp.access_token) {
                    _accessToken = resp.access_token;
                    _tokenExpiraEm = Date.now() + ((resp.expires_in || 3500) * 1000);
                    if (window.Notifications) Notifications.success('Google Agenda conectado.');
                    notificar();
                    importarNovosDoGoogle().then(function (novos) {
                        if (novos && window.Notifications) Notifications.success(novos + ' evento(s) importado(s) do Google Agenda.');
                        notificar();
                    });
                } else {
                    if (window.Notifications) Notifications.error('Não foi possível conectar ao Google Agenda.');
                    notificar();
                }
            }
        });
        return _tokenClient;
    }

    function conectar() {
        if (!configurado()) {
            if (window.Notifications) Notifications.error('Google Agenda ainda não configurado (falta o Client ID em google-calendar-sync.js).');
            return;
        }
        var tc = garantirTokenClient();
        if (!tc) {
            if (window.Notifications) Notifications.error('Biblioteca do Google ainda não carregou. Tente novamente em instantes.');
            return;
        }
        tc.requestAccessToken({ prompt: conectado() ? '' : 'consent' });
    }

    function desconectar() {
        if (_accessToken && bibliotecaCarregada()) {
            try { google.accounts.oauth2.revoke(_accessToken, function () {}); } catch (_) {}
        }
        _accessToken = null;
        _tokenExpiraEm = 0;
        notificar();
    }

    function chamarApi(metodo, url, corpo) {
        return fetch(url, {
            method: metodo,
            headers: {
                'Authorization': 'Bearer ' + _accessToken,
                'Content-Type': 'application/json'
            },
            body: corpo ? JSON.stringify(corpo) : undefined
        }).then(function (resp) {
            if (!resp.ok) {
                return resp.text().then(function (t) {
                    var erro = new Error('Google Calendar API (' + resp.status + '): ' + t);
                    erro.status = resp.status;
                    throw erro;
                });
            }
            if (resp.status === 204) return null;
            return resp.json();
        });
    }

    function somarMinutos(hhmm, minAdd) {
        var partes = String(hhmm).split(':');
        var totalMin = Number(partes[0]) * 60 + Number(partes[1]) + minAdd;
        var h = Math.floor((totalMin % (24 * 60)) / 60), m = totalMin % 60;
        return String(h).padStart(2, '0') + ':' + String(m).padStart(2, '0');
    }

    function somarDiaIso(dataIso) {
        var d = new Date(dataIso + 'T00:00:00Z');
        d.setUTCDate(d.getUTCDate() + 1);
        return d.toISOString().slice(0, 10);
    }

    function montarEvento(atividade, negocio, cliente) {
        var tipoInfo = (window.CrmModel && CrmModel.TIPOS_ATIVIDADE[atividade.tipo]) || { icone: '📌', rotulo: atividade.tipo };
        var titulo = tipoInfo.icone + ' ' + (atividade.assunto || '(sem assunto)') + (negocio ? (' — ' + negocio.titulo) : '');

        var descPartes = [];
        if (atividade.descricao) descPartes.push(atividade.descricao);
        if (negocio) descPartes.push('Negócio: ' + (negocio.titulo || ''));
        if (cliente) descPartes.push('Cliente: ' + (cliente.nome || ''));
        descPartes.push('— Sincronizado automaticamente do Controle de Estoque (CRM). Edições feitas aqui no Google não retornam ao sistema.');

        var evento = {
            summary: titulo,
            description: descPartes.join('\n'),
            colorId: EVENT_COLOR_ID,
            extendedProperties: { private: { crmAtividadeId: atividade.id } }
        };

        if (atividade.horaInicio) {
            var fim = atividade.horaFim || somarMinutos(atividade.horaInicio, 30);
            evento.start = { dateTime: atividade.data + 'T' + atividade.horaInicio + ':00', timeZone: TIMEZONE };
            evento.end = { dateTime: atividade.data + 'T' + fim + ':00', timeZone: TIMEZONE };
        } else {
            evento.start = { date: atividade.data };
            evento.end = { date: somarDiaIso(atividade.data) };
        }
        return evento;
    }

    /**
     * Cria, atualiza ou remove (se concluída) o evento correspondente no Google.
     * Grava o googleEventId de volta na atividade via CrmStore quando muda.
     */
    function sincronizarAtividade(atividade, negocio, cliente) {
        if (!configurado() || !conectado() || !atividade) return Promise.resolve(null);
        if (atividade.feito) return removerEvento(atividade);
        if (!atividade.data) return Promise.resolve(null);

        var corpo = montarEvento(atividade, negocio, cliente);

        if (atividade.googleEventId) {
            return chamarApi('PATCH', EVENTS_URL + '/' + atividade.googleEventId, corpo)
                .catch(function (e) {
                    // Só recria se o Google confirmar que o evento não existe mais (404).
                    // Qualquer outra falha (403, evento gerenciado por outra fonte, rede, etc.)
                    // é reportada como erro de verdade — nunca cria um duplicado às cegas.
                    if (e && e.status === 404) return chamarApi('POST', EVENTS_URL, corpo);
                    throw e;
                })
                .then(function (ev) {
                    if (ev && ev.id && ev.id !== atividade.googleEventId && window.CrmStore) {
                        CrmStore.atualizarAtividade(atividade.id, { googleEventId: ev.id });
                    }
                    return ev;
                });
        }

        return chamarApi('POST', EVENTS_URL, corpo).then(function (ev) {
            if (ev && ev.id && window.CrmStore) CrmStore.atualizarAtividade(atividade.id, { googleEventId: ev.id });
            return ev;
        });
    }

    function removerEvento(atividade) {
        if (!configurado() || !conectado() || !atividade || !atividade.googleEventId) return Promise.resolve(null);
        var id = atividade.googleEventId;
        return chamarApi('DELETE', EVENTS_URL + '/' + id)
            .catch(function () { /* evento já pode ter sido removido no Google */ })
            .then(function () {
                if (window.CrmStore) CrmStore.atualizarAtividade(atividade.id, { googleEventId: null });
            });
    }

    // ── Google → CRM (importação de eventos criados direto na Google Agenda) ──

    var JANELA_IMPORT_DIAS_PASSADO = 7;
    var JANELA_IMPORT_DIAS_FUTURO = 60;

    function googleDataHora(pontoEvento) {
        if (!pontoEvento) return null;
        if (pontoEvento.date) return { allDay: true, data: pontoEvento.date };
        if (pontoEvento.dateTime) {
            var d = new Date(pontoEvento.dateTime);
            if (isNaN(d.getTime())) return null;
            var local = d.toLocaleString('sv-SE', { timeZone: TIMEZONE }); // "2026-08-01 11:00:00"
            var partes = local.split(' ');
            return { allDay: false, data: partes[0], hora: (partes[1] || '00:00:00').slice(0, 5) };
        }
        return null;
    }

    /**
     * Reservas de voo/hotel etc. que o Gmail detecta e injeta automaticamente
     * na Google Agenda não são eventos "de verdade" — não aceitam edição via
     * PATCH normal (causou duplicidade quando o CRM tentou reenviar). Detecta
     * pela assinatura típica desses itens (campo `source` apontando pro Gmail).
     */
    function ehEventoAutoDetectadoPeloGmail(ev) {
        var origem = ev && ev.source && ev.source.url;
        return !!(origem && /mail\.google\.com/i.test(origem));
    }

    function mapearEventoGoogleParaAtividade(ev) {
        if (!ev || ev.status === 'cancelled') return null;
        var ini = googleDataHora(ev.start);
        if (!ini) return null;
        var fim = googleDataHora(ev.end);
        return {
            tipo: 'reuniao',
            assunto: ev.summary || '(sem assunto)',
            descricao: ev.description || '',
            data: ini.data,
            horaInicio: ini.allDay ? '' : (ini.hora || ''),
            horaFim: (fim && !fim.allDay) ? (fim.hora || '') : '',
            googleEventId: ev.id,
            origemGoogle: true
        };
    }

    /**
     * Busca eventos do Google (janela de -7 a +60 dias) e cria como novas
     * atividades no CRM os que ainda não existem por lá — pulando tanto os
     * que o próprio CRM criou (marca crmAtividadeId) quanto os já importados
     * antes (já existe uma atividade com esse googleEventId).
     */
    function importarNovosDoGoogle() {
        if (!configurado() || !conectado() || !window.CrmStore) return Promise.resolve(0);

        var agora = new Date();
        var min = new Date(agora.getTime() - JANELA_IMPORT_DIAS_PASSADO * 86400000).toISOString();
        var max = new Date(agora.getTime() + JANELA_IMPORT_DIAS_FUTURO * 86400000).toISOString();
        var url = EVENTS_URL + '?singleEvents=true&orderBy=startTime&maxResults=250' +
            '&timeMin=' + encodeURIComponent(min) + '&timeMax=' + encodeURIComponent(max);

        return chamarApi('GET', url).then(function (resp) {
            var itens = (resp && resp.items) || [];
            var googleIdsExistentes = {};
            CrmStore.listarAtividades().forEach(function (a) {
                if (a.googleEventId) googleIdsExistentes[a.googleEventId] = true;
            });

            var novos = 0;
            itens.forEach(function (ev) {
                var props = ev.extendedProperties && ev.extendedProperties.private;
                if (props && props.crmAtividadeId) return; // criado pelo próprio CRM
                if (googleIdsExistentes[ev.id]) return; // já importado antes
                if (ehEventoAutoDetectadoPeloGmail(ev)) return; // reserva de voo/hotel etc. detectada pelo Gmail — não é editável de verdade

                var dados = mapearEventoGoogleParaAtividade(ev);
                if (!dados) return;
                CrmStore.criarAtividade(dados);
                novos++;
            });
            return novos;
        }).catch(function (e) {
            if (window.Notifications) Notifications.error('Falha ao importar eventos do Google Agenda: ' + e.message);
            return 0;
        });
    }

    /**
     * Importa novidades do Google e depois varre todas as atividades do CRM,
     * sincronizando uma a uma (sequencial, para respeitar limites de taxa da
     * API). Usado pelo botão "Sincronizar tudo".
     */
    function sincronizarTudo() {
        if (!configurado()) {
            if (window.Notifications) Notifications.error('Configure o Client ID do Google em google-calendar-sync.js primeiro.');
            return;
        }
        if (!conectado()) { conectar(); return; }
        if (!window.CrmStore) return;

        importarNovosDoGoogle().then(function (novosImportados) {
            var atividades = CrmStore.listarAtividades().filter(function (a) { return a.data; });
            var negocios = CrmStore.listarNegocios();
            var negocioPorId = {};
            negocios.forEach(function (n) { negocioPorId[n.id] = n; });

            var fila = atividades.slice();
            var ok = 0, falha = 0;

            function proximo() {
                if (!fila.length) {
                    if (window.Notifications) {
                        Notifications.success('Sincronização com o Google Agenda concluída: ' + ok + ' enviada(s)' +
                            (novosImportados ? (', ' + novosImportados + ' importada(s)') : '') +
                            (falha ? (', ' + falha + ' falha(s)') : '') + '.');
                    }
                    notificar();
                    return;
                }
                var a = fila.shift();
                var negocio = negocioPorId[a.negocioId];
                var cliente = negocio ? CrmStore.getCliente(negocio.clienteId) : null;
                sincronizarAtividade(a, negocio, cliente)
                    .then(function () { ok++; proximo(); })
                    .catch(function () { falha++; proximo(); });
            }
            proximo();
        });
    }

    window.GoogleCalendarSync = {
        configurado: configurado,
        conectado: conectado,
        conectar: conectar,
        desconectar: desconectar,
        aoMudarStatus: aoMudarStatus,
        sincronizarAtividade: sincronizarAtividade,
        removerEvento: removerEvento,
        sincronizarTudo: sincronizarTudo,
        importarNovosDoGoogle: importarNovosDoGoogle
    };
})();
