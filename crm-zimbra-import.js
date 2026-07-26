/**
 * crm-zimbra-import.js - Importação unidirecional (Zimbra → CRM) de eventos
 * do calendário do Zimbra para a aba Calendário do Relacionamento.
 *
 * Fluxo:
 *  1. Botão "Importar do Zimbra" chama o endpoint /api/zimbra-sync (Cloudflare
 *     Pages Function em functions/api/zimbra-sync.js), que fala com o
 *     servidor Zimbra via CalDAV (contorna CORS) e devolve os eventos em JSON.
 *  2. O navegador grava esses eventos numa coleção Firestore própria
 *     ("zimbra_eventos", uma por evento, id = UID do Zimbra) — NUNCA no
 *     documento principal app_data/latest, para não arriscar sobrescrever
 *     uma edição concorrente feita por outra pessoa no sistema.
 *  3. A aba Calendário lê essa coleção e mistura os eventos na lista/semana,
 *     marcados com o selo "📥 Zimbra" — são somente leitura (não editáveis
 *     pelo modal de atividade do CRM).
 *
 * Requer que o proxy (functions/api/zimbra-sync.js) esteja configurado com
 * as variáveis de ambiente ZIMBRA_CALDAV_URL / ZIMBRA_USER / ZIMBRA_PASSWORD
 * no Cloudflare Pages.
 */
(function () {
    'use strict';

    var COLECAO = 'zimbra_eventos';
    var _cache = [];
    var _carregando = false;
    var _listeners = [];

    function aoMudar(fn) {
        if (typeof fn === 'function') _listeners.push(fn);
    }

    function notificar() {
        _listeners.forEach(function (fn) { try { fn(); } catch (_) {} });
    }

    /**
     * Devolve os eventos já em cache (formato compatível com o de uma
     * atividade do CRM, para reaproveitar os renders/filtros existentes).
     */
    function listar() {
        return _cache.slice();
    }

    function carregando() {
        return _carregando;
    }

    function mapearParaAtividade(id, d) {
        return {
            id: 'zimbra_' + id,
            negocioId: null,
            tipo: 'reuniao',
            assunto: d.assunto || '(sem assunto)',
            descricao: d.descricao || '',
            data: d.data || null,
            horaInicio: d.horaInicio || '',
            horaFim: d.horaFim || '',
            feito: false,
            origemZimbra: true,
            zimbraUid: id
        };
    }

    /**
     * Recarrega o cache local a partir da coleção Firestore (sem contatar o
     * Zimbra) — leve, chamado sempre que a aba Calendário é aberta.
     */
    function carregarDoFirestore() {
        if (!window.firestoreDB) return Promise.resolve([]);
        return window.firestoreDB.collection(COLECAO).get().then(function (snap) {
            var lista = [];
            snap.forEach(function (doc) { lista.push(mapearParaAtividade(doc.id, doc.data() || {})); });
            _cache = lista;
            notificar();
            return _cache;
        }).catch(function (e) {
            console.warn('Falha ao carregar eventos do Zimbra do Firestore:', e);
            return _cache;
        });
    }

    /**
     * Dispara a importação de verdade: chama o proxy, que consulta o Zimbra,
     * e grava o resultado no Firestore. Usado pelo botão "Importar do Zimbra".
     */
    function importarAgora() {
        if (_carregando) return Promise.resolve(null);
        _carregando = true;
        notificar();

        return fetch('/api/zimbra-sync')
            .then(function (resp) {
                return resp.json().catch(function () { return {}; }).then(function (corpo) {
                    if (!resp.ok) throw new Error(corpo.erro || ('Falha ao importar do Zimbra (HTTP ' + resp.status + ').'));
                    return corpo;
                });
            })
            .then(function (corpo) {
                var eventos = corpo.eventos || [];
                if (!window.firestoreDB) throw new Error('Firestore não disponível — faça login primeiro.');
                if (!eventos.length) return 0;

                var lote = window.firestoreDB.batch();
                eventos.forEach(function (ev) {
                    var id = String(ev.uid || '').replace(/\//g, '_').slice(0, 400);
                    if (!id) return;
                    var ref = window.firestoreDB.collection(COLECAO).doc(id);
                    lote.set(ref, {
                        assunto: ev.assunto || '',
                        descricao: ev.descricao || '',
                        data: ev.data || null,
                        horaInicio: ev.horaInicio || '',
                        horaFim: ev.horaFim || '',
                        atualizadoEm: new Date().toISOString()
                    }, { merge: true });
                });
                return lote.commit().then(function () { return eventos.length; });
            })
            .then(function (total) {
                if (window.Notifications) Notifications.success('Importação do Zimbra concluída: ' + total + ' evento(s).');
                return carregarDoFirestore();
            })
            .catch(function (e) {
                if (window.Notifications) Notifications.error('Falha ao importar do Zimbra: ' + e.message);
                throw e;
            })
            .then(
                function (r) { _carregando = false; notificar(); return r; },
                function (e) { _carregando = false; notificar(); throw e; }
            );
    }

    window.ZimbraImport = {
        listar: listar,
        carregando: carregando,
        carregarDoFirestore: carregarDoFirestore,
        importarAgora: importarAgora,
        aoMudar: aoMudar
    };
})();
