/**
 * ponto-calendar-bridge.js - Ponte de LEITURA entre o calendário do Relacionamento
 * (CRM) e o sistema Ponto (https://github.com/joffreribeiro/Ponto), que roda num
 * projeto Firebase totalmente separado ("ponto-68b4a"), com sua própria base de
 * usuários. Não há como uma consulta Firestore comum enxergar os dois projetos —
 * por isso este módulo abre uma SEGUNDA instância nomeada do Firebase (a mesma
 * técnica usada por multi-tenant apps), autenticada de forma independente da
 * sessão principal do Estoque.
 *
 * Só LÊ. Nunca escreve no Firestore do Ponto — nenhuma função de escrita é sequer
 * exposta aqui. Os dados trazidos (Eventos, que no Ponto também cobrem Férias —
 * ver tipoEvento === 'ferias') viram pseudo-atividades no calendário do CRM,
 * igual ao que já acontece com feriados nacionais (ver gerarFeriadosPseudoAtividades
 * em crm-ui.js): geradas na hora, nunca gravadas em estoque.crm, nunca editáveis
 * por aqui.
 *
 * Autenticação: o Ponto usa e-mail/senha (não OAuth do Google), então a conexão
 * pede as mesmas credenciais que você já usa para entrar no Ponto. A sessão fica
 * só neste navegador (padrão de persistência do Firebase Auth) — não é gravada
 * em lugar nenhum por este módulo.
 */
(function () {
    'use strict';

    // Config pública do projeto Firebase do Ponto (mesma exposta no próprio
    // firebase-init.js do Ponto — não é segredo, é a identificação do app).
    var PONTO_CONFIG = {
        apiKey: 'AIzaSyC3euGrayxErDKDJvHT5WN6ixkeqTwwp8M',
        authDomain: 'ponto-68b4a.firebaseapp.com',
        projectId: 'ponto-68b4a',
        storageBucket: 'ponto-68b4a.firebasestorage.app',
        messagingSenderId: '516394911771',
        appId: '1:516394911771:web:89c861d000cb95a594d0be'
    };
    var APP_NAME = 'pontoBridge';
    var CACHE_MS = 5 * 60 * 1000; // evita reler o Firestore a cada re-render do calendário

    var _app = null, _auth = null, _db = null;
    var _listeners = [];
    var _cache = null; // { eventos: [...], buscadoEm: <timestamp> }
    var _buscando = false;
    var _ultimoErro = null; // string amigável do último erro de busca, ou null

    function getUltimoErro() { return _ultimoErro; }

    function notificar() {
        _listeners.forEach(function (fn) { try { fn(); } catch (_) {} });
    }
    function aoMudarStatus(fn) {
        if (typeof fn === 'function') _listeners.push(fn);
    }

    function bibliotecaCarregada() {
        return !!(window.firebase && firebase.initializeApp && firebase.auth && firebase.firestore);
    }
    function configurado() {
        return bibliotecaCarregada();
    }

    function garantirApp() {
        if (_app) return _app;
        if (!bibliotecaCarregada()) return null;
        try {
            var existente = (firebase.apps || []).filter(function (a) { return a.name === APP_NAME; })[0];
            _app = existente || firebase.initializeApp(PONTO_CONFIG, APP_NAME);
            _auth = firebase.auth(_app);
            _db = firebase.firestore(_app);
            _auth.onAuthStateChanged(function () { _cache = null; notificar(); });
        } catch (e) {
            console.error('Ponte com o Ponto: falha ao inicializar app secundário do Firebase.', e);
            _app = null;
        }
        return _app;
    }

    function conectado() {
        garantirApp();
        return !!(_auth && _auth.currentUser);
    }

    function usuarioEmail() {
        return (_auth && _auth.currentUser) ? (_auth.currentUser.email || '') : '';
    }

    /** Mesmas credenciais de e-mail/senha usadas para entrar no Ponto. */
    function conectar(email, senha) {
        if (!garantirApp()) return Promise.reject(new Error('Firebase indisponível nesta página.'));
        if (!email || !senha) return Promise.reject(new Error('Informe e-mail e senha do Ponto.'));
        return _auth.signInWithEmailAndPassword(email, senha).then(function (cred) {
            notificar();
            return cred;
        });
    }

    function desconectar() {
        if (_auth) { try { _auth.signOut(); } catch (_) {} }
        _cache = null;
        _ultimoErro = null;
        notificar();
    }

    /** Leitura sem rede — o que já estiver em cache (pode estar vazio/desatualizado). */
    function getEventosCache() {
        return (_cache && _cache.eventos) || [];
    }

    /**
     * O Ponto às vezes acumula eventos EXATAMENTE duplicados (mesmo tipo,
     * descrição e intervalo de datas — ex: a mesma Férias salva duas vezes).
     * Mantém só a primeira ocorrência de cada assinatura, para não mostrar a
     * mesma coisa repetida no calendário do CRM.
     */
    function deduplicarEventos(eventos) {
        var vistos = {};
        return (eventos || []).filter(function (ev) {
            var chave = [ev.tipoEvento, ev.descricaoEvento, ev.dataInicioEvento, ev.dataFimEvento].join('|');
            if (vistos[chave]) return false;
            vistos[chave] = true;
            return true;
        });
    }

    /**
     * Busca `ponto/{uid}.dados.eventos` no Firestore do Ponto. Cacheia por
     * CACHE_MS para não bater no Firestore a cada renderização do calendário;
     * `forcar` ignora o cache (botão "Atualizar").
     */
    function atualizarEventos(forcar) {
        if (!conectado()) return Promise.resolve([]);
        if (!forcar && _cache && (Date.now() - _cache.buscadoEm) < CACHE_MS) {
            return Promise.resolve(_cache.eventos);
        }
        if (_buscando) return Promise.resolve(getEventosCache());
        _buscando = true;

        var uid = _auth.currentUser.uid;
        return _db.collection('ponto').doc(uid).get().then(function (snap) {
            _buscando = false;
            _ultimoErro = (snap && snap.exists) ? null : 'Nenhum documento encontrado em ponto/' + uid + ' — verifique se é a mesma conta usada no Ponto.';
            var dados = (snap && snap.exists) ? (snap.data().dados || {}) : {};
            var eventos = deduplicarEventos(Array.isArray(dados.eventos) ? dados.eventos : []);
            _cache = { eventos: eventos, buscadoEm: Date.now() };
            notificar();
            return eventos;
        }).catch(function (e) {
            _buscando = false;
            console.error('Falha ao buscar eventos do Ponto:', e);
            _ultimoErro = (e && e.code === 'permission-denied')
                ? 'Sem permissão para ler os dados do Ponto (confira o login e as regras do Firestore).'
                : ('Falha ao buscar dados do Ponto: ' + (e.message || e));
            if (window.Notifications) Notifications.error(_ultimoErro);
            notificar();
            return getEventosCache();
        });
    }

    window.PontoBridge = {
        configurado: configurado,
        conectado: conectado,
        usuarioEmail: usuarioEmail,
        conectar: conectar,
        desconectar: desconectar,
        aoMudarStatus: aoMudarStatus,
        getEventosCache: getEventosCache,
        atualizarEventos: atualizarEventos,
        getUltimoErro: getUltimoErro
    };
})();
