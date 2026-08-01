/**
 * ponto-ui.js - Orquestração da aba de Ponto (jornada/RH)
 * Fase 1 do plano de migração: só a navegação interna (sub-abas) e
 * placeholders. Conteúdo real de cada sub-aba (Folha, Registros, Acordos,
 * Eventos, Férias) entra na Fase 3, reaproveitando PontoStore/PontoCalculos.
 *
 * Segue o mesmo padrão delegado do crm-ui.js (ver crm-ui.js:3051-3066), mas
 * com seu próprio atributo (`data-ponto-action`) e mapa de ações — o módulo é
 * dono do próprio clique, sem depender do listener do CRM.
 */
(function () {
    var esc = function (v) { return (window.Utils && Utils.escapeHtml) ? Utils.escapeHtml(v) : String(v == null ? '' : v); };

    var SUBABAS = [
        { chave: 'folha', rotulo: 'Folha de Ponto' },
        { chave: 'registros', rotulo: 'Registros' },
        { chave: 'acordos', rotulo: 'Acordos' },
        { chave: 'eventos', rotulo: 'Eventos' },
        { chave: 'ferias', rotulo: 'Férias' },
        { chave: 'migracao', rotulo: 'Migração' }
    ];

    var _subaba = 'folha';
    var _migrando = false;
    var _resultadoMigracao = null; // { registros, eventos, acordos, periodosAquisitivos, importadoEm } | null

    function renderizarPlaceholder(chave) {
        var rotulo = (SUBABAS.filter(function (s) { return s.chave === chave; })[0] || {}).rotulo || chave;
        return '<div class="ponto-placeholder">' +
            '<p>' + esc(rotulo) + ' — em construção (Fase 3 do plano de migração do Ponto).</p>' +
        '</div>';
    }

    function renderizarConteudoSubaba(chave) {
        switch (chave) {
            case 'folha': return renderizarPlaceholder('folha');
            case 'registros': return renderizarPlaceholder('registros');
            case 'acordos': return renderizarPlaceholder('acordos');
            case 'eventos': return renderizarPlaceholder('eventos');
            case 'ferias': return renderizarPlaceholder('ferias');
            case 'migracao': return renderizarMigracao();
            default: return renderizarPlaceholder(chave);
        }
    }

    // ──────────────────────────────────────────────
    //  MIGRAÇÃO (Fase 2) — importação única do Ponto original
    // ──────────────────────────────────────────────

    /**
     * Painel de migração: conecta no Firestore do Ponto (mesmo mecanismo de
     * app secundário do ponto-calendar-bridge.js, credenciais próprias do
     * Ponto) e importa o documento inteiro pra dentro de estoque.ponto. Uma
     * ação admin, pensada pra rodar uma vez só — não é um bridge permanente.
     */
    function renderizarMigracao() {
        var pb = window.PontoBridge;
        if (!pb || !pb.configurado()) {
            return '<div class="ponto-placeholder"><p>Firebase indisponível nesta página — não é possível conectar ao Ponto.</p></div>';
        }

        var aviso = '<p class="ponto-migracao-aviso">Isto lê o Firestore do sistema Ponto original (projeto <code>ponto-68b4a</code>) ' +
            'e sobrescreve <strong>estoque.ponto</strong> com os dados normalizados. É uma ação para rodar uma vez, com as ' +
            'credenciais que você já usa para entrar no Ponto — nada é escrito de volta lá.</p>';

        var conexao;
        if (pb.conectado()) {
            var erro = pb.getUltimoErro && pb.getUltimoErro();
            conexao = '<div class="ponto-migracao-conexao">' +
                '<span class="ponto-migracao-status-on">🟢 Conectado ao Ponto (' + esc(pb.usuarioEmail()) + ')</span>' +
                '<button type="button" class="btn-secondary crm-btn-mini" data-ponto-action="migracaoImportar"' +
                    (_migrando ? ' disabled' : '') + '>' + (_migrando ? 'Importando…' : 'Importar dados do Ponto') + '</button>' +
                '<button type="button" class="btn-secondary crm-btn-mini" data-ponto-action="migracaoDesconectar">Desconectar</button>' +
                (erro ? ('<div class="ponto-migracao-erro">⚠ ' + esc(erro) + '</div>') : '') +
            '</div>';
        } else {
            conexao = '<div class="ponto-migracao-conexao">' +
                '<input type="email" id="pontoMigracaoEmail" class="crm-filtro-input" placeholder="e-mail do Ponto" autocomplete="username">' +
                '<input type="password" id="pontoMigracaoSenha" class="crm-filtro-input" placeholder="senha" autocomplete="current-password" ' +
                    'onkeydown="if(event.key===\'Enter\'){event.preventDefault();PontoUI.migracaoConectar();}">' +
                '<button type="button" class="btn-secondary crm-btn-mini" data-ponto-action="migracaoConectar">Conectar Ponto</button>' +
            '</div>';
        }

        var resultado = '';
        if (_resultadoMigracao) {
            var r = _resultadoMigracao;
            resultado = '<div class="ponto-migracao-resultado">' +
                '<p>✅ Importado em ' + esc(r.importadoEm) + ':</p>' +
                '<ul>' +
                    '<li>' + r.registros + ' registro(s) de ponto</li>' +
                    '<li>' + r.eventos + ' evento(s)</li>' +
                    '<li>' + r.acordos + ' acordo(s)</li>' +
                    '<li>' + r.periodosAquisitivos + ' período(s) aquisitivo(s)</li>' +
                '</ul>' +
                '<p class="ponto-migracao-verificar">Confira uma amostra desses números contra o Ponto original antes de confiar neles ' +
                    '(Fase 2 do plano de migração pede essa validação antes de seguir para as views de leitura).</p>' +
            '</div>';
        }

        return aviso + conexao + resultado;
    }

    function migracaoConectar() {
        var email = ((document.getElementById('pontoMigracaoEmail') || {}).value || '').trim();
        var senha = (document.getElementById('pontoMigracaoSenha') || {}).value || '';
        PontoBridge.conectar(email, senha).then(function () {
            renderizar();
        }).catch(function (e) {
            if (window.Notifications) Notifications.error('Falha ao conectar ao Ponto: ' + (e && e.message ? e.message : e));
        });
    }

    function migracaoImportar() {
        if (typeof requireAdminOrNotify === 'function' && !requireAdminOrNotify()) return;
        var executar = function () {
            _migrando = true;
            renderizar();
            PontoBridge.buscarDadosCompletos().then(function (dadosBrutos) {
                var semCrm = Object.assign({}, dadosBrutos);
                delete semCrm.crm; // decisão do plano: a cópia de CRM do Ponto fica pra trás
                var normalizado = PontoStore.importarDadosMigrados(semCrm);
                _resultadoMigracao = {
                    registros: normalizado.registros.length,
                    eventos: normalizado.eventos.length,
                    acordos: normalizado.acordos.length,
                    periodosAquisitivos: normalizado.periodosAquisitivos.length,
                    importadoEm: new Date().toLocaleString('pt-BR')
                };
                _migrando = false;
                if (window.Notifications) Notifications.success('Dados do Ponto importados.');
                renderizar();
            }).catch(function (e) {
                _migrando = false;
                if (window.Notifications) Notifications.error('Falha ao importar dados do Ponto: ' + (e && e.message ? e.message : e));
                renderizar();
            });
        };

        if (window.Notifications && Notifications.confirm) {
            Notifications.confirm('Importar agora vai sobrescrever estoque.ponto com os dados do Ponto original. Continuar?', executar);
        } else {
            executar();
        }
    }

    function migracaoDesconectar() {
        if (window.PontoBridge) PontoBridge.desconectar();
        renderizar();
    }

    function renderizar() {
        ligarListenersUmaVez();
        if (window.PontoStore) PontoStore.ensurePontoDefault();

        var el = document.getElementById('pontoConteudo');
        if (!el) return;

        var abas = '<div class="ponto-subabas" role="tablist" aria-label="Seção de Ponto">' +
            SUBABAS.map(function (s) {
                return '<button type="button" class="ponto-subaba' + (s.chave === _subaba ? ' active' : '') + '" ' +
                    'role="tab" aria-selected="' + (s.chave === _subaba) + '" ' +
                    'data-ponto-action="trocarSubaba" data-valor="' + esc(s.chave) + '">' + esc(s.rotulo) + '</button>';
            }).join('') +
        '</div>';

        el.innerHTML = abas + '<div class="ponto-subaba-corpo">' + renderizarConteudoSubaba(_subaba) + '</div>';
    }

    var ACOES = {
        trocarSubaba: function (el) {
            var valido = SUBABAS.some(function (s) { return s.chave === el.dataset.valor; });
            _subaba = valido ? el.dataset.valor : 'folha';
            renderizar();
        },
        migracaoConectar: function () { migracaoConectar(); },
        migracaoDesconectar: function () { migracaoDesconectar(); },
        migracaoImportar: function () { migracaoImportar(); }
    };

    function aoClicarPonto(e) {
        var el = e.target.closest('[data-ponto-action]');
        if (!el) return;
        var fn = ACOES[el.dataset.pontoAction];
        if (fn) fn(el);
    }

    function ligarListenersUmaVez() {
        if (window.__pontoListenersLigados) return;
        window.__pontoListenersLigados = true;
        document.addEventListener('click', aoClicarPonto);
    }

    window.PontoUI = {
        renderizar: renderizar,
        migracaoConectar: migracaoConectar
    };
})();
