/**
 * crm-ui.js - Orquestração e renderização da aba Relacionamento (CRM) no Controle-Estoque
 *
 * Depende de (globais já carregados antes deste arquivo):
 *  - CrmModel, CrmCalculos, CrmStore, CrmKanban (módulos do CRM)
 *  - Utils.escapeHtml, Notifications, DateUtils (crm-compat.js)
 *  - requireAdminOrNotify, fecharModal (app2.js)
 *
 * Estado de visualização (busca, visão, aba do detalhe etc.) fica em variáveis
 * de módulo — não é persistido em estoque.crm.config a cada interação, para não
 * disparar salvamentos (local + Firestore) a cada tecla digitada.
 */
(function () {
    function esc(v) { return window.Utils.escapeHtml(v); }

    // ── Estado de sessão da aba (não persistido) ──
    var _secao = 'negocios'; // 'negocios' | 'calendario' (seções de topo dentro de Relacionamento)
    var _visao = 'kanban';
    var _busca = '';
    var _mostrarFechados = false;
    var _ordenarPor = 'ordem';
    var _detalheId = null;
    var _abaDetalhe = 'atividade';
    var _filtroHistorico = 'todos';
    var _secoesColapsadas = {};
    var _atividadeEditandoId = null;
    var _itensTemp = [];
    var _dragId = null;

    // ── Estado de filtros da vista Lista (Anotações) ──
    var _crmListaFiltros = {
        busca: '',
        funilId: null,
        clienteId: null,
        prioridade: null,
        origemDemanda: null,
        prazoInicio: null,
        prazoFim: null,
        dataSolicitacaoInicio: null,
        dataSolicitacaoFim: null,
        mostrarFinalizadas: true,
        situacao: null,
        responsavel: null,
        encaminhadoPara: null,
        somenteVencidas: false,
        somenteAtrasoTerceiro: false,
        ordenarPor: 'prazo',
        ordenarDireccao: 'asc'
    };
    var _crmListaPagina = 1;
    var _crmListaPaginaSize = 50;
    var _anotacaoNegocioId = null; // negócio selecionado no modal de anotação
    var _anotacaoRelacionadaId = null; // anotação-pai (thread) selecionada no modal
    var _encaminhamentosTemp = []; // encaminhamentos em edição no modal (só gravados ao salvar)
    var _negocioTrocandoFunil = null; // negócio cuja troca de funil está em progresso

    // ── Estado da Caixa de Demandas ──
    var _demandaAba = 'comigo'; // comigo | aguardando | consolidar | concluidas | todas

    // ── Estado da vista Calendário (Atividades, todos os negócios) ──
    var _calModo = 'mes'; // 'lista' | 'semana' | 'mes'
    var _calPeriodo = 'todos'; // todos|paraFazer|vencido|hoje|amanha|semana|proximaSemana
    var _calTipo = ''; // '' (todos) | chamada|reuniao|tarefa|prazo|email|viagem|ferias|recesso|particular
    var _calBusca = '';
    var _calSemanaRef = CrmCalculos.inicioSemana(hojeIsoLocal()); // domingo da semana exibida na vista Calendário
    var _calMesRef = hojeIsoLocal().slice(0, 7) + '-01'; // 1º dia do mês exibido na vista Mês
    var _calMostrarFeriados = false; // filtro da Lista: mostrar/ocultar feriados nacionais
    var _calMostrarPonto = true; // filtro: mostrar/ocultar Eventos/Férias importados do Ponto
    var _pontoAutoFetchFeito = false; // garante no máx. 1 busca automática por carregamento de página
    var _calAtvNegocioId = null; // negócio selecionado no modal de atividade global

    var ICONES_HISTORICO = { criacao: '✨', campo: '✎', etapa: '➡️', exclusao: '🗑', atividade: '📅', nota: '📝', anotacao: '🗒️' };
    var PRIORIDADE_ROTULO = { baixa: 'Baixa', media: 'Média', alta: 'Alta', critico: 'Crítico' };
    var SITUACAO_ROTULO = {
        recebida: 'Recebida', comigo: 'Comigo', aguardando_terceiro: 'Aguardando terceiro',
        consolidando: 'Consolidando', respondida: 'Respondida'
    };
    var STATUS_ENC_ROTULO = { pendente: 'Pendente', respondido: 'Respondido', cancelado: 'Cancelado' };
    var DEMANDA_ABAS = [
        ['comigo', 'Comigo'], ['aguardando', 'Aguardando terceiros'],
        ['consolidar', 'Para consolidar'], ['concluidas', 'Concluídas'], ['todas', 'Todas']
    ];
    var MESES = ['janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho', 'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro'];

    // ──────────────────────────────────────────────
    //  ENTRADA
    // ──────────────────────────────────────────────

    function renderizar() {
        ligarListenersUmaVez();
        popularSelectFunil();
        var chkFechados = document.getElementById('crmMostrarFechados');
        var selOrdenar = document.getElementById('crmOrdenarPor');
        var inpBusca = document.getElementById('crmBusca');
        if (chkFechados) chkFechados.checked = _mostrarFechados;
        if (selOrdenar) selOrdenar.value = _ordenarPor;
        if (inpBusca) inpBusca.value = _busca;
        renderizarConteudoAtivo();
    }

    function ligarListenersUmaVez() {
        if (window.__crmListenersLigados) return;
        window.__crmListenersLigados = true;
        document.addEventListener('click', aoClicar);
        document.addEventListener('cliente:salvo', aoSalvarClienteExterno);
        if (window.GoogleCalendarSync) {
            GoogleCalendarSync.aoMudarStatus(function () { if (_secao === 'calendario') renderizarCalendarioView(); });
        }
        if (window.PontoBridge) {
            PontoBridge.aoMudarStatus(function () {
                if (!PontoBridge.conectado()) _pontoAutoFetchFeito = false;
                if (_secao === 'calendario') renderizarCalendarioView();
            });
        }
    }

    function popularSelectFunil() {
        var sel = document.getElementById('crmFunilSelect');
        if (!sel) return;
        var crm = CrmStore.getCrm();
        var ativos = crm.funis.filter(function (f) { return !f.arquivado; });
        sel.innerHTML = ativos.map(function (f) { return '<option value="' + esc(f.id) + '">' + esc(f.nome) + '</option>'; }).join('');
        var funilAtivo = CrmStore.getFunilAtivo();
        if (funilAtivo) sel.value = funilAtivo.id;
    }

    function trocarFunil(id) {
        CrmStore.setFunilAtivo(id);
        _detalheId = null;
        renderizarConteudoAtivo();
    }

    function setBusca(v) { _busca = v; renderizarConteudoAtivo(); }
    function setMostrarFechados(v) { _mostrarFechados = v; renderizarConteudoAtivo(); }
    function setOrdenarPor(v) { _ordenarPor = v; renderizarConteudoAtivo(); }

    // ──────────────────────────────────────────────
    //  RENDER PRINCIPAL
    // ──────────────────────────────────────────────

    function renderizarConteudoAtivo() {
        var secaoNegocios = document.getElementById('crmSecaoNegocios');
        var secaoCalendario = document.getElementById('crmSecaoCalendario');
        if (!secaoNegocios) return;

        document.querySelectorAll('.crm-secao-btn').forEach(function (b) {
            b.classList.toggle('active', b.dataset.valor === _secao);
        });
        secaoNegocios.style.display = _secao === 'negocios' ? '' : 'none';
        if (secaoCalendario) secaoCalendario.style.display = _secao === 'calendario' ? '' : 'none';

        if (_secao === 'calendario') { renderizarCalendarioView(); return; }

        var kanban = document.getElementById('crmKanban');
        var lista = document.getElementById('crmListaNegocios');
        var demandas = document.getElementById('crmDemandas');
        var previsao = document.getElementById('crmPrevisao');
        var excluidos = document.getElementById('crmExcluidos');
        var detalhe = document.getElementById('crmViewDetalhe');
        var barra = document.querySelector('.crm-barra-visoes');
        if (!kanban) return;

        if (_detalheId) {
            [kanban, lista, demandas, previsao, excluidos].forEach(function (el) { if (el) el.style.display = 'none'; });
            if (barra) barra.style.display = 'none';
            detalhe.style.display = 'block';
            renderizarDetalhe(_detalheId);
            return;
        }

        if (barra) barra.style.display = '';
        detalhe.style.display = 'none';
        kanban.style.display = _visao === 'kanban' ? '' : 'none';
        lista.style.display = _visao === 'lista' ? '' : 'none';
        if (demandas) demandas.style.display = _visao === 'demandas' ? '' : 'none';
        previsao.style.display = _visao === 'previsao' ? '' : 'none';
        excluidos.style.display = _visao === 'excluidos' ? '' : 'none';

        // A barra de filtros de demandas serve tanto à Caixa quanto à Lista completa.
        var ehVisaoDemanda = (_visao === 'lista' || _visao === 'demandas');
        var crmListaFiltros = document.getElementById('crmListaFiltros');
        if (crmListaFiltros) crmListaFiltros.style.display = ehVisaoDemanda ? '' : 'none';

        document.querySelectorAll('.crm-visao-btn').forEach(function (b) {
            b.classList.toggle('active', b.dataset.valor === _visao);
        });

        var funil = CrmStore.getFunilAtivo();
        if (!funil) return;

        if (_visao === 'kanban') renderizarKanban(funil);
        else if (_visao === 'demandas') renderizarDemandasView();
        else if (_visao === 'lista') renderizarListaView();
        else if (_visao === 'previsao') renderizarPrevisaoView(funil);
        else if (_visao === 'excluidos') renderizarExcluidosView(funil);
    }

    function setSecao(v) {
        _secao = v;
        _detalheId = null;
        renderizarConteudoAtivo();
    }

    function negociosVisiveis(funil) {
        var todos = CrmStore.listarNegocios(funil.id);
        if (!_mostrarFechados) todos = todos.filter(function (n) { return n.status === 'aberto'; });
        var filtrados = CrmCalculos.filtrarNegocios(todos, { busca: _busca });

        if (_ordenarPor === 'proxima') {
            var atividades = CrmStore.listarAtividades();
            filtrados = filtrados.slice().sort(function (a, b) {
                var pa = CrmCalculos.proximaAtividade(atividades, a.id);
                var pb = CrmCalculos.proximaAtividade(atividades, b.id);
                var da = pa ? (pa.data + (pa.horaInicio || '')) : '9999';
                var db = pb ? (pb.data + (pb.horaInicio || '')) : '9999';
                return da.localeCompare(db);
            });
        } else {
            filtrados = CrmCalculos.ordenarNegocios(filtrados, _ordenarPor);
        }
        return filtrados;
    }

    function atualizarContagem(n, singular, plural) {
        var el = document.getElementById('crmContagem');
        var s = singular || 'negócio';
        var p = plural || (s + 's');
        if (el) el.textContent = n + ' ' + (n !== 1 ? p : s);
    }

    // ── Kanban ──

    function renderizarKanban(funil) {
        var negocios = negociosVisiveis(funil);
        var atividades = CrmStore.listarAtividades();
        document.getElementById('crmKanban').innerHTML = CrmKanban.renderizarBoard(funil, negocios, { atividades: atividades });
        atualizarContagem(negocios.length);
        ligarDragDrop();
    }

    function ligarDragDrop() {
        document.querySelectorAll('#crmKanban .crm-card').forEach(function (card) {
            card.addEventListener('dragstart', function (e) {
                _dragId = card.dataset.crmNegocioId;
                try { e.dataTransfer.effectAllowed = 'move'; e.dataTransfer.setData('text/plain', _dragId); } catch (_) {}
            });
        });
        document.querySelectorAll('#crmKanban .kanban-list').forEach(function (zone) {
            zone.addEventListener('dragover', function (e) { e.preventDefault(); });
            zone.addEventListener('drop', function (e) {
                e.preventDefault();
                var id = _dragId || (e.dataTransfer && e.dataTransfer.getData('text/plain'));
                _dragId = null;
                if (!id) return;
                if (!requireAdminOrNotify()) return;
                CrmStore.moverNegocio(id, zone.dataset.crmEtapaId, null);
                renderizarConteudoAtivo();
            });
        });
    }

    // ── Caixa de Demandas ──

    /**
     * Contexto compartilhado pelas visões de demanda: índices de negócio/cliente/
     * funil e a lista já filtrada e ordenada pelos filtros da barra.
     */
    function contextoDemandas() {
        var anotacoes = CrmStore.listarAnotacoes();
        var negocios = CrmStore.listarNegocios();
        var clientes = CrmStore.listarClientes();
        var funis = CrmStore.listarFunis();
        var extra = { negocios: negocios, clientes: clientes, funis: funis, hoje: hojeIsoLocal() };

        var idx = { negocio: {}, cliente: {}, funil: {}, anotacao: {} };
        negocios.forEach(function (n) { idx.negocio[n.id] = n; });
        clientes.forEach(function (c) { idx.cliente[c.id] = c; });
        funis.forEach(function (f) { idx.funil[f.id] = f; });
        anotacoes.forEach(function (a) { idx.anotacao[a.id] = a; });

        var filtradas = CrmCalculos.filtrarAnotacoes(anotacoes, _crmListaFiltros, extra);
        var ordenadas = CrmCalculos.ordenarAnotacoes(filtradas, _crmListaFiltros.ordenarPor, _crmListaFiltros.ordenarDireccao, extra);
        return { anotacoes: anotacoes, extra: extra, idx: idx, ordenadas: ordenadas };
    }

    function nomeFunilDaDemanda(a, idx) {
        var vinculo = CrmCalculos.vinculoDaAnotacao(a, idx.negocio);
        var f = vinculo.funilId ? idx.funil[vinculo.funilId] : null;
        return f ? f.nome : '-';
    }

    function nomeClienteDaDemanda(a, idx) {
        var vinculo = CrmCalculos.vinculoDaAnotacao(a, idx.negocio);
        var c = vinculo.clienteId ? idx.cliente[vinculo.clienteId] : null;
        return c ? c.nome : '-';
    }

    /**
     * Nome do cliente como atalho para o cadastro. O `data-crm-action` próprio
     * tem precedência sobre o da linha (aoClicar usa closest), então clicar aqui
     * abre o cliente em vez da demanda.
     */
    function linkClienteDaDemanda(a, idx) {
        var vinculo = CrmCalculos.vinculoDaAnotacao(a, idx.negocio);
        var c = vinculo.clienteId ? idx.cliente[vinculo.clienteId] : null;
        if (!c) return '-';
        return '<button type="button" class="crm-link-cliente" data-crm-action="editarCliente" ' +
            'data-id="' + esc(c.id) + '" data-origem="demandas" title="Abrir cadastro do cliente">' +
            esc(c.nome) + '</button>';
    }

    function badgeSituacao(a) {
        return '<span class="crm-situacao-badge" data-situacao="' + esc(a.situacao) + '">' +
            esc(SITUACAO_ROTULO[a.situacao] || a.situacao) + '</span>';
    }

    function celulaPrazo(a, hoje) {
        if (!a.prazo) return '<td>-</td>';
        var sem = CrmCalculos.semaforoPrazo(a, hoje);
        var dias = CrmCalculos.diasAte(a.prazo, hoje);
        var sufixo = '';
        if (!a.finalizado) {
            if (dias < 0) sufixo = ' (' + Math.abs(dias) + 'd atrás)';
            else if (dias === 0) sufixo = ' (hoje)';
            else sufixo = ' (em ' + dias + 'd)';
        }
        return '<td class="crm-prazo-cel" data-semaforo="' + esc(sem) + '">' +
            esc(dataOuTraco(a.prazo)) + '<span class="crm-prazo-sufixo">' + esc(sufixo) + '</span></td>';
    }

    function renderizarDemandasView() {
        popularFiltrosLista();

        var ctx = contextoDemandas();
        var hoje = ctx.extra.hoje;
        var grupos = CrmCalculos.agruparPorSituacao(ctx.ordenadas);
        var doGrupo = {
            comigo: grupos.comigo, aguardando: grupos.aguardando,
            consolidar: grupos.consolidar, concluidas: grupos.concluidas, todas: ctx.ordenadas
        };

        // Alertas: prazo estourando (consumo real de lembrarDiasAntes) e terceiro atrasado.
        var emAlerta = 0, comTerceiroAtrasado = 0;
        ctx.ordenadas.forEach(function (a) {
            if (CrmCalculos.semaforoPrazo(a, hoje) !== 'ok') emAlerta++;
            if (CrmCalculos.resumoEncaminhamentos(a, hoje).algumVencido) comTerceiroAtrasado++;
        });

        var abas = DEMANDA_ABAS.map(function (par) {
            var chave = par[0];
            var lista = doGrupo[chave] || [];
            var atrasadas = lista.filter(function (a) {
                return CrmCalculos.semaforoPrazo(a, hoje) === 'vencida' ||
                    CrmCalculos.resumoEncaminhamentos(a, hoje).algumVencido;
            }).length;
            return '<button type="button" class="crm-demanda-aba' + (_demandaAba === chave ? ' active' : '') + '" ' +
                'data-crm-action="setDemandaAba" data-valor="' + esc(chave) + '">' +
                esc(par[1]) + ' <span class="crm-demanda-contador">' + lista.length + '</span>' +
                (atrasadas ? '<span class="crm-demanda-alerta" title="' + atrasadas + ' em atraso">!' + atrasadas + '</span>' : '') +
                '</button>';
        }).join('');

        var visiveis = doGrupo[_demandaAba] || [];

        var linhas = visiveis.map(function (a) {
            var resumo = CrmCalculos.resumoEncaminhamentos(a, hoje);
            var comQuem = resumo.comQuem.length
                ? resumo.comQuem.map(function (p) { return '<span class="crm-tag">' + esc(p) + '</span>'; }).join(' ')
                : (resumo.respondidos ? '<span class="crm-tag crm-tag-ok">' + resumo.respondidos + ' respondido(s)</span>' : '-');
            var espera = resumo.diasAguardando === null ? '-' :
                (resumo.diasAguardando + 'd' + (resumo.algumVencido ? ' ⚠' : ''));

            return '<tr data-crm-action="abrirModalAnotacao" data-id="' + esc(a.id) + '">' +
                '<td class="crm-cel-assunto">' + esc(a.assunto || '(sem assunto)') +
                    '<div class="crm-cel-sub">' + esc(nomeFunilDaDemanda(a, ctx.idx)) +
                    ' · ' + linkClienteDaDemanda(a, ctx.idx) + '</div></td>' +
                '<td>' + esc(a.origemDemanda || '-') + '</td>' +
                '<td>' + esc(a.remetente || '-') + '</td>' +
                celulaPrazo(a, hoje) +
                '<td>' + badgeSituacao(a) + '</td>' +
                '<td>' + comQuem + '</td>' +
                '<td class="crm-cel-espera"' + (resumo.algumVencido ? ' data-atraso="1"' : '') + '>' + esc(espera) + '</td>' +
                '<td><span class="crm-prioridade-badge" data-prioridade="' + esc(a.prioridade) + '">' +
                    esc(PRIORIDADE_ROTULO[a.prioridade] || a.prioridade) + '</span></td>' +
                '<td>' + esc(a.responsavel || '-') + '</td>' +
            '</tr>';
        }).join('');

        var cabecalhos = [
            ['assunto', 'Assunto'], ['origemDemanda', 'Origem'], ['remetente', 'Solicitante'],
            ['prazo', 'Prazo'], ['situacao', 'Situação'], [null, 'Com quem está'],
            ['diasAguardando', 'Aguardando'], ['prioridade', 'Prioridade'], ['responsavel', 'Responsável']
        ].map(function (c) {
            if (!c[0]) return '<th>' + esc(c[1]) + '</th>';
            var seta = _crmListaFiltros.ordenarPor === c[0] ? (_crmListaFiltros.ordenarDireccao === 'asc' ? ' ▲' : ' ▼') : '';
            return '<th onclick="Crm.setOrdenacaoLista(\'' + c[0] + '\')">' + esc(c[1]) + seta + '</th>';
        }).join('');

        var avisos = '';
        if (emAlerta || comTerceiroAtrasado) {
            var partes = [];
            if (emAlerta) partes.push(emAlerta + ' com prazo próximo ou vencido');
            if (comTerceiroAtrasado) partes.push(comTerceiroAtrasado + ' com terceiro atrasado');
            avisos = '<div class="crm-demanda-avisos">⚠ ' + esc(partes.join(' · ')) + '</div>';
        }

        var html = '<div class="crm-demanda-topo">' +
                '<div class="crm-demanda-abas">' + abas + '</div>' +
                '<button type="button" class="crm-btn-negocio" onclick="Crm.abrirModalAnotacao()">+ Nova Demanda</button>' +
            '</div>' +
            avisos +
            '<div class="crm-lista-wrapper"><table class="crm-lista-table crm-demanda-table"><thead><tr>' + cabecalhos + '</tr></thead><tbody>' +
            (visiveis.length ? linhas : '<tr><td colspan="9" style="text-align:center;padding:20px;color:#94a3b8;">Nenhuma demanda nesta caixa.</td></tr>') +
            '</tbody></table></div>';

        var el = document.getElementById('crmDemandas');
        if (el) el.innerHTML = html;
        atualizarContagem(visiveis.length, 'demanda', 'demandas');
    }

    function setDemandaAba(valor) {
        _demandaAba = valor;
        renderizarConteudoAtivo();
    }

    // ── Lista (Anotações) ──

    function renderizarListaView() {
        popularFiltrosLista();

        var ctx = contextoDemandas();
        var hoje = ctx.extra.hoje;
        var ordenadas = ctx.ordenadas;

        var inicio = (_crmListaPagina - 1) * _crmListaPaginaSize;
        var fim = inicio + _crmListaPaginaSize;
        var pagina = ordenadas.slice(inicio, fim);

        var linhas = pagina.map(function (a) {
            var vencida = (CrmCalculos.semaforoPrazo(a, hoje) === 'vencida') ? ' class="crm-lista-vencida"' : '';
            var finalizadoIcone = a.finalizado ? '✓' : '✗';
            var prioridadeBadge = '<span class="crm-prioridade-badge" data-prioridade="' + esc(a.prioridade) + '">' +
                esc(PRIORIDADE_ROTULO[a.prioridade] || a.prioridade) + '</span>';
            var tags = (a.tags || []).map(function (t) { return '<span class="crm-tag">' + esc(t) + '</span>'; }).join(' ');
            var pai = a.anotacaoRelacionadaId ? ctx.idx.anotacao[a.anotacaoRelacionadaId] : null;
            var assuntoCel = (pai ? '<span class="crm-anot-link-icone" title="Vinculada a: ' + esc(pai.assunto || '(sem assunto)') + '">🗒️</span> ' : '') + esc(a.assunto || '-');
            var resumo = CrmCalculos.resumoEncaminhamentos(a, hoje);
            var encaminhadoPara = (a.encaminhamentos || []).map(function (e) { return e.para; }).filter(Boolean).join(', ');

            // Demanda avulsa não tem negócio para trocar de funil: mostra texto simples.
            var celFunil = a.negocioId
                ? '<button type="button" class="crm-lista-btn-funil" data-crm-action="trocarFunilNegocio" data-negocio-id="' + esc(a.negocioId) + '">' + esc(nomeFunilDaDemanda(a, ctx.idx)) + '</button>'
                : esc(nomeFunilDaDemanda(a, ctx.idx));

            return '<tr' + vencida + ' data-crm-action="abrirModalAnotacao" data-id="' + esc(a.id) + '">' +
                '<td>' + celFunil + '</td>' +
                '<td>' + linkClienteDaDemanda(a, ctx.idx) + '</td>' +
                '<td>' + assuntoCel + '</td>' +
                '<td>' + badgeSituacao(a) + '</td>' +
                '<td>' + esc(a.responsavel || '-') + '</td>' +
                '<td>' + esc(a.remetente || '-') + '</td>' +
                '<td>' + esc(a.origemDemanda || '-') + '</td>' +
                '<td>' + esc(dataOuTraco(a.dataSolicitacao)) + '</td>' +
                '<td>' + esc(a.tipoDoc || '-') + '</td>' +
                '<td>' + esc(a.numeroDocumento || '-') + '</td>' +
                '<td>' + esc(encaminhadoPara || '-') + '</td>' +
                '<td>' + (resumo.pendentes ? esc(resumo.pendentes + ' pendente(s)') : (resumo.total ? 'todos responderam' : '-')) + '</td>' +
                '<td>' + esc(a.acaoRealizar || '-') + '</td>' +
                '<td>' + prioridadeBadge + '</td>' +
                '<td>' + esc(dataOuTraco(a.prazo)) + '</td>' +
                '<td>' + esc(a.lembrarDiasAntes === null ? '-' : String(a.lembrarDiasAntes)) + '</td>' +
                '<td>' + (tags || '-') + '</td>' +
                '<td>' + esc(a.observacoes || '-') + '</td>' +
                '<td>' + finalizadoIcone + '</td>' +
                '<td>' + esc(dataOuTraco(a.dataConclusao)) + '</td>' +
                '<td>' + esc(a.oQueFoiFeito || '-') + '</td>' +
                '<td>' + esc(a.respostaNumeroDocumento || '-') + '</td>' +
            '</tr>';
        }).join('');

        var totalPaginas = Math.max(1, Math.ceil(ordenadas.length / _crmListaPaginaSize));
        var colunas = [
            ['funil', 'Relacionamento'], ['cliente', 'Cliente'], ['assunto', 'Assunto'],
            ['situacao', 'Situação'], ['responsavel', 'Responsável'], ['remetente', 'Solicitante'],
            ['origemDemanda', 'Origem da Demanda'], ['dataSolicitacao', 'Data Solicitação'], ['tipoDoc', 'Tipo Doc'],
            ['numeroDocumento', 'Nº Documento'], [null, 'Encaminhado para'], [null, 'Respostas'],
            ['acaoRealizar', 'Ação a Realizar'],
            ['prioridade', 'Prioridade'], ['prazo', 'Prazo'], [null, 'Avisar (dias antes)'], [null, 'Tags'],
            [null, 'Observações'], ['finalizado', 'Finalizado'], ['dataConclusao', 'Data Conclusão'],
            [null, 'O que foi feito'], [null, 'Nº Doc saída']
        ];
        var thead = colunas.map(function (c) {
            if (!c[0]) return '<th>' + esc(c[1]) + '</th>';
            var seta = _crmListaFiltros.ordenarPor === c[0] ? (_crmListaFiltros.ordenarDireccao === 'asc' ? ' ▲' : ' ▼') : '';
            return '<th onclick="Crm.setOrdenacaoLista(\'' + c[0] + '\')">' + esc(c[1]) + seta + '</th>';
        }).join('');
        var html = '<div class="crm-lista-topo"><button type="button" class="crm-btn-negocio" onclick="Crm.abrirModalAnotacao()">+ Nova Demanda</button></div>' +
            '<div class="crm-lista-wrapper"><table class="crm-lista-table"><thead><tr>' +
            thead +
            '</tr></thead><tbody>' +
            (pagina.length ? linhas : '<tr><td colspan="' + colunas.length + '" style="text-align:center;padding:20px;color:#94a3b8;">Nenhuma demanda encontrada.</td></tr>') +
            '</tbody></table></div>' +
            '<div class="crm-paginacao">' +
            '<button type="button" onclick="Crm.setListaPagina(' + Math.max(1, _crmListaPagina - 1) + ')" ' + (_crmListaPagina === 1 ? 'disabled' : '') + '>← Anterior</button>' +
            '<span>Página ' + _crmListaPagina + ' de ' + totalPaginas + ' (' + ordenadas.length + ' resultados)</span>' +
            '<button type="button" onclick="Crm.setListaPagina(' + Math.min(totalPaginas, _crmListaPagina + 1) + ')" ' + (_crmListaPagina === totalPaginas ? 'disabled' : '') + '>Próxima →</button>' +
            '</div>';

        var listEl = document.getElementById('crmListaNegocios');
        if (listEl) listEl.innerHTML = html;
        atualizarContagem(ordenadas.length, 'demanda', 'demandas');
    }

    function dataOuTraco(iso) {
        if (!iso) return '-';
        return (DateUtils.formatBR ? DateUtils.formatBR(iso) : iso) || '-';
    }

    function hojeIsoLocal() {
        var d = new Date();
        return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
    }

    function popularSelectPreservandoValor(sel, opcoes, valorAtual) {
        if (!sel) return;
        sel.innerHTML = opcoes.map(function (o) {
            return '<option value="' + esc(o.valor) + '">' + esc(o.rotulo) + '</option>';
        }).join('');
        sel.value = valorAtual || '';
    }

    function popularFiltrosLista() {
        var funis = CrmStore.listarFunis();
        var clientes = CrmStore.listarClientes();
        var anotacoes = CrmStore.listarAnotacoes();

        // Origem, responsável e destinatários dos encaminhamentos são texto livre:
        // as opções saem dos valores efetivamente usados.
        var origens = [], responsaveis = [], destinos = [];
        anotacoes.forEach(function (a) {
            if (a.origemDemanda && origens.indexOf(a.origemDemanda) === -1) origens.push(a.origemDemanda);
            if (a.responsavel && responsaveis.indexOf(a.responsavel) === -1) responsaveis.push(a.responsavel);
            (a.encaminhamentos || []).forEach(function (e) {
                if (e.para && destinos.indexOf(e.para) === -1) destinos.push(e.para);
            });
        });
        origens.sort();
        responsaveis.sort();
        destinos.sort();

        popularSelectPreservandoValor(
            document.getElementById('crmListaFunil'),
            [{ valor: '', rotulo: 'Todos os funis' }].concat(funis.map(function (f) { return { valor: f.id, rotulo: f.nome }; })),
            _crmListaFiltros.funilId
        );
        popularSelectPreservandoValor(
            document.getElementById('crmListaCliente'),
            [{ valor: '', rotulo: 'Cliente' }].concat(clientes.map(function (c) { return { valor: c.id, rotulo: c.nome }; })),
            _crmListaFiltros.clienteId
        );
        popularSelectPreservandoValor(
            document.getElementById('crmListaOrigem'),
            [{ valor: '', rotulo: 'Origem da demanda' }].concat(origens.map(function (o) { return { valor: o, rotulo: o }; })),
            _crmListaFiltros.origemDemanda
        );

        popularSelectPreservandoValor(
            document.getElementById('crmListaResponsavel'),
            [{ valor: '', rotulo: 'Responsável' }].concat(responsaveis.map(function (r) { return { valor: r, rotulo: r }; })),
            _crmListaFiltros.responsavel
        );
        popularSelectPreservandoValor(
            document.getElementById('crmListaEncaminhadoPara'),
            [{ valor: '', rotulo: 'Encaminhado para' }].concat(destinos.map(function (d) { return { valor: d, rotulo: d }; })),
            _crmListaFiltros.encaminhadoPara
        );

        var inpBusca = document.getElementById('crmListaBusca');
        if (inpBusca && document.activeElement !== inpBusca) inpBusca.value = _crmListaFiltros.busca;
        var prioridade = document.getElementById('crmListaPrioridade');
        if (prioridade) prioridade.value = _crmListaFiltros.prioridade || '';
        var situacao = document.getElementById('crmListaSituacao');
        if (situacao) situacao.value = _crmListaFiltros.situacao || '';
        var soVencidas = document.getElementById('crmListaSomenteVencidas');
        if (soVencidas) soVencidas.checked = !!_crmListaFiltros.somenteVencidas;
        var soAtraso = document.getElementById('crmListaSomenteAtrasoTerceiro');
        if (soAtraso) soAtraso.checked = !!_crmListaFiltros.somenteAtrasoTerceiro;
        var mostrarFinalizadas = document.getElementById('crmListaMostrarConcluidas');
        if (mostrarFinalizadas) mostrarFinalizadas.checked = _crmListaFiltros.mostrarFinalizadas !== false;
        var prazoInicio = document.getElementById('crmListaDataPrazoInicio');
        if (prazoInicio) prazoInicio.value = _crmListaFiltros.prazoInicio || '';
        var prazoFim = document.getElementById('crmListaDataPrazoFim');
        if (prazoFim) prazoFim.value = _crmListaFiltros.prazoFim || '';
        var solInicio = document.getElementById('crmListaDataRecInicio');
        if (solInicio) solInicio.value = _crmListaFiltros.dataSolicitacaoInicio || '';
        var solFim = document.getElementById('crmListaDataRecFim');
        if (solFim) solFim.value = _crmListaFiltros.dataSolicitacaoFim || '';
    }

    function setListaBusca(valor) {
        _crmListaFiltros.busca = valor;
        _crmListaPagina = 1;
        renderizarConteudoAtivo();
    }

    function setListaFiltro() {
        var el = function (id) { return document.getElementById(id); };
        _crmListaFiltros.funilId = el('crmListaFunil') ? (el('crmListaFunil').value || null) : null;
        _crmListaFiltros.clienteId = el('crmListaCliente') ? (el('crmListaCliente').value || null) : null;
        _crmListaFiltros.origemDemanda = el('crmListaOrigem') ? (el('crmListaOrigem').value || null) : null;
        _crmListaFiltros.prioridade = el('crmListaPrioridade') ? (el('crmListaPrioridade').value || null) : null;
        _crmListaFiltros.mostrarFinalizadas = el('crmListaMostrarConcluidas') ? el('crmListaMostrarConcluidas').checked : true;
        _crmListaFiltros.prazoInicio = el('crmListaDataPrazoInicio') ? (el('crmListaDataPrazoInicio').value || null) : null;
        _crmListaFiltros.prazoFim = el('crmListaDataPrazoFim') ? (el('crmListaDataPrazoFim').value || null) : null;
        _crmListaFiltros.dataSolicitacaoInicio = el('crmListaDataRecInicio') ? (el('crmListaDataRecInicio').value || null) : null;
        _crmListaFiltros.dataSolicitacaoFim = el('crmListaDataRecFim') ? (el('crmListaDataRecFim').value || null) : null;
        _crmListaFiltros.situacao = el('crmListaSituacao') ? (el('crmListaSituacao').value || null) : null;
        _crmListaFiltros.responsavel = el('crmListaResponsavel') ? (el('crmListaResponsavel').value || null) : null;
        _crmListaFiltros.encaminhadoPara = el('crmListaEncaminhadoPara') ? (el('crmListaEncaminhadoPara').value || null) : null;
        _crmListaFiltros.somenteVencidas = el('crmListaSomenteVencidas') ? el('crmListaSomenteVencidas').checked : false;
        _crmListaFiltros.somenteAtrasoTerceiro = el('crmListaSomenteAtrasoTerceiro') ? el('crmListaSomenteAtrasoTerceiro').checked : false;
        _crmListaPagina = 1;
        renderizarConteudoAtivo();
    }

    function limparListaFiltros() {
        _crmListaFiltros = {
            busca: '', funilId: null, clienteId: null, prioridade: null, origemDemanda: null,
            prazoInicio: null, prazoFim: null, dataSolicitacaoInicio: null, dataSolicitacaoFim: null,
            mostrarFinalizadas: true, situacao: null, responsavel: null, encaminhadoPara: null,
            somenteVencidas: false, somenteAtrasoTerceiro: false,
            ordenarPor: 'prazo', ordenarDireccao: 'asc'
        };
        _crmListaPagina = 1;
        var inpBusca = document.getElementById('crmListaBusca');
        if (inpBusca) inpBusca.value = '';
        renderizarConteudoAtivo();
    }

    function setOrdenacaoLista(coluna) {
        if (_crmListaFiltros.ordenarPor === coluna) {
            _crmListaFiltros.ordenarDireccao = _crmListaFiltros.ordenarDireccao === 'asc' ? 'desc' : 'asc';
        } else {
            _crmListaFiltros.ordenarPor = coluna;
            _crmListaFiltros.ordenarDireccao = 'asc';
        }
        _crmListaPagina = 1;
        renderizarConteudoAtivo();
    }

    function setListaPagina(p) {
        _crmListaPagina = Math.max(1, p);
        renderizarConteudoAtivo();
    }

    function abrirSeletorFunilNegocio(negocioId) {
        var negocio = CrmStore.listarNegocios().filter(function (n) { return n.id === negocioId; })[0];
        if (!negocio) return;

        var funis = CrmStore.listarFunis();
        var opcoes = funis.map(function (f) {
            return '<button type="button" class="crm-funil-opcao' + (f.id === negocio.funilId ? ' ativo' : '') + '" data-funil-id="' + esc(f.id) + '" onclick="Crm.salvarTrocaFunilNegocio(\'' + esc(negocioId) + '\', this.dataset.funilId)">' + esc(f.nome) + '</button>';
        }).join('');

        var html = '<div class="crm-seletor-funil">' +
            '<div class="crm-seletor-funil-titulo">Trocar relacionamento para:</div>' +
            '<div class="crm-seletor-funil-opcoes">' + opcoes + '</div>' +
            '</div>';

        var container = document.getElementById('crmSeletorFunil');
        if (!container) {
            container = document.createElement('div');
            container.id = 'crmSeletorFunil';
            document.body.appendChild(container);
        }
        container.innerHTML = html;
        container.style.display = 'block';

        _negocioTrocandoFunil = negocioId;

        // Fecha ao clicar fora
        var fecharSeletor = function (e) {
            if (!container.contains(e.target) && e.target.getAttribute('data-crm-action') !== 'trocarFunilNegocio') {
                container.style.display = 'none';
                document.removeEventListener('click', fecharSeletor);
            }
        };
        setTimeout(function () { document.addEventListener('click', fecharSeletor); }, 0);
    }

    function salvarTrocaFunilNegocio(negocioId, novoFunilId) {
        if (!requireAdminOrNotify()) return;
        var novoFunil = CrmStore.listarFunis().filter(function (f) { return f.id === novoFunilId; })[0];
        var novaEtapa = novoFunil ? (novoFunil.etapas.filter(function (e) { return e.tipo === 'aberta'; })[0] || novoFunil.etapas[0]) : null;
        CrmStore.atualizarNegocio(negocioId, { funilId: novoFunilId, etapaId: novaEtapa ? novaEtapa.id : null });
        var container = document.getElementById('crmSeletorFunil');
        if (container) container.style.display = 'none';
        renderizarConteudoAtivo();
    }

    // ── Calendário (Atividades de todos os negócios) ──

    var CAL_PERIODOS = [
        ['todos', 'Tudo'], ['paraFazer', 'Para fazer'], ['vencido', 'Vencido'],
        ['hoje', 'Hoje'], ['amanha', 'Amanhã'], ['semana', 'Esta semana'], ['proximaSemana', 'Próxima semana']
    ];
    var CAL_GRADE_INICIO_MIN = 6 * 60;   // 06:00
    var CAL_GRADE_FIM_MIN = 21 * 60;     // 21:00
    var DIAS_SEMANA_ABREV = ['dom', 'seg', 'ter', 'qua', 'qui', 'sex', 'sáb'];
    var MESES_ABREV = ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez'];

    function ehFimDeSemana(dataIso) {
        var dia = new Date(dataIso + 'T00:00:00Z').getUTCDay();
        return dia === 0 || dia === 6;
    }

    // ── Ponte com o Google Agenda (GoogleCalendarSync, google-calendar-sync.js) ──

    function contextoNegocioCliente(negocioId) {
        var negocio = negocioId ? CrmStore.listarNegocios().filter(function (n) { return n.id === negocioId; })[0] : null;
        var cliente = negocio ? CrmStore.getCliente(negocio.clienteId) : null;
        return { negocio: negocio, cliente: cliente };
    }

    function sincronizarAtividadeComGoogle(id) {
        if (!id || !window.GoogleCalendarSync || !GoogleCalendarSync.conectado()) return;
        var atividade = CrmStore.listarAtividades().filter(function (a) { return a.id === id; })[0];
        if (!atividade) return;
        var ctx = contextoNegocioCliente(atividade.negocioId);
        GoogleCalendarSync.sincronizarAtividade(atividade, ctx.negocio, ctx.cliente).catch(function (e) {
            Notifications.error('Falha ao sincronizar com o Google Agenda: ' + (e && e.message ? e.message : 'erro desconhecido'));
        });
    }

    function removerAtividadeDoGoogle(atividade) {
        if (!atividade || !window.GoogleCalendarSync || !GoogleCalendarSync.conectado()) return;
        GoogleCalendarSync.removerEvento(atividade).catch(function () {});
    }

    function atividadesCalContexto() {
        var negocios = CrmStore.listarNegocios();
        var clientes = CrmStore.listarClientes();
        var negocioPorId = {};
        negocios.forEach(function (n) { negocioPorId[n.id] = n; });
        var clientePorId = {};
        clientes.forEach(function (c) { clientePorId[c.id] = c; });
        return { negocios: negocios, clientes: clientes, negocioPorId: negocioPorId, clientePorId: clientePorId };
    }

    /**
     * Feriados nacionais, calculados localmente (sem depender de conta do
     * Google nem de conexão) — sempre presentes no Calendário, como itens
     * somente leitura (nunca viram atividade de verdade no CRM).
     */
    /**
     * Migração: em versões anteriores, feriados importados do Google chegaram
     * a ser salvos como atividades reais (origemFeriado:true) no CRM. Feriados
     * agora são só calculados na hora (nunca persistidos) — remove qualquer
     * resquício real, senão o filtro "Mostrar feriados" não tem efeito sobre
     * eles. Roda a cada render (os dados do Firestore chegam de forma
     * assíncrona, então uma trava de "só uma vez" arriscaria rodar antes dos
     * dados reais carregarem e nunca mais tentar de novo); uma vez limpos, o
     * filtro em CrmStore.listarAtividades() já não acha nada, então o custo
     * cai a praticamente zero nos renders seguintes.
     */
    function limparFeriadosOrfaos() {
        var orfaos = CrmStore.listarAtividades().filter(function (a) { return a.origemFeriado; });
        if (!orfaos.length) return;
        orfaos.forEach(function (a) { CrmStore.removerAtividade(a.id); });
    }

    function gerarFeriadosPseudoAtividades() {
        var anoAtual = Number(hojeIsoLocal().slice(0, 4));
        var lista = [];
        for (var ano = anoAtual - 1; ano <= anoAtual + 2; ano++) {
            CrmCalculos.feriadosNacionais(ano).forEach(function (f) {
                lista.push({
                    id: 'feriado_' + f.data,
                    negocioId: null,
                    tipo: 'prazo',
                    assunto: f.nome,
                    descricao: '',
                    data: f.data,
                    horaInicio: '',
                    horaFim: '',
                    feito: false,
                    origemFeriado: true
                });
            });
        }
        return lista;
    }

    // ── Ponte com o sistema Ponto (PontoBridge, ponto-calendar-bridge.js) ──
    // Eventos e Férias do Ponto viram pseudo-atividades somente leitura, no
    // mesmo espírito dos feriados nacionais (gerarFeriadosPseudoAtividades):
    // geradas na hora a partir do cache do PontoBridge, nunca gravadas aqui.

    var PONTO_TIPO_META = {
        feriado: { icone: '🏖️', cor: '#dc2626', corTexto: '#ffffff', rotulo: 'Feriado' },
        ferias: { icone: '🏝️', cor: '#d97706', corTexto: '#ffffff', rotulo: 'Férias' },
        viagem: { icone: '✈️', cor: '#7c3aed', corTexto: '#ffffff', rotulo: 'Viagem' },
        abono: { icone: '📋', cor: '#10b981', corTexto: '#ffffff', rotulo: 'Abono' },
        abono_acordo: { icone: '🤝', cor: '#059669', corTexto: '#ffffff', rotulo: 'Abono (acordo)' },
        pagar_hora_acordo: { icone: '⏱️', cor: '#db2777', corTexto: '#ffffff', rotulo: 'Pagar hora (acordo)' },
        outro: { icone: '📌', cor: '#64748b', corTexto: '#ffffff', rotulo: 'Outro' }
    };
    var PONTO_MAX_DIAS_POR_EVENTO = 400; // trava de segurança contra datas absurdas
    var PONTO_LIMIAR_INTERVALO_SUSPEITO = 60; // acima disso não é plausível p/ Viagem/Férias/Abono etc — provável dado ruim

    /**
     * Normaliza a data de um evento do Ponto para YYYY-MM-DD. Aceita ISO direto
     * (com ou sem hora) e DD/MM/YYYY. Registros de Férias mais antigos usam
     * nomes de campo variados (`inicio`/`dataInicio` em vez de `dataInicioEvento`)
     * — o próprio Ponto trata essa variação em `computeVacationOverview`, com o
     * mesmo fallback de três nomes reproduzido em gerarPontoPseudoAtividades.
     */
    function normalizarDataPonto(v) {
        if (!v) return null;
        var s = String(v).trim();
        if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
        var m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
        if (m) return m[3] + '-' + m[2] + '-' + m[1];
        return null;
    }

    /** Quantidade de dias entre duas datas ISO, inclusive (sem gerar a lista). */
    function contarDiasIso(iniIso, fimIso) {
        var d = new Date(iniIso + 'T00:00:00Z');
        var f = new Date((fimIso || iniIso) + 'T00:00:00Z');
        if (isNaN(d) || isNaN(f)) return 1;
        if (f < d) f = new Date(d.getTime());
        return Math.round((f - d) / 86400000) + 1;
    }

    /**
     * Lista de datas ISO de `inicio` a `fim`, inclusive, capada em `max`.
     * IMPORTANTE: ao "grudar" fimD em d (quando fim < início, ex: erro de
     * digitação de ano no Ponto), usar `new Date(...)` — nunca `fimD = d`.
     * Isso faria os dois nomes apontarem pro MESMO objeto; o `d.setUTCDate`
     * do loop empurraria fimD junto, e `d <= fimD` nunca pararia até o `max`.
     */
    function expandirIntervaloDias(inicio, fim, max) {
        var out = [];
        var d = new Date(inicio + 'T00:00:00Z');
        var fimD = new Date((fim || inicio) + 'T00:00:00Z');
        if (isNaN(d) || isNaN(fimD)) return [inicio];
        if (fimD < d) fimD = new Date(d.getTime());
        var i = 0;
        while (d <= fimD && i < max) {
            out.push(d.toISOString().slice(0, 10));
            d.setUTCDate(d.getUTCDate() + 1);
            i++;
        }
        return out.length ? out : [inicio];
    }

    /**
     * Atividades reais que cobrem vários dias (hoje só acontece com reservas
     * de hotel/viagem importadas do Google, ver dataFim em crm-model.js) viram
     * um clone por dia coberto, todos com o mesmo `grupoBarraId` (o id da
     * atividade original) — é o que permite às grades Mês/Semana mesclar a
     * célula numa única barra contínua, no mesmo mecanismo já usado pros
     * eventos multi-dia do Ponto (ver acharBarrasPontoNaSemana). Atividades de
     * 1 dia só passam direto, sem clonar.
     */
    function expandirAtividadesMultidia(atividades) {
        var out = [];
        atividades.forEach(function (a) {
            if (!a.dataFim || a.dataFim <= a.data) { out.push(a); return; }
            var dias = expandirIntervaloDias(a.data, a.dataFim, PONTO_MAX_DIAS_POR_EVENTO);
            dias.forEach(function (dataIso) {
                var clone = {};
                for (var k in a) { if (Object.prototype.hasOwnProperty.call(a, k)) clone[k] = a[k]; }
                clone.data = dataIso;
                clone.grupoBarraId = a.id;
                out.push(clone);
            });
        });
        return out;
    }

    /**
     * `AppState.dados.eventos` no Ponto mistura dois tipos de registro:
     *  - eventos de verdade, criados na aba Eventos (Viagem, Férias, Feriado,
     *    Abono...);
     *  - pseudo-eventos AUTO-GERADOS a cada Registro de ponto salvo com período
     *    (`Registro: matutino/vespertino`), marcados com `linkedRegistroDate`
     *    (ou `nomeCSS: 'evento-registro'`) — ver app-refatorado.js, função que
     *    salva o registro do dia (~linha 4403).
     * Aqui só interessa a aba Eventos; os pseudo-eventos de registro são ruído
     * no calendário do CRM (nunca são "prazo"/"atividade", são reflexo do
     * ponto batido no dia).
     */
    function ehEventoDeRegistroPonto(ev) {
        return !!(ev.linkedRegistroDate || ev.nomeCSS === 'evento-registro');
    }

    function gerarPontoPseudoAtividades() {
        if (!window.PontoBridge || !PontoBridge.conectado()) return [];
        var eventos = PontoBridge.getEventosCache();
        var lista = [];
        eventos.forEach(function (ev, idx) {
            if (ehEventoDeRegistroPonto(ev)) return;
            var ini = normalizarDataPonto(ev.dataInicioEvento || ev.inicio || ev.dataInicio);
            var fim = normalizarDataPonto(ev.dataFimEvento || ev.fim || ev.dataFim) || ini;
            if (!ini) return;
            var meta = PONTO_TIPO_META[ev.tipoEvento] || PONTO_TIPO_META.outro;
            var totalDias = contarDiasIso(ini, fim);

            // Intervalo implausível para o tipo do evento (ex: "Feriado" de 400 dias) —
            // quase sempre é dado ruim no Ponto (fim ausente/errado), não uma escala real.
            // Em vez de inundar o mês inteiro com um card por dia, mostra 1 card de aviso.
            if (totalDias > PONTO_LIMIAR_INTERVALO_SUSPEITO) {
                console.warn('Ponto: evento com intervalo suspeito (' + totalDias + ' dias, ' + ini + ' a ' + fim + '). Mostrando 1 card em vez de expandir. Evento bruto:', ev);
                lista.push({
                    id: 'ponto_' + idx + '_suspeito',
                    negocioId: null,
                    tipo: 'prazo',
                    assunto: '⚠ ' + (ev.descricaoEvento || meta.rotulo) + ' (' + totalDias + ' dias — confira no Ponto)',
                    descricao: 'Intervalo de ' + ini + ' a ' + fim + ' parece incorreto para este tipo de evento.',
                    data: ini,
                    horaInicio: '',
                    horaFim: '',
                    feito: false,
                    origemPonto: true,
                    pontoGrupoId: 'ponto_' + idx,
                    pontoTipoRotulo: meta.rotulo,
                    pontoCor: ev.corFundo || meta.cor,
                    pontoCorTexto: ev.corTexto || meta.corTexto,
                    pontoIcone: meta.icone
                });
                return;
            }

            // `pontoGrupoId` é compartilhado por todos os dias do mesmo evento —
            // é o que permite à visão Mês mesclar os dias num único bloco
            // contínuo (renderizarBarrasPontoSemana), em vez de um card por dia.
            var dias = expandirIntervaloDias(ini, fim, PONTO_MAX_DIAS_POR_EVENTO);
            dias.forEach(function (dataIso, i) {
                lista.push({
                    id: 'ponto_' + idx + '_' + dataIso,
                    negocioId: null,
                    tipo: 'prazo',
                    assunto: ev.descricaoEvento || meta.rotulo,
                    descricao: '',
                    data: dataIso,
                    horaInicio: '',
                    horaFim: '',
                    feito: false,
                    origemPonto: true,
                    pontoGrupoId: 'ponto_' + idx,
                    pontoTipoRotulo: meta.rotulo,
                    pontoCor: ev.corFundo || meta.cor,
                    pontoCorTexto: ev.corTexto || meta.corTexto,
                    pontoIcone: meta.icone
                });
            });
        });
        return lista;
    }

    /** No máximo 1 busca automática por carregamento de página; o resto é sob demanda (conectar/atualizar). */
    function garantirEventosPontoCarregados() {
        if (!window.PontoBridge || !PontoBridge.conectado() || _pontoAutoFetchFeito) return;
        _pontoAutoFetchFeito = true;
        PontoBridge.atualizarEventos().then(function () {
            if (_secao === 'calendario') renderizarCalendarioView();
        });
    }

    function renderizarPontoStatus() {
        var pb = window.PontoBridge;
        if (!pb || !pb.configurado()) return '';
        if (pb.conectado()) {
            var qtd = pb.getEventosCache().length;
            var erro = pb.getUltimoErro && pb.getUltimoErro();
            return '<div class="crm-cal-ponto">' +
                '<span class="crm-cal-ponto-on">🟢 Ponto conectado (' + esc(pb.usuarioEmail()) + ') — ' + qtd + ' evento(s) carregado(s)</span>' +
                '<label class="crm-check"><input type="checkbox" ' + (_calMostrarPonto ? 'checked' : '') +
                    ' onchange="Crm.setMostrarPontoCal(this.checked)"> Mostrar no calendário</label>' +
                '<button type="button" class="btn-secondary crm-btn-mini" data-crm-action="pontoAtualizar">Atualizar</button>' +
                '<button type="button" class="btn-secondary crm-btn-mini" data-crm-action="pontoDesconectar">Desconectar</button>' +
                (erro ? ('<span class="crm-cal-ponto-erro" title="' + esc(erro) + '">⚠ ' + esc(erro) + '</span>') : '') +
            '</div>';
        }
        return '<div class="crm-cal-ponto">' +
            '<span class="crm-cal-ponto-off">📅 Ponto: desconectado</span>' +
            '<input type="email" id="crmPontoEmail" class="crm-filtro-input crm-cal-ponto-input" placeholder="e-mail do Ponto" autocomplete="username">' +
            '<input type="password" id="crmPontoSenha" class="crm-filtro-input crm-cal-ponto-input" placeholder="senha" autocomplete="current-password" ' +
                'onkeydown="if(event.key===\'Enter\'){event.preventDefault();Crm.pontoConectar();}">' +
            '<button type="button" class="btn-secondary crm-btn-mini" data-crm-action="pontoConectar">Conectar Ponto</button>' +
        '</div>';
    }

    function setMostrarPontoCal(v) { _calMostrarPonto = v; renderizarCalendarioView(); }

    /**
     * O Ponto já traz feriados nacionais, distritais e da empresa (mais
     * completo que o cálculo genérico do CRM). Enquanto ele estiver conectado
     * e visível, os dois catálogos duplicariam feriados nacionais (ex:
     * "Corpus Christi" aparecendo duas vezes) — então o gerador genérico do
     * CRM só entra em ação como fallback, quando o Ponto não está disponível.
     */
    function feriadosGenericosAtivos() {
        return _calMostrarFeriados && !(_calMostrarPonto && window.PontoBridge && PontoBridge.conectado());
    }

    function renderizarCalendarioView() {
        var ctx = atividadesCalContexto();
        limparFeriadosOrfaos(); // migração: remove feriados que ficaram salvos como atividade real (fase antiga, via Google)
        garantirEventosPontoCarregados();
        var reais = CrmStore.listarAtividades();
        var feriados = feriadosGenericosAtivos() ? gerarFeriadosPseudoAtividades() : [];
        var doPonto = _calMostrarPonto ? gerarPontoPseudoAtividades() : [];
        var todas = reais.concat(feriados).concat(doPonto);
        var porPeriodo = CrmCalculos.filtrarAtividadesPeriodo(todas, _calPeriodo, hojeIsoLocal());
        var porBusca = CrmCalculos.filtrarAtividadesBusca(porPeriodo, _calBusca, ctx);
        var filtradas = _calTipo ? porBusca.filter(function (a) { return a.tipo === _calTipo; }) : porBusca;

        var chips = CAL_PERIODOS.map(function (p) {
            return '<button type="button" class="crm-cal-chip' + (p[0] === _calPeriodo ? ' active' : '') + '" data-crm-action="setCalPeriodo" data-valor="' + p[0] + '">' + esc(p[1]) + '</button>';
        }).join('');

        var chipsTipo = '<button type="button" class="crm-cal-chip' + (!_calTipo ? ' active' : '') + '" data-crm-action="setCalTipo" data-valor="">Tudo</button>' +
            CrmStore.listarTiposAtividadeAtivos().map(function (t) {
                return '<button type="button" class="crm-cal-chip' + (_calTipo === t.chave ? ' active' : '') + '" data-crm-action="setCalTipo" data-valor="' + esc(t.chave) + '">' + t.icone + ' ' + esc(t.nome) + '</button>';
            }).join('') +
            '<button type="button" class="crm-cal-chip crm-cal-chip-config" data-crm-action="abrirModalTiposAtividade" title="Gerenciar tipos de atividade">⚙</button>';

        var toolbar = '' +
            '<div class="crm-cal-toolbar">' +
                '<button type="button" class="btn btn-primary crm-cal-btn-nova" data-crm-action="abrirModalAtividadeCal">+ Atividade</button>' +
                '<div class="crm-cal-modos">' +
                    '<button type="button" class="crm-cal-modo-btn' + (_calModo === 'mes' ? ' active' : '') + '" data-crm-action="setCalModo" data-valor="mes" title="Mês">📆</button>' +
                    '<button type="button" class="crm-cal-modo-btn' + (_calModo === 'semana' ? ' active' : '') + '" data-crm-action="setCalModo" data-valor="semana" title="Semana">🗓</button>' +
                    '<button type="button" class="crm-cal-modo-btn' + (_calModo === 'lista' ? ' active' : '') + '" data-crm-action="setCalModo" data-valor="lista" title="Lista">☰</button>' +
                '</div>' +
                '<input type="text" class="crm-busca" placeholder="Buscar atividade, negócio ou contato..." value="' + esc(_calBusca) + '" oninput="Crm.setCalBusca(this.value)">' +
                '<span class="crm-contagem">' + filtradas.length + (filtradas.length !== 1 ? ' atividades' : ' atividade') + '</span>' +
            '</div>' +
            renderizarGoogleStatus() +
            renderizarPontoStatus() +
            '<div class="crm-cal-tipos-filtro">' + chipsTipo + '</div>' +
            '<div class="crm-cal-periodos">' + chips + '</div>';

        var corpo = _calModo === 'semana' ? renderizarCalendarioSemana(filtradas)
            : (_calModo === 'mes' ? renderizarCalendarioMes(filtradas) : renderizarCalendarioLista(filtradas, ctx));

        document.getElementById('crmCalendario').innerHTML = toolbar + corpo;
    }

    function renderizarGoogleStatus() {
        var gs = window.GoogleCalendarSync;
        if (!gs || !gs.configurado()) {
            return '<div class="crm-cal-google">' +
                '<span class="crm-cal-google-off">🔗 Google Agenda: não configurado <span title="Defina o Client ID OAuth em google-calendar-sync.js">ⓘ</span></span>' +
            '</div>';
        }
        if (gs.conectado()) {
            return '<div class="crm-cal-google">' +
                '<span class="crm-cal-google-on">🟢 Google Agenda conectado</span>' +
                '<button type="button" class="btn-secondary crm-btn-mini" data-crm-action="googleSincronizarTudo">Sincronizar tudo</button>' +
                '<button type="button" class="btn-secondary crm-btn-mini" data-crm-action="googleDesconectar">Desconectar</button>' +
            '</div>';
        }
        return '<div class="crm-cal-google">' +
            '<span class="crm-cal-google-off">🔗 Google Agenda: desconectado</span>' +
            '<button type="button" class="btn-secondary crm-btn-mini" data-crm-action="googleConectar">Conectar Google Agenda</button>' +
        '</div>';
    }

    function renderizarCalendarioLista(atividades, ctx) {
        var ordenadas = CrmCalculos.ordenarAtividadesPorData(atividades);
        var mapaTipos = CrmStore.mapaTiposAtividade();
        var linhas = ordenadas.map(function (a) {
            var negocio = ctx.negocioPorId[a.negocioId];
            var cliente = negocio ? ctx.clientePorId[negocio.clienteId] : null;
            var tipoInfo = mapaTipos[a.tipo] || { icone: '📌', rotulo: a.tipo };
            var vencida = (!a.feito && a.data && a.data < hojeIsoLocal()) ? ' class="crm-lista-vencida"' : (a.feito ? ' class="crm-cal-linha-feita"' : '');
            var duracao = CrmCalculos.duracaoAtividade(a.horaInicio, a.horaFim);
            var horaTxt = a.horaInicio ? (a.horaInicio + (a.horaFim ? ('–' + a.horaFim) : '')) : '';
            var somenteLeitura = a.origemFeriado || a.origemPonto;
            var selo = a.origemFeriado
                ? ' <span class="crm-cal-badge-feriado" title="Feriado nacional — somente leitura">🎉 Feriado</span>'
                : (a.origemPonto ? ' <span class="crm-cal-badge-ponto" title="Do sistema Ponto (' + esc(a.pontoTipoRotulo || '') + ') — somente leitura">' + esc(a.pontoIcone || '📅') + ' Ponto</span>'
                    : (a.origemGoogle ? ' <span class="crm-cal-badge-google" title="' + (a.somenteLeituraGoogle ? 'Importado da Google Agenda (reserva de voo/hotel) — edições aqui não sincronizam de volta' : 'Importado da Google Agenda') + '">📥 Google' + (a.somenteLeituraGoogle ? ' 🔒' : '') + '</span>' : ''));
            var celAssunto = (a.origemPonto ? '' : (tipoInfo.icone + ' ')) + esc(a.assunto || '(sem assunto)') + selo;

            return '<tr' + vencida + '>' +
                '<td class="crm-cal-td-check">' + (somenteLeitura
                    ? '<input type="checkbox" disabled title="Somente leitura — não editável aqui">'
                    : '<input type="checkbox" ' + (a.feito ? 'checked' : '') + ' data-crm-action="concluirAtividadeCal" data-id="' + esc(a.id) + '" data-feito="' + a.feito + '" onclick="event.stopPropagation()">') + '</td>' +
                (somenteLeitura
                    ? '<td>' + celAssunto + '</td>'
                    : '<td data-crm-action="abrirModalAtividadeCal" data-id="' + esc(a.id) + '">' + celAssunto + '</td>') +
                '<td>' + (negocio ? esc(negocio.titulo || '-') : '-') + '</td>' +
                '<td>' + (cliente ? esc(cliente.nome) : '-') + '</td>' +
                '<td>' + (cliente && cliente.email ? esc(cliente.email) : '-') + '</td>' +
                '<td>' + (cliente && cliente.telefone ? esc(cliente.telefone) : '-') + '</td>' +
                '<td>' + esc(dataOuTraco(a.data)) + (horaTxt ? (' <span class="crm-cal-hora">' + esc(horaTxt) + '</span>') : '') + '</td>' +
                '<td>' + (duracao ? esc(duracao) : '-') + '</td>' +
                '<td>' + (negocio && negocio.responsavel ? esc(negocio.responsavel) : '-') + '</td>' +
            '</tr>';
        }).join('');

        var pontoAtivo = _calMostrarPonto && window.PontoBridge && PontoBridge.conectado();
        var filtroFeriados = '<div class="crm-cal-lista-filtros">' +
            '<label class="crm-check" title="' + (pontoAtivo ? 'Com o Ponto conectado, os feriados vêm de lá (nacionais, distritais e da empresa) — o cálculo genérico do CRM fica em espera para não duplicar.' : 'Cálculo genérico de feriados nacionais.') + '">' +
                '<input type="checkbox" ' + (_calMostrarFeriados ? 'checked' : '') + ' onchange="Crm.setMostrarFeriadosCal(this.checked)">' +
                'Mostrar feriados' + (pontoAtivo ? ' (via Ponto)' : '') +
            '</label>' +
        '</div>';

        return filtroFeriados +
            '<div class="crm-lista-wrapper"><table class="crm-lista-table crm-cal-table"><thead><tr>' +
            '<th></th><th>Assunto</th><th>Negócio</th><th>Pessoa de contato</th><th>E-mail</th><th>Telefone</th>' +
            '<th>Data de vencimento</th><th>Duração</th><th>Atribuído a</th>' +
            '</tr></thead><tbody>' +
            (ordenadas.length ? linhas : '<tr><td colspan="9" style="text-align:center;padding:20px;color:#94a3b8;">Nenhuma atividade encontrada.</td></tr>') +
            '</tbody></table></div>';
    }

    function formatarFaixaSemana(domingoIso) {
        var fim = CrmCalculos.somarDias(domingoIso, 6);
        var di = new Date(domingoIso + 'T00:00:00Z'), df = new Date(fim + 'T00:00:00Z');
        var mesmoMes = di.getUTCMonth() === df.getUTCMonth();
        var fimTxt = (mesmoMes ? '' : (MESES_ABREV[df.getUTCMonth()] + ' ')) + df.getUTCDate();
        return MESES_ABREV[di.getUTCMonth()] + ' ' + di.getUTCDate() + ' – ' + fimTxt + ', ' + df.getUTCFullYear();
    }

    function renderizarCalendarioSemana(atividades) {
        var domingo = _calSemanaRef;
        var porDia = CrmCalculos.agruparAtividadesPorDiaDaSemana(expandirAtividadesMultidia(atividades), domingo);
        var diasSemana = Object.keys(porDia);
        var hoje = hojeIsoLocal();

        var nav = '<div class="crm-cal-nav">' +
            '<button type="button" class="btn-secondary crm-btn-mini" data-crm-action="calHoje">Hoje</button>' +
            '<button type="button" class="crm-cal-seta" data-crm-action="calSemanaAnterior">‹</button>' +
            '<button type="button" class="crm-cal-seta" data-crm-action="calSemanaProxima">›</button>' +
            '<span class="crm-cal-faixa">' + esc(formatarFaixaSemana(domingo)) + '</span>' +
        '</div>';

        var totalMin = CAL_GRADE_FIM_MIN - CAL_GRADE_INICIO_MIN;

        var cabecalho = '<div class="crm-cal-grade-cab"><div class="crm-cal-cab-gutter"></div>' +
            Object.keys(porDia).map(function (dataIso) {
                var d = new Date(dataIso + 'T00:00:00Z');
                var ehHoje = dataIso === hoje;
                var numero = ehHoje ? '<span class="crm-cal-cab-numero-hoje">' + d.getUTCDate() + '</span>' : d.getUTCDate();
                return '<div class="crm-cal-cab-dia' + (ehHoje ? ' crm-cal-cab-hoje' : '') + (ehFimDeSemana(dataIso) ? ' crm-cal-fim-semana' : '') + '">' + DIAS_SEMANA_ABREV[d.getUTCDay()] + ' ' + numero + '</div>';
            }).join('') + '</div>';

        var trilhaHtml = renderizarTrilhaGrade(diasSemana, porDia, 2, 'crm-cal-grade-trilha-linha');

        var diaSemHora = diasSemana.map(function (dataIso) {
            // Itens multi-dia (Ponto ou reais) já viraram barra na trilha acima — não repetir aqui.
            var semHora = porDia[dataIso].filter(function (a) { return !a.horaInicio && !a.origemPonto && !a.grupoBarraId; });
            return '<div class="crm-cal-allday-col' + (ehFimDeSemana(dataIso) ? ' crm-cal-fim-semana' : '') + '" data-crm-action="novaAtividadeNoDia" data-data="' + esc(dataIso) + '" title="Clique para agendar uma atividade neste dia">' + semHora.map(function (a) { return renderizarEventoCard(a); }).join('') + '</div>';
        }).join('');
        var linhaSemHora = trilhaHtml + '<div class="crm-cal-grade-allday"><div class="crm-cal-cab-gutter"></div>' + diaSemHora + '</div>';

        var horas = [];
        for (var m = CAL_GRADE_INICIO_MIN; m < CAL_GRADE_FIM_MIN; m += 60) {
            horas.push('<div class="crm-cal-hora-linha">' + String(Math.floor(m / 60)).padStart(2, '0') + ':00</div>');
        }
        var gutter = '<div class="crm-cal-gutter">' + horas.join('') + '</div>';

        var colunas = Object.keys(porDia).map(function (dataIso) {
            var comHora = porDia[dataIso].filter(function (a) { return a.horaInicio; });
            var eventos = comHora.map(function (a) {
                var inicioMin = horaParaMin(a.horaInicio);
                var fimMin = a.horaFim ? horaParaMin(a.horaFim) : (inicioMin + 30);
                if (fimMin <= inicioMin) fimMin = inicioMin + 30;
                var topPct = Math.max(0, (inicioMin - CAL_GRADE_INICIO_MIN) / totalMin * 100);
                var alturaPct = Math.max(3, (fimMin - inicioMin) / totalMin * 100);
                return renderizarEventoCard(a, 'position:absolute;left:2px;right:2px;top:' + topPct.toFixed(2) + '%;height:' + alturaPct.toFixed(2) + '%');
            }).join('');
            var linhaAgoraCol = (dataIso === hoje) ? renderizarLinhaAgora(totalMin) : '';
            return '<div class="crm-cal-dia-col' + (ehFimDeSemana(dataIso) ? ' crm-cal-fim-semana' : '') + '" data-crm-action="novaAtividadeNoDia" data-data="' + esc(dataIso) + '" title="Clique para agendar uma atividade neste dia">' + eventos + linhaAgoraCol + '</div>';
        }).join('');
        var grade = '<div class="crm-cal-grade-corpo">' + gutter + '<div class="crm-cal-dias-wrap">' + colunas + '</div></div>';

        return nav + '<div class="crm-cal-grade">' + cabecalho + linhaSemHora + grade + '</div>';
    }

    /**
     * Linha do horário atual (marcador vermelho), usada tanto na visão de
     * Semana quanto na mini-agenda do modal de atividade. Sempre aparece
     * quando o dia em questão é hoje — se a hora estiver fora da faixa
     * exibida (06h-21h), fica "grudada" no topo ou no rodapé da grade.
     */
    function renderizarLinhaAgora(totalMin) {
        var agora = new Date();
        var minAgora = agora.getHours() * 60 + agora.getMinutes();
        var topAgora = Math.max(0, Math.min(100, (minAgora - CAL_GRADE_INICIO_MIN) / totalMin * 100));
        var hh = String(agora.getHours()).padStart(2, '0'), mm = String(agora.getMinutes()).padStart(2, '0');
        return '<div class="crm-atv-mini-agora" style="top:' + topAgora.toFixed(2) + '%"><span class="crm-atv-mini-agora-hora">' + hh + ':' + mm + '</span></div>';
    }

    function horaParaMin(hhmm) {
        var partes = String(hhmm || '').split(':');
        var h = Number(partes[0]), m = Number(partes[1]);
        return (isNaN(h) ? 0 : h) * 60 + (isNaN(m) ? 0 : m);
    }

    /**
     * Mini-agenda do dia exibida ao lado do modal de atividade (Calendário),
     * pra mostrar de cara o que já está marcado naquele dia — e, se for hoje,
     * a linha do horário atual. Atualiza sozinha quando a data/hora do form muda.
     */
    function atualizarMiniAgendaCal() {
        var el = document.getElementById('crmAtvCalMiniAgenda');
        if (!el) return;
        var dataIso = document.getElementById('crmAtvCalData').value;
        if (!dataIso) {
            el.innerHTML = '<div class="crm-atv-mini-vazio">Escolha uma data para ver a agenda do dia.</div>';
            return;
        }

        var idAtual = document.getElementById('crmAtvCalId').value;
        var doDia = CrmStore.listarAtividades().filter(function (a) { return a.data === dataIso && a.id !== idAtual; });
        var semHora = doDia.filter(function (a) { return !a.horaInicio; });
        var comHora = doDia.filter(function (a) { return a.horaInicio; });

        var d = new Date(dataIso + 'T00:00:00Z');
        var diaAbrev = DIAS_SEMANA_ABREV[d.getUTCDay()];
        var rotuloDia = diaAbrev.charAt(0).toUpperCase() + diaAbrev.slice(1) + ', ' + Number(dataIso.slice(8, 10)) + ' de ' + MESES[Number(dataIso.slice(5, 7)) - 1];

        var totalMin = CAL_GRADE_FIM_MIN - CAL_GRADE_INICIO_MIN;
        var horas = [];
        for (var m = CAL_GRADE_INICIO_MIN; m < CAL_GRADE_FIM_MIN; m += 60) {
            horas.push('<div class="crm-cal-hora-linha">' + String(Math.floor(m / 60)).padStart(2, '0') + ':00</div>');
        }

        var eventos = comHora.map(function (a) {
            var inicioMin = horaParaMin(a.horaInicio);
            var fimMin = a.horaFim ? horaParaMin(a.horaFim) : (inicioMin + 30);
            if (fimMin <= inicioMin) fimMin = inicioMin + 30;
            var topPct = Math.max(0, (inicioMin - CAL_GRADE_INICIO_MIN) / totalMin * 100);
            var alturaPct = Math.max(3, (fimMin - inicioMin) / totalMin * 100);
            return renderizarEventoCard(a, 'position:absolute;left:2px;right:2px;top:' + topPct.toFixed(2) + '%;height:' + alturaPct.toFixed(2) + '%');
        }).join('');

        // Prévia (fantasma) do horário que está sendo digitado no formulário agora.
        var horaInicioForm = document.getElementById('crmAtvCalHoraInicio').value;
        var previewHtml = '';
        if (horaInicioForm) {
            var inicioMinP = horaParaMin(horaInicioForm);
            var horaFimForm = document.getElementById('crmAtvCalHoraFim').value;
            var fimMinP = horaFimForm ? horaParaMin(horaFimForm) : (inicioMinP + 30);
            if (fimMinP <= inicioMinP) fimMinP = inicioMinP + 30;
            var topP = Math.max(0, (inicioMinP - CAL_GRADE_INICIO_MIN) / totalMin * 100);
            var altP = Math.max(3, (fimMinP - inicioMinP) / totalMin * 100);
            previewHtml = '<div class="crm-atv-mini-preview" style="top:' + topP.toFixed(2) + '%;height:' + altP.toFixed(2) + '%"></div>';
        }

        // Linha do horário atual — só quando a data escolhida é hoje.
        var linhaAgora = (dataIso === hojeIsoLocal()) ? renderizarLinhaAgora(totalMin) : '';

        var semHoraHtml = semHora.map(function (a) { return renderizarEventoCard(a); }).join('');

        el.innerHTML =
            '<div class="crm-atv-mini-titulo">' + esc(rotuloDia) + '</div>' +
            (semHoraHtml ? '<div class="crm-atv-mini-allday">' + semHoraHtml + '</div>' : '') +
            '<div class="crm-atv-mini-grade">' +
                '<div class="crm-atv-mini-gutter">' + horas.join('') + '</div>' +
                '<div class="crm-atv-mini-coluna">' + eventos + previewHtml + linhaAgora + '</div>' +
            '</div>';
    }

    function renderizarEventoCard(a, estiloExtra) {
        var tipoInfo = CrmStore.mapaTiposAtividade()[a.tipo] || { icone: '📌', rotulo: a.tipo };
        var classeOrigem = a.origemFeriado ? ' crm-cal-evento-feriado' : (a.origemPonto ? ' crm-cal-evento-ponto' : (a.origemGoogle ? ' crm-cal-evento-google' : ''));
        var icone = a.origemFeriado ? '🎉' : (a.origemPonto ? (a.pontoIcone || '📅') : tipoInfo.icone);
        var somenteLeitura = a.origemFeriado || a.origemPonto;
        var acao = somenteLeitura ? '' : ('data-crm-action="abrirModalAtividadeCal" data-id="' + esc(a.id) + '"');
        var estiloPonto = a.origemPonto ? ('background:' + esc(a.pontoCor || '#64748b') + ';color:' + esc(a.pontoCorTexto || '#fff') + ';border-left-color:' + esc(a.pontoCor || '#64748b') + ';') : '';
        var estiloFinal = estiloPonto + (estiloExtra || '');
        var titulo = a.origemFeriado ? 'Feriado nacional — somente leitura'
            : (a.origemPonto ? ('Do sistema Ponto (' + (a.pontoTipoRotulo || '') + ') — somente leitura')
                : (a.somenteLeituraGoogle ? 'Importado da Google Agenda (reserva de voo/hotel) — edições aqui não sincronizam de volta' : ''));
        return '<div class="crm-cal-evento' + (a.feito ? ' crm-cal-evento-feito' : '') + classeOrigem + '" ' + acao + ' data-tipo="' + esc(a.tipo) + '"' +
            (estiloFinal ? (' style="' + estiloFinal + '"') : '') +
            (titulo ? (' title="' + esc(titulo) + '"') : '') + '>' +
            '<span class="crm-cal-evento-icone">' + icone + '</span>' +
            '<span class="crm-cal-evento-txt">' + esc(a.assunto || '(sem assunto)') + '</span>' +
        '</div>';
    }

    /**
     * Eventos que cobrem vários dias vêm como um item por dia, todos
     * compartilhando um id de grupo — `pontoGrupoId` para eventos do Ponto
     * (ver gerarPontoPseudoAtividades), `grupoBarraId` para atividades reais
     * multi-dia (ver expandirAtividadesMultidia — hoje só reservas de
     * hotel/viagem importadas do Google). Nesta semana da grade, acha a
     * coluna (0-6) inicial e final de cada grupo presente, para desenhar UMA
     * barra contínua em vez de um card por dia.
     */
    function acharBarrasPontoNaSemana(semanaDias, porDia) {
        var porGrupo = {};
        semanaDias.forEach(function (dataIso, col) {
            (porDia[dataIso] || []).forEach(function (a) {
                var gid = (a.origemPonto && a.pontoGrupoId) ? a.pontoGrupoId : a.grupoBarraId;
                if (!gid) return;
                var g = porGrupo[gid];
                if (!g) porGrupo[gid] = { inicio: col, fim: col, item: a };
                else g.fim = col; // dias da semana vêm em ordem crescente
            });
        });
        return Object.keys(porGrupo).map(function (gid) { return porGrupo[gid]; })
            .sort(function (a, b) { return a.inicio - b.inicio || (b.fim - b.inicio) - (a.fim - a.inicio); });
    }

    /** Empacota barras em "faixas" (linhas) sem deixar duas com colunas sobrepostas na mesma linha. */
    function empacotarFaixas(barras) {
        var faixas = [];
        barras.forEach(function (b) {
            var faixa = faixas.filter(function (f) {
                return !f.some(function (o) { return !(b.fim < o.inicio || b.inicio > o.fim); });
            })[0];
            if (!faixa) { faixa = []; faixas.push(faixa); }
            faixa.push(b);
        });
        return faixas;
    }

    /**
     * Uma barra da trilha (um evento multi-dia mesclado). `offsetColuna` é 1
     * na grade Mês (7 colunas puras) e 2 na grade Semana (a coluna 1 é o
     * gutter fixo de horário, os dias começam na 2). Eventos do Ponto ficam
     * somente-leitura (cor própria, sem clique); os demais (hoje, reservas do
     * Google) abrem o modal de edição normal, como qualquer atividade.
     */
    function renderizarBarraMultidia(item, colInicio, colFim, offsetColuna) {
        var ehPonto = !!item.origemPonto;
        var tipoInfo = (!ehPonto && window.CrmStore && CrmStore.mapaTiposAtividade()[item.tipo]) || {};
        var cor = ehPonto ? (item.pontoCor || '#64748b') : '#eef2ff';
        var corTexto = ehPonto ? (item.pontoCorTexto || '#fff') : '#4338ca';
        var icone = ehPonto ? (item.pontoIcone || '📅') : (tipoInfo.icone || '📌');
        var estilo = 'grid-column:' + (colInicio + offsetColuna) + ' / ' + (colFim + offsetColuna + 1) + ';' +
            'background:' + esc(cor) + ';color:' + esc(corTexto) + ';' + (ehPonto ? '' : 'cursor:pointer;');
        var titulo = (ehPonto ? ((item.pontoTipoRotulo || '') + ': ' + (item.assunto || '')) : (item.assunto || '(sem assunto)')) +
            (colFim > colInicio ? ' (' + (colFim - colInicio + 1) + ' dias)' : '');
        var acao = ehPonto ? '' : (' data-crm-action="abrirModalAtividadeCal" data-id="' + esc(item.id) + '"');
        return '<div class="crm-cal-mes-barra-ponto" style="' + estilo + '"' + acao + ' title="' + esc(titulo) + '">' +
            '<span class="crm-cal-evento-icone">' + esc(icone) + '</span>' +
            '<span class="crm-cal-evento-txt">' + esc(item.assunto || '(sem assunto)') + '</span>' +
        '</div>';
    }

    function renderizarTrilhaGrade(semanaDias, porDia, offsetColuna, classeLinha) {
        var barras = acharBarrasPontoNaSemana(semanaDias, porDia);
        if (!barras.length) return '';
        var faixas = empacotarFaixas(barras);
        // A trilha é uma linha própria por cima da grade de dias — sem isto,
        // nas semanas em que ela aparece (por causa de uma barra em qualquer
        // dia), a faixa fica branca cruzando a coluna de sábado/domingo, que
        // nas células normais já está tingida (ver crm-cal-fim-semana).
        var fundoFimDeSemana = semanaDias.map(function (dataIso, col) {
            if (!ehFimDeSemana(dataIso)) return '';
            return '<div class="crm-cal-trilha-fundo-fds" style="grid-column:' + (col + offsetColuna) + ' / ' + (col + offsetColuna + 1) + '"></div>';
        }).join('');
        var linhasHtml = faixas.map(function (faixa) {
            var celulas = faixa.map(function (b) { return renderizarBarraMultidia(b.item, b.inicio, b.fim, offsetColuna); }).join('');
            return '<div class="' + classeLinha + '">' + fundoFimDeSemana + celulas + '</div>';
        }).join('');
        return '<div class="crm-cal-mes-trilha">' + linhasHtml + '</div>';
    }

    /**
     * Segmento de uma barra multi-dia DENTRO da célula do dia (mês) — logo
     * abaixo do número, como qualquer outro card (voo, conta...). Ao
     * contrário da trilha da semana (uma faixa à parte por cima da grade),
     * aqui a barra "mora" na própria célula e se conecta visualmente com a
     * célula vizinha: `posicao` controla o arredondamento e sangra o padding
     * de 4px da célula nos lados onde a barra continua (inicio sangra a
     * direita, fim sangra a esquerda, meio sangra os dois — ver CSS).
     */
    function renderizarSegmentoBarraMultidia(item, posicao) {
        var ehPonto = !!item.origemPonto;
        var tipoInfo = (!ehPonto && window.CrmStore && CrmStore.mapaTiposAtividade()[item.tipo]) || {};
        var cor = ehPonto ? (item.pontoCor || '#64748b') : '#eef2ff';
        var corTexto = ehPonto ? (item.pontoCorTexto || '#fff') : '#4338ca';
        var icone = ehPonto ? (item.pontoIcone || '📅') : (tipoInfo.icone || '📌');
        var titulo = (ehPonto ? ((item.pontoTipoRotulo || '') + ': ' + (item.assunto || '')) : (item.assunto || '(sem assunto)'));
        var acao = ehPonto ? '' : (' data-crm-action="abrirModalAtividadeCal" data-id="' + esc(item.id) + '"');
        var mostrarTexto = (posicao === 'inicio' || posicao === 'unica');
        return '<div class="crm-cal-mes-barra-seg crm-cal-mes-barra-seg-' + posicao + '" style="background:' + esc(cor) + ';color:' + esc(corTexto) + ';' + (ehPonto ? '' : 'cursor:pointer;') + '"' + acao + ' title="' + esc(titulo) + '">' +
            (mostrarTexto ? ('<span class="crm-cal-evento-icone">' + esc(icone) + '</span><span class="crm-cal-evento-txt">' + esc(item.assunto || '(sem assunto)') + '</span>') : '') +
        '</div>';
    }

    function renderizarSemanaMes(semanaDias, porDia, mesAtual, hoje, maxVisiveis) {
        var barras = acharBarrasPontoNaSemana(semanaDias, porDia);
        var faixas = empacotarFaixas(barras);

        var celulasDias = semanaDias.map(function (dataIso, col) {
            // Itens multi-dia (Ponto ou reais) já viram segmento de barra abaixo do número — não repetir aqui.
            var itens = (porDia[dataIso] || []).filter(function (a) { return !a.origemPonto && !a.grupoBarraId; });
            var foraDoMes = dataIso.slice(0, 7) !== mesAtual;
            var ehHoje = dataIso === hoje;
            var visiveis = itens.slice(0, maxVisiveis);
            var extras = itens.length - visiveis.length;
            var eventosHtml = visiveis.map(function (a) { return renderizarEventoCard(a); }).join('') +
                (extras > 0 ? '<div class="crm-cal-mes-mais">+' + extras + ' mais</div>' : '');

            var segmentosHtml = faixas.map(function (faixa) {
                var b = faixa.filter(function (bb) { return col >= bb.inicio && col <= bb.fim; })[0];
                if (!b) return '<div class="crm-cal-mes-barra-vazia"></div>';
                var posicao = (col === b.inicio && col === b.fim) ? 'unica' : (col === b.inicio ? 'inicio' : (col === b.fim ? 'fim' : 'meio'));
                return renderizarSegmentoBarraMultidia(b.item, posicao);
            }).join('');

            var numeroMes = ehHoje ? '<span class="crm-cal-cab-numero-hoje">' + Number(dataIso.slice(8, 10)) + '</span>' : Number(dataIso.slice(8, 10));
            return '<div class="crm-cal-mes-dia' + (foraDoMes ? ' crm-cal-mes-dia-fora' : '') + (ehHoje ? ' crm-cal-mes-dia-hoje' : '') + (ehFimDeSemana(dataIso) ? ' crm-cal-fim-semana' : '') + '" data-crm-action="novaAtividadeNoDia" data-data="' + esc(dataIso) + '" title="Clique para agendar uma atividade neste dia">' +
                '<div class="crm-cal-mes-numero">' + numeroMes + '</div>' +
                (faixas.length ? ('<div class="crm-cal-mes-segmentos">' + segmentosHtml + '</div>') : '') +
                '<div class="crm-cal-mes-eventos">' + eventosHtml + '</div>' +
            '</div>';
        }).join('');

        return '<div class="crm-cal-mes-semana"><div class="crm-cal-mes-semana-dias">' + celulasDias + '</div></div>';
    }

    function renderizarCalendarioMes(atividades) {
        var mesRef = _calMesRef;
        var porDia = CrmCalculos.agruparAtividadesPorGradeMes(expandirAtividadesMultidia(atividades), mesRef);
        var dias = Object.keys(porDia);
        var hoje = hojeIsoLocal();
        var mesAtual = mesRef.slice(0, 7);

        var nav = '<div class="crm-cal-nav">' +
            '<button type="button" class="btn-secondary crm-btn-mini" data-crm-action="calMesHoje">Hoje</button>' +
            '<button type="button" class="crm-cal-seta" data-crm-action="calMesAnterior">‹</button>' +
            '<button type="button" class="crm-cal-seta" data-crm-action="calMesProximo">›</button>' +
            '<span class="crm-cal-faixa">' + esc(MESES[Number(mesAtual.slice(5, 7)) - 1]) + ' de ' + mesAtual.slice(0, 4) + '</span>' +
        '</div>';

        var cabecalho = '<div class="crm-cal-mes-cab">' +
            DIAS_SEMANA_ABREV.map(function (d, i) { return '<div class="crm-cal-mes-cab-dia' + ((i === 0 || i === 6) ? ' crm-cal-fim-semana' : '') + '">' + d + '</div>'; }).join('') +
        '</div>';

        var MAX_VISIVEIS = 4;
        var semanasHtml = '';
        for (var w = 0; w < dias.length; w += 7) {
            semanasHtml += renderizarSemanaMes(dias.slice(w, w + 7), porDia, mesAtual, hoje, MAX_VISIVEIS);
        }

        return nav + '<div class="crm-cal-mes-grade">' + cabecalho + '<div class="crm-cal-mes-corpo">' + semanasHtml + '</div></div>';
    }

    function setCalModo(v) { _calModo = v; renderizarCalendarioView(); }
    function setCalPeriodo(v) { _calPeriodo = v; renderizarCalendarioView(); }
    function setCalBusca(v) { _calBusca = v; renderizarCalendarioView(); }
    function setMostrarFeriadosCal(v) { _calMostrarFeriados = v; renderizarCalendarioView(); }
    function setCalTipo(v) { _calTipo = v; renderizarCalendarioView(); }
    function calHoje() { _calSemanaRef = CrmCalculos.inicioSemana(hojeIsoLocal()); renderizarCalendarioView(); }
    function calSemanaAnterior() { _calSemanaRef = CrmCalculos.somarDias(_calSemanaRef, -7); renderizarCalendarioView(); }
    function calSemanaProxima() { _calSemanaRef = CrmCalculos.somarDias(_calSemanaRef, 7); renderizarCalendarioView(); }
    function calMesHoje() { _calMesRef = hojeIsoLocal().slice(0, 7) + '-01'; renderizarCalendarioView(); }
    function calMesAnterior() { _calMesRef = CrmCalculos.somarMeses(_calMesRef, -1); renderizarCalendarioView(); }
    function calMesProximo() { _calMesRef = CrmCalculos.somarMeses(_calMesRef, 1); renderizarCalendarioView(); }

    // ── Modal de Atividade global (vinculada a um negócio, aberto a partir do Calendário) ──

    function abrirModalAtividadeCal(id, dataPreenchida) {
        var atividade = id ? CrmStore.listarAtividades().filter(function (a) { return a.id === id; })[0] : null;
        document.getElementById('crmModalAtvCalTitulo').textContent = id ? 'Editar atividade' : 'Nova atividade';
        document.getElementById('crmAtvCalId').value = id || '';

        _calAtvNegocioId = atividade ? atividade.negocioId : null;
        if (_calAtvNegocioId) {
            selecionarNegocioAtividadeCal(_calAtvNegocioId);
        } else {
            document.getElementById('crmAtvCalNegocioBusca').value = '';
        }
        document.getElementById('crmAtvCalNegocioLista').style.display = 'none';
        document.getElementById('crmAtvCalNegocioBusca').disabled = !!_calAtvNegocioId;

        var tipo = atividade ? atividade.tipo : 'tarefa';
        document.getElementById('crmAtvCalTipo').value = tipo;
        var tiposEl = document.getElementById('crmAtvCalTipos');
        if (tiposEl) {
            tiposEl.innerHTML = CrmStore.listarTiposAtividadeAtivos().map(function (t) {
                return '<button type="button" class="crm-atv-tipo crm-atv-cal-tipo' + (t.chave === tipo ? ' active' : '') + '" data-crm-action="escolherTipoAtividadeCal" data-valor="' + esc(t.chave) + '">' + t.icone + ' ' + esc(t.nome) + '</button>';
            }).join('');
        }

        document.getElementById('crmAtvCalAssunto').value = atividade ? (atividade.assunto || '') : '';
        document.getElementById('crmAtvCalData').value = atividade ? (atividade.data || '') : (dataPreenchida || hojeIsoLocal());
        document.getElementById('crmAtvCalHoraInicio').value = atividade ? (atividade.horaInicio || '') : '';
        document.getElementById('crmAtvCalHoraFim').value = atividade ? (atividade.horaFim || '') : '';
        document.getElementById('crmAtvCalDescricao').value = atividade ? (atividade.descricao || '') : '';
        document.getElementById('crmAtvCalFeito').checked = atividade ? !!atividade.feito : false;

        var btnExcluir = document.getElementById('crmAtvCalBtnExcluir');
        if (btnExcluir) btnExcluir.style.display = id ? '' : 'none';

        document.getElementById('modalAtividadeCal').style.display = 'flex';
        atualizarMiniAgendaCal();
    }

    function escolherTipoAtividadeCal(el) {
        document.getElementById('crmAtvCalTipo').value = el.dataset.valor;
        document.querySelectorAll('.crm-atv-cal-tipo').forEach(function (btn) {
            btn.classList.toggle('active', btn.dataset.valor === el.dataset.valor);
        });
    }

    function buscarNegocioParaAtividadeCal(termo) {
        var lista = document.getElementById('crmAtvCalNegocioLista');
        if (!lista) return;
        var todos = CrmStore.listarNegocios();
        var clientes = CrmStore.listarClientes();
        var clientePorId = {};
        clientes.forEach(function (c) { clientePorId[c.id] = c; });
        var termoNorm = String(termo || '').trim().toLowerCase();
        var filtrados = (termoNorm ? todos.filter(function (n) {
            var cliente = clientePorId[n.clienteId];
            return (n.titulo || '').toLowerCase().indexOf(termoNorm) !== -1 ||
                (cliente && (cliente.nome || '').toLowerCase().indexOf(termoNorm) !== -1);
        }) : todos).slice(0, 8);

        var itens = filtrados.map(function (n) {
            var cliente = clientePorId[n.clienteId];
            return '<div class="crm-ac-item" data-crm-action="selecionarNegocioAtividadeCal" data-id="' + esc(n.id) + '">' +
                esc(n.titulo || '(sem título)') + (cliente ? (' <span style="opacity:.6">· ' + esc(cliente.nome) + '</span>') : '') +
                '</div>';
        }).join('');

        lista.innerHTML = itens || '<div class="crm-ac-vazio">Nenhum negócio encontrado.</div>';
        lista.style.display = 'block';
    }

    function selecionarNegocioAtividadeCal(id) {
        var negocio = CrmStore.listarNegocios().filter(function (n) { return n.id === id; })[0];
        if (!negocio) return;
        _calAtvNegocioId = negocio.id;
        document.getElementById('crmAtvCalNegocioBusca').value = negocio.titulo || '';
        document.getElementById('crmAtvCalNegocioLista').style.display = 'none';
    }

    /** Duas faixas de horário no mesmo dia se sobrepõem — sem hora definida conta como dia inteiro (sempre conflita). */
    function horasSeSobrepoem(a, b) {
        if (!a.horaInicio || !b.horaInicio) return true;
        var iniA = horaParaMin(a.horaInicio), fimA = a.horaFim ? horaParaMin(a.horaFim) : iniA + 30;
        var iniB = horaParaMin(b.horaInicio), fimB = b.horaFim ? horaParaMin(b.horaFim) : iniB + 30;
        if (fimA <= iniA) fimA = iniA + 30;
        if (fimB <= iniB) fimB = iniB + 30;
        return iniA < fimB && iniB < fimA;
    }

    /**
     * Verifica se o compromisso que está sendo salvo coincide com algo no
     * mesmo dia/horário: fim de semana, feriado nacional, evento do Ponto
     * (férias, viagem...) ou outra atividade real já marcada. Devolve uma
     * lista de descrições (vazia = sem conflito) — usada só pra alertar antes
     * de salvar, nunca bloqueia.
     */
    function detectarConflitosAtividade(dados, idExcluir) {
        var conflitos = [];

        if (ehFimDeSemana(dados.data)) {
            var diaSemana = DIAS_SEMANA_ABREV[new Date(dados.data + 'T00:00:00Z').getUTCDay()];
            conflitos.push('cai num fim de semana (' + diaSemana + ')');
        }

        gerarFeriadosPseudoAtividades().forEach(function (f) {
            if (f.data === dados.data) conflitos.push('feriado nacional: ' + (f.assunto || ''));
        });

        if (window.PontoBridge && PontoBridge.conectado() && _calMostrarPonto) {
            gerarPontoPseudoAtividades().forEach(function (p) {
                if (p.data === dados.data) conflitos.push((p.pontoTipoRotulo || 'Ponto') + ': ' + (p.assunto || ''));
            });
        }

        CrmStore.listarAtividades().forEach(function (a) {
            if (a.id === idExcluir || a.data !== dados.data || !horasSeSobrepoem(dados, a)) return;
            conflitos.push('"' + (a.assunto || '(sem assunto)') + '"' + (a.horaInicio ? (' às ' + a.horaInicio) : ''));
        });

        return conflitos;
    }

    function salvarAtividadeCal() {
        if (!requireAdminOrNotify()) return;
        if (!_calAtvNegocioId) { Notifications.error('Selecione o negócio vinculado a esta atividade.'); return; }
        var assunto = document.getElementById('crmAtvCalAssunto').value.trim();
        if (!assunto) { Notifications.error('Assunto é obrigatório.'); return; }
        var data = document.getElementById('crmAtvCalData').value;
        if (!data) { Notifications.error('Data é obrigatória.'); return; }

        var dados = {
            negocioId: _calAtvNegocioId,
            tipo: document.getElementById('crmAtvCalTipo').value,
            assunto: assunto,
            descricao: document.getElementById('crmAtvCalDescricao').value.trim(),
            data: data,
            horaInicio: document.getElementById('crmAtvCalHoraInicio').value || '',
            horaFim: document.getElementById('crmAtvCalHoraFim').value || '',
            feito: document.getElementById('crmAtvCalFeito').checked
        };

        var id = document.getElementById('crmAtvCalId').value || null;
        var conflitos = detectarConflitosAtividade(dados, id);

        if (conflitos.length) {
            confirmarEExecutar('Esse horário coincide com ' + conflitos.join('; ') + '. Salvar mesmo assim?', function () {
                gravarAtividadeCal(dados, id);
            });
            return;
        }
        gravarAtividadeCal(dados, id);
    }

    function gravarAtividadeCal(dados, id) {
        var criado = null;
        if (id) CrmStore.atualizarAtividade(id, dados);
        else criado = CrmStore.criarAtividade(dados);
        sincronizarAtividadeComGoogle(id || (criado && criado.id));

        fecharModal('modalAtividadeCal');
        renderizarConteudoAtivo();
    }

    function excluirAtividadeCal() {
        if (!requireAdminOrNotify()) return;
        var id = document.getElementById('crmAtvCalId').value;
        if (!id) return;
        confirmarEExecutar('Excluir esta atividade?', function () {
            var atividade = CrmStore.listarAtividades().filter(function (a) { return a.id === id; })[0];
            CrmStore.removerAtividade(id);
            removerAtividadeDoGoogle(atividade);
            fecharModal('modalAtividadeCal');
            renderizarConteudoAtivo();
        });
    }

    // ── Modal de Anotação ──

    function buscarNegocioParaAnotacao(termo) {
        var lista = document.getElementById('crmAnotacaoNegocioLista');
        if (!lista) return;
        var todos = CrmStore.listarNegocios();
        var clientes = CrmStore.listarClientes();
        var clientePorId = {};
        clientes.forEach(function (c) { clientePorId[c.id] = c; });
        var termoNorm = String(termo || '').trim().toLowerCase();
        var filtrados = (termoNorm ? todos.filter(function (n) {
            var cliente = clientePorId[n.clienteId];
            return (n.titulo || '').toLowerCase().indexOf(termoNorm) !== -1 ||
                (cliente && (cliente.nome || '').toLowerCase().indexOf(termoNorm) !== -1);
        }) : todos).slice(0, 8);

        var itens = filtrados.map(function (n) {
            var cliente = clientePorId[n.clienteId];
            return '<div class="crm-ac-item" data-crm-action="selecionarNegocioAnotacao" data-id="' + esc(n.id) + '">' +
                esc(n.titulo || '(sem título)') + (cliente ? (' <span style="opacity:.6">· ' + esc(cliente.nome) + '</span>') : '') +
                '</div>';
        }).join('');

        lista.innerHTML = itens || '<div class="crm-ac-vazio">Nenhum negócio encontrado.</div>';
        lista.style.display = 'block';
    }

    function selecionarNegocioAnotacao(id) {
        var crm = CrmStore.getCrm();
        var negocio = crm ? crm.negocios.filter(function (n) { return n.id === id; })[0] : null;
        if (!negocio) return;
        var cliente = CrmStore.getCliente(negocio.clienteId);
        var funil = crm.funis.filter(function (f) { return f.id === negocio.funilId; })[0];

        _anotacaoNegocioId = negocio.id;
        document.getElementById('crmAnotacaoNegocioBusca').value = negocio.titulo || '';
        document.getElementById('crmAnotacaoNegocioLista').style.display = 'none';
        document.getElementById('crmAnotacaoRelacionamento').textContent = funil ? funil.nome : '-';
        document.getElementById('crmAnotacaoClienteNome').textContent = cliente ? cliente.nome : '-';
        alternarModoVinculo();
    }

    function desvincularNegocioAnotacao() {
        _anotacaoNegocioId = null;
        var busca = document.getElementById('crmAnotacaoNegocioBusca');
        if (busca) { busca.value = ''; busca.disabled = false; }
        alternarModoVinculo();
    }

    /**
     * Com negócio, funil/cliente vêm dele (só leitura). Sem negócio, a demanda é
     * avulsa e escolhe funil/etapa/cliente por conta própria.
     */
    function alternarModoVinculo() {
        var temNegocio = !!_anotacaoNegocioId;
        var avulsa = document.getElementById('crmAnotacaoAvulsa');
        var info = document.getElementById('crmAnotacaoVinculoInfo');
        var btnDesvincular = document.getElementById('crmAnotacaoBtnDesvincular');
        if (avulsa) avulsa.style.display = temNegocio ? 'none' : '';
        if (info) info.style.display = temNegocio ? '' : 'none';
        if (btnDesvincular) btnDesvincular.style.display = temNegocio ? '' : 'none';
    }

    function popularSelectsAvulsa(funilId, etapaId, clienteId) {
        var funis = CrmStore.listarFunis();
        var alvoFunil = funilId || (CrmStore.getFunilAtivo() ? CrmStore.getFunilAtivo().id : (funis[0] ? funis[0].id : ''));
        popularSelectPreservandoValor(
            document.getElementById('crmAnotacaoFunil'),
            funis.map(function (f) { return { valor: f.id, rotulo: f.nome }; }),
            alvoFunil
        );
        popularSelectEtapaAnotacao(alvoFunil, etapaId);
        popularSelectPreservandoValor(
            document.getElementById('crmAnotacaoCliente'),
            [{ valor: '', rotulo: '(sem cliente)' }].concat(CrmStore.listarClientes().map(function (c) {
                return { valor: c.id, rotulo: c.nome };
            })),
            clienteId || ''
        );
    }

    function popularSelectEtapaAnotacao(funilId, etapaId) {
        var funil = CrmStore.listarFunis().filter(function (f) { return f.id === funilId; })[0];
        var etapas = funil ? funil.etapas : [];
        var alvo = etapaId && etapas.some(function (e) { return e.id === etapaId; })
            ? etapaId
            : ((etapas.filter(function (e) { return e.tipo === 'aberta'; })[0] || etapas[0] || {}).id || '');
        popularSelectPreservandoValor(
            document.getElementById('crmAnotacaoEtapa'),
            etapas.map(function (e) { return { valor: e.id, rotulo: e.nome }; }),
            alvo
        );
    }

    function aoTrocarFunilAnotacao() {
        var sel = document.getElementById('crmAnotacaoFunil');
        popularSelectEtapaAnotacao(sel ? sel.value : null, null);
    }

    function popularSelectResponsavelAnotacao(valorAtual) {
        var reps = (window.estoque && Array.isArray(estoque.representantes)) ? estoque.representantes : [];
        popularSelectPreservandoValor(
            document.getElementById('crmAnotacaoResponsavel'),
            [{ valor: '', rotulo: '(sem responsável)' }].concat(reps.map(function (r) { return { valor: r, rotulo: r }; })),
            valorAtual || ''
        );
    }

    // ── Encaminhamentos dentro do modal ──

    /**
     * Os encaminhamentos são editados em memória (`_encaminhamentosTemp`) e só
     * vão para o store ao salvar a demanda — mesmo padrão de `_itensTemp`.
     */
    function renderizarEncaminhamentosModal() {
        var cont = document.getElementById('crmAnotacaoEncaminhamentos');
        if (!cont) return;
        var hoje = hojeIsoLocal();

        if (!_encaminhamentosTemp.length) {
            cont.innerHTML = '<div class="crm-enc-vazio">Nenhum encaminhamento. Use "+ Encaminhar" quando pedir informação a outra pessoa.</div>';
            renderizarRespostasResumo();
            return;
        }

        cont.innerHTML = _encaminhamentosTemp.map(function (e, i) {
            var atrasado = (e.status === 'pendente' && e.prazoResposta && CrmCalculos.diasAte(e.prazoResposta, hoje) < 0);
            var respondido = e.status === 'respondido';
            return '<div class="crm-enc-linha' + (atrasado ? ' crm-enc-atrasada' : '') + '" data-idx="' + i + '">' +
                '<div class="crm-enc-cabecalho">' +
                    '<span class="crm-enc-num">#' + (i + 1) + '</span>' +
                    '<span class="crm-enc-status" data-status="' + esc(e.status) + '">' + esc(STATUS_ENC_ROTULO[e.status] || e.status) + '</span>' +
                    (atrasado ? '<span class="crm-enc-atraso">⚠ prazo vencido</span>' : '') +
                    '<button type="button" class="crm-btn-mini" data-crm-action="marcarEncRespondido" data-idx="' + i + '"' +
                        (respondido ? ' disabled' : '') + '>✓ Marcar respondida</button>' +
                    '<button type="button" class="crm-btn-mini crm-btn-mini-perigo" data-crm-action="removerEncTemp" data-idx="' + i + '">Remover</button>' +
                '</div>' +
                '<div class="crm-form-linha">' +
                    campoEnc(i, 'para', 'Para quem *', 'text', e.para, 'Ex: Setor Jurídico') +
                    campoEnc(i, 'canal', 'Canal', 'text', e.canal, 'Ex: E-mail, SEI') +
                    campoEnc(i, 'numeroDocumentoEnvio', 'Nº doc enviado', 'text', e.numeroDocumentoEnvio, '') +
                '</div>' +
                '<div class="crm-form-linha">' +
                    campoEnc(i, 'dataEnvio', 'Enviado em', 'date', e.dataEnvio, '') +
                    campoEnc(i, 'prazoResposta', 'Prazo p/ responder', 'date', e.prazoResposta, '') +
                    campoEnc(i, 'dataResposta', 'Respondido em', 'date', e.dataResposta, '') +
                '</div>' +
                '<div class="form-group">' +
                    '<label>O que foi pedido</label>' +
                    '<textarea rows="2" data-enc-idx="' + i + '" data-enc-campo="oQuePedido">' + esc(e.oQuePedido || '') + '</textarea>' +
                '</div>' +
                '<div class="form-group">' +
                    '<label>Resposta recebida</label>' +
                    '<textarea rows="2" data-enc-idx="' + i + '" data-enc-campo="resposta">' + esc(e.resposta || '') + '</textarea>' +
                '</div>' +
            '</div>';
        }).join('');

        renderizarRespostasResumo();
    }

    function campoEnc(idx, campo, rotulo, tipo, valor, placeholder) {
        return '<div class="form-group">' +
            '<label>' + esc(rotulo) + '</label>' +
            '<input type="' + tipo + '" data-enc-idx="' + idx + '" data-enc-campo="' + campo + '"' +
                ' value="' + esc(valor || '') + '"' +
                (placeholder ? ' placeholder="' + esc(placeholder) + '"' : '') + '>' +
        '</div>';
    }

    /** Lê de volta o que foi digitado nos campos dos encaminhamentos. */
    function coletarEncaminhamentosDoDom() {
        var cont = document.getElementById('crmAnotacaoEncaminhamentos');
        if (!cont) return;
        cont.querySelectorAll('[data-enc-idx]').forEach(function (input) {
            var idx = Number(input.dataset.encIdx);
            var alvo = _encaminhamentosTemp[idx];
            if (!alvo) return;
            var campo = input.dataset.encCampo;
            var valor = input.value;
            var ehData = (campo === 'dataEnvio' || campo === 'prazoResposta' || campo === 'dataResposta');
            alvo[campo] = ehData ? (valor || null) : valor.trim();
        });
    }

    function adicionarEncaminhamento() {
        coletarEncaminhamentosDoDom();
        _encaminhamentosTemp.push(CrmModel.criarEncaminhamento({
            dataEnvio: hojeIsoLocal(),
            status: 'pendente'
        }));
        renderizarEncaminhamentosModal();
        sugerirSituacaoNoModal();
    }

    function removerEncTemp(idx) {
        coletarEncaminhamentosDoDom();
        _encaminhamentosTemp.splice(idx, 1);
        renderizarEncaminhamentosModal();
        sugerirSituacaoNoModal();
    }

    function marcarEncRespondido(idx) {
        coletarEncaminhamentosDoDom();
        var e = _encaminhamentosTemp[idx];
        if (!e) return;
        e.status = 'respondido';
        if (!e.dataResposta) e.dataResposta = hojeIsoLocal();
        renderizarEncaminhamentosModal();
        sugerirSituacaoNoModal();
    }

    /** Reflete no select de situação o que os encaminhamentos indicam. */
    function sugerirSituacaoNoModal() {
        var sel = document.getElementById('crmAnotacaoSituacao');
        if (!sel) return;
        var sugerida = CrmCalculos.situacaoSugerida({
            situacao: sel.value,
            encaminhamentos: _encaminhamentosTemp
        }, hojeIsoLocal());
        if (sugerida !== sel.value) sel.value = sugerida;
    }

    function renderizarRespostasResumo() {
        var el = document.getElementById('crmAnotacaoRespostasResumo');
        if (!el) return;
        var respondidos = _encaminhamentosTemp.filter(function (e) { return e.status === 'respondido' && e.resposta; });
        var pendentes = _encaminhamentosTemp.filter(function (e) { return e.status === 'pendente'; });

        if (!_encaminhamentosTemp.length) {
            el.innerHTML = '<div class="crm-enc-vazio">Sem encaminhamentos — nada a consolidar.</div>';
            return;
        }
        var html = '';
        if (pendentes.length) {
            html += '<div class="crm-resumo-aviso">⏳ Ainda aguardando: ' +
                esc(pendentes.map(function (e) { return e.para || '(?)'; }).join(', ')) + '</div>';
        }
        html += respondidos.length
            ? respondidos.map(function (e) {
                return '<div class="crm-resposta-item"><strong>' + esc(e.para || '(?)') + '</strong>' +
                    (e.dataResposta ? ' <span class="crm-resposta-data">' + esc(dataOuTraco(e.dataResposta)) + '</span>' : '') +
                    '<div>' + esc(e.resposta) + '</div></div>';
            }).join('')
            : '<div class="crm-enc-vazio">Nenhuma resposta com conteúdo ainda.</div>';
        el.innerHTML = html;
    }

    /**
     * Anexa as respostas recebidas ao campo de resposta consolidada, sem
     * sobrescrever o que já estava escrito e sem repetir o que já foi colado.
     */
    function consolidarRespostas() {
        coletarEncaminhamentosDoDom();
        var campo = document.getElementById('crmAnotacaoOQueFoiFeito');
        if (!campo) return;
        var respondidos = _encaminhamentosTemp.filter(function (e) { return e.status === 'respondido' && e.resposta; });
        if (!respondidos.length) {
            Notifications.error('Nenhuma resposta recebida para consolidar.');
            return;
        }
        var atual = campo.value;
        var novos = respondidos
            .map(function (e) { return '— ' + (e.para || '(?)') + ': ' + e.resposta; })
            .filter(function (bloco) { return atual.indexOf(bloco) === -1; });

        if (!novos.length) {
            Notifications.error('As respostas já estão consolidadas no texto.');
            return;
        }
        campo.value = (atual ? atual.replace(/\s*$/, '') + '\n' : '') + novos.join('\n');
        Notifications.success(novos.length + ' resposta(s) adicionada(s).');
    }

    function abrirModalAnotacao(id, negocioForcado, relacionadaForcada) {
        var anotacao = id ? CrmStore.getAnotacao(id) : null;
        var val = function (campo, v) { document.getElementById(campo).value = v; };
        document.getElementById('crmModalAnotacaoTitulo').textContent = id ? 'Editar demanda' : 'Nova demanda';
        val('crmAnotacaoId', id || '');

        _anotacaoRelacionadaId = anotacao ? anotacao.anotacaoRelacionadaId : (relacionadaForcada || null);
        var pai = _anotacaoRelacionadaId ? CrmStore.getAnotacao(_anotacaoRelacionadaId) : null;
        var vinculoEl = document.getElementById('crmAnotacaoVinculo');
        if (vinculoEl) {
            vinculoEl.style.display = pai ? '' : 'none';
            vinculoEl.textContent = pai ? ('Vinculada a: ' + (pai.assunto || '(sem assunto)')) : '';
        }

        _anotacaoNegocioId = anotacao ? anotacao.negocioId : (negocioForcado || null);
        document.getElementById('crmAnotacaoNegocioLista').style.display = 'none';
        document.getElementById('crmAnotacaoNegocioBusca').disabled = false;
        if (_anotacaoNegocioId) {
            selecionarNegocioAnotacao(_anotacaoNegocioId);
        } else {
            val('crmAnotacaoNegocioBusca', '');
            document.getElementById('crmAnotacaoRelacionamento').textContent = '-';
            document.getElementById('crmAnotacaoClienteNome').textContent = '-';
        }
        popularSelectsAvulsa(
            anotacao ? anotacao.funilId : null,
            anotacao ? anotacao.etapaId : null,
            anotacao ? anotacao.clienteId : null
        );
        alternarModoVinculo();

        val('crmAnotacaoAssunto', anotacao ? (anotacao.assunto || '') : '');
        val('crmAnotacaoRemetente', anotacao ? (anotacao.remetente || '') : '');
        val('crmAnotacaoOrigemDemanda', anotacao ? (anotacao.origemDemanda || '') : '');
        val('crmAnotacaoDataSolicitacao', anotacao ? (anotacao.dataSolicitacao || '') : '');
        val('crmAnotacaoTipoDoc', anotacao ? (anotacao.tipoDoc || '') : '');
        val('crmAnotacaoNumeroDocumento', anotacao ? (anotacao.numeroDocumento || '') : '');
        val('crmAnotacaoAcaoRealizar', anotacao ? (anotacao.acaoRealizar || '') : '');
        val('crmAnotacaoSituacao', anotacao ? anotacao.situacao : 'recebida');
        popularSelectResponsavelAnotacao(anotacao ? anotacao.responsavel : '');
        val('crmAnotacaoPrioridade', anotacao ? anotacao.prioridade : 'media');
        val('crmAnotacaoPrazo', anotacao ? (anotacao.prazo || '') : '');
        val('crmAnotacaoLembrarDias', (anotacao && anotacao.lembrarDiasAntes !== null) ? String(anotacao.lembrarDiasAntes) : '');
        val('crmAnotacaoTags', anotacao ? (anotacao.tags || []).join(', ') : '');
        val('crmAnotacaoObservacoes', anotacao ? (anotacao.observacoes || '') : '');
        val('crmAnotacaoDataConclusao', anotacao ? (anotacao.dataConclusao || '') : '');
        val('crmAnotacaoOQueFoiFeito', anotacao ? (anotacao.oQueFoiFeito || '') : '');
        val('crmAnotacaoRespostaTipoDoc', anotacao ? (anotacao.respostaTipoDoc || '') : '');
        val('crmAnotacaoRespostaNumeroDocumento', anotacao ? (anotacao.respostaNumeroDocumento || '') : '');

        // Cópia profunda: editar no modal não deve tocar o store antes de salvar.
        _encaminhamentosTemp = anotacao
            ? (anotacao.encaminhamentos || []).map(function (e) { return Object.assign({}, e); })
            : [];
        renderizarEncaminhamentosModal();

        var btnExcluir = document.getElementById('crmAnotacaoBtnExcluir');
        if (btnExcluir) btnExcluir.style.display = id ? '' : 'none';

        document.getElementById('modalAnotacao').style.display = 'flex';
    }

    function salvarAnotacao() {
        if (!requireAdminOrNotify()) return;
        coletarEncaminhamentosDoDom();

        var el = function (campo) { return document.getElementById(campo); };
        var avulsa = !_anotacaoNegocioId;
        var tagsRaw = el('crmAnotacaoTags').value || '';
        var lembrarRaw = el('crmAnotacaoLembrarDias').value.trim();

        var dados = {
            negocioId: _anotacaoNegocioId,
            funilId: avulsa ? (el('crmAnotacaoFunil').value || null) : null,
            etapaId: avulsa ? (el('crmAnotacaoEtapa').value || null) : null,
            clienteId: avulsa ? (el('crmAnotacaoCliente').value || null) : null,
            anotacaoRelacionadaId: _anotacaoRelacionadaId,
            assunto: el('crmAnotacaoAssunto').value.trim(),
            remetente: el('crmAnotacaoRemetente').value.trim(),
            origemDemanda: el('crmAnotacaoOrigemDemanda').value.trim(),
            dataSolicitacao: el('crmAnotacaoDataSolicitacao').value || null,
            tipoDoc: el('crmAnotacaoTipoDoc').value.trim(),
            numeroDocumento: el('crmAnotacaoNumeroDocumento').value.trim(),
            acaoRealizar: el('crmAnotacaoAcaoRealizar').value.trim(),
            situacao: el('crmAnotacaoSituacao').value,
            responsavel: el('crmAnotacaoResponsavel').value,
            encaminhamentos: _encaminhamentosTemp,
            prioridade: el('crmAnotacaoPrioridade').value,
            prazo: el('crmAnotacaoPrazo').value || null,
            lembrarDiasAntes: lembrarRaw === '' ? null : Number(lembrarRaw),
            tags: tagsRaw.split(',').map(function (t) { return t.trim(); }).filter(Boolean),
            observacoes: el('crmAnotacaoObservacoes').value.trim(),
            dataConclusao: el('crmAnotacaoDataConclusao').value || null,
            oQueFoiFeito: el('crmAnotacaoOQueFoiFeito').value.trim(),
            respostaTipoDoc: el('crmAnotacaoRespostaTipoDoc').value.trim(),
            respostaNumeroDocumento: el('crmAnotacaoRespostaNumeroDocumento').value.trim()
        };

        // Validação central do model — antes isto era feito à mão aqui e os
        // formatos de data escapavam.
        var erros = CrmModel.validarAnotacao(dados);
        if (erros.length) { Notifications.error(erros[0]); return; }

        // Normaliza antes de gravar para que os derivados (finalizado,
        // dataConclusao, tipos) fiquem coerentes já na criação.
        var normalizada = CrmModel.normalizarAnotacao(dados);

        var id = el('crmAnotacaoId').value || null;
        if (id) {
            delete normalizada.id;
            delete normalizada.criadoEm;
            delete normalizada.excluidoEm; // salvar não deve ressuscitar uma demanda excluída
            CrmStore.atualizarAnotacao(id, normalizada);
        } else {
            CrmStore.criarAnotacao(normalizada);
        }
        fecharModal('modalAnotacao');
        renderizarConteudoAtivo();
    }

    function excluirAnotacao() {
        if (!requireAdminOrNotify()) return;
        var id = document.getElementById('crmAnotacaoId').value;
        if (!id) return;
        if (!confirm('Excluir esta demanda? Ela sai das listas, mas o registro é preservado.')) return;
        CrmStore.removerAnotacao(id);
        fecharModal('modalAnotacao');
        renderizarConteudoAtivo();
    }

    // ── Previsão ──

    function formatarMesAno(mesIso) {
        var partes = mesIso.split('-');
        return MESES[parseInt(partes[1], 10) - 1] + ' de ' + partes[0];
    }

    function renderizarPrevisaoView(funil) {
        var negocios = CrmCalculos.filtrarNegocios(
            CrmStore.listarNegocios(funil.id).filter(function (n) { return n.status === 'aberto'; }),
            { busca: _busca }
        );
        var grupos = CrmCalculos.agruparPorMesFechamento(negocios);
        var atividades = CrmStore.listarAtividades();

        var html = grupos.map(function (g) {
            var titulo = g.mes ? formatarMesAno(g.mes) : 'Sem previsão';
            var soma = funil.mostrarValor ? esc(CrmCalculos.formatarMoeda(CrmCalculos.somarValor(g.negocios))) : '';
            var cards = g.negocios.map(function (n) { return CrmKanban.renderizarCard(n, funil.mostrarValor, atividades); }).join('');
            return '' +
                '<div class="crm-previsao-grupo">' +
                    '<div class="crm-previsao-header"><h4>' + esc(titulo) + '</h4>' +
                        '<span>' + g.negocios.length + ' negócio' + (g.negocios.length !== 1 ? 's' : '') + (soma ? (' · ' + soma) : '') + '</span></div>' +
                    '<div class="crm-previsao-cards">' + (cards || '<div class="crm-coluna-vazia">Nenhum negócio</div>') + '</div>' +
                '</div>';
        }).join('');

        document.getElementById('crmPrevisao').innerHTML = html || '<p>Nenhum negócio em aberto.</p>';
        atualizarContagem(negocios.length);
    }

    // ── Excluídos ──

    function renderizarExcluidosView(funil) {
        var negocios = CrmStore.listarNegociosExcluidos(funil.id);
        var etapaPorId = {};
        funil.etapas.forEach(function (e) { etapaPorId[e.id] = e; });

        var linhas = negocios.map(function (n) {
            var etapa = etapaPorId[n.etapaId];
            var cliente = CrmKanban.nomeCliente(n);
            return '<tr>' +
                '<td>' + esc(n.titulo || '(sem título)') + '</td>' +
                '<td>' + esc(cliente || '—') + '</td>' +
                '<td>' + esc(etapa ? etapa.nome : '—') + '</td>' +
                '<td>' + esc(DateUtils.formatBR(n.excluidoEm)) + '</td>' +
                '<td>' +
                    '<button type="button" class="btn-secondary crm-btn-mini" data-crm-action="restaurarNegocio" data-id="' + esc(n.id) + '">Restaurar</button> ' +
                    '<button type="button" class="btn-secondary crm-btn-mini" data-crm-action="excluirDefinitivo" data-id="' + esc(n.id) + '">Excluir de vez</button>' +
                '</td>' +
            '</tr>';
        }).join('');

        document.getElementById('crmExcluidos').innerHTML =
            '<div class="table-container"><table><thead><tr><th>Título</th><th>Cliente</th><th>Etapa</th><th>Excluído em</th><th>Ações</th></tr></thead><tbody>' +
            (linhas || '<tr><td colspan="5">A lixeira está vazia.</td></tr>') + '</tbody></table></div>';
        atualizarContagem(negocios.length);
    }

    // ──────────────────────────────────────────────
    //  DETALHE
    // ──────────────────────────────────────────────

    function secaoColapsavel(chave, titulo, corpoHtml) {
        var colapsada = !!_secoesColapsadas[chave];
        return '' +
            '<div class="crm-secao">' +
                '<div class="crm-secao-header" data-crm-action="toggleSecao" data-secao="' + chave + '">' +
                    '<span class="crm-secao-seta">' + (colapsada ? '▸' : '▾') + '</span>' +
                    '<span class="crm-secao-titulo">' + esc(titulo) + '</span>' +
                '</div>' +
                (colapsada ? '' : ('<div class="crm-secao-corpo">' + corpoHtml + '</div>')) +
            '</div>';
    }

    function renderizarDetalhe(id) {
        var el = document.getElementById('crmViewDetalhe');
        var crm = CrmStore.getCrm();
        var negocio = crm.negocios.filter(function (n) { return n.id === id; })[0];
        if (!negocio) { el.innerHTML = '<p>Negócio não encontrado.</p>'; _detalheId = null; return; }
        var funil = crm.funis.filter(function (f) { return f.id === negocio.funilId; })[0];
        var etapaAtual = funil ? funil.etapas.filter(function (e) { return e.id === negocio.etapaId; })[0] : null;
        var cliente = negocio.clienteId ? CrmStore.getCliente(negocio.clienteId) : null;
        var atividades = CrmStore.listarAtividades(id);

        el.innerHTML = '' +
            '<button type="button" class="btn-secondary crm-btn-mini" data-crm-action="voltarLista" style="margin-bottom:10px">← Voltar</button>' +
            renderizarHeaderDetalhe(negocio, funil, etapaAtual) +
            '<div class="crm-detalhe-grid crm-detalhe-grid-2col">' +
                '<div class="crm-det-esquerda">' + renderizarPainelEsquerdo(negocio, cliente) + '</div>' +
                '<div>' + renderizarPainelCentro(negocio, atividades) + '</div>' +
            '</div>';
    }

    function renderizarHeaderDetalhe(negocio, funil, etapaAtual) {
        var badge = negocio.status === 'ganho' ? '<span class="crm-badge-status crm-badge-ganho">GANHO</span>'
            : negocio.status === 'perdido' ? '<span class="crm-badge-status crm-badge-perdido">PERDIDO</span>' : '';
        var valorTxt = (funil && funil.mostrarValor && negocio.valor !== null && negocio.valor !== undefined)
            ? esc(CrmCalculos.formatarMoeda(negocio.valor, negocio.moeda)) : '';

        var barra = '';
        if (funil) {
            var idxAtual = -1;
            funil.etapas.forEach(function (e, i) { if (e.id === negocio.etapaId) idxAtual = i; });
            barra = '<div class="crm-prog-barra">' + funil.etapas.map(function (e, i) {
                var cls = 'crm-prog-seg';
                var texto = esc(e.nome);
                if (i < idxAtual) cls += ' crm-prog-passada';
                else if (i === idxAtual) {
                    cls += ' crm-prog-atual';
                    texto += ' · ' + CrmCalculos.diasNaEtapa(CrmStore.getCrm().historico, negocio) + ' dias';
                }
                return '<button type="button" class="' + cls + '" data-crm-action="moverEtapaProgresso" data-etapa-id="' + esc(e.id) + '">' + texto + '</button>';
            }).join('') + '</div>';
        }

        return '' +
            '<div class="crm-det-header">' +
                '<div class="crm-det-header-linha">' +
                    '<h2 class="crm-det-titulo">' + esc(negocio.titulo || '(sem título)') + '</h2>' +
                    badge +
                    (valorTxt ? '<span class="crm-det-valor">' + valorTxt + '</span>' : '') +
                    '<div style="margin-left:auto;display:flex;gap:8px;align-items:center">' +
                        (negocio.status === 'aberto'
                            ? ('<button type="button" class="crm-btn-ganho" data-crm-action="marcarGanho" data-id="' + esc(negocio.id) + '">Ganho</button>' +
                               '<button type="button" class="crm-btn-perdido" data-crm-action="marcarPerdido" data-id="' + esc(negocio.id) + '">Perdido</button>')
                            : '') +
                        '<button type="button" class="btn-secondary crm-btn-mini" data-crm-action="editarNegocio" data-id="' + esc(negocio.id) + '">Editar</button>' +
                        '<button type="button" class="btn-secondary crm-btn-mini" data-crm-action="excluirNegocio" data-id="' + esc(negocio.id) + '">🗑</button>' +
                    '</div>' +
                '</div>' +
                barra +
            '</div>';
    }

    function renderizarPainelEsquerdo(negocio, cliente) {
        var resumo = '' +
            '<div>Cliente: ' + (cliente ? esc(cliente.nome) : '—') + '</div>' +
            '<div>Etiquetas: ' + ((negocio.tags || []).map(function (t) { return '<span class="crm-tag">' + esc(t) + '</span>'; }).join(' ') || '—') + '</div>' +
            '<div>Fechamento esperado: ' + (negocio.dataPrevisao ? esc(DateUtils.formatBR(negocio.dataPrevisao)) : '—') + '</div>' +
            (negocio.status === 'perdido' && negocio.motivoPerda ? ('<div>Motivo da perda: ' + esc(negocio.motivoPerda) + '</div>') : '');

        var itensHtml = (negocio.itens && negocio.itens.length)
            ? negocio.itens.map(function (it) {
                return '<div class="crm-item-linha"><span>' + esc(it.nome) + '</span><span>' + it.quantidade + ' × ' + esc(CrmCalculos.formatarMoeda(it.precoUnit)) + '</span></div>';
            }).join('')
            : '<div class="crm-coluna-vazia">Nenhum item.</div>';

        var fonte = '' +
            '<div>Origem: ' + (negocio.origem ? esc(negocio.origem) : '—') + '</div>' +
            '<div>Recebido em: ' + (negocio.dataRecebimento ? esc(DateUtils.formatBR(negocio.dataRecebimento)) : '—') + '</div>';

        var clienteHtml = cliente
            ? ('<div>' + esc(cliente.nome) + (cliente.contato ? (' (' + esc(cliente.contato) + ')') : '') + '</div>' +
               '<div>' + esc(cliente.telefone || '—') + '</div>' +
               '<div>' + esc(cliente.email || '—') + '</div>' +
               '<div>' + esc(cliente.cnpj || '—') + '</div>' +
               '<div>' + esc([cliente.endereco, cliente.cidade, cliente.uf].filter(Boolean).join(', ') || '—') + '</div>' +
               '<button type="button" class="crm-btn-mini" style="margin-top:8px" data-crm-action="editarCliente" ' +
                   'data-id="' + esc(cliente.id) + '" data-origem="detalhe">✎ Editar cadastro</button>')
            : '<div class="crm-coluna-vazia">Nenhum cliente vinculado.</div>';

        var visaoGeral = '' +
            '<div>Idade: ' + CrmCalculos.idadeEmDias(negocio) + ' dias</div>' +
            '<div>Inativo há: ' + CrmCalculos.diasInativo(negocio, CrmStore.listarAtividades()) + ' dias</div>' +
            '<div>Criado em: ' + esc(DateUtils.formatBR(negocio.criadoEm)) + '</div>';

        return '' +
            secaoColapsavel('resumo', 'Resumo', resumo) +
            secaoColapsavel('itens', 'Produtos', itensHtml) +
            secaoColapsavel('fonte', 'Fonte', fonte) +
            secaoColapsavel('cliente', 'Cliente', clienteHtml) +
            secaoColapsavel('visaogeral', 'Visão geral', visaoGeral);
    }

    function renderizarPainelCentro(negocio, atividades) {
        var abas = '' +
            '<div class="crm-det-abas">' +
                '<button type="button" class="crm-det-aba' + (_abaDetalhe === 'atividade' ? ' active' : '') + '" data-crm-action="trocarAbaDetalhe" data-valor="atividade">Atividade</button>' +
                '<button type="button" class="crm-det-aba' + (_abaDetalhe === 'anotacoes' ? ' active' : '') + '" data-crm-action="trocarAbaDetalhe" data-valor="anotacoes">Anotações</button>' +
                '<span class="crm-det-aba crm-det-aba-off" title="Indisponível nesta versão">Chamada</span>' +
                '<span class="crm-det-aba crm-det-aba-off" title="Indisponível nesta versão">E-mail</span>' +
                '<span class="crm-det-aba crm-det-aba-off" title="Indisponível nesta versão">Arquivos</span>' +
                '<span class="crm-det-aba crm-det-aba-off" title="Indisponível nesta versão">Documentos</span>' +
            '</div>';

        var corpoAba = _abaDetalhe === 'anotacoes' ? renderizarListaAnotacoesNegocio(negocio) : renderizarComposerAtividade(negocio);

        return abas + corpoAba +
            '<div class="crm-det-bloco">' + renderizarFoco(negocio, atividades) + '</div>' +
            '<div class="crm-det-bloco">' + renderizarHistoricoFiltrado(negocio) + '</div>';
    }

    function renderizarComposerAtividade(negocio) {
        var editando = _atividadeEditandoId ? CrmStore.listarAtividades(negocio.id).filter(function (a) { return a.id === _atividadeEditandoId; })[0] : null;
        var tipo = editando ? editando.tipo : 'tarefa';
        var pills = CrmStore.listarTiposAtividadeAtivos().map(function (t) {
            return '<button type="button" class="crm-atv-tipo' + (t.chave === tipo ? ' active' : '') + '" data-crm-action="escolherTipoAtividade" data-valor="' + esc(t.chave) + '">' + t.icone + ' ' + esc(t.nome) + '</button>';
        }).join('');

        return '' +
            '<div class="crm-atv-composer">' +
                '<input type="hidden" id="crmAtvId" value="' + (editando ? esc(editando.id) : '') + '">' +
                '<input type="hidden" id="crmAtvTipo" value="' + esc(tipo) + '">' +
                '<div class="crm-atv-tipos">' + pills + '</div>' +
                '<input type="text" id="crmAtvAssunto" placeholder="Assunto" value="' + (editando ? esc(editando.assunto) : '') + '">' +
                '<div class="crm-atv-linha">' +
                    '<input type="date" id="crmAtvData" value="' + (editando ? esc(editando.data || '') : '') + '">' +
                    '<input type="time" id="crmAtvHoraInicio" value="' + (editando ? esc(editando.horaInicio || '') : '') + '">' +
                    '<span>–</span>' +
                    '<input type="time" id="crmAtvHoraFim" value="' + (editando ? esc(editando.horaFim || '') : '') + '">' +
                '</div>' +
                '<textarea id="crmAtvDescricao" rows="2" placeholder="Descrição (opcional)">' + (editando ? esc(editando.descricao) : '') + '</textarea>' +
                '<div class="crm-atv-acoes">' +
                    (editando ? '<button type="button" class="btn-secondary" data-crm-action="cancelarEdicaoAtividade">Cancelar edição</button>' : '') +
                    '<button type="button" class="btn btn-primary" data-crm-action="salvarAtividade" data-negocio-id="' + esc(negocio.id) + '">' + (editando ? 'Salvar alterações' : 'Salvar') + '</button>' +
                '</div>' +
            '</div>';
    }

    function formatarDataHora(iso) {
        if (!iso) return '';
        var d = new Date(iso);
        if (isNaN(d.getTime())) return DateUtils.formatBR(iso);
        var hh = String(d.getHours()).padStart(2, '0');
        var mm = String(d.getMinutes()).padStart(2, '0');
        return DateUtils.formatBR(iso) + ' às ' + hh + ':' + mm;
    }

    function renderizarCardAnotacao(a, negocioId) {
        var prioridadeBadge = '<span class="crm-prioridade-badge" data-prioridade="' + esc(a.prioridade) + '">' +
            esc(PRIORIDADE_ROTULO[a.prioridade] || a.prioridade) + '</span>';
        var prazoTxt = a.prazo ? DateUtils.formatBR(a.prazo) : '';
        var resumo = CrmCalculos.resumoEncaminhamentos(a, hojeIsoLocal());
        var aguardando = resumo.comQuem.length
            ? '<div class="crm-anot-nota-aguardando' + (resumo.algumVencido ? ' crm-atrasado' : '') + '">⏳ Aguardando: ' +
                esc(resumo.comQuem.join(', ')) + (resumo.algumVencido ? ' ⚠' : '') + '</div>'
            : '';
        return '' +
            '<div class="crm-anot-linha">' +
                '<span class="crm-anot-icone" title="Demanda">🗒️</span>' +
                '<div class="crm-anot-nota' + (a.finalizado ? ' crm-anot-finalizada' : '') + '">' +
                    '<div data-crm-action="abrirModalAnotacao" data-id="' + esc(a.id) + '">' +
                        '<div class="crm-anot-nota-data">' + esc(formatarDataHora(a.criadoEm)) + (a.finalizado ? ' · ✓ Finalizado' : '') + '</div>' +
                        '<div class="crm-anot-nota-assunto">' + esc(a.assunto || '(sem assunto)') + '</div>' +
                        '<div class="crm-anot-nota-meta">' + badgeSituacao(a) + prioridadeBadge + (prazoTxt ? ('<span>Prazo: ' + esc(prazoTxt) + '</span>') : '') + '</div>' +
                        aguardando +
                        (a.observacoes ? '<div class="crm-anot-nota-obs">' + esc(a.observacoes) + '</div>' : '') +
                    '</div>' +
                    '<button type="button" class="crm-anot-btn-vincular" data-crm-action="novaAnotacaoNegocio" ' +
                        'data-negocio-id="' + esc(negocioId) + '" data-relacionada-id="' + esc(a.id) + '" title="Nova anotação vinculada a esta">' +
                        '🗒️ Vincular anotação' +
                    '</button>' +
                '</div>' +
            '</div>';
    }

    function renderizarListaAnotacoesNegocio(negocio) {
        var todas = CrmStore.listarAnotacoes(negocio.id);
        var idsValidos = {};
        todas.forEach(function (a) { idsValidos[a.id] = true; });

        var filhasPorPai = {};
        var raizes = [];
        todas.forEach(function (a) {
            if (a.anotacaoRelacionadaId && idsValidos[a.anotacaoRelacionadaId]) {
                (filhasPorPai[a.anotacaoRelacionadaId] = filhasPorPai[a.anotacaoRelacionadaId] || []).push(a);
            } else {
                raizes.push(a);
            }
        });
        raizes.sort(function (a, b) { return String(b.criadoEm || '').localeCompare(String(a.criadoEm || '')); });
        Object.keys(filhasPorPai).forEach(function (pid) {
            filhasPorPai[pid].sort(function (a, b) { return String(a.criadoEm || '').localeCompare(String(b.criadoEm || '')); });
        });

        var threads = raizes.map(function (raiz) {
            var filhas = filhasPorPai[raiz.id] || [];
            var filhasHtml = filhas.map(function (f) { return renderizarCardAnotacao(f, negocio.id); }).join('');
            return '<div class="crm-anot-thread">' +
                renderizarCardAnotacao(raiz, negocio.id) +
                (filhasHtml ? '<div class="crm-anot-filhos">' + filhasHtml + '</div>' : '') +
            '</div>';
        }).join('');

        return '' +
            '<div class="crm-atv-composer">' +
                '<button type="button" class="btn btn-primary" data-crm-action="novaAnotacaoNegocio" data-negocio-id="' + esc(negocio.id) + '">+ Nova Anotação</button>' +
            '</div>' +
            '<div class="crm-anot-notas-grid">' + (threads || '<p>Nenhuma anotação registrada para este negócio.</p>') + '</div>';
    }

    function renderizarFoco(negocio, atividades) {
        var pendentes = CrmCalculos.atividadesPendentesDe(atividades, negocio.id);
        if (!pendentes.length) {
            return '<div class="crm-bloco-titulo">Foco</div><p>Nenhum item de foco. Agende uma atividade acima.</p>';
        }
        var hoje = CrmCalculos.hojeIso();
        var mapaTipos = CrmStore.mapaTiposAtividade();
        var itens = pendentes.map(function (a) {
            var atrasada = a.data && a.data < hoje;
            var t = mapaTipos[a.tipo] || { icone: '', rotulo: a.tipo };
            return '' +
                '<div class="crm-foco-item' + (atrasada ? ' crm-foco-atrasada' : '') + '">' +
                    '<button type="button" class="crm-foco-check" data-crm-action="concluirAtividade" data-id="' + esc(a.id) + '" data-feito="' + a.feito + '">✓</button>' +
                    '<div class="crm-foco-corpo">' +
                        '<div class="crm-foco-assunto">' + t.icone + ' ' + esc(a.assunto) + '</div>' +
                        '<div class="crm-card-vinculos">' + (a.data ? esc(DateUtils.formatBR(a.data)) : '') + (a.horaInicio ? (' às ' + esc(a.horaInicio)) : '') + '</div>' +
                    '</div>' +
                    '<button type="button" class="btn-secondary crm-btn-mini" data-crm-action="editarAtividade" data-id="' + esc(a.id) + '">Editar</button>' +
                    '<button type="button" class="crm-chip-x" data-crm-action="excluirAtividade" data-id="' + esc(a.id) + '">✕</button>' +
                '</div>';
        }).join('');
        return '<div class="crm-bloco-titulo">Foco</div>' + itens;
    }

    function renderizarHistoricoFiltrado(negocio) {
        var todos = CrmStore.historicoDe('negocio', negocio.id);
        var filtros = [
            { valor: 'todos', rotulo: 'Todos' },
            { valor: 'atividades', rotulo: 'Atividades' },
            { valor: 'anotacoes', rotulo: 'Anotações' },
            { valor: 'alteracoes', rotulo: 'Alterações' }
        ];
        var mapaFiltro = { atividades: 'atividade', anotacoes: 'anotacao' };
        function pertenceAlteracoes(h) { return h.tipo === 'campo' || h.tipo === 'etapa' || h.tipo === 'criacao' || h.tipo === 'exclusao'; }

        var itensFiltrados = todos.filter(function (h) {
            if (_filtroHistorico === 'todos') return true;
            if (_filtroHistorico === 'alteracoes') return pertenceAlteracoes(h);
            return h.tipo === mapaFiltro[_filtroHistorico];
        });

        var pills = filtros.map(function (f) {
            var n = f.valor === 'todos' ? todos.length : todos.filter(function (h) {
                return f.valor === 'alteracoes' ? pertenceAlteracoes(h) : h.tipo === mapaFiltro[f.valor];
            }).length;
            return '<button type="button" class="crm-hist-pill' + (_filtroHistorico === f.valor ? ' active' : '') + '" data-crm-action="filtroHistorico" data-valor="' + f.valor + '">' + f.rotulo + ' (' + n + ')</button>';
        }).join('');

        var itensHtml = itensFiltrados.map(function (h) {
            var icone = ICONES_HISTORICO[h.tipo] || '•';
            return '<div class="crm-timeline-item"><span>' + icone + '</span> <span>' + esc(h.texto) + '</span> <span class="crm-card-vinculos">' + esc(DateUtils.formatBR(h.criadoEm)) + '</span></div>';
        }).join('');

        return '<div class="crm-bloco-titulo">Histórico</div><div class="crm-hist-pills">' + pills + '</div>' + (itensHtml || '<p>Nenhum registro.</p>');
    }

    // ──────────────────────────────────────────────
    //  ATIVIDADES / NOTAS
    // ──────────────────────────────────────────────

    function salvarAtividade(negocioId) {
        if (!requireAdminOrNotify()) return;
        var id = document.getElementById('crmAtvId').value || null;
        var dados = {
            negocioId: negocioId || _detalheId,
            tipo: document.getElementById('crmAtvTipo').value,
            assunto: document.getElementById('crmAtvAssunto').value.trim(),
            data: document.getElementById('crmAtvData').value || null,
            horaInicio: document.getElementById('crmAtvHoraInicio').value || '',
            horaFim: document.getElementById('crmAtvHoraFim').value || '',
            descricao: document.getElementById('crmAtvDescricao').value
        };
        var erros = CrmModel.validarAtividade(CrmModel.normalizarAtividade(dados));
        if (erros.length) { Notifications.error(erros[0]); return; }

        var criado = null;
        if (id) CrmStore.atualizarAtividade(id, dados);
        else criado = CrmStore.criarAtividade(dados);
        sincronizarAtividadeComGoogle(id || (criado && criado.id));

        _atividadeEditandoId = null;
        renderizarConteudoAtivo();
    }

    // ──────────────────────────────────────────────
    //  MODAL: NEGÓCIO (cliente, itens de produto, etapa, proposta)
    // ──────────────────────────────────────────────

    function popularSelectRepresentantesModal() {
        var sel = document.getElementById('crmNegocioResponsavel');
        var reps = CrmStore.listarRepresentantes();
        sel.innerHTML = '<option value="">Selecione...</option>' + reps.map(function (r) { return '<option value="' + esc(r) + '">' + esc(r) + '</option>'; }).join('');
    }

    function popularSelectProdutosModal() {
        var sel = document.getElementById('crmItemProdutoSelect');
        var produtos = CrmStore.listarProdutos();
        sel.innerHTML = '<option value="">Selecione um produto...</option>' + produtos.map(function (p) { return '<option value="' + esc(p.id) + '">' + esc(p.nome) + '</option>'; }).join('');
    }

    function popularSelectPropostasModal() {
        var sel = document.getElementById('crmNegocioProposta');
        var propostas = CrmStore.listarPropostas();
        sel.innerHTML = '<option value="">Nenhuma</option>' + propostas.map(function (p) {
            var rotulo = p.titulo || p.nome || ('Proposta #' + p.id);
            return '<option value="' + esc(p.id) + '">' + esc(rotulo) + '</option>';
        }).join('');
    }

    function renderizarChevronsModal(funil, etapaSelecionadaId) {
        var idx = -1;
        funil.etapas.forEach(function (e, i) { if (e.id === etapaSelecionadaId) idx = i; });
        var html = funil.etapas.map(function (e, i) {
            var cls = 'crm-chevron' + (i <= idx ? ' crm-chevron-ativo' : '');
            return '<button type="button" class="' + cls + '" title="' + esc(e.nome) + '" data-crm-action="setEtapaModal" data-etapa-id="' + esc(e.id) + '"></button>';
        }).join('');
        document.getElementById('crmNegocioEtapaChevrons').innerHTML = html;
    }

    function renderizarItensModal() {
        var lista = document.getElementById('crmItensLista');
        var linhas = _itensTemp.map(function (it, idx) {
            return '<div class="crm-item-linha"><span>' + esc(it.nome) + '</span>' +
                '<span>' + it.quantidade + ' × ' + esc(CrmCalculos.formatarMoeda(it.precoUnit)) + '</span>' +
                '<span>' + esc(CrmCalculos.formatarMoeda(it.quantidade * it.precoUnit)) + '</span>' +
                '<button type="button" class="crm-chip-x" data-crm-action="removerItem" data-idx="' + idx + '">✕</button></div>';
        }).join('');
        lista.innerHTML = linhas || '<div class="crm-coluna-vazia">Nenhum item adicionado.</div>';
        var total = CrmCalculos.somarItens(_itensTemp);
        document.getElementById('crmItensTotal').textContent = CrmCalculos.formatarMoeda(total);
        var valorWrap = document.getElementById('crmValorManualWrap');
        if (valorWrap) valorWrap.style.display = _itensTemp.length ? 'none' : '';
    }

    function adicionarItem() {
        var sel = document.getElementById('crmItemProdutoSelect');
        var produtoId = sel.value;
        if (!produtoId) { Notifications.error('Selecione um produto.'); return; }
        var produto = CrmStore.listarProdutos().filter(function (p) { return String(p.id) === String(produtoId); })[0];
        var qtd = Number(document.getElementById('crmItemQtd').value) || 1;
        var preco = Number(document.getElementById('crmItemPreco').value) || 0;
        _itensTemp.push({ produtoId: produtoId, nome: produto ? produto.nome : '', quantidade: qtd, precoUnit: preco });
        document.getElementById('crmItemQtd').value = 1;
        document.getElementById('crmItemPreco').value = '';
        sel.value = '';
        renderizarItensModal();
    }

    function buscarCliente(termo) {
        var lista = document.getElementById('crmClienteLista');
        var todos = CrmStore.listarClientes();
        var termoNorm = String(termo || '').trim().toLowerCase();
        var filtrados = (termoNorm ? todos.filter(function (c) { return (c.nome || '').toLowerCase().indexOf(termoNorm) !== -1; }) : todos).slice(0, 8);

        var itens = filtrados.map(function (c) {
            return '<div class="crm-ac-item" data-crm-action="selecionarCliente" data-id="' + esc(c.id) + '" data-nome="' + esc(c.nome) + '">' +
                esc(c.nome) + (c.cnpj ? (' <span style="opacity:.6">· ' + esc(c.cnpj) + '</span>') : '') + '</div>';
        }).join('');

        var temExato = todos.some(function (c) { return (c.nome || '').trim().toLowerCase() === termoNorm; });
        var criarNovo = (termoNorm && !temExato)
            ? '<div class="crm-ac-item crm-ac-novo" data-crm-action="criarClienteInline" data-nome="' + esc(termo.trim()) + '">+ Criar cliente "' + esc(termo.trim()) + '"</div>'
            : '';

        lista.innerHTML = (itens + criarNovo) || '<div class="crm-ac-vazio">Nenhum cliente encontrado.</div>';
        lista.style.display = 'block';
    }

    // ── Cadastro de cliente aberto de dentro do CRM ──

    /**
     * De onde o cadastro de cliente foi aberto, para saber o que atualizar
     * quando ele for salvo. Null quando a edição veio da própria aba Clientes.
     */
    var _clienteEditadoDe = null;

    /**
     * Abre o cadastro de cliente (modal do app2) empilhado sobre a tela atual,
     * sem trocar de aba — assim nada do que estiver sendo preenchido se perde.
     * O retorno é automático: ao salvar, o modal fecha e a tela de baixo continua
     * onde estava, já atualizada (ver aoSalvarClienteExterno).
     */
    function abrirCadastroCliente(clienteId, origem) {
        if (typeof abrirModalCliente !== 'function') {
            Notifications.error('Cadastro de clientes indisponível nesta tela.');
            return;
        }
        _clienteEditadoDe = origem || null;
        abrirModalCliente(clienteId || null);
        // Todos os modais compartilham z-index 3100; sem isto o modal de Negócio,
        // que vem depois no DOM, ficaria por cima do cadastro de cliente.
        var m = document.getElementById('modalCliente');
        if (m) m.classList.add('crm-modal-empilhado');
    }

    function aoSalvarClienteExterno(ev) {
        if (!_clienteEditadoDe) return; // edição feita pela aba Clientes: não é conosco
        var origem = _clienteEditadoDe;
        _clienteEditadoDe = null;

        var m = document.getElementById('modalCliente');
        if (m) m.classList.remove('crm-modal-empilhado');

        var id = (ev && ev.detail) ? ev.detail.id : null;
        var modalNegocio = document.getElementById('modalNegocio');
        var negocioAberto = modalNegocio && modalNegocio.style.display !== 'none' && modalNegocio.style.display !== '';
        if (origem === 'negocio' && id && negocioAberto) {
            var c = CrmStore.getCliente(id);
            if (c) selecionarCliente(c.id, c.nome);
        }
        renderizarConteudoAtivo();
    }

    function editarClienteDoNegocio() {
        var id = document.getElementById('crmNegocioClienteId').value;
        if (!id) { Notifications.error('Selecione um cliente primeiro.'); return; }
        abrirCadastroCliente(id, 'negocio');
    }

    function selecionarCliente(id, nome) {
        document.getElementById('crmNegocioClienteId').value = id;
        document.getElementById('crmNegocioClienteBusca').value = nome;
        document.getElementById('crmClienteLista').style.display = 'none';
        var btn = document.getElementById('crmNegocioBtnEditarCliente');
        if (btn) btn.disabled = !id;
    }

    function abrirModalNegocio(id) {
        var funil = CrmStore.getFunilAtivo();
        if (!funil) { Notifications.error('Nenhum funil ativo. Crie um funil primeiro.'); return; }

        document.getElementById('crmModalNegocioTitulo').textContent = id ? 'Editar negócio' : 'Novo negócio';
        document.getElementById('crmNegocioId').value = id || '';
        popularSelectRepresentantesModal();
        popularSelectProdutosModal();
        popularSelectPropostasModal();

        var crm = CrmStore.getCrm();
        var negocio = id ? crm.negocios.filter(function (n) { return n.id === id; })[0] : null;

        document.getElementById('crmNegocioTitulo').value = negocio ? (negocio.titulo || '') : '';
        document.getElementById('crmNegocioValor').value = (negocio && (!negocio.itens || !negocio.itens.length) && negocio.valor !== null) ? negocio.valor : '';
        document.getElementById('crmNegocioTags').value = negocio ? (negocio.tags || []).join(', ') : '';
        document.getElementById('crmNegocioDescricao').value = negocio ? (negocio.descricao || '') : '';
        document.getElementById('crmNegocioResponsavel').value = negocio ? (negocio.responsavel || '') : '';
        document.getElementById('crmNegocioProposta').value = negocio ? (negocio.propostaId || '') : '';
        document.getElementById('crmNegocioPrevisao').value = negocio ? (negocio.dataPrevisao || '') : '';
        document.getElementById('crmNegocioRecebimento').value = negocio ? (negocio.dataRecebimento || '') : '';
        document.getElementById('crmNegocioOrigem').value = negocio ? (negocio.origem || '') : '';

        var clienteNome = '';
        if (negocio && negocio.clienteId) {
            var c = CrmStore.getCliente(negocio.clienteId);
            clienteNome = c ? c.nome : '';
        }
        document.getElementById('crmNegocioClienteId').value = negocio ? (negocio.clienteId || '') : '';
        document.getElementById('crmNegocioClienteBusca').value = clienteNome;
        document.getElementById('crmClienteLista').style.display = 'none';
        var btnCli = document.getElementById('crmNegocioBtnEditarCliente');
        if (btnCli) btnCli.disabled = !(negocio && negocio.clienteId);

        _itensTemp = (negocio && negocio.itens) ? negocio.itens.map(function (it) { return Object.assign({}, it); }) : [];
        renderizarItensModal();

        var etapaId = negocio ? negocio.etapaId : ((funil.etapas.filter(function (e) { return e.tipo === 'aberta'; })[0] || funil.etapas[0] || {}).id);
        document.getElementById('crmNegocioEtapaId').value = etapaId || '';
        renderizarChevronsModal(funil, etapaId);

        document.getElementById('modalNegocio').style.display = 'flex';
    }

    function salvarNegocio() {
        if (!requireAdminOrNotify()) return;
        var funil = CrmStore.getFunilAtivo();
        if (!funil) return;

        var id = document.getElementById('crmNegocioId').value || null;
        var tagsRaw = document.getElementById('crmNegocioTags').value || '';
        var dados = {
            funilId: funil.id,
            etapaId: document.getElementById('crmNegocioEtapaId').value || null,
            titulo: document.getElementById('crmNegocioTitulo').value.trim(),
            clienteId: document.getElementById('crmNegocioClienteId').value || null,
            itens: _itensTemp,
            valor: _itensTemp.length ? null : (document.getElementById('crmNegocioValor').value || null),
            responsavel: document.getElementById('crmNegocioResponsavel').value || '',
            propostaId: document.getElementById('crmNegocioProposta').value || null,
            dataPrevisao: document.getElementById('crmNegocioPrevisao').value || null,
            dataRecebimento: document.getElementById('crmNegocioRecebimento').value || null,
            origem: document.getElementById('crmNegocioOrigem').value.trim(),
            tags: tagsRaw.split(',').map(function (t) { return t.trim(); }).filter(Boolean),
            descricao: document.getElementById('crmNegocioDescricao').value
        };

        var erros = CrmModel.validarNegocio(CrmModel.normalizarNegocio(dados), funil);
        if (erros.length) { Notifications.error(erros[0]); return; }

        if (id) CrmStore.atualizarNegocio(id, dados);
        else CrmStore.criarNegocio(dados);

        fecharModal('modalNegocio');
        Notifications.success('Negócio salvo com sucesso.');
        renderizarConteudoAtivo();
    }

    // ──────────────────────────────────────────────
    //  MODAL: TIPOS DE ATIVIDADE (personalizáveis)
    // ──────────────────────────────────────────────

    var ICONES_TIPO_ATIVIDADE = [
        '📞', '👥', '✔️', '🚩', '✉️', '🧳', '🏖️', '📴', '👤', '⏰',
        '📸', '✂️', '🔧', '🔒', '📋', '📱', '🎯', '💰', '📊', '🚗',
        '🌐', '🔍', '🔊', '🔑', '📎', '💼', '🏆', '🚀', '📡', '🎓',
        '🍽️', '🎉', '🩺', '⚖️', '📚', '🖥️'
    ];

    function abrirModalTiposAtividade() {
        renderizarListaTiposAtividadeModal();
        document.getElementById('modalTiposAtividade').style.display = 'flex';
    }

    function renderizarListaTiposAtividadeModal() {
        var lista = document.getElementById('crmTiposAtividadeLista');
        if (!lista) return;
        var tipos = CrmStore.listarTiposAtividade();
        lista.innerHTML = tipos.map(function (t, idx) {
            return '<div class="crm-tipo-atv-linha' + (!t.ativo ? ' crm-tipo-atv-inativo' : '') + '">' +
                '<div class="crm-tipo-atv-setas">' +
                    '<button type="button" class="crm-tipo-atv-seta" data-crm-action="moverTipoAtividade" data-chave="' + esc(t.chave) + '" data-dir="-1"' + (idx === 0 ? ' disabled' : '') + '>▲</button>' +
                    '<button type="button" class="crm-tipo-atv-seta" data-crm-action="moverTipoAtividade" data-chave="' + esc(t.chave) + '" data-dir="1"' + (idx === tipos.length - 1 ? ' disabled' : '') + '>▼</button>' +
                '</div>' +
                '<span class="crm-tipo-atv-icone">' + t.icone + '</span>' +
                '<span class="crm-tipo-atv-nome" data-crm-action="abrirFormTipoAtividade" data-chave="' + esc(t.chave) + '">' + esc(t.nome) + '</span>' +
                '<label class="crm-switch">' +
                    '<input type="checkbox"' + (t.ativo ? ' checked' : '') + ' data-crm-action="toggleAtivoTipoAtividade" data-chave="' + esc(t.chave) + '">' +
                    '<span class="crm-switch-track"></span>' +
                '</label>' +
            '</div>';
        }).join('');
    }

    function moverTipoAtividade(chave, dir) {
        if (!requireAdminOrNotify()) return;
        var tipos = CrmStore.listarTiposAtividade();
        var idx = -1;
        tipos.forEach(function (t, i) { if (t.chave === chave) idx = i; });
        var novoIdx = idx + dir;
        if (idx === -1 || novoIdx < 0 || novoIdx >= tipos.length) return;
        var chaves = tipos.map(function (t) { return t.chave; });
        var tmp = chaves[idx];
        chaves[idx] = chaves[novoIdx];
        chaves[novoIdx] = tmp;
        CrmStore.reordenarTiposAtividade(chaves);
        renderizarListaTiposAtividadeModal();
        renderizarConteudoAtivo();
    }

    function toggleAtivoTipoAtividade(chave, ativo) {
        if (!requireAdminOrNotify()) return;
        CrmStore.atualizarTipoAtividade(chave, { ativo: ativo });
        renderizarListaTiposAtividadeModal();
        renderizarConteudoAtivo();
    }

    function renderizarGradeIconesTipoAtividade(selecionado) {
        var grid = document.getElementById('crmTipoAtvIconeGrid');
        if (!grid) return;
        grid.innerHTML = ICONES_TIPO_ATIVIDADE.map(function (ic) {
            return '<button type="button" class="crm-tipo-icone-opcao' + (ic === selecionado ? ' active' : '') + '" data-crm-action="escolherIconeTipoAtividade" data-valor="' + ic + '">' + ic + '</button>';
        }).join('');
    }

    function escolherIconeTipoAtividade(icone) {
        document.getElementById('crmTipoAtvIconeEscolhido').value = icone;
        renderizarGradeIconesTipoAtividade(icone);
    }

    function abrirFormTipoAtividade(chave) {
        var t = chave ? CrmStore.listarTiposAtividade().filter(function (x) { return x.chave === chave; })[0] : null;
        document.getElementById('crmTipoAtvFormTitulo').textContent = t ? 'Editar tipo de atividade' : 'Novo tipo de atividade';
        document.getElementById('crmTipoAtvChave').value = t ? t.chave : '';
        document.getElementById('crmTipoAtvNome').value = t ? t.nome : '';
        var iconeEscolhido = t ? t.icone : ICONES_TIPO_ATIVIDADE[0];
        document.getElementById('crmTipoAtvIconeEscolhido').value = iconeEscolhido;
        renderizarGradeIconesTipoAtividade(iconeEscolhido);
        document.getElementById('modalTipoAtividadeForm').style.display = 'flex';
    }

    function salvarTipoAtividade() {
        if (!requireAdminOrNotify()) return;
        var nome = document.getElementById('crmTipoAtvNome').value.trim();
        if (!nome) { Notifications.error('Nome é obrigatório.'); return; }
        var icone = document.getElementById('crmTipoAtvIconeEscolhido').value || '📌';
        var chave = document.getElementById('crmTipoAtvChave').value;
        if (chave) {
            CrmStore.atualizarTipoAtividade(chave, { nome: nome, icone: icone });
        } else {
            CrmStore.criarTipoAtividade({ nome: nome, icone: icone });
        }
        fecharModal('modalTipoAtividadeForm');
        renderizarListaTiposAtividadeModal();
        renderizarConteudoAtivo();
    }

    // ──────────────────────────────────────────────
    //  MODAL: FUNIL
    // ──────────────────────────────────────────────

    function renderizarEtapasFunilModal(etapas) {
        var lista = document.getElementById('crmFunilEtapasLista');
        lista.innerHTML = etapas.map(function (e) {
            return '' +
                '<div class="crm-funil-etapa-linha">' +
                    '<input type="text" value="' + esc(e.nome) + '" data-crm-etapa-nome data-id="' + esc(e.id) + '">' +
                    '<select data-crm-etapa-tipo>' +
                        '<option value="aberta"' + (e.tipo === 'aberta' ? ' selected' : '') + '>Aberta</option>' +
                        '<option value="ganho"' + (e.tipo === 'ganho' ? ' selected' : '') + '>Ganho</option>' +
                        '<option value="perdido"' + (e.tipo === 'perdido' ? ' selected' : '') + '>Perdido</option>' +
                    '</select>' +
                    '<button type="button" class="crm-chip-x" onclick="this.closest(\'.crm-funil-etapa-linha\').remove()">✕</button>' +
                '</div>';
        }).join('');
    }

    function abrirModalFunil() {
        var funil = CrmStore.getFunilAtivo();
        if (!funil) return;
        document.getElementById('crmFunilNome').value = funil.nome;
        renderizarEtapasFunilModal(funil.etapas);
        document.getElementById('modalFunil').style.display = 'flex';
    }

    function adicionarEtapaFunil() {
        var lista = document.getElementById('crmFunilEtapasLista');
        var div = document.createElement('div');
        div.className = 'crm-funil-etapa-linha';
        div.innerHTML = '<input type="text" value="" placeholder="Nova etapa" data-crm-etapa-nome data-id="">' +
            '<select data-crm-etapa-tipo><option value="aberta" selected>Aberta</option><option value="ganho">Ganho</option><option value="perdido">Perdido</option></select>' +
            '<button type="button" class="crm-chip-x" onclick="this.closest(\'.crm-funil-etapa-linha\').remove()">✕</button>';
        lista.appendChild(div);
    }

    function salvarFunil() {
        if (!requireAdminOrNotify()) return;
        var funil = CrmStore.getFunilAtivo();
        if (!funil) return;

        var nome = document.getElementById('crmFunilNome').value.trim();
        if (nome) CrmStore.atualizarFunil(funil.id, { nome: nome });

        var linhas = document.querySelectorAll('#crmFunilEtapasLista .crm-funil-etapa-linha');
        var etapas = Array.prototype.map.call(linhas, function (linha, idx) {
            var inputNome = linha.querySelector('[data-crm-etapa-nome]');
            var selectTipo = linha.querySelector('[data-crm-etapa-tipo]');
            return {
                id: inputNome.dataset.id || undefined,
                nome: inputNome.value.trim() || ('Etapa ' + (idx + 1)),
                ordem: idx,
                tipo: selectTipo.value
            };
        });
        if (!etapas.length) { Notifications.error('O funil precisa de ao menos uma etapa.'); return; }

        CrmStore.definirEtapasFunil(funil.id, etapas);
        fecharModal('modalFunil');
        Notifications.success('Funil atualizado.');
        popularSelectFunil();
        renderizarConteudoAtivo();
    }

    function criarFunilDeTemplate() {
        if (!requireAdminOrNotify()) return;
        var chave = document.getElementById('crmFunilNovoTemplate').value;
        var funil = CrmStore.criarFunil({ template: chave });
        if (!funil) { Notifications.error('Modelo inválido.'); return; }
        CrmStore.setFunilAtivo(funil.id);
        popularSelectFunil();
        abrirModalFunil();
        Notifications.success('Funil criado a partir do modelo.');
    }

    // ──────────────────────────────────────────────
    //  DELEGAÇÃO DE CLIQUES (data-crm-action)
    // ──────────────────────────────────────────────

    function confirmarEExecutar(msg, cb) { Notifications.confirm(msg, cb); }

    var ACOES = {
        setSecao: function (el) { setSecao(el.dataset.valor); },
        setVisao: function (el) { _visao = el.dataset.valor; _detalheId = null; renderizarConteudoAtivo(); },
        abrirDetalhe: function (el) { _detalheId = el.dataset.id; _abaDetalhe = 'atividade'; _atividadeEditandoId = null; renderizarConteudoAtivo(); },
        voltarLista: function () { _detalheId = null; renderizarConteudoAtivo(); },

        excluirNegocio: function (el) {
            if (!requireAdminOrNotify()) return;
            confirmarEExecutar('Excluir este negócio? Ele irá para a lixeira.', function () {
                CrmStore.removerNegocio(el.dataset.id);
                _detalheId = null;
                renderizarConteudoAtivo();
            });
        },
        restaurarNegocio: function (el) {
            if (!requireAdminOrNotify()) return;
            CrmStore.restaurarNegocio(el.dataset.id);
            renderizarConteudoAtivo();
        },
        excluirDefinitivo: function (el) {
            if (!requireAdminOrNotify()) return;
            confirmarEExecutar('Excluir definitivamente? Esta ação não pode ser desfeita.', function () {
                CrmStore.excluirNegocioDefinitivo(el.dataset.id);
                renderizarConteudoAtivo();
            });
        },
        editarNegocio: function (el) { abrirModalNegocio(el.dataset.id); },
        marcarGanho: function (el) {
            if (!requireAdminOrNotify()) return;
            CrmStore.marcarGanho(el.dataset.id);
            renderizarConteudoAtivo();
        },
        marcarPerdido: function (el) {
            if (!requireAdminOrNotify()) return;
            var motivo = prompt('Motivo da perda (opcional):') || '';
            CrmStore.marcarPerdido(el.dataset.id, motivo);
            renderizarConteudoAtivo();
        },
        moverEtapaProgresso: function (el) {
            if (!requireAdminOrNotify()) return;
            CrmStore.moverNegocio(_detalheId, el.dataset.etapaId, null);
            renderizarConteudoAtivo();
        },

        toggleSecao: function (el) {
            var chave = el.dataset.secao;
            _secoesColapsadas[chave] = !_secoesColapsadas[chave];
            renderizarConteudoAtivo();
        },
        trocarAbaDetalhe: function (el) { _abaDetalhe = el.dataset.valor; _atividadeEditandoId = null; renderizarConteudoAtivo(); },
        trocarFunilNegocio: function (el) { abrirSeletorFunilNegocio(el.dataset.negocioId); },
        filtroHistorico: function (el) { _filtroHistorico = el.dataset.valor; renderizarConteudoAtivo(); },

        escolherTipoAtividade: function (el) {
            document.getElementById('crmAtvTipo').value = el.dataset.valor;
            document.querySelectorAll('.crm-atv-tipo').forEach(function (btn) {
                btn.classList.toggle('active', btn.dataset.valor === el.dataset.valor);
            });
        },
        salvarAtividade: function (el) { salvarAtividade(el.dataset.negocioId); },
        cancelarEdicaoAtividade: function () { _atividadeEditandoId = null; renderizarConteudoAtivo(); },
        editarAtividade: function (el) { _atividadeEditandoId = el.dataset.id; _abaDetalhe = 'atividade'; renderizarConteudoAtivo(); },
        concluirAtividade: function (el) {
            if (!requireAdminOrNotify()) return;
            CrmStore.concluirAtividade(el.dataset.id, el.dataset.feito !== 'true');
            sincronizarAtividadeComGoogle(el.dataset.id);
            renderizarConteudoAtivo();
        },
        excluirAtividade: function (el) {
            if (!requireAdminOrNotify()) return;
            confirmarEExecutar('Excluir esta atividade?', function () {
                var atividade = CrmStore.listarAtividades().filter(function (a) { return a.id === el.dataset.id; })[0];
                CrmStore.removerAtividade(el.dataset.id);
                removerAtividadeDoGoogle(atividade);
                renderizarConteudoAtivo();
            });
        },

        removerItem: function (el) { _itensTemp.splice(Number(el.dataset.idx), 1); renderizarItensModal(); },
        selecionarCliente: function (el) { selecionarCliente(el.dataset.id, el.dataset.nome); },
        editarCliente: function (el) { abrirCadastroCliente(el.dataset.id, el.dataset.origem); },
        criarClienteInline: function (el) {
            if (!requireAdminOrNotify()) return;
            var c = CrmStore.criarCliente({ nome: el.dataset.nome });
            if (c) selecionarCliente(c.id, c.nome);
        },
        setEtapaModal: function (el) {
            document.getElementById('crmNegocioEtapaId').value = el.dataset.etapaId;
            renderizarChevronsModal(CrmStore.getFunilAtivo(), el.dataset.etapaId);
        },

        abrirModalAnotacao: function (el) { abrirModalAnotacao(el.dataset.id); },
        novaAnotacaoNegocio: function (el) { abrirModalAnotacao(null, el.dataset.negocioId, el.dataset.relacionadaId || null); },
        selecionarNegocioAnotacao: function (el) { selecionarNegocioAnotacao(el.dataset.id); },
        setDemandaAba: function (el) { setDemandaAba(el.dataset.valor); },
        removerEncTemp: function (el) { removerEncTemp(Number(el.dataset.idx)); },
        marcarEncRespondido: function (el) { marcarEncRespondido(Number(el.dataset.idx)); },

        setCalModo: function (el) { setCalModo(el.dataset.valor); },
        setCalPeriodo: function (el) { setCalPeriodo(el.dataset.valor); },
        setCalTipo: function (el) { setCalTipo(el.dataset.valor); },
        calHoje: function () { calHoje(); },
        calSemanaAnterior: function () { calSemanaAnterior(); },
        calSemanaProxima: function () { calSemanaProxima(); },
        calMesHoje: function () { calMesHoje(); },
        calMesAnterior: function () { calMesAnterior(); },
        calMesProximo: function () { calMesProximo(); },
        concluirAtividadeCal: function (el) {
            if (!requireAdminOrNotify()) return;
            CrmStore.concluirAtividade(el.dataset.id, el.dataset.feito !== 'true');
            sincronizarAtividadeComGoogle(el.dataset.id);
            renderizarCalendarioView();
        },
        abrirModalAtividadeCal: function (el) { abrirModalAtividadeCal(el.dataset.id || null); },
        novaAtividadeNoDia: function (el) { abrirModalAtividadeCal(null, el.dataset.data); },
        escolherTipoAtividadeCal: function (el) { escolherTipoAtividadeCal(el); },
        selecionarNegocioAtividadeCal: function (el) { selecionarNegocioAtividadeCal(el.dataset.id); },

        googleConectar: function () { if (window.GoogleCalendarSync) GoogleCalendarSync.conectar(); },
        googleDesconectar: function () { if (window.GoogleCalendarSync) GoogleCalendarSync.desconectar(); },
        googleSincronizarTudo: function () { if (window.GoogleCalendarSync) GoogleCalendarSync.sincronizarTudo(); },

        pontoConectar: function () {
            if (!window.PontoBridge) return;
            var email = (document.getElementById('crmPontoEmail') || {}).value || '';
            var senha = (document.getElementById('crmPontoSenha') || {}).value || '';
            PontoBridge.conectar(email.trim(), senha).then(function () {
                _pontoAutoFetchFeito = true; // já vamos buscar abaixo; evita disparo duplicado
                if (window.Notifications) Notifications.success('Conectado ao Ponto.');
                return PontoBridge.atualizarEventos(true);
            }).then(function () {
                renderizarCalendarioView();
            }).catch(function (e) {
                if (window.Notifications) Notifications.error('Falha ao conectar ao Ponto: ' + (e && e.message ? e.message : e));
            });
        },
        pontoDesconectar: function () { if (window.PontoBridge) PontoBridge.desconectar(); },
        pontoAtualizar: function () {
            if (!window.PontoBridge) return;
            PontoBridge.atualizarEventos(true).then(function () {
                if (window.Notifications) Notifications.success('Eventos do Ponto atualizados.');
                renderizarCalendarioView();
            });
        },

        abrirModalTiposAtividade: function () { abrirModalTiposAtividade(); },
        moverTipoAtividade: function (el) { moverTipoAtividade(el.dataset.chave, Number(el.dataset.dir)); },
        toggleAtivoTipoAtividade: function (el) { toggleAtivoTipoAtividade(el.dataset.chave, el.checked); },
        abrirFormTipoAtividade: function (el) { abrirFormTipoAtividade(el && el.dataset ? el.dataset.chave : null); },
        escolherIconeTipoAtividade: function (el) { escolherIconeTipoAtividade(el.dataset.valor); }
    };

    function aoClicar(e) {
        var el = e.target.closest('[data-crm-action]');
        if (!el) {
            if (!e.target.closest('.crm-autocomplete')) {
                var lista = document.getElementById('crmClienteLista');
                if (lista) lista.style.display = 'none';
                var listaNeg = document.getElementById('crmAnotacaoNegocioLista');
                if (listaNeg) listaNeg.style.display = 'none';
                var listaAtvCal = document.getElementById('crmAtvCalNegocioLista');
                if (listaAtvCal) listaAtvCal.style.display = 'none';
            }
            return;
        }
        var fn = ACOES[el.dataset.crmAction];
        if (fn) fn(el);
    }

    // ──────────────────────────────────────────────
    //  EXPORT
    // ──────────────────────────────────────────────

    window.Crm = {
        renderizar: renderizar,
        trocarFunil: trocarFunil,
        setBusca: setBusca,
        setMostrarFechados: setMostrarFechados,
        setOrdenarPor: setOrdenarPor,

        setListaBusca: setListaBusca,
        setListaFiltro: setListaFiltro,
        limparListaFiltros: limparListaFiltros,
        setOrdenacaoLista: setOrdenacaoLista,
        setListaPagina: setListaPagina,

        abrirModalNegocio: abrirModalNegocio,
        salvarNegocio: salvarNegocio,
        buscarCliente: buscarCliente,
        abrirCadastroCliente: abrirCadastroCliente,
        editarClienteDoNegocio: editarClienteDoNegocio,
        adicionarItem: adicionarItem,

        abrirModalAnotacao: abrirModalAnotacao,
        salvarAnotacao: salvarAnotacao,
        excluirAnotacao: excluirAnotacao,
        buscarNegocioParaAnotacao: buscarNegocioParaAnotacao,
        desvincularNegocioAnotacao: desvincularNegocioAnotacao,
        aoTrocarFunilAnotacao: aoTrocarFunilAnotacao,
        adicionarEncaminhamento: adicionarEncaminhamento,
        consolidarRespostas: consolidarRespostas,

        abrirModalFunil: abrirModalFunil,
        adicionarEtapaFunil: adicionarEtapaFunil,
        criarFunilDeTemplate: criarFunilDeTemplate,
        salvarFunil: salvarFunil,

        setCalBusca: setCalBusca,
        setMostrarFeriadosCal: setMostrarFeriadosCal,
        setMostrarPontoCal: setMostrarPontoCal,
        abrirFormTipoAtividade: abrirFormTipoAtividade,
        salvarTipoAtividade: salvarTipoAtividade,
        abrirModalAtividadeCal: abrirModalAtividadeCal,
        salvarAtividadeCal: salvarAtividadeCal,
        excluirAtividadeCal: excluirAtividadeCal,
        buscarNegocioParaAtividadeCal: buscarNegocioParaAtividadeCal,
        atualizarMiniAgendaCal: atualizarMiniAgendaCal,

        salvarTrocaFunilNegocio: salvarTrocaFunilNegocio
    };
})();
