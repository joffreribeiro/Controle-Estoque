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
        ordenarPor: 'prazo',
        ordenarDireccao: 'asc'
    };
    var _crmListaPagina = 1;
    var _crmListaPaginaSize = 50;
    var _anotacaoNegocioId = null; // negócio selecionado no modal de anotação
    var _anotacaoRelacionadaId = null; // anotação-pai (thread) selecionada no modal

    // ── Estado da vista Calendário (Atividades, todos os negócios) ──
    var _calModo = 'lista'; // 'lista' | 'semana' | 'mes'
    var _calPeriodo = 'todos'; // todos|paraFazer|vencido|hoje|amanha|semana|proximaSemana
    var _calTipo = ''; // '' (todos) | chamada|reuniao|tarefa|prazo|email|viagem|ferias|recesso|particular
    var _calBusca = '';
    var _calSemanaRef = CrmCalculos.inicioSemana(hojeIsoLocal()); // domingo da semana exibida na vista Calendário
    var _calMesRef = hojeIsoLocal().slice(0, 7) + '-01'; // 1º dia do mês exibido na vista Mês
    var _calMostrarFeriados = false; // filtro da Lista: mostrar/ocultar feriados nacionais
    var _calAtvNegocioId = null; // negócio selecionado no modal de atividade global

    var ICONES_HISTORICO = { criacao: '✨', campo: '✎', etapa: '➡️', exclusao: '🗑', atividade: '📅', nota: '📝', anotacao: '🗒️' };
    var PRIORIDADE_ROTULO = { baixa: 'Baixa', media: 'Média', alta: 'Alta', critico: 'Crítico' };
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
        if (window.GoogleCalendarSync) {
            GoogleCalendarSync.aoMudarStatus(function () { if (_secao === 'calendario') renderizarCalendarioView(); });
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
        var previsao = document.getElementById('crmPrevisao');
        var excluidos = document.getElementById('crmExcluidos');
        var detalhe = document.getElementById('crmViewDetalhe');
        var barra = document.querySelector('.crm-barra-visoes');
        if (!kanban) return;

        if (_detalheId) {
            [kanban, lista, previsao, excluidos].forEach(function (el) { if (el) el.style.display = 'none'; });
            if (barra) barra.style.display = 'none';
            detalhe.style.display = 'block';
            renderizarDetalhe(_detalheId);
            return;
        }

        if (barra) barra.style.display = '';
        detalhe.style.display = 'none';
        kanban.style.display = _visao === 'kanban' ? '' : 'none';
        lista.style.display = _visao === 'lista' ? '' : 'none';
        previsao.style.display = _visao === 'previsao' ? '' : 'none';
        excluidos.style.display = _visao === 'excluidos' ? '' : 'none';

        var crmListaFiltros = document.getElementById('crmListaFiltros');
        if (crmListaFiltros) crmListaFiltros.style.display = _visao === 'lista' ? '' : 'none';

        document.querySelectorAll('.crm-visao-btn').forEach(function (b) {
            b.classList.toggle('active', b.dataset.valor === _visao);
        });

        var funil = CrmStore.getFunilAtivo();
        if (!funil) return;

        if (_visao === 'kanban') renderizarKanban(funil);
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

    // ── Lista (Anotações) ──

    function renderizarListaView() {
        popularFiltrosLista();

        var todasAnotacoes = CrmStore.listarAnotacoes();
        var negocios = CrmStore.listarNegocios();
        var clientes = CrmStore.listarClientes();
        var funis = CrmStore.listarFunis();
        var extra = { negocios: negocios, clientes: clientes, funis: funis };

        var negocioPorId = {};
        negocios.forEach(function (n) { negocioPorId[n.id] = n; });
        var clientePorId = {};
        clientes.forEach(function (c) { clientePorId[c.id] = c; });
        var funilPorId = {};
        funis.forEach(function (f) { funilPorId[f.id] = f; });
        var anotacaoPorId = {};
        todasAnotacoes.forEach(function (a) { anotacaoPorId[a.id] = a; });

        var filtradas = CrmCalculos.filtrarAnotacoes(todasAnotacoes, _crmListaFiltros, extra);
        var ordenadas = CrmCalculos.ordenarAnotacoes(filtradas, _crmListaFiltros.ordenarPor, _crmListaFiltros.ordenarDireccao, extra);

        var inicio = (_crmListaPagina - 1) * _crmListaPaginaSize;
        var fim = inicio + _crmListaPaginaSize;
        var pagina = ordenadas.slice(inicio, fim);

        var linhas = pagina.map(function (a) {
            var negocio = negocioPorId[a.negocioId];
            var funilObj = negocio ? funilPorId[negocio.funilId] : null;
            var clienteObj = negocio ? clientePorId[negocio.clienteId] : null;
            var vencida = (!a.finalizado && a.prazo && a.prazo < hojeIsoLocal()) ? ' class="crm-lista-vencida"' : '';
            var finalizadoIcone = a.finalizado ? '✓' : '✗';
            var prioridadeBadge = '<span class="crm-prioridade-badge" data-prioridade="' + esc(a.prioridade) + '">' +
                esc(PRIORIDADE_ROTULO[a.prioridade] || a.prioridade) + '</span>';
            var tags = (a.tags || []).map(function (t) { return '<span class="crm-tag">' + esc(t) + '</span>'; }).join(' ');
            var pai = a.anotacaoRelacionadaId ? anotacaoPorId[a.anotacaoRelacionadaId] : null;
            var assuntoCel = (pai ? '<span class="crm-anot-link-icone" title="Vinculada a: ' + esc(pai.assunto || '(sem assunto)') + '">🗒️</span> ' : '') + esc(a.assunto || '-');

            return '<tr' + vencida + ' data-crm-action="abrirModalAnotacao" data-id="' + esc(a.id) + '">' +
                '<td>' + esc(funilObj ? funilObj.nome : '-') + '</td>' +
                '<td>' + esc(clienteObj ? clienteObj.nome : '-') + '</td>' +
                '<td>' + assuntoCel + '</td>' +
                '<td>' + esc(a.remetente || '-') + '</td>' +
                '<td>' + esc(a.origemDemanda || '-') + '</td>' +
                '<td>' + esc(dataOuTraco(a.dataSolicitacao)) + '</td>' +
                '<td>' + esc(a.tipoDoc || '-') + '</td>' +
                '<td>' + esc(a.numeroDocumento || '-') + '</td>' +
                '<td>' + esc(a.destinatario || '-') + '</td>' +
                '<td>' + esc(a.acaoRealizar || '-') + '</td>' +
                '<td>' + prioridadeBadge + '</td>' +
                '<td>' + esc(dataOuTraco(a.prazo)) + '</td>' +
                '<td>' + esc(a.lembrarDiasAntes || '-') + '</td>' +
                '<td>' + (tags || '-') + '</td>' +
                '<td>' + esc(a.observacoes || '-') + '</td>' +
                '<td>' + finalizadoIcone + '</td>' +
                '<td>' + esc(dataOuTraco(a.dataConclusao)) + '</td>' +
                '<td>' + esc(a.oQueFoiFeito || '-') + '</td>' +
            '</tr>';
        }).join('');

        var totalPaginas = Math.max(1, Math.ceil(ordenadas.length / _crmListaPaginaSize));
        var colunas = [
            ['funil', 'Relacionamento'], ['cliente', 'Cliente'], ['assunto', 'Assunto'], ['remetente', 'Remetente'],
            ['origemDemanda', 'Origem da Demanda'], ['dataSolicitacao', 'Data Solicitação'], ['tipoDoc', 'Tipo Doc'],
            ['numeroDocumento', 'Nº Documento'], ['destinatario', 'Destinatário'], ['acaoRealizar', 'Ação a Realizar'],
            ['prioridade', 'Prioridade'], ['prazo', 'Prazo'], [null, 'Lembra (dias antes)'], [null, 'Tags'],
            [null, 'Observações'], ['finalizado', 'Finalizado'], ['dataConclusao', 'Data Conclusão'], [null, 'O que foi feito']
        ];
        var thead = colunas.map(function (c) {
            if (!c[0]) return '<th>' + esc(c[1]) + '</th>';
            var seta = _crmListaFiltros.ordenarPor === c[0] ? (_crmListaFiltros.ordenarDireccao === 'asc' ? ' ▲' : ' ▼') : '';
            return '<th onclick="Crm.setOrdenacaoLista(\'' + c[0] + '\')">' + esc(c[1]) + seta + '</th>';
        }).join('');
        var html = '<div class="crm-lista-topo"><button type="button" class="crm-btn-negocio" onclick="Crm.abrirModalAnotacao()">+ Nova Anotação</button></div>' +
            '<div class="crm-lista-wrapper"><table class="crm-lista-table"><thead><tr>' +
            thead +
            '</tr></thead><tbody>' +
            (pagina.length ? linhas : '<tr><td colspan="18" style="text-align:center;padding:20px;color:#94a3b8;">Nenhuma anotação encontrada.</td></tr>') +
            '</tbody></table></div>' +
            '<div class="crm-paginacao">' +
            '<button type="button" onclick="Crm.setListaPagina(' + Math.max(1, _crmListaPagina - 1) + ')" ' + (_crmListaPagina === 1 ? 'disabled' : '') + '>← Anterior</button>' +
            '<span>Página ' + _crmListaPagina + ' de ' + totalPaginas + ' (' + ordenadas.length + ' resultados)</span>' +
            '<button type="button" onclick="Crm.setListaPagina(' + Math.min(totalPaginas, _crmListaPagina + 1) + ')" ' + (_crmListaPagina === totalPaginas ? 'disabled' : '') + '>Próxima →</button>' +
            '</div>';

        var listEl = document.getElementById('crmListaNegocios');
        if (listEl) listEl.innerHTML = html;
        atualizarContagem(ordenadas.length, 'anotação', 'anotações');
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

        var origens = [];
        anotacoes.forEach(function (a) {
            if (a.origemDemanda && origens.indexOf(a.origemDemanda) === -1) origens.push(a.origemDemanda);
        });
        origens.sort();

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

        var inpBusca = document.getElementById('crmListaBusca');
        if (inpBusca && document.activeElement !== inpBusca) inpBusca.value = _crmListaFiltros.busca;
        var prioridade = document.getElementById('crmListaPrioridade');
        if (prioridade) prioridade.value = _crmListaFiltros.prioridade || '';
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
        _crmListaPagina = 1;
        renderizarConteudoAtivo();
    }

    function limparListaFiltros() {
        _crmListaFiltros = {
            busca: '', funilId: null, clienteId: null, prioridade: null, origemDemanda: null,
            prazoInicio: null, prazoFim: null, dataSolicitacaoInicio: null, dataSolicitacaoFim: null,
            mostrarFinalizadas: true, ordenarPor: 'prazo', ordenarDireccao: 'asc'
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

    // ── Calendário (Atividades de todos os negócios) ──

    var CAL_PERIODOS = [
        ['todos', 'Tudo'], ['paraFazer', 'Para fazer'], ['vencido', 'Vencido'],
        ['hoje', 'Hoje'], ['amanha', 'Amanhã'], ['semana', 'Esta semana'], ['proximaSemana', 'Próxima semana']
    ];
    var CAL_GRADE_INICIO_MIN = 6 * 60;   // 06:00
    var CAL_GRADE_FIM_MIN = 21 * 60;     // 21:00
    var DIAS_SEMANA_ABREV = ['dom', 'seg', 'ter', 'qua', 'qui', 'sex', 'sáb'];
    var MESES_ABREV = ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez'];

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

    function renderizarCalendarioView() {
        var ctx = atividadesCalContexto();
        limparFeriadosOrfaos(); // migração: remove feriados que ficaram salvos como atividade real (fase antiga, via Google)
        var reais = CrmStore.listarAtividades();
        var feriados = _calMostrarFeriados ? gerarFeriadosPseudoAtividades() : [];
        var todas = reais.concat(feriados);
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
                    '<button type="button" class="crm-cal-modo-btn' + (_calModo === 'lista' ? ' active' : '') + '" data-crm-action="setCalModo" data-valor="lista" title="Lista">☰</button>' +
                    '<button type="button" class="crm-cal-modo-btn' + (_calModo === 'semana' ? ' active' : '') + '" data-crm-action="setCalModo" data-valor="semana" title="Semana">🗓</button>' +
                    '<button type="button" class="crm-cal-modo-btn' + (_calModo === 'mes' ? ' active' : '') + '" data-crm-action="setCalModo" data-valor="mes" title="Mês">📆</button>' +
                '</div>' +
                '<input type="text" class="crm-busca" placeholder="Buscar atividade, negócio ou contato..." value="' + esc(_calBusca) + '" oninput="Crm.setCalBusca(this.value)">' +
                '<span class="crm-contagem">' + filtradas.length + (filtradas.length !== 1 ? ' atividades' : ' atividade') + '</span>' +
            '</div>' +
            renderizarGoogleStatus() +
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
            var selo = a.origemFeriado
                ? ' <span class="crm-cal-badge-feriado" title="Feriado nacional — somente leitura">🎉 Feriado</span>'
                : (a.origemGoogle ? ' <span class="crm-cal-badge-google" title="Importado da Google Agenda">📥 Google</span>' : '');
            var celAssunto = tipoInfo.icone + ' ' + esc(a.assunto || '(sem assunto)') + selo;

            return '<tr' + vencida + '>' +
                '<td class="crm-cal-td-check">' + (a.origemFeriado
                    ? '<input type="checkbox" disabled title="Feriado — não editável">'
                    : '<input type="checkbox" ' + (a.feito ? 'checked' : '') + ' data-crm-action="concluirAtividadeCal" data-id="' + esc(a.id) + '" data-feito="' + a.feito + '" onclick="event.stopPropagation()">') + '</td>' +
                (a.origemFeriado
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

        var filtroFeriados = '<div class="crm-cal-lista-filtros">' +
            '<label class="crm-check">' +
                '<input type="checkbox" ' + (_calMostrarFeriados ? 'checked' : '') + ' onchange="Crm.setMostrarFeriadosCal(this.checked)">' +
                'Mostrar feriados' +
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
        var porDia = CrmCalculos.agruparAtividadesPorDiaDaSemana(atividades, domingo);
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
                return '<div class="crm-cal-cab-dia' + (ehHoje ? ' crm-cal-cab-hoje' : '') + '">' + DIAS_SEMANA_ABREV[d.getUTCDay()] + ' ' + numero + '</div>';
            }).join('') + '</div>';

        var diaSemHora = Object.keys(porDia).map(function (dataIso) {
            var semHora = porDia[dataIso].filter(function (a) { return !a.horaInicio; });
            return '<div class="crm-cal-allday-col" data-crm-action="novaAtividadeNoDia" data-data="' + esc(dataIso) + '" title="Clique para agendar uma atividade neste dia">' + semHora.map(function (a) { return renderizarEventoCard(a); }).join('') + '</div>';
        }).join('');
        var linhaSemHora = '<div class="crm-cal-grade-allday"><div class="crm-cal-cab-gutter"></div>' + diaSemHora + '</div>';

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
            return '<div class="crm-cal-dia-col" data-crm-action="novaAtividadeNoDia" data-data="' + esc(dataIso) + '" title="Clique para agendar uma atividade neste dia">' + eventos + linhaAgoraCol + '</div>';
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
        var classeOrigem = a.origemFeriado ? ' crm-cal-evento-feriado' : (a.origemGoogle ? ' crm-cal-evento-google' : '');
        var icone = a.origemFeriado ? '🎉' : tipoInfo.icone;
        var acao = a.origemFeriado ? '' : ('data-crm-action="abrirModalAtividadeCal" data-id="' + esc(a.id) + '"');
        return '<div class="crm-cal-evento' + (a.feito ? ' crm-cal-evento-feito' : '') + classeOrigem + '" ' + acao + ' data-tipo="' + esc(a.tipo) + '"' +
            (estiloExtra ? (' style="' + estiloExtra + '"') : '') +
            (a.origemFeriado ? ' title="Feriado nacional — somente leitura"' : '') + '>' +
            '<span class="crm-cal-evento-icone">' + icone + '</span>' +
            '<span class="crm-cal-evento-txt">' + esc(a.assunto || '(sem assunto)') + '</span>' +
        '</div>';
    }

    function renderizarCalendarioMes(atividades) {
        var mesRef = _calMesRef;
        var porDia = CrmCalculos.agruparAtividadesPorGradeMes(atividades, mesRef);
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
            DIAS_SEMANA_ABREV.map(function (d) { return '<div class="crm-cal-mes-cab-dia">' + d + '</div>'; }).join('') +
        '</div>';

        var MAX_VISIVEIS = 4;
        var celulas = dias.map(function (dataIso) {
            var itens = porDia[dataIso];
            var foraDoMes = dataIso.slice(0, 7) !== mesAtual;
            var ehHoje = dataIso === hoje;
            var visiveis = itens.slice(0, MAX_VISIVEIS);
            var extras = itens.length - visiveis.length;
            var eventosHtml = visiveis.map(function (a) { return renderizarEventoCard(a); }).join('') +
                (extras > 0 ? '<div class="crm-cal-mes-mais">+' + extras + ' mais</div>' : '');
            var numeroMes = ehHoje ? '<span class="crm-cal-cab-numero-hoje">' + Number(dataIso.slice(8, 10)) + '</span>' : Number(dataIso.slice(8, 10));
            return '<div class="crm-cal-mes-dia' + (foraDoMes ? ' crm-cal-mes-dia-fora' : '') + (ehHoje ? ' crm-cal-mes-dia-hoje' : '') + '" data-crm-action="novaAtividadeNoDia" data-data="' + esc(dataIso) + '" title="Clique para agendar uma atividade neste dia">' +
                '<div class="crm-cal-mes-numero">' + numeroMes + '</div>' +
                '<div class="crm-cal-mes-eventos">' + eventosHtml + '</div>' +
            '</div>';
        }).join('');

        return nav + '<div class="crm-cal-mes-grade">' + cabecalho + '<div class="crm-cal-mes-corpo">' + celulas + '</div></div>';
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
    }

    function abrirModalAnotacao(id, negocioForcado, relacionadaForcada) {
        var anotacao = id ? CrmStore.getAnotacao(id) : null;
        document.getElementById('crmModalAnotacaoTitulo').textContent = id ? 'Editar anotação' : 'Nova anotação';
        document.getElementById('crmAnotacaoId').value = id || '';

        _anotacaoRelacionadaId = anotacao ? anotacao.anotacaoRelacionadaId : (relacionadaForcada || null);
        var pai = _anotacaoRelacionadaId ? CrmStore.getAnotacao(_anotacaoRelacionadaId) : null;
        var vinculoEl = document.getElementById('crmAnotacaoVinculo');
        if (vinculoEl) {
            vinculoEl.style.display = pai ? '' : 'none';
            vinculoEl.textContent = pai ? ('Vinculada a: ' + (pai.assunto || '(sem assunto)')) : '';
        }

        _anotacaoNegocioId = anotacao ? anotacao.negocioId : (negocioForcado || null);
        if (_anotacaoNegocioId) {
            selecionarNegocioAnotacao(_anotacaoNegocioId);
        } else {
            document.getElementById('crmAnotacaoNegocioBusca').value = '';
            document.getElementById('crmAnotacaoRelacionamento').textContent = '-';
            document.getElementById('crmAnotacaoClienteNome').textContent = '-';
        }
        document.getElementById('crmAnotacaoNegocioLista').style.display = 'none';
        document.getElementById('crmAnotacaoNegocioBusca').disabled = !!_anotacaoNegocioId;

        document.getElementById('crmAnotacaoAssunto').value = anotacao ? (anotacao.assunto || '') : '';
        document.getElementById('crmAnotacaoRemetente').value = anotacao ? (anotacao.remetente || '') : '';
        document.getElementById('crmAnotacaoOrigemDemanda').value = anotacao ? (anotacao.origemDemanda || '') : '';
        document.getElementById('crmAnotacaoDataSolicitacao').value = anotacao ? (anotacao.dataSolicitacao || '') : '';
        document.getElementById('crmAnotacaoTipoDoc').value = anotacao ? (anotacao.tipoDoc || '') : '';
        document.getElementById('crmAnotacaoNumeroDocumento').value = anotacao ? (anotacao.numeroDocumento || '') : '';
        document.getElementById('crmAnotacaoDestinatario').value = anotacao ? (anotacao.destinatario || '') : '';
        document.getElementById('crmAnotacaoAcaoRealizar').value = anotacao ? (anotacao.acaoRealizar || '') : '';
        document.getElementById('crmAnotacaoPrioridade').value = anotacao ? anotacao.prioridade : 'media';
        document.getElementById('crmAnotacaoPrazo').value = anotacao ? (anotacao.prazo || '') : '';
        document.getElementById('crmAnotacaoLembrarDias').value = anotacao ? (anotacao.lembrarDiasAntes || '') : '';
        document.getElementById('crmAnotacaoTags').value = anotacao ? (anotacao.tags || []).join(', ') : '';
        document.getElementById('crmAnotacaoObservacoes').value = anotacao ? (anotacao.observacoes || '') : '';
        document.getElementById('crmAnotacaoFinalizado').checked = anotacao ? !!anotacao.finalizado : false;
        document.getElementById('crmAnotacaoDataConclusao').value = anotacao ? (anotacao.dataConclusao || '') : '';
        document.getElementById('crmAnotacaoOQueFoiFeito').value = anotacao ? (anotacao.oQueFoiFeito || '') : '';

        var btnExcluir = document.getElementById('crmAnotacaoBtnExcluir');
        if (btnExcluir) btnExcluir.style.display = id ? '' : 'none';

        document.getElementById('modalAnotacao').style.display = 'flex';
    }

    function salvarAnotacao() {
        if (!requireAdminOrNotify()) return;
        if (!_anotacaoNegocioId) { Notifications.error('Selecione o negócio (Relacionamento/Cliente) desta anotação.'); return; }

        var assunto = document.getElementById('crmAnotacaoAssunto').value.trim();
        if (!assunto) { Notifications.error('Assunto é obrigatório.'); return; }

        var tagsRaw = document.getElementById('crmAnotacaoTags').value || '';
        var dados = {
            negocioId: _anotacaoNegocioId,
            anotacaoRelacionadaId: _anotacaoRelacionadaId,
            assunto: assunto,
            remetente: document.getElementById('crmAnotacaoRemetente').value.trim(),
            origemDemanda: document.getElementById('crmAnotacaoOrigemDemanda').value.trim(),
            dataSolicitacao: document.getElementById('crmAnotacaoDataSolicitacao').value || null,
            tipoDoc: document.getElementById('crmAnotacaoTipoDoc').value.trim(),
            numeroDocumento: document.getElementById('crmAnotacaoNumeroDocumento').value.trim(),
            destinatario: document.getElementById('crmAnotacaoDestinatario').value.trim(),
            acaoRealizar: document.getElementById('crmAnotacaoAcaoRealizar').value.trim(),
            prioridade: document.getElementById('crmAnotacaoPrioridade').value,
            prazo: document.getElementById('crmAnotacaoPrazo').value || null,
            lembrarDiasAntes: document.getElementById('crmAnotacaoLembrarDias').value.trim(),
            tags: tagsRaw.split(',').map(function (t) { return t.trim(); }).filter(Boolean),
            observacoes: document.getElementById('crmAnotacaoObservacoes').value.trim(),
            finalizado: document.getElementById('crmAnotacaoFinalizado').checked,
            dataConclusao: document.getElementById('crmAnotacaoDataConclusao').value || null,
            oQueFoiFeito: document.getElementById('crmAnotacaoOQueFoiFeito').value.trim()
        };

        var id = document.getElementById('crmAnotacaoId').value || null;
        if (id) {
            CrmStore.atualizarAnotacao(id, dados);
        } else {
            CrmStore.criarAnotacao(dados);
        }
        fecharModal('modalAnotacao');
        renderizarConteudoAtivo();
    }

    function excluirAnotacao() {
        if (!requireAdminOrNotify()) return;
        var id = document.getElementById('crmAnotacaoId').value;
        if (!id) return;
        if (!confirm('Excluir esta anotação? Esta ação não pode ser desfeita.')) return;
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
               '<div>' + esc([cliente.endereco, cliente.cidade, cliente.uf].filter(Boolean).join(', ') || '—') + '</div>')
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
        return '' +
            '<div class="crm-anot-linha">' +
                '<span class="crm-anot-icone" title="Anotação">🗒️</span>' +
                '<div class="crm-anot-nota' + (a.finalizado ? ' crm-anot-finalizada' : '') + '">' +
                    '<div data-crm-action="abrirModalAnotacao" data-id="' + esc(a.id) + '">' +
                        '<div class="crm-anot-nota-data">' + esc(formatarDataHora(a.criadoEm)) + (a.finalizado ? ' · ✓ Finalizado' : '') + '</div>' +
                        '<div class="crm-anot-nota-assunto">' + esc(a.assunto || '(sem assunto)') + '</div>' +
                        '<div class="crm-anot-nota-meta">' + prioridadeBadge + (prazoTxt ? ('<span>Prazo: ' + esc(prazoTxt) + '</span>') : '') + '</div>' +
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

    function selecionarCliente(id, nome) {
        document.getElementById('crmNegocioClienteId').value = id;
        document.getElementById('crmNegocioClienteBusca').value = nome;
        document.getElementById('crmClienteLista').style.display = 'none';
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
        adicionarItem: adicionarItem,

        abrirModalAnotacao: abrirModalAnotacao,
        salvarAnotacao: salvarAnotacao,
        excluirAnotacao: excluirAnotacao,
        buscarNegocioParaAnotacao: buscarNegocioParaAnotacao,

        abrirModalFunil: abrirModalFunil,
        adicionarEtapaFunil: adicionarEtapaFunil,
        criarFunilDeTemplate: criarFunilDeTemplate,
        salvarFunil: salvarFunil,

        setCalBusca: setCalBusca,
        setMostrarFeriadosCal: setMostrarFeriadosCal,
        abrirFormTipoAtividade: abrirFormTipoAtividade,
        salvarTipoAtividade: salvarTipoAtividade,
        abrirModalAtividadeCal: abrirModalAtividadeCal,
        salvarAtividadeCal: salvarAtividadeCal,
        excluirAtividadeCal: excluirAtividadeCal,
        buscarNegocioParaAtividadeCal: buscarNegocioParaAtividadeCal,
        atualizarMiniAgendaCal: atualizarMiniAgendaCal
    };
})();
