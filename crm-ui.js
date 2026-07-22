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

    var ICONES_HISTORICO = { criacao: '✨', campo: '✎', etapa: '➡️', exclusao: '🗑', atividade: '📅', nota: '📝' };
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
        var kanban = document.getElementById('crmKanban');
        var lista = document.getElementById('crmListaNegocios');
        var previsao = document.getElementById('crmPrevisao');
        var excluidos = document.getElementById('crmExcluidos');
        var detalhe = document.getElementById('crmViewDetalhe');
        var barra = document.querySelector('.crm-barra-visoes');
        if (!kanban) return;

        if (_detalheId) {
            [kanban, lista, previsao, excluidos].forEach(function (el) { el.style.display = 'none'; });
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

        document.querySelectorAll('.crm-visao-btn').forEach(function (b) {
            b.classList.toggle('active', b.dataset.valor === _visao);
        });

        var funil = CrmStore.getFunilAtivo();
        if (!funil) return;

        if (_visao === 'kanban') renderizarKanban(funil);
        else if (_visao === 'lista') renderizarListaView(funil);
        else if (_visao === 'previsao') renderizarPrevisaoView(funil);
        else if (_visao === 'excluidos') renderizarExcluidosView(funil);
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

    function atualizarContagem(n) {
        var el = document.getElementById('crmContagem');
        if (el) el.textContent = n + ' negócio' + (n !== 1 ? 's' : '');
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

    // ── Lista ──

    function renderizarListaView(funil) {
        var negocios = negociosVisiveis(funil);
        var atividades = CrmStore.listarAtividades();
        var etapaPorId = {};
        funil.etapas.forEach(function (e) { etapaPorId[e.id] = e; });

        var linhas = negocios.map(function (n) {
            var etapa = etapaPorId[n.etapaId];
            var cliente = CrmKanban.nomeCliente(n);
            var prox = CrmCalculos.proximaAtividade(atividades, n.id);
            return '<tr>' +
                '<td><span class="crm-link" data-crm-action="abrirDetalhe" data-id="' + esc(n.id) + '">' + esc(n.titulo || '(sem título)') + '</span></td>' +
                '<td>' + esc(cliente || '—') + '</td>' +
                '<td>' + esc(etapa ? etapa.nome : '—') + '</td>' +
                (funil.mostrarValor ? '<td>' + esc(CrmCalculos.formatarMoeda(n.valor, n.moeda)) + '</td>' : '') +
                '<td>' + (prox ? ('📅 ' + esc(DateUtils.formatBR(prox.data))) : '<span class="crm-alerta-atividade">⚠️ Nenhuma</span>') + '</td>' +
                '<td>' + esc(n.responsavel || '—') + '</td>' +
                '<td><button type="button" class="btn-secondary crm-btn-mini" data-crm-action="excluirNegocio" data-id="' + esc(n.id) + '">Excluir</button></td>' +
            '</tr>';
        }).join('');

        document.getElementById('crmListaNegocios').innerHTML =
            '<div class="table-container"><table><thead><tr><th>Título</th><th>Cliente</th><th>Etapa</th>' +
            (funil.mostrarValor ? '<th>Valor</th>' : '') +
            '<th>Próxima atividade</th><th>Proprietário</th><th>Ações</th></tr></thead><tbody>' +
            (linhas || '<tr><td colspan="7">Nenhum negócio encontrado.</td></tr>') + '</tbody></table></div>';
        atualizarContagem(negocios.length);
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

        var corpoAba = _abaDetalhe === 'anotacoes' ? renderizarComposerNota() : renderizarComposerAtividade(negocio);

        return abas + corpoAba +
            '<div class="crm-det-bloco">' + renderizarFoco(negocio, atividades) + '</div>' +
            '<div class="crm-det-bloco">' + renderizarHistoricoFiltrado(negocio) + '</div>';
    }

    function renderizarComposerAtividade(negocio) {
        var editando = _atividadeEditandoId ? CrmStore.listarAtividades(negocio.id).filter(function (a) { return a.id === _atividadeEditandoId; })[0] : null;
        var tipo = editando ? editando.tipo : 'tarefa';
        var pills = Object.keys(CrmModel.TIPOS_ATIVIDADE).map(function (chave) {
            var t = CrmModel.TIPOS_ATIVIDADE[chave];
            return '<button type="button" class="crm-atv-tipo' + (chave === tipo ? ' active' : '') + '" data-crm-action="escolherTipoAtividade" data-valor="' + chave + '">' + t.icone + ' ' + esc(t.rotulo) + '</button>';
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

    function renderizarComposerNota() {
        return '' +
            '<div class="crm-atv-composer">' +
                '<textarea id="crmNotaTexto" rows="3" placeholder="Escreva uma anotação..."></textarea>' +
                '<div class="crm-atv-acoes"><button type="button" class="btn btn-primary" data-crm-action="adicionarNota" data-negocio-id="' + esc(_detalheId || '') + '">Registrar</button></div>' +
            '</div>';
    }

    function renderizarFoco(negocio, atividades) {
        var pendentes = CrmCalculos.atividadesPendentesDe(atividades, negocio.id);
        if (!pendentes.length) {
            return '<div class="crm-bloco-titulo">Foco</div><p>Nenhum item de foco. Agende uma atividade acima.</p>';
        }
        var hoje = CrmCalculos.hojeIso();
        var itens = pendentes.map(function (a) {
            var atrasada = a.data && a.data < hoje;
            var t = CrmModel.TIPOS_ATIVIDADE[a.tipo] || { icone: '', rotulo: a.tipo };
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
        var mapaFiltro = { atividades: 'atividade', anotacoes: 'nota' };
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

        if (id) CrmStore.atualizarAtividade(id, dados);
        else CrmStore.criarAtividade(dados);

        _atividadeEditandoId = null;
        renderizarConteudoAtivo();
    }

    function salvarNota(negocioId) {
        if (!requireAdminOrNotify()) return;
        var el = document.getElementById('crmNotaTexto');
        var texto = el ? el.value : '';
        if (!texto || !texto.trim()) return;
        CrmStore.adicionarNota('negocio', negocioId || _detalheId, texto);
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
            renderizarConteudoAtivo();
        },
        excluirAtividade: function (el) {
            if (!requireAdminOrNotify()) return;
            confirmarEExecutar('Excluir esta atividade?', function () {
                CrmStore.removerAtividade(el.dataset.id);
                renderizarConteudoAtivo();
            });
        },
        adicionarNota: function (el) { salvarNota(el.dataset.negocioId); },

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
        }
    };

    function aoClicar(e) {
        var el = e.target.closest('[data-crm-action]');
        if (!el) {
            if (!e.target.closest('.crm-autocomplete')) {
                var lista = document.getElementById('crmClienteLista');
                if (lista) lista.style.display = 'none';
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

        abrirModalNegocio: abrirModalNegocio,
        salvarNegocio: salvarNegocio,
        buscarCliente: buscarCliente,
        adicionarItem: adicionarItem,

        abrirModalFunil: abrirModalFunil,
        adicionarEtapaFunil: adicionarEtapaFunil,
        criarFunilDeTemplate: criarFunilDeTemplate,
        salvarFunil: salvarFunil
    };
})();
