/**
 * crm-store.js - Único ponto que escreve em estoque.crm (Controle-Estoque)
 *
 * Adaptador da camada de estado: no Ponto havia um Proxy reativo com auto-save;
 * aqui o estado é o objeto global `estoque` e a persistência é manual via
 * `salvarDados()` (que serializa `estoque` inteiro para localStorage comprimido
 * e agenda o save no Firestore `app_data/latest`). Portanto `emLote(fn)` apenas
 * executa a mutação e chama `salvarDados()`.
 *
 * O CRM NÃO guarda clientes/produtos: lê-os ao vivo de estoque.clientes /
 * estoque.produtos (fonte única). Depende de CrmModel e CrmCalculos (globais,
 * carregados antes) e de `estoque`/`salvarDados` (definidos em app2.js).
 */
(function () {

    var LIMITE_HISTORICO_POR_ENTIDADE = 150;
    var LIMITE_HISTORICO_GLOBAL = 1500;
    var PODAR_A_CADA_N_ESCRITAS = 20;
    var _contadorEscritasHistorico = 0;

    function getEstoque() {
        return (typeof estoque !== 'undefined' && estoque) ? estoque : (window.estoque || null);
    }

    function getCrm() {
        var e = getEstoque();
        return e ? e.crm : null;
    }

    function persistir(imediato) {
        if (typeof salvarDados === 'function') {
            try { salvarDados(imediato ? { imediato: true } : {}); return; } catch (_) { /* fallback */ }
        }
        if (window.salvarDados) { try { window.salvarDados(imediato ? { imediato: true } : {}); } catch (_) {} }
    }

    /**
     * Agrupa uma mutação num único save. Sem Proxy: executa e persiste.
     */
    function emLote(fn) {
        fn();
        persistir(false);
    }

    /**
     * Garante que estoque.crm existe e está normalizado. Roda na inicialização
     * e após import de backup. Só grava se algo mudou, para não disparar um
     * write desnecessário no Firestore a cada abertura.
     */
    function ensureCrmDefault() {
        var e = getEstoque();
        if (!e) return;

        var atual = e.crm;
        var antes = JSON.stringify(atual || null);
        var normalizado = CrmModel.normalizarCrm(atual);

        if (!normalizado.funis.length) {
            var seed = CrmModel.funilDeTemplate('vendas');
            normalizado.funis.push(seed);
            normalizado.config.funilAtivoId = seed.id;
        }

        var depois = JSON.stringify(normalizado);
        if (depois !== antes) {
            e.crm = normalizado;
            persistir(false);
        } else {
            e.crm = normalizado; // garante referência normalizada em memória
        }
    }

    // ──────────────────────────────────────────────
    //  LEITURA — CRM
    // ──────────────────────────────────────────────

    function listarFunis() {
        var crm = getCrm();
        return crm ? crm.funis.slice() : [];
    }

    function getFunilAtivo() {
        var crm = getCrm();
        if (!crm) return null;
        return crm.funis.filter(function (f) { return f.id === crm.config.funilAtivoId; })[0] || crm.funis[0] || null;
    }

    function listarNegocios(funilId) {
        var crm = getCrm();
        if (!crm) return [];
        return crm.negocios.filter(function (n) {
            return !n.excluidoEm && (!funilId || n.funilId === funilId);
        });
    }

    function listarNegociosExcluidos(funilId) {
        var crm = getCrm();
        if (!crm) return [];
        return crm.negocios.filter(function (n) {
            return !!n.excluidoEm && (!funilId || n.funilId === funilId);
        });
    }

    function listarAtividades(negocioId) {
        var crm = getCrm();
        if (!crm) return [];
        var todas = crm.atividades || [];
        return negocioId ? todas.filter(function (a) { return a.negocioId === negocioId; }) : todas.slice();
    }

    function historicoDe(entidade, entidadeId) {
        var crm = getCrm();
        if (!crm) return [];
        return CrmCalculos.timelineDe(crm.historico, entidade, entidadeId);
    }

    // ──────────────────────────────────────────────
    //  LEITURA — coleções do ESTOQUE reusadas pelo CRM
    // ──────────────────────────────────────────────

    function listarClientes() {
        var e = getEstoque();
        return (e && Array.isArray(e.clientes)) ? e.clientes.slice() : [];
    }

    function getCliente(id) {
        return listarClientes().filter(function (c) { return String(c.id) === String(id); })[0] || null;
    }

    function listarProdutos() {
        var e = getEstoque();
        return (e && Array.isArray(e.produtos)) ? e.produtos.slice() : [];
    }

    function listarRepresentantes() {
        var e = getEstoque();
        return (e && Array.isArray(e.representantes)) ? e.representantes.slice() : [];
    }

    function listarPropostas() {
        var e = getEstoque();
        return (e && Array.isArray(e.propostas)) ? e.propostas.slice() : [];
    }

    /**
     * Cria um cliente novo no cadastro do Estoque (reuso — o CRM não tem
     * entidade própria de contatos). Mantém o formato usado por salvarCliente.
     */
    function criarCliente(dados) {
        var e = getEstoque();
        if (!e) return null;
        if (!Array.isArray(e.clientes)) e.clientes = [];
        var d = dados || {};
        var novo = {
            id: Date.now().toString(),
            nome: (d.nome || '').trim(),
            cnpj: (d.cnpj || '').trim(),
            tipoPessoa: d.tipoPessoa || 'PJ',
            endereco: (d.endereco || '').trim(),
            cidade: (d.cidade || '').trim(),
            uf: (d.uf || '').trim().toUpperCase(),
            telefone: (d.telefone || '').trim(),
            email: (d.email || '').trim(),
            contato: (d.contato || '').trim(),
            representante: d.representante || '',
            observacoes: (d.observacoes || '').trim()
        };
        emLote(function () {
            e.clientes.push(novo);
            // mantém o espelho global `clientes` em sincronia se ele divergiu
            try { if (typeof clientes !== 'undefined' && clientes !== e.clientes) clientes = e.clientes; } catch (_) {}
        });
        return novo;
    }

    // ──────────────────────────────────────────────
    //  CONFIG DA ABA
    // ──────────────────────────────────────────────

    function setFunilAtivo(id) {
        var crm = getCrm();
        if (!crm) return;
        emLote(function () { crm.config.funilAtivoId = id; });
    }

    function setVisao(visao) {
        var crm = getCrm();
        if (!crm) return;
        var validas = ['kanban', 'lista', 'previsao', 'excluidos'];
        emLote(function () { crm.config.visao = validas.indexOf(visao) !== -1 ? visao : 'kanban'; });
    }

    function setSubaba(subaba) {
        var crm = getCrm();
        if (!crm) return;
        emLote(function () { crm.config.subaba = subaba; });
    }

    function setDetalheAberto(id) {
        var crm = getCrm();
        if (!crm) return;
        emLote(function () { crm.config.detalheAbertoId = id || null; });
    }

    function setFiltros(filtros) {
        var crm = getCrm();
        if (!crm) return;
        emLote(function () {
            crm.config.filtros = Object.assign({}, crm.config.filtros, filtros || {});
        });
    }

    // ──────────────────────────────────────────────
    //  HISTÓRICO / NOTAS
    // ──────────────────────────────────────────────

    function registrarHistorico(entidade, entidadeId, tipo, texto, dadosExtra) {
        var crm = getCrm();
        if (!crm) return null;
        var item = {
            id: CrmModel.novoId('hst'),
            entidade: entidade,
            entidadeId: entidadeId,
            tipo: tipo,
            texto: texto || '',
            dados: dadosExtra || null,
            autor: '',
            editavel: (tipo === 'nota'),
            criadoEm: new Date().toISOString()
        };
        crm.historico.push(item);

        _contadorEscritasHistorico++;
        if (_contadorEscritasHistorico >= PODAR_A_CADA_N_ESCRITAS) {
            _contadorEscritasHistorico = 0;
            podarHistorico();
        }
        return item;
    }

    function adicionarNota(entidade, entidadeId, texto) {
        if (!texto || !String(texto).trim()) return null;
        var crm = getCrm();
        if (!crm) return null;
        var nota = null;
        emLote(function () {
            nota = registrarHistorico(entidade, entidadeId, 'nota', String(texto).trim());
        });
        return nota;
    }

    function editarNota(historicoId, novoTexto) {
        var crm = getCrm();
        if (!crm) return false;
        var item = crm.historico.filter(function (h) { return h.id === historicoId && h.editavel; })[0];
        if (!item) return false;
        emLote(function () {
            item.texto = novoTexto;
            item.editadoEm = new Date().toISOString();
        });
        return true;
    }

    function removerNota(historicoId) {
        var crm = getCrm();
        if (!crm) return false;
        var idx = -1;
        crm.historico.forEach(function (h, i) { if (h.id === historicoId && h.editavel) idx = i; });
        if (idx === -1) return false;
        emLote(function () { crm.historico.splice(idx, 1); });
        return true;
    }

    /**
     * Mantém no máximo LIMITE_HISTORICO_POR_ENTIDADE entradas por entidade e um
     * teto agregado; nunca descarta notas. Faz UMA reatribuição de crm.historico.
     */
    function podarHistorico() {
        var crm = getCrm();
        if (!crm || !crm.historico.length) return false;

        var porEntidade = {};
        var todos = [];
        for (var i = 0; i < crm.historico.length; i++) {
            var h = crm.historico[i];
            todos.push(h);
            var chave = h.entidade + ':' + h.entidadeId;
            (porEntidade[chave] = porEntidade[chave] || []).push(h);
        }

        var idsParaRemover = {};
        Object.keys(porEntidade).forEach(function (chave) {
            var itens = porEntidade[chave]
                .slice()
                .sort(function (a, b) { return String(a.criadoEm).localeCompare(String(b.criadoEm)); });
            var podaveis = itens.filter(function (h) { return !h.editavel; });
            var excedente = podaveis.length - LIMITE_HISTORICO_POR_ENTIDADE;
            if (excedente > 0) {
                podaveis.slice(0, excedente).forEach(function (h) { idsParaRemover[h.id] = true; });
            }
        });

        var sobreviventes = todos.filter(function (h) { return !idsParaRemover[h.id]; });
        if (sobreviventes.length > LIMITE_HISTORICO_GLOBAL) {
            var podaveisGlobal = sobreviventes
                .filter(function (h) { return !h.editavel; })
                .sort(function (a, b) { return String(a.criadoEm).localeCompare(String(b.criadoEm)); });
            var excedenteGlobal = sobreviventes.length - LIMITE_HISTORICO_GLOBAL;
            for (var k = 0; k < excedenteGlobal && k < podaveisGlobal.length; k++) {
                idsParaRemover[podaveisGlobal[k].id] = true;
            }
        }

        if (!Object.keys(idsParaRemover).length) return false;

        var mantidos = todos.filter(function (h) { return !idsParaRemover[h.id]; });
        emLote(function () { crm.historico = mantidos; });
        return true;
    }

    // O Estoque não tem o indicador de tamanho do Ponto — no-op seguro p/ paridade
    function atualizarIndicadorTamanho() { /* sem UI equivalente no Estoque */ }

    // ──────────────────────────────────────────────
    //  NEGÓCIOS
    // ──────────────────────────────────────────────

    function criarNegocio(dados) {
        var crm = getCrm();
        if (!crm) return null;
        var negocio = CrmModel.criarNegocio(dados);
        var irmaos = crm.negocios.filter(function (n) { return n.etapaId === negocio.etapaId && !n.excluidoEm; });
        negocio.ordem = irmaos.length;

        emLote(function () {
            crm.negocios.push(negocio);
            registrarHistorico('negocio', negocio.id, 'criacao', 'Negócio criado');
        });
        return negocio;
    }

    function atualizarNegocio(id, patch) {
        var crm = getCrm();
        if (!crm) return null;
        var n = crm.negocios.filter(function (x) { return x.id === id; })[0];
        if (!n) return null;

        var antes = {};
        CrmModel.CAMPOS_AUDITAVEIS_NEGOCIO.forEach(function (campo) { antes[campo] = n[campo]; });

        emLote(function () {
            Object.keys(patch || {}).forEach(function (campo) {
                if (campo === 'id') return;
                n[campo] = patch[campo];
            });
            n.atualizadoEm = new Date().toISOString();

            CrmModel.CAMPOS_AUDITAVEIS_NEGOCIO.forEach(function (campo) {
                if (Object.prototype.hasOwnProperty.call(patch || {}, campo) && antes[campo] !== n[campo]) {
                    registrarHistorico('negocio', id, 'campo',
                        'Campo "' + campo + '" alterado',
                        { campo: campo, de: antes[campo], para: n[campo] });
                }
            });
        });
        return n;
    }

    function removerNegocio(id) {
        var crm = getCrm();
        if (!crm) return false;
        var n = crm.negocios.filter(function (x) { return x.id === id; })[0];
        if (!n || n.excluidoEm) return false;
        emLote(function () {
            n.excluidoEm = new Date().toISOString();
            registrarHistorico('negocio', id, 'exclusao', 'Negócio movido para Excluídos');
        });
        return true;
    }

    function restaurarNegocio(id) {
        var crm = getCrm();
        if (!crm) return false;
        var n = crm.negocios.filter(function (x) { return x.id === id; })[0];
        if (!n || !n.excluidoEm) return false;
        emLote(function () {
            n.excluidoEm = null;
            registrarHistorico('negocio', id, 'exclusao', 'Negócio restaurado da lixeira');
        });
        return true;
    }

    function excluirNegocioDefinitivo(id) {
        var crm = getCrm();
        if (!crm) return false;
        var idx = -1;
        crm.negocios.forEach(function (n, i) { if (n.id === id) idx = i; });
        if (idx === -1) return false;
        emLote(function () {
            crm.negocios.splice(idx, 1);
            crm.atividades = (crm.atividades || []).filter(function (a) { return a.negocioId !== id; });
            crm.historico = crm.historico.filter(function (h) {
                return !(h.entidade === 'negocio' && h.entidadeId === id);
            });
        });
        return true;
    }

    function setParticipantes(negocioId, pessoaIds) {
        var crm = getCrm();
        if (!crm) return false;
        var n = crm.negocios.filter(function (x) { return x.id === negocioId; })[0];
        if (!n) return false;
        emLote(function () {
            n.participantes = (pessoaIds || []).slice();
            n.atualizadoEm = new Date().toISOString();
        });
        return true;
    }

    function moverNegocio(id, etapaId, indice) {
        var crm = getCrm();
        if (!crm) return false;
        var n = crm.negocios.filter(function (x) { return x.id === id; })[0];
        if (!n) return false;
        var funil = crm.funis.filter(function (f) { return f.id === n.funilId; })[0];
        if (!funil) return false;
        var etapaAnt = funil.etapas.filter(function (e) { return e.id === n.etapaId; })[0] || null;
        var etapaNova = funil.etapas.filter(function (e) { return e.id === etapaId; })[0];
        if (!etapaNova) return false;

        emLote(function () {
            var etapaMudou = !etapaAnt || etapaAnt.id !== etapaNova.id;

            n.etapaId = etapaNova.id;
            n.status = (etapaNova.tipo === 'ganho') ? 'ganho' : (etapaNova.tipo === 'perdido' ? 'perdido' : 'aberto');
            if (n.status !== 'aberto' && !n.dataFechamento) {
                n.dataFechamento = new Date().toISOString().slice(0, 10);
            }
            if (n.status === 'aberto') n.dataFechamento = null;
            n.atualizadoEm = new Date().toISOString();

            var negociosCrus = crm.negocios.map(function (x) {
                return { id: x.id, etapaId: x.etapaId, ordem: x.ordem };
            });
            var novasOrdens = CrmCalculos.reordenarNaEtapa(negociosCrus, etapaNova.id, id, indice);
            novasOrdens.forEach(function (par) {
                var alvo = crm.negocios.filter(function (x) { return x.id === par.id; })[0];
                if (alvo) alvo.ordem = par.ordem;
            });

            if (etapaMudou) {
                registrarHistorico('negocio', id, 'etapa',
                    'Movido de "' + (etapaAnt ? etapaAnt.nome : '—') + '" para "' + etapaNova.nome + '"',
                    { campo: 'etapaId', de: etapaAnt ? etapaAnt.id : null, para: etapaNova.id });
            }
        });
        return true;
    }

    function marcarGanho(id) {
        var crm = getCrm();
        if (!crm) return false;
        var n = crm.negocios.filter(function (x) { return x.id === id; })[0];
        if (!n) return false;
        var funil = crm.funis.filter(function (f) { return f.id === n.funilId; })[0];
        if (!funil) return false;
        var etapaGanho = funil.etapas.filter(function (e) { return e.tipo === 'ganho'; })[0];
        return etapaGanho ? moverNegocio(id, etapaGanho.id, null) : false;
    }

    function marcarPerdido(id, motivo) {
        var crm = getCrm();
        if (!crm) return false;
        var n = crm.negocios.filter(function (x) { return x.id === id; })[0];
        if (!n) return false;
        var funil = crm.funis.filter(function (f) { return f.id === n.funilId; })[0];
        if (!funil) return false;
        var etapaPerdido = funil.etapas.filter(function (e) { return e.tipo === 'perdido'; })[0];
        if (!etapaPerdido) return false;
        var ok = moverNegocio(id, etapaPerdido.id, null);
        if (ok && motivo) emLote(function () { n.motivoPerda = motivo; });
        return ok;
    }

    // ──────────────────────────────────────────────
    //  ATIVIDADES
    // ──────────────────────────────────────────────

    function criarAtividade(dados) {
        var crm = getCrm();
        if (!crm) return null;
        var atividade = CrmModel.criarAtividade(dados);
        emLote(function () {
            if (!crm.atividades) crm.atividades = [];
            crm.atividades.push(atividade);
            var rotulo = (CrmModel.TIPOS_ATIVIDADE[atividade.tipo] || {}).rotulo || atividade.tipo;
            registrarHistorico('negocio', atividade.negocioId, 'atividade',
                rotulo + ' agendada: ' + atividade.assunto,
                { atividadeId: atividade.id, acao: 'criada' });
        });
        return atividade;
    }

    function atualizarAtividade(id, patch) {
        var crm = getCrm();
        if (!crm) return null;
        var a = (crm.atividades || []).filter(function (x) { return x.id === id; })[0];
        if (!a) return null;
        emLote(function () {
            Object.keys(patch || {}).forEach(function (campo) {
                if (campo === 'id') return;
                a[campo] = patch[campo];
            });
            a.atualizadoEm = new Date().toISOString();
        });
        return a;
    }

    function concluirAtividade(id, feito) {
        var crm = getCrm();
        if (!crm) return false;
        var a = (crm.atividades || []).filter(function (x) { return x.id === id; })[0];
        if (!a) return false;
        var marcar = (feito !== false);
        emLote(function () {
            a.feito = marcar;
            a.feitoEm = marcar ? new Date().toISOString() : null;
            a.atualizadoEm = new Date().toISOString();
            var rotulo = (CrmModel.TIPOS_ATIVIDADE[a.tipo] || {}).rotulo || a.tipo;
            registrarHistorico('negocio', a.negocioId, 'atividade',
                rotulo + (marcar ? ' concluída: ' : ' reaberta: ') + a.assunto,
                { atividadeId: a.id, acao: marcar ? 'concluida' : 'reaberta' });
        });
        return true;
    }

    function removerAtividade(id) {
        var crm = getCrm();
        if (!crm) return false;
        var idx = -1;
        (crm.atividades || []).forEach(function (a, i) { if (a.id === id) idx = i; });
        if (idx === -1) return false;
        emLote(function () { crm.atividades.splice(idx, 1); });
        return true;
    }

    // ──────────────────────────────────────────────
    //  FUNIS
    // ──────────────────────────────────────────────

    function criarFunil(dados) {
        var crm = getCrm();
        if (!crm) return null;
        var funil = (dados && dados.template) ? CrmModel.funilDeTemplate(dados.template) : CrmModel.criarFunil(dados);
        if (!funil) return null;
        if (dados && dados.nome) funil.nome = dados.nome;
        emLote(function () {
            crm.funis.push(funil);
            if (!crm.config.funilAtivoId) crm.config.funilAtivoId = funil.id;
        });
        return funil;
    }

    function atualizarFunil(id, patch) {
        var crm = getCrm();
        if (!crm) return null;
        var f = crm.funis.filter(function (x) { return x.id === id; })[0];
        if (!f) return null;
        emLote(function () {
            Object.keys(patch || {}).forEach(function (campo) {
                if (campo === 'id' || campo === 'etapas') return;
                f[campo] = patch[campo];
            });
            f.atualizadoEm = new Date().toISOString();
        });
        return f;
    }

    function definirEtapasFunil(funilId, etapasBrutas) {
        var crm = getCrm();
        if (!crm) return null;
        var f = crm.funis.filter(function (x) { return x.id === funilId; })[0];
        if (!f) return null;
        var etapas = (etapasBrutas || [])
            .map(function (e, idx) { return CrmModel.normalizarEtapa(e, idx); })
            .sort(function (a, b) { return a.ordem - b.ordem; });
        emLote(function () {
            f.etapas = etapas;
            f.atualizadoEm = new Date().toISOString();
        });
        return f;
    }

    function arquivarFunil(id, arquivado) {
        var crm = getCrm();
        if (!crm) return false;
        var f = crm.funis.filter(function (x) { return x.id === id; })[0];
        if (!f) return false;
        emLote(function () { f.arquivado = (arquivado !== false); });
        return true;
    }

    // Força persistência imediata (usado no fim do drag-and-drop)
    function flush() { persistir(true); }

    window.CrmStore = {
        ensureCrmDefault: ensureCrmDefault,
        getCrm: getCrm,
        atualizarIndicadorTamanho: atualizarIndicadorTamanho,
        flush: flush,

        listarFunis: listarFunis,
        getFunilAtivo: getFunilAtivo,
        listarNegocios: listarNegocios,
        listarNegociosExcluidos: listarNegociosExcluidos,
        listarAtividades: listarAtividades,
        historicoDe: historicoDe,

        listarClientes: listarClientes,
        getCliente: getCliente,
        criarCliente: criarCliente,
        listarProdutos: listarProdutos,
        listarRepresentantes: listarRepresentantes,
        listarPropostas: listarPropostas,

        setFunilAtivo: setFunilAtivo,
        setVisao: setVisao,
        setSubaba: setSubaba,
        setDetalheAberto: setDetalheAberto,
        setFiltros: setFiltros,

        registrarHistorico: registrarHistorico,
        adicionarNota: adicionarNota,
        editarNota: editarNota,
        removerNota: removerNota,
        podarHistorico: podarHistorico,

        criarNegocio: criarNegocio,
        atualizarNegocio: atualizarNegocio,
        removerNegocio: removerNegocio,
        restaurarNegocio: restaurarNegocio,
        excluirNegocioDefinitivo: excluirNegocioDefinitivo,
        setParticipantes: setParticipantes,
        moverNegocio: moverNegocio,
        marcarGanho: marcarGanho,
        marcarPerdido: marcarPerdido,

        criarAtividade: criarAtividade,
        atualizarAtividade: atualizarAtividade,
        concluirAtividade: concluirAtividade,
        removerAtividade: removerAtividade,

        criarFunil: criarFunil,
        atualizarFunil: atualizarFunil,
        definirEtapasFunil: definirEtapasFunil,
        arquivarFunil: arquivarFunil,

        emLote: emLote
    };
})();
