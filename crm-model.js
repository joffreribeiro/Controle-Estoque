/**
 * crm-model.js - Forma dos dados do módulo de Relacionamento (CRM) do Controle-Estoque
 * Funções puras: factories, normalização/migração e validação.
 * Sem DOM, sem estado global — testável isoladamente com Vitest.
 *
 * Diferenças em relação ao CRM do Ponto:
 *  - o negócio referencia um CLIENTE existente (estoque.clientes) via `clienteId`,
 *    em vez de o CRM manter suas próprias pessoas/organizações;
 *  - o negócio pode ter ITENS de produto (estoque.produtos) e uma referência
 *    opcional a uma PROPOSTA (estoque.propostas).
 */

const CAMPOS_AUDITAVEIS_NEGOCIO = ['titulo', 'valor', 'etapaId', 'responsavel', 'dataPrevisao', 'status', 'clienteId', 'propostaId'];

const TIPOS_FUNIL = ['vendas', 'demandas', 'projetos'];
const TIPOS_ETAPA = ['aberta', 'ganho', 'perdido'];
const STATUS_NEGOCIO = ['aberto', 'ganho', 'perdido'];

// Tipos de atividade agendável (padrão Pipedrive). O ícone é um emoji para
// não depender de bibliotecas de ícone nos módulos puros.
const TIPOS_ATIVIDADE = {
    chamada: { rotulo: 'Chamada', icone: '📞' },
    reuniao: { rotulo: 'Reunião', icone: '👥' },
    tarefa: { rotulo: 'Tarefa', icone: '✔️' },
    prazo: { rotulo: 'Prazo', icone: '🚩' },
    email: { rotulo: 'E-mail', icone: '✉️' },
    almoco: { rotulo: 'Almoço', icone: '🍽️' }
};

const TEMPLATES_FUNIL = {
    vendas: {
        nome: 'Comercial',
        mostrarValor: true,
        etapas: ['Qualificação', 'Contato feito', 'Proposta', 'Negociação', 'Ganho', 'Perdido']
    },
    demandas: {
        nome: 'Demandas',
        mostrarValor: false,
        etapas: ['Recebida', 'Em análise', 'Em execução', 'Aguardando terceiros', 'Concluída', 'Cancelada']
    },
    projetos: {
        nome: 'Projetos',
        mostrarValor: true,
        etapas: ['Prospecção', 'Planejamento', 'Execução', 'Homologação', 'Entregue', 'Cancelado']
    }
};

function nowIso() {
    return new Date().toISOString();
}

function novoId(prefixo) {
    return prefixo + '_' + Date.now().toString(36) + '_' + Math.random().toString(36).substr(2, 6);
}

function ehObjeto(v) {
    return !!v && typeof v === 'object' && !Array.isArray(v);
}

// ──────────────────────────────────────────────
//  NORMALIZAÇÃO — uma função por entidade, todas puras e idempotentes
// ──────────────────────────────────────────────

function normalizarEtapa(eBruta, idx) {
    const e = ehObjeto(eBruta) ? eBruta : {};
    const tipo = TIPOS_ETAPA.indexOf(e.tipo) !== -1 ? e.tipo : 'aberta';
    const probabilidadeDefault = tipo === 'ganho' ? 100 : (tipo === 'perdido' ? 0 : 20);
    return {
        id: e.id || novoId('etp'),
        nome: (typeof e.nome === 'string' && e.nome.trim()) ? e.nome : ('Etapa ' + (idx + 1)),
        ordem: Number.isFinite(e.ordem) ? e.ordem : idx,
        cor: (typeof e.cor === 'string' && e.cor) ? e.cor : '#64748b',
        tipo,
        probabilidade: Number.isFinite(e.probabilidade) ? e.probabilidade : probabilidadeDefault
    };
}

function normalizarFunil(fBruto) {
    const f = ehObjeto(fBruto) ? fBruto : {};
    const etapasBrutas = Array.isArray(f.etapas) ? f.etapas : [];
    const etapas = etapasBrutas
        .map(normalizarEtapa)
        .sort((a, b) => a.ordem - b.ordem);

    return {
        id: f.id || novoId('fnl'),
        nome: (typeof f.nome === 'string' && f.nome.trim()) ? f.nome : 'Funil sem nome',
        tipo: TIPOS_FUNIL.indexOf(f.tipo) !== -1 ? f.tipo : 'vendas',
        mostrarValor: f.mostrarValor !== false,
        moeda: (typeof f.moeda === 'string' && f.moeda) ? f.moeda : 'BRL',
        ordem: Number.isFinite(f.ordem) ? f.ordem : 0,
        arquivado: !!f.arquivado,
        etapas,
        criadoEm: f.criadoEm || nowIso(),
        atualizadoEm: f.atualizadoEm || nowIso()
    };
}

function normalizarItem(iBruto) {
    const i = ehObjeto(iBruto) ? iBruto : {};
    const q = Number(i.quantidade);
    const p = Number(i.precoUnit);
    return {
        produtoId: (i.produtoId !== undefined && i.produtoId !== null) ? i.produtoId : null,
        nome: typeof i.nome === 'string' ? i.nome : '',
        quantidade: (Number.isFinite(q) && q > 0) ? q : 1,
        precoUnit: (Number.isFinite(p) && p >= 0) ? p : 0
    };
}

function normalizarNegocio(nBruto) {
    const n = ehObjeto(nBruto) ? nBruto : {};
    const itens = Array.isArray(n.itens) ? n.itens.map(normalizarItem) : [];

    // Se há itens, o valor é derivado deles; senão aceita valor manual.
    let valor;
    if (itens.length) {
        valor = itens.reduce((acc, it) => acc + (it.quantidade * it.precoUnit), 0);
    } else {
        const temValor = n.valor !== null && n.valor !== undefined && n.valor !== '';
        const valorNumerico = temValor ? Number(n.valor) : null;
        valor = (valorNumerico !== null && !isNaN(valorNumerico)) ? valorNumerico : null;
    }

    return {
        id: n.id || novoId('ngc'),
        funilId: n.funilId || null,
        etapaId: n.etapaId || null,
        titulo: typeof n.titulo === 'string' ? n.titulo : '',
        clienteId: (n.clienteId !== undefined && n.clienteId !== null) ? n.clienteId : null,
        itens: itens,
        valor: valor,
        moeda: (typeof n.moeda === 'string' && n.moeda) ? n.moeda : 'BRL',
        propostaId: (n.propostaId !== undefined && n.propostaId !== null) ? n.propostaId : null,
        responsavel: typeof n.responsavel === 'string' ? n.responsavel : '',
        status: STATUS_NEGOCIO.indexOf(n.status) !== -1 ? n.status : 'aberto',
        motivoPerda: typeof n.motivoPerda === 'string' ? n.motivoPerda : '',
        origem: typeof n.origem === 'string' ? n.origem : '',
        dataRecebimento: n.dataRecebimento || null,
        dataPrevisao: n.dataPrevisao || null,
        dataFechamento: n.dataFechamento || null,
        ordem: Number.isFinite(n.ordem) ? n.ordem : 0,
        tags: Array.isArray(n.tags) ? n.tags.slice() : [],
        participantes: Array.isArray(n.participantes) ? n.participantes.slice() : [],
        excluidoEm: n.excluidoEm || null,
        descricao: typeof n.descricao === 'string' ? n.descricao : '',
        criadoEm: n.criadoEm || nowIso(),
        atualizadoEm: n.atualizadoEm || nowIso()
    };
}

function normalizarAtividade(aBruta) {
    const a = ehObjeto(aBruta) ? aBruta : {};
    const tipo = Object.prototype.hasOwnProperty.call(TIPOS_ATIVIDADE, a.tipo) ? a.tipo : 'tarefa';
    return {
        id: a.id || novoId('atv'),
        negocioId: a.negocioId || null,
        tipo,
        assunto: typeof a.assunto === 'string' ? a.assunto : '',
        descricao: typeof a.descricao === 'string' ? a.descricao : '',
        data: a.data || null,
        horaInicio: typeof a.horaInicio === 'string' ? a.horaInicio : '',
        horaFim: typeof a.horaFim === 'string' ? a.horaFim : '',
        feito: !!a.feito,
        feitoEm: a.feitoEm || null,
        criadoEm: a.criadoEm || nowIso(),
        atualizadoEm: a.atualizadoEm || nowIso()
    };
}

function normalizarHistoricoItem(hBruto) {
    const h = ehObjeto(hBruto) ? hBruto : {};
    const tipo = typeof h.tipo === 'string' && h.tipo ? h.tipo : 'campo';
    return {
        id: h.id || novoId('hst'),
        entidade: typeof h.entidade === 'string' ? h.entidade : 'negocio',
        entidadeId: h.entidadeId || null,
        tipo,
        texto: typeof h.texto === 'string' ? h.texto : '',
        dados: ehObjeto(h.dados) ? h.dados : null,
        autor: typeof h.autor === 'string' ? h.autor : '',
        editavel: typeof h.editavel === 'boolean' ? h.editavel : (tipo === 'nota'),
        criadoEm: h.criadoEm || nowIso()
    };
}

function normalizarConfig(cBruta, funis) {
    const c = ehObjeto(cBruta) ? cBruta : {};
    const idsValidos = funis.map(f => f.id);
    const funilAtivoId = idsValidos.indexOf(c.funilAtivoId) !== -1 ? c.funilAtivoId : (funis[0] ? funis[0].id : null);
    const filtrosBrutos = ehObjeto(c.filtros) ? c.filtros : {};
    return {
        funilAtivoId,
        visao: ['kanban', 'lista', 'previsao', 'excluidos'].indexOf(c.visao) !== -1 ? c.visao : 'kanban',
        subaba: 'negocios',
        detalheAbertoId: c.detalheAbertoId || null,
        filtros: {
            busca: typeof filtrosBrutos.busca === 'string' ? filtrosBrutos.busca : '',
            responsavel: typeof filtrosBrutos.responsavel === 'string' ? filtrosBrutos.responsavel : '',
            status: typeof filtrosBrutos.status === 'string' ? filtrosBrutos.status : ''
        }
    };
}

/**
 * Normaliza o objeto crm inteiro: garante todos os arrays, preenche defaults,
 * gera IDs faltantes e realoca negócios órfãos (funil/etapa inexistente)
 * para a primeira etapa aberta do primeiro funil. Pura e idempotente.
 *
 * NÃO guarda clientes/produtos — o CRM lê essas coleções ao vivo de
 * estoque.clientes/estoque.produtos pelo store adaptador.
 */
function normalizarCrm(crmBruto) {
    const crm = ehObjeto(crmBruto) ? crmBruto : {};

    const funis = (Array.isArray(crm.funis) ? crm.funis : []).map(normalizarFunil);

    const idsFunilValidos = funis.map(f => f.id);
    const primeiraEtapaAbertaPorFunil = {};
    funis.forEach(f => {
        const aberta = f.etapas.filter(e => e.tipo === 'aberta')[0] || f.etapas[0] || null;
        primeiraEtapaAbertaPorFunil[f.id] = aberta ? aberta.id : null;
    });

    const negocios = (Array.isArray(crm.negocios) ? crm.negocios : [])
        .map(normalizarNegocio)
        .filter(() => idsFunilValidos.length > 0)
        .map(n => {
            let funilId = n.funilId;
            if (idsFunilValidos.indexOf(funilId) === -1) {
                funilId = idsFunilValidos[0];
            }
            const funil = funis[idsFunilValidos.indexOf(funilId)];
            const idsEtapaDoFunil = funil ? funil.etapas.map(e => e.id) : [];
            let etapaId = n.etapaId;
            if (!etapaId || idsEtapaDoFunil.indexOf(etapaId) === -1) {
                etapaId = primeiraEtapaAbertaPorFunil[funilId] || null;
            }
            return Object.assign({}, n, { funilId, etapaId });
        });

    const historico = (Array.isArray(crm.historico) ? crm.historico : []).map(normalizarHistoricoItem);

    const idsNegocioValidos = {};
    negocios.forEach(n => { idsNegocioValidos[n.id] = true; });
    const atividades = (Array.isArray(crm.atividades) ? crm.atividades : [])
        .map(normalizarAtividade)
        .filter(a => a.negocioId && idsNegocioValidos[a.negocioId]);

    const config = normalizarConfig(crm.config, funis);

    return {
        versao: 1,
        funis,
        negocios,
        atividades,
        historico,
        config
    };
}

// ──────────────────────────────────────────────
//  FACTORIES
// ──────────────────────────────────────────────

function criarFunil(dados) { return normalizarFunil(dados); }
function criarNegocio(dados) { return normalizarNegocio(dados); }
function criarAtividade(dados) { return normalizarAtividade(dados); }

function funilDeTemplate(chave) {
    const tpl = TEMPLATES_FUNIL[chave];
    if (!tpl) return null;
    const n = tpl.etapas.length;
    const etapas = tpl.etapas.map((nome, idx) => {
        let tipo = 'aberta';
        if (idx === n - 2) tipo = 'ganho';
        if (idx === n - 1) tipo = 'perdido';
        return normalizarEtapa({ nome, ordem: idx, tipo }, idx);
    });
    return normalizarFunil({ nome: tpl.nome, tipo: chave, mostrarValor: tpl.mostrarValor, etapas });
}

// ──────────────────────────────────────────────
//  VALIDAÇÃO — devolve array de mensagens (nunca lança)
// ──────────────────────────────────────────────

function validarNegocio(negocio, funil) {
    const erros = [];
    if (!ehObjeto(negocio)) {
        erros.push('Negócio deve ser um objeto válido');
        return erros;
    }
    if (!negocio.titulo || !String(negocio.titulo).trim()) {
        erros.push('Título é obrigatório');
    }
    const temItens = Array.isArray(negocio.itens) && negocio.itens.length > 0;
    const temValor = negocio.valor !== null && negocio.valor !== undefined && negocio.valor !== '';
    if (funil && funil.mostrarValor === false) {
        if (temValor && !temItens) erros.push('Este funil não utiliza valor monetário');
    } else if (temValor && !temItens) {
        const v = Number(negocio.valor);
        if (isNaN(v) || v < 0) erros.push('Valor deve ser um número não-negativo');
    }
    if (negocio.dataPrevisao && !/^\d{4}-\d{2}-\d{2}$/.test(negocio.dataPrevisao)) {
        erros.push('Data de previsão inválida (use formato YYYY-MM-DD)');
    }
    if (negocio.dataRecebimento && !/^\d{4}-\d{2}-\d{2}$/.test(negocio.dataRecebimento)) {
        erros.push('Data de recebimento inválida (use formato YYYY-MM-DD)');
    }
    return erros;
}

function validarAtividade(atividade) {
    const erros = [];
    if (!ehObjeto(atividade)) {
        erros.push('Atividade deve ser um objeto válido');
        return erros;
    }
    if (!atividade.assunto || !String(atividade.assunto).trim()) {
        erros.push('Assunto é obrigatório');
    }
    if (!atividade.negocioId) {
        erros.push('Atividade precisa estar vinculada a um negócio');
    }
    if (!atividade.data || !/^\d{4}-\d{2}-\d{2}$/.test(atividade.data)) {
        erros.push('Data da atividade é obrigatória (formato YYYY-MM-DD)');
    }
    if (atividade.horaInicio && !/^\d{2}:\d{2}$/.test(atividade.horaInicio)) {
        erros.push('Hora de início inválida (use HH:MM)');
    }
    if (atividade.horaFim && !/^\d{2}:\d{2}$/.test(atividade.horaFim)) {
        erros.push('Hora de fim inválida (use HH:MM)');
    }
    return erros;
}

const CrmModel = {
    CAMPOS_AUDITAVEIS_NEGOCIO,
    TIPOS_FUNIL,
    TIPOS_ETAPA,
    STATUS_NEGOCIO,
    TIPOS_ATIVIDADE,
    TEMPLATES_FUNIL,

    novoId,

    normalizarCrm,
    normalizarFunil,
    normalizarEtapa,
    normalizarNegocio,
    normalizarItem,
    normalizarAtividade,
    normalizarHistoricoItem,
    normalizarConfig,

    criarFunil,
    criarNegocio,
    criarAtividade,
    funilDeTemplate,

    validarNegocio,
    validarAtividade
};

if (typeof window !== 'undefined') {
    window.CrmModel = CrmModel;
}

if (typeof module !== 'undefined' && module.exports) {
    module.exports = CrmModel;
}
