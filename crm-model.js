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
const PRIORIDADES_ANOTACAO = ['baixa', 'media', 'alta', 'critico'];

// Ciclo de vida de uma demanda: chega ('recebida'), é assumida ('comigo'),
// depende de terceiros ('aguardando_terceiro'), volta para mim juntar as
// respostas ('consolidando') e é devolvida ao solicitante ('respondida').
const SITUACOES_ANOTACAO = ['recebida', 'comigo', 'aguardando_terceiro', 'consolidando', 'respondida'];
const STATUS_ENCAMINHAMENTO = ['pendente', 'respondido', 'cancelado'];

const CAMPOS_AUDITAVEIS_ANOTACAO = ['assunto', 'situacao', 'responsavel', 'prazo', 'prioridade', 'negocioId', 'funilId', 'etapaId', 'clienteId'];

// Tipos de atividade agendável (padrão Pipedrive). O ícone é um emoji para
// não depender de bibliotecas de ícone nos módulos puros.
const TIPOS_ATIVIDADE = {
    chamada: { rotulo: 'Chamada', icone: '📞' },
    reuniao: { rotulo: 'Reunião', icone: '👥' },
    tarefa: { rotulo: 'Tarefa', icone: '✔️' },
    prazo: { rotulo: 'Prazo', icone: '🚩' },
    email: { rotulo: 'E-mail', icone: '✉️' },
    viagem: { rotulo: 'Viagem', icone: '🧳' },
    ferias: { rotulo: 'Férias', icone: '🏖️' },
    recesso: { rotulo: 'Recesso', icone: '📴' },
    particular: { rotulo: 'Particular', icone: '👤' }
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

function normalizarTipoAtividade(tBruto, idx) {
    const t = ehObjeto(tBruto) ? tBruto : {};
    return {
        chave: (typeof t.chave === 'string' && t.chave.trim()) ? t.chave.trim() : novoId('tpa'),
        nome: (typeof t.nome === 'string' && t.nome.trim()) ? t.nome.trim() : ('Tipo ' + (idx + 1)),
        icone: (typeof t.icone === 'string' && t.icone) ? t.icone : '📌',
        ativo: t.ativo !== false,
        ordem: Number.isFinite(t.ordem) ? t.ordem : idx
    };
}

/**
 * Seed inicial dos tipos de atividade (primeira vez que o CRM é usado, antes
 * de qualquer customização) — a partir do conjunto padrão fixo TIPOS_ATIVIDADE.
 */
function tiposAtividadePadrao() {
    return Object.keys(TIPOS_ATIVIDADE).map((chave, idx) => ({
        chave,
        nome: TIPOS_ATIVIDADE[chave].rotulo,
        icone: TIPOS_ATIVIDADE[chave].icone,
        ativo: true,
        ordem: idx
    }));
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
    // O tipo é validado contra a lista de tipos conhecidos — força para 'tarefa' se inválido.
    const tipo = (typeof a.tipo === 'string' && TIPOS_ATIVIDADE[a.tipo]) ? a.tipo : 'tarefa';
    return {
        id: a.id || novoId('atv'),
        negocioId: a.negocioId || null,
        tipo,
        assunto: typeof a.assunto === 'string' ? a.assunto : '',
        descricao: typeof a.descricao === 'string' ? a.descricao : '',
        data: a.data || null,
        // Só usado por atividades importadas do Google que cobrem vários dias
        // (ex: reserva de hotel) — permite mesclar a célula nos calendários
        // Mês/Semana em vez de mostrar só no dia de início. null = evento de 1 dia só.
        dataFim: (typeof a.dataFim === 'string' && a.dataFim > a.data) ? a.dataFim : null,
        horaInicio: typeof a.horaInicio === 'string' ? a.horaInicio : '',
        horaFim: typeof a.horaFim === 'string' ? a.horaFim : '',
        feito: !!a.feito,
        feitoEm: a.feitoEm || null,
        googleEventId: typeof a.googleEventId === 'string' ? a.googleEventId : null,
        origemGoogle: !!a.origemGoogle,
        // Reservas de voo/hotel detectadas pelo Gmail: importadas só pra
        // aparecer no calendário — o app nunca escreve de volta nelas no
        // Google (ver google-calendar-sync.js: sincronizarAtividade/removerEvento).
        somenteLeituraGoogle: !!a.somenteLeituraGoogle,
        origemFeriado: !!a.origemFeriado,
        criadoEm: a.criadoEm || nowIso(),
        atualizadoEm: a.atualizadoEm || nowIso()
    };
}

/**
 * Encaminhamento: um pedido de informação feito a um terceiro dentro de uma
 * demanda. Vive embutido em `anotacao.encaminhamentos[]` (mesmo padrão de
 * `negocio.itens[]`) porque a mesma demanda costuma ser encaminhada a várias
 * áreas em paralelo e as respostas precisam ser consolidadas na demanda-mãe.
 */
function normalizarEncaminhamento(eBruto) {
    const e = ehObjeto(eBruto) ? eBruto : {};
    return {
        id: e.id || novoId('enc'),
        para: typeof e.para === 'string' ? e.para : '',
        canal: typeof e.canal === 'string' ? e.canal : '',
        oQuePedido: typeof e.oQuePedido === 'string' ? e.oQuePedido : '',
        numeroDocumentoEnvio: typeof e.numeroDocumentoEnvio === 'string' ? e.numeroDocumentoEnvio : '',
        dataEnvio: e.dataEnvio || null,
        prazoResposta: e.prazoResposta || null,
        status: STATUS_ENCAMINHAMENTO.indexOf(e.status) !== -1 ? e.status : 'pendente',
        dataResposta: e.dataResposta || null,
        resposta: typeof e.resposta === 'string' ? e.resposta : '',
        criadoEm: e.criadoEm || nowIso(),
        atualizadoEm: e.atualizadoEm || nowIso()
    };
}

/**
 * Converte `lembrarDiasAntes` para number|null. O campo era gravado como string
 * (e nunca lido); agora alimenta o semáforo de prazo em CrmCalculos.
 */
function normalizarDiasAntes(v) {
    if (v === null || v === undefined || v === '') return null;
    const n = Number(v);
    return (Number.isFinite(n) && n >= 0) ? Math.floor(n) : null;
}

/**
 * Anotação (demanda): unidade de trabalho da aba Relacionamento.
 *
 * O vínculo com negócio é OPCIONAL — uma demanda avulsa anda sozinha no funil
 * pelos seus próprios `funilId`/`etapaId`/`clienteId`. Quando `negocioId` está
 * preenchido, o store/UI derivam funil e cliente do negócio pai e os campos
 * próprios ficam ociosos.
 *
 * Migração idempotente de registros legados (feita aqui por não haver
 * migrations — ver normalizarCrm):
 *  - `situacao` ausente é inferida de `finalizado`/`destinatario`;
 *  - `destinatario` (texto livre, um único terceiro) vira o primeiro item de
 *    `encaminhamentos[]`. O campo permanece no schema apenas como legado, para
 *    não perder o dado histórico; saiu do formulário e da tabela.
 */
function normalizarAnotacao(aBruta) {
    const a = ehObjeto(aBruta) ? aBruta : {};
    const finalizado = !!a.finalizado;
    const destinatario = typeof a.destinatario === 'string' ? a.destinatario : '';

    let situacao = a.situacao;
    if (SITUACOES_ANOTACAO.indexOf(situacao) === -1) {
        situacao = finalizado ? 'respondida' : (destinatario ? 'aguardando_terceiro' : 'recebida');
    }

    let encaminhamentos;
    if (Array.isArray(a.encaminhamentos)) {
        encaminhamentos = a.encaminhamentos.map(normalizarEncaminhamento);
    } else if (destinatario) {
        encaminhamentos = [normalizarEncaminhamento({
            para: destinatario,
            status: finalizado ? 'respondido' : 'pendente',
            dataResposta: finalizado ? (a.dataConclusao || null) : null,
            criadoEm: a.criadoEm || nowIso()
        })];
    } else {
        encaminhamentos = [];
    }

    return {
        id: a.id || novoId('ant'),
        negocioId: a.negocioId || null,
        funilId: a.funilId || null,
        etapaId: a.etapaId || null,
        clienteId: (a.clienteId !== undefined && a.clienteId !== null) ? a.clienteId : null,
        anotacaoRelacionadaId: (a.anotacaoRelacionadaId && a.anotacaoRelacionadaId !== a.id) ? a.anotacaoRelacionadaId : null,
        assunto: typeof a.assunto === 'string' ? a.assunto : '',
        remetente: typeof a.remetente === 'string' ? a.remetente : '',
        origemDemanda: typeof a.origemDemanda === 'string' ? a.origemDemanda : '',
        dataSolicitacao: a.dataSolicitacao || null,
        tipoDoc: typeof a.tipoDoc === 'string' ? a.tipoDoc : '',
        numeroDocumento: typeof a.numeroDocumento === 'string' ? a.numeroDocumento : '',
        destinatario,
        acaoRealizar: typeof a.acaoRealizar === 'string' ? a.acaoRealizar : '',
        responsavel: typeof a.responsavel === 'string' ? a.responsavel : '',
        situacao,
        encaminhamentos,
        prioridade: PRIORIDADES_ANOTACAO.indexOf(a.prioridade) !== -1 ? a.prioridade : 'media',
        prazo: a.prazo || null,
        lembrarDiasAntes: normalizarDiasAntes(a.lembrarDiasAntes),
        tags: Array.isArray(a.tags) ? a.tags.slice() : [],
        observacoes: typeof a.observacoes === 'string' ? a.observacoes : '',
        finalizado: situacao === 'respondida',
        dataConclusao: a.dataConclusao || null,
        oQueFoiFeito: typeof a.oQueFoiFeito === 'string' ? a.oQueFoiFeito : '',
        respostaNumeroDocumento: typeof a.respostaNumeroDocumento === 'string' ? a.respostaNumeroDocumento : '',
        excluidoEm: a.excluidoEm || null,
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
        visao: ['kanban', 'lista', 'demandas', 'previsao', 'excluidos'].indexOf(c.visao) !== -1 ? c.visao : 'kanban',
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
 * Auto-cura de dados antigos: se qualquer demanda de uma thread vinculada
 * está "respondida", a thread inteira (mãe, filhas, netas, em qualquer
 * direção do vínculo) deve estar — cobre registros gravados antes do
 * cascateamento existir em CrmStore, sem depender de o usuário reabrir e
 * salvar cada uma manualmente. Roda a cada normalização (idempotente).
 */
function reconciliarThreadsRespondidas(anotacoes) {
    const porId = {};
    anotacoes.forEach(a => { porId[a.id] = a; });

    const filhasPorPai = {};
    anotacoes.forEach(a => {
        if (a.anotacaoRelacionadaId) {
            (filhasPorPai[a.anotacaoRelacionadaId] = filhasPorPai[a.anotacaoRelacionadaId] || []).push(a);
        }
    });

    function vizinhos(a) {
        const lista = [];
        if (a.anotacaoRelacionadaId && porId[a.anotacaoRelacionadaId]) lista.push(porId[a.anotacaoRelacionadaId]);
        (filhasPorPai[a.id] || []).forEach(f => lista.push(f));
        return lista;
    }

    const visitado = {};
    anotacoes.forEach(raiz => {
        if (visitado[raiz.id]) return;
        const fila = [raiz];
        const componente = [];
        visitado[raiz.id] = true;
        while (fila.length) {
            const atual = fila.pop();
            componente.push(atual);
            vizinhos(atual).forEach(v => {
                if (!visitado[v.id]) { visitado[v.id] = true; fila.push(v); }
            });
        }
        if (componente.length > 1 && componente.some(a => a.situacao === 'respondida')) {
            componente.forEach(a => {
                if (a.situacao !== 'respondida') {
                    a.situacao = 'respondida';
                    a.finalizado = true;
                    if (!a.dataConclusao) a.dataConclusao = (a.criadoEm || nowIso()).slice(0, 10);
                }
            });
        }
    });
}

/**
 * Mantém a etapa do negócio no Quadro coerente com o conjunto de demandas
 * vinculadas a ele: todas respondidas → etapa "ganho" do funil; qualquer uma
 * reaberta → volta para a etapa "aberta". Negócios sem demanda alguma ou já
 * marcados como perdidos ficam de fora (etapa continua manual).
 */
function reconciliarEtapasComDemandas(negocios, funis, anotacoes) {
    const funilPorId = {};
    funis.forEach(f => { funilPorId[f.id] = f; });

    const demandasPorNegocio = {};
    anotacoes.forEach(a => {
        if (a.negocioId && !a.excluidoEm) {
            (demandasPorNegocio[a.negocioId] = demandasPorNegocio[a.negocioId] || []).push(a);
        }
    });

    negocios.forEach(n => {
        const demandas = demandasPorNegocio[n.id];
        if (!demandas || !demandas.length) return;

        const funil = funilPorId[n.funilId];
        if (!funil) return;
        const etapaGanho = funil.etapas.filter(e => e.tipo === 'ganho')[0];
        const etapaAberta = funil.etapas.filter(e => e.tipo === 'aberta')[0] || funil.etapas[0];
        if (!etapaGanho || !etapaAberta) return;

        const etapaAtual = funil.etapas.filter(e => e.id === n.etapaId)[0];
        if (etapaAtual && etapaAtual.tipo === 'perdido') return;

        const todasRespondidas = demandas.every(a => a.situacao === 'respondida');

        if (todasRespondidas && (!etapaAtual || etapaAtual.tipo !== 'ganho')) {
            n.etapaId = etapaGanho.id;
            n.status = 'ganho';
            if (!n.dataFechamento) n.dataFechamento = nowIso().slice(0, 10);
        } else if (!todasRespondidas && etapaAtual && etapaAtual.tipo === 'ganho') {
            n.etapaId = etapaAberta.id;
            n.status = 'aberto';
            n.dataFechamento = null;
        }
    });
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
        .filter(a => !a.negocioId || idsNegocioValidos[a.negocioId]);

    // Demandas NÃO são mais descartadas por falta de negócio: o vínculo é
    // opcional. Um `negocioId` que aponta para negócio inexistente é zerado e a
    // demanda é realocada num funil válido — mesmo tratamento dado aos negócios
    // órfãos acima, em vez da perda silenciosa que havia antes.
    const anotacoesValidas = (Array.isArray(crm.anotacoes) ? crm.anotacoes : [])
        .map(normalizarAnotacao)
        .map(a => {
            if (a.negocioId && idsNegocioValidos[a.negocioId]) return a;
            let funilId = a.funilId;
            if (idsFunilValidos.indexOf(funilId) === -1) {
                funilId = idsFunilValidos[0] || null;
            }
            const funil = funilId ? funis[idsFunilValidos.indexOf(funilId)] : null;
            const idsEtapaDoFunil = funil ? funil.etapas.map(e => e.id) : [];
            let etapaId = a.etapaId;
            if (!etapaId || idsEtapaDoFunil.indexOf(etapaId) === -1) {
                etapaId = funilId ? (primeiraEtapaAbertaPorFunil[funilId] || null) : null;
            }
            return Object.assign({}, a, { negocioId: null, funilId, etapaId });
        });
    const idsAnotacaoValidos = {};
    anotacoesValidas.forEach(a => { idsAnotacaoValidos[a.id] = true; });
    const anotacoes = anotacoesValidas.map(a => (
        a.anotacaoRelacionadaId && idsAnotacaoValidos[a.anotacaoRelacionadaId]
            ? a
            : Object.assign({}, a, { anotacaoRelacionadaId: null })
    ));
    reconciliarThreadsRespondidas(anotacoes);
    reconciliarEtapasComDemandas(negocios, funis, anotacoes);

    const config = normalizarConfig(crm.config, funis);

    const tiposAtividadeBrutos = Array.isArray(crm.tiposAtividade) ? crm.tiposAtividade : null;
    const tiposAtividade = (tiposAtividadeBrutos && tiposAtividadeBrutos.length ? tiposAtividadeBrutos : tiposAtividadePadrao())
        .map(normalizarTipoAtividade)
        .sort((a, b) => a.ordem - b.ordem);

    return {
        versao: 1,
        funis,
        negocios,
        atividades,
        anotacoes,
        historico,
        tiposAtividade,
        config
    };
}

// ──────────────────────────────────────────────
//  FACTORIES
// ──────────────────────────────────────────────

function criarFunil(dados) { return normalizarFunil(dados); }
function criarNegocio(dados) { return normalizarNegocio(dados); }
function criarAtividade(dados) { return normalizarAtividade(dados); }
function criarAnotacao(dados) { return normalizarAnotacao(dados); }
function criarEncaminhamento(dados) { return normalizarEncaminhamento(dados); }

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

function validarAnotacao(anotacao) {
    const erros = [];
    if (!ehObjeto(anotacao)) {
        erros.push('Anotação deve ser um objeto válido');
        return erros;
    }
    // O vínculo com negócio é opcional, mas a demanda precisa de algum lugar no
    // mundo: ou um negócio, ou um funil próprio.
    if (!anotacao.negocioId && !anotacao.funilId) {
        erros.push('Demanda precisa estar vinculada a um negócio ou a um funil');
    }
    if (!anotacao.assunto || !String(anotacao.assunto).trim()) {
        erros.push('Assunto é obrigatório');
    }
    if (anotacao.situacao && SITUACOES_ANOTACAO.indexOf(anotacao.situacao) === -1) {
        erros.push('Situação inválida');
    }
    if (anotacao.prazo && !/^\d{4}-\d{2}-\d{2}$/.test(anotacao.prazo)) {
        erros.push('Prazo inválido (use formato YYYY-MM-DD)');
    }
    if (anotacao.dataSolicitacao && !/^\d{4}-\d{2}-\d{2}$/.test(anotacao.dataSolicitacao)) {
        erros.push('Data de solicitação inválida (use formato YYYY-MM-DD)');
    }
    (Array.isArray(anotacao.encaminhamentos) ? anotacao.encaminhamentos : []).forEach((e, i) => {
        const ref = 'Encaminhamento ' + (i + 1) + ': ';
        if (!e || !String(e.para || '').trim()) erros.push(ref + 'informe para quem foi encaminhado');
        ['dataEnvio', 'prazoResposta', 'dataResposta'].forEach(campo => {
            if (e && e[campo] && !/^\d{4}-\d{2}-\d{2}$/.test(e[campo])) {
                erros.push(ref + campo + ' inválida (use formato YYYY-MM-DD)');
            }
        });
    });
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
    CAMPOS_AUDITAVEIS_ANOTACAO,
    TIPOS_FUNIL,
    TIPOS_ETAPA,
    STATUS_NEGOCIO,
    TIPOS_ATIVIDADE,
    PRIORIDADES_ANOTACAO,
    SITUACOES_ANOTACAO,
    STATUS_ENCAMINHAMENTO,
    TEMPLATES_FUNIL,

    novoId,

    normalizarCrm,
    normalizarFunil,
    normalizarEtapa,
    normalizarNegocio,
    normalizarItem,
    normalizarAtividade,
    normalizarAnotacao,
    normalizarEncaminhamento,
    normalizarHistoricoItem,
    normalizarConfig,
    normalizarTipoAtividade,
    tiposAtividadePadrao,

    criarFunil,
    criarNegocio,
    criarAtividade,
    criarAnotacao,
    criarEncaminhamento,
    funilDeTemplate,

    validarNegocio,
    validarAtividade,
    validarAnotacao
};

if (typeof window !== 'undefined') {
    window.CrmModel = CrmModel;
}

if (typeof module !== 'undefined' && module.exports) {
    module.exports = CrmModel;
}
