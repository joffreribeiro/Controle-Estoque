/**
 * crm-calculos.js - Cálculos e transformações puras do módulo de Relacionamento (CRM)
 * Sem DOM, sem estado global — testável isoladamente com Vitest.
 */

// Faixa Unicode "Combining Diacritical Marks" (U+0300–U+036F), escrita via
// código de caractere para não depender de caracteres literais no arquivo-fonte.
const INICIO_DIACRITICOS = 0x0300;
const FIM_DIACRITICOS = 0x036f;

function normalizarParaBusca(str) {
    const decomposto = String(str || '').toLowerCase().normalize('NFD');
    let out = '';
    for (let i = 0; i < decomposto.length; i++) {
        const code = decomposto.charCodeAt(i);
        if (code < INICIO_DIACRITICOS || code > FIM_DIACRITICOS) {
            out += decomposto[i];
        }
    }
    return out;
}

/**
 * Soma os itens de produto de um negócio: Σ quantidade × precoUnit.
 */
function somarItens(itens) {
    return (itens || []).reduce((acc, it) => {
        const q = Number(it && it.quantidade);
        const p = Number(it && it.precoUnit);
        if (isNaN(q) || isNaN(p)) return acc;
        return acc + q * p;
    }, 0);
}

/**
 * Agrupa negócios por etapa, na ordem das próprias etapas.
 * Etapas sem negócios recebem array vazio; negócios com etapaId
 * sem correspondência nas etapas fornecidas são ignorados (órfãos).
 */
function agruparPorEtapa(negocios, etapas) {
    const out = {};
    (etapas || []).forEach(e => { out[e.id] = []; });
    (negocios || []).forEach(n => {
        if (n && n.etapaId && Object.prototype.hasOwnProperty.call(out, n.etapaId)) {
            out[n.etapaId].push(n);
        }
    });
    Object.keys(out).forEach(etapaId => {
        out[etapaId] = out[etapaId].slice().sort((a, b) => (a.ordem || 0) - (b.ordem || 0));
    });
    return out;
}

/**
 * Soma o campo `valor` de uma lista de negócios.
 * Aceita string numérica; ignora null/undefined/NaN sem lançar.
 */
function somarValor(negocios) {
    return (negocios || []).reduce((acc, n) => {
        const bruto = n && n.valor;
        if (bruto === null || bruto === undefined || bruto === '') return acc;
        const num = Number(bruto);
        return acc + (isNaN(num) ? 0 : num);
    }, 0);
}

/**
 * Resumo agregado de um funil: contagens por status, valor em aberto/ganho
 * e ticket médio (calculado só sobre negócios ganhos).
 */
function resumoFunil(negocios) {
    const lista = negocios || [];
    const abertos = lista.filter(n => n.status === 'aberto');
    const ganhos = lista.filter(n => n.status === 'ganho');
    const perdidos = lista.filter(n => n.status === 'perdido');
    const valorAberto = somarValor(abertos);
    const valorGanho = somarValor(ganhos);
    return {
        total: lista.length,
        abertos: abertos.length,
        ganhos: ganhos.length,
        perdidos: perdidos.length,
        valorAberto,
        valorGanho,
        ticketMedio: ganhos.length ? valorGanho / ganhos.length : 0
    };
}

/**
 * Filtra negócios por busca textual (sem acento, case-insensitive, no título),
 * responsável, status e funil. Qualquer filtro ausente/vazio é ignorado.
 */
function filtrarNegocios(negocios, filtros, extra) {
    const f = filtros || {};
    const busca = normalizarParaBusca(f.busca);
    const clientes = (extra && extra.clientes) || [];
    return (negocios || []).filter(n => {
        if (f.funilId && n.funilId !== f.funilId) return false;
        if (f.responsavel && n.responsavel !== f.responsavel) return false;
        if (f.status && n.status !== f.status) return false;
        if (f.origem && n.origem !== f.origem) return false;
        if (f.clienteId && n.clienteId !== f.clienteId) return false;
        if (f.dataRecebimentoInicio && String(n.dataRecebimento || '') < f.dataRecebimentoInicio) return false;
        if (f.dataRecebimentoFim && String(n.dataRecebimento || '') > f.dataRecebimentoFim) return false;
        if (f.dataPrevisaoInicio && String(n.dataPrevisao || '') < f.dataPrevisaoInicio) return false;
        if (f.dataPrevisaoFim && String(n.dataPrevisao || '') > f.dataPrevisaoFim) return false;
        if (f.mostrarConcluidas === false && n.status === 'ganho') return false;
        if (busca) {
            const cliente = clientes.find(c => c.id === n.clienteId);
            const alvo = normalizarParaBusca(n.titulo) + ' ' + normalizarParaBusca(cliente && cliente.nome);
            if (alvo.indexOf(busca) === -1) return false;
        }
        return true;
    });
}

/**
 * Ordena negócios sem mutar a entrada. Critério pode ser uma string simples
 * ('valor', 'previsao', 'atualizado', 'ordem') ou uma coluna da vista Lista
 * ('id', 'recebimento', 'titulo', 'cliente', 'origem', 'responsavel', 'prazo',
 * 'status', 'concluido', 'funil', 'ultima_atividade'), combinada com `direcao`
 * ('asc'|'desc', default 'asc' para colunas da Lista).
 */
function ordenarNegocios(negocios, criterio, direcao, extra) {
    const lista = (negocios || []).slice();
    if (criterio === 'valor' && !direcao) {
        return lista.sort((a, b) => (Number(b.valor) || 0) - (Number(a.valor) || 0));
    }
    if (criterio === 'previsao' && !direcao) {
        return lista.sort((a, b) => String(a.dataPrevisao || '9999-99-99').localeCompare(String(b.dataPrevisao || '9999-99-99')));
    }
    if (criterio === 'atualizado' && !direcao) {
        return lista.sort((a, b) => String(b.atualizadoEm || '').localeCompare(String(a.atualizadoEm || '')));
    }
    if (!criterio || criterio === 'ordem') {
        return lista.sort((a, b) => (a.ordem || 0) - (b.ordem || 0));
    }

    const clientes = (extra && extra.clientes) || [];
    const funis = (extra && extra.funis) || [];
    const atividades = (extra && extra.atividades) || [];
    const nomeCliente = id => { const c = clientes.find(x => x.id === id); return c ? c.nome : ''; };
    const nomeFunil = id => { const f = funis.find(x => x.id === id); return f ? f.nome : ''; };
    const ultimaAtividadeData = n => {
        const dasDele = atividades.filter(a => a.negocioId === n.id);
        const at = dasDele.length ? dasDele.slice().sort((a, b) => String(b.data || '').localeCompare(String(a.data || '')))[0] : null;
        return (at && at.data) || n.atualizadoEm || '';
    };

    const chave = n => {
        switch (criterio) {
            case 'id': return String(n.id || '');
            case 'recebimento': return String(n.dataRecebimento || '');
            case 'titulo': return normalizarParaBusca(n.titulo);
            case 'cliente': return normalizarParaBusca(nomeCliente(n.clienteId));
            case 'origem': return normalizarParaBusca(n.origem);
            case 'responsavel': return normalizarParaBusca(n.responsavel);
            case 'prazo': return String(n.dataPrevisao || '');
            case 'status': return String(n.status || '');
            case 'concluido': return n.status === 'ganho' ? 1 : 0;
            case 'valor': return Number(n.valor) || 0;
            case 'funil': return normalizarParaBusca(nomeFunil(n.funilId));
            case 'ultima_atividade': return String(ultimaAtividadeData(n));
            default: return String(n.ordem || 0);
        }
    };

    const dir = direcao === 'desc' ? -1 : 1;
    return lista.sort((a, b) => {
        const va = chave(a), vb = chave(b);
        if (va < vb) return -1 * dir;
        if (va > vb) return 1 * dir;
        return 0;
    });
}

// ──────────────────────────────────────────────
//  DEMANDAS (anotações) — derivações do ciclo encaminhar → aguardar → consolidar
// ──────────────────────────────────────────────

/**
 * Diferença em dias COM SINAL: negativo = a data já passou.
 * `diasEntre` trunca em zero, o que serve para idade mas não para prazo.
 */
function diasAte(dataIso, hoje) {
    if (!dataIso) return null;
    const alvo = new Date(String(dataIso).slice(0, 10) + 'T00:00:00Z');
    const ref = new Date(hojeIso(hoje) + 'T00:00:00Z');
    if (isNaN(alvo) || isNaN(ref)) return null;
    return Math.round((alvo - ref) / 86400000);
}

function diasDesde(dataIso, hoje) {
    const d = diasAte(dataIso, hoje);
    return d === null ? null : -d;
}

/**
 * Fotografia dos pedidos feitos a terceiros numa demanda. É o que responde
 * "com quem isso está parado, e há quanto tempo".
 */
function resumoEncaminhamentos(anotacao, hoje) {
    const lista = (anotacao && Array.isArray(anotacao.encaminhamentos)) ? anotacao.encaminhamentos : [];
    const pendentes = lista.filter(e => e.status === 'pendente');
    const respondidos = lista.filter(e => e.status === 'respondido');

    let prazoTerceiroMaisProximo = null;
    let algumVencido = false;
    let maiorEspera = null;

    pendentes.forEach(e => {
        if (e.prazoResposta) {
            if (!prazoTerceiroMaisProximo || e.prazoResposta < prazoTerceiroMaisProximo) {
                prazoTerceiroMaisProximo = e.prazoResposta;
            }
            if (diasAte(e.prazoResposta, hoje) < 0) algumVencido = true;
        }
        const espera = diasDesde(e.dataEnvio, hoje);
        if (espera !== null && (maiorEspera === null || espera > maiorEspera)) maiorEspera = espera;
    });

    return {
        total: lista.length,
        pendentes: pendentes.length,
        respondidos: respondidos.length,
        comQuem: pendentes.map(e => e.para).filter(Boolean),
        prazoTerceiroMaisProximo,
        algumVencido,
        diasAguardando: maiorEspera
    };
}

/**
 * Situação que os encaminhamentos sugerem. Não decide sozinha: a UI/store usa
 * como sugestão e o usuário mantém a palavra final. Uma demanda já respondida
 * nunca é puxada de volta.
 */
function situacaoSugerida(anotacao, hoje) {
    const atual = (anotacao && anotacao.situacao) || 'recebida';
    if (atual === 'respondida') return atual;
    const r = resumoEncaminhamentos(anotacao, hoje);
    if (r.pendentes > 0) return 'aguardando_terceiro';
    if (r.respondidos > 0) return 'consolidando';
    return atual === 'aguardando_terceiro' ? 'comigo' : atual;
}

/**
 * Semáforo de prazo da demanda. É o único consumidor de `lembrarDiasAntes` —
 * até então o campo era gravado e nunca lido.
 */
function semaforoPrazo(anotacao, hoje) {
    if (!anotacao || anotacao.finalizado || !anotacao.prazo) return 'ok';
    const dias = diasAte(anotacao.prazo, hoje);
    if (dias === null) return 'ok';
    if (dias < 0) return 'vencida';
    if (dias === 0) return 'hoje';
    const antecedencia = Number.isFinite(anotacao.lembrarDiasAntes) ? anotacao.lembrarDiasAntes : null;
    if (antecedencia !== null && dias <= antecedencia) return 'alerta';
    return 'ok';
}

/**
 * Buckets das abas da Caixa de Demandas. 'recebida' e 'comigo' caem juntas em
 * `comigo` porque, do ponto de vista da fila de trabalho, ambas dependem de mim.
 */
function agruparPorSituacao(anotacoes) {
    const out = { comigo: [], aguardando: [], consolidar: [], concluidas: [] };
    (anotacoes || []).forEach(a => {
        switch (a.situacao) {
            case 'aguardando_terceiro': out.aguardando.push(a); break;
            case 'consolidando': out.consolidar.push(a); break;
            case 'respondida': out.concluidas.push(a); break;
            default: out.comigo.push(a);
        }
    });
    return out;
}

/**
 * Resolve funil/cliente de uma demanda: pelo negócio pai quando existe, senão
 * pelos campos próprios da demanda avulsa.
 */
function vinculoDaAnotacao(anotacao, negocioPorId) {
    const negocio = anotacao.negocioId ? negocioPorId[anotacao.negocioId] : null;
    return negocio
        ? { funilId: negocio.funilId, clienteId: negocio.clienteId }
        : { funilId: anotacao.funilId || null, clienteId: anotacao.clienteId || null };
}

/**
 * Filtra demandas. `extra.negocios` resolve funil/cliente das que têm negócio
 * pai; as avulsas usam os próprios campos.
 */
function filtrarAnotacoes(anotacoes, filtros, extra) {
    const f = filtros || {};
    const busca = normalizarParaBusca(f.busca);
    const negocios = (extra && extra.negocios) || [];
    const hoje = extra && extra.hoje;
    const negocioPorId = {};
    negocios.forEach(n => { negocioPorId[n.id] = n; });

    return (anotacoes || []).filter(a => {
        if (a.excluidoEm) return false;
        const vinculo = vinculoDaAnotacao(a, negocioPorId);
        if (f.funilId && vinculo.funilId !== f.funilId) return false;
        if (f.clienteId && vinculo.clienteId !== f.clienteId) return false;
        if (f.prioridade && a.prioridade !== f.prioridade) return false;
        if (f.situacao && a.situacao !== f.situacao) return false;
        if (f.responsavel && a.responsavel !== f.responsavel) return false;
        if (f.origemDemanda && a.origemDemanda !== f.origemDemanda) return false;
        if (f.prazoInicio && String(a.prazo || '') < f.prazoInicio) return false;
        if (f.prazoFim && String(a.prazo || '') > f.prazoFim) return false;
        if (f.dataSolicitacaoInicio && String(a.dataSolicitacao || '') < f.dataSolicitacaoInicio) return false;
        if (f.dataSolicitacaoFim && String(a.dataSolicitacao || '') > f.dataSolicitacaoFim) return false;
        if (f.mostrarFinalizadas === false && a.finalizado) return false;

        const encaminhamentos = Array.isArray(a.encaminhamentos) ? a.encaminhamentos : [];
        if (f.encaminhadoPara) {
            if (!encaminhamentos.some(e => e.para === f.encaminhadoPara)) return false;
        }
        if (f.somenteVencidas && semaforoPrazo(a, hoje) !== 'vencida') return false;
        if (f.somenteAtrasoTerceiro && !resumoEncaminhamentos(a, hoje).algumVencido) return false;

        if (busca) {
            const dosEncaminhamentos = encaminhamentos
                .map(e => normalizarParaBusca(e.para) + ' ' + normalizarParaBusca(e.oQuePedido) + ' ' + normalizarParaBusca(e.resposta))
                .join(' ');
            const alvo = normalizarParaBusca(a.assunto) + ' ' + normalizarParaBusca(a.remetente) + ' ' +
                normalizarParaBusca(a.acaoRealizar) + ' ' + normalizarParaBusca(a.observacoes) + ' ' + dosEncaminhamentos;
            if (alvo.indexOf(busca) === -1) return false;
        }
        return true;
    });
}

const PRIORIDADE_PESO = { baixa: 0, media: 1, alta: 2, critico: 3 };
const SITUACAO_PESO = { recebida: 0, comigo: 1, aguardando_terceiro: 2, consolidando: 3, respondida: 4 };

/**
 * Ordena anotações sem mutar a entrada. Critérios: 'assunto', 'remetente',
 * 'origemDemanda', 'dataSolicitacao', 'tipoDoc', 'numeroDocumento',
 * 'destinatario', 'acaoRealizar', 'prioridade', 'prazo', 'finalizado',
 * 'dataConclusao', 'situacao', 'responsavel', 'diasAguardando', 'funil',
 * 'cliente' (do negócio pai, ou da própria demanda quando avulsa).
 * `direcao`: 'asc'|'desc'.
 */
function ordenarAnotacoes(anotacoes, criterio, direcao, extra) {
    const lista = (anotacoes || []).slice();
    const negocios = (extra && extra.negocios) || [];
    const clientes = (extra && extra.clientes) || [];
    const funis = (extra && extra.funis) || [];
    const negocioPorId = {};
    negocios.forEach(n => { negocioPorId[n.id] = n; });
    const nomeCliente = id => { const c = clientes.find(x => x.id === id); return c ? c.nome : ''; };
    const nomeFunil = id => { const f = funis.find(x => x.id === id); return f ? f.nome : ''; };

    const chave = a => {
        const vinculo = vinculoDaAnotacao(a, negocioPorId);
        switch (criterio) {
            case 'funil': return normalizarParaBusca(nomeFunil(vinculo.funilId));
            case 'cliente': return normalizarParaBusca(nomeCliente(vinculo.clienteId));
            case 'situacao': return SITUACAO_PESO[a.situacao] || 0;
            case 'responsavel': return normalizarParaBusca(a.responsavel);
            case 'diasAguardando': {
                const d = resumoEncaminhamentos(a, extra && extra.hoje).diasAguardando;
                return d === null ? -1 : d;
            }
            case 'assunto': return normalizarParaBusca(a.assunto);
            case 'remetente': return normalizarParaBusca(a.remetente);
            case 'origemDemanda': return normalizarParaBusca(a.origemDemanda);
            case 'dataSolicitacao': return String(a.dataSolicitacao || '');
            case 'tipoDoc': return normalizarParaBusca(a.tipoDoc);
            case 'numeroDocumento': return normalizarParaBusca(a.numeroDocumento);
            case 'destinatario': return normalizarParaBusca(a.destinatario);
            case 'acaoRealizar': return normalizarParaBusca(a.acaoRealizar);
            case 'prioridade': return PRIORIDADE_PESO[a.prioridade] || 0;
            case 'prazo': return String(a.prazo || '');
            case 'finalizado': return a.finalizado ? 1 : 0;
            case 'dataConclusao': return String(a.dataConclusao || '');
            default: return String(a.criadoEm || '');
        }
    };

    const dir = direcao === 'desc' ? -1 : 1;
    return lista.sort((a, b) => {
        const va = chave(a), vb = chave(b);
        if (va < vb) return -1 * dir;
        if (va > vb) return 1 * dir;
        return 0;
    });
}

/**
 * Calcula a nova ordenação de uma etapa após mover um negócio para dentro dela
 * (de outra etapa ou reordenando na mesma). Pura: não muta `negocios`.
 * Devolve pares {id, ordem} já densos (0..n-1) para a etapa de destino.
 */
function reordenarNaEtapa(negocios, etapaId, idMovido, indice) {
    const daEtapaSemMovido = (negocios || [])
        .filter(n => n.etapaId === etapaId && n.id !== idMovido)
        .slice()
        .sort((a, b) => (a.ordem || 0) - (b.ordem || 0));

    const movido = (negocios || []).find(n => n.id === idMovido) || { id: idMovido };
    const listaFinal = daEtapaSemMovido.slice();
    const pos = Math.max(0, Math.min(indice == null ? listaFinal.length : indice, listaFinal.length));
    listaFinal.splice(pos, 0, movido);

    return listaFinal.map((n, idx) => ({ id: n.id, ordem: idx }));
}

/**
 * Taxa de conversão (ganhos / (ganhos + perdidos)). 0 quando não há fechados.
 */
function taxaConversao(negocios) {
    const lista = negocios || [];
    const ganhos = lista.filter(n => n.status === 'ganho').length;
    const perdidos = lista.filter(n => n.status === 'perdido').length;
    const total = ganhos + perdidos;
    return total ? ganhos / total : 0;
}

/**
 * Formata um valor monetário em pt-BR. Cai num fallback manual se a moeda
 * informada não for reconhecida pelo Intl.
 */
function formatarMoeda(valor, moeda) {
    const num = Number(valor) || 0;
    try {
        return new Intl.NumberFormat('pt-BR', { style: 'currency', currency: moeda || 'BRL' }).format(num);
    } catch (_) {
        return 'R$ ' + num.toFixed(2).replace('.', ',');
    }
}

function negociosDoCliente(negocios, clienteId) {
    return (negocios || []).filter(n => n.clienteId === clienteId);
}

// ──────────────────────────────────────────────
//  ATIVIDADES E MÉTRICAS DERIVADAS (padrão Pipedrive)
//  Todas aceitam `hoje` (YYYY-MM-DD) como parâmetro para serem testáveis.
// ──────────────────────────────────────────────

function hojeIso(hoje) {
    return hoje || new Date().toISOString().slice(0, 10);
}

function diasEntre(isoAntigo, isoNovo) {
    const a = new Date(String(isoAntigo).slice(0, 10) + 'T00:00:00Z');
    const b = new Date(String(isoNovo).slice(0, 10) + 'T00:00:00Z');
    if (isNaN(a) || isNaN(b)) return 0;
    return Math.max(0, Math.round((b - a) / 86400000));
}

function atividadesPendentesDe(atividades, negocioId) {
    return (atividades || [])
        .filter(a => a.negocioId === negocioId && !a.feito)
        .slice()
        .sort((a, b) => String(a.data || '9999').localeCompare(String(b.data || '9999'))
            || String(a.horaInicio || '99:99').localeCompare(String(b.horaInicio || '99:99')));
}

function proximaAtividade(atividades, negocioId) {
    return atividadesPendentesDe(atividades, negocioId)[0] || null;
}

function temAtividadePendente(atividades, negocioId) {
    return !!proximaAtividade(atividades, negocioId);
}

function diasNaEtapa(historico, negocio, hoje) {
    const mudancas = (historico || [])
        .filter(h => h.entidade === 'negocio' && h.entidadeId === negocio.id && h.tipo === 'etapa')
        .sort((a, b) => String(b.criadoEm || '').localeCompare(String(a.criadoEm || '')));
    const desde = (mudancas[0] && mudancas[0].criadoEm) || negocio.criadoEm;
    return diasEntre(desde, hojeIso(hoje));
}

function idadeEmDias(negocio, hoje) {
    return diasEntre(negocio.criadoEm, hojeIso(hoje));
}

function diasInativo(negocio, atividades, hoje) {
    let ultimo = negocio.atualizadoEm || negocio.criadoEm;
    (atividades || []).forEach(a => {
        if (a.negocioId === negocio.id && a.feitoEm && String(a.feitoEm) > String(ultimo)) {
            ultimo = a.feitoEm;
        }
    });
    return diasEntre(ultimo, hojeIso(hoje));
}

/**
 * Agrupa negócios pelo mês da data de fechamento esperada (visão Previsão).
 * Devolve lista ordenada de { mes: 'YYYY-MM'|null, negocios: [...] };
 * sem data entra no grupo mes:null, sempre por último.
 */
function agruparPorMesFechamento(negocios) {
    const porMes = {};
    const semData = [];
    (negocios || []).forEach(n => {
        const mes = (n.dataPrevisao && /^\d{4}-\d{2}/.test(n.dataPrevisao)) ? n.dataPrevisao.slice(0, 7) : null;
        if (!mes) { semData.push(n); return; }
        (porMes[mes] = porMes[mes] || []).push(n);
    });
    const grupos = Object.keys(porMes).sort().map(mes => ({ mes, negocios: porMes[mes] }));
    if (semData.length) grupos.push({ mes: null, negocios: semData });
    return grupos;
}

// ──────────────────────────────────────────────
//  CALENDÁRIO DE ATIVIDADES (visão global, todas os negócios)
// ──────────────────────────────────────────────

/**
 * Domingo (YYYY-MM-DD) da semana que contém `dataIso`. Semana dom→sáb.
 */
function inicioSemana(dataIso) {
    const d = new Date(String(dataIso).slice(0, 10) + 'T00:00:00Z');
    if (isNaN(d)) return dataIso;
    d.setUTCDate(d.getUTCDate() - d.getUTCDay());
    return d.toISOString().slice(0, 10);
}

function somarDias(dataIso, dias) {
    const d = new Date(String(dataIso).slice(0, 10) + 'T00:00:00Z');
    if (isNaN(d)) return dataIso;
    d.setUTCDate(d.getUTCDate() + dias);
    return d.toISOString().slice(0, 10);
}

/**
 * Filtra atividades por período nomeado (padrão Pipedrive), relativo a `hoje`.
 * Períodos: 'todos' | 'paraFazer' | 'vencido' | 'hoje' | 'amanha' | 'semana' | 'proximaSemana'.
 */
function filtrarAtividadesPeriodo(atividades, periodo, hoje) {
    const h = hojeIso(hoje);
    const lista = atividades || [];
    switch (periodo) {
        case 'paraFazer': return lista.filter(a => !a.feito);
        case 'vencido': return lista.filter(a => !a.feito && a.data && a.data < h);
        case 'hoje': return lista.filter(a => a.data === h);
        case 'amanha': return lista.filter(a => a.data === somarDias(h, 1));
        case 'semana': {
            const ini = inicioSemana(h), fim = somarDias(ini, 6);
            return lista.filter(a => a.data && a.data >= ini && a.data <= fim);
        }
        case 'proximaSemana': {
            const ini = somarDias(inicioSemana(h), 7), fim = somarDias(ini, 6);
            return lista.filter(a => a.data && a.data >= ini && a.data <= fim);
        }
        default: return lista.slice();
    }
}

/**
 * Filtra atividades por texto livre (assunto) e, via `extra`, pelo negócio/cliente pai.
 */
function filtrarAtividadesBusca(atividades, busca, extra) {
    const termo = normalizarParaBusca(busca);
    if (!termo) return atividades || [];
    const negocios = (extra && extra.negocios) || [];
    const clientes = (extra && extra.clientes) || [];
    const negocioPorId = {};
    negocios.forEach(n => { negocioPorId[n.id] = n; });
    const clientePorId = {};
    clientes.forEach(c => { clientePorId[c.id] = c; });

    return (atividades || []).filter(a => {
        const negocio = negocioPorId[a.negocioId];
        const cliente = negocio ? clientePorId[negocio.clienteId] : null;
        const alvo = normalizarParaBusca(a.assunto) + ' ' +
            normalizarParaBusca(negocio ? negocio.titulo : '') + ' ' +
            normalizarParaBusca(cliente ? cliente.nome : '');
        return alvo.indexOf(termo) !== -1;
    });
}

/**
 * Ordena atividades por data + hora de início (ascendente); sem data vai por último.
 */
function ordenarAtividadesPorData(atividades) {
    return (atividades || []).slice().sort((a, b) =>
        String(a.data || '9999-99-99').localeCompare(String(b.data || '9999-99-99')) ||
        String(a.horaInicio || '99:99').localeCompare(String(b.horaInicio || '99:99'))
    );
}

/**
 * Duração formatada 'Hh Mmin' entre horaInicio e horaFim (HH:MM); null se não houver ambos.
 */
function duracaoAtividade(horaInicio, horaFim) {
    if (!horaInicio || !horaFim) return null;
    const [h1, m1] = horaInicio.split(':').map(Number);
    const [h2, m2] = horaFim.split(':').map(Number);
    if ([h1, m1, h2, m2].some(n => isNaN(n))) return null;
    let min = (h2 * 60 + m2) - (h1 * 60 + m1);
    if (min < 0) return null;
    const hh = Math.floor(min / 60), mm = min % 60;
    if (hh && mm) return hh + 'h ' + mm + 'min';
    if (hh) return hh + 'h';
    return mm + 'min';
}

/**
 * Agrupa atividades de uma semana (dom→sáb) por dia. Devolve mapa 'YYYY-MM-DD' -> [atividades],
 * sempre com as 7 chaves presentes (mesmo vazias), ordenadas por hora dentro do dia.
 */
function agruparAtividadesPorDiaDaSemana(atividades, domingoIso) {
    const mapa = {};
    for (let i = 0; i < 7; i++) mapa[somarDias(domingoIso, i)] = [];
    (atividades || []).forEach(a => {
        if (Object.prototype.hasOwnProperty.call(mapa, a.data)) mapa[a.data].push(a);
    });
    Object.keys(mapa).forEach(k => {
        mapa[k].sort((a, b) => String(a.horaInicio || '99:99').localeCompare(String(b.horaInicio || '99:99')));
    });
    return mapa;
}

/**
 * Data da Páscoa (domingo) de um ano, calculada localmente — algoritmo
 * "anônimo gregoriano" (Meeus/Jones/Butcher). Usada para derivar os feriados
 * móveis (Carnaval, Sexta-feira Santa, Corpus Christi).
 */
function calcularPascoa(ano) {
    const a = ano % 19;
    const b = Math.floor(ano / 100);
    const c = ano % 100;
    const d = Math.floor(b / 4);
    const e = b % 4;
    const f = Math.floor((b + 8) / 25);
    const g = Math.floor((b - f + 1) / 3);
    const h = (19 * a + b - d - g + 15) % 30;
    const i = Math.floor(c / 4);
    const k = c % 4;
    const l = (32 + 2 * e + 2 * i - h - k) % 7;
    const m = Math.floor((a + 11 * h + 22 * l) / 451);
    const mes = Math.floor((h + l - 7 * m + 114) / 31);
    const dia = ((h + l - 7 * m + 114) % 31) + 1;
    return new Date(Date.UTC(ano, mes - 1, dia));
}

/**
 * Feriados nacionais do Brasil de um ano (fixos + móveis, calculados a partir
 * da Páscoa). Não inclui feriados estaduais/municipais. Devolve lista
 * ordenada de { data: 'YYYY-MM-DD', nome }.
 */
function feriadosNacionais(ano) {
    const pascoa = calcularPascoa(ano);
    const comData = (offsetDias) => {
        const d = new Date(pascoa);
        d.setUTCDate(d.getUTCDate() + offsetDias);
        return d.toISOString().slice(0, 10);
    };

    const fixos = [
        ['01-01', 'Confraternização Universal'],
        ['04-21', 'Tiradentes'],
        ['05-01', 'Dia do Trabalho'],
        ['09-07', 'Independência do Brasil'],
        ['10-12', 'Nossa Senhora Aparecida'],
        ['11-02', 'Finados'],
        ['11-15', 'Proclamação da República'],
        ['11-20', 'Dia Nacional de Zumbi e da Consciência Negra'],
        ['12-25', 'Natal']
    ].map(([md, nome]) => ({ data: ano + '-' + md, nome }));

    const moveis = [
        { data: comData(-47), nome: 'Carnaval' },
        { data: comData(-2), nome: 'Sexta-feira Santa' },
        { data: comData(0), nome: 'Páscoa' },
        { data: comData(60), nome: 'Corpus Christi' }
    ];

    return fixos.concat(moveis).sort((x, y) => x.data.localeCompare(y.data));
}

/**
 * Domingo (YYYY-MM-DD) da primeira semana exibida na grade mensal (a semana
 * que contém o dia 1 do mês de `mesRefIso`, seja qual for o dia informado).
 */
function inicioGradeMes(mesRefIso) {
    const d = new Date(String(mesRefIso).slice(0, 7) + '-01T00:00:00Z');
    if (isNaN(d)) return mesRefIso;
    d.setUTCDate(d.getUTCDate() - d.getUTCDay());
    return d.toISOString().slice(0, 10);
}

/**
 * Soma/subtrai meses a uma referência 'YYYY-MM' (ou 'YYYY-MM-DD'), sempre
 * devolvendo o dia 1 do mês resultante ('YYYY-MM-01').
 */
function somarMeses(mesRefIso, qtd) {
    const d = new Date(String(mesRefIso).slice(0, 7) + '-01T00:00:00Z');
    if (isNaN(d)) return mesRefIso;
    d.setUTCMonth(d.getUTCMonth() + qtd);
    return d.toISOString().slice(0, 7) + '-01';
}

/**
 * Agrupa atividades numa grade mensal de 6 semanas (42 dias, dom→sáb),
 * começando no domingo que contém o dia 1 do mês de `mesRefIso`. Devolve
 * mapa 'YYYY-MM-DD' -> [atividades], sempre com as 42 chaves presentes.
 */
function agruparAtividadesPorGradeMes(atividades, mesRefIso) {
    const inicio = inicioGradeMes(mesRefIso);
    const mapa = {};
    for (let i = 0; i < 42; i++) mapa[somarDias(inicio, i)] = [];
    (atividades || []).forEach(a => {
        if (Object.prototype.hasOwnProperty.call(mapa, a.data)) mapa[a.data].push(a);
    });
    Object.keys(mapa).forEach(k => {
        mapa[k].sort((a, b) => String(a.horaInicio || '99:99').localeCompare(String(b.horaInicio || '99:99')));
    });
    return mapa;
}

/**
 * Itens de histórico de uma entidade específica, mais recentes primeiro.
 */
function timelineDe(historico, entidade, entidadeId) {
    return (historico || [])
        .filter(h => h.entidade === entidade && h.entidadeId === entidadeId)
        .slice()
        .sort((a, b) => String(b.criadoEm || '').localeCompare(String(a.criadoEm || '')));
}

const CrmCalculos = {
    somarItens,
    agruparPorEtapa,
    somarValor,
    resumoFunil,
    filtrarNegocios,
    ordenarNegocios,
    filtrarAnotacoes,
    ordenarAnotacoes,
    diasAte,
    diasDesde,
    resumoEncaminhamentos,
    situacaoSugerida,
    semaforoPrazo,
    agruparPorSituacao,
    vinculoDaAnotacao,
    reordenarNaEtapa,
    taxaConversao,
    formatarMoeda,
    negociosDoCliente,
    timelineDe,

    hojeIso,
    diasEntre,
    atividadesPendentesDe,
    proximaAtividade,
    temAtividadePendente,
    diasNaEtapa,
    idadeEmDias,
    diasInativo,
    agruparPorMesFechamento,

    inicioSemana,
    somarDias,
    filtrarAtividadesPeriodo,
    filtrarAtividadesBusca,
    ordenarAtividadesPorData,
    duracaoAtividade,
    agruparAtividadesPorDiaDaSemana,
    inicioGradeMes,
    somarMeses,
    agruparAtividadesPorGradeMes,
    calcularPascoa,
    feriadosNacionais
};

if (typeof window !== 'undefined') {
    window.CrmCalculos = CrmCalculos;
}

if (typeof module !== 'undefined' && module.exports) {
    module.exports = CrmCalculos;
}
