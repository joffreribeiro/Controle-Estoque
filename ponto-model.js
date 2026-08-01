/**
 * ponto-model.js - Forma dos dados do módulo de Ponto (jornada/RH) do Controle-Estoque
 * Funções puras: factories, normalização/migração e validação.
 * Sem DOM, sem estado global — testável isoladamente com Vitest.
 *
 * Schema e nomes de campo portados quase literalmente do sistema Ponto
 * (https://github.com/joffreribeiro/Ponto, storage.js/validators.js) de propósito:
 * a Fase 2 (migração) faz um pass-through dos dados brutos do Firestore do Ponto
 * direto por normalizarPonto(), e a Fase 3 valida os números calculados contra o
 * que o Ponto original mostrava — divergir nomes de campo sem necessidade
 * dificultaria as duas coisas.
 */

const TIPOS_ATESTADO = [
    'afastamento',
    'comparecimento_matutino', 'comparecimento_vespertino',
    'abono_matutino', 'abono_vespertino', 'abono_dia_todo',
    'abono_acordo_matutino', 'abono_acordo_vespertino', 'abono_acordo_dia_todo',
    'pagar_hora_acordo_matutino', 'pagar_hora_acordo_vespertino', 'pagar_hora_acordo_dia_todo'
];

const TIPOS_EVENTO_PADRAO = [
    { id: 'feriado', nome: 'Feriado', cor: '#dc2626' },
    { id: 'ferias', nome: 'Férias', cor: '#d97706' },
    { id: 'afastamento', nome: 'Afastamento', cor: '#0891b2' },
    { id: 'viagem', nome: 'Viagem', cor: '#7c3aed' },
    { id: 'abono_acordo', nome: 'Abono (acordo)', cor: '#059669' },
    { id: 'compensar_acordo', nome: 'Pagar Hora (acordo)', cor: '#db2777' },
    { id: 'outro', nome: 'Outro', cor: '#64748b' }
];

function novoId(prefixo) {
    return prefixo + '_' + Date.now().toString(36) + '_' + Math.random().toString(36).substr(2, 6);
}

function ehObjeto(v) {
    return !!v && typeof v === 'object' && !Array.isArray(v);
}

// Aceita "YYYY-MM-DD" ou "DD/MM/YYYY" e normaliza para ISO. null se inválida.
function normalizarDataIso(v) {
    if (!v || typeof v !== 'string') return null;
    if (/^\d{4}-\d{2}-\d{2}$/.test(v)) {
        return isNaN(Date.parse(v)) ? null : v;
    }
    const m = v.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
    if (m) {
        const [, d, mo, y] = m;
        const dt = new Date(Number(y), Number(mo) - 1, Number(d));
        const valida = dt.getFullYear() === Number(y) && dt.getMonth() === Number(mo) - 1 && dt.getDate() === Number(d);
        return valida ? (y + '-' + mo + '-' + d) : null;
    }
    return null;
}

function horaValida(v) {
    if (!v || typeof v !== 'string') return false;
    const m = v.match(/^(\d{1,2}):(\d{2})$/);
    if (!m) return false;
    const h = Number(m[1]), min = Number(m[2]);
    return h >= 0 && h < 24 && min >= 0 && min < 60;
}

// ──────────────────────────────────────────────
//  NORMALIZAÇÃO — uma função por entidade, todas puras e idempotentes
// ──────────────────────────────────────────────

function normalizarRegistro(rBruto) {
    const r = ehObjeto(rBruto) ? rBruto : {};
    return {
        data: normalizarDataIso(r.data),
        entrada: horaValida(r.entrada) ? r.entrada : null,
        saidaAlmoco: horaValida(r.saidaAlmoco) ? r.saidaAlmoco : null,
        retornoAlmoco: horaValida(r.retornoAlmoco) ? r.retornoAlmoco : null,
        saida: horaValida(r.saida) ? r.saida : null,
        observacoes: typeof r.observacoes === 'string' ? r.observacoes : '',
        periodoEvento: typeof r.periodoEvento === 'string' ? r.periodoEvento : null,
        tipoEventoRegistro: typeof r.tipoEventoRegistro === 'string' ? r.tipoEventoRegistro : null,
        createLinkedEvent: !!r.createLinkedEvent,
        tipoAtestado: TIPOS_ATESTADO.indexOf(r.tipoAtestado) !== -1 ? r.tipoAtestado : null
    };
}

function normalizarPeriodoAcordo(pBruto) {
    const p = ehObjeto(pBruto) ? pBruto : {};
    return {
        inicio: normalizarDataIso(p.inicio),
        fim: normalizarDataIso(p.fim),
        minutosExtras: Number.isFinite(Number(p.minutosExtras)) ? Number(p.minutosExtras) : 0
    };
}

function normalizarRegraHorario(rBruto) {
    const r = ehObjeto(rBruto) ? rBruto : {};
    return {
        inicio: normalizarDataIso(r.inicio),
        fim: normalizarDataIso(r.fim),
        almocoMin: Number.isFinite(Number(r.almocoMin)) ? Number(r.almocoMin) : 60,
        tolAlmoco: Number.isFinite(Number(r.tolAlmoco)) ? Number(r.tolAlmoco) : 5,
        tolSaida: Number.isFinite(Number(r.tolSaida)) ? Number(r.tolSaida) : 5,
        minutosExtras: Number.isFinite(Number(r.minutosExtras)) ? Number(r.minutosExtras) : 0,
        vale: Number.isFinite(Number(r.vale)) ? Number(r.vale) : 0,
        inicioExpediente: (typeof r.inicioExpediente === 'string' && horaValida(r.inicioExpediente)) ? r.inicioExpediente : null
    };
}

function normalizarAcordo(aBruto) {
    const a = ehObjeto(aBruto) ? aBruto : {};
    return {
        id: (typeof a.id === 'string' && a.id) ? a.id : novoId('acd'),
        nome: (typeof a.nome === 'string' && a.nome.trim()) ? a.nome.trim() : '',
        periodos: (Array.isArray(a.periodos) ? a.periodos : []).map(normalizarPeriodoAcordo),
        regrasHorario: (Array.isArray(a.regrasHorario) ? a.regrasHorario : []).map(normalizarRegraHorario)
    };
}

function normalizarEvento(eBruto) {
    const e = ehObjeto(eBruto) ? eBruto : {};
    return {
        id: (typeof e.id === 'string' && e.id) ? e.id : novoId('evt'),
        tipoEvento: (typeof e.tipoEvento === 'string' && e.tipoEvento) ? e.tipoEvento : 'outro',
        descricaoEvento: (typeof e.descricaoEvento === 'string') ? e.descricaoEvento : '',
        dataInicioEvento: normalizarDataIso(e.dataInicioEvento),
        dataFimEvento: normalizarDataIso(e.dataFimEvento) || normalizarDataIso(e.dataInicioEvento),
        periodo: (typeof e.periodo === 'string' && e.periodo) ? e.periodo : 'dia_todo',
        impactoEvento: (typeof e.impactoEvento === 'string' && e.impactoEvento) ? e.impactoEvento : 'folga',
        acordoIndex: (Number.isInteger(e.acordoIndex) && e.acordoIndex >= 0) ? e.acordoIndex : null,
        corFundo: (typeof e.corFundo === 'string' && e.corFundo) ? e.corFundo : null,
        corTexto: (typeof e.corTexto === 'string' && e.corTexto) ? e.corTexto : null,
        nomeCSS: (typeof e.nomeCSS === 'string') ? e.nomeCSS : ''
    };
}

function normalizarTipoEvento(tBruto, idx) {
    const t = ehObjeto(tBruto) ? tBruto : {};
    const padrao = TIPOS_EVENTO_PADRAO[idx] || TIPOS_EVENTO_PADRAO[TIPOS_EVENTO_PADRAO.length - 1];
    return {
        id: (typeof t.id === 'string' && t.id) ? t.id : padrao.id,
        nome: (typeof t.nome === 'string' && t.nome.trim()) ? t.nome.trim() : padrao.nome,
        cor: (typeof t.cor === 'string' && t.cor) ? t.cor : padrao.cor
    };
}

function normalizarPeriodoAquisitivo(pBruto) {
    const p = ehObjeto(pBruto) ? pBruto : {};
    return {
        id: (typeof p.id === 'string' && p.id) ? p.id : novoId('per'),
        periodoIndex: Number.isFinite(Number(p.periodoIndex)) ? Number(p.periodoIndex) : 1,
        inicio: normalizarDataIso(p.inicio),
        termino: normalizarDataIso(p.termino),
        limite: normalizarDataIso(p.limite),
        subIndex: Number.isFinite(Number(p.subIndex)) ? Number(p.subIndex) : null,
        subTotal: Number.isFinite(Number(p.subTotal)) ? Number(p.subTotal) : null,
        feriasInicio: normalizarDataIso(p.feriasInicio),
        feriasFim: normalizarDataIso(p.feriasFim),
        adto13: typeof p.adto13 === 'string' ? p.adto13 : '',
        dias: Number.isFinite(Number(p.dias)) ? Number(p.dias) : null,
        documento: typeof p.documento === 'string' ? p.documento : ''
    };
}

function normalizarConfiguracoes(cBruto) {
    const c = ehObjeto(cBruto) ? cBruto : {};
    return {
        tipoJornada: Number.isFinite(Number(c.tipoJornada)) ? Number(c.tipoJornada) : 44,
        entradaPadrao: (typeof c.entradaPadrao === 'string') ? c.entradaPadrao : '',
        saidaPadrao: (typeof c.saidaPadrao === 'string') ? c.saidaPadrao : '',
        almocoMinutos: Number.isFinite(Number(c.almocoMinutos)) ? Number(c.almocoMinutos) : 60,
        toleranciaAtraso: Number.isFinite(Number(c.toleranciaAtraso)) ? Number(c.toleranciaAtraso) : 5,
        inicioPeriodoBanco: normalizarDataIso(c.inicioPeriodoBanco),
        fimPeriodoBanco: normalizarDataIso(c.fimPeriodoBanco),
        feriasDias: Number.isFinite(Number(c.feriasDias)) && Number(c.feriasDias) > 0 ? Number(c.feriasDias) : 30
    };
}

/**
 * Normaliza o objeto ponto inteiro: garante todos os arrays/campos, preenche
 * defaults, deduplica registros por data (o Ponto usa `data` como chave única
 * — um só registro por dia). Pura e idempotente.
 */
function normalizarPonto(pontoBruto) {
    const p = ehObjeto(pontoBruto) ? pontoBruto : {};

    const registrosBrutos = (Array.isArray(p.registros) ? p.registros : []).map(normalizarRegistro).filter(r => r.data);
    const porData = {};
    registrosBrutos.forEach(r => { porData[r.data] = r; }); // último ganha, mesma regra do Ponto (findIndex + substitui)
    const registros = Object.keys(porData).sort().map(d => porData[d]);

    const acordos = (Array.isArray(p.acordos) ? p.acordos : []).map(normalizarAcordo);
    const eventos = (Array.isArray(p.eventos) ? p.eventos : []).map(normalizarEvento).filter(e => e.dataInicioEvento);
    const tiposEvento = (Array.isArray(p.tiposEvento) && p.tiposEvento.length)
        ? p.tiposEvento.map(normalizarTipoEvento)
        : TIPOS_EVENTO_PADRAO.map(normalizarTipoEvento);
    const periodosAquisitivos = (Array.isArray(p.periodosAquisitivos) ? p.periodosAquisitivos : []).map(normalizarPeriodoAquisitivo);

    return {
        versao: 1,
        admissao: normalizarDataIso(p.admissao),
        registros,
        configuracoes: normalizarConfiguracoes(p.configuracoes),
        eventos,
        acordos,
        tiposEvento,
        periodosAquisitivos
    };
}

// ──────────────────────────────────────────────
//  FACTORIES
// ──────────────────────────────────────────────

function criarRegistro(dados) {
    return normalizarRegistro(dados);
}

function criarEvento(dados) {
    return normalizarEvento(Object.assign({}, dados, { id: undefined }));
}

function criarAcordo(dados) {
    return normalizarAcordo(Object.assign({}, dados, { id: undefined }));
}

function criarPeriodoAquisitivo(dados) {
    return normalizarPeriodoAquisitivo(Object.assign({}, dados, { id: undefined }));
}

// ──────────────────────────────────────────────
//  VALIDAÇÃO — porte de validators.js do Ponto
// ──────────────────────────────────────────────

function validarRegistro(registro) {
    const erros = [];
    if (!ehObjeto(registro)) { erros.push('Registro deve ser um objeto válido'); return erros; }
    if (!normalizarDataIso(registro.data)) erros.push('Data inválida (use formato YYYY-MM-DD)');
    if (registro.entrada && !horaValida(registro.entrada)) erros.push('Hora de entrada inválida (use formato HH:MM)');
    if (registro.saida && !horaValida(registro.saida)) erros.push('Hora de saída inválida (use formato HH:MM)');
    if (registro.saidaAlmoco && !horaValida(registro.saidaAlmoco)) erros.push('Hora de saída de almoço inválida');
    if (registro.retornoAlmoco && !horaValida(registro.retornoAlmoco)) erros.push('Hora de retorno de almoço inválida');
    if (horaValida(registro.entrada) && registro.saida && horaValida(registro.saida)) {
        const [hE, mE] = registro.entrada.split(':').map(Number);
        const [hS, mS] = registro.saida.split(':').map(Number);
        if ((hS * 60 + mS) <= (hE * 60 + mE)) erros.push('Hora de saída deve ser posterior à entrada');
    }
    return erros;
}

function validarPeriodoAcordo(periodo) {
    const erros = [];
    if (!ehObjeto(periodo)) { erros.push('Período deve ser um objeto válido'); return erros; }
    const isoInicio = normalizarDataIso(periodo.inicio);
    const isoFim = normalizarDataIso(periodo.fim);
    if (!isoInicio) erros.push('Data de início do período inválida');
    if (!isoFim) erros.push('Data de fim do período inválida');
    if (isoInicio && isoFim && isoFim < isoInicio) erros.push('Data de fim não pode ser anterior à de início');
    const minExtras = Number(periodo.minutosExtras || 0);
    if (isNaN(minExtras) || minExtras < 0) erros.push('Minutos extras deve ser um número não-negativo');
    return erros;
}

function validarRegraHorario(regra) {
    const erros = [];
    if (!ehObjeto(regra)) { erros.push('Regra deve ser um objeto válido'); return erros; }
    const isoInicio = normalizarDataIso(regra.inicio);
    const isoFim = normalizarDataIso(regra.fim);
    if (!isoInicio) erros.push('Data de início da regra inválida');
    if (!isoFim) erros.push('Data de fim da regra inválida');
    if (isoInicio && isoFim && isoFim < isoInicio) erros.push('Data de fim não pode ser anterior à de início');
    const almocoMin = Number(regra.almocoMin != null ? regra.almocoMin : 60);
    if (isNaN(almocoMin) || almocoMin < 0 || almocoMin > 480) erros.push('Almoço deve estar entre 0 e 480 minutos');
    const tolAlmoco = Number(regra.tolAlmoco != null ? regra.tolAlmoco : 5);
    if (isNaN(tolAlmoco) || tolAlmoco < 0) erros.push('Tolerância de almoço deve ser não-negativa');
    const tolSaida = Number(regra.tolSaida != null ? regra.tolSaida : 5);
    if (isNaN(tolSaida) || tolSaida < 0) erros.push('Tolerância de saída deve ser não-negativa');
    const vale = Number(regra.vale || 0);
    if (isNaN(vale) || vale < 0) erros.push('Valor do vale deve ser não-negativo');
    return erros;
}

function validarEvento(evento) {
    const erros = [];
    if (!ehObjeto(evento)) { erros.push('Evento deve ser um objeto válido'); return erros; }
    if (!evento.tipoEvento || typeof evento.tipoEvento !== 'string') erros.push('Tipo de evento é obrigatório');
    if (!evento.descricaoEvento || typeof evento.descricaoEvento !== 'string') erros.push('Descrição do evento é obrigatória');
    const isoInicio = normalizarDataIso(evento.dataInicioEvento);
    const isoFim = normalizarDataIso(evento.dataFimEvento);
    if (!isoInicio) erros.push('Data de início do evento inválida');
    if (!isoFim) erros.push('Data de fim do evento inválida');
    if (isoInicio && isoFim && isoFim < isoInicio) erros.push('Data de fim não pode ser anterior à de início');
    if (evento.acordoIndex !== undefined && evento.acordoIndex !== null) {
        if (!Number.isInteger(evento.acordoIndex) || evento.acordoIndex < 0) erros.push('Índice de acordo inválido');
    }
    return erros;
}

function validarAcordo(acordo) {
    const erros = [];
    if (!ehObjeto(acordo)) { erros.push('Acordo deve ser um objeto válido'); return erros; }
    if (!acordo.nome || typeof acordo.nome !== 'string' || !acordo.nome.trim()) erros.push('Nome do acordo é obrigatório');
    if (!Array.isArray(acordo.periodos) || acordo.periodos.length === 0) erros.push('Acordo deve ter pelo menos um período');
    (acordo.periodos || []).forEach((p, idx) => {
        validarPeriodoAcordo(p).forEach(e => erros.push('Período ' + (idx + 1) + ': ' + e));
    });
    (acordo.regrasHorario || []).forEach((r, idx) => {
        validarRegraHorario(r).forEach(e => erros.push('Regra ' + (idx + 1) + ': ' + e));
    });
    return erros;
}

function validarConfiguracoes(config) {
    const erros = [];
    if (!ehObjeto(config)) { erros.push('Configurações devem ser um objeto válido'); return erros; }
    const tipoJornada = Number(config.tipoJornada || 44);
    if (isNaN(tipoJornada) || tipoJornada <= 0 || tipoJornada > 168) erros.push('Tipo de jornada deve estar entre 1 e 168 horas');
    const almocoMinutos = Number(config.almocoMinutos != null ? config.almocoMinutos : 60);
    if (isNaN(almocoMinutos) || almocoMinutos < 0 || almocoMinutos > 480) erros.push('Almoço deve estar entre 0 e 480 minutos');
    const toleranciaAtraso = Number(config.toleranciaAtraso != null ? config.toleranciaAtraso : 5);
    if (isNaN(toleranciaAtraso) || toleranciaAtraso < 0) erros.push('Tolerância deve ser não-negativa');
    if (config.inicioPeriodoBanco && !normalizarDataIso(config.inicioPeriodoBanco)) erros.push('Data de início do banco de horas inválida');
    if (config.fimPeriodoBanco && !normalizarDataIso(config.fimPeriodoBanco)) erros.push('Data de fim do banco de horas inválida');
    return erros;
}

const PontoModel = {
    TIPOS_ATESTADO,
    TIPOS_EVENTO_PADRAO,

    novoId,
    normalizarDataIso,
    horaValida,

    normalizarPonto,
    normalizarRegistro,
    normalizarAcordo,
    normalizarPeriodoAcordo,
    normalizarRegraHorario,
    normalizarEvento,
    normalizarTipoEvento,
    normalizarPeriodoAquisitivo,
    normalizarConfiguracoes,

    criarRegistro,
    criarEvento,
    criarAcordo,
    criarPeriodoAquisitivo,

    validarRegistro,
    validarPeriodoAcordo,
    validarRegraHorario,
    validarEvento,
    validarAcordo,
    validarConfiguracoes
};

if (typeof window !== 'undefined') {
    window.PontoModel = PontoModel;
}

if (typeof module !== 'undefined' && module.exports) {
    module.exports = PontoModel;
}
