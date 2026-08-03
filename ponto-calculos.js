/**
 * ponto-calculos.js - Cálculos de horas do módulo de Ponto (jornada/RH)
 * Sem DOM, sem estado global — testável isoladamente com Vitest.
 *
 * Porte de calculations.js do sistema Ponto (https://github.com/joffreribeiro/Ponto).
 * As regras de negócio (tolerâncias, atestados, banco de horas, acordos) foram
 * mantidas EXATAMENTE como no original — isto tem peso de conformidade
 * trabalhista, então divergência de comportamento aqui não é "melhoria", é bug.
 * Qualquer ajuste de regra deve ser deliberado e verificado contra o Ponto
 * original antes de ser considerado correto (ver Fase 3 do plano de migração).
 */

function timeToMinutes(hhmm) {
    if (!hhmm || typeof hhmm !== 'string') return null;
    const m = hhmm.match(/^(\d{1,2}):(\d{2})$/);
    if (!m) return null;
    const h = Number(m[1]), min = Number(m[2]);
    if (isNaN(h) || isNaN(min)) return null;
    return h * 60 + min;
}

function normalizeDateForComparison(dateStr) {
    if (!dateStr) return null;
    if (/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) return dateStr;
    const m = dateStr.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
    if (m) {
        const [, day, month, year] = m;
        return `${year}-${month}-${day}`;
    }
    return null;
}

function getEventoByData(eventos, dataStr) {
    if (!Array.isArray(eventos) || !dataStr) return null;
    const dataNormalizada = normalizeDateForComparison(dataStr);
    return eventos.find(e =>
        normalizeDateForComparison(e.dataInicioEvento) <= dataNormalizada &&
        normalizeDateForComparison(e.dataFimEvento) >= dataNormalizada
    ) || null;
}

function getAcordoByData(acordos, dataStr) {
    if (!Array.isArray(acordos) || !dataStr) return null;
    const dataNormalizada = normalizeDateForComparison(dataStr);
    for (const acordo of acordos) {
        if (!Array.isArray(acordo.periodos)) continue;
        const periodo = acordo.periodos.find(p =>
            normalizeDateForComparison(p.inicio) <= dataNormalizada &&
            normalizeDateForComparison(p.fim) >= dataNormalizada
        );
        if (periodo) return acordo;
    }
    return null;
}

function getMinutosExtrasForDay(acordo, dataStr) {
    if (!acordo || !Array.isArray(acordo.periodos) || !dataStr) return 0;
    const dataNormalizada = normalizeDateForComparison(dataStr);
    const periodo = acordo.periodos.find(p =>
        normalizeDateForComparison(p.inicio) <= dataNormalizada &&
        normalizeDateForComparison(p.fim) >= dataNormalizada
    );
    return periodo ? Number(periodo.minutosExtras || 0) : 0;
}

function getDefaultRegra() {
    return { almocoMin: 60, tolAlmoco: 0, tolSaida: 0, minutosExtras: 0 };
}

function getRegraHorarioForDay(acordo, dataStr) {
    if (!acordo || !Array.isArray(acordo.regrasHorario) || !dataStr) return getDefaultRegra();
    const dataNormalizada = normalizeDateForComparison(dataStr);
    return acordo.regrasHorario.find(r =>
        normalizeDateForComparison(r.inicio) <= dataNormalizada &&
        normalizeDateForComparison(r.fim) >= dataNormalizada
    ) || getDefaultRegra();
}

/** Calcula horas trabalhadas para um dia com detalhes (sem contexto de evento/acordo). */
function calculateDayDetail(registro, minutosExtrasPeriodo, regra) {
    if (!registro || !registro.entrada || !registro.saida) {
        return { trabalhadas: 0, saldo: 0, temRegistro: !!registro, status: 'sem_registro', detalhes: {} };
    }

    const entrada = timeToMinutes(registro.entrada);
    const saida = timeToMinutes(registro.saida);
    const saidaAlm = registro.saidaAlmoco ? timeToMinutes(registro.saidaAlmoco) : null;
    const retornoAlm = registro.retornoAlmoco ? timeToMinutes(registro.retornoAlmoco) : null;

    if (entrada === null || saida === null) {
        return { trabalhadas: 0, saldo: 0, temRegistro: true, status: 'erro_hora', detalhes: { erro: 'Horários inválidos' } };
    }

    const almocoPadrao = (regra.almocoMin != null) ? regra.almocoMin : 60;
    const tolAlmoco = (regra.tolAlmoco != null) ? regra.tolAlmoco : 5;
    const tolSaida = (regra.tolSaida != null) ? regra.tolSaida : 5;

    let duracaoAlmoco = almocoPadrao;
    if (saidaAlm !== null && retornoAlm !== null && retornoAlm > saidaAlm) {
        duracaoAlmoco = retornoAlm - saidaAlm;
        const diffAlmoco = duracaoAlmoco - almocoPadrao;
        if (Math.abs(diffAlmoco) <= tolAlmoco) duracaoAlmoco = almocoPadrao;
    }

    const trabalhadas = (saida - entrada) - duracaoAlmoco;
    const minutosExtrasRegra = regra.minutosExtras || 0;
    const carga = 480 + minutosExtrasRegra;
    let saldo = trabalhadas - carga;
    if (Math.abs(saldo) <= tolSaida) saldo = 0;

    let status = 'ok';
    if (saldo > 0) status = 'extra';
    if (saldo < 0) status = 'falta';

    return {
        trabalhadas, saldo, temRegistro: true, status,
        detalhes: { entrada, saida, duracaoAlmoco, carga, minutosExtrasRegra, minutosExtrasPeriodo }
    };
}

/** Calcula horas de um dia com contexto de acordo/eventos/atestados. */
function calculateDayWithContext(registros, eventos, acordos, dataStr, registro) {
    const evento = getEventoByData(eventos, dataStr);
    const acordo = getAcordoByData(acordos, dataStr);
    const minutosExtras = getMinutosExtrasForDay(acordo, dataStr);
    const regra = getRegraHorarioForDay(acordo, dataStr);

    if (registro && registro.tipoAtestado === 'comparecimento_vespertino') {
        const minutosExtrasRegra = regra.minutosExtras || 0;
        const carga = 480 + minutosExtrasRegra;
        return {
            trabalhadas: carga, saldo: 0, temRegistro: true, status: 'ok',
            evento, acordo, regra, tipoAtestado: 'comparecimento_vespertino',
            detalhes: { carga, atestado: true }
        };
    }

    if (registro && registro.tipoAtestado === 'comparecimento_matutino') {
        const entradaEfetiva = (regra && regra.inicioExpediente) || '07:45';
        const registroAjustado = Object.assign({}, registro, { entrada: entradaEfetiva });
        const calc = calculateDayDetail(registroAjustado, minutosExtras, regra);
        return Object.assign({}, calc, { evento, acordo, regra, tipoAtestado: 'comparecimento_matutino' });
    }

    if (registro && (registro.tipoAtestado === 'abono_matutino' ||
                     registro.tipoAtestado === 'abono_vespertino' ||
                     registro.tipoAtestado === 'abono_dia_todo' ||
                     registro.tipoAtestado === 'abono_acordo_matutino' ||
                     registro.tipoAtestado === 'abono_acordo_vespertino' ||
                     registro.tipoAtestado === 'abono_acordo_dia_todo')) {
        const minutosExtrasRegra = regra.minutosExtras || 0;
        const carga = 480 + minutosExtrasRegra;
        return {
            trabalhadas: carga, saldo: 0, temRegistro: true, status: 'ok',
            evento, acordo, regra, tipoAtestado: registro.tipoAtestado,
            detalhes: { carga, atestado: true }
        };
    }

    if (registro && (registro.tipoAtestado === 'pagar_hora_acordo_matutino' ||
                     registro.tipoAtestado === 'pagar_hora_acordo_vespertino' ||
                     registro.tipoAtestado === 'pagar_hora_acordo_dia_todo')) {
        const minutosExtrasRegra = regra.minutosExtras || 0;
        const cargaDia = 480 + minutosExtrasRegra;
        const meioPeriodo = registro.tipoAtestado !== 'pagar_hora_acordo_dia_todo';
        const cargaPeriodo = meioPeriodo ? cargaDia / 2 : cargaDia;
        const temPonto = registro.entrada || registro.saida;
        if (!temPonto) {
            return {
                trabalhadas: 0, saldo: -cargaPeriodo, temRegistro: false, status: 'falta',
                evento, acordo, regra, tipoAtestado: registro.tipoAtestado,
                detalhes: { carga: cargaPeriodo }
            };
        }
    }

    if (registro && registro.tipoAtestado === 'afastamento') {
        const minutosExtrasRegra = regra.minutosExtras || 0;
        const carga = 480 + minutosExtrasRegra;
        return {
            trabalhadas: carga, saldo: 0, temRegistro: true, status: 'ok',
            evento, acordo, regra, tipoAtestado: 'afastamento',
            detalhes: { carga, atestado: true }
        };
    }

    if (evento && evento.tipoEvento === 'pagar_hora_acordo') {
        const minutosExtrasRegra = regra.minutosExtras || 0;
        const cargaDia = 480 + minutosExtrasRegra;
        const periodo = evento.periodo || 'dia_todo';
        const meioPeriodo = (periodo === 'matutino' || periodo === 'vespertino');
        const temPonto = registro && (registro.entrada || registro.saida);
        if (!temPonto) {
            const cargaPeriodo = meioPeriodo ? cargaDia / 2 : cargaDia;
            return {
                trabalhadas: 0, saldo: -cargaPeriodo, temRegistro: false, status: 'falta',
                evento, acordo, regra, detalhes: { carga: cargaPeriodo }
            };
        }
    }

    const calc = calculateDayDetail(registro, minutosExtras, regra);
    return Object.assign({}, calc, { evento, acordo, regra });
}

/**
 * Uso de Abono e Pagar Hora de um acordo (porte de calcularUsoBeneficiosAcordo do Ponto).
 * `qtdAbono` é expresso em MEIO-PERÍODOS (1 = meio período, 2 = dia inteiro);
 * `qtdPagarHora` em horas. Consumo vem de eventos (abono_acordo/pagar_hora_acordo
 * vinculados via acordoIndex) e de registros com tipoAtestado abono_acordo_ / pagar_hora_acordo_.
 */
function calcularUsoBeneficiosAcordo(acordoIndex, acordo, eventos, registros, acordos) {
    let usadosMeioPeriodo = 0;
    let horasPagasUsadas = 0;

    (eventos || []).forEach(ev => {
        if (ev.acordoIndex == null || Number(ev.acordoIndex) !== Number(acordoIndex)) return;
        const tipo = String(ev.tipoEvento || '').toLowerCase();
        const periodo = String(ev.periodo || '').toLowerCase();
        const fatorPeriodo = (periodo === 'matutino' || periodo === 'vespertino') ? 0.5 : 1;

        if (tipo === 'abono_acordo') {
            const inicio = new Date(ev.dataInicioEvento), fim = new Date(ev.dataFimEvento);
            if (!isNaN(inicio) && !isNaN(fim)) {
                const dias = Math.floor((fim - inicio) / 86400000) + 1;
                usadosMeioPeriodo += dias * (fatorPeriodo * 2);
            }
        }
        if (tipo === 'pagar_hora_acordo' || tipo === 'compensar_acordo' || tipo === 'pagar_hora') {
            const inicio = new Date(ev.dataInicioEvento), fim = new Date(ev.dataFimEvento);
            if (!isNaN(inicio) && !isNaN(fim)) {
                const dias = Math.floor((fim - inicio) / 86400000) + 1;
                const horasPorDia = (periodo === 'matutino' || periodo === 'vespertino') ? 4 : 8;
                horasPagasUsadas += dias * horasPorDia;
            }
        }
    });

    const periodos = acordo.periodos || [];
    (registros || []).forEach(reg => {
        const ta = String(reg.tipoAtestado || '');
        if (!ta.startsWith('abono_acordo_') && !ta.startsWith('pagar_hora_acordo_')) return;
        const dataReg = reg.data;
        const pertence = periodos.some(p => dataReg >= p.inicio && dataReg <= p.fim);
        if (!pertence) return;
        const acordoDaData = getAcordoByData(acordos, dataReg);
        if (!acordoDaData) return;
        if ((acordos || []).indexOf(acordoDaData) !== Number(acordoIndex)) return;

        const sufixo = ta.replace('abono_acordo_', '').replace('pagar_hora_acordo_', '');
        const fator = (sufixo === 'matutino' || sufixo === 'vespertino') ? 0.5 : 1;

        if (ta.startsWith('abono_acordo_')) {
            usadosMeioPeriodo += fator * 2;
        } else if (ta.startsWith('pagar_hora_acordo_')) {
            const regraReg = getRegraHorarioForDay(acordoDaData, dataReg);
            const extrasMin = (regraReg && regraReg.minutosExtras) || 0;
            const cargaMin = 480 + extrasMin;
            horasPagasUsadas += (fator * cargaMin) / 60;
        }
    });

    const qtdAbono = acordo.qtdAbono || 0;
    const qtdPagarHora = acordo.qtdPagarHora || 0;

    return {
        usadoAbono: usadosMeioPeriodo,
        restanteAbono: Math.max(0, qtdAbono - usadosMeioPeriodo),
        usadoPagarHora: horasPagasUsadas,
        restantePagarHora: Math.max(0, qtdPagarHora - horasPagasUsadas)
    };
}

/** Totalizações para um período (soma de calculateDayWithContext em cada registro). */
function calculatePeriodTotals(registros, eventos, acordos) {
    let totalTrabalhadas = 0, totalSaldo = 0, horasExtras = 0, horasFaltas = 0, horasAcordo = 0;
    let diasComExtra = 0, diasComFalta = 0;

    (registros || []).forEach(registro => {
        const calc = calculateDayWithContext(registros, eventos, acordos, registro.data, registro);
        totalTrabalhadas += calc.trabalhadas;
        totalSaldo += calc.saldo;
        if (calc.saldo > 0) { horasExtras += calc.saldo; diasComExtra++; }
        else if (calc.saldo < 0) { horasFaltas += Math.abs(calc.saldo); diasComFalta++; }
        if (calc.minutosExtras && calc.minutosExtras > 0) horasAcordo += calc.minutosExtras;
    });

    return {
        totalTrabalhadas, totalSaldo, horasExtras, horasFaltas, horasAcordo,
        diasComExtra, diasComFalta, diasProcessados: (registros || []).length
    };
}

/**
 * Formata minutos (podem ser negativos) como "HH:MM", ex: -27 -> "-00:27".
 * Porte de DateUtils.minutesToTime do Ponto — é este o formato usado no
 * timesheet, então trocá-lo por outro mudaria a folha de ponto visualmente.
 */
function minutosParaHHMM(totalMinutes) {
    const sinal = totalMinutes < 0 ? '-' : '';
    const abs = Math.abs(Math.round(totalMinutes));
    const h = Math.floor(abs / 60), m = abs % 60;
    return sinal + String(h).padStart(2, '0') + ':' + String(m).padStart(2, '0');
}

/** Formata minutos (podem ser negativos) como "Hh MMmin", ex: -90 -> "-1h 30min". */
function formatarMinutos(min) {
    const sinal = min < 0 ? '-' : '';
    const abs = Math.abs(Math.round(min));
    const h = Math.floor(abs / 60), m = abs % 60;
    if (h && m) return sinal + h + 'h ' + m + 'min';
    if (h) return sinal + h + 'h';
    return sinal + m + 'min';
}

const PontoCalculos = {
    timeToMinutes,
    normalizeDateForComparison,
    getEventoByData,
    getAcordoByData,
    getMinutosExtrasForDay,
    getDefaultRegra,
    getRegraHorarioForDay,
    calculateDayDetail,
    calculateDayWithContext,
    calculatePeriodTotals,
    formatarMinutos,
    minutosParaHHMM,
    calcularUsoBeneficiosAcordo
};

if (typeof window !== 'undefined') {
    window.PontoCalculos = PontoCalculos;
}

if (typeof module !== 'undefined' && module.exports) {
    module.exports = PontoCalculos;
}
