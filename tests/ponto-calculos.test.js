import { describe, it, expect } from 'vitest';
import PontoCalculos from '../ponto-calculos.js';

describe('PontoCalculos.timeToMinutes', () => {
  it('converte HH:MM em minutos desde meia-noite', () => {
    expect(PontoCalculos.timeToMinutes('08:00')).toBe(480);
    expect(PontoCalculos.timeToMinutes('17:30')).toBe(1050);
    expect(PontoCalculos.timeToMinutes('')).toBe(null);
    expect(PontoCalculos.timeToMinutes(null)).toBe(null);
  });
});

describe('PontoCalculos.calculateDayDetail — dia comum, sem acordo', () => {
  const regraPadrao = PontoCalculos.getDefaultRegra();

  it('sem registro (ausência) devolve status sem_registro e zero horas', () => {
    const calc = PontoCalculos.calculateDayDetail(null, 0, regraPadrao);
    expect(calc.status).toBe('sem_registro');
    expect(calc.trabalhadas).toBe(0);
  });

  it('jornada exata de 8h com 1h de almoço fecha saldo zero', () => {
    const registro = { entrada: '08:00', saidaAlmoco: '12:00', retornoAlmoco: '13:00', saida: '17:00' };
    const calc = PontoCalculos.calculateDayDetail(registro, 0, { almocoMin: 60, tolAlmoco: 0, tolSaida: 0 });
    expect(calc.trabalhadas).toBe(480); // 9h de portão a portão - 1h de almoço = 8h
    expect(calc.saldo).toBe(0);
    expect(calc.status).toBe('ok');
  });

  it('trabalhar além da carga gera saldo positivo (status extra)', () => {
    const registro = { entrada: '08:00', saidaAlmoco: '12:00', retornoAlmoco: '13:00', saida: '18:00' };
    const calc = PontoCalculos.calculateDayDetail(registro, 0, { almocoMin: 60, tolAlmoco: 0, tolSaida: 0 });
    expect(calc.trabalhadas).toBe(540); // 9h líquidas
    expect(calc.saldo).toBe(60);
    expect(calc.status).toBe('extra');
  });

  it('sair antes da carga completa gera saldo negativo (status falta)', () => {
    const registro = { entrada: '08:00', saidaAlmoco: '12:00', retornoAlmoco: '13:00', saida: '16:00' };
    const calc = PontoCalculos.calculateDayDetail(registro, 0, { almocoMin: 60, tolAlmoco: 0, tolSaida: 0 });
    expect(calc.saldo).toBe(-60);
    expect(calc.status).toBe('falta');
  });

  it('tolerância de saída absorve pequenas diferenças (saldo zerado)', () => {
    const registro = { entrada: '08:00', saidaAlmoco: '12:00', retornoAlmoco: '13:00', saida: '17:03' };
    const calc = PontoCalculos.calculateDayDetail(registro, 0, { almocoMin: 60, tolAlmoco: 0, tolSaida: 5 });
    expect(calc.saldo).toBe(0);
    expect(calc.status).toBe('ok');
  });

  it('almoço maior que o padrão, mas dentro da tolerância, não penaliza', () => {
    const registro = { entrada: '08:00', saidaAlmoco: '12:00', retornoAlmoco: '13:04', saida: '17:00' };
    const calc = PontoCalculos.calculateDayDetail(registro, 0, { almocoMin: 60, tolAlmoco: 5, tolSaida: 0 });
    expect(calc.detalhes.duracaoAlmoco).toBe(60); // absorvido pela tolerância
  });

  it('minutosExtras da regra aumenta a carga esperada do dia', () => {
    const registro = { entrada: '08:00', saidaAlmoco: '12:00', retornoAlmoco: '13:00', saida: '17:30' };
    const calc = PontoCalculos.calculateDayDetail(registro, 0, { almocoMin: 60, tolAlmoco: 0, tolSaida: 0, minutosExtras: 30 });
    expect(calc.detalhes.carga).toBe(510); // 480 + 30
    expect(calc.saldo).toBe(0);
  });
});

describe('PontoCalculos.calculateDayWithContext — atestados e eventos', () => {
  it('afastamento: carga cheia, saldo zero, sem depender do registro real', () => {
    const registro = { data: '2026-07-01', tipoAtestado: 'afastamento' };
    const calc = PontoCalculos.calculateDayWithContext([registro], [], [], '2026-07-01', registro);
    expect(calc.saldo).toBe(0);
    expect(calc.trabalhadas).toBe(480);
    expect(calc.tipoAtestado).toBe('afastamento');
  });

  it('comparecimento_matutino ajusta a entrada efetiva para 07:45 quando não há regra de acordo', () => {
    const registro = { data: '2026-07-01', tipoAtestado: 'comparecimento_matutino', saidaAlmoco: '12:00', retornoAlmoco: '13:00', saida: '17:00' };
    const calc = PontoCalculos.calculateDayWithContext([registro], [], [], '2026-07-01', registro);
    // entrada efetiva 07:45 -> saída 17:00, almoço 1h => 8h15 trabalhadas
    expect(calc.trabalhadas).toBe(495);
  });

  it('evento "pagar_hora_acordo" sem ponto batido gera saldo negativo proporcional ao período', () => {
    const evento = { tipoEvento: 'pagar_hora_acordo', dataInicioEvento: '2026-07-01', dataFimEvento: '2026-07-01', periodo: 'matutino' };
    const calc = PontoCalculos.calculateDayWithContext([], [evento], [], '2026-07-01', null);
    expect(calc.status).toBe('falta');
    expect(calc.saldo).toBe(-240); // metade da carga de 480
  });

  it('getEventoByData encontra o evento vigente numa data dentro do intervalo', () => {
    const eventos = [{ dataInicioEvento: '2026-08-01', dataFimEvento: '2026-08-10', tipoEvento: 'ferias' }];
    expect(PontoCalculos.getEventoByData(eventos, '2026-08-05')).toBe(eventos[0]);
    expect(PontoCalculos.getEventoByData(eventos, '2026-08-11')).toBe(null);
  });
});

describe('PontoCalculos.calculatePeriodTotals', () => {
  it('soma trabalhadas/saldo e conta dias de extra e falta corretamente', () => {
    const registros = [
      { data: '2026-07-01', entrada: '08:00', saidaAlmoco: '12:00', retornoAlmoco: '13:00', saida: '18:00' }, // +60
      { data: '2026-07-02', entrada: '08:00', saidaAlmoco: '12:00', retornoAlmoco: '13:00', saida: '16:00' }  // -60
    ];
    const totais = PontoCalculos.calculatePeriodTotals(registros, [], []);
    expect(totais.diasComExtra).toBe(1);
    expect(totais.diasComFalta).toBe(1);
    expect(totais.totalSaldo).toBe(0);
    expect(totais.diasProcessados).toBe(2);
  });
});

describe('PontoCalculos.formatarMinutos', () => {
  it('formata horas e minutos, incluindo negativos', () => {
    expect(PontoCalculos.formatarMinutos(90)).toBe('1h 30min');
    expect(PontoCalculos.formatarMinutos(60)).toBe('1h');
    expect(PontoCalculos.formatarMinutos(45)).toBe('45min');
    expect(PontoCalculos.formatarMinutos(-90)).toBe('-1h 30min');
  });
});
