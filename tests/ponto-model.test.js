import { describe, it, expect } from 'vitest';
import PontoModel from '../ponto-model.js';

describe('PontoModel.normalizarPonto', () => {
  it('devolve estrutura completa com todos os arrays a partir de undefined', () => {
    const p = PontoModel.normalizarPonto(undefined);
    expect(p.versao).toBe(1);
    expect(Array.isArray(p.registros)).toBe(true);
    expect(Array.isArray(p.eventos)).toBe(true);
    expect(Array.isArray(p.acordos)).toBe(true);
    expect(Array.isArray(p.periodosAquisitivos)).toBe(true);
    expect(p.registros.length).toBe(0);
    // catálogo padrão de tipos de evento é preenchido quando ausente
    expect(p.tiposEvento.length).toBe(PontoModel.TIPOS_EVENTO_PADRAO.length);
    expect(p.configuracoes.tipoJornada).toBe(44);
  });

  it('é idempotente: normalizar(normalizar(x)) é igual a normalizar(x)', () => {
    const entrada = {
      admissao: '2020-01-15',
      registros: [{ data: '2026-07-01', entrada: '08:00', saida: '17:00' }],
      eventos: [{ tipoEvento: 'ferias', descricaoEvento: 'Férias', dataInicioEvento: '2026-08-01', dataFimEvento: '2026-08-10' }],
      acordos: [{ nome: 'Acordo X', periodos: [{ inicio: '2026-01-01', fim: '2026-12-31', minutosExtras: 30 }] }]
    };
    const uma = PontoModel.normalizarPonto(entrada);
    const duas = PontoModel.normalizarPonto(uma);
    expect(duas).toEqual(uma);
  });

  it('deduplica registros por data, mantendo o último (mesma regra do Ponto original)', () => {
    const p = PontoModel.normalizarPonto({
      registros: [
        { data: '2026-07-01', entrada: '08:00', saida: '17:00' },
        { data: '2026-07-01', entrada: '09:00', saida: '18:00' }
      ]
    });
    expect(p.registros.length).toBe(1);
    expect(p.registros[0].entrada).toBe('09:00');
  });

  it('descarta registro sem data válida', () => {
    const p = PontoModel.normalizarPonto({ registros: [{ entrada: '08:00', saida: '17:00' }] });
    expect(p.registros.length).toBe(0);
  });

  it('descarta evento sem data de início válida', () => {
    const p = PontoModel.normalizarPonto({ eventos: [{ tipoEvento: 'ferias', descricaoEvento: 'x' }] });
    expect(p.eventos.length).toBe(0);
  });

  it('evento sem dataFimEvento herda dataInicioEvento (evento de um dia só)', () => {
    const p = PontoModel.normalizarPonto({
      eventos: [{ tipoEvento: 'feriado', descricaoEvento: 'Feriado local', dataInicioEvento: '2026-09-07' }]
    });
    expect(p.eventos[0].dataFimEvento).toBe('2026-09-07');
  });

  it('aceita data em formato DD/MM/YYYY e normaliza para ISO', () => {
    expect(PontoModel.normalizarDataIso('15/01/2026')).toBe('2026-01-15');
    expect(PontoModel.normalizarDataIso('2026-01-15')).toBe('2026-01-15');
    expect(PontoModel.normalizarDataIso('31/02/2026')).toBe(null); // 31 de fevereiro não existe
    expect(PontoModel.normalizarDataIso('lixo')).toBe(null);
  });
});

describe('PontoModel.horaValida', () => {
  it('aceita HH:MM dentro da faixa válida e rejeita o resto', () => {
    expect(PontoModel.horaValida('08:00')).toBe(true);
    expect(PontoModel.horaValida('23:59')).toBe(true);
    expect(PontoModel.horaValida('24:00')).toBe(false);
    expect(PontoModel.horaValida('08:60')).toBe(false);
    expect(PontoModel.horaValida('')).toBe(false);
    expect(PontoModel.horaValida(null)).toBe(false);
  });
});

describe('PontoModel.validarRegistro', () => {
  it('rejeita data ausente/ inválida', () => {
    expect(PontoModel.validarRegistro({}).length).toBeGreaterThan(0);
    expect(PontoModel.validarRegistro({ data: 'lixo' }).length).toBeGreaterThan(0);
  });

  it('aceita registro só com data (horários são opcionais)', () => {
    expect(PontoModel.validarRegistro({ data: '2026-07-01' })).toEqual([]);
  });

  it('rejeita saída anterior ou igual à entrada', () => {
    const erros = PontoModel.validarRegistro({ data: '2026-07-01', entrada: '17:00', saida: '08:00' });
    expect(erros.some(e => /posterior/i.test(e))).toBe(true);
  });

  it('aceita registro completo válido', () => {
    expect(PontoModel.validarRegistro({
      data: '2026-07-01', entrada: '08:00', saidaAlmoco: '12:00', retornoAlmoco: '13:00', saida: '17:00'
    })).toEqual([]);
  });
});

describe('PontoModel.validarEvento / validarAcordo', () => {
  it('exige tipo, descrição e datas de um evento', () => {
    const erros = PontoModel.validarEvento({});
    expect(erros.length).toBeGreaterThanOrEqual(4);
  });

  it('rejeita evento com data fim anterior à data início', () => {
    const erros = PontoModel.validarEvento({
      tipoEvento: 'ferias', descricaoEvento: 'x', dataInicioEvento: '2026-08-10', dataFimEvento: '2026-08-01'
    });
    expect(erros.some(e => /anterior/i.test(e))).toBe(true);
  });

  it('exige nome e ao menos um período de um acordo', () => {
    expect(PontoModel.validarAcordo({}).length).toBeGreaterThanOrEqual(2);
    expect(PontoModel.validarAcordo({ nome: 'Acordo X', periodos: [{ inicio: '2026-01-01', fim: '2026-12-31' }] })).toEqual([]);
  });
});
