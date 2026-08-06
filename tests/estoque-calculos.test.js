import { describe, it, expect } from 'vitest';
import EstoqueCalculos from '../estoque-calculos.js';

describe('EstoqueCalculos.calcularStatusCusto', () => {
  it('retorna "sem_custo" quando custoTotal é null, undefined ou 0', () => {
    expect(EstoqueCalculos.calcularStatusCusto({ custoTotal: null })).toBe('sem_custo');
    expect(EstoqueCalculos.calcularStatusCusto({ custoTotal: undefined })).toBe('sem_custo');
    expect(EstoqueCalculos.calcularStatusCusto({ custoTotal: 0 })).toBe('sem_custo');
    expect(EstoqueCalculos.calcularStatusCusto({})).toBe('sem_custo');
    expect(EstoqueCalculos.calcularStatusCusto(null)).toBe('sem_custo');
  });

  it('retorna "sem_custo" para valores negativos ou não numéricos', () => {
    expect(EstoqueCalculos.calcularStatusCusto({ custoTotal: -50 })).toBe('sem_custo');
    expect(EstoqueCalculos.calcularStatusCusto({ custoTotal: 'abc' })).toBe('sem_custo');
    expect(EstoqueCalculos.calcularStatusCusto({ custoTotal: NaN })).toBe('sem_custo');
  });

  it('retorna "ok" quando há custo cadastrado maior que zero', () => {
    expect(EstoqueCalculos.calcularStatusCusto({ custoTotal: 150.5 })).toBe('ok');
    expect(EstoqueCalculos.calcularStatusCusto({ custoTotal: 0.01 })).toBe('ok');
  });
});
