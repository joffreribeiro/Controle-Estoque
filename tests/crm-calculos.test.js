import { describe, it, expect } from 'vitest';
import CrmCalculos from '../crm-calculos.js';

describe('CrmCalculos — export completo do módulo', () => {
  it('expõe todas as funções usadas pela UI (hojeIso e diasEntre incluídos)', () => {
    ['somarItens', 'agruparPorEtapa', 'somarValor', 'resumoFunil', 'filtrarNegocios',
      'ordenarNegocios', 'reordenarNaEtapa', 'taxaConversao', 'formatarMoeda', 'negociosDoCliente',
      'timelineDe', 'hojeIso', 'diasEntre', 'atividadesPendentesDe', 'proximaAtividade',
      'temAtividadePendente', 'diasNaEtapa', 'idadeEmDias', 'diasInativo', 'agruparPorMesFechamento'
    ].forEach(function (nome) {
      expect(typeof CrmCalculos[nome]).toBe('function');
    });
  });
});

describe('CrmCalculos.somarItens', () => {
  it('soma quantidade × preço de cada item', () => {
    const itens = [
      { quantidade: 3, precoUnit: 1000 },
      { quantidade: 10, precoUnit: 50 }
    ];
    expect(CrmCalculos.somarItens(itens)).toBe(3500);
  });

  it('devolve 0 para lista vazia/indefinida e ignora item inválido', () => {
    expect(CrmCalculos.somarItens([])).toBe(0);
    expect(CrmCalculos.somarItens(undefined)).toBe(0);
    expect(CrmCalculos.somarItens([{ quantidade: 'x', precoUnit: 10 }])).toBe(0);
  });
});

describe('CrmCalculos.agruparPorEtapa', () => {
  const etapas = [{ id: 'e1', ordem: 0 }, { id: 'e2', ordem: 1 }];

  it('ignora negócio com etapaId órfão e ordena por ordem', () => {
    const negocios = [
      { id: 'n1', etapaId: 'e1', ordem: 2 },
      { id: 'n2', etapaId: 'e1', ordem: 0 },
      { id: 'n3', etapaId: 'inexistente', ordem: 0 }
    ];
    const out = CrmCalculos.agruparPorEtapa(negocios, etapas);
    expect(out.e1.map(n => n.id)).toEqual(['n2', 'n1']);
    expect(out.e2).toEqual([]);
  });
});

describe('CrmCalculos.somarValor / resumoFunil', () => {
  it('soma valores e calcula ticket médio só sobre ganhos', () => {
    const negocios = [
      { status: 'aberto', valor: 100 },
      { status: 'ganho', valor: 300 },
      { status: 'ganho', valor: 100 },
      { status: 'perdido', valor: 999 }
    ];
    const resumo = CrmCalculos.resumoFunil(negocios);
    expect(resumo.valorGanho).toBe(400);
    expect(resumo.ticketMedio).toBe(200);
    expect(resumo.abertos).toBe(1);
  });

  it('ticket médio é 0 (não NaN) sem ganhos', () => {
    expect(CrmCalculos.resumoFunil([{ status: 'aberto', valor: 100 }]).ticketMedio).toBe(0);
  });
});

describe('CrmCalculos.filtrarNegocios', () => {
  it('acha "Manutenção" buscando "manutencao"', () => {
    const negocios = [
      { id: 'n1', titulo: 'Contrato de Manutenção 2027' },
      { id: 'n2', titulo: 'Outra coisa' }
    ];
    expect(CrmCalculos.filtrarNegocios(negocios, { busca: 'manutencao' }).map(n => n.id)).toEqual(['n1']);
  });
});

describe('CrmCalculos.reordenarNaEtapa', () => {
  it('produz ordens densas e não muta a entrada', () => {
    const negocios = [
      { id: 'a', etapaId: 'e1', ordem: 0 },
      { id: 'b', etapaId: 'e1', ordem: 1 },
      { id: 'c', etapaId: 'e1', ordem: 2 }
    ];
    const copia = JSON.parse(JSON.stringify(negocios));
    const resultado = CrmCalculos.reordenarNaEtapa(negocios, 'e1', 'a', 2);
    const porId = Object.fromEntries(resultado.map(r => [r.id, r.ordem]));
    expect(porId.b).toBe(0);
    expect(porId.c).toBe(1);
    expect(porId.a).toBe(2);
    expect(negocios).toEqual(copia);
  });
});

describe('CrmCalculos.formatarMoeda', () => {
  it('formata em pt-BR', () => {
    expect(CrmCalculos.formatarMoeda(1234.5)).toContain('1.234,50');
    expect(CrmCalculos.formatarMoeda(undefined)).toContain('0,00');
  });
});

describe('CrmCalculos — atividades e métricas derivadas', () => {
  const atividades = [
    { id: 'a1', negocioId: 'n1', data: '2026-07-30', horaInicio: '10:00', feito: false },
    { id: 'a2', negocioId: 'n1', data: '2026-07-25', horaInicio: '14:00', feito: false },
    { id: 'a3', negocioId: 'n1', data: '2026-07-20', feito: true, feitoEm: '2026-07-20T15:00:00.000Z' }
  ];

  it('proximaAtividade devolve a pendente de menor data', () => {
    expect(CrmCalculos.proximaAtividade(atividades, 'n1').id).toBe('a2');
  });

  it('temAtividadePendente false quando só há feitas', () => {
    expect(CrmCalculos.temAtividadePendente([{ negocioId: 'n9', feito: true }], 'n9')).toBe(false);
    expect(CrmCalculos.temAtividadePendente(atividades, 'n1')).toBe(true);
  });

  it('diasNaEtapa conta desde a última mudança de etapa', () => {
    const negocio = { id: 'n1', criadoEm: '2026-07-01T12:00:00.000Z' };
    const historico = [
      { entidade: 'negocio', entidadeId: 'n1', tipo: 'etapa', criadoEm: '2026-07-10T09:00:00.000Z' },
      { entidade: 'negocio', entidadeId: 'n1', tipo: 'nota', criadoEm: '2026-07-18T09:00:00.000Z' }
    ];
    expect(CrmCalculos.diasNaEtapa(historico, negocio, '2026-07-22')).toBe(12);
    expect(CrmCalculos.diasNaEtapa([], negocio, '2026-07-22')).toBe(21);
  });

  it('agruparPorMesFechamento ordena meses e põe "sem data" por último', () => {
    const negocios = [
      { id: 'n1', dataPrevisao: '2026-08-15' },
      { id: 'n2', dataPrevisao: '2026-07-30' },
      { id: 'n3', dataPrevisao: null }
    ];
    const grupos = CrmCalculos.agruparPorMesFechamento(negocios);
    expect(grupos.map(g => g.mes)).toEqual(['2026-07', '2026-08', null]);
  });
});
