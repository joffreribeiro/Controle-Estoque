import { describe, it, expect } from 'vitest';
import CrmModel from '../crm-model.js';

describe('CrmModel.normalizarCrm', () => {
  it('devolve estrutura completa com todos os arrays a partir de undefined', () => {
    const crm = CrmModel.normalizarCrm(undefined);
    expect(crm.versao).toBe(1);
    expect(Array.isArray(crm.funis)).toBe(true);
    expect(Array.isArray(crm.negocios)).toBe(true);
    expect(Array.isArray(crm.atividades)).toBe(true);
    expect(Array.isArray(crm.historico)).toBe(true);
    expect(crm.funis.length).toBe(0);
    expect(crm.config).toBeTruthy();
    // Não guarda clientes/produtos — lidos ao vivo do estoque
    expect(crm.pessoas).toBeUndefined();
    expect(crm.organizacoes).toBeUndefined();
  });

  it('preenche defaults sem perder id/clienteId de um negócio existente', () => {
    const crm = CrmModel.normalizarCrm({
      funis: [{ id: 'fnl_x', etapas: [{ id: 'etp_x', tipo: 'aberta' }] }],
      negocios: [{ id: 'a', funilId: 'fnl_x', etapaId: 'etp_x', clienteId: 'cli9' }]
    });
    expect(crm.negocios[0].id).toBe('a');
    expect(crm.negocios[0].clienteId).toBe('cli9');
    expect(crm.negocios[0].titulo).toBe('');
    expect(crm.negocios[0].status).toBe('aberto');
  });

  it('realoca negócio com etapaId inexistente para a primeira etapa aberta', () => {
    const crm = CrmModel.normalizarCrm({
      funis: [{ id: 'fnl_x', etapas: [{ id: 'etp_aberta', tipo: 'aberta' }, { id: 'etp_ganho', tipo: 'ganho' }] }],
      negocios: [{ id: 'a', funilId: 'fnl_x', etapaId: 'nao-existe' }]
    });
    expect(crm.negocios[0].etapaId).toBe('etp_aberta');
  });

  it('é idempotente: normalizar(normalizar(x)) é igual a normalizar(x)', () => {
    const entrada = {
      funis: [{ nome: 'Comercial', etapas: [{ nome: 'Qualificação' }, { nome: 'Ganho', tipo: 'ganho' }] }],
      negocios: [{ titulo: 'Negócio 1', clienteId: 'c1', itens: [{ produtoId: 5, nome: 'X', quantidade: 2, precoUnit: 10 }] }]
    };
    const uma = CrmModel.normalizarCrm(entrada);
    const duas = CrmModel.normalizarCrm(uma);
    expect(duas).toEqual(uma);
  });
});

describe('CrmModel.normalizarNegocio — itens de produto', () => {
  it('deriva valor da soma dos itens (quantidade × preço)', () => {
    const n = CrmModel.normalizarNegocio({
      titulo: 'Venda', itens: [
        { produtoId: 1, nome: 'Pistola', quantidade: 3, precoUnit: 1000 },
        { produtoId: 2, nome: 'Munição', quantidade: 10, precoUnit: 50 }
      ]
    });
    expect(n.valor).toBe(3500);
    expect(n.itens.length).toBe(2);
  });

  it('sem itens, aceita valor manual', () => {
    const n = CrmModel.normalizarNegocio({ titulo: 'Serviço', valor: 800 });
    expect(n.valor).toBe(800);
    expect(n.itens).toEqual([]);
  });

  it('normaliza item com quantidade/preço inválidos para defaults', () => {
    const it = CrmModel.normalizarItem({ produtoId: 7 });
    expect(it.quantidade).toBe(1);
    expect(it.precoUnit).toBe(0);
    expect(it.produtoId).toBe(7);
  });
});

describe('CrmModel.funilDeTemplate', () => {
  it('template "demandas" não mostra valor e tem uma etapa ganho e uma perdido', () => {
    const funil = CrmModel.funilDeTemplate('demandas');
    expect(funil.mostrarValor).toBe(false);
    expect(funil.etapas.filter(e => e.tipo === 'ganho').length).toBe(1);
    expect(funil.etapas.filter(e => e.tipo === 'perdido').length).toBe(1);
  });

  it('devolve null para chave desconhecida', () => {
    expect(CrmModel.funilDeTemplate('inexistente')).toBeNull();
  });
});

describe('CrmModel.validarNegocio', () => {
  const funilComValor = { mostrarValor: true };
  const funilSemValor = { mostrarValor: false };

  it('rejeita título vazio', () => {
    expect(CrmModel.validarNegocio({ titulo: '' }, funilComValor).length).toBeGreaterThan(0);
  });

  it('aceita negócio com itens mesmo em funil sem valor monetário manual', () => {
    const dados = { titulo: 'X', itens: [{ produtoId: 1, nome: 'A', quantidade: 1, precoUnit: 100 }], valor: 100 };
    expect(CrmModel.validarNegocio(dados, funilSemValor)).toEqual([]);
  });

  it('rejeita valor manual em funil sem valor', () => {
    const erros = CrmModel.validarNegocio({ titulo: 'X', valor: 100 }, funilSemValor);
    expect(erros.some(e => /valor/i.test(e))).toBe(true);
  });
});

describe('CrmModel.normalizarAtividade / validarAtividade', () => {
  it('preenche defaults e força tipo válido', () => {
    const a = CrmModel.normalizarAtividade({ tipo: 'inexistente' });
    expect(a.tipo).toBe('tarefa');
    expect(a.feito).toBe(false);
    expect(a.id).toMatch(/^atv_/);
  });

  it('validarAtividade exige assunto, negócio e data', () => {
    expect(CrmModel.validarAtividade({ assunto: '', negocioId: null, data: null }).length).toBe(3);
    expect(CrmModel.validarAtividade({ assunto: 'Ligar', negocioId: 'n1', data: '2026-07-25' })).toEqual([]);
  });
});

describe('CrmModel.normalizarCrm — fase 2', () => {
  it('descarta atividades órfãs e aceita visão previsao/excluidos', () => {
    const crm = CrmModel.normalizarCrm({
      funis: [{ id: 'f1', etapas: [{ id: 'e1', tipo: 'aberta' }] }],
      negocios: [{ id: 'n1', funilId: 'f1', etapaId: 'e1' }],
      atividades: [{ id: 'a1', negocioId: 'n1', assunto: 'ok' }, { id: 'a2', negocioId: 'x', assunto: 'orfã' }],
      config: { visao: 'previsao' }
    });
    expect(crm.atividades.map(a => a.id)).toEqual(['a1']);
    expect(crm.config.visao).toBe('previsao');
  });

  it('preserva excluidoEm e participantes do negócio', () => {
    const crm = CrmModel.normalizarCrm({
      funis: [{ id: 'f1', etapas: [{ id: 'e1', tipo: 'aberta' }] }],
      negocios: [{ id: 'n1', funilId: 'f1', etapaId: 'e1', excluidoEm: '2026-07-01T00:00:00.000Z', participantes: ['cli1'] }]
    });
    expect(crm.negocios[0].excluidoEm).toBe('2026-07-01T00:00:00.000Z');
    expect(crm.negocios[0].participantes).toEqual(['cli1']);
  });
});
