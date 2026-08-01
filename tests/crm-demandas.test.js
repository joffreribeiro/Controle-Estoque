import { describe, it, expect } from 'vitest';
import CrmModel from '../crm-model.js';
import CrmCalculos from '../crm-calculos.js';

const FUNIL = { id: 'fnl_1', etapas: [{ id: 'etp_ab', tipo: 'aberta' }, { id: 'etp_ok', tipo: 'ganho' }] };
const HOJE = '2026-07-30';

describe('CrmModel.normalizarAnotacao — migração de registros legados', () => {
  it('é idempotente', () => {
    const entrada = {
      assunto: 'Parecer técnico',
      destinatario: 'Setor Jurídico',
      lembrarDiasAntes: '3',
      prazo: '2026-08-10'
    };
    const uma = CrmModel.normalizarAnotacao(entrada);
    const duas = CrmModel.normalizarAnotacao(uma);
    // ids e timestamps são gerados uma única vez e preservados na 2ª passada
    expect(duas).toEqual(uma);
  });

  it('anotação legada com destinatário e não finalizada vira aguardando_terceiro com 1 encaminhamento pendente', () => {
    const a = CrmModel.normalizarAnotacao({
      assunto: 'x', destinatario: 'Fulano', finalizado: false
    });
    expect(a.situacao).toBe('aguardando_terceiro');
    expect(a.encaminhamentos.length).toBe(1);
    expect(a.encaminhamentos[0].para).toBe('Fulano');
    expect(a.encaminhamentos[0].status).toBe('pendente');
  });

  it('anotação legada finalizada vira respondida com encaminhamento respondido', () => {
    const a = CrmModel.normalizarAnotacao({
      assunto: 'x', destinatario: 'Fulano', finalizado: true, dataConclusao: '2026-07-01'
    });
    expect(a.situacao).toBe('respondida');
    expect(a.finalizado).toBe(true);
    expect(a.encaminhamentos[0].status).toBe('respondido');
    expect(a.encaminhamentos[0].dataResposta).toBe('2026-07-01');
  });

  it('anotação legada sem destinatário e não finalizada vira recebida, sem encaminhamentos', () => {
    const a = CrmModel.normalizarAnotacao({ assunto: 'x' });
    expect(a.situacao).toBe('recebida');
    expect(a.encaminhamentos).toEqual([]);
  });

  it('não recria encaminhamento a partir do destinatário quando o array já existe', () => {
    const a = CrmModel.normalizarAnotacao({
      assunto: 'x', destinatario: 'Legado', encaminhamentos: []
    });
    expect(a.encaminhamentos).toEqual([]);
    expect(a.destinatario).toBe('Legado');
  });

  it('finalizado é derivado da situação, nunca fonte da verdade', () => {
    const a = CrmModel.normalizarAnotacao({ assunto: 'x', situacao: 'comigo', finalizado: true });
    expect(a.finalizado).toBe(false);
  });

  it('converte lembrarDiasAntes para number e descarta lixo', () => {
    expect(CrmModel.normalizarAnotacao({ lembrarDiasAntes: '3' }).lembrarDiasAntes).toBe(3);
    expect(CrmModel.normalizarAnotacao({ lembrarDiasAntes: 5 }).lembrarDiasAntes).toBe(5);
    expect(CrmModel.normalizarAnotacao({ lembrarDiasAntes: '' }).lembrarDiasAntes).toBe(null);
    expect(CrmModel.normalizarAnotacao({ lembrarDiasAntes: 'abc' }).lembrarDiasAntes).toBe(null);
    expect(CrmModel.normalizarAnotacao({ lembrarDiasAntes: -2 }).lembrarDiasAntes).toBe(null);
  });

  it('rejeita situação inválida caindo na inferência', () => {
    const a = CrmModel.normalizarAnotacao({ assunto: 'x', situacao: 'inventada' });
    expect(a.situacao).toBe('recebida');
  });

  it('não emite mais respostaTipoDoc, e descarta o valor de registros antigos', () => {
    const a = CrmModel.normalizarAnotacao({ assunto: 'x', respostaTipoDoc: 'Ofício de resposta' });
    expect('respostaTipoDoc' in a).toBe(false);
    // o irmão que a Lista realmente exibe continua vivo
    const b = CrmModel.normalizarAnotacao({ assunto: 'x', respostaNumeroDocumento: '123/2026' });
    expect(b.respostaNumeroDocumento).toBe('123/2026');
    // e segue idempotente sem o campo removido
    expect(CrmModel.normalizarAnotacao(a)).toEqual(a);
  });
});

describe('CrmModel.normalizarCrm — demandas avulsas', () => {
  it('REGRESSÃO: demanda sem negocioId mas com funil válido sobrevive', () => {
    const crm = CrmModel.normalizarCrm({
      funis: [FUNIL],
      anotacoes: [{ id: 'ant_1', assunto: 'avulsa', funilId: 'fnl_1', etapaId: 'etp_ab' }]
    });
    expect(crm.anotacoes.length).toBe(1);
    expect(crm.anotacoes[0].id).toBe('ant_1');
    expect(crm.anotacoes[0].funilId).toBe('fnl_1');
  });

  it('demanda com negocioId inexistente é realocada, não deletada', () => {
    const crm = CrmModel.normalizarCrm({
      funis: [FUNIL],
      anotacoes: [{ id: 'ant_1', assunto: 'órfã', negocioId: 'ngc_sumiu' }]
    });
    expect(crm.anotacoes.length).toBe(1);
    expect(crm.anotacoes[0].negocioId).toBe(null);
    expect(crm.anotacoes[0].funilId).toBe('fnl_1');
    expect(crm.anotacoes[0].etapaId).toBe('etp_ab');
  });

  it('demanda com negócio válido preserva o vínculo', () => {
    const crm = CrmModel.normalizarCrm({
      funis: [FUNIL],
      negocios: [{ id: 'ngc_1', funilId: 'fnl_1', etapaId: 'etp_ab' }],
      anotacoes: [{ id: 'ant_1', assunto: 'x', negocioId: 'ngc_1' }]
    });
    expect(crm.anotacoes[0].negocioId).toBe('ngc_1');
  });

  it('demanda avulsa com etapa inexistente cai na primeira etapa aberta', () => {
    const crm = CrmModel.normalizarCrm({
      funis: [FUNIL],
      anotacoes: [{ id: 'ant_1', assunto: 'x', funilId: 'fnl_1', etapaId: 'nao-existe' }]
    });
    expect(crm.anotacoes[0].etapaId).toBe('etp_ab');
  });

  it('zera anotacaoRelacionadaId que aponta para demanda inexistente', () => {
    const crm = CrmModel.normalizarCrm({
      funis: [FUNIL],
      anotacoes: [{ id: 'ant_1', assunto: 'x', funilId: 'fnl_1', anotacaoRelacionadaId: 'ant_fantasma' }]
    });
    expect(crm.anotacoes[0].anotacaoRelacionadaId).toBe(null);
  });
});

describe('CrmModel.validarAnotacao', () => {
  it('aceita demanda avulsa com funil', () => {
    expect(CrmModel.validarAnotacao({ assunto: 'x', funilId: 'fnl_1' })).toEqual([]);
  });

  it('recusa demanda sem negócio e sem funil', () => {
    const erros = CrmModel.validarAnotacao({ assunto: 'x' });
    expect(erros.length).toBe(1);
    expect(erros[0]).toMatch(/negócio ou a um funil/);
  });

  it('recusa encaminhamento sem destinatário e com data malformada', () => {
    const erros = CrmModel.validarAnotacao({
      assunto: 'x', funilId: 'fnl_1',
      encaminhamentos: [{ para: '', dataEnvio: '30/07/2026' }]
    });
    expect(erros.length).toBe(2);
  });
});

describe('CrmCalculos.diasAte / diasDesde', () => {
  it('devolve valor negativo para datas passadas (diferente de diasEntre, que trunca em 0)', () => {
    expect(CrmCalculos.diasAte('2026-07-28', HOJE)).toBe(-2);
    expect(CrmCalculos.diasEntre('2026-07-28', HOJE)).toBe(2);
  });

  it('zero no próprio dia e positivo no futuro', () => {
    expect(CrmCalculos.diasAte(HOJE, HOJE)).toBe(0);
    expect(CrmCalculos.diasAte('2026-08-02', HOJE)).toBe(3);
    expect(CrmCalculos.diasDesde('2026-07-25', HOJE)).toBe(5);
  });

  it('devolve null para data ausente', () => {
    expect(CrmCalculos.diasAte(null, HOJE)).toBe(null);
  });
});

describe('CrmCalculos.resumoEncaminhamentos', () => {
  const demanda = enc => CrmModel.normalizarAnotacao({ assunto: 'x', encaminhamentos: enc });

  it('sem encaminhamentos', () => {
    const r = CrmCalculos.resumoEncaminhamentos(demanda([]), HOJE);
    expect(r).toMatchObject({ total: 0, pendentes: 0, respondidos: 0, algumVencido: false });
    expect(r.comQuem).toEqual([]);
  });

  it('lista quem ainda deve resposta e ignora os já respondidos', () => {
    const r = CrmCalculos.resumoEncaminhamentos(demanda([
      { para: 'Jurídico', status: 'pendente' },
      { para: 'Financeiro', status: 'respondido' },
      { para: 'TI', status: 'pendente' }
    ]), HOJE);
    expect(r.total).toBe(3);
    expect(r.pendentes).toBe(2);
    expect(r.respondidos).toBe(1);
    expect(r.comQuem).toEqual(['Jurídico', 'TI']);
  });

  it('detecta prazo de terceiro vencido e pega o mais próximo', () => {
    const r = CrmCalculos.resumoEncaminhamentos(demanda([
      { para: 'A', status: 'pendente', prazoResposta: '2026-08-15' },
      { para: 'B', status: 'pendente', prazoResposta: '2026-07-20' }
    ]), HOJE);
    expect(r.prazoTerceiroMaisProximo).toBe('2026-07-20');
    expect(r.algumVencido).toBe(true);
  });

  it('prazo vencido de encaminhamento já respondido não conta como atraso', () => {
    const r = CrmCalculos.resumoEncaminhamentos(demanda([
      { para: 'A', status: 'respondido', prazoResposta: '2026-07-01' }
    ]), HOJE);
    expect(r.algumVencido).toBe(false);
  });

  it('diasAguardando é a maior espera entre os pendentes', () => {
    const r = CrmCalculos.resumoEncaminhamentos(demanda([
      { para: 'A', status: 'pendente', dataEnvio: '2026-07-28' },
      { para: 'B', status: 'pendente', dataEnvio: '2026-07-10' }
    ]), HOJE);
    expect(r.diasAguardando).toBe(20);
  });

  it('cancelado não conta como pendente nem como respondido', () => {
    const r = CrmCalculos.resumoEncaminhamentos(demanda([
      { para: 'A', status: 'cancelado' }
    ]), HOJE);
    expect(r.total).toBe(1);
    expect(r.pendentes).toBe(0);
    expect(r.respondidos).toBe(0);
  });
});

describe('CrmCalculos.situacaoSugerida', () => {
  const demanda = (situacao, enc) => CrmModel.normalizarAnotacao({ assunto: 'x', situacao, encaminhamentos: enc || [] });

  it('vira aguardando_terceiro quando há pedido pendente', () => {
    expect(CrmCalculos.situacaoSugerida(demanda('comigo', [{ para: 'A', status: 'pendente' }]))).toBe('aguardando_terceiro');
  });

  it('vira consolidando quando todos responderam', () => {
    expect(CrmCalculos.situacaoSugerida(demanda('aguardando_terceiro', [{ para: 'A', status: 'respondido' }]))).toBe('consolidando');
  });

  it('um pendente entre vários respondidos mantém aguardando', () => {
    expect(CrmCalculos.situacaoSugerida(demanda('aguardando_terceiro', [
      { para: 'A', status: 'respondido' },
      { para: 'B', status: 'pendente' }
    ]))).toBe('aguardando_terceiro');
  });

  it('nunca reabre uma demanda já respondida', () => {
    expect(CrmCalculos.situacaoSugerida(demanda('respondida', [{ para: 'A', status: 'pendente' }]))).toBe('respondida');
  });

  it('remover o último pendente devolve a demanda para mim', () => {
    expect(CrmCalculos.situacaoSugerida(demanda('aguardando_terceiro', []))).toBe('comigo');
  });

  it('preserva a situação quando não há encaminhamento algum', () => {
    expect(CrmCalculos.situacaoSugerida(demanda('recebida', []))).toBe('recebida');
  });
});

describe('CrmCalculos.semaforoPrazo', () => {
  const demanda = extra => CrmModel.normalizarAnotacao(Object.assign({ assunto: 'x' }, extra));

  it('vencida quando o prazo passou', () => {
    expect(CrmCalculos.semaforoPrazo(demanda({ prazo: '2026-07-29' }), HOJE)).toBe('vencida');
  });

  it('hoje no dia do prazo', () => {
    expect(CrmCalculos.semaforoPrazo(demanda({ prazo: HOJE }), HOJE)).toBe('hoje');
  });

  it('alerta dentro da antecedência pedida — o consumo real de lembrarDiasAntes', () => {
    expect(CrmCalculos.semaforoPrazo(demanda({ prazo: '2026-08-02', lembrarDiasAntes: 3 }), HOJE)).toBe('alerta');
  });

  it('ok logo fora da antecedência', () => {
    expect(CrmCalculos.semaforoPrazo(demanda({ prazo: '2026-08-03', lembrarDiasAntes: 3 }), HOJE)).toBe('ok');
  });

  it('sem lembrarDiasAntes não há alerta antecipado', () => {
    expect(CrmCalculos.semaforoPrazo(demanda({ prazo: '2026-08-02' }), HOJE)).toBe('ok');
  });

  it('demanda respondida ou sem prazo é sempre ok', () => {
    expect(CrmCalculos.semaforoPrazo(demanda({ prazo: '2026-01-01', situacao: 'respondida' }), HOJE)).toBe('ok');
    expect(CrmCalculos.semaforoPrazo(demanda({}), HOJE)).toBe('ok');
  });
});

describe('CrmCalculos.agruparPorSituacao', () => {
  it('junta recebida e comigo na mesma fila de trabalho', () => {
    const g = CrmCalculos.agruparPorSituacao([
      { situacao: 'recebida' }, { situacao: 'comigo' },
      { situacao: 'aguardando_terceiro' }, { situacao: 'consolidando' }, { situacao: 'respondida' }
    ]);
    expect(g.comigo.length).toBe(2);
    expect(g.aguardando.length).toBe(1);
    expect(g.consolidar.length).toBe(1);
    expect(g.concluidas.length).toBe(1);
  });
});

describe('CrmCalculos.filtrarAnotacoes', () => {
  const negocios = [{ id: 'ngc_1', funilId: 'fnl_A', clienteId: 'cli_1' }];
  const comNegocio = CrmModel.normalizarAnotacao({ id: 'a1', assunto: 'com negócio', negocioId: 'ngc_1' });
  const avulsa = CrmModel.normalizarAnotacao({ id: 'a2', assunto: 'avulsa', funilId: 'fnl_B', clienteId: 'cli_2' });
  const todas = [comNegocio, avulsa];
  const extra = { negocios, hoje: HOJE };

  it('resolve funil pelo negócio pai', () => {
    const r = CrmCalculos.filtrarAnotacoes(todas, { funilId: 'fnl_A' }, extra);
    expect(r.map(a => a.id)).toEqual(['a1']);
  });

  it('resolve funil pela própria demanda quando avulsa', () => {
    const r = CrmCalculos.filtrarAnotacoes(todas, { funilId: 'fnl_B' }, extra);
    expect(r.map(a => a.id)).toEqual(['a2']);
  });

  it('resolve cliente da demanda avulsa', () => {
    const r = CrmCalculos.filtrarAnotacoes(todas, { clienteId: 'cli_2' }, extra);
    expect(r.map(a => a.id)).toEqual(['a2']);
  });

  it('esconde demandas com soft delete', () => {
    const excluida = CrmModel.normalizarAnotacao({ id: 'a3', assunto: 'x', funilId: 'fnl_B', excluidoEm: '2026-07-01T00:00:00Z' });
    const r = CrmCalculos.filtrarAnotacoes(todas.concat([excluida]), {}, extra);
    expect(r.map(a => a.id)).toEqual(['a1', 'a2']);
  });

  it('filtra por situação e responsável', () => {
    const lista = [
      CrmModel.normalizarAnotacao({ id: 'b1', funilId: 'f', situacao: 'consolidando', responsavel: 'KOLTE' }),
      CrmModel.normalizarAnotacao({ id: 'b2', funilId: 'f', situacao: 'comigo', responsavel: 'KOLTE' }),
      CrmModel.normalizarAnotacao({ id: 'b3', funilId: 'f', situacao: 'consolidando', responsavel: 'ISA' })
    ];
    expect(CrmCalculos.filtrarAnotacoes(lista, { situacao: 'consolidando' }, extra).map(a => a.id)).toEqual(['b1', 'b3']);
    expect(CrmCalculos.filtrarAnotacoes(lista, { responsavel: 'ISA' }, extra).map(a => a.id)).toEqual(['b3']);
  });

  it('filtra por quem recebeu o encaminhamento', () => {
    const lista = [
      CrmModel.normalizarAnotacao({ id: 'c1', funilId: 'f', encaminhamentos: [{ para: 'Jurídico' }] }),
      CrmModel.normalizarAnotacao({ id: 'c2', funilId: 'f', encaminhamentos: [{ para: 'TI' }] })
    ];
    expect(CrmCalculos.filtrarAnotacoes(lista, { encaminhadoPara: 'TI' }, extra).map(a => a.id)).toEqual(['c2']);
  });

  it('somenteVencidas e somenteAtrasoTerceiro', () => {
    const lista = [
      CrmModel.normalizarAnotacao({ id: 'd1', funilId: 'f', prazo: '2026-07-01' }),
      CrmModel.normalizarAnotacao({ id: 'd2', funilId: 'f', prazo: '2026-12-01' }),
      CrmModel.normalizarAnotacao({ id: 'd3', funilId: 'f', encaminhamentos: [{ para: 'A', status: 'pendente', prazoResposta: '2026-07-01' }] })
    ];
    expect(CrmCalculos.filtrarAnotacoes(lista, { somenteVencidas: true }, extra).map(a => a.id)).toEqual(['d1']);
    expect(CrmCalculos.filtrarAnotacoes(lista, { somenteAtrasoTerceiro: true }, extra).map(a => a.id)).toEqual(['d3']);
  });

  it('a busca livre alcança o conteúdo dos encaminhamentos', () => {
    const lista = [
      CrmModel.normalizarAnotacao({ id: 'e1', assunto: 'nada', funilId: 'f', encaminhamentos: [{ para: 'X', resposta: 'Segue o parecer favoravel' }] }),
      CrmModel.normalizarAnotacao({ id: 'e2', assunto: 'nada', funilId: 'f' })
    ];
    expect(CrmCalculos.filtrarAnotacoes(lista, { busca: 'parecer' }, extra).map(a => a.id)).toEqual(['e1']);
  });
});

describe('CrmCalculos.ordenarAnotacoes', () => {
  const extra = { negocios: [], hoje: HOJE };

  it('ordena por situação seguindo o ciclo do processo', () => {
    const lista = [
      CrmModel.normalizarAnotacao({ id: 's1', situacao: 'respondida' }),
      CrmModel.normalizarAnotacao({ id: 's2', situacao: 'recebida' }),
      CrmModel.normalizarAnotacao({ id: 's3', situacao: 'aguardando_terceiro' })
    ];
    expect(CrmCalculos.ordenarAnotacoes(lista, 'situacao', 'asc', extra).map(a => a.id)).toEqual(['s2', 's3', 's1']);
  });

  it('ordena por dias aguardando, com quem não aguarda no fim', () => {
    const lista = [
      CrmModel.normalizarAnotacao({ id: 't1' }),
      CrmModel.normalizarAnotacao({ id: 't2', encaminhamentos: [{ para: 'A', status: 'pendente', dataEnvio: '2026-07-01' }] }),
      CrmModel.normalizarAnotacao({ id: 't3', encaminhamentos: [{ para: 'B', status: 'pendente', dataEnvio: '2026-07-28' }] })
    ];
    expect(CrmCalculos.ordenarAnotacoes(lista, 'diasAguardando', 'desc', extra).map(a => a.id)).toEqual(['t2', 't3', 't1']);
  });

  it('ordena por funil resolvendo demandas avulsas', () => {
    const negocios = [{ id: 'n1', funilId: 'fnl_z' }];
    const funis = [{ id: 'fnl_a', nome: 'Alfa' }, { id: 'fnl_z', nome: 'Zulu' }];
    const lista = [
      CrmModel.normalizarAnotacao({ id: 'u1', negocioId: 'n1' }),
      CrmModel.normalizarAnotacao({ id: 'u2', funilId: 'fnl_a' })
    ];
    const r = CrmCalculos.ordenarAnotacoes(lista, 'funil', 'asc', { negocios, funis, hoje: HOJE });
    expect(r.map(a => a.id)).toEqual(['u2', 'u1']);
  });
});
