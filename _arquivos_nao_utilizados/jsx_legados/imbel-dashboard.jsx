// IMBEL · Dashboard — visão geral do controle de facas
const { useState: useStateID, useMemo: useMemoID } = React;

function ImbelDashboard() {
  const [period, setPeriod] = useStateID('30d');

  const kpis = useMemoID(() => {
    const totalEstoque = IMBEL_PRODUTOS.reduce((s, p) => s + p.estoqueGalpao, 0);
    const totalSKUs = IMBEL_PRODUTOS.filter(p => p.status === 'ativo').length;
    const valorEstoque = IMBEL_PRODUTOS.reduce((s, p) => s + p.ci * p.estoqueGalpao, 0);
    const abaixoMin = IMBEL_PRODUTOS.filter(p => p.status === 'ativo' && p.estoqueGalpao < p.estoqueMin).length;
    const semEstoque = IMBEL_PRODUTOS.filter(p => p.status === 'ativo' && p.estoqueGalpao === 0).length;
    const corte = new Date(Date.now() - 30 * 24 * 3600 * 1000).toISOString().slice(0, 10);
    const movsRecentes = IMBEL_MOVIMENTACOES.filter(m => m.timestamp >= corte);
    const entradas30 = movsRecentes.filter(m => m.tipo === 'entrada').reduce((s, m) => s + m.quantidade, 0);
    const saidas30 = movsRecentes.filter(m => m.tipo === 'saida').reduce((s, m) => s + Math.abs(m.quantidade), 0);
    const fat30 = movsRecentes.filter(m => m.tipo === 'saida').reduce((s, m) => s + m.valorTotal, 0);
    return { totalEstoque, totalSKUs, valorEstoque, abaixoMin, semEstoque, entradas30, saidas30, fat30, ndocs: movsRecentes.length };
  }, [period]);

  const topGiro = useMemoID(() => {
    const giro = {};
    IMBEL_MOVIMENTACOES.filter(m => m.tipo === 'saida').forEach(m => {
      giro[m.produtoId] = (giro[m.produtoId] || 0) + Math.abs(m.quantidade);
    });
    return Object.entries(giro)
      .map(([id, qtd]) => ({ ...IMBEL_PRODUTOS.find(p => p.id === id), giro: qtd }))
      .filter(p => p && p.id)
      .sort((a, b) => b.giro - a.giro)
      .slice(0, 6);
  }, []);

  const porCategoria = useMemoID(() => {
    const map = {};
    IMBEL_PRODUTOS.forEach(p => {
      if (p.status !== 'ativo') return;
      if (!map[p.categoria]) map[p.categoria] = { cat: p.categoria, qtd: 0, valor: 0, skus: 0 };
      map[p.categoria].qtd += p.estoqueGalpao;
      map[p.categoria].valor += p.estoqueGalpao * p.ci;
      map[p.categoria].skus += 1;
    });
    return Object.values(map).sort((a, b) => b.valor - a.valor);
  }, []);

  const alertas = useMemoID(() =>
    IMBEL_PRODUTOS
      .filter(p => p.status === 'ativo' && p.estoqueGalpao < p.estoqueMin)
      .sort((a, b) => (a.estoqueGalpao / Math.max(1, a.estoqueMin)) - (b.estoqueGalpao / Math.max(1, b.estoqueMin))),
  []);

  const serieMov = useMemoID(() => {
    const dias = [];
    for (let i = 13; i >= 0; i--) {
      const d = new Date(Date.now() - i * 24 * 3600 * 1000);
      const k = d.toISOString().slice(0, 10);
      dias.push({ key: k, label: d.toLocaleDateString('pt-BR', { day: '2-digit', month: '2-digit' }), entrada: 0, saida: 0 });
    }
    IMBEL_MOVIMENTACOES.forEach(m => {
      const k = m.timestamp.slice(0, 10);
      const d = dias.find(x => x.key === k);
      if (!d) return;
      if (m.tipo === 'entrada') d.entrada += m.quantidade;
      if (m.tipo === 'saida')   d.saida   += Math.abs(m.quantidade);
    });
    return dias;
  }, []);

  const maxSerie = Math.max(...serieMov.map(d => Math.max(d.entrada, d.saida)), 100);
  const totalRepEstoque = IMBEL_REP_COMPARE.reduce((s, r) => s + r.estoque, 0);

  return h('div', { className: 'imbel-subtab' }, [

    // 1. COMMAND BAR
    h('div', { className: 'imbel-cmdbar', key: 'cmd' }, [
      h('div', { style: { display: 'flex', alignItems: 'center', gap: 8 }, key: 'h' }, [
        h('span', { className: 'imbel-dot-large', key: 'd' }),
        h('span', { style: { fontFamily: 'var(--tv-font-display)', fontSize: 13, fontWeight: 600, color: '#0f1e31', letterSpacing: 0.5 }, key: 'tt' }, 'CONTROLE DE FACAS — VISÃO GERAL'),
      ]),
      h('div', { className: 'divider', key: 'd1' }),
      h('div', { className: 'seg', key: 'p' }, [
        h('button', { className: period === '7d'  ? 'active' : '', onClick: () => setPeriod('7d'),  key: 1 }, '7d'),
        h('button', { className: period === '30d' ? 'active' : '', onClick: () => setPeriod('30d'), key: 2 }, '30d'),
        h('button', { className: period === '90d' ? 'active' : '', onClick: () => setPeriod('90d'), key: 3 }, '90d'),
        h('button', { className: period === 'ytd' ? 'active' : '', onClick: () => setPeriod('ytd'), key: 4 }, 'YTD'),
      ]),
      h('div', { className: 'actions', key: 'ac' }, [
        h('button', { className: 'btn', key: 1 }, [
          h('svg', { width: 14, height: 14, viewBox: '0 0 24 24', fill: 'none', stroke: 'currentColor', strokeWidth: 2, key: 'i' },
            h('polyline', { points: '23 4 23 10 17 10' }),
            h('path', { d: 'M20.49 15a9 9 0 1 1-2.12-9.36L23 10' })
          ),
          'Atualizar'
        ]),
        h('button', { className: 'btn', key: 2 }, [
          h('svg', { width: 14, height: 14, viewBox: '0 0 24 24', fill: 'none', stroke: 'currentColor', strokeWidth: 2, key: 'i' },
            h('path', { d: 'M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4' }),
            h('polyline', { points: '7 10 12 15 17 10' }),
            h('line', { x1: 12, y1: 15, x2: 12, y2: 3 })
          ),
          'Exportar relatório'
        ]),
        h('button', { className: 'btn icon', key: 3, title: 'Imprimir' },
          h('svg', { width: 14, height: 14, viewBox: '0 0 24 24', fill: 'none', stroke: 'currentColor', strokeWidth: 2 },
            h('polyline', { points: '6 9 6 2 18 2 18 9' }),
            h('path', { d: 'M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2' }),
            h('rect', { x: 6, y: 14, width: 12, height: 8 })
          )
        ),
      ])
    ]),

    // 3. KPI STRIP (sem type bar no Dashboard)
    h('div', { className: 'imbel-kpis', key: 'kpi' }, [
      h('div', { className: 'kpi', key: 1 }, [
        h('div', { className: 'kpi-label', key: 'l' }, 'Estoque galpão IMBEL'),
        h('div', { className: 'kpi-value accent', key: 'v' }, [fmt(kpis.totalEstoque, 0), h('span', { className: 'kpi-unit', key: 'u' }, ' un')]),
        h('div', { className: 'kpi-sub', key: 'd' }, [kpis.totalSKUs, ' SKUs ativos']),
      ]),
      h('div', { className: 'kpi', key: 2 }, [
        h('div', { className: 'kpi-label', key: 'l' }, 'Valor em estoque (CI)'),
        h('div', { className: 'kpi-value pos', key: 'v' }, ['R$ ', fmt(kpis.valorEstoque / 1000, 0), h('span', { className: 'kpi-unit', key: 'u' }, ' mil')]),
        h('div', { className: 'kpi-sub', key: 'd' }, 'CI × qtd em galpão'),
      ]),
      h('div', { className: 'kpi', key: 3 }, [
        h('div', { className: 'kpi-label', key: 'l' }, ['Movimentações ', h('span', { className: 'kpi-tag', key: 't' }, period)]),
        h('div', { className: 'kpi-value', key: 'v', style: { fontSize: 18 } }, [
          h('span', { style: { color: '#16a34a' }, key: 'e' }, '+' + fmt(kpis.entradas30, 0)),
          h('span', { style: { color: '#cbd5e1', margin: '0 5px' }, key: 's' }, '/'),
          h('span', { style: { color: '#dc2626' }, key: 'a' }, '-' + fmt(kpis.saidas30, 0)),
        ]),
        h('div', { className: 'kpi-sub', key: 'd' }, [kpis.ndocs + ' lançamentos · saldo +', fmt(kpis.entradas30 - kpis.saidas30, 0)]),
      ]),
      h('div', { className: 'kpi', key: 4 }, [
        h('div', { className: 'kpi-label', key: 'l' }, ['Faturamento ', h('span', { className: 'kpi-tag', key: 't' }, period)]),
        h('div', { className: 'kpi-value', key: 'v' }, ['R$ ', fmt(kpis.fat30 / 1000, 0), h('span', { className: 'kpi-unit', key: 'u' }, ' mil')]),
        h('div', { className: 'kpi-sub', key: 'd' }, 'ticket méd. R$ ' + fmt(kpis.fat30 / Math.max(1, IMBEL_MOVIMENTACOES.filter(m => m.tipo === 'saida').length), 0)),
      ]),
      h('div', { className: 'kpi', key: 5 }, [
        h('div', { className: 'kpi-label', key: 'l' }, ['Alertas de estoque', alertas.length > 0 && h('span', { className: 'kpi-tag', key: 't', style: { background: '#fee2e2', color: '#dc2626' } }, 'crítico')]),
        h('div', { className: `kpi-value ${alertas.length > 0 ? 'neg' : ''}`, key: 'v' }, alertas.length),
        h('div', { className: 'kpi-sub', key: 'd' }, [kpis.semEstoque + ' sem estoque · ', alertas.length - kpis.semEstoque, ' abaixo do mín']),
      ]),
    ]),

    // 4. BODY
    h('div', { className: 'imbel-body', key: 'body' },
      h('div', { className: 'imbel-dash-grid' }, [
        h('div', { className: 'imbel-dash-col', key: 'L' }, [
          h('div', { className: 'imbel-card', key: 'mc' }, [
            h('div', { className: 'imbel-card-head', key: 'h' }, [
              h('div', { className: 'imbel-card-ttl', key: 't' }, 'Movimentações — últimos 14 dias'),
              h('div', { className: 'imbel-card-legend', key: 'l' }, [
                h('span', { key: 1 }, [h('i', { style: { background: '#16a34a' }, key: 'i' }), 'Entradas']),
                h('span', { key: 2 }, [h('i', { style: { background: '#dc2626' }, key: 'i' }), 'Saídas']),
              ])
            ]),
            h('div', { className: 'imbel-bars', key: 'b' },
              serieMov.map(d =>
                h('div', { key: d.key, className: 'imbel-bar', title: d.label }, [
                  h('div', { className: 'imbel-bar-stack', key: 's' }, [
                    h('div', { className: 'imbel-bar-e', style: { height: ((d.entrada / maxSerie) * 100).toFixed(1) + '%' }, key: 'e' }),
                    h('div', { className: 'imbel-bar-s', style: { height: ((d.saida   / maxSerie) * 100).toFixed(1) + '%' }, key: 's2' }),
                  ]),
                  h('div', { className: 'imbel-bar-lbl', key: 'l' }, d.label),
                ])
              )
            )
          ]),
          h('div', { className: 'imbel-card', key: 'tg' }, [
            h('div', { className: 'imbel-card-head', key: 'h' }, [
              h('div', { className: 'imbel-card-ttl', key: 't' }, 'Top produtos · maior giro (saídas)'),
              h('div', { className: 'imbel-card-sub', key: 's' }, 'unidades vendidas no período'),
            ]),
            h('div', { className: 'imbel-top-list', key: 'l' },
              topGiro.map((p, i) => {
                const max = topGiro[0].giro;
                return h('div', { className: 'imbel-top-row', key: p.id }, [
                  h('div', { className: 'imbel-top-rank', key: 'r' }, '#' + (i + 1)),
                  h('div', { key: 'b' }, [
                    h('div', { className: 'imbel-top-name', key: 'n' }, p.nome),
                    h('div', { className: 'imbel-top-meta', key: 'm' }, [h('span', { key: 1 }, p.id), h('span', { key: 2 }, p.categoria)])
                  ]),
                  h('div', { className: 'imbel-top-bar', key: 'br' },
                    h('div', { className: 'fill', style: { width: ((p.giro / max) * 100).toFixed(0) + '%' } })),
                  h('div', { className: 'imbel-top-val', key: 'v' }, [fmt(p.giro, 0), h('span', { key: 'u' }, ' un')]),
                ]);
              })
            )
          ])
        ]),
        h('div', { className: 'imbel-dash-col', key: 'R' }, [
          h('div', { className: 'imbel-card', key: 'cat' }, [
            h('div', { className: 'imbel-card-head', key: 'h' }, [
              h('div', { className: 'imbel-card-ttl', key: 't' }, 'Estoque por categoria'),
              h('div', { className: 'imbel-card-sub', key: 's' }, fmt(kpis.totalEstoque, 0) + ' unidades · ' + porCategoria.length + ' categorias'),
            ]),
            h('div', { className: 'imbel-cat-list', key: 'l' },
              porCategoria.map(c => h('div', { className: 'imbel-cat-row', key: c.cat }, [
                h('div', { className: 'imbel-cat-head', key: 'h' }, [
                  h('div', { className: 'imbel-cat-nm', key: 'n' }, c.cat),
                  h('div', { className: 'imbel-cat-vl', key: 'v' }, [
                    fmt(c.qtd, 0), h('span', { className: 'unit', key: 'u' }, ' un'),
                    h('span', { className: 'skus', key: 's' }, c.skus + ' SKU'),
                  ])
                ]),
                h('div', { className: 'imbel-cat-bar', key: 'b' }, [
                  h('div', { className: 'fill', style: { width: ((c.qtd / kpis.totalEstoque) * 100).toFixed(1) + '%' }, key: 'f' }),
                  h('div', { className: 'pct', key: 'p' }, ((c.qtd / kpis.totalEstoque) * 100).toFixed(1) + '%')
                ]),
                h('div', { className: 'imbel-cat-vl-money', key: 'm' }, 'R$ ' + fmt(c.valor / 1000, 0) + ' mil')
              ]))
            )
          ]),
          alertas.length > 0 && h('div', { className: 'imbel-card alert', key: 'al' }, [
            h('div', { className: 'imbel-card-head', key: 'h' }, [
              h('div', { className: 'imbel-card-ttl', key: 't', style: { color: '#dc2626' } }, '⚡ Reposição urgente'),
              h('div', { className: 'imbel-card-sub', key: 's' }, alertas.length + ' produtos abaixo do mínimo'),
            ]),
            h('div', { className: 'imbel-alert-list', key: 'l' },
              alertas.slice(0, 5).map(p => {
                const pct = p.estoqueMin > 0 ? (p.estoqueGalpao / p.estoqueMin) * 100 : 0;
                return h('div', { className: 'imbel-alert-row', key: p.id }, [
                  h('div', { className: 'imbel-alert-id', key: 'i' }, p.id),
                  h('div', { key: 'b' }, [
                    h('div', { className: 'imbel-alert-nm', key: 'n' }, p.nome),
                    h('div', { className: 'imbel-alert-bar', key: 'r' },
                      h('div', { className: 'fill', style: { width: Math.min(100, pct).toFixed(0) + '%', background: pct === 0 ? '#dc2626' : pct < 50 ? '#f87171' : '#e5a128' }, key: 'f' })
                    )
                  ]),
                  h('div', { className: 'imbel-alert-val', key: 'v' }, [
                    h('div', { className: 'cur', key: 'c', style: { color: p.estoqueGalpao === 0 ? '#dc2626' : '#b8651f' } }, p.estoqueGalpao),
                    h('div', { className: 'min', key: 'm' }, ['/ mín ', p.estoqueMin])
                  ])
                ]);
              })
            )
          ]),
          h('div', { className: 'imbel-card', key: 'rc' }, [
            h('div', { className: 'imbel-card-head', key: 'h' }, [
              h('div', { className: 'imbel-card-ttl', key: 't' }, 'IMBEL vs outros reps'),
              h('div', { className: 'imbel-card-sub', key: 's' }, 'Posição relativa no portfólio'),
            ]),
            h('div', { className: 'imbel-rep-list', key: 'l' },
              IMBEL_REP_COMPARE.map(r => h('div', { className: 'imbel-rep-row', key: r.rep }, [
                h('div', { className: 'imbel-rep-name', key: 'n', style: { borderLeftColor: r.cor, fontWeight: r.rep === 'IMBEL' ? 700 : 500 } }, [
                  r.rep, r.rep === 'IMBEL' && h('span', { className: 'rep-badge', key: 'b' }, 'aqui')
                ]),
                h('div', { className: 'imbel-rep-bar', key: 'b' },
                  h('div', { className: 'fill', style: { width: ((r.estoque / totalRepEstoque) * 100).toFixed(1) + '%', background: r.cor } })),
                h('div', { className: 'imbel-rep-val', key: 'v' }, [
                  h('div', { key: 1 }, fmt(r.estoque, 0)),
                  h('div', { className: 'sub', key: 2 }, 'R$ ' + fmt(r.faturamento30d / 1000, 0) + 'k')
                ]),
              ]))
            )
          ])
        ])
      ])
    ),

    // 5. FOOTBAR
    h('div', { className: 'imbel-footbar', key: 'foot' }, [
      h('div', { className: 'seg', key: 1 }, [h('span', { className: 'lbl' }, 'SKU'), h('span', { className: 'val accent' }, kpis.totalSKUs)]),
      h('div', { className: 'divider', key: 'd1' }),
      h('div', { className: 'seg', key: 2 }, [h('span', { className: 'lbl' }, 'ESTOQUE'), h('span', { className: 'val' }, fmt(kpis.totalEstoque, 0))]),
      h('div', { className: 'divider', key: 'd2' }),
      h('div', { className: 'seg', key: 3 }, [h('span', { className: 'lbl' }, 'VALOR'), h('span', { className: 'val' }, 'R$ ' + fmt(kpis.valorEstoque / 1000, 0) + 'k')]),
      h('div', { className: 'divider', key: 'd3' }),
      h('div', { className: 'seg', key: 4 }, [h('span', { className: 'lbl' }, 'ALERTAS'), h('span', { className: `val ${alertas.length > 0 ? 'accent' : ''}` }, alertas.length)]),
      h('div', { className: 'right', key: 'rt' },
        h('div', { className: 'keys' }, [
          h('span', { key: 1 }, [h('b', { key: 'b' }, 'R'), ' Atualizar']),
          h('span', { key: 2 }, [h('b', { key: 'b' }, '⌥M'), ' Trocar módulo']),
          h('span', { key: 3 }, [h('b', { key: 'b' }, '⌥[/]'), ' Aba']),
        ])
      )
    ])
  ]);
}

window.ImbelDashboard = ImbelDashboard;
