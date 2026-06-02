// IMBEL · Tabela de Preços — gestão de períodos com simulador What-If
const { useState: useStateIP, useMemo: useMemoIP } = React;

function ImbelPrecos() {
  const [periodoId,      setPeriodoId]      = useStateIP('P2026-Q2');
  const [whatifOpen,     setWhatifOpen]     = useStateIP(true);
  const [showHistorico,  setShowHistorico]  = useStateIP(false);

  const periodo = IMBEL_PRECO_PERIODOS.find(p => p.id === periodoId);

  const [taxa, setTaxa] = useStateIP(periodo.taxa);
  const [roi,  setRoi]  = useStateIP(periodo.roi);
  const [desc, setDesc] = useStateIP(periodo.descTabelado);

  React.useEffect(() => {
    setTaxa(periodo.taxa); setRoi(periodo.roi); setDesc(periodo.descTabelado);
  }, [periodoId]);

  const ICMS_FACA = 18.4;
  function calcPrecoSede(ci, t, r) {
    const denom = 1 - t/100 - ICMS_FACA/100 - 0.05;
    if (denom <= 0.05) return null;
    return ci * (1 + r/100) / denom;
  }

  const baseTaxa = periodo.taxa;
  const baseRoi  = periodo.roi;
  const baseDesc = periodo.descTabelado;

  const rows = useMemoIP(() =>
    IMBEL_PRODUTOS.filter(p => p.status === 'ativo').map(p => {
      const pvBase   = calcPrecoSede(p.ci, baseTaxa, baseRoi);
      const pvNovo   = calcPrecoSede(p.ci, taxa, roi);
      const pvLojista = pvNovo * (1 - desc / 100);
      const margem   = ((pvLojista - p.ci) / pvLojista) * 100;
      return { ...p, pvBase, pvNovo, pvLojista, margem, delta: pvNovo - pvBase, deltaPct: pvBase ? (pvNovo - pvBase) / pvBase * 100 : 0 };
    }),
  [taxa, roi, desc, baseTaxa, baseRoi]);

  const kpis = useMemoIP(() => {
    const sumCI  = rows.reduce((s, r) => s + r.ci, 0);
    const sumPV  = rows.reduce((s, r) => s + r.pvNovo, 0);
    const sumLoj = rows.reduce((s, r) => s + r.pvLojista, 0);
    const avgMargem = rows.reduce((s, r) => s + r.margem, 0) / Math.max(1, rows.length);
    const sumBase = rows.reduce((s, r) => s + r.pvBase, 0);
    const totalDelta = sumPV - sumBase;
    return { sumCI, sumPV, sumLoj, avgMargem, totalDelta, totalDeltaPct: sumBase ? totalDelta / sumBase * 100 : 0, count: rows.length };
  }, [rows]);

  const taxaPct = ((taxa - 10) / (50 - 10)) * 100;
  const roiPct  = ((roi  - 0)  / (200 - 0)) * 100;
  const descPct = ((desc - 0)  / (50  - 0)) * 100;

  return h('div', { className: 'imbel-subtab' }, [

    // 0. PARAMS CARD (só Tabela de Preços)
    h('div', { className: 'imbel-params', key: 'pc' }, [
      h('div', { className: 'group', key: 1 }, [
        h('span', { className: 'lbl', key: 'l' }, 'Período ativo'),
        h('span', { className: 'pill', key: 'p' }, [h('span', { className: 'dot', key: 'd' }), periodo.label]),
        h('span', { style: { color: '#94a3b8', fontSize: 11 }, key: 'd' },
          `${new Date(periodo.dataInicio).toLocaleDateString('pt-BR')} – ${new Date(periodo.dataFim).toLocaleDateString('pt-BR')}`),
      ]),
      h('div', { className: 'group', key: 2 }, [
        h('span', { className: 'lbl', key: 'l' }, 'Taxa'),
        h('span', { className: 'val', key: 'v' }, `${taxa.toFixed(2)}%`),
        taxa !== baseTaxa && h('span', { style: { color: '#b8651f', fontSize: 10, fontFamily: 'IBM Plex Mono,monospace' }, key: 'd' }, `Δ ${(taxa-baseTaxa >= 0 ? '+' : '') + (taxa-baseTaxa).toFixed(2)}pp`),
      ]),
      h('div', { className: 'group', key: 3 }, [
        h('span', { className: 'lbl', key: 'l' }, 'ROI'),
        h('span', { className: 'val accent', key: 'v' }, `${roi}%`),
        roi !== baseRoi && h('span', { style: { color: '#b8651f', fontSize: 10, fontFamily: 'IBM Plex Mono,monospace' }, key: 'd' }, `Δ ${(roi-baseRoi >= 0 ? '+' : '') + (roi-baseRoi)}pp`),
      ]),
      h('div', { className: 'group', key: 4 }, [
        h('span', { className: 'lbl', key: 'l' }, 'Desc. Lojista'),
        h('span', { className: 'val', key: 'v', style: { color: '#b8651f' } }, `${desc}%`),
        desc !== baseDesc && h('span', { style: { color: '#b8651f', fontSize: 10, fontFamily: 'IBM Plex Mono,monospace' }, key: 'd' }, `Δ ${(desc-baseDesc >= 0 ? '+' : '') + (desc-baseDesc)}pp`),
      ]),
      h('div', { className: 'group', key: 5, style: { marginLeft: 'auto' } }, [
        h('button', { className: 'imbel-cmdbar-like-btn', onClick: () => setShowHistorico(true), key: 1, style: { height: 28, padding: '0 10px', fontSize: 11, border: '1px solid #e2e8f0', borderRadius: 4, background: '#fff', cursor: 'pointer' } }, `Histórico (${IMBEL_PRECO_PERIODOS.length})`),
        h('button', { className: 'imbel-cmdbar-like-btn', onClick: () => {}, key: 2, style: { height: 28, padding: '0 10px', fontSize: 11, border: '1px solid #e2e8f0', borderRadius: 4, background: '#fff', cursor: 'pointer' } }, '+ Novo período'),
        h('button', { key: 3, style: { height: 28, padding: '0 12px', fontSize: 11, border: 'none', borderRadius: 4, background: '#0f1e31', color: '#e5a128', cursor: 'pointer', fontWeight: 600 } }, 'Aplicar ao período'),
      ])
    ]),

    // 1. COMMAND BAR
    h('div', { className: 'imbel-cmdbar', key: 'cmd' }, [
      h('span', { className: 'filter-label', key: 'pl' }, 'Período'),
      h('select', {
        value: periodoId, onChange: e => setPeriodoId(e.target.value),
        key: 's', style: { fontFamily: 'IBM Plex Mono,monospace', fontSize: 12, fontWeight: 600, minWidth: 180 }
      }, IMBEL_PRECO_PERIODOS.map(p => h('option', { value: p.id, key: p.id },
        (p.active ? '● ' : '') + p.label + ' · ' + p.status
      ))),
      h('div', { className: 'divider', key: 'd1' }),
      h('div', { style: { fontSize: 11, color: '#94a3b8', fontFamily: 'IBM Plex Mono,monospace' }, key: 'h' }, [
        'ICMS faca SP ', h('b', { style: { color: '#0f1e31' }, key: 'b' }, ICMS_FACA + '%'),
        ' · base UF SP · ', h('b', { style: { color: '#0f1e31' }, key: 'r' }, kpis.count + ' produtos'),
      ]),
      h('div', { className: 'actions', key: 'ac' }, [
        h('button', { className: `btn ${whatifOpen ? 'accent' : ''}`, key: 1, onClick: () => setWhatifOpen(o => !o) }, 'What-If'),
        h('button', { className: 'btn icon', key: 2, title: 'Exportar' }, '↓'),
        h('button', { className: 'btn icon', key: 3, title: 'Imprimir' }, '⎙'),
      ])
    ]),

    // 3. KPI STRIP (sem type bar)
    h('div', { className: 'imbel-kpis', key: 'kpi' }, [
      h('div', { className: 'kpi', key: 1 }, [
        h('div', { className: 'kpi-label', key: 'l' }, 'Produtos no período'),
        h('div', { className: 'kpi-value', key: 'v' }, [kpis.count, h('span', { className: 'kpi-unit', key: 'u' }, ' SKU')]),
        h('div', { className: 'kpi-sub', key: 'd' }, 'IMBEL Sede'),
      ]),
      h('div', { className: 'kpi', key: 2 }, [
        h('div', { className: 'kpi-label', key: 'l' }, 'Σ CI'),
        h('div', { className: 'kpi-value accent', key: 'v' }, ['R$ ', fmt(kpis.sumCI, 0)]),
        h('div', { className: 'kpi-sub', key: 'd' }, 'custo agregado'),
      ]),
      h('div', { className: 'kpi', key: 3 }, [
        h('div', { className: 'kpi-label', key: 'l' }, 'Σ PV Sede'),
        h('div', { className: 'kpi-value', key: 'v' }, ['R$ ', fmt(kpis.sumPV, 0)]),
        h('div', { className: 'kpi-sub', key: 'd' }, 'preço de balcão IMBEL'),
      ]),
      h('div', { className: 'kpi', key: 4 }, [
        h('div', { className: 'kpi-label', key: 'l' }, 'Margem média'),
        h('div', { className: 'kpi-value pos', key: 'v' }, [kpis.avgMargem.toFixed(1), h('span', { className: 'kpi-unit', key: 'u' }, ' %')]),
        h('div', { className: 'kpi-sub', key: 'd' }, 'CI → PV lojista'),
      ]),
      h('div', { className: 'kpi', key: 5 }, [
        h('div', { className: 'kpi-label', key: 'l' }, ['Impacto vs base ', h('span', { className: 'kpi-tag', key: 't' }, periodo.label)]),
        h('div', { className: `kpi-value ${kpis.totalDelta >= 0 ? 'pos' : 'neg'}`, key: 'v' }, [
          kpis.totalDelta >= 0 ? '+' : '', fmt(kpis.totalDeltaPct, 2), h('span', { className: 'kpi-unit', key: 'u' }, ' %')
        ]),
        h('div', { className: `kpi-sub`, key: 'd' }, `${kpis.totalDelta >= 0 ? '+' : ''}R$ ${fmt(kpis.totalDelta, 0)} agregado`),
      ]),
    ]),

    // 4. BODY (tabela + painel what-if lateral)
    h('div', { className: `imbel-preco-main${whatifOpen ? ' with-panel' : ''}`, key: 'body' }, [
      h('div', { className: 'imbel-preco-workspace', key: 'ws' },
        h('table', { className: 'price-table' }, [
          h('thead', { key: 'h' }, h('tr', {}, [
            h('th', { key: 1, style: { minWidth: 70  } }, 'ID'),
            h('th', { key: 2, style: { minWidth: 100 } }, 'PN'),
            h('th', { key: 3, style: { minWidth: 220 } }, 'PRODUTO'),
            h('th', { key: 4, style: { minWidth: 120 } }, 'CATEGORIA'),
            h('th', { className: 'num', key: 5, style: { background: 'linear-gradient(180deg,#0f1e31,#4d2818 50%)' } }, 'CI'),
            h('th', { className: 'num', key: 6 }, ['PV BASE ', h('span', { style: { color: '#6b7a94', fontWeight: 400 }, key: 's' }, periodo.label)]),
            h('th', { className: 'num', key: 7, style: { background: taxa !== baseTaxa || roi !== baseRoi ? 'linear-gradient(180deg,#0f1e31,#4d3d1f 50%)' : undefined } }, [
              'PV NOVO',
              (taxa !== baseTaxa || roi !== baseRoi) && h('span', { style: { color: '#e5a128', fontWeight: 400, marginLeft: 4 }, key: 's' }, 'sim.')
            ]),
            h('th', { className: 'num', key: 8 }, 'Δ ABS'),
            h('th', { className: 'num', key: 9 }, 'Δ %'),
            h('th', { className: 'num', key: 10 }, ['PV LOJISTA ', h('span', { style: { color: '#6b7a94', fontWeight: 400 }, key: 's' }, '-' + desc + '%')]),
            h('th', { className: 'num', key: 11 }, 'MARGEM'),
            h('th', { className: 'num', key: 12 }, 'ESTOQUE'),
          ])),
          h('tbody', { key: 'b' },
            rows.map(p => h('tr', { key: p.id }, [
              h('td', { className: 'ncm',     key: 1 }, p.id),
              h('td', { className: 'pn',      key: 2 }, p.pn),
              h('td', { className: 'name',    key: 3 }, p.nome),
              h('td', { className: 'grupo',   key: 4 }, p.categoria),
              h('td', { className: 'num ci',  key: 5 }, fmt(p.ci)),
              h('td', { className: 'num price med', key: 6 }, fmt(p.pvBase)),
              h('td', { className: 'num', key: 7, style: { color: p.delta !== 0 ? '#92400e' : '#0f1e31', fontWeight: 700, background: p.delta !== 0 ? '#fef3c7' : undefined } }, fmt(p.pvNovo)),
              h('td', { className: 'num', key: 8, style: { color: p.delta > 0 ? '#16a34a' : p.delta < 0 ? '#dc2626' : '#94a3b8' } },
                p.delta === 0 ? '—' : (p.delta > 0 ? '+' : '') + 'R$ ' + fmt(p.delta)),
              h('td', { className: 'num', key: 9, style: { color: p.deltaPct > 0 ? '#16a34a' : p.deltaPct < 0 ? '#dc2626' : '#94a3b8' } },
                p.deltaPct === 0 ? '—' : (p.deltaPct > 0 ? '+' : '') + p.deltaPct.toFixed(2) + '%'),
              h('td', { className: 'num', key: 10, style: { fontWeight: 700 } }, fmt(p.pvLojista)),
              h('td', { className: 'num', key: 11, style: { color: p.margem > 30 ? '#16a34a' : '#b8651f', fontWeight: 600 } }, p.margem.toFixed(1) + '%'),
              h('td', { className: 'num', key: 12, style: { color: p.estoqueGalpao < p.estoqueMin ? '#dc2626' : '#64748b', fontWeight: p.estoqueGalpao < p.estoqueMin ? 600 : 400 } }, fmt(p.estoqueGalpao, 0)),
            ]))
          )
        ])
      ),
      whatifOpen && h('div', { className: 'imbel-whatif', key: 'wi' }, [
        h('div', { className: 'imbel-whatif-head', key: 'h' }, [
          h('div', { className: 'ttl', key: 't' }, 'What-If · ' + periodo.label),
          h('div', { style: { display: 'flex', gap: 6, alignItems: 'center' }, key: 'r' }, [
            h('span', { className: 'live', key: 'l' }, [h('span', { className: 'dot', key: 'd' }), 'live']),
            h('button', { onClick: () => setWhatifOpen(false), title: 'Fechar', key: 'x', style: { background: 'none', border: 'none', cursor: 'pointer', color: '#94a3b8', fontSize: 16 } }, '✕'),
          ])
        ]),
        h('div', { className: 'imbel-whatif-body', key: 'b' }, [
          h('div', { key: 'glb' }, [
            h('div', { className: 'imbel-sec-ttl', key: 't' }, [
              'Parâmetros',
              (taxa !== baseTaxa || roi !== baseRoi || desc !== baseDesc) &&
                h('button', { className: 'reset', onClick: () => { setTaxa(baseTaxa); setRoi(baseRoi); setDesc(baseDesc); }, key: 'r' }, '↺ reset'),
            ]),
            // Taxa
            h('div', { className: 'imbel-param-row', key: 'taxa' }, [
              h('div', { className: 'top', key: 't' }, [
                h('span', { className: 'name', key: 'n' }, ['Taxa %', h('span', { className: 'hint', key: 'h' }, ' impostos+simples')]),
                h('span', { className: 'value', key: 'v' }, [taxa.toFixed(2), h('span', { className: 'unit', key: 'u' }, '%'),
                  taxa !== baseTaxa && h('span', { className: 'delta ' + (taxa - baseTaxa > 0 ? 'pos' : 'neg'), key: 'd' }, `${taxa-baseTaxa > 0 ? '+' : ''}${(taxa-baseTaxa).toFixed(2)}pp`)
                ]),
              ]),
              h('input', { type: 'range', min: 10, max: 50, step: 0.01, value: taxa, style: { '--p': taxaPct + '%' }, onChange: e => setTaxa(parseFloat(e.target.value)), key: 'i' }),
              h('div', { className: 'ticks', key: 'tk' }, ['10%','20%','30%','40%','50%'].map((t,i) => h('span',{key:i},t))),
            ]),
            // ROI
            h('div', { className: 'imbel-param-row', key: 'roi' }, [
              h('div', { className: 'top', key: 't' }, [
                h('span', { className: 'name', key: 'n' }, ['ROI %', h('span', { className: 'hint', key: 'h' }, ' margem sobre CI')]),
                h('span', { className: 'value', key: 'v' }, [roi.toFixed(0), h('span', { className: 'unit', key: 'u' }, '%'),
                  roi !== baseRoi && h('span', { className: 'delta ' + (roi - baseRoi > 0 ? 'pos' : 'neg'), key: 'd' }, `${roi-baseRoi > 0 ? '+' : ''}${roi-baseRoi}pp`)
                ]),
              ]),
              h('input', { type: 'range', min: 0, max: 200, step: 1, value: roi, style: { '--p': roiPct + '%' }, onChange: e => setRoi(parseFloat(e.target.value)), key: 'i' }),
              h('div', { className: 'ticks', key: 'tk' }, ['0%','50%','100%','150%','200%'].map((t,i) => h('span',{key:i},t))),
            ]),
            // Desc
            h('div', { className: 'imbel-param-row', key: 'desc' }, [
              h('div', { className: 'top', key: 't' }, [
                h('span', { className: 'name', key: 'n' }, ['Desc. Lojista', h('span', { className: 'hint', key: 'h' }, ' sobre PV')]),
                h('span', { className: 'value', key: 'v' }, [desc.toFixed(0), h('span', { className: 'unit', key: 'u' }, '%'),
                  desc !== baseDesc && h('span', { className: 'delta ' + (desc - baseDesc > 0 ? 'neg' : 'pos'), key: 'd' }, `${desc-baseDesc > 0 ? '+' : ''}${desc-baseDesc}pp`)
                ]),
              ]),
              h('input', { type: 'range', min: 0, max: 50, step: 1, value: desc, style: { '--p': descPct + '%' }, onChange: e => setDesc(parseFloat(e.target.value)), key: 'i' }),
              h('div', { className: 'ticks', key: 'tk' }, ['0%','15%','30%','50%'].map((t,i) => h('span',{key:i},t))),
            ]),
          ]),
          h('div', { key: 'imp' }, [
            h('div', { className: 'imbel-sec-ttl', key: 't' }, 'Impacto sobre o catálogo'),
            h('div', { className: 'imbel-impact-grid', key: 'g' }, [
              h('div', { className: 'imbel-impact-cell', key: 1 }, [h('div',{className:'l'},'Σ PV Sede'), h('div',{className:'v'},'R$ '+fmt(kpis.sumPV,0)), h('div',{className:`d ${kpis.totalDelta>=0?'pos':'neg'}`},`${kpis.totalDelta>=0?'+':''}R$ ${fmt(kpis.totalDelta,0)}`)]),
              h('div', { className: 'imbel-impact-cell', key: 2 }, [h('div',{className:'l'},'Σ PV Lojista'), h('div',{className:'v'},'R$ '+fmt(kpis.sumLoj,0)), h('div',{className:'d'},`-${desc}%`)]),
              h('div', { className: 'imbel-impact-cell', key: 3 }, [h('div',{className:'l'},'Margem méd'), h('div',{className:'v',style:{color:kpis.avgMargem>30?'#16a34a':'#b8651f'}},kpis.avgMargem.toFixed(1)+'%'), h('div',{className:'d'},'CI → lojista')]),
              h('div', { className: 'imbel-impact-cell', key: 4 }, [h('div',{className:'l'},'Δ vs base'), h('div',{className:'v',style:{color:kpis.totalDelta>=0?'#16a34a':'#dc2626'}},`${kpis.totalDeltaPct>=0?'+':''}${kpis.totalDeltaPct.toFixed(2)}%`), h('div',{className:'d'},periodo.label)]),
            ])
          ]),
        ]),
        h('div', { className: 'imbel-whatif-foot', key: 'f' }, [
          h('button', { onClick: () => { setTaxa(baseTaxa); setRoi(baseRoi); setDesc(baseDesc); }, key: 1, style: { flex: 1, height: 30, border: '1px solid #e2e8f0', borderRadius: 4, background: '#fff', cursor: 'pointer', fontSize: 12 } }, '↺ Descartar'),
          h('button', { key: 2, style: { flex: 1, height: 30, border: 'none', borderRadius: 4, background: '#0f1e31', color: '#e5a128', cursor: 'pointer', fontSize: 12, fontWeight: 600 } }, 'Salvar rascunho'),
        ])
      ]),
    ]),

    // 5. FOOTBAR
    h('div', { className: 'imbel-footbar', key: 'foot' }, [
      h('div', { className: 'seg', key: 1 }, [h('span', { className: 'lbl' }, 'PERÍODO'), h('span', { className: 'val accent' }, periodo.label)]),
      h('div', { className: 'divider', key: 'd1' }),
      h('div', { className: 'seg', key: 2 }, [h('span', { className: 'lbl' }, 'SKU'), h('span', { className: 'val' }, kpis.count)]),
      h('div', { className: 'divider', key: 'd2' }),
      h('div', { className: 'seg', key: 3 }, [h('span', { className: 'lbl' }, 'Σ PV'), h('span', { className: 'val' }, 'R$ ' + fmt(kpis.sumPV / 1000, 0) + 'k')]),
      h('div', { className: 'divider', key: 'd3' }),
      h('div', { className: 'seg', key: 4 }, [h('span', { className: 'lbl' }, 'MARGEM'), h('span', { className: 'val' }, kpis.avgMargem.toFixed(1) + '%')]),
      h('div', { className: 'divider', key: 'd4' }),
      h('div', { className: 'seg', key: 5 }, [h('span', { className: 'lbl' }, 'Δ'), h('span', { className: `val ${kpis.totalDelta !== 0 ? 'accent' : ''}` }, kpis.totalDelta === 0 ? '—' : (kpis.totalDelta > 0 ? '+' : '') + 'R$ ' + fmt(kpis.totalDelta, 0))]),
      h('div', { className: 'right', key: 'rt' },
        h('div', { className: 'keys' }, [
          h('span', { key: 1 }, [h('b', { key: 'b' }, 'W'), ' What-If']),
          h('span', { key: 2 }, [h('b', { key: 'b' }, 'N'), ' Novo período']),
          h('span', { key: 3 }, [h('b', { key: 'b' }, '⌘S'), ' Salvar']),
        ])
      )
    ])
  ]);
}

window.ImbelPrecos = ImbelPrecos;
