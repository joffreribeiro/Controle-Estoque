// IMBEL · Cadastro — gestão de produtos com edição inline e agrupamento por categoria
const { useState: useStateIC, useMemo: useMemoIC } = React;

function ImbelCadastro() {
  const [search,       setSearch]       = useStateIC('');
  const [filterCat,    setFilterCat]    = useStateIC(new Set());
  const [filterStatus, setFilterStatus] = useStateIC('ativo');
  const [collapsed,    setCollapsed]    = useStateIC(new Set());
  const [editingId,    setEditingId]    = useStateIC(null);
  const [edits,        setEdits]        = useStateIC({});
  const [showImport,   setShowImport]   = useStateIC(false);
  const [bulkSelected, setBulkSelected] = useStateIC(new Set());

  const filtered = useMemoIC(() => {
    let arr = IMBEL_PRODUTOS;
    if (search.trim()) {
      const q = search.toLowerCase();
      arr = arr.filter(p =>
        p.nome.toLowerCase().includes(q) ||
        p.fabrica.toLowerCase().includes(q) ||
        p.pn.includes(q) ||
        p.id.toLowerCase().includes(q) ||
        p.ncm.includes(q)
      );
    }
    if (filterCat.size) arr = arr.filter(p => filterCat.has(p.categoria));
    if (filterStatus !== 'todos') arr = arr.filter(p => p.status === filterStatus);
    return arr;
  }, [search, filterCat, filterStatus]);

  const grouped = useMemoIC(() => {
    const map = new Map();
    filtered.forEach(p => {
      if (!map.has(p.categoria)) map.set(p.categoria, []);
      map.get(p.categoria).push(p);
    });
    return [...map.entries()].sort((a, b) => a[0].localeCompare(b[0]));
  }, [filtered]);

  const stats = useMemoIC(() => {
    const sumEst  = IMBEL_PRODUTOS.reduce((s, p) => s + p.estoqueGalpao, 0);
    const sumVal  = IMBEL_PRODUTOS.reduce((s, p) => s + p.estoqueGalpao * p.ci, 0);
    const ativos  = IMBEL_PRODUTOS.filter(p => p.status === 'ativo').length;
    const inativos= IMBEL_PRODUTOS.filter(p => p.status === 'inativo').length;
    const baixoEst= IMBEL_PRODUTOS.filter(p => p.status === 'ativo' && p.estoqueGalpao < p.estoqueMin).length;
    return { sumEst, sumVal, ativos, inativos, baixoEst, total: IMBEL_PRODUTOS.length };
  }, []);

  function toggleCat(c) {
    const s = new Set(filterCat); if (s.has(c)) s.delete(c); else s.add(c); setFilterCat(s);
  }
  function toggleCollapse(c) {
    const s = new Set(collapsed); if (s.has(c)) s.delete(c); else s.add(c); setCollapsed(s);
  }
  function toggleBulk(id) {
    const s = new Set(bulkSelected); if (s.has(id)) s.delete(id); else s.add(id); setBulkSelected(s);
  }
  function getEdited(id, field) {
    return edits[id]?.[field] ?? IMBEL_PRODUTOS.find(p => p.id === id)[field];
  }
  function setField(id, field, value) {
    setEdits(prev => ({ ...prev, [id]: { ...prev[id], [field]: value } }));
  }
  function discardRow(id) {
    setEdits(prev => { const n = { ...prev }; delete n[id]; return n; });
    setEditingId(null);
  }

  const editCount = Object.keys(edits).length;

  return h('div', { className: 'imbel-subtab' }, [

    // 1. COMMAND BAR
    h('div', { className: 'imbel-cmdbar', key: 'cmd' }, [
      h('div', { className: 'search', key: 'sr' }, [
        h('svg', { width: 14, height: 14, viewBox: '0 0 24 24', fill: 'none', stroke: 'currentColor', strokeWidth: 2, key: 'i' },
          h('circle', { cx: 11, cy: 11, r: 8 }), h('line', { x1: 21, y1: 21, x2: 16.65, y2: 16.65 })
        ),
        h('input', { type: 'text', value: search, onChange: e => setSearch(e.target.value), placeholder: 'Buscar por ID, PN, nome, NCM…', key: 'in' }),
      ]),
      h('div', { className: 'divider', key: 'd1' }),
      h('div', { className: 'seg', key: 'st' }, [
        h('button', { className: filterStatus === 'ativo'   ? 'active' : '', onClick: () => setFilterStatus('ativo'),   key: 1 }, ['Ativos ',   h('span', { style: { fontFamily: 'IBM Plex Mono,monospace', fontSize: 10, marginLeft: 4, opacity: 0.6 }, key: 'n' }, stats.ativos)]),
        h('button', { className: filterStatus === 'inativo' ? 'active' : '', onClick: () => setFilterStatus('inativo'), key: 2 }, ['Inativos ', h('span', { style: { fontFamily: 'IBM Plex Mono,monospace', fontSize: 10, marginLeft: 4, opacity: 0.6 }, key: 'n' }, stats.inativos)]),
        h('button', { className: filterStatus === 'todos'   ? 'active' : '', onClick: () => setFilterStatus('todos'),   key: 3 }, 'Todos'),
      ]),
      bulkSelected.size > 0 && [
        h('div', { className: 'divider', key: 'd2' }),
        h('div', { className: 'edits-badge', key: 'bs', style: { background: '#dbeafe', color: '#1e3a5f', borderColor: '#93c5fd' } }, [
          h('span', { className: 'cnt', key: 'c', style: { background: '#1e3a5f', color: '#fff' } }, bulkSelected.size),
          ' selecionado', bulkSelected.size > 1 ? 's' : '',
        ]),
        h('button', { className: 'btn small ghost', key: 'bef', onClick: () => setBulkSelected(new Set()) }, 'Limpar'),
      ],
      editCount > 0 && !bulkSelected.size && [
        h('div', { className: 'divider', key: 'd3' }),
        h('div', { className: 'edits-badge', key: 'eb' }, [
          h('span', { className: 'cnt', key: 'c' }, editCount),
          ' alteração', editCount > 1 ? 'ões' : '',
        ]),
        h('button', { className: 'btn small ghost', onClick: () => setEdits({}), key: 'des' }, 'Descartar'),
      ],
      h('div', { className: 'actions', key: 'ac' }, [
        h('button', { className: 'btn', key: 1, onClick: () => setShowImport(true) }, 'Importar'),
        h('button', { className: 'btn', key: 2 }, 'Exportar'),
        editCount > 0 && h('button', { className: 'btn accent', key: 3 }, `Salvar (${editCount})`),
        h('button', { className: `btn ${editCount > 0 ? '' : 'accent'}`, key: 4 }, '+ Novo produto'),
      ])
    ]),

    // 2. TYPE BAR (categoria pills)
    h('div', { className: 'imbel-typebar', key: 'cb' }, [
      h('span', { className: 'imbel-typebar-lbl', key: 'l' }, 'Categoria:'),
      ...IMBEL_CATEGORIAS.map(c => {
        const cnt = IMBEL_PRODUTOS.filter(p => p.categoria === c).length;
        return h('button', {
          key: c,
          className: `imbel-pill ${filterCat.has(c) ? 'on' : ''}`,
          onClick: () => toggleCat(c),
        }, [c, h('span', { className: 'imbel-pill-cnt', key: 'cnt' }, cnt)]);
      }),
      filterCat.size > 0 && h('button', {
        className: 'imbel-pill', onClick: () => setFilterCat(new Set()),
        style: { marginLeft: 'auto', color: '#94a3b8' }, key: 'cl'
      }, 'Limpar')
    ]),

    // 3. KPI STRIP
    h('div', { className: 'imbel-kpis', key: 'kpi' }, [
      h('div', { className: 'kpi', key: 1 }, [
        h('div', { className: 'kpi-label', key: 'l' }, 'SKUs cadastrados'),
        h('div', { className: 'kpi-value', key: 'v' }, [stats.total, h('span', { className: 'kpi-unit', key: 'u' }, ' itens')]),
        h('div', { className: 'kpi-sub', key: 'd' }, [stats.ativos, ' ativos · ', stats.inativos, ' inativos']),
      ]),
      h('div', { className: 'kpi', key: 2 }, [
        h('div', { className: 'kpi-label', key: 'l' }, 'Estoque consolidado'),
        h('div', { className: 'kpi-value accent', key: 'v' }, [fmt(stats.sumEst, 0), h('span', { className: 'kpi-unit', key: 'u' }, ' un')]),
        h('div', { className: 'kpi-sub', key: 'd' }, ['galpão IMBEL · valor CI R$ ', fmt(stats.sumVal / 1000, 0), 'k']),
      ]),
      h('div', { className: 'kpi', key: 3 }, [
        h('div', { className: 'kpi-label', key: 'l' }, 'Categorias'),
        h('div', { className: 'kpi-value', key: 'v' }, IMBEL_CATEGORIAS.length),
        h('div', { className: 'kpi-sub', key: 'd' }, 'classificações'),
      ]),
      h('div', { className: 'kpi', key: 4 }, [
        h('div', { className: 'kpi-label', key: 'l' }, ['Em alerta', stats.baixoEst > 0 && h('span', { className: 'kpi-tag', key: 't', style: { background: '#fee2e2', color: '#dc2626' } }, '!')]),
        h('div', { className: `kpi-value ${stats.baixoEst > 0 ? 'neg' : ''}`, key: 'v' }, stats.baixoEst),
        h('div', { className: 'kpi-sub', key: 'd' }, 'abaixo do mínimo'),
      ]),
      h('div', { className: 'kpi', key: 5 }, [
        h('div', { className: 'kpi-label', key: 'l' }, 'Filtrados'),
        h('div', { className: 'kpi-value', key: 'v' }, [filtered.length, h('span', { className: 'kpi-unit', key: 'u' }, ' / ' + stats.total)]),
        h('div', { className: 'kpi-sub', key: 'd' }, grouped.length + ' categoria(s) visível(is)'),
      ]),
    ]),

    // 4. BODY
    h('div', { className: 'imbel-body', key: 'body' },
      h('table', { className: 'price-table imbel-cad-table' }, [
        h('thead', { key: 'h' }, h('tr', {}, [
          h('th', { key: 'cb', style: { minWidth: 36, width: 36 } }, h('input', { type: 'checkbox', onChange: e => setBulkSelected(e.target.checked ? new Set(filtered.map(p => p.id)) : new Set()) })),
          h('th', { key: 1, style: { minWidth: 70  } }, 'ID'),
          h('th', { key: 2, style: { minWidth: 110 } }, 'PN'),
          h('th', { key: 3, style: { minWidth: 220 } }, 'NOME'),
          h('th', { key: 4, style: { minWidth: 200 } }, 'FÁBRICA'),
          h('th', { key: 5, style: { minWidth: 90  } }, 'NCM'),
          h('th', { className: 'num', key: 6, style: { minWidth: 80 } }, 'PESO'),
          h('th', { className: 'num', key: 7, style: { minWidth: 100, background: 'linear-gradient(180deg,#0f1e31,#4d2818 50%)' } }, 'CI'),
          h('th', { className: 'num', key: 8, style: { minWidth: 100 } }, 'ESTOQUE'),
          h('th', { className: 'num', key: 9, style: { minWidth: 90  } }, 'MÍN'),
          h('th', { key: 10, style: { minWidth: 90  } }, 'STATUS'),
          h('th', { key: 11, style: { minWidth: 100 } }, 'ÚLT MOV'),
          h('th', { key: 12, style: { minWidth: 60  } }, ''),
        ])),
        h('tbody', { key: 'b' },
          grouped.flatMap(([cat, items]) => {
            const isCollapsed = collapsed.has(cat);
            const allSel  = items.length > 0 && items.every(p => bulkSelected.has(p.id));
            const rows = [];
            const sumQtd = items.reduce((s, p) => s + p.estoqueGalpao, 0);
            const sumCI  = items.reduce((s, p) => s + p.estoqueGalpao * p.ci, 0);
            rows.push(
              h('tr', { className: 'imbel-group-row', key: 'grp-' + cat, onClick: () => toggleCollapse(cat) },
                h('td', { colSpan: 13, key: 'c' },
                  h('div', { className: 'imbel-group-inner', key: 'r' }, [
                    h('input', { type: 'checkbox', checked: allSel, onChange: () => {}, onClick: e => e.stopPropagation(), key: 'cb' }),
                    h('span', { className: 'imbel-group-collapse', key: 'c' }, isCollapsed ? '▶' : '▼'),
                    h('span', { className: 'imbel-group-name', key: 'n' }, cat),
                    h('span', { className: 'imbel-group-cnt', key: 'q' }, items.length + ' SKU'),
                    h('span', { className: 'imbel-group-stat', key: 's' }, [h('span', { className: 'l', key: 1 }, 'Σ estoque'), h('b', { key: 2 }, fmt(sumQtd, 0))]),
                    h('span', { className: 'imbel-group-stat', key: 'v' }, [h('span', { className: 'l', key: 1 }, 'Σ valor CI'), h('b', { key: 2 }, 'R$ ' + fmt(sumCI / 1000, 0) + 'k')]),
                  ])
                )
              )
            );
            if (!isCollapsed) {
              items.forEach(p => {
                const editing  = editingId === p.id;
                const dirty    = !!edits[p.id];
                const selected = bulkSelected.has(p.id);
                rows.push(
                  h('tr', {
                    key: p.id,
                    className: `${selected ? 'selected' : ''} ${dirty ? 'dirty' : ''}`,
                    onDoubleClick: () => setEditingId(p.id),
                    style: dirty ? { background: 'rgba(229,161,40,0.06)' } : {},
                  }, [
                    h('td', { key: 'cb' }, h('input', { type: 'checkbox', checked: selected, onChange: () => toggleBulk(p.id) })),
                    h('td', { className: 'ncm', key: 1 }, p.id),
                    h('td', { className: 'pn', key: 2 }, editing
                      ? h('input', { type: 'text', value: getEdited(p.id, 'pn'), onChange: e => setField(p.id, 'pn', e.target.value), className: 'imbel-cad-input' })
                      : p.pn),
                    h('td', { className: 'name', key: 3 }, editing
                      ? h('input', { type: 'text', value: getEdited(p.id, 'nome'), onChange: e => setField(p.id, 'nome', e.target.value), className: 'imbel-cad-input' })
                      : (dirty && edits[p.id]?.nome ? edits[p.id].nome : p.nome)),
                    h('td', { className: 'fabrica', key: 4, title: p.fabrica }, editing
                      ? h('input', { type: 'text', value: getEdited(p.id, 'fabrica'), onChange: e => setField(p.id, 'fabrica', e.target.value), className: 'imbel-cad-input' })
                      : p.fabrica),
                    h('td', { className: 'pn', key: 5 }, p.ncm),
                    h('td', { className: 'num', key: 6 }, p.peso.toFixed(3) + ' kg'),
                    h('td', { className: 'num ci', key: 7 }, editing
                      ? h('input', { type: 'number', step: 0.01, value: getEdited(p.id, 'ci'), onChange: e => setField(p.id, 'ci', parseFloat(e.target.value) || 0), className: 'imbel-cad-input', style: { textAlign: 'right' } })
                      : fmt(p.ci)),
                    h('td', { className: 'num', key: 8, style: { fontWeight: 700, color: p.estoqueGalpao < p.estoqueMin ? '#dc2626' : '#0f1e31' } }, fmt(p.estoqueGalpao, 0)),
                    h('td', { className: 'num', key: 9 }, fmt(p.estoqueMin, 0)),
                    h('td', { key: 10 }, h('span', { className: 'imbel-status ' + p.status }, p.status)),
                    h('td', { className: 'pn', key: 11 }, p.ultMov.split('-').reverse().join('/')),
                    h('td', { key: 12, style: { textAlign: 'right' } },
                      editing
                        ? h('div', { style: { display: 'flex', gap: 4, justifyContent: 'flex-end' } }, [
                            h('button', { key: 1, onClick: () => setEditingId(null), style: { background: 'none', border: 'none', cursor: 'pointer', fontSize: 14, color: '#16a34a' } }, '✓'),
                            h('button', { key: 2, onClick: () => discardRow(p.id), style: { background: 'none', border: 'none', cursor: 'pointer', fontSize: 14, color: '#dc2626' } }, '✕'),
                          ])
                        : h('button', { onClick: () => setEditingId(p.id), title: 'Editar', style: { background: 'none', border: 'none', cursor: 'pointer', color: '#94a3b8' } }, '⋯')
                    ),
                  ])
                );
              });
            }
            return rows;
          })
        )
      ])
    ),

    // 5. FOOTBAR
    h('div', { className: 'imbel-footbar', key: 'foot' }, [
      h('div', { className: 'seg', key: 1 }, [h('span', { className: 'lbl' }, 'SKU'), h('span', { className: 'val' }, filtered.length + '/' + stats.total)]),
      h('div', { className: 'divider', key: 'd1' }),
      h('div', { className: 'seg', key: 2 }, [h('span', { className: 'lbl' }, 'CAT'), h('span', { className: 'val accent' }, grouped.length)]),
      h('div', { className: 'divider', key: 'd2' }),
      h('div', { className: 'seg', key: 3 }, [h('span', { className: 'lbl' }, 'SELECT'), h('span', { className: 'val' }, bulkSelected.size)]),
      h('div', { className: 'divider', key: 'd3' }),
      h('div', { className: 'seg', key: 4 }, [h('span', { className: 'lbl' }, 'EDIT'), h('span', { className: `val ${editCount > 0 ? 'accent' : ''}` }, editCount)]),
      h('div', { className: 'right', key: 'rt' },
        h('div', { className: 'keys' }, [
          h('span', { key: 1 }, [h('b', { key: 'b' }, 'Dbl-clk'), ' Editar']),
          h('span', { key: 2 }, [h('b', { key: 'b' }, 'N'), ' Novo']),
          h('span', { key: 3 }, [h('b', { key: 'b' }, '⌘S'), ' Salvar']),
        ])
      )
    ])
  ]);
}

window.ImbelCadastro = ImbelCadastro;
