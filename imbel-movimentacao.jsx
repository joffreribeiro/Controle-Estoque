// IMBEL · Movimentação — extrato de entradas/saídas com filtros e edição inline
const { useState: useStateIM, useMemo: useMemoIM } = React;

const TIPO_MOV_META = {
  entrada:       { label: 'Entrada',       sign: '+', color: '#16a34a', bg: '#dcfce7', icon: 'Plus' },
  saida:         { label: 'Saída',         sign: '−', color: '#dc2626', bg: '#fee2e2', icon: 'Trend' },
  ajuste:        { label: 'Ajuste',        sign: '±', color: '#92400e', bg: '#fef3c7', icon: 'Settings' },
  transferencia: { label: 'Transferência', sign: '→', color: '#1e3a5f', bg: '#dbeafe', icon: 'Refresh' },
};

function ImbelMovimentacao() {
  const [search,       setSearch]       = useStateIM('');
  const [filterTipo,   setFilterTipo]   = useStateIM(new Set());
  const [filterProduto,setFilterProduto]= useStateIM(null);
  const [filterUser,   setFilterUser]   = useStateIM('todos');
  const [period,       setPeriod]       = useStateIM('30d');
  const [editingId,    setEditingId]    = useStateIM(null);
  const [editData,     setEditData]     = useStateIM(null);
  const [editReason,   setEditReason]   = useStateIM('');
  const [showImport,   setShowImport]   = useStateIM(false);
  const [novaForm,     setNovaForm]     = useStateIM(false);

  const users = useMemoIM(() => [...new Set(IMBEL_MOVIMENTACOES.map(m => m.usuario))], []);

  const filtered = useMemoIM(() => {
    let arr = IMBEL_MOVIMENTACOES;
    if (search.trim()) {
      const q = search.toLowerCase();
      arr = arr.filter(m =>
        m.id.toLowerCase().includes(q) ||
        m.produtoId.toLowerCase().includes(q) ||
        m.produtoNome.toLowerCase().includes(q) ||
        (m.nf || '').toLowerCase().includes(q) ||
        (m.cliente || '').toLowerCase().includes(q) ||
        m.usuario.toLowerCase().includes(q) ||
        (m.observacao || '').toLowerCase().includes(q)
      );
    }
    if (filterTipo.size) arr = arr.filter(m => filterTipo.has(m.tipo));
    if (filterProduto)   arr = arr.filter(m => m.produtoId === filterProduto);
    if (filterUser !== 'todos') arr = arr.filter(m => m.usuario === filterUser);
    return arr;
  }, [search, filterTipo, filterProduto, filterUser]);

  const withBalance = useMemoIM(() => {
    const saldos = {};
    IMBEL_PRODUTOS.forEach(p => { saldos[p.id] = p.estoqueGalpao; });
    return [...filtered].map(m => {
      const saldoApos = saldos[m.produtoId];
      const delta = m.tipo === 'entrada' ? m.quantidade : m.tipo === 'saida' ? -Math.abs(m.quantidade) : m.quantidade;
      saldos[m.produtoId] = saldoApos - delta;
      return { ...m, saldoApos };
    });
  }, [filtered]);

  const stats = useMemoIM(() => {
    const byTipo = { entrada: 0, saida: 0, ajuste: 0, transferencia: 0 };
    let totalEntrada = 0, totalSaida = 0, totalFat = 0;
    filtered.forEach(m => {
      byTipo[m.tipo] = (byTipo[m.tipo] || 0) + 1;
      if (m.tipo === 'entrada') totalEntrada += m.quantidade;
      if (m.tipo === 'saida')  { totalSaida += Math.abs(m.quantidade); totalFat += m.valorTotal; }
    });
    return { total: filtered.length, byTipo, totalEntrada, totalSaida, totalFat };
  }, [filtered]);

  function toggleTipo(t) {
    const s = new Set(filterTipo);
    if (s.has(t)) s.delete(t); else s.add(t);
    setFilterTipo(s);
  }
  function startEdit(m) { setEditingId(m.id); setEditData({ ...m }); setEditReason(''); }
  function cancelEdit() { setEditingId(null); setEditData(null); setEditReason(''); }

  function fmtDateLong(ts) {
    const d = new Date(ts);
    return {
      data: d.toLocaleDateString('pt-BR'),
      hora: d.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' }),
    };
  }

  return h('div', { className: 'imbel-subtab' }, [

    // 1. COMMAND BAR
    h('div', { className: 'imbel-cmdbar', key: 'cmd' }, [
      h('div', { className: 'search', key: 'sr' }, [
        h('svg', { width: 14, height: 14, viewBox: '0 0 24 24', fill: 'none', stroke: 'currentColor', strokeWidth: 2, key: 'i' },
          h('circle', { cx: 11, cy: 11, r: 8 }), h('line', { x1: 21, y1: 21, x2: 16.65, y2: 16.65 })
        ),
        h('input', { type: 'text', value: search, onChange: e => setSearch(e.target.value), placeholder: 'Buscar NF, cliente, produto…', key: 'in' }),
      ]),
      h('div', { className: 'divider', key: 'd1' }),
      h('span', { className: 'filter-label', key: 'ul' }, 'Usuário'),
      h('select', { value: filterUser, onChange: e => setFilterUser(e.target.value), key: 'su' }, [
        h('option', { value: 'todos', key: 'all' }, 'Todos'),
        ...users.map(u => h('option', { value: u, key: u }, u))
      ]),
      h('span', { className: 'filter-label', key: 'pl' }, 'Produto'),
      h('select', { value: filterProduto || '', onChange: e => setFilterProduto(e.target.value || null), key: 'sp', style: { minWidth: 160, fontFamily: 'IBM Plex Mono, monospace', fontSize: 11 } }, [
        h('option', { value: '', key: 'all' }, 'Todos os produtos'),
        ...IMBEL_PRODUTOS.map(p => h('option', { value: p.id, key: p.id }, p.id + ' · ' + p.nome))
      ]),
      h('div', { className: 'seg', key: 'pd' }, [
        h('button', { className: period === '7d'  ? 'active' : '', onClick: () => setPeriod('7d'),  key: 1 }, '7d'),
        h('button', { className: period === '30d' ? 'active' : '', onClick: () => setPeriod('30d'), key: 2 }, '30d'),
        h('button', { className: period === '90d' ? 'active' : '', onClick: () => setPeriod('90d'), key: 3 }, '90d'),
        h('button', { className: period === 'all' ? 'active' : '', onClick: () => setPeriod('all'), key: 4 }, 'Tudo'),
      ]),
      h('div', { className: 'actions', key: 'ac' }, [
        h('button', { className: 'btn', key: 1, onClick: () => setShowImport(true) }, 'Importar Excel'),
        h('button', { className: 'btn', key: 2 }, 'Exportar'),
        h('button', { className: 'btn accent', key: 3, onClick: () => setNovaForm(true) }, '+ Nova movimentação'),
      ])
    ]),

    // 2. TYPE BAR
    h('div', { className: 'imbel-typebar', key: 'tb' }, [
      h('span', { className: 'imbel-typebar-lbl', key: 'l' }, 'Tipo:'),
      ...Object.entries(TIPO_MOV_META).map(([k, meta]) => {
        const cnt = stats.byTipo[k] || 0;
        return h('button', {
          key: k,
          className: `imbel-pill ${filterTipo.has(k) ? 'on' : ''} ${cnt === 0 ? 'empty' : ''}`,
          style: filterTipo.has(k) ? { background: meta.bg, color: meta.color, borderColor: meta.color + '40' } : {},
          onClick: () => toggleTipo(k),
        }, [meta.label, h('span', { className: 'imbel-pill-cnt', key: 'c' }, cnt)]);
      }),
      filterTipo.size > 0 && h('button', {
        className: 'imbel-pill', onClick: () => setFilterTipo(new Set()),
        style: { marginLeft: 'auto', color: '#94a3b8' }, key: 'cl'
      }, 'Limpar')
    ]),

    // 3. KPI STRIP
    h('div', { className: 'imbel-kpis', key: 'kpi' }, [
      h('div', { className: 'kpi', key: 1 }, [
        h('div', { className: 'kpi-label', key: 'l' }, ['Lançamentos ', h('span', { className: 'kpi-tag', key: 't' }, period)]),
        h('div', { className: 'kpi-value', key: 'v' }, [stats.total, h('span', { className: 'kpi-unit', key: 'u' }, ' registros')]),
        h('div', { className: 'kpi-sub', key: 'd' }, ['de ', IMBEL_MOVIMENTACOES.length, ' totais']),
      ]),
      h('div', { className: 'kpi', key: 2 }, [
        h('div', { className: 'kpi-label', key: 'l' }, 'Entradas'),
        h('div', { className: 'kpi-value pos', key: 'v' }, ['+', fmt(stats.totalEntrada, 0), h('span', { className: 'kpi-unit', key: 'u' }, ' un')]),
        h('div', { className: 'kpi-sub', key: 'd' }, [stats.byTipo.entrada || 0, ' notas de entrada']),
      ]),
      h('div', { className: 'kpi', key: 3 }, [
        h('div', { className: 'kpi-label', key: 'l' }, 'Saídas'),
        h('div', { className: 'kpi-value neg', key: 'v' }, ['-', fmt(stats.totalSaida, 0), h('span', { className: 'kpi-unit', key: 'u' }, ' un')]),
        h('div', { className: 'kpi-sub', key: 'd' }, [stats.byTipo.saida || 0, ' notas de saída']),
      ]),
      h('div', { className: 'kpi', key: 4 }, [
        h('div', { className: 'kpi-label', key: 'l' }, 'Saldo líquido'),
        h('div', { className: `kpi-value ${(stats.totalEntrada - stats.totalSaida) >= 0 ? 'pos' : 'neg'}`, key: 'v' },
          [(stats.totalEntrada - stats.totalSaida) >= 0 ? '+' : '', fmt(stats.totalEntrada - stats.totalSaida, 0), h('span', { className: 'kpi-unit', key: 'u' }, ' un')]),
        h('div', { className: 'kpi-sub', key: 'd' }, 'entrada − saída'),
      ]),
      h('div', { className: 'kpi', key: 5 }, [
        h('div', { className: 'kpi-label', key: 'l' }, 'Faturamento bruto'),
        h('div', { className: 'kpi-value accent', key: 'v' }, ['R$ ', fmt(stats.totalFat / 1000, 0), h('span', { className: 'kpi-unit', key: 'u' }, ' mil')]),
        h('div', { className: 'kpi-sub', key: 'd' }, 'soma das saídas'),
      ]),
    ]),

    // 4. BODY
    h('div', { className: 'imbel-body', key: 'body' }, [
      h('table', { className: 'price-table imbel-mov-table', key: 't' }, [
        h('thead', { key: 'h' }, h('tr', {}, [
          h('th', { key: 1, style: { minWidth: 90  } }, 'ID'),
          h('th', { key: 2, style: { minWidth: 110 } }, 'DATA / HORA'),
          h('th', { key: 3, style: { minWidth: 110 } }, 'TIPO'),
          h('th', { key: 4, style: { minWidth: 180 } }, 'PRODUTO'),
          h('th', { className: 'num', key: 5, style: { minWidth: 80  } }, 'QTD'),
          h('th', { className: 'num', key: 6, style: { minWidth: 100 } }, 'SALDO APÓS'),
          h('th', { key: 7, style: { minWidth: 90  } }, 'NF'),
          h('th', { key: 8, style: { minWidth: 180 } }, 'CLIENTE'),
          h('th', { key: 9, style: { minWidth: 100 } }, 'USUÁRIO'),
          h('th', { className: 'num', key: 10, style: { minWidth: 110 } }, 'VALOR'),
          h('th', { key: 11, style: { minWidth: 180 } }, 'OBSERVAÇÃO'),
          h('th', { key: 12, style: { minWidth: 80  } }, ''),
        ])),
        h('tbody', { key: 'b' },
          withBalance.map(m => {
            const meta = TIPO_MOV_META[m.tipo];
            const { data, hora } = fmtDateLong(m.timestamp);
            const editing = editingId === m.id;
            return h('tr', { key: m.id, className: editing ? 'editing' : '' }, [
              h('td', { className: 'ncm', key: 1 }, m.id),
              h('td', { key: 2 }, [
                h('div', { style: { fontFamily: 'IBM Plex Mono,monospace', fontSize: 11, color: '#0f1e31', fontWeight: 600 }, key: 1 }, data),
                h('div', { style: { fontFamily: 'IBM Plex Mono,monospace', fontSize: 10, color: '#94a3b8' }, key: 2 }, hora),
              ]),
              h('td', { key: 3 }, h('span', {
                style: { color: meta.color, background: meta.bg, padding: '2px 8px', borderRadius: 2, fontSize: 10, fontWeight: 600, letterSpacing: 0.4, textTransform: 'uppercase', display: 'inline-flex', alignItems: 'center', gap: 4 }
              }, meta.label)),
              h('td', { className: 'name', key: 4 }, [
                h('div', { style: { fontWeight: 600, color: '#0f1e31' }, key: 1 }, m.produtoNome),
                h('div', { style: { fontFamily: 'IBM Plex Mono,monospace', fontSize: 10, color: '#94a3b8' }, key: 2 }, m.produtoId),
              ]),
              h('td', { className: 'num', key: 5, style: { fontWeight: 700, color: meta.color, fontSize: 13 } },
                editing
                  ? h('input', { type: 'number', value: editData.quantidade, onChange: e => setEditData({ ...editData, quantidade: parseInt(e.target.value) || 0 }), style: { width: 70, padding: '4px 6px', textAlign: 'right', fontFamily: 'IBM Plex Mono,monospace', border: '1px solid #e5a128', borderRadius: 2 } })
                  : (meta.sign + fmt(Math.abs(m.quantidade), 0))),
              h('td', { className: 'num', key: 6, style: { color: m.saldoApos < (IMBEL_PRODUTOS.find(p => p.id === m.produtoId)?.estoqueMin || 0) ? '#dc2626' : '#64748b' } }, fmt(m.saldoApos, 0)),
              h('td', { className: 'pn',     key: 7 }, m.nf || '—'),
              h('td', { className: 'fabrica',key: 8 }, m.cliente || h('span', { style: { color: '#94a3b8' } }, '—')),
              h('td', { key: 9, style: { fontSize: 11 } }, m.usuario),
              h('td', { className: 'num', key: 10, style: { color: '#16a34a', fontWeight: 600 } }, m.tipo === 'saida' ? 'R$ ' + fmt(m.valorTotal, 0) : '—'),
              h('td', { key: 11, style: { fontSize: 11, fontStyle: 'italic', color: '#64748b' } },
                editing
                  ? h('input', { type: 'text', value: editData.observacao, onChange: e => setEditData({ ...editData, observacao: e.target.value }), placeholder: 'Observação', style: { width: '100%', padding: '4px 6px', fontSize: 11, border: '1px solid #e5a128', borderRadius: 2 } })
                  : (m.observacao || '—')),
              h('td', { key: 12, style: { textAlign: 'right' } },
                editing
                  ? h('div', { style: { display: 'flex', gap: 4 } }, [
                      h('button', { className: 'btn small', key: 1, style: { color: '#16a34a' }, onClick: cancelEdit }, '✓'),
                      h('button', { className: 'btn small ghost', key: 2, onClick: cancelEdit }, '✕'),
                    ])
                  : h('button', { className: 'icon-btn', onClick: () => startEdit(m), title: 'Editar' }, '⋯')
              ),
            ]);
          })
        )
      ]),
      editingId && h('div', { className: 'imbel-edit-bar', key: 'eb' }, [
        h('div', { className: 'imbel-edit-ttl', key: 1 }, ['Editando ', h('b', { key: 'b' }, editingId), ' · justificativa obrigatória']),
        h('input', { type: 'text', value: editReason, onChange: e => setEditReason(e.target.value), placeholder: 'Por que está editando este lançamento?', className: 'imbel-edit-input', key: 2 }),
        h('button', { className: 'btn ghost', onClick: cancelEdit, key: 3, style: { border: '1px solid #e2e8f0', borderRadius: 4, padding: '0 12px', height: 32, cursor: 'pointer', fontSize: 12 } }, 'Cancelar'),
        h('button', { className: 'btn accent', onClick: cancelEdit, disabled: !editReason.trim(), key: 4, style: { background: '#0f1e31', color: '#e5a128', border: 'none', borderRadius: 4, padding: '0 12px', height: 32, cursor: 'pointer', fontSize: 12, fontWeight: 600 } }, 'Salvar'),
      ]),
    ]),

    // 5. FOOTBAR
    h('div', { className: 'imbel-footbar', key: 'foot' }, [
      h('div', { className: 'seg', key: 1 }, [h('span', { className: 'lbl' }, 'MOV'), h('span', { className: 'val' }, `${filtered.length}/${IMBEL_MOVIMENTACOES.length}`)]),
      h('div', { className: 'divider', key: 'd1' }),
      h('div', { className: 'seg', key: 2 }, [h('span', { className: 'lbl' }, 'ENTRADAS'), h('span', { className: 'val', style: { color: '#4ade80' } }, '+' + fmt(stats.totalEntrada, 0))]),
      h('div', { className: 'divider', key: 'd2' }),
      h('div', { className: 'seg', key: 3 }, [h('span', { className: 'lbl' }, 'SAÍDAS'), h('span', { className: 'val', style: { color: '#f87171' } }, '-' + fmt(stats.totalSaida, 0))]),
      h('div', { className: 'divider', key: 'd3' }),
      h('div', { className: 'seg', key: 4 }, [h('span', { className: 'lbl' }, 'FAT'), h('span', { className: 'val accent' }, 'R$ ' + fmt(stats.totalFat / 1000, 0) + 'k')]),
      h('div', { className: 'right', key: 'rt' },
        h('div', { className: 'keys' }, [
          h('span', { key: 1 }, [h('b', { key: 'b' }, 'N'), ' Novo']),
          h('span', { key: 2 }, [h('b', { key: 'b' }, 'I'), ' Importar']),
          h('span', { key: 3 }, [h('b', { key: 'b' }, 'E'), ' Editar linha']),
        ])
      )
    ])
  ]);
}

window.ImbelMovimentacao = ImbelMovimentacao;
