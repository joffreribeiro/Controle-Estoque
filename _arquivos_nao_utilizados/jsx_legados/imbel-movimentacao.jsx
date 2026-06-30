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
        (m.id || '').toLowerCase().includes(q) ||
        (m.produtoId || '').toLowerCase().includes(q) ||
        (m.produtoNome || '').toLowerCase().includes(q) ||
        (m.destinatario || '').toLowerCase().includes(q) ||
        (m.cpfCnpj || '').toLowerCase().includes(q) ||
        (m.nf || '').toLowerCase().includes(q) ||
        (m.usuario || '').toLowerCase().includes(q) ||
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
      const pid = m.produtoId || (m.items && m.items[0]?.produtoId);
      const saldoApos = saldos[pid] || 0;
      const isEnt = typeof imbelTipoAumentaEstoque === 'function'
        ? imbelTipoAumentaEstoque(m.tipo)
        : (m.tipo || '').toUpperCase().includes('ENTRADA') || (m.tipo || '') === 'entrada';
      const delta = isEnt ? Number(m.quantidade) || 0 : -(Math.abs(Number(m.quantidade) || 0));
      if (pid) saldos[pid] = saldoApos - delta;
      return { ...m, saldoApos };
    });
  }, [filtered]);

  const stats = useMemoIM(() => {
    const byTipo = { entrada: 0, saida: 0, ajuste: 0, transferencia: 0 };
    let totalEntrada = 0, totalSaida = 0, totalFat = 0;
    filtered.forEach(m => {
      const tipoKey = (m.tipo || '').toLowerCase().includes('ajuste') ? 'ajuste'
        : (m.tipo || '').toLowerCase().includes('transfer') ? 'transferencia'
        : (typeof imbelTipoAumentaEstoque === 'function' ? imbelTipoAumentaEstoque(m.tipo) : (m.tipo||'').toUpperCase().includes('ENTRADA'))
          ? 'entrada' : 'saida';
      byTipo[tipoKey] = (byTipo[tipoKey] || 0) + 1;
      if (tipoKey === 'entrada') totalEntrada += Number(m.quantidade) || 0;
      if (tipoKey === 'saida')  { totalSaida += Math.abs(Number(m.quantidade) || 0); totalFat += Number(m.valor) || Number(m.valorTotal) || 0; }
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

  function fmtDateLong(m) {
    // suporta tanto m.timestamp (ISO) quanto m.data + m.hora (campos separados)
    if (m.data) {
      const d = new Date(m.data + 'T12:00:00');
      return {
        data: d.toLocaleDateString('pt-BR'),
        hora: m.hora || '',
      };
    }
    if (m.timestamp) {
      const d = new Date(m.timestamp);
      return {
        data: d.toLocaleDateString('pt-BR'),
        hora: d.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' }),
      };
    }
    return { data: '—', hora: '' };
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
          h('th', { key: 'cb', style: { width: 36, minWidth: 36 } },
            h('input', { type: 'checkbox', title: 'Selecionar todos' })
          ),
          h('th', { key: 1,  style: { minWidth: 80,  textAlign: 'left'  } }, 'ID'),
          h('th', { key: 2,  style: { minWidth: 180, textAlign: 'left'  } }, 'DESTINATÁRIO'),
          h('th', { key: 3,  style: { minWidth: 100 } }, 'DATA'),
          h('th', { key: 4,  style: { minWidth: 110 } }, 'TIPO'),
          h('th', { key: 5,  style: { minWidth: 180, textAlign: 'left'  } }, 'PRODUTO'),
          h('th', { className: 'num', key: 6, style: { minWidth: 70  } }, 'QTD'),
          h('th', { className: 'num', key: 7, style: { minWidth: 110 } }, 'VALOR'),
          h('th', { key: 8,  style: { minWidth: 50  }, title: 'Entregue'  }, 'ENTG.'),
          h('th', { key: 9,  style: { minWidth: 50  }, title: 'Pagamento' }, 'PGTO'),
          h('th', { key: 10, style: { minWidth: 40  }, title: 'FI'        }, 'FI'),
          h('th', { key: 11, style: { minWidth: 70  } }, 'AÇÕES'),
        ])),
        h('tbody', { key: 'b' },
          withBalance.map(m => {
            const { data, hora } = fmtDateLong(m);
            const editing = editingId === m.id;
            // normaliza tipo para lógica de cores
            const tipoUp   = (m.tipo || '').toUpperCase();
            const isEntrada = typeof imbelTipoAumentaEstoque === 'function'
              ? imbelTipoAumentaEstoque(m.tipo)
              : tipoUp.includes('ENTRADA');
            const isAjuste  = tipoUp.includes('AJUSTE');
            const isTransf  = tipoUp.includes('TRANSFER');
            const badgeBg   = isAjuste ? 'rgba(37,99,235,.10)' : isTransf ? 'rgba(30,58,95,.10)' : isEntrada ? 'rgba(22,163,74,.12)' : 'rgba(220,38,38,.10)';
            const badgeClr  = isAjuste ? '#1d4ed8' : isTransf ? '#1e3a5f' : isEntrada ? '#15803d' : '#dc2626';
            const badgePfx  = isAjuste ? '⬡ ' : isTransf ? '→ ' : isEntrada ? '+ ' : '↗ ';
            const qtdSign   = isEntrada ? '+' : '−';
            const qtdColor  = isEntrada ? '#15803d' : '#dc2626';
            // label do tipo: usa IMBEL_TIPOS se disponível, senão usa m.tipo direto
            const tipoLabel = (typeof IMBEL_TIPOS !== 'undefined' && IMBEL_TIPOS[m.tipo])
              ? IMBEL_TIPOS[m.tipo].label
              : (m.tipo || '—');

            // checkbox cell helper
            const chkStyle = { width: 16, height: 16, cursor: 'pointer', accentColor: isEntrada ? '#16a34a' : '#1e3a5f' };

            return h('tr', { key: m.id, className: editing ? 'editing' : '', style: { fontWeight: editing ? 600 : 400 } }, [
              // ☐ checkbox
              h('td', { key: 'cb', style: { textAlign: 'center', padding: '5px 8px' } },
                h('input', { type: 'checkbox', style: chkStyle })
              ),
              // ID
              h('td', { key: 1, style: { whiteSpace: 'nowrap', padding: '5px 8px' } },
                h('span', { style: { fontFamily: 'IBM Plex Mono,monospace', fontSize: 11, fontWeight: 700, color: '#1e3a5f', background: '#f0f4f9', padding: '1px 6px', borderRadius: 2, letterSpacing: '.04em' } }, m.id)
              ),
              // DESTINATÁRIO
              h('td', { key: 2, style: { fontWeight: 600, color: '#1e293b', padding: '5px 8px', maxWidth: 200, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' } }, [
                m.destinatario || h('span', { style: { color: '#94a3b8' }, key: 'd' }, '—'),
                m.cpfCnpj && h('div', { style: { fontFamily: 'IBM Plex Mono,monospace', fontSize: 10, color: '#94a3b8', fontWeight: 400, marginTop: 1 }, key: 'c' }, m.cpfCnpj),
              ]),
              // DATA
              h('td', { key: 3, style: { textAlign: 'center', whiteSpace: 'nowrap', fontFamily: 'IBM Plex Mono,monospace', fontSize: 11, color: '#374151', padding: '5px 8px' } }, [
                h('div', { key: 'd' }, data),
                hora && h('div', { key: 'h', style: { fontSize: 10, color: '#94a3b8' } }, hora),
              ]),
              // TIPO badge
              h('td', { key: 4, style: { textAlign: 'center', padding: '5px 8px' } },
                h('span', { style: { display: 'inline-flex', alignItems: 'center', gap: 3, background: badgeBg, color: badgeClr, fontFamily: 'var(--tv-font-display,inherit)', fontSize: 10, fontWeight: 700, letterSpacing: '.06em', padding: '2px 7px', borderRadius: 3, whiteSpace: 'nowrap', border: `1px solid ${badgeClr}33` } },
                  badgePfx + tipoLabel
                )
              ),
              // PRODUTO
              h('td', { key: 5, style: { fontWeight: 500, padding: '5px 8px', maxWidth: 220, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' } }, [
                h('div', { style: { fontWeight: 600, color: '#0f1e31' }, key: 'n' }, m.produtoNome || m.produtoId || '—'),
                m.produtoId && m.produtoNome && h('div', { style: { fontFamily: 'IBM Plex Mono,monospace', fontSize: 10, color: '#94a3b8' }, key: 'id' }, m.produtoId),
              ]),
              // QTD
              h('td', { key: 6, className: 'num', style: { fontFamily: 'IBM Plex Mono,monospace', fontWeight: 700, fontSize: 13, color: qtdColor, padding: '5px 8px' } },
                editing
                  ? h('input', { type: 'number', value: editData.quantidade, onChange: e => setEditData({ ...editData, quantidade: parseInt(e.target.value) || 0 }), style: { width: 60, padding: '3px 5px', textAlign: 'right', fontFamily: 'IBM Plex Mono,monospace', border: '1px solid #e5a128', borderRadius: 2 } })
                  : qtdSign + fmt(Math.abs(m.quantidade), 0)
              ),
              // VALOR
              h('td', { key: 7, className: 'num', style: { fontFamily: 'IBM Plex Mono,monospace', fontWeight: 700, fontSize: 12, color: '#15803d', padding: '5px 8px' } },
                m.valorTotal ? 'R$ ' + fmt(m.valorTotal, 0) : h('span', { style: { color: '#94a3b8', fontWeight: 400 } }, '—')
              ),
              // ENTG. checkbox
              h('td', { key: 8, style: { textAlign: 'center', padding: '5px 8px' } },
                h('input', { type: 'checkbox', checked: (m.entregue || '').toUpperCase() === 'SIM', readOnly: true, style: { width: 15, height: 15, cursor: 'pointer', accentColor: '#16a34a' } })
              ),
              // PGTO checkbox
              h('td', { key: 9, style: { textAlign: 'center', padding: '5px 8px' } },
                h('input', { type: 'checkbox', checked: (m.pagamento || '').toUpperCase() === 'SIM', readOnly: true, style: { width: 15, height: 15, cursor: 'pointer', accentColor: '#16a34a' } })
              ),
              // FI checkbox
              h('td', { key: 10, style: { textAlign: 'center', padding: '5px 8px' } },
                h('input', { type: 'checkbox', checked: (m.fi || '').toUpperCase() === 'SIM', readOnly: true, style: { width: 15, height: 15, cursor: 'pointer', accentColor: '#1e3a5f' } })
              ),
              // AÇÕES
              h('td', { key: 11, style: { textAlign: 'center', padding: '5px 8px' } },
                editing
                  ? h('div', { style: { display: 'flex', gap: 4, justifyContent: 'center' } }, [
                      h('button', { key: 1, onClick: cancelEdit, style: { background: 'none', border: '1px solid #16a34a', borderRadius: 3, padding: '2px 7px', cursor: 'pointer', color: '#16a34a', fontSize: 12 } }, '✓'),
                      h('button', { key: 2, onClick: cancelEdit, style: { background: 'none', border: '1px solid #dc2626', borderRadius: 3, padding: '2px 7px', cursor: 'pointer', color: '#dc2626', fontSize: 12 } }, '✕'),
                    ])
                  : h('div', { style: { display: 'flex', gap: 4, justifyContent: 'center' } }, [
                      h('button', { key: 1, onClick: () => startEdit(m), title: 'Editar', style: { background: 'none', border: '1px solid #e2e8f0', borderRadius: 3, padding: '2px 7px', cursor: 'pointer', color: '#64748b', fontSize: 12 } }, '✎'),
                      h('button', { key: 2, title: 'Excluir', style: { background: 'none', border: '1px solid #fca5a5', borderRadius: 3, padding: '2px 7px', cursor: 'pointer', color: '#dc2626', fontSize: 12 } }, '✕'),
                    ])
              ),
            ]);
          })
        )
      ]),
      editingId && h('div', { className: 'imbel-edit-bar', key: 'eb' }, [
        h('div', { className: 'imbel-edit-ttl', key: 1 }, ['Editando ', h('b', { key: 'b' }, editingId), ' · justificativa obrigatória']),
        h('input', { type: 'text', value: editReason, onChange: e => setEditReason(e.target.value), placeholder: 'Por que está editando este lançamento?', className: 'imbel-edit-input', key: 2 }),
        h('button', { key: 3, onClick: cancelEdit, style: { border: '1px solid #e2e8f0', borderRadius: 4, padding: '0 12px', height: 32, cursor: 'pointer', fontSize: 12, background: '#fff' } }, 'Cancelar'),
        h('button', { key: 4, onClick: cancelEdit, disabled: !editReason.trim(), style: { background: '#0f1e31', color: '#e5a128', border: 'none', borderRadius: 4, padding: '0 12px', height: 32, cursor: 'pointer', fontSize: 12, fontWeight: 600 } }, 'Salvar'),
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
