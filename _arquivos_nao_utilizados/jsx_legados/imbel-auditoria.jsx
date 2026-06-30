// IMBEL · Auditoria — timeline de mudanças
const { useState: useStateIA, useMemo: useMemoIA } = React;

const IMBEL_AUDIT_META = {
  entrada:  { label: 'Entrada',        color: '#16a34a', bg: '#dcfce7' },
  saida:    { label: 'Saída',          color: '#dc2626', bg: '#fee2e2' },
  ajuste:   { label: 'Ajuste estoque', color: '#92400e', bg: '#fef3c7' },
  edicao:   { label: 'Edição',         color: '#1e3a5f', bg: '#dbeafe' },
  cadastro: { label: 'Cadastro',       color: '#0891b2', bg: '#cffafe' },
  preco:    { label: 'Preço',          color: '#b8651f', bg: '#fdebd8' },
  periodo:  { label: 'Período',        color: '#7c3aed', bg: '#ede9fe' },
};

function fmtTsIA(ts) {
  const d = new Date(ts);
  return {
    data: d.toLocaleDateString('pt-BR'),
    hora: d.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' }),
    rel: (function() {
      const diff = Date.now() - d.getTime();
      const min  = diff / 60000;
      if (min < 60)   return Math.floor(min) + ' min atrás';
      if (min < 1440) return Math.floor(min / 60) + ' h atrás';
      return Math.floor(min / 1440) + ' dias atrás';
    })(),
    full: d.toLocaleString('pt-BR'),
  };
}

function ImbelAuditoria() {
  const [search,      setSearch]      = useStateIA('');
  const [filterTipo,  setFilterTipo]  = useStateIA(new Set());
  const [filterUser,  setFilterUser]  = useStateIA('todos');
  const [period,      setPeriod]      = useStateIA('30d');
  const [selectedId,  setSelectedId]  = useStateIA(IMBEL_AUDIT[0]?.id || null);

  const users = useMemoIA(() => [...new Set(IMBEL_AUDIT.map(a => a.usuario))], []);

  const filtered = useMemoIA(() => {
    let arr = IMBEL_AUDIT;
    if (search.trim()) {
      const q = search.toLowerCase();
      arr = arr.filter(a =>
        a.target.toLowerCase().includes(q) ||
        a.usuario.toLowerCase().includes(q) ||
        a.id.toLowerCase().includes(q) ||
        String(a.before).toLowerCase().includes(q) ||
        String(a.after).toLowerCase().includes(q)
      );
    }
    if (filterTipo.size) arr = arr.filter(a => filterTipo.has(a.tipo));
    if (filterUser !== 'todos') arr = arr.filter(a => a.usuario === filterUser);
    return arr;
  }, [search, filterTipo, filterUser]);

  const selected = filtered.find(a => a.id === selectedId) || filtered[0];

  const grouped = useMemoIA(() => {
    const map = new Map();
    filtered.forEach(a => {
      const key = new Date(a.timestamp).toISOString().slice(0, 10);
      if (!map.has(key)) map.set(key, []);
      map.get(key).push(a);
    });
    return [...map.entries()].sort((a, b) => b[0].localeCompare(a[0])).map(([key, entries]) => ({
      key, date: new Date(key), entries,
    }));
  }, [filtered]);

  const stats = useMemoIA(() => {
    const byTipo = {}, byUser = {};
    filtered.forEach(a => {
      byTipo[a.tipo]    = (byTipo[a.tipo]    || 0) + 1;
      byUser[a.usuario] = (byUser[a.usuario] || 0) + 1;
    });
    const topTipo = Object.keys(byTipo).sort((a, b) => byTipo[b] - byTipo[a])[0];
    return { total: filtered.length, byTipo, byUser, uniqueUsers: Object.keys(byUser).length, uniqueDays: grouped.length, topTipo };
  }, [filtered, grouped]);

  function toggleTipo(t) {
    const s = new Set(filterTipo); if (s.has(t)) s.delete(t); else s.add(t); setFilterTipo(s);
  }

  return h('div', { className: 'imbel-subtab' }, [

    // 1. COMMAND BAR
    h('div', { className: 'imbel-cmdbar', key: 'cmd' }, [
      h('div', { className: 'search', key: 'sr' }, [
        h('svg', { width: 14, height: 14, viewBox: '0 0 24 24', fill: 'none', stroke: 'currentColor', strokeWidth: 2, key: 'i' },
          h('circle', { cx: 11, cy: 11, r: 8 }), h('line', { x1: 21, y1: 21, x2: 16.65, y2: 16.65 })
        ),
        h('input', { type: 'text', value: search, onChange: e => setSearch(e.target.value), placeholder: 'Buscar em alvo, usuário, valor…', key: 'in' }),
      ]),
      h('div', { className: 'divider', key: 'd1' }),
      h('span', { className: 'filter-label', key: 'ul' }, 'Usuário'),
      h('select', { value: filterUser, onChange: e => setFilterUser(e.target.value), key: 'su' }, [
        h('option', { value: 'todos', key: 'all' }, 'Todos'),
        ...users.map(u => h('option', { value: u, key: u }, u))
      ]),
      h('div', { className: 'seg', key: 'pd' }, [
        h('button', { className: period === '7d'  ? 'active' : '', onClick: () => setPeriod('7d'),  key: 1 }, '7d'),
        h('button', { className: period === '30d' ? 'active' : '', onClick: () => setPeriod('30d'), key: 2 }, '30d'),
        h('button', { className: period === '90d' ? 'active' : '', onClick: () => setPeriod('90d'), key: 3 }, '90d'),
        h('button', { className: period === 'all' ? 'active' : '', onClick: () => setPeriod('all'), key: 4 }, 'Tudo'),
      ]),
      h('div', { className: 'actions', key: 'ac' }, [
        h('button', { className: 'btn', key: 1 }, 'Exportar CSV'),
        h('button', { className: 'btn icon', key: 2, title: 'Imprimir' }, '⎙'),
      ])
    ]),

    // 2. TYPE BAR (7 tipos de evento)
    h('div', { className: 'imbel-typebar', key: 'tb' }, [
      h('span', { className: 'imbel-typebar-lbl', key: 'l' }, 'Filtrar por tipo:'),
      ...Object.entries(IMBEL_AUDIT_META).map(([k, meta]) => {
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
        h('div', { className: 'kpi-label', key: 'l' }, ['Eventos ', h('span', { className: 'kpi-tag', key: 't' }, period)]),
        h('div', { className: 'kpi-value', key: 'v' }, [stats.total, h('span', { className: 'kpi-unit', key: 'u' }, ' registros')]),
        h('div', { className: 'kpi-sub', key: 'd' }, ['ao longo de ', stats.uniqueDays, ' dia(s)']),
      ]),
      h('div', { className: 'kpi', key: 2 }, [
        h('div', { className: 'kpi-label', key: 'l' }, 'Usuários ativos'),
        h('div', { className: 'kpi-value', key: 'v' }, stats.uniqueUsers),
        h('div', { className: 'kpi-sub', key: 'd' }, Object.keys(stats.byUser).slice(0, 3).join(' · ')),
      ]),
      h('div', { className: 'kpi', key: 3 }, [
        h('div', { className: 'kpi-label', key: 'l' }, 'Tipo mais frequente'),
        h('div', { className: 'kpi-value', key: 'v', style: { fontSize: 14 } }, IMBEL_AUDIT_META[stats.topTipo]?.label || '—'),
        h('div', { className: 'kpi-sub', key: 'd' }, [(stats.byTipo[stats.topTipo] || 0), ' eventos']),
      ]),
      h('div', { className: 'kpi', key: 4 }, [
        h('div', { className: 'kpi-label', key: 'l' }, 'Último evento'),
        h('div', { className: 'kpi-value', key: 'v', style: { fontSize: 14 } }, filtered[0] ? fmtTsIA(filtered[0].timestamp).rel : '—'),
        h('div', { className: 'kpi-sub', key: 'd' }, filtered[0] ? filtered[0].usuario : ''),
      ]),
      h('div', { className: 'kpi', key: 5 }, [
        h('div', { className: 'kpi-label', key: 'l' }, 'Selecionado'),
        h('div', { className: 'kpi-value accent', key: 'v', style: { fontSize: 14 } }, selected ? selected.id : '—'),
        h('div', { className: 'kpi-sub', key: 'd' }, selected ? (IMBEL_AUDIT_META[selected.tipo]?.label || selected.tipo) : ''),
      ]),
    ]),

    // 4. BODY (timeline + detalhe)
    h('div', { className: 'imbel-audit-main', key: 'body' }, [

      // Timeline
      h('div', { className: 'imbel-timeline', key: 'L' },
        grouped.map(g => h('div', { className: 'imbel-day', key: g.key }, [
          h('div', { className: 'imbel-day-head', key: 'h' }, [
            h('div', { className: 'imbel-day-date', key: 1 }, [
              h('span', { className: 'imbel-day-d', key: 'd' }, String(g.date.getDate()).padStart(2, '0')),
              h('div', { className: 'imbel-day-m', key: 'm' }, [
                g.date.toLocaleDateString('pt-BR', { month: 'short' }).toUpperCase().replace('.', ''),
                h('span', { className: 'imbel-day-y', key: 'y' }, g.date.getFullYear()),
              ])
            ]),
            h('div', { className: 'imbel-day-meta', key: 2 }, [
              g.date.toLocaleDateString('pt-BR', { weekday: 'long' }),
              h('span', { className: 'imbel-day-cnt', key: 'c' }, g.entries.length + ' evento' + (g.entries.length === 1 ? '' : 's')),
            ])
          ]),
          h('div', { className: 'imbel-entries', key: 'e' },
            g.entries.map(a => {
              const meta = IMBEL_AUDIT_META[a.tipo] || IMBEL_AUDIT_META.edicao;
              const ts = fmtTsIA(a.timestamp);
              return h('div', {
                key: a.id,
                className: `imbel-entry ${a.id === selectedId ? 'active' : ''}`,
                onClick: () => setSelectedId(a.id),
                style: { borderLeftColor: meta.color },
              }, [
                h('div', { className: 'imbel-entry-time', key: 1 }, ts.hora),
                h('div', { className: 'imbel-entry-body', key: 2 }, [
                  h('div', { className: 'imbel-entry-row1', key: 'r1' }, [
                    h('span', { className: 'imbel-entry-tipo', style: { color: meta.color, background: meta.bg }, key: 't' }, meta.label),
                    h('span', { className: 'imbel-entry-user', key: 'u' }, a.usuario),
                  ]),
                  h('div', { className: 'imbel-entry-target', key: 'tg' }, a.target),
                  h('div', { className: 'imbel-entry-diff', key: 'df' }, [
                    h('span', { className: 'before', key: 'b' }, a.before),
                    h('span', { style: { color: '#94a3b8' }, key: 'arr' }, '→'),
                    h('span', { className: 'after', key: 'a' }, a.after),
                    a.impact !== '—' && h('span', { className: 'impact ' + (a.impact.startsWith('+') ? 'pos' : 'neg'), key: 'i' }, a.impact),
                  ])
                ])
              ]);
            })
          )
        ]))
      ),

      // Painel de detalhe
      selected && h('div', { className: 'imbel-detail', key: 'D' }, [
        h('div', { className: 'imbel-detail-head', key: 'h' }, [
          h('div', { style: { display: 'flex', alignItems: 'center', gap: 8 }, key: 1 }, [
            (function() {
              const meta = IMBEL_AUDIT_META[selected.tipo] || IMBEL_AUDIT_META.edicao;
              return h('div', {
                style: { background: meta.bg, color: meta.color, padding: '3px 10px', borderRadius: 3, fontSize: 10, fontWeight: 600, letterSpacing: 0.5, textTransform: 'uppercase' },
              }, meta.label);
            })(),
            h('span', { style: { fontFamily: 'IBM Plex Mono,monospace', fontSize: 11, color: '#94a3b8' }, key: 'i' }, selected.id),
          ]),
          h('div', { className: 'imbel-detail-target', key: 2 }, selected.target),
          h('div', { className: 'imbel-detail-meta', key: 3 }, [
            h('span', { key: 1 }, selected.usuario),
            h('span', { key: 2 }, fmtTsIA(selected.timestamp).full),
            h('span', { key: 3, style: { color: '#e5a128' } }, '⚡ Módulo IMBEL'),
          ])
        ]),
        h('div', { className: 'imbel-detail-body', key: 'b' }, [
          h('div', { key: 1 }, [
            h('div', { className: 'imbel-detail-section-ttl', key: 't' }, 'Mudança'),
            h('div', { className: 'imbel-change-card', key: 'c' }, [
              h('div', { className: 'imbel-rcc-row', key: 1 }, [h('div', { className: 'imbel-rcc-l' }, 'Antes'), h('div', { className: 'imbel-rcc-v before' }, selected.before)]),
              h('div', { className: 'imbel-rcc-arrow', key: 2 }, '↓'),
              h('div', { className: 'imbel-rcc-row', key: 3 }, [h('div', { className: 'imbel-rcc-l' }, 'Depois'), h('div', { className: 'imbel-rcc-v after' }, selected.after)]),
              selected.impact !== '—' && h('div', {
                className: 'imbel-rcc-impact',
                style: { color: selected.impact.startsWith('+') ? '#16a34a' : '#dc2626' },
                key: 4
              }, ['⚡ Impacto: ', h('b', { key: 'b' }, selected.impact)])
            ])
          ]),
          h('div', { key: 2 }, [
            h('div', { className: 'imbel-detail-section-ttl', key: 't' }, 'Detalhes técnicos'),
            h('div', { className: 'imbel-tech', key: 'c' }, [
              h('div', { key: 1 }, [h('span', { className: 'l' }, 'Ação'),    h('span', { className: 'v' }, selected.acao)]),
              h('div', { key: 2 }, [h('span', { className: 'l' }, 'Entidade'),h('span', { className: 'v' }, selected.entidade)]),
              h('div', { key: 3 }, [h('span', { className: 'l' }, 'Origem'),  h('span', { className: 'v' }, 'web client')]),
              h('div', { key: 4 }, [h('span', { className: 'l' }, 'Hash'),    h('span', { className: 'v', style: { fontSize: 10 } }, 'sha256:…' + selected.id.toLowerCase())]),
            ])
          ]),
          h('div', { key: 3 }, [
            h('div', { className: 'imbel-detail-section-ttl', key: 't' }, 'Ações'),
            h('div', { style: { display: 'flex', flexDirection: 'column', gap: 6 }, key: 'c' }, [
              h('button', { key: 1, style: { width: '100%', height: 30, border: '1px solid #e2e8f0', borderRadius: 4, background: '#fff', cursor: 'pointer', fontSize: 12, textAlign: 'left', padding: '0 10px' } }, '↺ Reverter para "antes"'),
              h('button', { key: 2, style: { width: '100%', height: 30, border: '1px solid #e2e8f0', borderRadius: 4, background: '#fff', cursor: 'pointer', fontSize: 12, textAlign: 'left', padding: '0 10px' } }, '👁 Ver estado completo'),
              h('button', { key: 3, style: { width: '100%', height: 30, border: '1px solid #e2e8f0', borderRadius: 4, background: '#fff', cursor: 'pointer', fontSize: 12, textAlign: 'left', padding: '0 10px' } }, '↓ Baixar log assinado'),
            ])
          ])
        ])
      ])
    ]),

    // 5. FOOTBAR
    h('div', { className: 'imbel-footbar', key: 'foot' }, [
      h('div', { className: 'seg', key: 1 }, [h('span', { className: 'lbl' }, 'EVENTOS'), h('span', { className: 'val' }, `${filtered.length}/${IMBEL_AUDIT.length}`)]),
      h('div', { className: 'divider', key: 'd1' }),
      h('div', { className: 'seg', key: 2 }, [h('span', { className: 'lbl' }, 'DIAS'), h('span', { className: 'val accent' }, stats.uniqueDays)]),
      h('div', { className: 'divider', key: 'd2' }),
      h('div', { className: 'seg', key: 3 }, [h('span', { className: 'lbl' }, 'USERS'), h('span', { className: 'val' }, stats.uniqueUsers)]),
      h('div', { className: 'divider', key: 'd3' }),
      h('div', { className: 'seg', key: 4 }, [h('span', { className: 'lbl' }, 'SEL'), h('span', { className: 'val accent' }, selected ? selected.id : '—')]),
      h('div', { className: 'right', key: 'rt' },
        h('div', { className: 'keys' }, [
          h('span', { key: 1 }, [h('b', { key: 'b' }, 'R'), ' Reverter']),
          h('span', { key: 2 }, [h('b', { key: 'b' }, '↑↓'), ' Navegar']),
          h('span', { key: 3 }, [h('b', { key: 'b' }, '⌥[/]'), ' Aba']),
        ])
      )
    ])
  ]);
}

window.ImbelAuditoria = ImbelAuditoria;
