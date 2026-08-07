/**
 * icms-regras.js - CRUD da Tabela de Regras de ICMS (NCM × Estado × Tipo × Categoria)
 *
 * Extraído de app2.js como prova de conceito de divisão em módulos: bloco
 * autocontido (nenhuma função aqui é chamada de fora dele, e nada fora dele
 * chama as internas tipo _ordenarTodosNoFim/_normalizarCabecalho) que só lê e
 * escreve o array global `tabelaICMS`.
 *
 * Script clássico (sem type="module", sem bundler) — mesmo padrão de todos os
 * outros arquivos do projeto. Depende de globais definidos em app2.js e
 * carregados antes deste arquivo: `tabelaICMS` (array, populado no boot pelo
 * carregamento de dados), `ESTADOS_BR`, `CATEGORIAS_PRODUTO`, `_escapeHtml`,
 * `_escapeJsString`, `salvarDados`, `mostrarNotificacao`, e a biblioteca
 * externa `XLSX` (import/export de planilha).
 */

function _ordenarTodosNoFim(a, b) {
    if (a === b) return 0;
    if (a === 'Todos') return 1;
    if (b === 'Todos') return -1;
    return String(a).localeCompare(String(b), 'pt-BR');
}

function renderizarTabelaICMS() {
    const tbody = document.getElementById('tabelaICMSBody');
    if (!tbody) return;

    const regrasOrdenadas = [...tabelaICMS].sort((a, b) => {
        const e = _ordenarTodosNoFim(a.estado, b.estado);
        if (e !== 0) return e;
        const t = _ordenarTodosNoFim(a.tipoPessoa, b.tipoPessoa);
        if (t !== 0) return t;
        return _ordenarTodosNoFim(a.categoriaProduto, b.categoriaProduto);
    });

    tbody.innerHTML = regrasOrdenadas.map(rule => `
        <tr id="icms_row_${_escapeHtml(rule.id)}">
            <td>${_escapeHtml(rule.ncm || '—')}</td>
            <td>${_escapeHtml(rule.estado)}</td>
            <td>${_escapeHtml(rule.tipoPessoa === 'PJ' ? 'Lojista' : rule.tipoPessoa === 'PF' ? 'Cons. Final' : rule.tipoPessoa)}</td>
            <td>${_escapeHtml(rule.categoriaProduto)}</td>
            <td>${Number(rule.aliquota || 0).toFixed(2)}%</td>
            <td>
                <button onclick="editarRegraICMS('${_escapeJsString(rule.id)}')" style="border:none;background:none;cursor:pointer">✏️</button>
                <button onclick="excluirRegraICMS('${_escapeJsString(rule.id)}')" style="border:none;background:none;cursor:pointer">🗑️</button>
            </td>
        </tr>
    `).join('');
}

function adicionarRegraICMS() {
    const tbody = document.getElementById('tabelaICMSBody');
    if (!tbody) return;
    if (document.getElementById('novaRegraRow')) return;

    const opcoesEstado = ['Todos', ...ESTADOS_BR].map(e => `<option value="${e}">${e}</option>`).join('');
    const opcoesCategoria = ['Todos', ...CATEGORIAS_PRODUTO].map(c => `<option value="${c}">${c}</option>`).join('');
        const opcoesNCM = ['','9301.90.00','9305.91.00','9302.00.00','9305.10.00','8211.10.00','8201.40.00']
            .map(n => n === '' ? `<option value="">Todos</option>` : `<option value="${n}">${n}</option>`).join('');

    const html = `
            <tr id="novaRegraRow">
                <td>
                    <select id="nr_ncm">${opcoesNCM}</select>
                </td>
                <td>
                    <select id="nr_estado">${opcoesEstado}</select>
                </td>
                <td>
                    <select id="nr_tipoPessoa">
                        <option value="Todos">Todos</option>
                        <option value="PJ">Lojista</option>
                        <option value="PF">Cons. Final</option>
                    </select>
                </td>
                <td>
                    <select id="nr_categoria">${opcoesCategoria}</select>
                </td>
                <td>
                    <input type="number" id="nr_aliquota" step="0.01" min="0" placeholder="12" style="width:70px">
                </td>
                <td>
                    <button onclick="confirmarNovaRegraICMS()" style="color:#16a34a;font-size:1.1rem;background:none;border:none;cursor:pointer">✓</button>
                    <button onclick="cancelarNovaRegraICMS()" style="color:#dc2626;font-size:1.1rem;background:none;border:none;cursor:pointer">✗</button>
                </td>
            </tr>`;

    tbody.insertAdjacentHTML('afterbegin', html);
}

function confirmarNovaRegraICMS() {
    const ncm = (document.getElementById('nr_ncm')?.value || '').trim() || 'Todos';
    const estado = document.getElementById('nr_estado')?.value || 'Todos';
    const tipoPessoa = document.getElementById('nr_tipoPessoa')?.value || 'Todos';
    const categoria = document.getElementById('nr_categoria')?.value || 'Todos';
    const aliquota = parseFloat(document.getElementById('nr_aliquota')?.value);

    if (!Number.isFinite(aliquota) || aliquota < 0) {
        alert('Informe uma alíquota válida (>= 0).');
        return;
    }

    tabelaICMS.push({
        id: Date.now().toString(),
        ncm,
        estado,
        tipoPessoa,
        categoriaProduto: categoria,
        aliquota
    });

    renderizarTabelaICMS();
    salvarDados();
}

function cancelarNovaRegraICMS() {
    const row = document.getElementById('novaRegraRow');
    if (row) row.remove();
}

function editarRegraICMS(id) {
    const rule = tabelaICMS.find(r => String(r.id) === String(id));
    if (!rule) return;
    const row = document.getElementById(`icms_row_${id}`);
    if (!row) return;

    const opcoesNCM = ['','9301.90.00','9305.91.00','9302.00.00','9305.10.00','8211.10.00','8201.40.00']
        .map(n => n === '' ? `<option value="">Todos</option>` : `<option value="${n}" ${rule.ncm === n ? 'selected' : ''}>${n}</option>`).join('');
    const opcoesEstado = ['Todos', ...ESTADOS_BR]
        .map(e => `<option value="${e}" ${e === rule.estado ? 'selected' : ''}>${e}</option>`)
        .join('');
    const opcoesCategoria = ['Todos', ...CATEGORIAS_PRODUTO]
        .map(c => `<option value="${c}" ${c === rule.categoriaProduto ? 'selected' : ''}>${c}</option>`)
        .join('');

    row.innerHTML = `
        <td><select id="er_ncm_${id}">${opcoesNCM}</select></td>
        <td><select id="er_estado_${id}">${opcoesEstado}</select></td>
        <td>
            <select id="er_tipoPessoa_${id}">
                <option value="Todos" ${rule.tipoPessoa === 'Todos' ? 'selected' : ''}>Todos</option>
                <option value="PJ" ${rule.tipoPessoa === 'PJ' ? 'selected' : ''}>Lojista</option>
                <option value="PF" ${rule.tipoPessoa === 'PF' ? 'selected' : ''}>Cons. Final</option>
            </select>
        </td>
        <td><select id="er_categoria_${id}">${opcoesCategoria}</select></td>
        <td><input type="number" id="er_aliquota_${id}" step="0.01" min="0" value="${Number(rule.aliquota || 0)}" style="width:70px"></td>
        <td>
            <button onclick="atualizarRegraICMS('${_escapeJsString(id)}')" style="color:#16a34a;font-size:1.1rem;background:none;border:none;cursor:pointer">✓</button>
            <button onclick="renderizarTabelaICMS()" style="color:#dc2626;font-size:1.1rem;background:none;border:none;cursor:pointer">✗</button>
        </td>
    `;
}

function atualizarRegraICMS(id) {
    const ncm = (document.getElementById(`er_ncm_${id}`)?.value || '').trim() || 'Todos';
    const estado = document.getElementById(`er_estado_${id}`)?.value || 'Todos';
    const tipoPessoa = document.getElementById(`er_tipoPessoa_${id}`)?.value || 'Todos';
    const categoria = document.getElementById(`er_categoria_${id}`)?.value || 'Todos';
    const aliquota = parseFloat(document.getElementById(`er_aliquota_${id}`)?.value);

    if (!Number.isFinite(aliquota) || aliquota < 0) {
        alert('Informe uma alíquota válida (>= 0).');
        return;
    }

    const idx = tabelaICMS.findIndex(r => String(r.id) === String(id));
    if (idx === -1) return;

    tabelaICMS[idx] = {
        ...tabelaICMS[idx],
        ncm,
        estado,
        tipoPessoa,
        categoriaProduto: categoria,
        aliquota
    };

    renderizarTabelaICMS();
    salvarDados();
}

function excluirRegraICMS(id) {
    if (confirm('Excluir esta regra de ICMS?')) {
        tabelaICMS = tabelaICMS.filter(r => String(r.id) !== String(id));
        renderizarTabelaICMS();
        salvarDados();
    }
}

function exportarTabelaICMS() {
    if (typeof XLSX === 'undefined') {
        mostrarNotificacao('Biblioteca XLSX não encontrada.', 'error');
        return;
    }
    const dados = tabelaICMS.map(r => ({
        'NCM': r.ncm || '',
        'Estado': r.estado,
        'Tipo Pessoa': r.tipoPessoa,
        'Categoria Produto': r.categoriaProduto,
        'Aliquota ICMS (%)': r.aliquota
    }));
    const ws = XLSX.utils.json_to_sheet(dados);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Tabela ICMS');
    XLSX.writeFile(wb, 'tabela_icms.xlsx');
}

function importarTabelaICMS() {
    document.getElementById('inputImportarICMS').click();
}

function _normalizarCabecalho(txt) {
    return String(txt || '')
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .toLowerCase()
        .replace(/[^a-z0-9]/g, '');
}

function importarTabelaICMSArquivo(event) {
    const file = event.target.files?.[0];
    if (!file) return;
    if (typeof XLSX === 'undefined') {
        mostrarNotificacao('Biblioteca XLSX não encontrada.', 'error');
        event.target.value = '';
        return;
    }

    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const sheet = workbook.Sheets[workbook.SheetNames[0]];
            const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });

            let importadas = 0;
            let atualizadas = 0;

            rows.forEach((row, index) => {
                const normalized = {};
                Object.keys(row).forEach(k => {
                    normalized[_normalizarCabecalho(k)] = row[k];
                });

                const ncm = String(normalized.ncm || '').trim() || '';
                const estado = String(normalized.estado || 'Todos').trim() || 'Todos';
                const tipoPessoa = String(normalized.tipopessoa || 'Todos').trim() || 'Todos';
                const categoria = String(normalized.categoriaproduto || 'Todos').trim() || 'Todos';
                const aliquota = parseFloat(normalized.aliquotaicms ?? normalized.aliquotaicmspercent ?? normalized.aliquota ?? '');

                if (!Number.isFinite(aliquota) || aliquota < 0) return;

                const idxExistente = tabelaICMS.findIndex(r =>
                    String(r.estado).toUpperCase() === estado.toUpperCase()
                    && String(r.tipoPessoa).toUpperCase() === tipoPessoa.toUpperCase()
                    && String(r.categoriaProduto).toUpperCase() === categoria.toUpperCase()
                    && (String(r.ncm || '').toUpperCase() === String(ncm || '').toUpperCase())
                );

                if (idxExistente >= 0) {
                    tabelaICMS[idxExistente].aliquota = aliquota;
                    atualizadas++;
                } else {
                    tabelaICMS.push({
                        id: `${Date.now()}_${index}`,
                        ncm: ncm || 'Todos',
                        estado,
                        tipoPessoa,
                        categoriaProduto: categoria,
                        aliquota
                    });
                    importadas++;
                }
            });

            renderizarTabelaICMS();
            salvarDados();
            alert(`${importadas} regras importadas, ${atualizadas} atualizadas.`);
        } catch (err) {
            console.error('Erro ao importar tabela ICMS:', err);
            mostrarNotificacao('Falha ao importar tabela de ICMS.', 'error');
        } finally {
            event.target.value = '';
        }
    };

    reader.readAsArrayBuffer(file);
}

function abrirModalICMS() {
    renderizarTabelaICMS();
    document.getElementById('modalICMS').style.display = 'block';
}
