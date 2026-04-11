// ========================================
// SISTEMA DE CONTROLE DE ESTOQUE
// Material Bélico - v2.0
// ========================================

// Estrutura de dados principal
if (location.hostname !== 'localhost') {
    console.log = () => {};
    console.debug = () => {};
}
// Flag para indicar que há dados locais alterados que ainda não foram sincronizados com o cloud
let _dadosAlterados = false;
// Marca que houve sincronização recente com o cloud (evita warning ao fechar)
let _cloudSyncedRecently = false;
// Timer para resetar o estado de sincronização recente
let __cloudSyncResetTimer = null;

window.addEventListener('beforeunload', e => {
    try {
        if (_dadosAlterados && !_cloudSyncedRecently) {
            e.preventDefault();
            e.returnValue = 'Há dados não sincronizados com o cloud. Sair mesmo assim?';
        }
    } catch (err) {}
});
let estoque = {
    produtos: [],
    representantes: ['KOLTE', 'ISA', 'LC', 'ADES', 'FL', 'IMBEL'],
    registroVendas: [],
    registroDistribuicao: [],
    registroDevolucoes: [],
    registroEntradas: [],
    controleEnvio: {},
    auditoriaVendas: [],
    fechamentosComissoes: [],
    clientes: [],
    propostas: [],
    precificacao: {},
    precificacaoConfig: null
};

let tabelaAliquotas = {};
/*
Shape:
{
    "CARABINA IA2 5,56": {
        pis: 1.65,
        cofins: 7.60,
        ipi: 0,
        icmsBase: 12
    }
}

// ========================================
// GRÁFICO: COMISSÕES POR REPRESENTANTE
// ========================================

function renderizarGraficoComissoes() {
    try {
        if (typeof Chart === 'undefined') return;
        const canvas = document.getElementById('chartComissoesRep');
        if (!canvas) return;

        const periodo = document.getElementById('filtroComissoesGraficoPeriodo')?.value || 'todos';
        const agora = new Date();
        const vendasFiltradas = obterVendasDashboardFiltradas() || [];

        const reps = ['KOLTE', 'ISA', 'LC', 'ADES', 'FL', 'IMBEL'];
        const comissoes = {};
        reps.forEach(r => comissoes[r] = 0);
        let totalComissoes = 0;

        vendasFiltradas.forEach(venda => {
            const d = new Date(venda.data || 0);
            if (isNaN(d.getTime())) return;
            if (periodo === 'mes' && !(d.getMonth() === agora.getMonth() && d.getFullYear() === agora.getFullYear())) return;
            if (periodo === 'trimestre' && !(Math.floor(d.getMonth()/3) === Math.floor(agora.getMonth()/3) && d.getFullYear() === agora.getFullYear())) return;
            if (periodo === 'ano' && !(d.getFullYear() === agora.getFullYear())) return;

            const rep = (venda.representante || '').toUpperCase();
            const itens = obterItensVendaNormalizados(venda) || [];
            itens.forEach(it => {
                const produtoKey = it.produtoNome || it.produtoId || it.produto;
                // tentar achar margem/comissão na precificação do cliente
                let margem = null;
                try {
                    const precs = (precificacoesCliente || []).filter(p => p && p.clienteId === venda.clienteId);
                    if (precs.length) {
                        const found = (precs[0].itens || []).find(vi => (vi.produtoNome || vi.produtoId) == produtoKey || String(vi.produtoId) === String(produtoKey));
                        if (found) margem = Number(found.margem ?? found.margemPercent ?? found.comissao ?? 0);
                    }
                } catch (e) {}

                const valor = Number(it.valorTotal || it.valor || 0) || 0;
                let comissao = 0;
                if (margem && !isNaN(margem) && margem !== 0) {
                    comissao = valor * (Number(margem) / 100);
                } else {
                    comissao = valor * 0.05; // fallback 5%
                }

                if (!comissoes[rep]) comissoes[rep] = 0;
                comissoes[rep] += comissao;
                totalComissoes += comissao;
            });
        });

        const labels = reps;
        const dataValues = labels.map(l => comissoes[l] || 0);

        // tabela lateral
        const tabelaDiv = document.getElementById('tabelaComissoesGrafico');
        if (tabelaDiv) {
            tabelaDiv.innerHTML = '';
            const tbl = document.createElement('table');
            tbl.style.width = '100%';
            const tb = document.createElement('tbody');
            labels.forEach((lab, i) => {
                const tr = document.createElement('tr');
                tr.innerHTML = `<td style="padding:6px 0">${lab}</td><td style="text-align:right;padding:6px 0">${formatarMoedaValor(dataValues[i])}</td>`;
                tb.appendChild(tr);
            });
            tbl.appendChild(tb);
            tabelaDiv.appendChild(tbl);
        }

        const totalEl = document.getElementById('totalComissoesGrafico');
        if (totalEl) totalEl.textContent = 'Total Comissões: ' + formatarMoedaValor(totalComissoes);

        // chart
        if (canvas) {
            if (_chartComissoesRep) try { _chartComissoesRep.destroy(); } catch (e) {}
            const palette = ['#79c0ff', '#7ee787', '#58a6ff', '#ffa657', '#d2a8ff', '#ff7b72'];
            _chartComissoesRep = new Chart(canvas, {
                type: 'bar',
                data: {
                    labels,
                    datasets: [{ label: 'Comissões (R$)', data: dataValues, backgroundColor: palette, borderRadius: 6 }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {
                        legend: { display: false },
                        tooltip: { callbacks: { label: function(ctx) { return formatarMoedaValor(ctx.raw || ctx.parsed?.y || 0); } } }
                    },
                    scales: { y: { beginAtZero: true, ticks: { callback: function(v){ return 'R$ ' + Number(v).toLocaleString('pt-BR'); } } } }
                }
            });
        }
    } catch (e) {
        console.warn('renderizarGraficoComissoes erro', e);
    }
}
*/

function verificarExpiracaoPrecificacoes() {
    try {
        const agora = new Date();
        let expiradas = 0;
        (precificacoesCliente || []).forEach(p => {
            if (p && p.status === 'ativa' && p.dataExpiracao) {
                try {
                    if (new Date(p.dataExpiracao) < agora) {
                        p.status = 'expirada';
                        expiradas++;
                    }
                } catch (e) { /* ignore parse errors */ }
            }
        });
        try { estoque.precificacoesCliente = precificacoesCliente; } catch (e) {}
        return expiradas;
    } catch (e) {
        console.warn('verificarExpiracaoPrecificacoes erro', e);
        return 0;
    }
}

let tabelaICMS = [];
/*
Shape per rule:
{
    id: "rule_001",
    estado: "SP",
    tipoPessoa: "PJ",
    categoriaProduto: "Arma Curta",
    aliquota: 12
}
*/

let precificacao = {};
/*
Shape per product:
{
    "CARABINA IA2 5,56": {
        ci: 0,
        taxa: 1,
        roi: 1,
        comissao: 5,
        precoFinalManual: null
    }
}
*/

// Precificações salvas por cliente (temporário/permanente salvo em estado)
let precificacoesCliente = [];
let ultimaPrecificacaoCalculada = null;
let ultimaVersaoSalva = null;
let exibindoPrecifSalva = false;
let precifSalvaCarregada = null;
// RETID + Benefícios fiscais (estado por sessão de precificação)
let retidPorProduto = {};
let beneficioFiscalAtivo = false;
let beneficiosPorProduto = {};

const CATEGORIAS_PRODUTO = [
        'Arma Curta', 'Arma Longa', 'Acessório', 'Faca', 'Munição', 'Outro'
];

let categoriaPorProduto = {};

// ----------------------------------------
// Debug helpers: overlay de erro e painel de diagnóstico
// (ajuda a capturar erros que antes eram suprimidos)
;(function(){
    function showRuntimeErrorOverlay(msg) {
        try {
            const id = 'runtime-error-overlay';
            let el = document.getElementById(id);
            if (!el) {
                el = document.createElement('div');
                el.id = id;
                el.style.position = 'fixed';
                el.style.right = '12px';
                el.style.bottom = '12px';
                el.style.zIndex = 2000;
                el.style.maxWidth = '520px';
                el.style.background = 'rgba(220,38,38,0.95)';
                el.style.color = '#fff';
                el.style.padding = '12px';
                el.style.borderRadius = '8px';
                el.style.fontSize = '0.9rem';
                el.style.boxShadow = '0 8px 30px rgba(0,0,0,0.3)';
                el.style.whiteSpace = 'pre-wrap';
                el.style.fontFamily = 'Inter, system-ui, -apple-system, "Segoe UI", Roboto, "Helvetica Neue", Arial';
                el.style.cursor = 'pointer';
                el.title = 'Clique para esconder este painel de erro';
                el.addEventListener('click', () => { el.style.display = 'none'; });
                (document.body || document.documentElement).appendChild(el);
            }
            const text = (typeof msg === 'string') ? msg : (msg && msg.stack) ? msg.stack : JSON.stringify(msg, null, 2);
            el.textContent = 'Erro em tempo de execução detectado:\n' + text;
        } catch (e) { try { console.error('Falha ao exibir overlay de erro', e); } catch(_){} }
    }

    function createDebugPanel(){
        try {
            const id = 'debug-panel-mini';
            if (document.getElementById(id)) return;
            const p = document.createElement('div');
            p.id = id;
            p.style.position = 'fixed';
            p.style.left = '12px';
            p.style.bottom = '12px';
            p.style.zIndex = 2000;
            p.style.background = 'rgba(15,23,42,0.95)';
            p.style.color = '#cbd5e1';
            p.style.padding = '10px';
            p.style.borderRadius = '8px';
            p.style.fontSize = '0.85rem';
            p.style.boxShadow = '0 6px 24px rgba(0,0,0,0.2)';
            p.style.fontFamily = 'Inter, system-ui, -apple-system, "Segoe UI", Roboto, Arial';
            p.style.maxWidth = '360px';
            p.style.whiteSpace = 'nowrap';
            p.innerHTML = '<strong style="color:#fff;display:block;margin-bottom:6px">Debug</strong>' +
                '<div id="debug-panel-content">carregando...</div>' +
                '<div style="margin-top:8px;text-align:right"><button id="debug-copy-json" style="background:#1f2937;color:#fff;border:none;padding:6px 8px;border-radius:6px;cursor:pointer">Copiar JSON</button></div>';
            (document.body || document.documentElement).appendChild(p);
            document.getElementById('debug-copy-json')?.addEventListener('click', () => {
                try {
                    const raw = localStorage.getItem('estoqueArmasV2') || '{}';
                    navigator.clipboard?.writeText(raw);
                    mostrarNotificacao && mostrarNotificacao('JSON copiado para área de transferência', 'info');
                } catch(e) { showRuntimeErrorOverlay(e); }
            });
        } catch(e){}
    }

    function updateDebugPanel(){
        try {
            createDebugPanel();
            const el = document.getElementById('debug-panel-content');
            if (!el) return;
            const raw = localStorage.getItem('estoqueArmasV2');
            let parsed = null;
            try { parsed = raw ? JSON.parse(raw) : null; } catch(e) { parsed = null; }
            const cfgRaw = localStorage.getItem('configContrato');
            let cfg = null;
            try { cfg = cfgRaw ? JSON.parse(cfgRaw) : null; } catch(e){ cfg = cfgRaw; }
            const lines = [];
            lines.push('produtos: ' + (Array.isArray(parsed?.produtos) ? parsed.produtos.length : '-'));
            lines.push('registroVendas: ' + (Array.isArray(parsed?.registroVendas) ? parsed.registroVendas.length : '-'));
            lines.push('registroDistribuicao: ' + (Array.isArray(parsed?.registroDistribuicao) ? parsed.registroDistribuicao.length : '-'));
            lines.push('clientes: ' + (Array.isArray(parsed?.clientes) ? parsed.clients?.length || parsed?.clientes.length : '-'));
            lines.push('configContrato: ' + (cfg ? (cfg.proximo || '-') + '/' + (cfg.ano || '-') : 'ausente'));
            lines.push('abaAtiva: ' + (window.abaAtiva || '-'));
            el.textContent = lines.join('\n');
        } catch(e) { showRuntimeErrorOverlay(e); }
    }

    window.__showRuntimeErrorOverlay = showRuntimeErrorOverlay;
    window.__updateDebugPanel = updateDebugPanel;

    window.addEventListener('error', function(ev){ try { showRuntimeErrorOverlay(ev.error || ev.message || ev); console.error('Unhandled error', ev.error || ev.message || ev); } catch(e){} });
    window.addEventListener('unhandledrejection', function(ev){ try { showRuntimeErrorOverlay(ev.reason || ev); console.error('Unhandled rejection', ev.reason || ev); } catch(e){} });
    document.addEventListener('DOMContentLoaded', function(){ setTimeout(updateDebugPanel, 400); });
})();
// ===== NCM / Predefinições fiscais =====
const NCM_PRODUTOS = {
    "9301.90.00": "FUZIL DE ALTA PRECISÃO IMBEL 308 AGLC",
    "9305.91.00": "PEÇAS",
    "9302.00.00": "PISTOLA",
    "9305.10.00": "CARREGADOR",
    "8211.10.00": "FACA",
    "8201.40.00": "MACHADINHA"
};

const NCM_POR_CATEGORIA = {
    "FUZIL":      "9301.90.00",
    "PEÇAS":      "9305.91.00",
    "PISTOLA":    "9302.00.00",
    "REVÓLVER":   "9302.00.00",
    "CARABINA":   "9302.00.00",
    "CARREGADOR": "9305.10.00",
    "FACA":       "8211.10.00",
    "MACHADINHA": "8201.40.00"
};

const IMPOSTOS_FEDERAIS_POR_NCM = {
    "9301.90.00": { pis: 1.65, cofins: 7.60, ipi: 55.00 },
    "9305.91.00": { pis: 1.65, cofins: 7.60, ipi:  0.00 },
    "9302.00.00": { pis: 1.65, cofins: 7.60, ipi: 55.00 },
    "9305.10.00": { pis: 1.65, cofins: 7.60, ipi: 29.25 },
    "8211.10.00": { pis: 1.65, cofins: 7.60, ipi:  7.80 },
    "8201.40.00": { pis: 1.65, cofins: 7.60, ipi:  0.00 }
};

const ICMS_PJ_POR_NCM = {
    "9301.90.00": {
        AC:7, AL:7, AP:7, AM:7, BA:7, CE:7, DF:7, ES:7, GO:7,
        MA:7, MS:7, PA:7, PB:7, PR:12, PE:7, PI:7, RJ:12,
        RN:7, RS:12, RO:7, RR:7, SC:12, SP:12, SE:7, TO:7, MG:12
    },
    "9305.91.00": {
        AC:7, AL:7, AP:7, AM:7, BA:7, CE:7, DF:7, ES:7, GO:7,
        MA:7, MS:7, PA:7, PB:7, PR:12, PE:7, PI:7, RJ:12,
        RN:7, RS:12, RO:7, RR:7, SC:12, SP:12, SE:7, TO:7, MG:12
    },
    "9302.00.00": {
        AC:7, AL:7, AP:7, AM:7, BA:7, CE:7, DF:7, ES:7, GO:7,
        MA:7, MS:7, PA:7, PB:7, PR:12, PE:7, PI:7, RJ:12,
        RN:7, RS:12, RO:7, RR:7, SC:12, SP:12, SE:7, TO:7, MG:12
    },
    "9305.10.00": {
        AC:7, AL:7, AP:7, AM:7, BA:7, CE:7, DF:7, ES:7, GO:7,
        MA:7, MS:7, PA:7, PB:7, PR:12, PE:7, PI:7, RJ:12,
        RN:7, RS:12, RO:7, RR:7, SC:12, SP:12, SE:7, TO:7, MG:12
    },
    "8211.10.00": {
        AC:7, AL:7, AP:7, AM:7, BA:7, CE:7, DF:7, ES:7, GO:7,
        MA:7, MS:7, PA:7, PB:7, PR:12, PE:7, PI:7, RJ:12,
        RN:7, RS:12, RO:7, RR:7, SC:12, SP:12, SE:7, TO:7, MG:12
    },
    "8201.40.00": {
        AC:7, AL:7, AP:7, AM:7, BA:7, CE:7, DF:7, ES:7, GO:7,
        MA:7, MS:7, PA:7, PB:7, PR:12, PE:7, PI:7, RJ:12,
        RN:7, RS:12, RO:7, RR:7, SC:12, SP:12, SE:7, TO:7, MG:12
    }
};

const ICMS_PF_POR_NCM = {
    "9301.90.00": {
        AC:25, AL:29, AP:29, AM:25, BA:38, CE:28, DF:25, ES:25, GO:25,
        MA:30.5, MT:35, MS:25, PA:30, PB:25, PR:25, PE:27, PI:33,
        RJ:37, RN:25, RS:25, RO:25, RR:25, SC:25, SP:25, SE:28,
        TO:27, MG:25
    },
    "9305.91.00": {
        AC:19, AL:29, AP:29, AM:20, BA:20.5, CE:20, DF:20, ES:25, GO:25,
        MA:23, MT:35, MS:25, PA:30, PB:20, PR:25, PE:27, PI:22.5,
        RJ:37, RN:20, RS:25, RO:25, RR:25, SC:25, SP:25, SE:28,
        TO:20, MG:18
    },
    "9302.00.00": {
        AC:25, AL:29, AP:29, AM:25, BA:38, CE:28, DF:25, ES:25, GO:25,
        MA:30.5, MT:35, MS:25, PA:30, PB:25, PR:25, PE:27, PI:33,
        RJ:37, RN:25, RS:25, RO:25, RR:25, SC:25, SP:25, SE:28,
        TO:27, MG:25
    },
    "9305.10.00": {
        AC:19, AL:29, AP:29, AM:20, BA:20.5, CE:20, DF:20, ES:25, GO:25,
        MA:23, MT:35, MS:25, PA:30, PB:20, PR:25, PE:27, PI:22.5,
        RJ:37, RN:20, RS:25, RO:25, RR:25, SC:25, SP:25, SE:28,
        TO:20, MG:18
    },
    "8211.10.00": {
        AC:19, AL:19, AP:18, AM:20, BA:20.5, CE:20, DF:20, ES:17, GO:19,
        MA:23, MT:17, MS:17, PA:19, PB:20, PR:19.5, PE:20.5, PI:22.5,
        RJ:20, RN:20, RS:17, RO:19.5, RR:20, SC:17, SP:18, SE:19,
        TO:20, MG:18
    },
    "8201.40.00": {
        AC:19, AL:19, AP:18, AM:20, BA:20.5, CE:20, DF:20, ES:17, GO:19,
        MA:23, MT:17, MS:17, PA:19, PB:20, PR:19.5, PE:20.5, PI:22.5,
        RJ:20, RN:20, RS:17, RO:19.5, RR:20, SC:17, SP:18, SE:19,
        TO:20, MG:18
    }
};


let abaAtiva = 'estoque';

// Configuração de alertas (persistida separadamente)
let configAlertas = { limite: 5, ativo: true };

// Helper: cálculo de estoque/vendas/distribuições
function obterTotalVendasProduto(produto) {
    if (!produto) return 0;
    // Preferir registros detalhados quando existirem (fonte de verdade)
    if (Array.isArray(estoque.registroVendas) && estoque.registroVendas.length > 0) {
        return estoque.registroVendas.reduce((sum, v) => {
            if (Array.isArray(v.items) && v.items.length) {
                return sum + v.items.reduce((s, it) => {
                    if (Number(it.produtoId) === Number(produto.id) || (it.produto === produto.nome) || (it.produtoNome === produto.nome)) return s + (Number(it.quantidade) || 0);
                    return s;
                }, 0);
            }
            if (Number(v.produtoId) === Number(produto.id) || v.produtoNome === produto.nome) return sum + (Number(v.quantidade) || 0);
            return sum;
        }, 0);
    }
    // Fallback para agregado em produto.vendas quando não há registros
    if (!produto.vendas) return 0;
    return Object.keys(produto.vendas).reduce((s, k) => s + (Number(produto.vendas[k]) || 0), 0);
}

function obterDistribuicaoTotalExcluindoImbel(produto) {
    if (!produto) return 0;
    // Preferir registroDistribuicao como fonte de verdade quando disponível
    if (Array.isArray(estoque.registroDistribuicao) && estoque.registroDistribuicao.length > 0) {
        return estoque.registroDistribuicao.reduce((s, d) => {
            try {
                if (Number(d.produtoId) === Number(produto.id) || (d.produtoNome && d.produtoNome === produto.nome)) {
                    const rep = (d.representante || '').toString().toUpperCase();
                    if (rep === 'IMBEL') return s;
                    return s + (Number(d.quantidade) || 0);
                }
            } catch (e) {}
            return s;
        }, 0);
    }
    // Fallback para o objeto produto.distribuicao (legado)
    if (!produto.distribuicao) return 0;
    return Object.keys(produto.distribuicao).reduce((s, k) => {
        if ((k || '').toUpperCase() === 'IMBEL') return s;
        return s + (Number(produto.distribuicao[k]) || 0);
    }, 0);
}

function calcularSaldoConsolidado(produto) {
    if (!produto) return 0;
    const estoqueConc = Number(produto.estoqueConsolidado || 0);
    const vendas = obterTotalVendasProduto(produto);
    return estoqueConc - vendas;
}

function calcularImbelDisponivel(produto) {
    if (!produto) return 0;
    const saldoConsol = calcularSaldoConsolidado(produto);
    const distribExcl = obterDistribuicaoTotalExcluindoImbel(produto);
    return saldoConsol - distribExcl;
}

// Calcula estoque disponível na IMBEL baseado exclusivamente em registros
function calcularEstoqueIMBEL(nomeProduto) {
    if (!nomeProduto) return 0;
    const produtos = estoque.produtos || [];
    const produto = produtos.find(p => {
        try {
            if (!p || !p.nome) return false;
            return p.nome.toString().toUpperCase() === nomeProduto.toString().toUpperCase();
        } catch (e) { return false; }
    });
    if (!produto) return 0;

    const estoqueTotal = Number(produto.estoqueConsolidado || produto.estoqueTotal || produto.estoque || 0);

    const totalDistribuido = (estoque.registroDistribuicao || []).reduce((sum, d) => {
        try {
            const match = (Number(d.produtoId) === Number(produto.id)) || ((d.produtoNome || '').toString().toUpperCase() === produto.nome.toString().toUpperCase());
            if (!match) return sum;
            const rep = (d.representante || '').toString().toUpperCase();
            if (rep === 'IMBEL') return sum; // não descontar distribuições para IMBEL
            return sum + (Number(d.quantidade) || 0);
        } catch (e) { return sum; }
    }, 0);

    const totalDevolvido = (estoque.registroDevolucoes || []).reduce((sum, d) => {
        try {
            const match = (Number(d.produtoId) === Number(produto.id)) || ((d.produtoNome || '').toString().toUpperCase() === produto.nome.toString().toUpperCase());
            if (!match) return sum;
            const destino = (d.destino || '').toString().toUpperCase();
            if (destino === 'IMBEL') return sum + (Number(d.quantidade) || 0);
        } catch (e) {}
        return sum;
    }, 0);

    return estoqueTotal - totalDistribuido + totalDevolvido;
}

function obterMetricasImbelProduto(produto) {
    if (!produto) {
        return { estoqueTotal: 0, imbelDisp: 0, imbelVenda: 0, imbelSaldo: 0 };
    }

    const produtoId = produto.id;
    const estoqueTotal = Number(
        (produto.estoqueConsolidado ?? produto.estoqueTotal ?? produto.estoque ?? produto.qtdTotal ?? produto.estoqueInicial) || 0
    );

    const distPorRep = {};
    (estoque.registroDistribuicao || []).forEach(d => {
        try {
            if (Number(d.produtoId) === Number(produtoId) || (d.produtoNome && d.produtoNome === produto.nome)) {
                const r = (d.representante || '').toString().toUpperCase();
                distPorRep[r] = (distPorRep[r] || 0) + (Number(d.quantidade) || 0);
            }
        } catch (e) {}
    });
    const totalDistribuido = Object.values(distPorRep).reduce((s, v) => s + v, 0);

    let totalDevolvidoParaImbel = 0;
    (estoque.registroDevolucoes || []).forEach(d => {
        try {
            if (Number(d.produtoId) === Number(produtoId) || (d.produtoNome && d.produtoNome === produto.nome)) {
                const destino = (d.destino || '').toString().toUpperCase();
                if (destino === 'IMBEL') totalDevolvidoParaImbel += (Number(d.quantidade) || 0);
            }
        } catch (e) {}
    });

    const vendasPorRep = {};
    let agregadoTemValores = false;
    if (produto.vendas && typeof produto.vendas === 'object') {
        Object.keys(produto.vendas).forEach(k => {
            const val = Number(produto.vendas[k]) || 0;
            if (val !== 0) agregadoTemValores = true;
            vendasPorRep[(k || '').toString().toUpperCase()] = val;
        });
    }
    if (!agregadoTemValores) {
        (estoque.registroVendas || []).forEach(v => {
            try {
                const rep = (v.representante || '').toString().toUpperCase();
                if (Array.isArray(v.items) && v.items.length) {
                    v.items.forEach(it => {
                        if (Number(it.produtoId) === Number(produtoId) || (it.produto === produto.nome) || (it.produtoNome === produto.nome)) {
                            vendasPorRep[rep] = (vendasPorRep[rep] || 0) + (Number(it.quantidade) || 0);
                        }
                    });
                } else if (Number(v.produtoId) === Number(produtoId) || (v.produtoNome === produto.nome)) {
                    vendasPorRep[rep] = (vendasPorRep[rep] || 0) + (Number(v.quantidade) || 0);
                }
            } catch (e) {}
        });
    }

    const totalVendasProduto = Object.values(vendasPorRep).reduce((s, v) => s + v, 0);
    const consolidadoDisp = estoqueTotal;
    const consolidadoVenda = totalVendasProduto;
    const consolidadoSaldo = consolidadoDisp - consolidadoVenda;

    const imbelDisp = estoqueTotal - totalDistribuido + totalDevolvidoParaImbel;
    const imbelVenda = Number(vendasPorRep['IMBEL'] || 0);
    const imbelSaldo = imbelDisp - imbelVenda;

    return {
        estoqueTotal,
        imbelDisp,
        imbelVenda,
        imbelSaldo,
        consolidadoDisp,
        consolidadoVenda,
        consolidadoSaldo
    };
}

// Reconstrói o objeto `produto.distribuicao` a partir dos registros de distribuição e devolução
function reconstruirDistribuicaoAPartirDeRegistros() {
    if (!Array.isArray(estoque.produtos)) return;
    // Inicializar chaves
    estoque.produtos.forEach(p => {
        try {
            if (!p.distribuicao) p.distribuicao = {};
            (estoque.representantes || []).forEach(r => { p.distribuicao[r] = 0; });
            // manter IMBEL sempre 0 (é derivado)
            p.distribuicao.IMBEL = 0;
        } catch (e) {}
    });

    // Somar todas as distribuições
    (estoque.registroDistribuicao || []).forEach(d => {
        try {
            const prod = estoque.produtos.find(p => Number(p.id) === Number(d.produtoId) || (d.produtoNome && p.nome === d.produtoNome));
            if (!prod) return;
            const rep = (d.representante || '').toString().toUpperCase();
            if (!rep) return;
            prod.distribuicao[rep] = (prod.distribuicao[rep] || 0) + (Number(d.quantidade) || 0);
        } catch (e) {}
    });

    // Aplicar devoluções: subtrair do origem, somar para destino quando não for IMBEL
    (estoque.registroDevolucoes || []).forEach(dev => {
        try {
            const prod = estoque.produtos.find(p => Number(p.id) === Number(dev.produtoId) || (dev.produtoNome && p.nome === dev.produtoNome));
            if (!prod) return;
            const origem = (dev.origem || '').toString().toUpperCase();
            const destino = (dev.destino || '').toString().toUpperCase();
            prod.distribuicao[origem] = Math.max(0, (prod.distribuicao[origem] || 0) - (Number(dev.quantidade) || 0));
            if (destino && destino !== 'IMBEL') {
                prod.distribuicao[destino] = (prod.distribuicao[destino] || 0) + (Number(dev.quantidade) || 0);
            }
        } catch (e) {}
    });
}

// Função de diagnóstico disponível no console para inspecionar um produto rapidamente
function diagnosticarProduto(produtoId) {
    try {
        const p = estoque.produtos.find(x => Number(x.id) === Number(produtoId) || String(x.id) === String(produtoId));
        if (!p) { console.warn('Produto não encontrado:', produtoId); return null; }
        const estoqueConsolidado = Number(p.estoqueConsolidado) || 0;
        const totalDistribuido = (estoque.registroDistribuicao || []).reduce((s, d) => {
            try {
                if (Number(d.produtoId) === Number(p.id) || (d.produtoNome && d.produtoNome === p.nome)) return s + (Number(d.quantidade) || 0);
            } catch (e) {}
            return s;
        }, 0);
        const totalDevolvidoParaImbel = (estoque.registroDevolucoes || []).reduce((s, d) => {
            try {
                if ((Number(d.produtoId) === Number(p.id) || (d.produtoNome && d.produtoNome === p.nome)) && ((d.destino || '').toString().toUpperCase() === 'IMBEL')) return s + (Number(d.quantidade) || 0);
            } catch (e) {}
            return s;
        }, 0);
        const totalVendas = obterTotalVendasProduto(p);
        const imbelDisponivel = estoqueConsolidado - totalDistribuido + totalDevolvidoParaImbel;
        const distribuicaoPorRep = {};
        (estoque.representantes || []).forEach(r => { distribuicaoPorRep[r] = p.distribuicao ? (p.distribuicao[r] || 0) : 0; });
        const result = { produtoId: p.id, nome: p.nome, estoqueConsolidado, totalDistribuido, totalDevolvidoParaImbel, totalVendas, imbelDisponivel, distribuicaoPorRep };
        console.debug('Diagnóstico do produto:', result);
        return result;
    } catch (e) { console.error('Erro no diagnóstico:', e); return null; }
}
window.diagnosticarProduto = diagnosticarProduto;

// Atalho por nome (útil para console): `diagnosticarProdutoPorNome("NOME DO PRODUTO")`
function diagnosticarProdutoPorNome(nome) {
    if (!nome) return null;
    const p = (estoque.produtos || []).find(x => (x.nome || '').toString().toUpperCase() === nome.toString().toUpperCase());
    if (!p) { console.warn('Produto não encontrado por nome:', nome); return null; }
    return diagnosticarProduto(p.id);
}
window.diagnosticarProdutoPorNome = diagnosticarProdutoPorNome;

// =============================
// Firebase (inicialização e helpers)
// =============================
// Atualiza o indicador visual de status do Firestore (presente no header)
function updateFirestoreStatus(connected, lastSyncDate, message) {
    try {
        const el = document.getElementById('firestoreStatus');
        const dot = document.getElementById('fsDot');
        const text = document.getElementById('fsText');
        if (!el || !dot || !text) return;
        if (connected) {
            dot.classList.remove('fs-offline');
            dot.classList.remove('fs-warning');
            dot.classList.add('fs-online');
            const label = message || 'Cloud: conectado';
            if (lastSyncDate) {
                const dt = (lastSyncDate instanceof Date) ? lastSyncDate : new Date(lastSyncDate);
                text.textContent = `${label} — último sync: ${dt.toLocaleString('pt-BR')}`;
            } else {
                text.textContent = message || 'Cloud: conectado — sem sync';
            }
        } else {
            dot.classList.remove('fs-online');
            dot.classList.remove('fs-warning');
            dot.classList.add('fs-offline');
            text.textContent = message || 'Cloud: desconectado';
        }
    } catch (e) {
        // ignore UI update errors
    }
}

try {
    if (typeof firebase !== 'undefined') {
        const firebaseConfig = {
            apiKey: "AIzaSyBizembCnAJpVe4TCcTTJvCickREOa_f1Y",
            authDomain: "estoquefi.firebaseapp.com",
            databaseURL: "https://estoquefi-default-rtdb.firebaseio.com",
            projectId: "estoquefi",
            storageBucket: "estoquefi.firebasestorage.app",
            messagingSenderId: "339770116384",
            appId: "1:339770116384:web:3b51acfbc9f18162c5af45",
            measurementId: "G-RVK6BC5TDP"
        };

        // Inicializa apenas se ainda não inicializado
        if (!firebase.apps || firebase.apps.length === 0) {
            firebase.initializeApp(firebaseConfig);
        }

        // Instância do Firestore para uso nas funções abaixo
        try {
            window.firestoreDB = firebase.firestore();
            // Não tentar ler do cloud antes da autenticação para evitar erros de permissão
            updateFirestoreStatus(true, null, 'Cloud: aguardando login');
        } catch (e) {
            console.warn('Firestore não disponível:', e);
            window.firestoreDB = null;
            updateFirestoreStatus(false, null, 'Cloud: não disponível');
        }
    } else {
        console.warn('Firebase SDK não carregado — funções cloud desativadas.');
        // atualizar UI se possível
        try { updateFirestoreStatus(false, null, 'SDK não carregado'); } catch (e) {}
    }
} catch (e) {
    console.error('Erro inicializando Firebase:', e);
}



// ID da venda que está sendo editada (null quando criando nova)
let vendaEditandoId = null;
// ID do produto que está sendo editado no modal de produto (null quando criando novo)
let produtoEditandoId = null;

// Dados iniciais com PREÇOS baseados na planilha - SEM dados de distribuição/vendas (zerados)
const dadosIniciais = [
    {
        nome: 'CARABINA IA2 5,56',
        preco: 10420.75,
        distribuicao: { KOLTE: 0, ISA: 0, LC: 0, ADES: 0, FL: 0, IMBEL: 0 },
        vendas: { KOLTE: 0, ISA: 0, LC: 0, ADES: 0, FL: 0, IMBEL: 0 }
    },
    {
        nome: 'CARABINA IA2 7,62',
        preco: 12690.21,
        distribuicao: { KOLTE: 0, ISA: 0, LC: 0, ADES: 0, FL: 0, IMBEL: 0 },
        vendas: { KOLTE: 0, ISA: 0, LC: 0, ADES: 0, FL: 0, IMBEL: 0 }
    },
    {
        nome: 'FACA CAMPANHA AMZ',
        preco: 360.00,
        distribuicao: { KOLTE: 0, ISA: 0, LC: 0, ADES: 0, FL: 0, IMBEL: 0 },
        vendas: { KOLTE: 0, ISA: 0, LC: 0, ADES: 0, FL: 0, IMBEL: 0 }
    },
    {
        nome: 'FACA POLICIAL AMZ',
        preco: 352.45,
        distribuicao: { KOLTE: 0, ISA: 0, LC: 0, ADES: 0, FL: 0, IMBEL: 0 },
        vendas: { KOLTE: 0, ISA: 0, LC: 0, ADES: 0, FL: 0, IMBEL: 0 }
    },
    {
        nome: 'FACA POLICIAL IA2',
        preco: 380.00,
        distribuicao: { KOLTE: 0, ISA: 0, LC: 0, ADES: 0, FL: 0, IMBEL: 0 },
        vendas: { KOLTE: 0, ISA: 0, LC: 0, ADES: 0, FL: 0, IMBEL: 0 }
    },
    {
        nome: 'FUZIL DE ALTA PRECISÃO IMBEL 308 AGLC (COMPLETO)',
        preco: 13500.00,
        distribuicao: { KOLTE: 0, ISA: 0, LC: 0, ADES: 0, FL: 0, IMBEL: 0 },
        vendas: { KOLTE: 0, ISA: 0, LC: 0, ADES: 0, FL: 0, IMBEL: 0 }
    },
    {
        nome: 'PISTOLA .40 GC MD7 C/ ADC',
        preco: 5159.71,
        distribuicao: { KOLTE: 0, ISA: 0, LC: 0, ADES: 0, FL: 0, IMBEL: 0 },
        vendas: { KOLTE: 0, ISA: 0, LC: 0, ADES: 0, FL: 0, IMBEL: 0 }
    },
    {
        nome: 'PISTOLA 380 GC MD1 C/ ADC',
        preco: 5219.54,
        distribuicao: { KOLTE: 0, ISA: 0, LC: 0, ADES: 0, FL: 0, IMBEL: 0 },
        vendas: { KOLTE: 0, ISA: 0, LC: 0, ADES: 0, FL: 0, IMBEL: 0 }
    },
    {
        nome: 'PISTOLA 380 GC MD1 S/ ADC',
        preco: 5406.19,
        distribuicao: { KOLTE: 0, ISA: 0, LC: 0, ADES: 0, FL: 0, IMBEL: 0 },
        vendas: { KOLTE: 0, ISA: 0, LC: 0, ADES: 0, FL: 0, IMBEL: 0 }
    },
    {
        nome: 'PISTOLA 380 GC MD2 C/ ADC',
        preco: 5162.57,
        distribuicao: { KOLTE: 0, ISA: 0, LC: 0, ADES: 0, FL: 0, IMBEL: 0 },
        vendas: { KOLTE: 0, ISA: 0, LC: 0, ADES: 0, FL: 0, IMBEL: 0 }
    },
    {
        nome: 'PISTOLA 380 GC MD2 S/ ADC',
        preco: 5207.87,
        distribuicao: { KOLTE: 0, ISA: 0, LC: 0, ADES: 0, FL: 0, IMBEL: 0 },
        vendas: { KOLTE: 0, ISA: 0, LC: 0, ADES: 0, FL: 0, IMBEL: 0 }
    },
    {
        nome: 'PISTOLA 9 GC MD1 S/ ADC',
        preco: 5236.30,
        distribuicao: { KOLTE: 0, ISA: 0, LC: 0, ADES: 0, FL: 0, IMBEL: 0 },
        vendas: { KOLTE: 0, ISA: 0, LC: 0, ADES: 0, FL: 0, IMBEL: 0 }
    }
];

// ========================================
// FUNÇÕES DE INICIALIZAÇÃO
// ========================================

function popularSelectRepresentantes(selectId, incluirImbel = true) {
    const sel = document.getElementById(selectId);
    if (!sel) return;
    const reps = estoque.representantes || 
        ['KOLTE', 'ISA', 'LC', 'ADES', 'FL'];
    // evitar duplicar IMBEL caso esteja presente em `estoque.representantes`
    const repsSemImbel = reps.filter(r => String(r).toUpperCase() !== 'IMBEL');
    sel.innerHTML = '<option value="">Selecione...</option>'
        + (incluirImbel ? '<option value="IMBEL">IMBEL (Venda Direta)</option>' : '')
        + repsSemImbel.map(r => `<option value="${r}">${r}</option>`).join('');
}

async function inicializar() {
    carregarDados();
    // Load contract config into Configurações tab
    try {
        const cfg = carregarConfigContrato();
        const anoEl = document.getElementById('configContratoAno');
        const proxEl = document.getElementById('configContratoProximo');
        if (anoEl) anoEl.value = cfg.ano;
        if (proxEl) proxEl.value = cfg.proximo;
        atualizarPreviaContrato();
    } catch (e) {}
    try { verificarExpiracaoPrecificacoes(); } catch (e) {}
    try { inicializarImpostosPreDefinidos(); } catch (e) { console.warn('Inicialização de impostos predefinidos falhou:', e); }

    try { inicializarImpostosEditaveis(); } catch (e) {}
    try { inicializarICMSEditavel(); } catch (e) {}

    // Sync automático com cloud ocorre após autenticação (onAuthStateChanged)

    try { renderizarTabela(); } catch (e) { console.error('renderizarTabela:', e); }
    try { renderizarCadastroProdutos(); } catch (e) { console.error('renderizarCadastroProdutos:', e); }
    try { renderizarDashboard(); } catch (e) { console.error('renderizarDashboard:', e); }
    try { renderizarRegistroVendas(); } catch (e) { console.error('renderizarVendas:', e); }
    try { renderizarRegistroDistribuicao(); } catch (e) { console.error('renderizarDist:', e); }
    try { renderizarControleEnvio(); } catch (e) { console.error('renderizarEnvio:', e); }
    try { renderizarClientes(); } catch (e) { console.error('renderizarClientes:', e); }
    try { atualizarKPIsClientes(); } catch (e) {}
    try { atualizarDatalistClientes(); } catch (e) {}
    try { renderizarPropostas(); } catch (e) { console.error('renderizarPropostas:', e); }
    try { atualizarKPIsPropostas(); } catch (e) {}
    try { renderizarPrecificacao(); } catch (e) { console.error('renderizarPrecif:', e); }
    try { atualizarSelectsProdutos(); } catch (e) {}
    try { atualizarSelectsRelatorios(); } catch (e) {}
    try { atualizarEstatisticas(); } catch (e) {}
    try { atualizarData(); } catch (e) {}

    // Popular selects visíveis de representantes (filtros) na inicialização
    try { popularSelectRepresentantes('filtroDistribuicaoRep', true); } catch (e) {}
    try { popularSelectRepresentantes('filtroRepresentante', true); } catch (e) {}
    try { popularSelectRepresentantes('filtroControleEnvioRep', true); } catch (e) {}
    try { popularSelectRepresentantes('filtroRelatoriosRep', true); } catch (e) {}

    // Check proposal alerts every hour
    if (!window.__PROPOSTA_ALERTAS_INTERVALO__) {
        window.__PROPOSTA_ALERTAS_INTERVALO__ = setInterval(() => {
            try { verificarAlertasEstoque(); } catch (e) {}
        }, 60 * 60 * 1000);
    }

    // Re-check when user returns to tab
    if (!window.__PROPOSTA_ALERTAS_VISIBILITY__) {
        document.addEventListener('visibilitychange', () => {
            if (!document.hidden) {
                try { verificarAlertasEstoque(); } catch (e) {}
            }
        });
        window.__PROPOSTA_ALERTAS_VISIBILITY__ = true;
    }

    // Reativar auto-save: habilita salvamento automático (debounced + periódico)
    try { window.__AUTO_SAVE_CLOUD.enabled = true; } catch (e) {}
    try { iniciarAutoSaveCloud(); } catch (e) { console.warn('Falha ao iniciar auto-save:', e); }
}

function normalizarPrecificacoesCliente(origem) {
    const lista = Array.isArray(origem) ? origem : [];
    return lista.map((p, i) => {
        const item = (p && typeof p === 'object') ? p : {};
        return {
            ...item,
            versao: item.versao || (i + 1),
            status: item.status || 'ativa',
            propostaId: item.propostaId || null,
            descricao: item.descricao || ''
        };
    });
}

function carregarDados() {
    const dadosSalvos = localStorage.getItem('estoqueArmasV2');
    if (dadosSalvos) {
        try {
            estoque = JSON.parse(dadosSalvos);
        } catch (e) {
            console.error('Falha ao interpretar localStorage estoqueArmasV2:', e);
            localStorage.removeItem('estoqueArmasV2');
            estoque = {
                produtos: [],
                representantes: ['KOLTE', 'ISA', 'LC', 'ADES', 'FL', 'IMBEL'],
                registroVendas: [],
                registroDistribuicao: [],
                registroDevolucoes: [],
                registroEntradas: [],
                controleEnvio: {},
                auditoriaVendas: [],
                fechamentosComissoes: [],
                clientes: [],
                propostas: [],
                precificacao: {},
                precificacaoConfig: null,
                precificacoesCliente: []
            };
        }
        // Garantir que registroVendas existe
        if (!estoque.registroVendas) {
            estoque.registroVendas = [];
        }
        // Garantir que registroDistribuicao existe
        if (!estoque.registroDistribuicao) {
            estoque.registroDistribuicao = [];
        }
        // Garantir que registroDevolucoes existe
        if (!estoque.registroDevolucoes) {
            estoque.registroDevolucoes = [];
        }
        // Garantir que registroEntradas existe (novidade)
        if (!estoque.registroEntradas) {
            estoque.registroEntradas = [];
        }
        // Garantir que controleEnvio existe
        if (!estoque.controleEnvio) {
            estoque.controleEnvio = {};
        }
        if (!Array.isArray(estoque.auditoriaVendas)) {
            estoque.auditoriaVendas = [];
        }
        if (!Array.isArray(estoque.fechamentosComissoes)) {
            estoque.fechamentosComissoes = [];
        }
        if (!Array.isArray(estoque.clientes)) {
            estoque.clientes = [];
        }
        clientes = estoque.clientes;
        if (!Array.isArray(estoque.propostas)) {
            estoque.propostas = [];
        }
        propostas = estoque.propostas;
        if (!Array.isArray(estoque.precificacoesCliente)) {
            estoque.precificacoesCliente = [];
        }
        estoque.precificacoesCliente = normalizarPrecificacoesCliente(estoque.precificacoesCliente);
        // Carregar configuração de alertas se presente
        try {
            const cfg = localStorage.getItem('configAlertas');
            if (cfg) configAlertas = JSON.parse(cfg);
        } catch (e) { console.warn('Erro ao carregar configAlertas:', e); }
        // Migrar produtos antigos para nova propriedade `estoqueConsolidado` quando necessário
        try {
            if (Array.isArray(estoque.produtos)) {
                estoque.produtos.forEach(p => {
                    if (typeof p.estoqueConsolidado === 'undefined' || p.estoqueConsolidado === null) {
                        // fallback: somar todas as chaves de distribuição armazenadas
                        try {
                            const totalDisp = Object.keys(p.distribuicao || {}).reduce((s, k) => s + (Number(p.distribuicao[k]) || 0), 0);
                            p.estoqueConsolidado = totalDisp;
                        } catch (inner) { p.estoqueConsolidado = 0; }
                    }
                    // padronizar: manter distribuições por representantes; zerar/ignorar IMBEL armazenado (IMBEL será calculado dinamicamente a partir do consolidado)
                    try { if (!p.distribuicao) p.distribuicao = {}; p.distribuicao.IMBEL = 0; } catch (e) {}
                });
            }
        } catch (e) { console.warn('Migração estoqueConsolidado falhou:', e); }
        if (estoque.precificacao && typeof estoque.precificacao === 'object') {
            precificacao = estoque.precificacao;
        } else {
            precificacao = {};
            estoque.precificacao = precificacao;
        }
        // carregar precificações por cliente, se presente
        try { precificacoesCliente = normalizarPrecificacoesCliente(estoque.precificacoesCliente); } catch (e) { precificacoesCliente = []; }
        estoque.precificacoesCliente = precificacoesCliente;
        tabelaAliquotas = (estoque.tabelaAliquotas && typeof estoque.tabelaAliquotas === 'object')
            ? estoque.tabelaAliquotas
            : {};
        tabelaICMS = Array.isArray(estoque.tabelaICMS) ? estoque.tabelaICMS : [];
        categoriaPorProduto = (estoque.categoriaPorProduto && typeof estoque.categoriaPorProduto === 'object')
            ? estoque.categoriaPorProduto
            : {};
        // Reconstruir distribuições por produto a partir dos registros (corrige inconsistências legadas)
        try { reconstruirDistribuicaoAPartirDeRegistros(); } catch (e) { /* ignore */ }
    } else {
        estoque.produtos = dadosIniciais.map((item, index) => ({
            id: index + 1,
            nome: item.nome,
            preco: item.preco,
            distribuicao: { ...item.distribuicao },
            vendas: { ...item.vendas }
        }));
        estoque.registroVendas = [];
        estoque.registroDistribuicao = [];
        estoque.registroDevolucoes = [];
        estoque.registroEntradas = [];
        estoque.controleEnvio = {};
        estoque.auditoriaVendas = [];
        estoque.fechamentosComissoes = [];
        estoque.clientes = [];
        clientes = estoque.clientes;
        estoque.precificacao = {};
        estoque.precificacaoConfig = null;
        precificacao = estoque.precificacao;
        tabelaAliquotas = {};
        tabelaICMS = [];
        categoriaPorProduto = {};
        estoque.precificacoesCliente = [];
        estoque.tabelaAliquotas = tabelaAliquotas;
        estoque.tabelaICMS = tabelaICMS;
        estoque.categoriaPorProduto = categoriaPorProduto;
        salvarDados();
    }
}

function salvarDados() {
    // Sincroniza os objetos de precificação no estado persistido
    estoque.precificacao = precificacao;
    estoque.tabelaAliquotas = tabelaAliquotas;
    estoque.tabelaICMS = tabelaICMS;
    estoque.categoriaPorProduto = categoriaPorProduto;
    // Persistir precificações salvas por cliente
    try { estoque.precificacoesCliente = precificacoesCliente || []; } catch (e) {}
    // Persistir tabelas de impostos editáveis
    try { estoque.impostosEditaveis = impostosEditaveis || {}; } catch (e) {}
    try { estoque.icmsEditavelPJ = icmsEditavelPJ || {}; } catch (e) {}
    try { estoque.icmsEditavelPF = icmsEditavelPF || {}; } catch (e) {}
    // marca hora local de atualização para comparação com o remoto
    try { estoque._localUpdatedAt = new Date().toISOString(); } catch (e) {}
    localStorage.setItem('estoqueArmasV2', JSON.stringify(estoque));
    atualizarEstatisticas();

    // agendar salvamento no cloud (debounced) se habilitado
    try { scheduleCloudSaveDebounced(); } catch (e) {}
    try {
        if (!_cloudSyncedRecently) _dadosAlterados = true;
    } catch (e) {}
}

// =============================
// Funções para salvar/carregar no Firestore
// =============================
async function salvarNoCloud() {
    if (!window.firestoreDB) {
        console.warn('Firestore não inicializado. Impossível salvar no cloud.');
        return false;
    }
    const indicator = document.getElementById('cloudSaveIndicator');
    if (indicator) { indicator.textContent = '☁️ Salvando...'; indicator.style.color = '#d97706'; }
    try {
        const docRef = window.firestoreDB.collection('app_data').doc('latest');
        await docRef.set({
            estado: estoque,
            precificacao,
            tabelaAliquotas,
            tabelaICMS,
            categoriaPorProduto,
            precificacoesCliente: precificacoesCliente || [],
            impostosEditaveis: impostosEditaveis || {},
            icmsEditavelPJ: icmsEditavelPJ || {},
            icmsEditavelPF: icmsEditavelPF || {},
            updatedAt: firebase.firestore.FieldValue.serverTimestamp()
        });
        // ler o documento para obter o updatedAt do servidor
            try {
            const savedDoc = await docRef.get();
            const d = savedDoc && savedDoc.exists ? savedDoc.data() : null;
            const updatedAt = d && d.updatedAt ? d.updatedAt.toDate() : new Date();
            updateFirestoreStatus(true, updatedAt, 'Cloud: salvo');
        } catch (inner) {
            updateFirestoreStatus(true, new Date(), 'Cloud: salvo');
        }
        if (indicator) {
            indicator.textContent = '✅ Salvo';
            indicator.style.color = '#16a34a';
            setTimeout(() => { if (indicator) indicator.textContent = ''; }, 3000);
        }
        try {
            _cloudSyncedRecently = true;
            _dadosAlterados = false;
            if (__cloudSyncResetTimer) clearTimeout(__cloudSyncResetTimer);
            __cloudSyncResetTimer = setTimeout(() => { _cloudSyncedRecently = false; }, 30000);
        } catch (e) {}
        console.debug('Dados salvos no Firestore (coleção app_data / doc latest)');
        return true;
    } catch (e) {
        console.error('Erro salvando no Firestore:', e);
        updateFirestoreStatus(false, null, 'Cloud: erro ao salvar');
        if (indicator) { indicator.textContent = '❌ Erro ao salvar'; indicator.style.color = '#dc2626'; }
        return false;
    }
}

// UI wrappers for manual save/load triggered by user buttons
async function salvarNoCloudUI() {
    try {
        mostrarNotificacao('Salvando dados no cloud...', 'info');
        const ok = await salvarNoCloud();
        if (ok) mostrarNotificacao('Dados salvos no Firestore com sucesso.', 'success');
        else mostrarNotificacao('Falha ao salvar no Firestore. Veja o console para detalhes.', 'error');
        return ok;
    } catch (e) {
        console.error('salvarNoCloudUI erro:', e);
        mostrarNotificacao('Erro ao salvar no cloud.', 'error');
        return false;
    }
}

async function carregarDoCloudUI() {
    try {
        const confirmed = confirm('Carregar do cloud substituirá os dados locais. Deseja continuar?');
        if (!confirmed) return false;
        mostrarNotificacao('Carregando dados do cloud...', 'info');
        const ok = await carregarDoCloud({ confirmOverwrite: false });
        if (ok) mostrarNotificacao('Dados carregados do Firestore com sucesso.', 'success');
        else mostrarNotificacao('Nenhum backup encontrado no Firestore ou falha ao carregar.', 'warning');
        return ok;
    } catch (e) {
        console.error('carregarDoCloudUI erro:', e);
        mostrarNotificacao('Erro ao carregar do cloud.', 'error');
        return false;
    }
}

async function carregarDoCloud({confirmOverwrite=true} = {}) {
    if (!window.firestoreDB) {
        console.warn('Firestore não inicializado. Impossível carregar do cloud.');
        return false;
    }
    try {
        const docRef = window.firestoreDB.collection('app_data').doc('latest');
        const doc = await docRef.get();
        if (!doc.exists) {
            console.warn('Nenhum backup encontrado no Firestore.');
            updateFirestoreStatus(true, null, 'Cloud: pronto (sem backup)');
            return false;
        }
        const data = doc.data();
        // Restaurar precificações salvas por cliente (se houver)
        try {
            precificacoesCliente = normalizarPrecificacoesCliente(data.precificacoesCliente || []);
            estoque.precificacoesCliente = precificacoesCliente;
        } catch (e) { precificacoesCliente = []; }
        if (!data || !data.estado) {
            console.warn('Documento encontrado não contém campo estado.');
            return false;
        }
        // obter timestamp de atualização remoto se disponível
        try {
            const updatedAt = data.updatedAt ? data.updatedAt.toDate() : null;
            updateFirestoreStatus(true, updatedAt, 'Cloud: carregado');
        } catch (inner) { updateFirestoreStatus(true, null, 'Cloud: carregado'); }
        if (confirmOverwrite) {
            const ok = confirm('Carregar dados do cloud irá substituir os dados locais. Deseja continuar?');
            if (!ok) return false;
        }
        estoque = data.estado;
        estoque.precificacoesCliente = normalizarPrecificacoesCliente(estoque.precificacoesCliente);
        if (!Array.isArray(estoque.auditoriaVendas)) estoque.auditoriaVendas = [];
        if (!Array.isArray(estoque.fechamentosComissoes)) estoque.fechamentosComissoes = [];
        if (!Array.isArray(estoque.clientes)) estoque.clientes = [];
        clientes = estoque.clientes;
        if (!Array.isArray(estoque.propostas)) estoque.propostas = [];
        propostas = estoque.propostas;
        if (estoque.precificacao && typeof estoque.precificacao === 'object') precificacao = estoque.precificacao;
        else { estoque.precificacao = {}; precificacao = estoque.precificacao; }

        precificacao = (data.precificacao && typeof data.precificacao === 'object')
            ? data.precificacao
            : (estoque.precificacao || {});
        tabelaAliquotas = (data.tabelaAliquotas && typeof data.tabelaAliquotas === 'object')
            ? data.tabelaAliquotas
            : (estoque.tabelaAliquotas || {});
        tabelaICMS = Array.isArray(data.tabelaICMS)
            ? data.tabelaICMS
            : (Array.isArray(estoque.tabelaICMS) ? estoque.tabelaICMS : []);
        categoriaPorProduto = (data.categoriaPorProduto && typeof data.categoriaPorProduto === 'object')
            ? data.categoriaPorProduto
            : (estoque.categoriaPorProduto || {});
        try { verificarExpiracaoPrecificacoes(); } catch (e) {}

        try { inicializarImpostosPreDefinidos(); } catch (e) { console.warn('Inicializar impostos predefinidos após cloud falhou:', e); }

        salvarDados();
        renderizarTabela();
        renderizarDashboard();
        renderizarRegistroVendas();
        renderizarRegistroDistribuicao();
        renderizarControleEnvio();
        renderizarClientes();
        atualizarKPIsClientes();
        atualizarDatalistClientes();
        renderizarPropostas();
        atualizarKPIsPropostas();
        if (abaAtiva === 'precificacao') renderizarPrecificacao();
        atualizarSelectsProdutos();
        atualizarSelectsRelatorios();
        console.debug('Dados carregados do Firestore com sucesso.');
        return true;
    } catch (e) {
        console.error('Erro carregando do Firestore:', e);
        return false;
    }
}

// Carrega automaticamente do cloud se o documento remoto for mais recente que o local
async function carregarDoCloudAuto() {
    if (!window.firestoreDB) return false;
    // Evitar leitura sem autenticação (causa Missing or insufficient permissions)
    try {
        if (!firebase || !firebase.auth || !firebase.auth().currentUser) {
            updateFirestoreStatus(true, null, 'Cloud: aguardando login');
            return false;
        }
    } catch (e) {
        return false;
    }
    try {
        const docRef = window.firestoreDB.collection('app_data').doc('latest');
        const doc = await docRef.get();
        if (!doc.exists) return false;
        const data = doc.data();
        const remoteUpdated = data.updatedAt ? data.updatedAt.toDate().getTime() : null;
        const localUpdated = estoque._localUpdatedAt ? new Date(estoque._localUpdatedAt).getTime() : 0;
        if (remoteUpdated && remoteUpdated > localUpdated) {
            // substituir local automaticamente
            estoque = data.estado;
            estoque.precificacoesCliente = normalizarPrecificacoesCliente(estoque.precificacoesCliente);
            if (!Array.isArray(estoque.auditoriaVendas)) estoque.auditoriaVendas = [];
            if (!Array.isArray(estoque.fechamentosComissoes)) estoque.fechamentosComissoes = [];
            if (!Array.isArray(estoque.clientes)) estoque.clientes = [];
            clientes = estoque.clientes;
            if (!Array.isArray(estoque.propostas)) estoque.propostas = [];
            propostas = estoque.propostas;
            if (estoque.precificacao && typeof estoque.precificacao === 'object') precificacao = estoque.precificacao;
            else { estoque.precificacao = {}; precificacao = estoque.precificacao; }

            precificacao = (data.precificacao && typeof data.precificacao === 'object')
                ? data.precificacao
                : (estoque.precificacao || {});
            tabelaAliquotas = (data.tabelaAliquotas && typeof data.tabelaAliquotas === 'object')
                ? data.tabelaAliquotas
                : (estoque.tabelaAliquotas || {});
            tabelaICMS = Array.isArray(data.tabelaICMS)
                ? data.tabelaICMS
                : (Array.isArray(estoque.tabelaICMS) ? estoque.tabelaICMS : []);
            categoriaPorProduto = (data.categoriaPorProduto && typeof data.categoriaPorProduto === 'object')
                ? data.categoriaPorProduto
                : (estoque.categoriaPorProduto || {});

            // Restaurar tabelas de impostos editáveis quando disponíveis
            impostosEditaveis = (data.impostosEditaveis && typeof data.impostosEditaveis === 'object') ? data.impostosEditaveis : (estoque.impostosEditaveis || {});
            icmsEditavelPJ    = (data.icmsEditavelPJ && typeof data.icmsEditavelPJ === 'object') ? data.icmsEditavelPJ : (estoque.icmsEditavelPJ || {});
            icmsEditavelPF    = (data.icmsEditavelPF && typeof data.icmsEditavelPF === 'object') ? data.icmsEditavelPF : (estoque.icmsEditavelPF || {});
            try { inicializarImpostosEditaveis(); } catch (e) {}
            try { inicializarICMSEditavel(); } catch (e) {}

            salvarDados();
            renderizarTabela();
            renderizarDashboard();
            renderizarRegistroVendas();
            renderizarRegistroDistribuicao();
            renderizarControleEnvio();
            renderizarClientes();
            atualizarKPIsClientes();
            atualizarDatalistClientes();
            renderizarPropostas();
            atualizarKPIsPropostas();
            if (abaAtiva === 'precificacao') renderizarPrecificacao();
            atualizarSelectsProdutos();
            atualizarSelectsRelatorios();
            console.debug('Dados carregados automaticamente do Firestore (remoto mais recente).');
            return true;
        }
        return false;
    } catch (e) {
        // Quando regras bloquearem leitura, não poluir console com erro fatal
        console.warn('carregarDoCloudAuto: leitura não permitida pelo perfil atual.');
        return false;
    }
}

// ============================
// Auto-save (debounced) helpers
// ============================
window.__AUTO_SAVE_CLOUD = {
    enabled: true,
    debounceMs: 2500,
    timerId: null,
    inProgress: false
};

function scheduleCloudSaveDebounced() {
    if (!window.__AUTO_SAVE_CLOUD.enabled) return;
    if (!window.firestoreDB) return;
    if (!isCurrentUserAdmin()) return;
    if (window.__AUTO_SAVE_CLOUD.timerId) clearTimeout(window.__AUTO_SAVE_CLOUD.timerId);
    window.__AUTO_SAVE_CLOUD.timerId = setTimeout(async () => {
        window.__AUTO_SAVE_CLOUD.timerId = null;
        if (window.__AUTO_SAVE_CLOUD.inProgress) return;
        window.__AUTO_SAVE_CLOUD.inProgress = true;
        try {
            await salvarNoCloud();
        } catch (e) {
            console.error('Auto-save falhou:', e);
        } finally {
            window.__AUTO_SAVE_CLOUD.inProgress = false;
        }
    }, window.__AUTO_SAVE_CLOUD.debounceMs);
}

function iniciarAutoSaveCloud() {
    // ativa auto-save se Firestore presente
    if (!window.firestoreDB) return;
    window.__AUTO_SAVE_CLOUD.enabled = true;
    // salvar a cada X minutos também (fallback periódico)
    if (!window.__AUTO_SAVE_CLOUD.periodicId) {
        window.__AUTO_SAVE_CLOUD.periodicId = setInterval(() => {
            if (!window.__AUTO_SAVE_CLOUD.inProgress) {
                salvarNoCloud().catch(e => console.error('Auto-save periódico falhou:', e));
            }
        }, 1000 * 60 * 5); // a cada 5 minutos
    }
}

function pararAutoSaveCloud() {
    window.__AUTO_SAVE_CLOUD.enabled = false;
    if (window.__AUTO_SAVE_CLOUD.timerId) {
        clearTimeout(window.__AUTO_SAVE_CLOUD.timerId);
        window.__AUTO_SAVE_CLOUD.timerId = null;
    }
    if (window.__AUTO_SAVE_CLOUD.periodicId) {
        clearInterval(window.__AUTO_SAVE_CLOUD.periodicId);
        window.__AUTO_SAVE_CLOUD.periodicId = null;
    }
}

function atualizarData() {
    const agora = new Date();
    const opcoes = { weekday: 'short', day: 'numeric', month: 'short', year: 'numeric' };
    const dataFormatada = agora.toLocaleDateString('pt-BR', opcoes);
    const dataAtualEl = document.getElementById('dataAtual');
    if (dataAtualEl) dataAtualEl.textContent = dataFormatada;
}

function atualizarEstatisticas() {
    let totalEstoque = 0;
    let totalVendas = 0;
    let valorTotalVendas = 0;

    estoque.produtos.forEach(produto => {
        estoque.representantes.forEach(rep => {
            const disp = produto.distribuicao[rep] || 0;
            const venda = produto.vendas[rep] || 0;
            totalEstoque += (disp - venda);
            totalVendas += venda;
        });
    });

    // Corrigir cálculo de faturamento: somar valorTotal de todas vendas registradas
    valorTotalVendas = 0;
    if (Array.isArray(estoque.registroVendas)) {
        estoque.registroVendas.forEach(venda => {
            if (venda && venda.cancelado) return; // ignorar vendas canceladas ao somar faturamento
            if (Array.isArray(venda.items) && venda.items.length > 0) {
                venda.items.forEach(it => {
                    valorTotalVendas += it.valorTotal || 0;
                });
            } else {
                valorTotalVendas += venda.valorTotal || 0;
            }
        });
    }

    const totalProdutosEl = document.getElementById('totalProdutos');
    if (totalProdutosEl) totalProdutosEl.textContent = estoque.produtos.length;

    const totalEstoqueEl = document.getElementById('totalEstoque');
    if (totalEstoqueEl) totalEstoqueEl.textContent = totalEstoque.toLocaleString('pt-BR');

    const totalVendasEl = document.getElementById('totalVendas');
    if (totalVendasEl) totalVendasEl.textContent = totalVendas.toLocaleString('pt-BR');

    const dashContainer = document.getElementById('tab-dashboard');
    const valorTotalVendasEl = (dashContainer && dashContainer.querySelector)
        ? (dashContainer.querySelector('#valorTotalVendas') || document.getElementById('valorTotalVendas'))
        : document.getElementById('valorTotalVendas');
    if (valorTotalVendasEl) valorTotalVendasEl.textContent = formatarMoedaValor(valorTotalVendas);
    // Calcular total de comissões (5%) excluindo vendas da IMBEL
    let totalComissoes = 0;
    if (Array.isArray(estoque.registroVendas)) {
        estoque.registroVendas.forEach(venda => {
            if (venda && venda.cancelado) return; // ignorar vendas canceladas
            const rep = (venda.representante || '').toString().trim().toUpperCase();
            if (rep === 'IMBEL') return; // sem comissão
            let valor = 0;
            if (Array.isArray(venda.items) && venda.items.length > 0) {
                venda.items.forEach(it => {
                    const valorItem = typeof it.valorTotal === 'number'
                        ? it.valorTotal
                        : ((Number(it.valorUnitario) || 0) * (Number(it.quantidade) || 0));
                    valor += valorItem;
                });
            } else {
                valor = typeof venda.valorTotal === 'number'
                    ? venda.valorTotal
                    : ((Number(venda.valorUnitario) || 0) * (Number(venda.quantidade) || 0));
            }
            totalComissoes += (Math.round((valor * 0.05) * 100) / 100);
        });
    }
    try {
        const totalComissoesEl = (dashContainer && dashContainer.querySelector)
            ? (dashContainer.querySelector('#totalComissoes') || document.getElementById('totalComissoes'))
            : document.getElementById('totalComissoes');
        if (totalComissoesEl) totalComissoesEl.textContent = formatarMoedaValor(totalComissoes);
    } catch (e) {}
}

// Helper global: normaliza várias formas de data para YYYY-MM-DD
function parseDateToYYYYMMDD(input) {
    if (!input && input !== 0) return null;
    // Firestore Timestamp-like objects (has toDate)
    try {
        if (input && typeof input.toDate === 'function') {
            const dt = input.toDate();
            if (dt instanceof Date && !isNaN(dt.getTime())) return dt.toISOString().slice(0,10);
        }
    } catch (e) {}
    // Objects with seconds (e.g., { seconds, nanoseconds })
    if (input && typeof input === 'object') {
        if (typeof input.seconds === 'number') {
            const dt = new Date(input.seconds * 1000);
            if (!isNaN(dt.getTime())) return dt.toISOString().slice(0,10);
        }
        if (typeof input._seconds === 'number') {
            const dt = new Date(input._seconds * 1000);
            if (!isNaN(dt.getTime())) return dt.toISOString().slice(0,10);
        }
    }

    if (input instanceof Date) {
        if (isNaN(input.getTime())) return null;
        return input.toISOString().slice(0,10);
    }

    let s = String(input).trim();
    // ISO-like (starts with YYYY-MM-DD)
    const iso = s.match(/^(\d{4}-\d{2}-\d{2})/);
    if (iso) return iso[1];
    // BR format DD/MM/YYYY
    const br = s.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
    if (br) return `${br[3]}-${br[2]}-${br[1]}`;
    // Try parsing general string
    const d = new Date(s);
    if (!isNaN(d.getTime())) return d.toISOString().slice(0,10);
    return null;
}

// Formata uma data (aceita 'YYYY-MM-DD', Date ou strings) para 'DD/MM/YYYY'
function formatDateToDDMMYYYY(input) {
    if (!input && input !== 0) return '-';
    try {
        // Already in YYYY-MM-DD
        if (typeof input === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(input)) {
            const parts = input.split('-');
            return `${parts[2].padStart(2,'0')}/${parts[1].padStart(2,'0')}/${parts[0]}`;
        }
        // Date-like
        const d = (input instanceof Date) ? input : new Date(input);
        if (!isNaN(d.getTime())) {
            const dd = String(d.getDate()).padStart(2,'0');
            const mm = String(d.getMonth() + 1).padStart(2,'0');
            const yyyy = d.getFullYear();
            return `${dd}/${mm}/${yyyy}`;
        }
    } catch (e) {}
    return '-';
}

function formatCpfCnpjMask(value) {
    const digits = String(value || '').replace(/\D/g, '').slice(0, 14);
    if (!digits) return '';

    if (digits.length <= 11) {
        if (digits.length <= 3) return digits;
        if (digits.length <= 6) return `${digits.slice(0,3)}.${digits.slice(3)}`;
        if (digits.length <= 9) return `${digits.slice(0,3)}.${digits.slice(3,6)}.${digits.slice(6)}`;
        return `${digits.slice(0,3)}.${digits.slice(3,6)}.${digits.slice(6,9)}-${digits.slice(9)}`;
    }

    if (digits.length <= 2) return digits;
    if (digits.length <= 5) return `${digits.slice(0,2)}.${digits.slice(2)}`;
    if (digits.length <= 8) return `${digits.slice(0,2)}.${digits.slice(2,5)}.${digits.slice(5)}`;
    if (digits.length <= 12) return `${digits.slice(0,2)}.${digits.slice(2,5)}.${digits.slice(5,8)}/${digits.slice(8)}`;
    return `${digits.slice(0,2)}.${digits.slice(2,5)}.${digits.slice(5,8)}/${digits.slice(8,12)}-${digits.slice(12)}`;
}

function formatPhoneMask(value) {
    const digits = String(value || '').replace(/\D/g, '').slice(0, 11);
    if (!digits) return '';
    if (digits.length <= 2) return `(${digits}`;
    if (digits.length <= 6) return `(${digits.slice(0,2)}) ${digits.slice(2)}`;
    if (digits.length <= 10) return `(${digits.slice(0,2)}) ${digits.slice(2,6)}-${digits.slice(6)}`;
    return `(${digits.slice(0,2)}) ${digits.slice(2,7)}-${digits.slice(7)}`;
}

function formatCurrencyBRLInput(value) {
    const digits = String(value || '').replace(/\D/g, '');
    if (!digits) return '';
    const cents = Number(digits);
    return (cents / 100).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function parseCurrencyBRLToNumber(value) {
    const text = String(value || '').trim();
    if (!text) return 0;
    const normalized = text
        .replace(/\s/g, '')
        .replace(/[R$r$]/g, '')
        .replace(/\./g, '')
        .replace(',', '.');
    const parsed = Number(normalized.replace(/[^0-9.-]/g, ''));
    return Number.isFinite(parsed) ? parsed : 0;
}

function bindImbelMovInputMasks() {
    const cpfInput = document.getElementById('imbel_mov_cpf');
    const telInput = document.getElementById('imbel_mov_tel');
    const valorInput = document.getElementById('imbel_mov_valor');

    if (cpfInput && !cpfInput.dataset.maskBound) {
        cpfInput.addEventListener('input', function() {
            this.value = formatCpfCnpjMask(this.value);
        });
        cpfInput.dataset.maskBound = '1';
    }

    if (telInput && !telInput.dataset.maskBound) {
        telInput.addEventListener('input', function() {
            this.value = formatPhoneMask(this.value);
        });
        telInput.dataset.maskBound = '1';
    }

    if (valorInput && !valorInput.dataset.maskBound) {
        valorInput.addEventListener('input', function() {
            this.value = formatCurrencyBRLInput(this.value);
        });
        valorInput.dataset.maskBound = '1';
    }
}

function normalizarContratoKey(valor) {
    const bruto = (valor ?? '').toString().normalize('NFKC');
    const clean = bruto.replace(/[\u200B-\u200D\uFEFF\s]+/g, '');
    const digitos = clean.replace(/\D+/g, '');
    return digitos ? String(parseInt(digitos, 10)) : clean.toUpperCase();
}

function getUsuarioAtual() {
    let usuario = '';
    try { usuario = (localStorage.getItem('estoqueUsuarioAtual') || '').trim(); } catch (e) {}
    if (!usuario) {
        usuario = (prompt('Informe seu nome/usuário para auditoria:') || '').trim();
        if (!usuario) usuario = 'Usuário';
        try { localStorage.setItem('estoqueUsuarioAtual', usuario); } catch (e) {}
    }
    return usuario;
}

function registrarAuditoriaVenda(acao, vendaAntes, vendaDepois, detalhes = '') {
    if (!Array.isArray(estoque.auditoriaVendas)) estoque.auditoriaVendas = [];
    const base = vendaDepois || vendaAntes || {};
    const contrato = normalizarContratoKey(base.contrato || '');
    const entry = {
        id: Date.now() + Math.floor(Math.random() * 1000),
        quando: new Date().toISOString(),
        quem: getUsuarioAtual(),
        acao: acao,
        contrato: contrato || '-',
        vendaId: base.id || null,
        antes: vendaAntes || null,
        depois: vendaDepois || null,
        detalhes: detalhes || ''
    };
    estoque.auditoriaVendas.push(entry);
    if (estoque.auditoriaVendas.length > 1000) {
        estoque.auditoriaVendas = estoque.auditoriaVendas.slice(-1000);
    }
}

function obterAuditoriaPorContrato(contrato) {
    const key = normalizarContratoKey(contrato || '');
    const lista = Array.isArray(estoque.auditoriaVendas) ? estoque.auditoriaVendas : [];
    return lista.filter(a => normalizarContratoKey(a.contrato || '') === key)
        .sort((a, b) => new Date(b.quando).getTime() - new Date(a.quando).getTime());
}

// ========================================
// RENDER: AUDITORIA DE VENDAS
// ========================================

function renderizarAuditoria() {
    const tbody = document.getElementById('auditoriaTbody');
    if (!tbody) return;

    const busca  = (document.getElementById('auditoriaBusca')?.value || '').toLowerCase();
    const filtroAcao = document.getElementById('auditoriaFiltroAcao')?.value || '';

    const lista = (estoque.auditoriaVendas || [])
        .slice()
        .reverse()
        .filter(a => {
            if (filtroAcao && !(a.acao||'').toUpperCase().includes(filtroAcao)) return false;
            if (busca) {
                return (a.contrato||'').toLowerCase().includes(busca) ||
                             (a.quem||'').toLowerCase().includes(busca) ||
                             (a.acao||'').toLowerCase().includes(busca) ||
                             (a.detalhes||'').toLowerCase().includes(busca);
            }
            return true;
        });

    const acaoColor = {
        'CRIAÇÃO':      { bg:'#f0fdf4', text:'#16a34a', icon:'➕' },
        'EDIÇÃO':       { bg:'#eff6ff', text:'#1d4ed8', icon:'✏️' },
        'EXCLUSÃO':     { bg:'#fef2f2', text:'#dc2626', icon:'🗑️' },
        'CANCELAMENTO': { bg:'#fff8f0', text:'#d97706', icon:'⛔' },
    };

    if (!lista.length) {
        tbody.innerHTML = `
            <tr>
                <td colspan="5" style="text-align:center;color:#94a3b8;padding:40px">Nenhum registro de auditoria encontrado</td>
            </tr>`;
        const cnt = document.getElementById('auditoriaCount');
        if (cnt) cnt.textContent = '0 registros';
        return;
    }

    tbody.innerHTML = lista.map((a, i) => {
        const dt = a.quando ? new Date(a.quando).toLocaleString('pt-BR') : '—';
        const acaoUpper = (a.acao||'').toUpperCase();
        const colorKey = Object.keys(acaoColor).find(k => acaoUpper.includes(k));
        const ac = acaoColor[colorKey] || { bg:'#f8fafc', text:'#475569', icon:'📝' };

        let delta = a.detalhes || '';
        if (a.antes && a.depois) {
            const campos = [];
            ['valorTotal','representante','contrato','loja'].forEach(campo => {
                const antes  = a.antes[campo];
                const depois = a.depois[campo];
                if (antes !== undefined && depois !== undefined && antes !== depois) {
                    campos.push(`${campo}: ${antes} → ${depois}`);
                }
            });
            if (campos.length) delta = campos.join(' | ');
        }

        return `
            <tr style="background:${i%2===0?'#fff':'#f8fafc'};border-bottom:1px solid #f1f5f9">
                <td style="font-size:0.78rem;color:#475569;white-space:nowrap">${dt}</td>
                <td style="font-size:0.8rem;font-weight:500;color:#1e293b">${a.quem || 'Sistema'}</td>
                <td><span style="background:${ac.bg};color:${ac.text};font-size:0.72rem;font-weight:600;padding:2px 8px;border-radius:20px;white-space:nowrap">${ac.icon} ${a.acao || '—'}</span></td>
                <td style="font-weight:700;color:#1e3a5f">${a.contrato ? '#' + a.contrato : '—'}</td>
                <td style="font-size:0.78rem;color:#64748b;max-width:300px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${delta}">${delta || '—'}</td>
            </tr>`;
    }).join('');

    const cnt = document.getElementById('auditoriaCount');
    if (cnt) cnt.textContent = `${lista.length} de ${(estoque.auditoriaVendas||[]).length} registros`;
}

function exportarAuditoria() {
    const lista = (estoque.auditoriaVendas||[]).slice().reverse();
    if (!lista.length) { mostrarNotificacao('Nenhum registro para exportar.', 'warning'); return; }
    const rows = lista.map(a => ({
        'Data/Hora': a.quando ? new Date(a.quando).toLocaleString('pt-BR') : '',
        'Usuário':   a.quem   || 'Sistema',
        'Ação':      a.acao   || '',
        'Contrato':  a.contrato || '',
        'Detalhes':  a.detalhes || '',
    }));
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Auditoria');
    XLSX.writeFile(wb, 'auditoria_' + new Date().toISOString().split('T')[0] + '.xlsx');
}

// ========================================
// NAVEGAÇÃO POR ABAS
// ========================================

function trocarAba(aba) {
    try {
        const targetId = `tab-${aba}`;
        const target = document.getElementById(targetId);
        if (!target) {
            const msg = `Aba não encontrada: ${targetId}`;
            console.error(msg);
            if (window.__showRuntimeErrorOverlay) window.__showRuntimeErrorOverlay(msg);
            return;
        }

        abaAtiva = aba;
        try { window.abaAtiva = aba; } catch (e) {}

        // Atualizar botões de navegação (sidebar e fallback antigo)
        document.querySelectorAll('.sidebar .nav-item, .tabs-navigation .tab-btn').forEach(btn => {
            btn.classList.remove('active');
        });
        const activeSidebarBtn = document.querySelector(`.sidebar .nav-item[data-tab="${aba}"]`);
        const activeLegacyBtn = document.querySelector(`.tabs-navigation .tab-btn[data-tab="${aba}"]`);
        if (activeSidebarBtn) activeSidebarBtn.classList.add('active');
        if (activeLegacyBtn) activeLegacyBtn.classList.add('active');

        // Atualizar conteúdo de forma segura
        document.querySelectorAll('.tab-content').forEach(content => content.classList.remove('active'));
        target.classList.add('active');

        // Mostrar/ocultar barra de ações
        const acoesEstoque = document.getElementById('acoesEstoque');
        if (acoesEstoque) acoesEstoque.style.display = (aba === 'estoque') ? 'flex' : 'none';

        if (aba === 'dashboard') {
            try { renderizarDashboard(); } catch (e) { console.error('dashboard:', e); if (window.__showRuntimeErrorOverlay) window.__showRuntimeErrorOverlay(e); }
        } else if (aba === 'cadastro-produtos') {
            try { renderizarCadastroProdutos(); } catch (e) { console.error('cadastro-produtos:', e); if (window.__showRuntimeErrorOverlay) window.__showRuntimeErrorOverlay(e); }
        } else if (aba === 'vendas') {
            try { renderizarRegistroVendas(); } catch (e) { console.error('vendas:', e); if (window.__showRuntimeErrorOverlay) window.__showRuntimeErrorOverlay(e); }
        } else if (aba === 'distribuicao') {
            try { renderizarRegistroDistribuicao(); } catch (e) { if (window.__showRuntimeErrorOverlay) window.__showRuntimeErrorOverlay(e); }
            try { atualizarSelectDistribuicaoProduto(); } catch (e) { if (window.__showRuntimeErrorOverlay) window.__showRuntimeErrorOverlay(e); }
        } else if (aba === 'relatorios') {
            try { prepararRelatorioInventario(); } catch (e) { if (window.__showRuntimeErrorOverlay) window.__showRuntimeErrorOverlay(e); }
            try { gerarRelatorioRentabilidade(); } catch (e) { if (window.__showRuntimeErrorOverlay) window.__showRuntimeErrorOverlay(e); }
        } else if (aba === 'controleenvio') {
            try { renderizarControleEnvio(); } catch (e) { if (window.__showRuntimeErrorOverlay) window.__showRuntimeErrorOverlay(e); }
        } else if (aba === 'controleimbel') {
            try { trocarSubAbaControleImbel('estoque'); } catch (e) { if (window.__showRuntimeErrorOverlay) window.__showRuntimeErrorOverlay(e); }
        } else if (aba === 'precificacao') {
            try { trocarSubabaPrecif('federais'); } catch (e) { if (window.__showRuntimeErrorOverlay) window.__showRuntimeErrorOverlay(e); }
            try { renderizarPrecificacao(); } catch (e) { if (window.__showRuntimeErrorOverlay) window.__showRuntimeErrorOverlay(e); }
        } else if (aba === 'clientes') {
            try { renderizarClientes(); } catch (e) { if (window.__showRuntimeErrorOverlay) window.__showRuntimeErrorOverlay(e); }
        } else if (aba === 'propostas') {
            try { renderizarPropostas(); } catch (e) { if (window.__showRuntimeErrorOverlay) window.__showRuntimeErrorOverlay(e); }
        } else if (aba === 'estoque') {
            try { renderizarTabela(); } catch (e) { if (window.__showRuntimeErrorOverlay) window.__showRuntimeErrorOverlay(e); }
        } else if (aba === 'configuracoes') {
            try { atualizarPreviaContrato(); } catch (e) { if (window.__showRuntimeErrorOverlay) window.__showRuntimeErrorOverlay(e); }
            try { renderizarAuditoria(); } catch (e) { if (window.__showRuntimeErrorOverlay) window.__showRuntimeErrorOverlay(e); }
        }

        try { if (window.__updateDebugPanel) window.__updateDebugPanel(); } catch (e) {}
    } catch (err) {
        console.error('trocarAba falhou:', err);
        if (window.__showRuntimeErrorOverlay) window.__showRuntimeErrorOverlay(err);
    }
}

// ========================================
// RENDERIZAÇÃO DA TABELA DE ESTOQUE
// ========================================

function renderizarTabela() {
    const tbody = document.getElementById('corpoTabela');
    if (!tbody) return;
    tbody.innerHTML = '';
    const totais = { GERAL: { disp: 0, venda: 0, saldo: 0 } };
    estoque.representantes.forEach(rep => { totais[rep] = { disp: 0, venda: 0, saldo: 0 }; });

    // Ordem de exibição: usa a lista de representantes do objeto `estoque`
    const produtosOrdenados = [...estoque.produtos].sort((a, b) => a.nome.localeCompare(b.nome, 'pt-BR'));
    let countConsolidadoZeradoOuNegativo = 0;

    produtosOrdenados.forEach(produto => {
        const tr = document.createElement('tr');
        tr.dataset.id = produto.id;
        let produtoHtml = produto.nome;
        const metricaImbel = obterMetricasImbelProduto(produto);

        const produtoId = produto.id;
        const estoqueTotal = Number(
            (produto.estoqueConsolidado ?? produto.estoqueTotal ?? produto.estoque ?? produto.qtdTotal ?? produto.estoqueInicial) || 0
        ); // fonte de verdade (compatível com vários nomes)

        // === STEP 2: distribuições por representante (agregar do registro)
        const distPorRep = {};
        (estoque.registroDistribuicao || []).forEach(d => {
            try {
                if (Number(d.produtoId) === Number(produtoId) || (d.produtoNome && d.produtoNome === produto.nome)) {
                    const r = (d.representante || '').toString().toUpperCase();
                    distPorRep[r] = (distPorRep[r] || 0) + (Number(d.quantidade) || 0);
                }
            } catch (e) {}
        });
        const totalDistribuido = Object.values(distPorRep).reduce((s, v) => s + v, 0);

        // === STEP 3: devoluções por representante (origem da devolução)
        const devPorRep = {};
        let totalDevolvidoParaImbel = 0;
        (estoque.registroDevolucoes || []).forEach(d => {
            try {
                if (Number(d.produtoId) === Number(produtoId) || (d.produtoNome && d.produtoNome === produto.nome)) {
                    const origem = (d.origem || d.representante || '').toString().toUpperCase();
                    const destino = (d.destino || '').toString().toUpperCase();
                    devPorRep[origem] = (devPorRep[origem] || 0) + (Number(d.quantidade) || 0);
                    if (destino === 'IMBEL') totalDevolvidoParaImbel += (Number(d.quantidade) || 0);
                }
            } catch (e) {}
        });
        const totalDevolvido = Object.values(devPorRep).reduce((s, v) => s + v, 0);

        // === STEP 4: vendas por representante - preferir o agregado em produto.vendas
        const vendasPorRep = {};
        let agregadoTemValores = false;
        if (produto.vendas && typeof produto.vendas === 'object') {
            Object.keys(produto.vendas).forEach(k => {
                const val = Number(produto.vendas[k]) || 0;
                if (val !== 0) agregadoTemValores = true;
                vendasPorRep[(k||'').toString().toUpperCase()] = val;
            });
        }
        // Se o agregado estiver zerado (produto.vendas não confiável), popular a partir de registroVendas
        if (!agregadoTemValores) {
            (estoque.registroVendas || []).forEach(v => {
                try {
                    const rep = (v.representante || '').toString().toUpperCase();
                    if (Array.isArray(v.items) && v.items.length) {
                        v.items.forEach(it => {
                            if (Number(it.produtoId) === Number(produtoId) || (it.produto === produto.nome) || (it.produtoNome === produto.nome)) {
                                vendasPorRep[rep] = (vendasPorRep[rep] || 0) + (Number(it.quantidade) || 0);
                            }
                        });
                    } else {
                        if (Number(v.produtoId) === Number(produtoId) || (v.produtoNome === produto.nome)) {
                            vendasPorRep[rep] = (vendasPorRep[rep] || 0) + (Number(v.quantidade) || 0);
                        }
                    }
                } catch (e) {}
            });
        }

        // total vendas do produto (em todos os reps)
        const totalVendasProduto = Object.values(vendasPorRep).reduce((s, v) => s + v, 0);

        // === STEP 5: montar colunas por representante (KOLTE, ISA, LC, ADES, FL) e IMBEL
        estoque.representantes.forEach(rep => {
            const repKey = (rep || '').toString().toUpperCase();
            let disp = 0, venda = 0, saldo = 0;
            if (repKey === 'IMBEL') {
                // IMBEL centralizado para evitar divergência entre abas
                disp = metricaImbel.imbelDisp;
                venda = metricaImbel.imbelVenda;
                saldo = metricaImbel.imbelSaldo;
            } else {
                const dist = distPorRep[repKey] || 0;
                const dev = devPorRep[repKey] || 0;
                disp = dist - dev; // enviado ao rep menos o que ele devolveu
                venda = vendasPorRep[repKey] || 0;
                saldo = disp - venda;
            }

            // Determinar se há movimento para este rep (usado para mostrar '-')
            const repTeveMovimento = (Number(distPorRep[repKey] || 0) !== 0) || (Number(devPorRep[repKey] || 0) !== 0) || (Number(vendasPorRep[repKey] || 0) !== 0);

            // Formatação de células e classes conforme regras
            const dispText = (repKey === 'IMBEL' && estoqueTotal === 0) ? '-' : (repTeveMovimento || repKey === 'IMBEL' ? formatarNumero(disp) : '-');
            const vendaText = repTeveMovimento || (repKey === 'IMBEL' && (vendasPorRep['IMBEL']||0) > 0) ? formatarNumero(venda) : '-';
            let saldoText = '-';
            if (repKey === 'IMBEL' && estoqueTotal === 0) {
                saldoText = '-';
            } else if (!repTeveMovimento && repKey !== 'IMBEL') {
                saldoText = '-';
            } else {
                saldoText = formatarNumero(saldo);
            }

            const vendaClass = (Number(venda) > 0) ? 'cell-venda' : 'cell-zero';
            // saldo sempre usa 'cell-saldo' e adiciona modificadores
            let saldoClass = 'cell-saldo';
            if (Number(saldo) > 0) saldoClass += ' saldo-positivo';
            else if (Number(saldo) === 0) {
                const dispVal = Number(disp) || 0;
                saldoClass += (dispVal > 0) ? ' negativo' : ' saldo-zero';
            } else { // negativo
                saldoClass += ' negativo';
            }

            tr.innerHTML += `
                <td class="cell-disp numeric-cell ${(!isFinite(disp) || disp === 0) ? 'cell-zero' : ''}">${dispText}</td>
                <td class="cell-venda numeric-cell ${venda === 0 ? 'cell-zero' : ''} ${vendaClass}">${vendaText}</td>
                <td class="cell-saldo numeric-cell ${saldoClass} ${saldo === 0 ? 'cell-zero' : ''}">${saldoText}</td>
            `;

            // Atualizar totais por rep
            totais[repKey].disp += Number(disp) || 0;
            totais[repKey].venda += Number(venda) || 0;
            totais[repKey].saldo += Number(saldo) || 0;
        });

        // === STEP 7: CONSOLIDADO (sempre exibir número)
        const consolidadoDisp = metricaImbel.consolidadoDisp;
        const consolidadoVenda = metricaImbel.consolidadoVenda;
        const consolidadoSaldo = metricaImbel.consolidadoSaldo;

        tr.innerHTML = `<td class="produto-nome col-produto" title="${produto.nome}" onclick="abrirModalEditarProduto(${produtoId})" style="cursor:pointer">${produtoHtml}</td>` + tr.innerHTML;

        const saldoGeralClass = consolidadoSaldo > 0 ? 'saldo-positivo' : 'negativo';
        tr.innerHTML += `
            <td class="geral-disp numeric-cell">${formatarNumero(consolidadoDisp)}</td>
            <td class="geral-venda numeric-cell ${consolidadoVenda > 0 ? 'venda-positiva' : ''}">${formatarNumero(consolidadoVenda)}</td>
            <td class="geral-saldo numeric-cell ${saldoGeralClass}">${formatarNumero(consolidadoSaldo)}</td>
        `;

        // Flag produto com consolidado zerado ou negativo (para KPI)
        if (consolidadoSaldo <= 0) countConsolidadoZeradoOuNegativo += 1;

        // Atualizar totais gerais
        totais.GERAL.disp += consolidadoDisp;
        totais.GERAL.venda += consolidadoVenda;
        totais.GERAL.saldo += consolidadoSaldo;

        tbody.appendChild(tr);
    });

    // Linha de totais
    const trTotal = document.createElement('tr');
    trTotal.className = 'total-row';
    trTotal.innerHTML = `<td class="produto-nome col-produto"><strong>TOTAL GERAL</strong></td>`;

    estoque.representantes.forEach(rep => {
        const repKey = (rep || '').toString().toUpperCase();
        const saldoRep = totais[repKey].saldo;
        const saldoRepClass = saldoRep > 0 ? 'saldo-positivo' : 'negativo';
        trTotal.innerHTML += `
            <td class="cell-disp numeric-cell"><strong>${formatarNumero(totais[repKey].disp)}</strong></td>
            <td class="cell-venda numeric-cell ${totais[repKey].venda > 0 ? 'venda-positiva' : ''}"><strong>${formatarNumero(totais[repKey].venda)}</strong></td>
            <td class="cell-saldo numeric-cell ${saldoRepClass}"><strong>${formatarNumero(totais[repKey].saldo)}</strong></td>
        `;
    });

    const saldoGeralTotalClass = totais.GERAL.saldo > 0 ? 'saldo-positivo' : 'negativo';
    trTotal.innerHTML += `
        <td class="geral-disp numeric-cell"><strong>${formatarNumero(totais.GERAL.disp)}</strong></td>
        <td class="geral-venda numeric-cell"><strong>${formatarNumero(totais.GERAL.venda)}</strong></td>
        <td class="geral-saldo numeric-cell ${saldoGeralTotalClass}"><strong>${formatarNumero(totais.GERAL.saldo)}</strong></td>
    `;

    tbody.appendChild(trTotal);

    // KPIs da aba estoque (agora renderizados no Dashboard)
    const dashContainer = document.getElementById('tab-dashboard');
    const kpiTotalVendidoEl = (dashContainer && dashContainer.querySelector)
        ? (dashContainer.querySelector('#kpiTotalVendido') || document.getElementById('kpiTotalVendido'))
        : document.getElementById('kpiTotalVendido');
    const kpiSaldoZeroEl = (dashContainer && dashContainer.querySelector)
        ? (dashContainer.querySelector('#kpiSaldoZero') || document.getElementById('kpiSaldoZero'))
        : document.getElementById('kpiSaldoZero');
    // Calcular total vendido com fallback: prefer registroVendas, se vazio usar produto.vendas agregado
    let totalUnidadesVendidas = 0;
    if (Array.isArray(estoque.registroVendas) && estoque.registroVendas.length > 0) {
        totalUnidadesVendidas = estoque.registroVendas.reduce((sum, v) => {
            if (Array.isArray(v.items)) return sum + v.items.reduce((s, i) => s + (Number(i.quantidade) || 0), 0);
            return sum + (Number(v.quantidade) || 0);
        }, 0);
    } else {
        totalUnidadesVendidas = (estoque.produtos || []).reduce((sumP, p) => {
            if (p.vendas && typeof p.vendas === 'object') return sumP + Object.values(p.vendas).reduce((s, vv) => s + (Number(vv) || 0), 0);
            return sumP;
        }, 0);
    }
    if (kpiTotalVendidoEl) kpiTotalVendidoEl.textContent = `${totalUnidadesVendidas.toLocaleString('pt-BR')} un.`;
    if (kpiSaldoZeroEl) kpiSaldoZeroEl.textContent = `${countConsolidadoZeradoOuNegativo.toLocaleString('pt-BR')} produtos`;

    // Ajustar posição sticky da segunda linha do header
    ajustarStickyHeader();
    try { renderizarCadastroProdutos(); } catch (e) {}
}

function renderizarCadastroProdutos() {
    const tbody = document.getElementById('corpoCadastroProdutos');
    const empty = document.getElementById('cadastroProdutosEmpty');
    const tabelaWrap = document.getElementById('cadastroProdutosTabelaWrap');
    if (!tbody || !empty || !tabelaWrap) return;

    const produtos = Array.isArray(estoque.produtos) ? [...estoque.produtos] : [];
    produtos.sort((a, b) => String(a.nome || '').localeCompare(String(b.nome || ''), 'pt-BR'));

    const total = produtos.length;
    let comCI = 0;

    tbody.innerHTML = '';

    produtos.forEach(produto => {
        const nome = produto.nome || '-';
        const regra = (precificacao && precificacao[nome]) ? precificacao[nome] : {};
        const ci = Number(regra.ci ?? produto.ci ?? 0) || 0;
        const margemMin = Number(regra.margemMinima ?? produto.margemMinima ?? 0) || 0;
        const descontoMax = Number(regra.descontoMaximo ?? produto.descontoMaximo ?? 0) || 0;
        const categoria = (categoriaPorProduto && categoriaPorProduto[nome]) || produto.categoria || '-';
        const ncm = produto.ncm || '-';
        const metricaImbel = obterMetricasImbelProduto(produto);
        const imbelTexto = metricaImbel.estoqueTotal === 0
            ? '-'
            : formatarNumero(metricaImbel.imbelDisp);
        // Mesmo valor do bloco CONSOLIDADO > SALDO da aba Estoque
        const saldoConsolidado = metricaImbel.consolidadoSaldo;
        const saldoConsolidadoTexto = formatarNumero(saldoConsolidado);
        const saldoConsolidadoCor = saldoConsolidado > 0 ? '#2da44e' : '#cf222e';
        const saldoConsolidadoClasse = saldoConsolidado < 0 ? 'negativo' : '';
        const atualizadoEm = produto.atualizadoEm || produto.dataAtualizacao || produto.criadoEm || '';
        const atualizadoTxt = atualizadoEm ? formatarData(atualizadoEm) : '-';

        if (ci > 0) comCI += 1;

        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td style="text-align:left; font-weight:600">${nome}</td>
            <td>${ncm}</td>
            <td>${categoria}</td>
            <td>${ci > 0 ? formatarMoedaValor(ci) : '-'}</td>
            <td>${margemMin > 0 ? margemMin.toFixed(1).replace('.', ',') : '-'}</td>
            <td>${descontoMax > 0 ? descontoMax.toFixed(1).replace('.', ',') : '-'}</td>
            <td>${imbelTexto}</td>
            <td class="${saldoConsolidadoClasse}" style="color:${saldoConsolidadoCor}; font-weight:700; font-family:monospace;">${saldoConsolidadoTexto}</td>
            <td>${atualizadoTxt}</td>
            <td>
                <button class="btn btn-outline btn-sm" data-admin="true" onclick="abrirModalEditarProduto(${Number(produto.id)})">Editar</button>
            </td>
        `;

        // Garantia defensiva: tbody deve ter o mesmo número de colunas do thead
        const expectedCols = document.querySelectorAll('#tabelaCadastroProdutos thead th').length;
        let actualCols = tr.querySelectorAll('td').length;
        if (actualCols !== expectedCols) {
            while (actualCols < expectedCols) {
                tr.appendChild(document.createElement('td'));
                actualCols += 1;
            }
            while (actualCols > expectedCols) {
                tr.removeChild(tr.lastElementChild);
                actualCols -= 1;
            }
        }

        tbody.appendChild(tr);
    });

    const semCI = Math.max(0, total - comCI);
    const totalEl = document.getElementById('cadProdTotal');
    const comCIEl = document.getElementById('cadProdComCI');
    const semCIEl = document.getElementById('cadProdSemCI');
    if (totalEl) totalEl.textContent = String(total);
    if (comCIEl) comCIEl.textContent = String(comCI);
    if (semCIEl) semCIEl.textContent = String(semCI);

    const vazio = total === 0;
    empty.style.display = vazio ? 'block' : 'none';
    tabelaWrap.style.display = vazio ? 'none' : 'block';
}

// Calcula e aplica o top correto para a segunda linha do thead (sub-headers)
function ajustarStickyHeader() {
    const tabela = document.getElementById('tabelaEstoque');
    if (!tabela) return;
    const firstRow = tabela.querySelector('thead tr:first-child');
    if (!firstRow) return;
    requestAnimationFrame(() => {
        const h = firstRow.getBoundingClientRect().height;
        const secondRowThs = tabela.querySelectorAll('thead tr:nth-child(2) th');
        secondRowThs.forEach(th => { th.style.top = h + 'px'; });
    });
}

// ========================================
// RENDERIZAÇÃO DO DASHBOARD
// ========================================

// ========================================
// RELATÓRIOS / IMPRESSÃO
// ========================================

function prepararRelatorioInventario() {
    // Only render if the relatorios tab is currently active
    const tabEl = document.getElementById('tab-relatorios');
    if (!tabEl || !tabEl.classList.contains('active')) return;
    const preview = document.getElementById('relatoriosPreview');
    if (!preview) return;

    const tabela = document.getElementById('tabelaEstoque');
    if (!tabela) {
        preview.innerHTML = '<p>Tabela de estoque não encontrada.</p>';
        return;
    }
    // Atualizar selects de relatorio (caso a lista de produtos tenha mudado)
    atualizarSelectsRelatorios();

    const filtroRep = document.getElementById('filtroRelatoriosRep').value;
    const filtroProduto = document.getElementById('filtroRelatoriosProduto').value;

    // Clonar a tabela para preview/print, evitando ids duplicados
    const clone = tabela.cloneNode(true);
    clone.id = 'tabelaEstoqueRelatorio';

    // Remover possíveis estilos de posicionamento que atrapalham impressão
    clone.querySelectorAll('thead th').forEach(th => { th.style.position = 'static'; th.style.left = 'auto'; th.style.top = 'auto'; });
    clone.querySelectorAll('td').forEach(td => { td.style.position = 'static'; td.style.left = 'auto'; });

    // Filtrar por produto (se selecionado) e por representante (linha com valores)
    const corpo = clone.querySelector('tbody');
    if (corpo) {
        const rows = Array.from(corpo.querySelectorAll('tr'));
        rows.forEach(row => {
            const pid = row.dataset.id;
            // Se é a linha de totais sem dataset, manter
            if (!pid) return;

            // Filtrar por produto
            if (filtroProduto && filtroProduto !== '' && pid !== filtroProduto) {
                row.remove();
                return;
            }

            // Filtrar por representante: manter apenas se houver quantidade ou venda para esse rep
            if (filtroRep && filtroRep !== '') {
                const produto = estoque.produtos.find(p => String(p.id) === String(pid));
                if (produto) {
                    const disp = produto.distribuicao[filtroRep] || 0;
                    const venda = produto.vendas[filtroRep] || 0;
                    if ((disp + venda) === 0) {
                        row.remove();
                        return;
                    }
                }
            }
        });
    }

    // Se foi selecionado um representante, reconstruir o THEAD reduzido e remover colunas que não pertencem
    const selRep = filtroRep || '';
    if (selRep) {
        const repsCount = estoque.representantes.length;
        const repIndex = estoque.representantes.indexOf(selRep);
        // índices baseados na ordem dos tds no tbody: 0 = produto, then reps*3, then GERAL*3
        const produtoColIndex = 0;
        const repStart = 1 + (repIndex * 3);
        const repCols = [repStart, repStart + 1, repStart + 2];
        const geralStart = 1 + (repsCount * 3);
        const geralCols = [geralStart, geralStart + 1, geralStart + 2];

        // Reconstruir THEAD com apenas PRODUTOS | REP (colspan=3) | CONSOLIDADO (colspan=3)
        const newThead = document.createElement('thead');
        const first = document.createElement('tr');
        const thProdutos = document.createElement('th');
        thProdutos.className = 'col-produto';
        thProdutos.rowSpan = 2;
        thProdutos.textContent = 'PRODUTOS';
        first.appendChild(thProdutos);

        const thRep = document.createElement('th');
        thRep.className = 'header-rep ' + selRep.toLowerCase();
        thRep.colSpan = 3;
        thRep.textContent = selRep;
        first.appendChild(thRep);

        const thGeral = document.createElement('th');
        thGeral.className = 'header-geral';
        thGeral.colSpan = 3;
        thGeral.textContent = 'CONSOLIDADO';
        first.appendChild(thGeral);

        const second = document.createElement('tr');
        const subLabels = ['DIST', 'VENDA', 'SALDO', 'TOTAL', 'VENDA', 'SALDO'];
        const ariaMap = {
            'DIST': 'Quantidade distribuída para este representante',
            'ESTOQUE': 'Quantidade disponível no galpão IMBEL',
            'TOTAL': 'Estoque total cadastrado no sistema',
            'VENDA': 'Total vendido',
            'SALDO': 'Saldo disponível para venda'
        };
        subLabels.forEach(text => {
            const t = document.createElement('th');
            t.className = 'sub-header';
            t.textContent = text;
            const aria = ariaMap[text] || '';
            if (aria) {
                t.setAttribute('aria-label', aria);
                t.setAttribute('title', aria);
            }
            second.appendChild(t);
        });

        newThead.appendChild(first);
        newThead.appendChild(second);

        // Substituir thead do clone
        const oldThead = clone.querySelector('thead');
        if (oldThead) oldThead.remove();
        clone.insertBefore(newThead, clone.firstChild);

        // Agora remover das linhas do corpo as colunas que não estão em repCols ou geralCols
        if (corpo) {
            const rowsAll = Array.from(corpo.querySelectorAll('tr'));
            rowsAll.forEach(row => {
                const cells = Array.from(row.children);
                // construir lista de índices a manter
                const keep = new Set([produtoColIndex, ...repCols, ...geralCols]);
                // iterar de trás para frente ao remover
                for (let i = cells.length - 1; i >= 0; i--) {
                    if (!keep.has(i)) {
                        cells[i].remove();
                    }
                }
            });
        }
    }

    preview.innerHTML = '';
    const wrapper = document.createElement('div');
    wrapper.className = 'report-printable';
    wrapper.appendChild(clone);
    preview.appendChild(wrapper);
}

function imprimirInventario() {
    prepararRelatorioInventario();
    const preview = document.getElementById('relatoriosPreview');
    if (!preview) return;
    const content = preview.innerHTML;
    const filtroRep = document.getElementById('filtroRelatoriosRep') ? document.getElementById('filtroRelatoriosRep').value : '';
    const filtroProduto = document.getElementById('filtroRelatoriosProduto') ? document.getElementById('filtroRelatoriosProduto').value : '';
    const orient = document.getElementById('filtroRelatoriosOrientacao') ? document.getElementById('filtroRelatoriosOrientacao').value : 'landscape';

    // Montar cabeçalho resumido para o relatório
    const produtoNome = filtroProduto ? (estoque.produtos.find(p => String(p.id) === String(filtroProduto)) || {}).nome : '';
    const dataAgora = new Date().toLocaleString('pt-BR');
    const headerHTML = `<div style="margin-bottom:8px;font-size:13px;color:#222"><strong>Representante:</strong> ${filtroRep || 'Todos'} &nbsp;|&nbsp; <strong>Produto:</strong> ${produtoNome || 'Todos'} &nbsp;|&nbsp; <strong>Data:</strong> ${dataAgora}</div>`;
    const win = window.open('', '_blank', 'width=1000,height=700');
    if (!win) {
        alert('Não foi possível abrir a janela de impressão. Permita popups ou use a impressão do navegador.');
        return;
    }

    win.document.write(`
        <!doctype html>
        <html lang="pt-BR">
        <head>
            <meta charset="utf-8">
            <title>Relatório - Inventário</title>
            <link rel="stylesheet" href="styles.css">
            <style>
                @page { size: A4 ${orient}; margin: 10mm; }
                body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Arial, sans-serif; padding: 12px; color: #222; }
                h1 { margin-bottom: 12px; font-size: 18px; }
                .report-printable table { width: 100%; border-collapse: collapse; font-size: 12px; }
                .report-printable th, .report-printable td { border: 1px solid #ddd; padding: 6px 8px; text-align: left; }
                thead { background: #1e3a5f; color: white; }
                th.col-produto, td.produto-nome { position: static !important; left: auto !important; }
                @media print { body { margin: 0; } }
            </style>
        </head>
        <body>
            <h1>Inventário de Produtos</h1>
            ${headerHTML}
            ${content}
            <script>window.onload = function(){ setTimeout(function(){ window.print(); }, 200); };</script>
        </body>
        </html>
    `);
    win.document.close();
}

function aprovarProposta(id) {
    const p = (propostas || []).find(x => x.id === id);
    if (!p) return;
    p.aguardandoAprovacao = false;
    try {
        p.aprovadoPor = (typeof auth !== 'undefined' && auth.currentUser && auth.currentUser.email) ? auth.currentUser.email : 'Admin';
    } catch (e) { p.aprovadoPor = 'Admin'; }
    p.dataAprovacao = new Date().toISOString();
    p.status = 'rascunho';
    salvarDados();
    try { renderizarPropostas(); } catch (e) {}
    mostrarNotificacao(`✅ Proposta ${p.numero} aprovada.`, 'success');
}

function recusarAprovacaoProposta(id) {
    const p = (propostas || []).find(x => x.id === id);
    if (!p) return;
    const motivo = prompt('Motivo da recusa:') || 'Sem motivo';
    p.aguardandoAprovacao = false;
    p.status = 'rascunho';
    p.observacoes = (p.observacoes || '') + `\n[RECUSADO: ${motivo} — ${new Date().toLocaleDateString('pt-BR')}]`;
    salvarDados();
    try { renderizarPropostas(); } catch (e) {}
    mostrarNotificacao('Desconto recusado. Proposta voltou para rascunho.', 'warning');
}
 

// ==============================
// RELATÓRIO DE RENTABILIDADE
// ==============================
function gerarRelatorioRentabilidade() {
    // Only render if the relatorios tab is currently active
    const tabEl = document.getElementById('tab-relatorios');
    if (!tabEl || !tabEl.classList.contains('active')) return;
    try {
        const periodo = document.getElementById('filtroRentabilidadePeriodo')?.value || 'todos';
        const hoje = new Date();
        const vendas = Array.isArray(estoque.registroVendas) ? estoque.registroVendas : (estoque.registroVendas || []);

        const vendasFiltradas = (vendas || []).filter(v => {
            if (!v || !v.data) return periodo === 'todos';
            const dataVenda = new Date(v.data);
            if (isNaN(dataVenda)) return periodo === 'todos';
            if (periodo === 'todos') return true;
            if (periodo === 'mes') return dataVenda.getMonth() === hoje.getMonth() && dataVenda.getFullYear() === hoje.getFullYear();
            if (periodo === 'trimestre') {
                const trimAtual = Math.floor(hoje.getMonth() / 3);
                return Math.floor(dataVenda.getMonth() / 3) === trimAtual && dataVenda.getFullYear() === hoje.getFullYear();
            }
            if (periodo === 'ano') return dataVenda.getFullYear() === hoje.getFullYear();
            return true;
        });

        // 2. Agregar por produto
        const porProduto = {};
        vendasFiltradas.forEach(venda => {
            const itens = Array.isArray(venda.items) && venda.items.length ? venda.items : (venda.itemsLegacy ? venda.itemsLegacy : []);
            if (!Array.isArray(itens) || itens.length === 0) {
                // Suporta formato legado com campos diretos
                const nome = venda.produtoNome || venda.produto || '';
                const qtd = Number(venda.quantidade) || 0;
                const valorUnit = Number(venda.valorUnitario || venda.valorUnit || (venda.valorTotal && qtd ? (venda.valorTotal / qtd) : 0)) || 0;
                if (!porProduto[nome]) porProduto[nome] = { qtd: 0, receita: 0 };
                porProduto[nome].qtd += qtd;
                porProduto[nome].receita += qtd * valorUnit;
                return;
            }
            itens.forEach(item => {
                let nome = item.produtoNome || item.produto || '';
                if (!nome && item.produtoId) {
                    const p = (estoque.produtos || []).find(pp => String(pp.id) === String(item.produtoId));
                    nome = p ? p.nome : String(item.produtoId);
                }
                const qtd = Number(item.quantidade) || 0;
                const valorUnit = Number(item.valorUnitario || item.valorUnit || (item.valorTotal && qtd ? (item.valorTotal / qtd) : 0)) || 0;
                if (!porProduto[nome]) porProduto[nome] = { qtd: 0, receita: 0 };
                porProduto[nome].qtd += qtd;
                porProduto[nome].receita += qtd * valorUnit;
            });
        });

        // 3. Enriquecer com custo da `precificacao`
        let totalReceita = 0, totalCusto = 0;
        const rows = Object.entries(porProduto).map(([nome, dados]) => {
            const prec = (typeof precificacao === 'object' && precificacao) ? (precificacao[nome] || {}) : {};
            const custoUnit = Number(prec.custoTotal || prec.custo_unitario || prec.custo || 0) || 0;
            const custoTotalProd = custoUnit * dados.qtd;
            const lucro = dados.receita - custoTotalProd;
            const margem = dados.receita > 0 ? (lucro / dados.receita) * 100 : 0;
            totalReceita += dados.receita;
            totalCusto += custoTotalProd;
            return { nome, qtd: dados.qtd, receita: dados.receita, custoUnit, custoTotal: custoTotalProd, lucro, margem };
        }).sort((a, b) => b.lucro - a.lucro);

        // 4. Montar linhas da tabela por produto
        const corMargem = m => m >= 20 ? '#16a34a' : m >= 10 ? '#d97706' : '#dc2626';
        const emojiRank = i => ['🥇','🥈','🥉'][i] || (i + 1);
        const tbodyProd = document.getElementById('tabelaRentabilidadeBody');
        if (tbodyProd) {
            tbodyProd.innerHTML = rows.map((r, i) => `
                <tr>
                  <td style="text-align:left; padding-left:15px; font-weight:500">${r.nome}</td>
                  <td>${r.qtd}</td>
                  <td style="color:#16a34a; font-weight:600">R$ ${r.receita.toLocaleString('pt-BR',{minimumFractionDigits:2})}</td>
                  <td>${r.custoUnit > 0 ? 'R$ ' + r.custoUnit.toLocaleString('pt-BR',{minimumFractionDigits:2}) : '<span style="color:#94a3b8">Sem custo</span>'}</td>
                  <td>${r.custoTotal > 0 ? 'R$ ' + r.custoTotal.toLocaleString('pt-BR',{minimumFractionDigits:2}) : '—'}</td>
                  <td style="font-weight:600; color:${r.lucro >= 0 ? '#16a34a' : '#dc2626'}">R$ ${r.lucro.toLocaleString('pt-BR',{minimumFractionDigits:2})}</td>
                  <td style="font-weight:700; color:${corMargem(r.margem)}">${r.margem.toFixed(1)}%</td>
                  <td style="text-align:center; font-size:1.1rem">${emojiRank(i)}</td>
                </tr>
            `).join('');
        }

        // 5. Agregar por representante
        const porRep = {};
        (vendasFiltradas || []).forEach(v => {
            const rep = (v.representante || 'IMBEL').toString().toUpperCase();
            if (!porRep[rep]) porRep[rep] = { qtd: 0, receita: 0, custo: 0 };
            const itens = Array.isArray(v.items) && v.items.length ? v.items : (v.itemsLegacy ? v.itemsLegacy : []);
            if (!Array.isArray(itens) || itens.length === 0) {
                const nome = v.produtoNome || v.produto || '';
                const qtd = Number(v.quantidade) || 0;
                const valorUnit = Number(v.valorUnitario || v.valorUnit || (v.valorTotal && qtd ? (v.valorTotal / qtd) : 0)) || 0;
                const prec = (precificacao && precificacao[nome]) ? precificacao[nome] : {};
                porRep[rep].qtd += qtd;
                porRep[rep].receita += qtd * valorUnit;
                porRep[rep].custo += (Number(prec.custoTotal) || 0) * qtd;
            } else {
                itens.forEach(item => {
                    let nome = item.produtoNome || item.produto || '';
                    if (!nome && item.produtoId) {
                        const p = (estoque.produtos || []).find(pp => String(pp.id) === String(item.produtoId));
                        nome = p ? p.nome : String(item.produtoId);
                    }
                    const qtd = Number(item.quantidade) || 0;
                    const valorUnit = Number(item.valorUnitario || item.valorUnit || (item.valorTotal && qtd ? (item.valorTotal / qtd) : 0)) || 0;
                    const prec = (precificacao && precificacao[nome]) ? precificacao[nome] : {};
                    porRep[rep].qtd += qtd;
                    porRep[rep].receita += qtd * valorUnit;
                    porRep[rep].custo += (Number(prec.custoTotal) || 0) * qtd;
                });
            }
        });

        const repColors = { KOLTE:'#3d5a80', ISA:'#5c4d7d', LC:'#2d6a4f', ADES:'#9c4a1a', FL:'#7b2d26', IMBEL:'#1e3a5f' };
        const tbodyRep = document.getElementById('tabelaRentabilidadeRepBody');
        if (tbodyRep) {
            tbodyRep.innerHTML = Object.entries(porRep).map(([rep, d]) => {
                const lucro = d.receita - d.custo;
                const margem = d.receita > 0 ? (lucro / d.receita) * 100 : 0;
                const cor = repColors[rep] || '#1e3a5f';
                return `
                    <tr>
                      <td><span class="badge-rep" style="background:${cor}20; color:${cor}; font-weight:700; padding:3px 10px; border-radius:20px">${rep}</span></td>
                      <td>${d.qtd}</td>
                      <td style="color:#16a34a; font-weight:600">R$ ${d.receita.toLocaleString('pt-BR',{minimumFractionDigits:2})}</td>
                      <td>${d.custo > 0 ? 'R$ ' + d.custo.toLocaleString('pt-BR',{minimumFractionDigits:2}) : '—'}</td>
                      <td style="font-weight:600; color:${lucro>=0?'#16a34a':'#dc2626'}">R$ ${lucro.toLocaleString('pt-BR',{minimumFractionDigits:2})}</td>
                      <td style="font-weight:700; color:${corMargem(margem)}">${margem.toFixed(1)}%</td>
                    </tr>
                `;
            }).join('');
        }

        // 6. Atualizar KPIs
        const lucroTotal = totalReceita - totalCusto;
        const margemTotal = totalReceita > 0 ? (lucroTotal / totalReceita) * 100 : 0;
        const fmt = v => 'R$ ' + v.toLocaleString('pt-BR',{minimumFractionDigits:2});
        if (document.getElementById('rentFaturamento')) document.getElementById('rentFaturamento').textContent = fmt(totalReceita);
        if (document.getElementById('rentCusto')) document.getElementById('rentCusto').textContent = fmt(totalCusto);
        if (document.getElementById('rentLucro')) { document.getElementById('rentLucro').textContent = fmt(lucroTotal); document.getElementById('rentLucro').style.color = lucroTotal >= 0 ? '#16a34a' : '#dc2626'; }
        if (document.getElementById('rentMargem')) document.getElementById('rentMargem').textContent = margemTotal.toFixed(1) + '%';
    } catch (e) {
        console.error('Erro gerando relatório de rentabilidade:', e);
    }
}

function exportarRentabilidadeExcel() {
    try {
        const rows = [];
        const body = document.getElementById('tabelaRentabilidadeBody');
        if (body) {
            Array.from(body.querySelectorAll('tr')).forEach(tr => {
                const cols = Array.from(tr.children).map(td => td.textContent.trim().replace(/\s+/g, ' '));
                rows.push(cols.join(';'));
            });
        }
        let csv = 'Produto;Qtd Vendida;Receita Total;Custo Unit.;Custo Total;Lucro Bruto;Margem %;Ranking\n';
        csv += rows.join('\n');
        const blob = new Blob(['\ufeff' + csv], { type: 'text/csv;charset=utf-8;' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `rentabilidade_${new Date().toISOString().slice(0,10)}.csv`;
        document.body.appendChild(a);
        a.click();
        a.remove();
        URL.revokeObjectURL(url);
    } catch (e) {
        console.error('Erro exportando rentabilidade:', e);
        mostrarNotificacao('Erro ao exportar relatório.', 'error');
    }
}

function imprimirRentabilidade() {
    try {
        gerarRelatorioRentabilidade();
        const prodHtml = document.getElementById('tabelaRentabilidade') ? document.getElementById('tabelaRentabilidade').outerHTML : '';
        const repHtml = document.getElementById('tabelaRentabilidadeRep') ? document.getElementById('tabelaRentabilidadeRep').outerHTML : '';
        const dataAgora = new Date().toLocaleString('pt-BR');
        const win = window.open('', '_blank', 'width=1000,height=800');
        if (!win) { alert('Permita popups para imprimir.'); return; }
        win.document.write(`
            <!doctype html><html lang="pt-BR"><head><meta charset="utf-8"><title>Relatório de Rentabilidade</title>
            <link rel="stylesheet" href="styles.css"><style>body{font-family:Arial,Helvetica,sans-serif;padding:12px;color:#222}table{width:100%;border-collapse:collapse}th,td{border:1px solid #ddd;padding:6px}</style></head><body>
            <h1>Relatório de Rentabilidade</h1>
            <div style="margin-bottom:8px;font-size:13px;color:#222"><strong>Gerado em:</strong> ${dataAgora}</div>
            ${prodHtml}
            <div style="height:18px"></div>
            ${repHtml}
            <script>window.onload=function(){setTimeout(()=>window.print(),200);};</script>
            </body></html>
        `);
        win.document.close();
    } catch (e) {
        console.error('Erro imprimindo rentabilidade:', e);
    }
}

function gerarPdfProposta(propostaId, tipo = 'simples') {
    try {
        const proposta = (propostas || []).find(p => String(p.id) === String(propostaId));
        if (!proposta) { mostrarNotificacao('Proposta não encontrada.', 'error'); return; }

        if (!window.jspdf || !window.jspdf.jsPDF) {
            mostrarNotificacao('jsPDF não está disponível.', 'error');
            return;
        }
        const { jsPDF } = window.jspdf;
        const doc = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });

        // HEADER
        doc.setFillColor(30, 58, 95);
        doc.rect(0, 0, 210, 35, 'F');
        doc.setFont('helvetica', 'bold');
        doc.setFontSize(16);
        doc.setTextColor(255, 255, 255);
        doc.text('FÁBRICA DE ITAJUBÁ', 15, 15);
        doc.setFontSize(9);
        doc.setFont('helvetica', 'normal');
        doc.text('Controle de Estoque — Material Bélico', 15, 22);
        doc.setFont('helvetica', 'bold');
        doc.setFontSize(14);
        doc.setTextColor(201, 162, 39);
        doc.text('PROPOSTA ' + (proposta.numero || ''), 195, 15, { align: 'right' });
        doc.setFontSize(9);
        doc.setTextColor(200, 200, 200);
        doc.text('Data: ' + (proposta.data ? new Date(proposta.data).toLocaleDateString('pt-BR') : ''), 195, 22, { align: 'right' });
        doc.text('Validade: ' + (proposta.dataExpiracao ? new Date(proposta.dataExpiracao).toLocaleDateString('pt-BR') : ''), 195, 28, { align: 'right' });

        // CLIENT INFO
        let y = 45;
        doc.setFillColor(248, 250, 252);
        doc.rect(10, y - 5, 190, 28, 'F');
        doc.setDrawColor(226, 232, 240);
        doc.rect(10, y - 5, 190, 28, 'S');
        doc.setFont('helvetica', 'bold');
        doc.setFontSize(8);
        doc.setTextColor(100, 116, 139);
        doc.text('CLIENTE', 15, y);
        doc.setFont('helvetica', 'bold');
        doc.setFontSize(12);
        doc.setTextColor(30, 41, 59);
        doc.text(proposta.cliente || '-', 15, y + 7);

        const clienteObj = (clientes || []).find(c => (c.nome || '').toLowerCase() === (proposta.cliente || '').toLowerCase());
        if (clienteObj) {
            doc.setFont('helvetica', 'normal');
            doc.setFontSize(9);
            doc.setTextColor(71, 85, 105);
            if (clienteObj.cnpj) doc.text('CNPJ: ' + clienteObj.cnpj, 15, y + 13);
            if (clienteObj.contato) doc.text('Contato: ' + clienteObj.contato, 100, y + 13);
            if (clienteObj.telefone) doc.text('Tel: ' + clienteObj.telefone, 15, y + 19);
            if (clienteObj.email) doc.text('E-mail: ' + clienteObj.email, 100, y + 19);
        }

        doc.setFont('helvetica', 'bold');
        doc.setFontSize(8);
        doc.setTextColor(100, 116, 139);
        doc.text('REPRESENTANTE', 160, y);
        doc.setFontSize(10);
        doc.setTextColor(30, 41, 59);
        doc.text(proposta.representante || '-', 160, y + 7);

        // ITENS
        y += 35;
        let finalY = y;
        const cliente = clienteObj;

        if (tipo === 'fiscal') {
            doc.setFont('helvetica', 'bold');
            doc.setFontSize(10);
            doc.setTextColor(30, 58, 95);
            doc.text('ITENS E COMPOSIÇÃO DE PREÇO', 15, y);
            y += 6;

            const precifRef = (precificacoesCliente || [])
                .filter(pc => String(pc.clienteId) === String(cliente?.id))
                .sort((a, b) => new Date(b.dataCriacao) - new Date(a.dataCriacao))[0];

            (proposta.itens || []).forEach((item, idx) => {
                const itemPrecif = precifRef?.itens?.find(i => i.produto === item.produto);
                const fmt = v => 'R$ ' + parseFloat(v || 0).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
                const pct = v => parseFloat(v || 0).toFixed(2) + '%';
                const qtd = Number(item.quantidade || 0);
                const valorUnit = Number(item.valorUnitario || item.valor || item.valorUnit || 0);

                if (y > 260 && idx < (proposta.itens || []).length - 1) {
                    doc.addPage();
                    y = 20;
                }

                doc.setFillColor(30, 58, 95);
                doc.rect(10, y, 190, 7, 'F');
                doc.setFont('helvetica', 'bold');
                doc.setFontSize(9);
                doc.setTextColor(255, 255, 255);
                doc.text(`${idx + 1}. ${item.produto || item.produtoNome || ''}`, 14, y + 5);
                doc.text(`Qtd: ${qtd}`, 160, y + 5);
                y += 10;

                if (itemPrecif) {
                    doc.autoTable({
                        startY: y,
                        body: [
                            ['Custo de Industrialização (CI)', '—', fmt(itemPrecif.ci)],
                            ['Taxa Diretoria', pct(itemPrecif.taxa), '×' + (1 + itemPrecif.taxa / 100).toFixed(4)],
                            ['ROI', pct(itemPrecif.roi), '×' + (1 + itemPrecif.roi / 100).toFixed(4)],
                            ['Valor Base', '—', fmt(itemPrecif.valorBase)],
                            ['ICMS', pct(itemPrecif.icms), fmt(itemPrecif.icmsR)],
                            ['PIS', pct(itemPrecif.pis), fmt(itemPrecif.pisR)],
                            ['COFINS', pct(itemPrecif.cofins), fmt(itemPrecif.cofinsR)],
                            ['c/ ICMS + PIS + COFINS', '—', fmt(itemPrecif.valorImpostos)],
                            ['IPI', pct(itemPrecif.ipi), fmt(itemPrecif.ipiR)],
                            ['Valor Total s/ Comissão', '—', fmt(itemPrecif.valorTotal)],
                            ['Comissão (s/ Valor Base)', pct(itemPrecif.comissao), fmt(itemPrecif.comissaoR)],
                            ['PREÇO UNITÁRIO FINAL', '—', fmt(valorUnit)]
                        ],
                        styles: { fontSize: 8, cellPadding: 2.5 },
                        columnStyles: {
                            0: { cellWidth: 110, textColor: [71, 85, 105] },
                            1: { cellWidth: 25, halign: 'center', textColor: [100, 116, 139] },
                            2: { cellWidth: 45, halign: 'right', fontStyle: 'bold', textColor: [30, 41, 59] }
                        },
                        didParseCell(data) {
                            if (data.row.index === 11) {
                                data.cell.styles.fillColor = [30, 58, 95];
                                data.cell.styles.textColor = [201, 162, 39];
                                data.cell.styles.fontStyle = 'bold';
                            }
                            if ([3, 7, 9].includes(data.row.index)) {
                                data.cell.styles.fillColor = [240, 244, 248];
                                data.cell.styles.fontStyle = 'bold';
                            }
                        },
                        margin: { left: 10, right: 10 }
                    });
                    y = doc.lastAutoTable.finalY + 3;
                    doc.setFont('helvetica', 'bold');
                    doc.setFontSize(8);
                    doc.setTextColor(22, 163, 74);
                    doc.text(`Total: ${qtd} × ${fmt(valorUnit)} = ` + fmt(qtd * valorUnit), 195, y, { align: 'right' });
                    y += 8;
                } else {
                    doc.autoTable({
                        startY: y,
                        body: [
                            ['Produto', item.produto || item.produtoNome || ''],
                            ['Qtd', qtd],
                            ['Valor Unitário', fmt(valorUnit)],
                            ['Total', fmt(qtd * valorUnit)]
                        ],
                        styles: { fontSize: 8 },
                        margin: { left: 10, right: 10 }
                    });
                    y = doc.lastAutoTable.finalY + 4;
                }
            });
            finalY = y + 2;
        } else {
            doc.setFont('helvetica', 'bold');
            doc.setFontSize(10);
            doc.setTextColor(30, 58, 95);
            doc.text('ITENS DA PROPOSTA', 15, y);
            y += 5;

            const tableData = (proposta.itens || []).map((item, idx) => {
                const nome = item.produtoNome || item.produto || '';
                const quantidade = Number(item.quantidade || 0);
                const valorUnit = Number(item.valorUnitario || item.valor || item.valorUnit || 0);
                const total = quantidade * valorUnit;
                return [
                    idx + 1,
                    nome,
                    quantidade,
                    'R$ ' + valorUnit.toLocaleString('pt-BR', { minimumFractionDigits: 2 }),
                    'R$ ' + total.toLocaleString('pt-BR', { minimumFractionDigits: 2 })
                ];
            });

            doc.autoTable({
                startY: y,
                head: [['#', 'Produto', 'Qtd', 'Valor Unit.', 'Total']],
                body: tableData,
                styles: { fontSize: 9, cellPadding: 4 },
                headStyles: { fillColor: [30, 58, 95], textColor: [255, 255, 255], fontStyle: 'bold' },
                columnStyles: { 0: { cellWidth: 12, halign: 'center' }, 1: { cellWidth: 90 }, 2: { cellWidth: 20, halign: 'center' }, 3: { cellWidth: 35, halign: 'right' }, 4: { cellWidth: 35, halign: 'right' } },
                alternateRowStyles: { fillColor: [248, 250, 252] },
                margin: { left: 10, right: 10 }
            });

            finalY = (doc.lastAutoTable && doc.lastAutoTable.finalY) ? doc.lastAutoTable.finalY + 8 : y + 60;
        }

        doc.setFillColor(30, 58, 95);
        doc.rect(130, finalY, 70, 14, 'F');
        doc.setFont('helvetica', 'bold');
        doc.setFontSize(9);
        doc.setTextColor(201, 162, 39);
        doc.text('VALOR TOTAL', 135, finalY + 5);
        doc.setFontSize(12);
        doc.setTextColor(255, 255, 255);
        doc.text('R$ ' + Number(proposta.valorTotal || 0).toLocaleString('pt-BR', { minimumFractionDigits: 2 }), 195, finalY + 10, { align: 'right' });

        // OBS
        if (proposta.observacoes) {
            const obsY = finalY + 22;
            doc.setFillColor(255, 251, 240);
            doc.setDrawColor(201, 162, 39);
            doc.rect(10, obsY - 4, 190, 20, 'FD');
            doc.setFont('helvetica', 'bold');
            doc.setFontSize(8);
            doc.setTextColor(100, 116, 139);
            doc.text('CONDIÇÕES COMERCIAIS', 15, obsY + 1);
            doc.setFont('helvetica', 'normal');
            doc.setFontSize(9);
            doc.setTextColor(30, 41, 59);
            const lines = doc.splitTextToSize(proposta.observacoes, 180);
            doc.text(lines, 15, obsY + 7);
        }

        // FOOTER
        const pageHeight = doc.internal.pageSize.height;
        doc.setFillColor(30, 58, 95);
        doc.rect(0, pageHeight - 15, 210, 15, 'F');
        doc.setFont('helvetica', 'normal');
        doc.setFontSize(8);
        doc.setTextColor(200, 200, 200);
        doc.text('Fábrica de Itajubá — Gestão de Material Bélico', 15, pageHeight - 6);
        doc.text('Proposta válida até ' + (proposta.dataExpiracao ? new Date(proposta.dataExpiracao).toLocaleDateString('pt-BR') : ''), 195, pageHeight - 6, { align: 'right' });

        const safeName = (proposta.cliente || 'cliente').replace(/[^a-z0-9]/gi, '_');
        const sufixo = tipo === 'fiscal' ? '_fiscal' : '';
        doc.save('Proposta_' + (proposta.numero || '') + '_' + safeName + sufixo + '.pdf');

    } catch (err) {
        console.error('Erro gerarPdfProposta:', err);
        mostrarNotificacao('Erro ao gerar PDF da proposta.', 'error');
    }
}

// Imprimir exatamente o preview atual em `#relatoriosPreview`
function imprimirRelatorioPreview() {
    const preview = document.getElementById('relatoriosPreview');
    if (!preview) { mostrarNotificacao('Preview de relatórios não encontrado.', 'error'); return; }
    // Se não há conteúdo, tentar visualizar antes de imprimir
    if (!preview.innerHTML || preview.innerHTML.trim() === '') {
        // Tenta re-gerar o preview atual com a função existente
        try { visualizarRelatorioSelecionado(); } catch (e) {}
        if (!preview.innerHTML || preview.innerHTML.trim() === '') {
            mostrarNotificacao('Nenhum relatório visualizado para imprimir.', 'warning');
            return;
        }
    }

    const content = preview.innerHTML;
    const tipo = document.getElementById('filtroRelatoriosTipo')?.value || '';
    const orient = document.getElementById('filtroRelatoriosOrientacao')?.value || 'landscape';
    const titulo = tipo === 'comissoes' ? 'Relatório - Comissões' : tipo === 'distribuicao' ? 'Relatório - Distribuição' : 'Relatório - Inventário';
    const dataAgora = new Date().toLocaleString('pt-BR');

    const win = window.open('', '_blank', 'width=1000,height=700');
    if (!win) { alert('Não foi possível abrir a janela de impressão. Permita popups.'); return; }

    win.document.write(`
        <!doctype html>
        <html lang="pt-BR">
        <head>
            <meta charset="utf-8">
            <title>${titulo}</title>
            <link rel="stylesheet" href="styles.css">
            <style>
                @page { size: A4 ${orient}; margin: 10mm; }
                body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Arial, sans-serif; padding:12px; color:#222 }
                h1 { margin-bottom:8px }
                .report-printable table { width:100%; border-collapse:collapse }
                .report-printable th, .report-printable td { border:1px solid #ddd; padding:6px 8px }
            </style>
        </head>
        <body>
            <h1>${titulo}</h1>
            <div style="margin-bottom:8px;font-size:13px;color:#222"><strong>Data:</strong> ${dataAgora}</div>
            <div class="report-printable">${content}</div>
            <script>window.onload = function(){ setTimeout(function(){ window.print(); }, 200); };</script>
        </body>
        </html>
    `);
    win.document.close();
}

// Imprimir diretamente o relatório selecionado no filtro (sem depender do preview)
function imprimirRelatorioSelecionado() {
    // Garantir que o preview esteja atualizado e então imprimir o conteúdo do preview
    try {
        visualizarRelatorioSelecionado();
    } catch (e) {
        // fallback: tentar preparar manualmente por tipo
        const tipo = document.getElementById('filtroRelatoriosTipo') ? document.getElementById('filtroRelatoriosTipo').value : 'inventario';
        if (tipo === 'comissoes') prepararRelatorioComissoes();
        else if (tipo === 'distribuicao') prepararRelatorioDistribuicao && prepararRelatorioDistribuicao();
        else prepararRelatorioInventario();
    }

    // Pequeno delay para garantir que o DOM do preview foi atualizado
    setTimeout(function(){ imprimirRelatorioPreview(); }, 80);
}

// =============================
// RELATÓRIO: COMISSÕES (5%)
// =============================

function obterComissoesConsolidadas({ filtroRep = '', dataInicio = '', dataFim = '' } = {}) {
    const vendas = (Array.isArray(estoque.registroVendas) ? [...estoque.registroVendas] : []).filter(v => !v.cancelado);
    const vendasSemImbel = vendas.filter(v => ((v.representante || '').toString().trim().toUpperCase() !== 'IMBEL'));

    const vendasFiltradas = vendasSemImbel.filter(v => {
        if (filtroRep && (v.representante || '') !== filtroRep) return false;
        if ((!dataInicio || dataInicio === '') && (!dataFim || dataFim === '')) return true;
        const d = parseDateToYYYYMMDD(v.data);
        if (!d) return false;
        if (dataInicio && d < dataInicio) return false;
        if (dataFim && d > dataFim) return false;
        return true;
    });

    const obterValorVenda = (venda) => {
        if (typeof venda.valorTotal === 'number') return venda.valorTotal;
        if (Array.isArray(venda.items) && venda.items.length > 0) {
            return venda.items.reduce((s, it) => s + (Number(it.valorTotal) || ((Number(it.valorUnitario) || 0) * (Number(it.quantidade) || 0))), 0);
        }
        return ((Number(venda.valorUnitario) || 0) * (Number(venda.quantidade) || 0));
    };

    const ordenarContrato = (a, b) => {
        const na = parseInt(a);
        const nb = parseInt(b);
        if (!isNaN(na) && !isNaN(nb)) return na - nb;
        return a.localeCompare(b);
    };

    const contratosMap = new Map();
    vendasFiltradas.forEach(v => {
        const contratoKey = normalizarContratoKey(v.contrato);
        if (!contratoKey) return;
        const dataNorm = parseDateToYYYYMMDD(v.data);
        const atual = contratosMap.get(contratoKey) || {
            contrato: contratoKey,
            loja: v.loja || '',
            representantes: new Set(),
            valorContrato: 0,
            dataMin: null,
            dataMax: null
        };
        atual.valorContrato += obterValorVenda(v);
        if (!atual.loja && v.loja) atual.loja = v.loja;
        if (v.representante) atual.representantes.add(v.representante);
        if (dataNorm) {
            if (!atual.dataMin || dataNorm < atual.dataMin) atual.dataMin = dataNorm;
            if (!atual.dataMax || dataNorm > atual.dataMax) atual.dataMax = dataNorm;
        }
        contratosMap.set(contratoKey, atual);
    });

    const contratos = Array.from(contratosMap.values()).sort((a, b) => ordenarContrato(a.contrato, b.contrato));
    let totalComissoes = 0;
    contratos.forEach(c => {
        c.comissao = Math.round((c.valorContrato * 0.05) * 100) / 100;
        totalComissoes += c.comissao;
    });

    return { contratos, totalComissoes };
}

function prepararRelatorioComissoes() {
    const preview = document.getElementById('relatoriosPreview');
    if (!preview) return;

    const filtroRep = document.getElementById('filtroRelatoriosRep') ? document.getElementById('filtroRelatoriosRep').value : '';
    const dataInicio = document.getElementById('filtroRelatoriosDataInicio') ? document.getElementById('filtroRelatoriosDataInicio').value : '';
    const dataFim = document.getElementById('filtroRelatoriosDataFim') ? document.getElementById('filtroRelatoriosDataFim').value : '';

    // Agrupar vendas por representante (ignorar vendas da IMBEL — sem comissão)
    const vendas = (Array.isArray(estoque.registroVendas) ? [...estoque.registroVendas] : []).filter(v => !v.cancelado);
    const vendasSemImbel = vendas.filter(v => ((v.representante || '').toString().trim().toUpperCase() !== 'IMBEL'));
    
    // Filtrar por intervalo de datas se fornecido (comparação por DATA apenas, formato YYYY-MM-DD)
    const vendasFiltradas = vendasSemImbel.filter(v => {
        if ((!dataInicio || dataInicio === '') && (!dataFim || dataFim === '')) return true;
        if (!v.data) return false;
        // Normalizar data do registro para YYYY-MM-DD
        const registroDateStr = parseDateToYYYYMMDD(v.data);
        if (!registroDateStr) return false;
        if (dataInicio && dataInicio !== '' && registroDateStr < dataInicio) return false;
        if (dataFim && dataFim !== '' && registroDateStr > dataFim) return false;
        return true;
    });

    // (debug panels removed)

    // Ordenar por contrato
    vendasFiltradas.sort((a, b) => (parseInt(a.contrato) || 0) - (parseInt(b.contrato) || 0));

    let totalComissoes = 0;

    const container = document.createElement('div');
    container.className = 'report-comissoes';
    // Card resumo
    const resumo = document.createElement('div');
    resumo.className = 'comissoes-resumo';
    resumo.style.marginBottom = '12px';
    resumo.innerHTML = `<strong>Total Comissões:</strong> <span id="totalComissoesCard">R$ 0,00</span>`;
    container.appendChild(resumo);

    const obterValorVenda = (venda) => {
        if (typeof venda.valorTotal === 'number') return venda.valorTotal;
        if (Array.isArray(venda.items) && venda.items.length > 0) {
            return venda.items.reduce((s, it) => s + (Number(it.valorTotal) || ((Number(it.valorUnitario) || 0) * (Number(it.quantidade) || 0))), 0);
        }
        return ((Number(venda.valorUnitario) || 0) * (Number(venda.quantidade) || 0));
    };

    const normalizarContrato = (valor) => {
        const bruto = (valor ?? '').toString().normalize('NFKC').replace(/[\u200B-\u200D\uFEFF\s]+/g, '');
        const digitos = bruto.replace(/\D+/g, '');
        return digitos ? String(parseInt(digitos, 10)) : bruto.toUpperCase();
    };

    const ordenarContrato = (a, b) => {
        const na = parseInt(a);
        const nb = parseInt(b);
        if (!isNaN(na) && !isNaN(nb)) return na - nb;
        return a.localeCompare(b);
    };

    const vendasConsideradas = filtroRep
        ? vendasFiltradas.filter(v => (v.representante || '') === filtroRep)
        : vendasFiltradas;

    const contratosMap = new Map();
    vendasConsideradas.forEach(v => {
        const contratoKey = normalizarContrato(v.contrato);
        if (!contratoKey) return;
        const dataNorm = parseDateToYYYYMMDD(v.data);
        const atual = contratosMap.get(contratoKey) || {
            contrato: contratoKey,
            loja: v.loja || '',
            representantes: new Set(),
            valorContrato: 0,
            dataMin: null,
            dataMax: null
        };
        atual.valorContrato += obterValorVenda(v);
        if (!atual.loja && v.loja) atual.loja = v.loja;
        if (v.representante) atual.representantes.add(v.representante);
        if (dataNorm) {
            if (!atual.dataMin || dataNorm < atual.dataMin) atual.dataMin = dataNorm;
            if (!atual.dataMax || dataNorm > atual.dataMax) atual.dataMax = dataNorm;
        }
        contratosMap.set(contratoKey, atual);
    });

    const contratos = Array.from(contratosMap.values()).sort((a, b) => ordenarContrato(a.contrato, b.contrato));

    const table = document.createElement('table');
    table.className = 'tabela-relatorio comissoes-table';
    table.style.width = '100%';
    table.style.borderCollapse = 'collapse';
    table.innerHTML = `
        <thead>
            <tr>
                <th style="text-align:left;padding:6px;border:1px solid #ddd">Contrato</th>
                <th style="text-align:left;padding:6px;border:1px solid #ddd">Cliente / Loja</th>
                <th style="text-align:left;padding:6px;border:1px solid #ddd">Representante(s)</th>
                <th style="text-align:left;padding:6px;border:1px solid #ddd">Data</th>
                <th style="text-align:right;padding:6px;border:1px solid #ddd">Valor Contrato (R$)</th>
                <th style="text-align:right;padding:6px;border:1px solid #ddd">Comissão 5% (R$)</th>
            </tr>
        </thead>
        <tbody></tbody>
    `;

    const tbody = table.querySelector('tbody');
    contratos.forEach(c => {
        const valor = c.valorContrato || 0;
        const comissao = Math.round((valor * 0.05) * 100) / 100;
        totalComissoes += comissao;
        const repsTexto = Array.from(c.representantes || []).join(', ');
        const dataTexto = c.dataMin
            ? (c.dataMax && c.dataMax !== c.dataMin
                ? `${new Date(c.dataMin + 'T00:00:00').toLocaleDateString('pt-BR')} até ${new Date(c.dataMax + 'T00:00:00').toLocaleDateString('pt-BR')}`
                : new Date(c.dataMin + 'T00:00:00').toLocaleDateString('pt-BR'))
            : '-';

        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td style="padding:6px;border:1px solid #ddd">${c.contrato || ''}</td>
            <td style="padding:6px;border:1px solid #ddd">${c.loja || ''}</td>
            <td style="padding:6px;border:1px solid #ddd">${repsTexto || '-'}</td>
            <td style="padding:6px;border:1px solid #ddd">${dataTexto}</td>
            <td style="padding:6px;border:1px solid #ddd;text-align:right">${formatarMoedaValor(valor)}</td>
            <td style="padding:6px;border:1px solid #ddd;text-align:right">${formatarMoedaValor(comissao)}</td>
        `;
        tbody.appendChild(tr);
    });

    const trTotal = document.createElement('tr');
    trTotal.innerHTML = `
        <td colspan="5" style="padding:6px;border:1px solid #ddd;text-align:right"><strong>Total Geral de Comissões</strong></td>
        <td style="padding:6px;border:1px solid #ddd;text-align:right"><strong>${formatarMoedaValor(totalComissoes)}</strong></td>
    `;
    tbody.appendChild(trTotal);

    container.appendChild(table);

    // Atualizar card total
    const totalEl = container.querySelector('#totalComissoesCard');
    if (totalEl) totalEl.textContent = formatarMoedaValor(totalComissoes);

    // Renderizar no preview (substitui o conteúdo atual de relatoriosPreview)
    preview.innerHTML = '';
    const wrapper = document.createElement('div');
    wrapper.className = 'report-printable';
    wrapper.appendChild(container);
    preview.appendChild(wrapper);
}

function imprimirComissoes() {
    prepararRelatorioComissoes();
    const preview = document.getElementById('relatoriosPreview');
    if (!preview) return;
    const content = preview.innerHTML;
    const filtroRep = document.getElementById('filtroRelatoriosRep') ? document.getElementById('filtroRelatoriosRep').value : 'Todos';
    const dataInicio = document.getElementById('filtroRelatoriosDataInicio') ? document.getElementById('filtroRelatoriosDataInicio').value : '';
    const dataFim = document.getElementById('filtroRelatoriosDataFim') ? document.getElementById('filtroRelatoriosDataFim').value : '';
    const dataAgora = (dataInicio || dataFim) ? `${dataInicio || '-'} até ${dataFim || '-'}` : new Date().toLocaleString('pt-BR');

    const win = window.open('', '_blank', 'width=900,height=700');
    if (!win) { alert('Não foi possível abrir janela de impressão. Permita popups.'); return; }

    win.document.write(`
        <!doctype html>
        <html lang="pt-BR">
        <head>
            <meta charset="utf-8">
            <title>Relatório - Comissões</title>
            <link rel="stylesheet" href="styles.css">
            <style>
                @page { size: A4 portrait; margin: 10mm; }
                body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Arial, sans-serif; padding: 12px; color: #222; }
                h1 { margin-bottom: 12px; font-size: 16px; }
                table { width: 100%; border-collapse: collapse; font-size: 12px; }
                th, td { border:1px solid #ddd; padding:6px 8px; }
                thead { background:#1e3a5f; color:white; }
                .comissoes-resumo { margin-bottom:12px; font-size:14px; }
            </style>
        </head>
        <body>
            <h1>Relatório de Comissões (5%)</h1>
            <div style="margin-bottom:8px;font-size:13px;color:#222"><strong>Representante:</strong> ${filtroRep || 'Todos'} &nbsp;|&nbsp; <strong>Data:</strong> ${dataAgora}</div>
            ${content}
            <script>window.onload = function(){ setTimeout(function(){ window.print(); }, 200); };</script>
        </body>
        </html>
    `);
    win.document.close();
}

function exportarComissoesCSV() {
    const filtroRep = document.getElementById('filtroRelatoriosRep') ? document.getElementById('filtroRelatoriosRep').value : '';
    // Excluir vendas da IMBEL (sem comissão) e aplicar filtro de datas se fornecido
    const dataInicio = document.getElementById('filtroRelatoriosDataInicio') ? document.getElementById('filtroRelatoriosDataInicio').value : '';
    const dataFim = document.getElementById('filtroRelatoriosDataFim') ? document.getElementById('filtroRelatoriosDataFim').value : '';
    const vendasRaw = (Array.isArray(estoque.registroVendas) ? [...estoque.registroVendas] : []).filter(v => !v.cancelado);
    const vendasFiltradasPorRep = vendasRaw.filter(v => ((v.representante || '').toString().trim().toUpperCase() !== 'IMBEL'));
    let startTs = null, endTs = null;
    try {
        if (dataInicio) startTs = new Date(dataInicio + 'T00:00:00').getTime();
        if (dataFim) endTs = new Date(dataFim + 'T23:59:59').getTime();
    } catch (e) { startTs = null; endTs = null; }
    const vendas = vendasFiltradasPorRep.filter(v => {
        if (!startTs && !endTs) return true;
        if (!v.data) return false;
        const t = new Date(v.data).getTime();
        if (startTs && t < startTs) return false;
        if (endTs && t > endTs) return false;
        return true;
    });

    const obterValorVenda = (venda) => {
        if (typeof venda.valorTotal === 'number') return venda.valorTotal;
        if (Array.isArray(venda.items) && venda.items.length > 0) {
            return venda.items.reduce((s, it) => s + (Number(it.valorTotal) || ((Number(it.valorUnitario) || 0) * (Number(it.quantidade) || 0))), 0);
        }
        return ((Number(venda.valorUnitario) || 0) * (Number(venda.quantidade) || 0));
    };

    const normalizarContrato = (valor) => {
        const bruto = (valor ?? '').toString().normalize('NFKC').replace(/[\u200B-\u200D\uFEFF\s]+/g, '');
        const digitos = bruto.replace(/\D+/g, '');
        return digitos ? String(parseInt(digitos, 10)) : bruto.toUpperCase();
    };

    const ordenarContrato = (a, b) => {
        const na = parseInt(a);
        const nb = parseInt(b);
        if (!isNaN(na) && !isNaN(nb)) return na - nb;
        return a.localeCompare(b);
    };

    const contratosMap = new Map();
    vendas.forEach(v => {
        if (filtroRep && filtroRep !== '' && v.representante !== filtroRep) return;
        const contratoKey = normalizarContrato(v.contrato);
        if (!contratoKey) return;
        const mapKey = `${v.representante || ''}||${contratoKey}`;
        const dataNorm = parseDateToYYYYMMDD(v.data);
        const atual = contratosMap.get(mapKey) || {
            representante: v.representante || '',
            contrato: contratoKey,
            loja: v.loja || '',
            valorContrato: 0,
            dataMin: null,
            dataMax: null
        };
        atual.valorContrato += obterValorVenda(v);
        if (!atual.loja && v.loja) atual.loja = v.loja;
        if (dataNorm) {
            if (!atual.dataMin || dataNorm < atual.dataMin) atual.dataMin = dataNorm;
            if (!atual.dataMax || dataNorm > atual.dataMax) atual.dataMax = dataNorm;
        }
        contratosMap.set(mapKey, atual);
    });

    const contratos = Array.from(contratosMap.values()).sort((a, b) => {
        const repCmp = (a.representante || '').localeCompare(b.representante || '');
        if (repCmp !== 0) return repCmp;
        return ordenarContrato(a.contrato, b.contrato);
    });

    const sep = ';';
    let csv = `REPRESENTANTE${sep}CONTRATO${sep}CLIENTE/LOJA${sep}DATA${sep}VALOR_CONTRATO${sep}COMISSAO_5%\n`;

    contratos.forEach(c => {
        const valor = c.valorContrato || 0;
        const comissao = Math.round((valor * 0.05) * 100) / 100;
        const dataTexto = c.dataMin
            ? (c.dataMax && c.dataMax !== c.dataMin
                ? `${new Date(c.dataMin + 'T00:00:00').toLocaleDateString('pt-BR')} até ${new Date(c.dataMax + 'T00:00:00').toLocaleDateString('pt-BR')}`
                : new Date(c.dataMin + 'T00:00:00').toLocaleDateString('pt-BR'))
            : '-';
        csv += `${c.representante || ''}${sep}${c.contrato || ''}${sep}${(c.loja || '').replace(/\n/g,' ')}${sep}${dataTexto}${sep}${valor.toFixed(2).replace('.',',')}${sep}${comissao.toFixed(2).replace('.',',')}\n`;
    });

    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `comissoes_${new Date().toISOString().slice(0,10)}.csv`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
}

function gerarFechamentoMensalComissoes() {
    const hoje = new Date();
    const competenciaDefault = `${hoje.getFullYear()}-${String(hoje.getMonth() + 1).padStart(2, '0')}`;
    const competencia = (prompt('Informe a competência (AAAA-MM):', competenciaDefault) || '').trim();
    if (!/^\d{4}-\d{2}$/.test(competencia)) {
        mostrarNotificacao('Competência inválida. Use AAAA-MM.', 'error');
        return;
    }

    const [ano, mes] = competencia.split('-').map(n => parseInt(n, 10));
    const dataInicio = `${ano}-${String(mes).padStart(2, '0')}-01`;
    const ultimoDia = new Date(ano, mes, 0).getDate();
    const dataFim = `${ano}-${String(mes).padStart(2, '0')}-${String(ultimoDia).padStart(2, '0')}`;
    const filtroRep = document.getElementById('filtroRelatoriosRep')?.value || '';

    const { contratos, totalComissoes } = obterComissoesConsolidadas({ filtroRep, dataInicio, dataFim });
    const chave = `${competencia}||${filtroRep || 'TODOS'}`;
    if (!Array.isArray(estoque.fechamentosComissoes)) estoque.fechamentosComissoes = [];
    const idxExistente = estoque.fechamentosComissoes.findIndex(f => f.chave === chave);
    if (idxExistente !== -1) {
        const ok = confirm(`Já existe fechamento para ${competencia} (${filtroRep || 'Todos'}). Deseja substituir?`);
        if (!ok) return;
    }

    const tabela = `
        <table class="dashboard-table">
            <thead>
                <tr>
                    <th>Referência</th>
                    <th>Lâmina</th>
                    <th>Matéria</th>
                    <th>Qtd</th>
                    <th>Preço</th>
                    <th>Local</th>
                </tr>
            </thead>
            <tbody>
                ${linhas}
            </tbody>
            <tfoot>
                <tr>
                    <td colspan="5">Total no Estoque</td>
                    <td>${totalGeral}</td>
                </tr>
            </tfoot>
        </table>
    `;
    abrirFechamentosComissoes();
}

function abrirFechamentosComissoes() {
    const container = document.getElementById('fechamentosComissoesConteudo');
    const modal = document.getElementById('modalFechamentosComissoes');
    if (!container || !modal) return;

    const lista = Array.isArray(estoque.fechamentosComissoes) ? [...estoque.fechamentosComissoes] : [];
    lista.sort((a, b) => (a.competencia < b.competencia ? 1 : -1));

    if (lista.length === 0) {
        container.innerHTML = '<p style="text-align:center;color:var(--text-secondary);padding:20px">Nenhum fechamento mensal registrado.</p>';
    } else {
        container.innerHTML = lista.map(f => {
            const criado = f.criadoEm ? new Date(f.criadoEm).toLocaleString('pt-BR') : '-';
            return `
                <div class="historico-item">
                    <span class="hist-data"><strong>${f.competencia}</strong><br><small>${criado}</small></span>
                    <span class="hist-tipo venda">SNAPSHOT</span>
                    <span class="hist-descricao">
                        ${f.filtroRep ? `Rep: ${f.filtroRep} | ` : ''}${f.linhas.length} contrato(s) | Total comissão: ${formatarMoedaValor(f.totalComissoes || 0)}
                        <div style="margin-top:6px">
                            <button class="btn btn-outline btn-sm" onclick="visualizarFechamentoComissoes(${f.id})">Visualizar</button>
                            <button class="btn btn-outline btn-sm" onclick="excluirFechamentoComissoes(${f.id})">Excluir</button>
                        </div>
                    </span>
                </div>
            `;
        }).join('');
    }

    modal.style.display = 'flex';
}

function visualizarFechamentoComissoes(id) {
    const fechamento = (estoque.fechamentosComissoes || []).find(f => f.id === id);
    if (!fechamento) return;
    const preview = document.getElementById('relatoriosPreview');
    if (!preview) return;

    const container = document.createElement('div');
    container.className = 'report-comissoes';
    container.innerHTML = `<div class="comissoes-resumo" style="margin-bottom:12px"><strong>Fechamento:</strong> ${fechamento.competencia} ${fechamento.filtroRep ? `| Rep: ${fechamento.filtroRep}` : ''} | <strong>Total:</strong> ${formatarMoedaValor(fechamento.totalComissoes || 0)}</div>`;

    const table = document.createElement('table');
    table.className = 'tabela-relatorio comissoes-table';
    table.style.width = '100%';
    table.style.borderCollapse = 'collapse';
    table.innerHTML = `
        <thead><tr>
            <th style="text-align:left;padding:6px;border:1px solid #ddd">Contrato</th>
            <th style="text-align:left;padding:6px;border:1px solid #ddd">Cliente / Loja</th>
            <th style="text-align:left;padding:6px;border:1px solid #ddd">Representante(s)</th>
            <th style="text-align:left;padding:6px;border:1px solid #ddd">Data</th>
            <th style="text-align:right;padding:6px;border:1px solid #ddd">Valor Contrato (R$)</th>
            <th style="text-align:right;padding:6px;border:1px solid #ddd">Comissão 5% (R$)</th>
        </tr></thead>
        <tbody></tbody>
    `;
    const tbody = table.querySelector('tbody');
    (fechamento.linhas || []).forEach(l => {
        const dataTexto = l.dataMin
            ? (l.dataMax && l.dataMax !== l.dataMin
                ? `${new Date(l.dataMin + 'T00:00:00').toLocaleDateString('pt-BR')} até ${new Date(l.dataMax + 'T00:00:00').toLocaleDateString('pt-BR')}`
                : new Date(l.dataMin + 'T00:00:00').toLocaleDateString('pt-BR'))
            : '-';
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td style="padding:6px;border:1px solid #ddd">${l.contrato || ''}</td>
            <td style="padding:6px;border:1px solid #ddd">${l.loja || ''}</td>
            <td style="padding:6px;border:1px solid #ddd">${(l.representantes || []).join(', ') || '-'}</td>
            <td style="padding:6px;border:1px solid #ddd">${dataTexto}</td>
            <td style="padding:6px;border:1px solid #ddd;text-align:right">${formatarMoedaValor(l.valorContrato || 0)}</td>
            <td style="padding:6px;border:1px solid #ddd;text-align:right">${formatarMoedaValor(l.comissao || 0)}</td>
        `;
        tbody.appendChild(tr);
    });

    container.appendChild(table);
    preview.innerHTML = '';
    const wrapper = document.createElement('div');
    wrapper.className = 'report-printable';
    wrapper.appendChild(container);
    preview.appendChild(wrapper);
    trocarAba('relatorios');
    fecharModal('modalFechamentosComissoes');
}

function excluirFechamentoComissoes(id) {
    const fechamento = (estoque.fechamentosComissoes || []).find(f => f.id === id);
    if (!fechamento) return;
    if (!confirm(`Excluir fechamento ${fechamento.competencia}?`)) return;
    estoque.fechamentosComissoes = (estoque.fechamentosComissoes || []).filter(f => f.id !== id);
    salvarDados();
    abrirFechamentosComissoes();
}

function exportarExcelCompleto() {
    if (typeof XLSX === 'undefined') {
        mostrarNotificacao('Biblioteca XLSX não carregada.', 'error');
        return;
    }

    const wb = XLSX.utils.book_new();

    const invRows = estoque.produtos.map(p => {
        const row = { Produto: p.nome };
        let totalDisp = 0, totalVenda = 0;
        (estoque.representantes || []).forEach(rep => {
            const disp = p.distribuicao?.[rep] || 0;
            const venda = p.vendas?.[rep] || 0;
            row[`${rep}_Disp`] = disp;
            row[`${rep}_Venda`] = venda;
            row[`${rep}_Saldo`] = disp - venda;
            totalDisp += disp;
            totalVenda += venda;
        });
        row.Total_Disp = totalDisp;
        row.Total_Venda = totalVenda;
        row.Total_Saldo = totalDisp - totalVenda;
        return row;
    });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(invRows), 'Inventario');

    const vendasRows = [];
    (estoque.registroVendas || []).forEach(v => {
        const data = parseDateToYYYYMMDD(v.data) || '';
        if (Array.isArray(v.items) && v.items.length > 0) {
            v.items.forEach(it => vendasRows.push({
                Contrato: v.contrato,
                Cliente: v.loja,
                Representante: v.representante,
                Produto: it.produtoNome,
                Quantidade: it.quantidade || 0,
                Valor_Unitario: it.valorUnitario || 0,
                Valor_Total_Item: it.valorTotal || 0,
                Data: data,
                Observacoes: v.observacoes || ''
            }));
        } else {
            vendasRows.push({
                Contrato: v.contrato,
                Cliente: v.loja,
                Representante: v.representante,
                Produto: v.produtoNome || '',
                Quantidade: v.quantidade || 0,
                Valor_Unitario: v.valorUnitario || 0,
                Valor_Total_Item: v.valorTotal || 0,
                Data: data,
                Observacoes: v.observacoes || ''
            });
        }
    });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(vendasRows), 'Vendas');

    const distRows = (estoque.registroDistribuicao || []).map(d => ({
        Representante: d.representante,
        Produto: d.produtoNome,
        Quantidade: d.quantidade || 0,
        Data: parseDateToYYYYMMDD(d.data) || d.data || '',
        Observacoes: d.observacoes || ''
    }));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(distRows), 'Distribuicao');

    // Incluir devoluções em aba separada
    const devolRows = (estoque.registroDevolucoes || []).map(d => ({
        Origem: d.origem,
        Destino: d.destino,
        Produto: d.produtoNome,
        Quantidade: d.quantidade || 0,
        Data: d.data || '',
        Observacoes: d.observacoes || ''
    }));
    if (devolRows.length > 0) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(devolRows), 'Devolucoes');

    const filtroRep = document.getElementById('filtroRelatoriosRep')?.value || '';
    const dataInicio = document.getElementById('filtroRelatoriosDataInicio')?.value || '';
    const dataFim = document.getElementById('filtroRelatoriosDataFim')?.value || '';
    const { contratos } = obterComissoesConsolidadas({ filtroRep, dataInicio, dataFim });
    const comRows = contratos.map(c => {
        const dataTexto = c.dataMin
            ? (c.dataMax && c.dataMax !== c.dataMin
                ? `${new Date(c.dataMin + 'T00:00:00').toLocaleDateString('pt-BR')} até ${new Date(c.dataMax + 'T00:00:00').toLocaleDateString('pt-BR')}`
                : new Date(c.dataMin + 'T00:00:00').toLocaleDateString('pt-BR'))
            : '-';
        return {
            Contrato: c.contrato,
            Cliente: c.loja,
            Representantes: Array.from(c.representantes || []).join(', '),
            Data: dataTexto,
            Valor_Contrato: c.valorContrato || 0,
            Comissao_5: c.comissao || 0
        };
    });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(comRows), 'Comissoes');

    XLSX.writeFile(wb, `controle_estoque_${new Date().toISOString().slice(0,10)}.xlsx`);
    mostrarNotificacao('Excel exportado com sucesso!', 'success');
}

function visualizarRelatorioSelecionado() {
    const tipo = document.getElementById('filtroRelatoriosTipo') ? document.getElementById('filtroRelatoriosTipo').value : 'inventario';
    if (tipo === 'comissoes') {
        prepararRelatorioComissoes();
    } else {
        prepararRelatorioInventario();
    }
}

function imprimirTabelaSeparada(tableId, titulo, subtitulo = '', orientacao = 'landscape') {
    const tabela = document.getElementById(tableId);
    if (!tabela) {
        mostrarNotificacao('Tabela não encontrada para impressão.', 'error');
        return;
    }

    const clone = tabela.cloneNode(true);

    // Converter campos editáveis em texto para impressão limpa
    clone.querySelectorAll('td').forEach(td => {
        const checkbox = td.querySelector('input[type="checkbox"]');
        if (checkbox) {
            td.textContent = checkbox.checked ? 'Sim' : 'Não';
            return;
        }

        const select = td.querySelector('select');
        if (select) {
            const opt = select.options[select.selectedIndex];
            td.textContent = opt ? opt.text : '-';
            return;
        }

        const inputTexto = td.querySelector('input[type="text"], input:not([type]), textarea');
        if (inputTexto) {
            td.textContent = inputTexto.value || '-';
            return;
        }
    });

    // Remover coluna de ações, se existir
    const headerRow = clone.querySelector('thead tr');
    let indiceAcoes = -1;
    if (headerRow) {
        const headers = Array.from(headerRow.children);
        indiceAcoes = headers.findIndex(th => {
            const txt = (th.textContent || '').trim().toUpperCase();
            return txt === 'AÇÕES' || txt === 'ACOES';
        });
        if (indiceAcoes >= 0 && headers[indiceAcoes]) {
            headers[indiceAcoes].remove();
        }
    }

    if (indiceAcoes >= 0) {
        clone.querySelectorAll('tbody tr, tfoot tr').forEach(row => {
            const cells = Array.from(row.children);
            if (cells[indiceAcoes]) {
                cells[indiceAcoes].remove();
            }
        });
    }

    const dataAgora = new Date().toLocaleString('pt-BR');
    const subtituloFinal = subtitulo ? `<div style="margin-bottom:8px;font-size:13px;color:#222">${subtitulo}</div>` : '';

    const win = window.open('', '_blank', 'width=1100,height=700');
    if (!win) {
        alert('Não foi possível abrir a janela de impressão. Permita popups ou use a impressão do navegador.');
        return;
    }

    win.document.write(`
        <!doctype html>
        <html lang="pt-BR">
        <head>
            <meta charset="utf-8">
            <title>${titulo}</title>
            <link rel="stylesheet" href="styles.css">
            <style>
                @page { size: A4 ${orientacao}; margin: 10mm; }
                body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Arial, sans-serif; padding: 12px; color: #222; }
                h1 { margin-bottom: 8px; font-size: 18px; }
                .meta { margin-bottom: 10px; font-size: 13px; color: #222; }
                .report-printable table { width: 100%; border-collapse: collapse; font-size: 12px; }
                .report-printable th, .report-printable td { border: 1px solid #ddd; padding: 6px 8px; text-align: left; }
                .report-printable thead th { background: #1e3a5f; color: #fff; }
                @media print { body { margin: 0; } }
            </style>
        </head>
        <body>
            <h1>${titulo}</h1>
            <div class="meta"><strong>Data:</strong> ${dataAgora}</div>
            ${subtituloFinal}
            <div class="report-printable">${clone.outerHTML}</div>
            <script>window.onload = function(){ setTimeout(function(){ window.print(); }, 200); };</script>
        </body>
        </html>
    `);
    win.document.close();
}

function imprimirControleEnvio() {
    renderizarControleEnvio();

    const tabela = document.getElementById('tabelaControleEnvio');
    if (!tabela) { mostrarNotificacao('Tabela não encontrada.', 'error'); return; }

    const filtroRep    = document.getElementById('filtroControleEnvioRep')?.value      || '';
    const filtroSistema  = document.getElementById('filtroControleEnvioSistema')?.value  || '';
    const filtroAssinado = document.getElementById('filtroControleEnvioAssinado')?.value || '';
    const filtroEnviado  = document.getElementById('filtroControleEnvioEnviado')?.value  || '';

    const fmt = v => v === 'sim' ? 'Marcado' : v === 'nao' ? 'Não marcado' : 'Todos';

    // Clonar tabela e converter campos editáveis para texto simples
    const clone = tabela.cloneNode(true);

    clone.querySelectorAll('td').forEach(td => {
        const cb = td.querySelector('input[type="checkbox"]');
        if (cb) { td.innerHTML = cb.checked ? '✔' : ''; return; }

        const status = td.querySelector('.status-indicator');
        if (status) { td.innerHTML = status.classList.contains('checked') ? '✔' : ''; return; }

        const inp = td.querySelector('input[type="text"], textarea');
        if (inp) { td.textContent = inp.value || '-'; return; }

        const sel = td.querySelector('select');
        if (sel) { td.textContent = sel.options[sel.selectedIndex]?.text || '-'; return; }
    });

    // Remover coluna Ações
    const headerCells = Array.from(clone.querySelectorAll('thead tr:first-child th'));
    const idxAcoes = headerCells.findIndex(th => /^a[çc][oõ]es$/i.test(th.textContent.trim()));
    if (idxAcoes >= 0) {
        clone.querySelectorAll('thead tr, tbody tr, tfoot tr').forEach(row => {
            const cells = Array.from(row.children);
            if (cells[idxAcoes]) cells[idxAcoes].remove();
        });
    }

    const dataAgora = new Date().toLocaleString('pt-BR');
    const filtrosHtml = `<div class="filtros-info">
        <strong>Representante:</strong> ${filtroRep || 'Todos'} &nbsp;|&nbsp;
        <strong>Sistema:</strong> ${fmt(filtroSistema)} &nbsp;|&nbsp;
        <strong>Assinado:</strong> ${fmt(filtroAssinado)} &nbsp;|&nbsp;
        <strong>Enviado:</strong> ${fmt(filtroEnviado)}
    </div>`;

    const win = window.open('', '_blank', 'width=1200,height=700');
    if (!win) { alert('Permita popups para imprimir.'); return; }

    win.document.write(`
        <!doctype html>
        <html lang="pt-BR">
        <head>
            <meta charset="utf-8">
            <title>Controle de Envio de Contratos</title>
            <style>
                @page {
                    size: A4 portrait;
                    margin: 8mm 6mm;
                }
                * { box-sizing: border-box; margin: 0; padding: 0; }
                body {
                    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Arial, sans-serif;
                    font-size: 8px;
                    color: #222;
                    padding: 6px 6px;
                }
                h1 {
                    font-size: 12px;
                    margin-bottom: 3px;
                    color: #1e3a5f;
                }
                .meta {
                    font-size: 8px;
                    color: #555;
                    margin-bottom: 2px;
                }
                .filtros-info {
                    font-size: 8px;
                    color: #333;
                    margin-bottom: 5px;
                    padding: 2px 5px;
                    background: #f4f6f9;
                    border-left: 3px solid #1e3a5f;
                    line-height: 1.3;
                }
                table {
                    width: 100%;
                    border-collapse: collapse;
                    table-layout: fixed;
                    font-size: 7.5px;
                }
                th, td {
                    border: 1px solid #ccc;
                    padding: 2px 3px;
                    text-align: left;
                    overflow: hidden;
                    text-overflow: ellipsis;
                    white-space: nowrap;
                    line-height: 1.2;
                }
                /* Permitir quebra de linha na coluna NOME para impressão */
                th:nth-child(2), td:nth-child(2) {
                    white-space: normal;
                    overflow-wrap: anywhere;
                }
                thead th {
                    background: #1e3a5f;
                    color: #fff;
                    font-size: 7.5px;
                    font-weight: 600;
                    text-align: center;
                    padding: 3px 3px;
                }
                tbody tr:nth-child(even) { background: #f7f9fc; }
                /* Larguras ajustadas para impressão: CTR menor, REPRESENTANTE menor, NOME maior */
                /* CTR | NOME | REP | SISTEMA | ASSINADO | ENVIADO | SOLICITAÇÃO */
                col.col-ctr        { width: 5%; }
                col.col-nome       { width: 45%; }
                col.col-rep        { width: 8%; }
                col.col-sistema    { width: 7%; }
                col.col-assinado   { width: 7%; }
                col.col-enviado    { width: 7%; }
                col.col-solic      { width: 21%; }
                /* Centralizar colunas de marcação */
                td:nth-child(4), td:nth-child(5), td:nth-child(6) { text-align: center; }
                .badge-rep {
                    display: inline-block;
                    padding: 0px 4px;
                    border-radius: 3px;
                    font-size: 7px;
                    font-weight: 700;
                    color: #fff;
                    background: #1e3a5f;
                }
                @media print { body { padding: 0; } }
            </style>
        </head>
        <body>
            <h1>Controle de Envio de Contratos</h1>
            <div class="meta"><strong>Data:</strong> ${dataAgora}</div>
            ${filtrosHtml}
            <colgroup>
                <col class="col-ctr">
                <col class="col-nome">
                <col class="col-rep">
                <col class="col-sistema">
                <col class="col-assinado">
                <col class="col-enviado">
                <col class="col-solic">
            </colgroup>
            ${clone.outerHTML}
            <script>window.onload = function(){ setTimeout(function(){ window.print(); }, 250); };<\/script>
        </body>
        </html>
    `);
    win.document.close();
}

function imprimirDashboardQtdProduto() {
    renderizarDashboard();
    imprimirTabelaSeparada('tabelaDashboardQtdProduto', 'Dashboard - Quantidade Vendida por Produto', '', 'portrait');
}

function imprimirDashboardValorProduto() {
    renderizarDashboard();
    imprimirTabelaSeparada('tabelaDashboardValorProduto', 'Dashboard - Valor das Vendas por Produto', '', 'portrait');
}

function imprimirDashboardVendasRepresentante() {
    renderizarDashboard();
    imprimirTabelaSeparada('tabelaVendasRep', 'Dashboard - Quantidade de Vendas por Representante', '', 'landscape');
}

// Dispatcher: imprime o relatório associado à aba informada
function imprimirRelatorioAba(tabId) {
    try {
        if (!tabId) tabId = document.querySelector('.tabs-navigation .tab-btn.active')?.dataset.tab || 'estoque';
        switch ((tabId || '').toString()) {
            case 'estoque':
                // utiliza o mecanismo de relatório/inventário já existente
                prepararRelatorioInventario();
                imprimirInventario();
                break;
            case 'distribuicao':
                imprimirDistribuicao();
                break;
            case 'vendas':
                imprimirVendas();
                break;
            case 'controleenvio':
                imprimirControleEnvio();
                break;
            case 'dashboard':
                // imprime o principal quadro de vendas por representante (mais completo)
                renderizarDashboard();
                imprimirDashboardVendasRepresentante();
                break;
            case 'relatorios':
                // decide com base no tipo selecionado
                const tipo = document.getElementById('filtroRelatoriosTipo')?.value || 'inventario';
                if (tipo === 'comissoes') imprimirComissoes();
                else if (tipo === 'distribuicao') imprimirDistribuicao();
                else imprimirInventario();
                break;
            case 'configuracoes':
                imprimirConfiguracoes();
                break;
            case 'clientes':
                imprimirClientes();
                break;
            case 'precificacao':
                imprimirPrecificacao();
                break;
            case 'propostas':
                imprimirPropostas();
                break;
            default:
                alert('Não há relatório configurado para esta aba.');
        }
    } catch (e) {
        console.error('Erro imprimindo aba', tabId, e);
        mostrarNotificacao('Falha ao iniciar impressão. Veja o console.', 'error');
    }
}

function imprimirDistribuicao() {
    prepararRelatorioDistribuicao();
    const preview = document.getElementById('relatoriosPreview');
    if (!preview) { mostrarNotificacao('Preview não disponível para distribuição.', 'error'); return; }
    const content = preview.innerHTML;
    const win = window.open('', '_blank', 'width=1000,height=700');
    if (!win) { alert('Não foi possível abrir a janela de impressão. Permita popups.'); return; }
    win.document.write(`
        <!doctype html>
        <html lang="pt-BR">
        <head>
            <meta charset="utf-8">
            <title>Relatório - Distribuição</title>
            <link rel="stylesheet" href="styles.css">
            <style> @page { size: A4 portrait; margin: 10mm; } body{font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Arial, sans-serif; padding:12px; color:#222 } .report-printable table{width:100%;border-collapse:collapse} .report-printable th, .report-printable td{border:1px solid #ddd;padding:6px 8px} thead{background:#1e3a5f;color:#fff} </style>
        </head>
        <body>
            <h1>Relatório de Distribuição</h1>
            ${content}
            <script>window.onload=function(){ setTimeout(function(){ window.print(); },200); };</script>
        </body>
        </html>
    `);
    win.document.close();
}

function imprimirVendas() {
    // Gera relatório em uma única tabela: cada produto/linha de venda em uma linha, com subtotal por contrato
    renderizarRegistroVendas();

    const vendas = Array.isArray(estoque.registroVendas) ? [...estoque.registroVendas] : [];
    if (vendas.length === 0) {
        mostrarNotificacao('Não há vendas registradas para imprimir.', 'warning');
        return;
    }

    // Montar linhas expandidas
    const linhas = [];
    vendas.forEach(venda => {
        const contratoKey = normalizarContratoKey(venda.contrato || '');
        const dataNorm = parseDateToYYYYMMDD(venda.data) || '';
        if (Array.isArray(venda.items) && venda.items.length > 0) {
            venda.items.forEach(it => {
                linhas.push({
                    contratoKey,
                    contratoRaw: venda.contrato || contratoKey,
                    vendaId: venda.id,
                    dataNorm,
                    loja: venda.loja || '-',
                    representante: venda.representante || '-',
                    produtoNome: it.produtoNome || '-',
                    quantidade: Number(it.quantidade || 0),
                    valorUnitario: Number(it.valorUnitario || 0),
                    valorTotal: Number(it.valorTotal !== undefined ? it.valorTotal : ((Number(it.valorUnitario)||0) * (Number(it.quantidade)||0))),
                    observacoes: venda.observacoes || '-'
                });
            });
        } else {
            linhas.push({
                contratoKey,
                contratoRaw: venda.contrato || contratoKey,
                vendaId: venda.id,
                dataNorm,
                loja: venda.loja || '-',
                representante: venda.representante || '-',
                produtoNome: venda.produtoNome || '-',
                quantidade: Number(venda.quantidade || 0),
                valorUnitario: Number(venda.valorUnitario || 0),
                valorTotal: Number(venda.valorTotal !== undefined ? venda.valorTotal : ((Number(venda.valorUnitario)||0) * (Number(venda.quantidade)||0))),
                observacoes: venda.observacoes || '-'
            });
        }
    });

    // Agrupar por contrato e ordenar
    const grupos = {};
    linhas.forEach(l => { if (!grupos[l.contratoKey]) grupos[l.contratoKey] = []; grupos[l.contratoKey].push(l); });
    const chavesOrdenadas = Object.keys(grupos).sort((a,b) => {
        const na = parseInt(a); const nb = parseInt(b);
        if (!isNaN(na) && !isNaN(nb)) return na - nb;
        return a.localeCompare(b);
    });

    // Construir tabela única (removida coluna OBS; mesclar células de contrato/loja/total por contrato quando for >1 linha)
    let tabelaHtml = `
        <table class="tabela-relatorio vendas-table" style="width:100%;border-collapse:collapse">
            <colgroup>
                <col style="width:8%" />
                <col style="width:18%" />
                <col style="width:10%" />
                <col style="width:18%" />
                <col style="width:4%" />
                <col style="width:10%" />
                <col style="width:14%" />
                <col style="width:12%" />
                <col style="width:6%" />
            </colgroup>
            <thead>
                <tr>
                    <th style="padding:6px;border:1px solid #ddd;vertical-align:middle;text-align:center">CONTRATO</th>
                    <th style="padding:6px;border:1px solid #ddd;vertical-align:middle;text-align:center">LOJA / CLIENTE</th>
                    <th style="padding:6px;border:1px solid #ddd;vertical-align:middle;text-align:center">REPRESENTANTE</th>
                    <th style="padding:6px;border:1px solid #ddd;vertical-align:middle;text-align:center">PRODUTO</th>
                    <th class="numeric" style="padding:6px;border:1px solid #ddd;vertical-align:middle;text-align:center">QTD</th>
                    <th class="numeric" style="padding:6px;border:1px solid #ddd;vertical-align:middle;text-align:center">VALOR UN.</th>
                    <th class="numeric" style="padding:6px;border:1px solid #ddd;vertical-align:middle;text-align:center">VALOR TOTAL</th>
                    <th class="numeric" style="padding:6px;border:1px solid #ddd;vertical-align:middle;text-align:center">TOTAL CONTRATO (R$)</th>
                    <th style="padding:6px;border:1px solid #ddd;vertical-align:middle;text-align:center">DATA</th>
                </tr>
            </thead>
            <tbody>`;

    let grandTotalQtd = 0;
    let grandTotalValor = 0;

    chavesOrdenadas.forEach(ck => {
        const grupo = grupos[ck] || [];
        if (!grupo.length) return;
        // calcular subtotal do contrato
        let subtotalQtd = 0; let subtotalValor = 0;
        grupo.forEach(r => { subtotalQtd += r.quantidade || 0; subtotalValor += r.valorTotal || 0; });

        // adicionar linhas do contrato — se houver mais de uma linha, usar rowspan para CONTRATO, LOJA e TOTAL CONTRATO
        const rowspanAttr = grupo.length > 1 ? ` rowspan="${grupo.length}"` : '';
        let primeiraLinha = true;
        grupo.forEach(r => {
            const dataFmt = formatDateToDDMMYYYY(r.dataNorm);
            tabelaHtml += `<tr>`;

            if (primeiraLinha) {
                tabelaHtml += `<td style="padding:6px;border:1px solid #ddd;text-align:center"${rowspanAttr}>${r.contratoRaw || ck}</td>`;
                tabelaHtml += `<td style="padding:6px;border:1px solid #ddd;text-align:center"${rowspanAttr}>${r.loja}</td>`;
            }

            tabelaHtml += `
                <td style="padding:6px;border:1px solid #ddd;vertical-align:middle;text-align:center">${r.representante}</td>
                <td style="padding:6px;border:1px solid #ddd;vertical-align:middle;white-space:normal;word-break:break-word;overflow-wrap:anywhere;text-align:center">${r.produtoNome}</td>
                <td class="numeric" style="padding:6px;border:1px solid #ddd;text-align:center;vertical-align:middle">${r.quantidade}</td>
                <td class="numeric" style="padding:6px;border:1px solid #ddd;text-align:center;vertical-align:middle">${r.valorUnitario ? formatarMoedaValor(r.valorUnitario) : '-'}</td>
                <td class="numeric" style="padding:6px;border:1px solid #ddd;text-align:center;vertical-align:middle">${formatarMoedaValor(r.valorTotal || 0)}</td>`;

            if (primeiraLinha) {
                tabelaHtml += `<td class="numeric" style="padding:6px;border:1px solid #ddd;text-align:center;vertical-align:middle"${rowspanAttr}><strong>${formatarMoedaValor(subtotalValor)}</strong></td>`;
            }

            tabelaHtml += `<td style="padding:6px;border:1px solid #ddd;vertical-align:middle;text-align:center">${dataFmt}</td>`;
            tabelaHtml += `</tr>`;

            primeiraLinha = false;
            grandTotalQtd += r.quantidade || 0;
            grandTotalValor += r.valorTotal || 0;
        });
        // não adicionar linha de subtotal separada — valor mostrado na primeira linha do contrato
    });

    tabelaHtml += `</tbody></table>`;

    tabelaHtml += `<div style="margin-top:12px;font-weight:700">Total Geral: ${grandTotalQtd} item(ns) — ${formatarMoedaValor(grandTotalValor)}</div>`;

    const win = window.open('', '_blank', 'width=1200,height=900');
    if (!win) { alert('Não foi possível abrir a janela de impressão. Permita popups.'); return; }

    // Forçar estilos inline para evitar CSS de impressão que esconda o conteúdo
    win.document.write(`
        <!doctype html>
        <html lang="pt-BR">
        <head>
            <meta charset="utf-8">
            <title>Relatório - Registro de Vendas</title>
            <style>
                @page { size: A4 landscape; margin: 8mm; }
                html, body { height: auto; margin: 0; padding:8px; }
                body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Arial, sans-serif; color:#000; background:#fff; font-size:11px; -webkit-print-color-adjust: exact; }
                h1 { margin:0 0 6px 0; font-size:14px }
                /* reduzir fonte da tabela para caber mais conteúdo */
                .tabela-relatorio, .tabela-relatorio table { font-size:9px; border-collapse:collapse; width:100% !important; table-layout: fixed; }
                /* garantir box-sizing e reduzir padding para evitar corte do conteúdo */
                .tabela-relatorio th, .tabela-relatorio td { box-sizing: border-box; border:1px solid #ddd; padding:3px !important; vertical-align: middle; word-break: break-word; }
                /* evitar quebra nas colunas numéricas */
                .tabela-relatorio th.numeric, .tabela-relatorio td.numeric { white-space: nowrap; }
                .tabela-relatorio thead th { background: #1e3a5f; color: #fff; text-align: center; font-size:10px }
                .tabela-relatorio tbody tr:nth-child(even) { background: #f7f9fc; }
                .tabela-relatorio td { background: #fff; }
                .tabela-relatorio td[colspan] { white-space: normal; }
                /* garantir que o wrapper de impressão não seja posicionado fora da página */
                .report-printable { position: static !important; left: auto !important; top: auto !important; width: 100% !important; }
                table { page-break-inside: auto; }
                /* permitir que a tabela quebre entre páginas — evitando empurrar toda a tabela para a próxima página */
                tr { page-break-inside: auto; page-break-after: auto }
                thead { display: table-header-group }
                tfoot { display: table-footer-group }
                @media print { body { margin: 0; } .tabela-relatorio { break-inside: auto; } }
            </style>
        </head>
        <body>
            <div class="report-printable">
                <h1 style="margin-bottom:8px">Registro de Vendas</h1>
                ${tabelaHtml}
            </div>
            <script>window.onload=function(){ setTimeout(function(){ window.print(); },200); };</script>
        </body>
        </html>
    `);
    win.document.close();
}

function imprimirConfiguracoes() {
    const tab = document.getElementById('tab-configuracoes');
    if (!tab) { mostrarNotificacao('Aba de configurações não encontrada.', 'error'); return; }
    const contentArea = tab.querySelector('.content-area');
    const cloneHtml = contentArea ? contentArea.innerHTML : tab.innerHTML;
    const win = window.open('', '_blank', 'width=900,height=700');
    if (!win) { alert('Não foi possível abrir a janela de impressão. Permita popups.'); return; }
    win.document.write(`
        <!doctype html>
        <html lang="pt-BR">
        <head>
            <meta charset="utf-8">
            <title>Configurações</title>
            <link rel="stylesheet" href="styles.css">
            <style> @page{size:A4 portrait;margin:10mm} body{font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Arial, sans-serif;padding:12px;color:#222} </style>
        </head>
        <body>
            <h1>Configurações do Sistema</h1>
            <div>${cloneHtml}</div>
            <script>window.onload=function(){ setTimeout(function(){ window.print(); },200); };</script>
        </body>
        </html>
    `);
    win.document.close();
}

// ================================
// CONTROLE IMBEL (separado) - storage chave: 'estoqueImbelV1'
// ================================
const IMBEL_KEY = 'estoqueImbelV1';

// Tenta migrar dados IMBEL caso exista alguma chave antiga no localStorage
function migrateImbelStorage() {
    try {
        // já existe a chave atual
        if (localStorage.getItem(IMBEL_KEY)) return false;

        const candidates = Object.keys(localStorage).filter(k => /imbel/i.test(k) && k !== IMBEL_KEY);
        if (candidates.length === 0) return false;

        for (const k of candidates) {
            try {
                const raw = localStorage.getItem(k);
                if (!raw) continue;
                const v = JSON.parse(raw);
                if (v && (Array.isArray(v.produtos) || Array.isArray(v.movimentacoes) || (v.produtos && v.movimentacoes))) {
                    localStorage.setItem(IMBEL_KEY, JSON.stringify(v));
                    console.info(`IMBEL: migrado ${k} -> ${IMBEL_KEY}`);
                    try { mostrarNotificacao(`Dados IMBEL restaurados de "${k}"`, 'success'); } catch(e) {}
                    return true;
                }
            } catch(e) {
                // ignora chaves que não são JSON válidos
            }
        }
    } catch(e) {}
    return false;
}

function loadImbel() {
    try {
        return JSON.parse(localStorage.getItem(IMBEL_KEY) || '{"produtos":[],"movimentacoes":[]}');
    } catch (e) { return {produtos:[], movimentacoes:[]}; }
}

function saveImbel(data) {
    try { localStorage.setItem(IMBEL_KEY, JSON.stringify(data)); } catch (e) { console.error('Erro salvando IMBEL', e); }
}

function initControleImbel() {
    // mostrar estoques por padrão quando a aba for aberta
    // hook: chamar trocarSubAbaControleImbel('estoque') para inicializar
    try { trocarSubAbaControleImbel('estoque'); } catch(e) {}
}

function trocarSubAbaControleImbel(sub) {
    // Atualizar botões ativos (sub-navegação IMBEL)
    document.querySelectorAll('#imbelSubNav .tab-btn').forEach(btn => btn.classList.remove('active'));
    const activeBtn = document.querySelector(`#imbelSubNav .tab-btn[data-imbeltab="${sub}"]`);
    if (activeBtn) activeBtn.classList.add('active');

    // Mostrar / ocultar painéis
    document.querySelectorAll('.controleimbel-subtab').forEach(el => el.style.display = 'none');
    const sel = document.getElementById(`controleImbel-${sub}`);
    if (sel) sel.style.display = 'block';

    // Renderizar conteúdo
    if (sub === 'estoque') renderControleImbelEstoque();
    else if (sub === 'cadastro') renderControleImbelCadastro();
    else if (sub === 'movimentacao') renderControleImbelMovimentacao();
    else if (sub === 'dashboard') renderControleImbelDashboard();
}

// ========== MODAL FUNCTIONS ==========
function openImbelProdModal() {
    const modal = document.getElementById('imbel_prod_modal');
    if (modal) modal.classList.add('open');
}

function closeImbelProdModal() {
    const modal = document.getElementById('imbel_prod_modal');
    if (modal) modal.classList.remove('open');
    // limpar formulário
    document.getElementById('imbel_prod_nome').value = '';
    document.getElementById('imbel_prod_codigo').value = '';
    document.getElementById('imbel_prod_qtd_inicial').value = '0';
    document.getElementById('imbel_prod_obs').value = '';
    const editField = document.getElementById('imbel_prod_edit_id'); if (editField) editField.value = '';
    document.getElementById('imbel_prod_salvar').textContent = 'Salvar';
}

function openImbelMovModal() {
    const modal = document.getElementById('imbel_mov_modal');
    if (modal) modal.classList.add('open');
}

function closeImbelMovModal() {
    const modal = document.getElementById('imbel_mov_modal');
    if (modal) modal.classList.remove('open');
    // limpar formulário
    document.getElementById('imbel_mov_prod').value = '';
    document.getElementById('imbel_mov_tipo').value = 'Entrada';
    const hoje = new Date().toISOString().slice(0,10);
    document.getElementById('imbel_mov_data').value = hoje;
    document.getElementById('imbel_mov_qtd').value = '1';
    document.getElementById('imbel_mov_dest').value = '';
    document.getElementById('imbel_mov_cpf').value = '';
    document.getElementById('imbel_mov_valor').value = '';
    document.getElementById('imbel_mov_endereco').value = '';
    document.getElementById('imbel_mov_tel').value = '';
    document.getElementById('imbel_mov_email').value = '';
    document.getElementById('imbel_mov_obs').value = '';
    const editField = document.getElementById('imbel_mov_edit_id'); if (editField) editField.value = '';
    document.getElementById('imbel_mov_salvar').textContent = 'Registrar Movimentação';
}

// Bind modal close buttons to modals
document.addEventListener('DOMContentLoaded', function(){
    // Fechar modal ao clicar no backdrop
    document.getElementById('imbel_prod_modal').querySelector('.imbel-modal-backdrop').onclick = closeImbelProdModal;
    document.getElementById('imbel_mov_modal').querySelector('.imbel-modal-backdrop').onclick = closeImbelMovModal;
    
    // Fechar modal ao clicar no X
    document.getElementById('imbel_prod_modal').querySelector('.imbel-modal-close').onclick = closeImbelProdModal;
    document.getElementById('imbel_mov_modal').querySelector('.imbel-modal-close').onclick = closeImbelMovModal;
    
    // Fechar modal ao clicar em Cancelar
    document.querySelectorAll('#imbel_prod_modal .imbel-modal-cancel').forEach(btn => {
        btn.onclick = closeImbelProdModal;
    });
    document.querySelectorAll('#imbel_mov_modal .imbel-modal-cancel').forEach(btn => {
        btn.onclick = closeImbelMovModal;
    });
    
    // Fechar modal com ESC
    document.addEventListener('keydown', function(e){
        if (e.key === 'Escape') {
            closeImbelProdModal();
            closeImbelMovModal();
        }
    });

    // Tentar migrar dados IMBEL de chaves antigas (se houver)
    try { migrateImbelStorage(); } catch(e) {}
}, { once: true });

function renderControleImbelDashboard() {
    const data = loadImbel();
    const container = document.getElementById('controleImbelDashboardContainer');
    if (!container) return;
    container.innerHTML = '';

    // Top indicators
    const movimentacoes = (data.movimentacoes || []).slice();
    // Considerar apenas SAÍDA como venda para faturamento
    const vendas = movimentacoes.filter(m => ((m.tipo||'').toString().toUpperCase() === 'SAÍDA' || (m.tipo||'').toString().toUpperCase() === 'SAÍDA'));
    const receitaTotal = vendas.reduce((s,m) => s + (Number(m.valor)||0), 0);
    const unidadesVendidas = vendas.reduce((s,m) => s + (Number(m.quantidade)||0), 0);
    const clientes = new Set(vendas.filter(v=> (v.destinatario||'').trim() ).map(v => (v.destinatario||'').toString().toUpperCase()));
    const clientesAtivos = clientes.size;
    const pedidosCount = vendas.length;
    const ticketMedio = pedidosCount > 0 ? receitaTotal / pedidosCount : 0;

    const indicadores = document.createElement('div');
    indicadores.style.cssText = 'display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:12px;margin-bottom:12px';
    const card = (titulo, valor, destaque) => {
        const c = document.createElement('div');
        c.style.cssText = 'background:#fff;padding:14px;border-radius:10px;box-shadow:0 1px 6px rgba(0,0,0,.06)';
        const innerVal = destaque ? `<span style="color:#1e3a5f">${valor}</span>` : `${valor}`;
        c.innerHTML = `<div style="font-size:.78rem;color:#666">${titulo}</div><div style="font-size:1.3rem;font-weight:700;margin-top:6px">${innerVal}</div>`;
        return c;
    };

    indicadores.appendChild(card('Receita total', 'R$ ' + receitaTotal.toLocaleString('pt-BR',{minimumFractionDigits:2}), true));
    indicadores.appendChild(card('Unidades vendidas', unidadesVendidas.toString(), false));
    indicadores.appendChild(card('Clientes ativos', clientesAtivos.toString(), false));
    indicadores.appendChild(card('Ticket médio', 'R$ ' + ticketMedio.toLocaleString('pt-BR',{minimumFractionDigits:2}), false));

    container.appendChild(indicadores);

    // Controle financeiro: pendentes e comparação
    const pendentes = vendas.filter(v => (v.pagamento||'').toString().toUpperCase() !== 'SIM');
    const valorPendentes = pendentes.reduce((s,m) => s + (Number(m.valor)||0), 0);
    const pagosCount = vendas.filter(v => (v.pagamento||'').toString().toUpperCase() === 'SIM').length;
    const naoConfirmadosCount = vendas.length - pagosCount;

    const financeiroWrap = document.createElement('div');
    financeiroWrap.style.cssText = 'display:flex;gap:12px;flex-wrap:wrap;margin-bottom:12px;align-items:stretch';
    const finCard = (title, value, sub) => {
        const el = document.createElement('div');
        el.style.cssText = 'background:#fff;padding:12px;border-radius:10px;box-shadow:0 1px 6px rgba(0,0,0,.06);min-width:200px;flex:1';
        el.innerHTML = `<div style="font-size:.78rem;color:#666">${title}</div><div style="font-size:1.1rem;font-weight:700;margin-top:6px">${value}</div>${sub?`<div style="font-size:.75rem;color:#888;margin-top:6px">${sub}</div>`:''}`;
        return el;
    };

    financeiroWrap.appendChild(finCard('Valor pendente de recebimento', 'R$ ' + valorPendentes.toLocaleString('pt-BR',{minimumFractionDigits:2}), `${pendentes.length} pedidos pendentes`));
    financeiroWrap.appendChild(finCard('Pedidos pagos', pagosCount.toString(), 'Confirmados'));
    financeiroWrap.appendChild(finCard('Pedidos não confirmados', naoConfirmadosCount.toString(), 'Aguardando comprovante/GRU'));

    // Chart area (paid x unconfirmed)
    const chartCard = document.createElement('div');
    chartCard.style.cssText = 'background:#fff;padding:12px;border-radius:10px;box-shadow:0 1px 6px rgba(0,0,0,.06);min-width:260px;flex:1';
    chartCard.innerHTML = `<div style="font-size:.78rem;color:#666">Comparação: Pagos vs Não confirmados</div><canvas id="imbelPaidChart" style="height:80px;margin-top:8px"></canvas>`;
    financeiroWrap.appendChild(chartCard);

    container.appendChild(financeiroWrap);

    // render chart if Chart available
    setTimeout(() => {
        try {
            const ctx = document.getElementById('imbelPaidChart');
            if (ctx && window.Chart) {
                // destroy previous chart instance if any
                if (ctx._chartInstance) try { ctx._chartInstance.destroy(); } catch(e){}
                const ch = new Chart(ctx.getContext('2d'), {
                    type: 'bar',
                    data: {
                        labels: ['Pagos','Não confirmados'],
                        datasets: [{
                            label: 'Pedidos',
                            data: [pagosCount, naoConfirmadosCount],
                            backgroundColor: ['#28a745','#ffc107']
                        }]
                    },
                    options: {plugins:{legend:{display:false}},scales:{y:{beginAtZero:true,ticks:{precision:0}}}}
                });
                ctx._chartInstance = ch;
            }
        } catch (e) { console.warn('Chart render falhou', e); }
    }, 40);

    // Gestão de estoque em tempo real com semáforo
    const estoqueWrap = document.createElement('div');
    estoqueWrap.style.cssText = 'background:#fff;border-radius:10px;padding:12px;box-shadow:0 1px 6px rgba(0,0,0,.06);margin-top:12px';
    estoqueWrap.innerHTML = `<h3 style="margin:0 0 8px 0;font-size:1rem">Gestão de Estoque</h3>`;
    const tabela = document.createElement('table');
    tabela.style.cssText = 'width:100%;border-collapse:collapse;font-size:.9rem';
    tabela.innerHTML = `<thead><tr style="background:#1e3a5f;color:#fff"><th style="padding:8px">Produto</th><th style="padding:8px">Estoque Atual</th><th style="padding:8px">Ponto Reposição</th><th style="padding:8px">Semáforo</th><th style="padding:8px">Tempo p/ zerar lote anterior</th></tr></thead><tbody></tbody>`;
    const tb = tabela.querySelector('tbody');

    const produtos = data.produtos || [];
    // função auxiliar para calcular saldo atual
    const calcSaldo = (prodId) => {
        const totE = (data.movimentacoes||[]).filter(m=>m.produtoId===prodId && (m.tipo||'').toString().toUpperCase()==='ENTRADA').reduce((s,m)=>s + (Number(m.quantidade)||0),0);
        const totS = (data.movimentacoes||[]).filter(m=>m.produtoId===prodId && ((m.tipo||'').toString().toUpperCase()==='SAÍDA' || (m.tipo||'').toString().toUpperCase()==='SAÍDA')).reduce((s,m)=>s + (Number(m.quantidade)||0),0);
        const inicial = Number((data.produtos||[]).find(p=>p.id===prodId)?.quantidadeInicial) || 0;
        return inicial + totE - totS;
    };

    // calcula tempo para zerar lote anterior (último ciclo completo)
    function tempoParaZerar(prodId) {
        const movs = (data.movimentacoes||[]).filter(m=>m.produtoId===prodId).slice().sort((a,b)=> (a.data||'').localeCompare(b.data||''));
        if (!movs.length) return null;
        let balance = 0;
        let lastEntradaDate = null;
        let lastZeroCycleDays = null;
        movs.forEach(m => {
            const tipo = (m.tipo||'').toString().toUpperCase();
            const q = Number(m.quantidade)||0;
            if (tipo === 'ENTRADA') {
                balance += q;
                lastEntradaDate = m.data || null;
            } else if (tipo === 'SAÍDA' || tipo === 'SAÍDA') {
                balance -= q;
            }
            if (lastEntradaDate && balance <= 0) {
                // ciclo zerou
                try {
                    const d1 = new Date(lastEntradaDate);
                    const d2 = new Date(m.data || lastEntradaDate);
                    const diff = Math.ceil((d2 - d1)/(1000*60*60*24));
                    lastZeroCycleDays = diff >= 0 ? diff : null;
                } catch(e) { lastZeroCycleDays = null; }
                // reset to look for more recent cycles
                lastEntradaDate = null;
                balance = 0;
            }
        });
        return lastZeroCycleDays;
    }

    // vamos acumular estatísticas de estoque enquanto construímos a tabela
    const esgotado = [];
    const baixo = [];
    const ok = [];
    const zerados = [];

    produtos.forEach(p => {
        const estoqueAtual = calcSaldo(p.id);
        const ponto = (p.pontoReposicao !== undefined) ? Number(p.pontoReposicao) : Math.max(1, Math.floor((Number(p.quantidadeInicial)||0) * 0.2));
        let status = 'OK';
        let color = '#28a745';
        if (estoqueAtual <= 0) { status = 'ESGOTADO'; color = '#dc3545'; esgotado.push(p); if (estoqueAtual === 0) zerados.push(p); }
        else if (estoqueAtual <= ponto) { status = 'BAIXO'; color = '#ffc107'; baixo.push(p); }
        else { ok.push(p); }

        const tempo = tempoParaZerar(p.id);
        const tr = document.createElement('tr');
        tr.style.background = '#fff';
        tr.innerHTML = `<td style="padding:8px;border:1px solid #eee">${p.nome}</td>
                        <td style="padding:8px;border:1px solid #eee;text-align:center">${estoqueAtual}</td>
                        <td style="padding:8px;border:1px solid #eee;text-align:center">${p.pontoReposicao !== undefined ? p.pontoReposicao : ponto}</td>
                        <td style="padding:8px;border:1px solid #eee;text-align:center"><span style="display:inline-block;padding:6px 10px;border-radius:12px;background:${color};color:#fff;font-weight:700">${status}</span></td>
                        <td style="padding:8px;border:1px solid #eee;text-align:center">${tempo===null?'-':(tempo + ' dias')}</td>`;
        tb.appendChild(tr);
    });

    // resumo rápido acima da tabela com alertas
    const resumo = document.createElement('div');
    resumo.style.cssText = 'display:flex;gap:12px;flex-wrap:wrap;margin-bottom:8px;align-items:center';
    const resumoItem = (label, value, color) => {
        const el = document.createElement('div');
        el.style.cssText = 'background:#fff;padding:8px;border-radius:8px;box-shadow:0 1px 4px rgba(0,0,0,.04);min-width:160px';
        el.innerHTML = `<div style="font-size:.78rem;color:#666">${label}</div><div style="font-weight:700;margin-top:6px;color:${color}">${value}</div>`;
        return el;
    };
    resumo.appendChild(resumoItem('Produtos esgotados', esgotado.length.toString(), '#dc3545'));
    resumo.appendChild(resumoItem('Produtos abaixo do ponto', baixo.length.toString(), '#856404'));
    resumo.appendChild(resumoItem('Produtos OK', ok.length.toString(), '#155724'));
    if (zerados.length) {
        const listaZ = document.createElement('div');
        listaZ.style.cssText = 'background:#fff;padding:8px;border-radius:8px;box-shadow:0 1px 4px rgba(0,0,0,.04);min-width:220px';
        listaZ.innerHTML = `<div style="font-size:.78rem;color:#666">Zerados</div><div style="font-size:.85rem;margin-top:6px;color:#333">${zerados.map(z=>z.nome).join(', ')}</div>`;
        resumo.appendChild(listaZ);
    }

    estoqueWrap.insertBefore(resumo, tabela);
    estoqueWrap.appendChild(tabela);
    container.appendChild(estoqueWrap);

    // Acompanhamento de pedidos (pipeline)
    const pipelineWrap = document.createElement('div');
    pipelineWrap.style.cssText = 'background:#fff;border-radius:10px;padding:12px;box-shadow:0 1px 6px rgba(0,0,0,.06);margin-top:12px';
    pipelineWrap.innerHTML = `<h3 style="margin:0 0 8px 0;font-size:1rem">Acompanhamento de Pedidos</h3>`;
    const tableP = document.createElement('table');
    tableP.style.cssText = 'width:100%;border-collapse:collapse;font-size:.9rem';
    tableP.innerHTML = `<thead><tr style="background:#1e3a5f;color:#fff"><th style="padding:8px">Nº</th><th style="padding:8px">Produto</th><th style="padding:8px">Data</th><th style="padding:8px">Cliente</th><th style="padding:8px">Pipeline</th><th style="padding:8px">Ações</th></tr></thead><tbody></tbody>`;
    const tpb = tableP.querySelector('tbody');

    const vendasLista = movimentacoes.filter(m => ((m.tipo||'').toString().toUpperCase()==='SAÍDA' || (m.tipo||'').toString().toUpperCase()==='SAÍDA')).slice().reverse();
    vendasLista.forEach((v, idx) => {
        const prod = (data.produtos||[]).find(p=>p.id===v.produtoId) || {nome: v.descricao || '-'};
        const dateFmt = formatDateToDDMMYYYY(v.data || '');
        const passos = [
            {label:'Pedido', ok:true},
            {label:'Comprovante', ok: (v.pagamento||'').toString().toUpperCase()==='SIM'},
            {label:'GRU paga', ok: (v.gruPago||'') === true ? true : false},
            {label:'Entregue', ok: (v.entregue||'').toString().toUpperCase()==='SIM'},
            {label:'FI', ok: (v.fi||'').toString().toUpperCase()==='SIM'}
        ];

        const passoHtml = passos.map(pas => `<span style="display:inline-block;margin-right:6px;padding:6px 8px;border-radius:12px;background:${pas.ok? '#d4edda':'#f8d7da'};color:${pas.ok? '#155724':'#721c24'};font-weight:600;font-size:.8rem">${pas.label}</span>`).join('');

        const tr = document.createElement('tr');
        tr.style.background = idx % 2 === 0 ? '#fff' : '#f7f9fc';
        tr.innerHTML = `<td style="padding:8px;border:1px solid #eee;text-align:center">${idx+1}</td>
                        <td style="padding:8px;border:1px solid #eee">${prod.nome}</td>
                        <td style="padding:8px;border:1px solid #eee;text-align:center">${dateFmt}</td>
                        <td style="padding:8px;border:1px solid #eee">${v.destinatario||'-'}</td>
                        <td style="padding:8px;border:1px solid #eee">${passoHtml}</td>
                        <td style="padding:8px;border:1px solid #eee;text-align:center"><button class="btn btn-outline" data-toggle-pag="${v.id}">Toggle Pag.</button> <button class="btn btn-outline" data-toggle-ent="${v.id}">Toggle Entregue</button> <button class="btn btn-outline" data-toggle-fi="${v.id}">Toggle FI</button></td>`;
        tpb.appendChild(tr);
    });

    pipelineWrap.appendChild(tableP);
    container.appendChild(pipelineWrap);

    // bind action buttons (toggle status)
    tpb.querySelectorAll('button[data-toggle-pag]').forEach(btn => btn.onclick = function(){
        const id = this.getAttribute('data-toggle-pag');
        const mov = (data.movimentacoes||[]).find(m=>m.id===id);
        if (!mov) return; mov.pagamento = (mov.pagamento||'').toString().toUpperCase()==='SIM' ? 'NÃO' : 'SIM'; saveImbel(data); renderControleImbelDashboard(); renderControleImbelMovimentacao(); renderControleImbelEstoque();
    });
    tpb.querySelectorAll('button[data-toggle-ent]').forEach(btn => btn.onclick = function(){
        const id = this.getAttribute('data-toggle-ent');
        const mov = (data.movimentacoes||[]).find(m=>m.id===id);
        if (!mov) return; mov.entregue = (mov.entregue||'').toString().toUpperCase()==='SIM' ? 'NÃO' : 'SIM'; saveImbel(data); renderControleImbelDashboard(); renderControleImbelMovimentacao(); renderControleImbelEstoque();
    });
    tpb.querySelectorAll('button[data-toggle-fi]').forEach(btn => btn.onclick = function(){
        const id = this.getAttribute('data-toggle-fi');
        const mov = (data.movimentacoes||[]).find(m=>m.id===id);
        if (!mov) return; mov.fi = (mov.fi||'').toString().toUpperCase()==='SIM' ? 'NÃO' : 'SIM'; saveImbel(data); renderControleImbelDashboard(); renderControleImbelMovimentacao(); renderControleImbelEstoque();
    });

    // --- Seções adicionais: Receita por produto e Análise de clientes ---
    // Receita por produto
    try {
        const receitaWrap = document.createElement('div');
        receitaWrap.style.cssText = 'background:#fff;border-radius:10px;padding:12px;box-shadow:0 1px 6px rgba(0,0,0,.06);margin-top:12px';
        receitaWrap.innerHTML = `<h3 style="margin:0 0 8px 0;font-size:1rem">Receita por Produto</h3>`;
        const tableR = document.createElement('table');
        tableR.style.cssText = 'width:100%;border-collapse:collapse;font-size:.9rem;margin-top:8px';
        tableR.innerHTML = `<thead><tr style="background:#1e3a5f;color:#fff"><th style="padding:8px">Produto</th><th style="padding:8px;text-align:center">Unid. Vendidas</th><th style="padding:8px;text-align:right">Receita</th></tr></thead><tbody></tbody>`;
        const tbr = tableR.querySelector('tbody');

        const receitaPorProduto = {};
        const unidadesPorProduto = {};
        (data.movimentacoes||[]).forEach(m => {
            if (!m.produtoId) return;
            const tipo = (m.tipo||'').toString().toUpperCase();
            if (tipo !== 'SAÍDA') return;
            const id = m.produtoId;
            receitaPorProduto[id] = (receitaPorProduto[id]||0) + (Number(m.valor)||0);
            unidadesPorProduto[id] = (unidadesPorProduto[id]||0) + (Number(m.quantidade)||0);
        });

        const produtosList = (data.produtos||[]).slice();
        produtosList.sort((a,b) => (receitaPorProduto[b.id]||0) - (receitaPorProduto[a.id]||0));
        produtosList.forEach(p => {
            const rec = receitaPorProduto[p.id] || 0;
            const unid = unidadesPorProduto[p.id] || 0;
            const tr = document.createElement('tr');
            tr.style.background = '#fff';
            tr.innerHTML = `<td style="padding:8px;border:1px solid #eee">${p.nome}</td><td style="padding:8px;border:1px solid #eee;text-align:center">${unid}</td><td style="padding:8px;border:1px solid #eee;text-align:right">R$ ${rec.toLocaleString('pt-BR',{minimumFractionDigits:2})}</td>`;
            tbr.appendChild(tr);
        });
        receitaWrap.appendChild(tableR);
        container.appendChild(receitaWrap);
    } catch(e){ console.warn('Erro ao gerar receita por produto', e); }

    // Análise de clientes: top 10 e clientes com múltiplos produtos
    try {
        const clientes = {};
        (data.movimentacoes||[]).forEach(m => {
            const tipo = (m.tipo||'').toString().toUpperCase();
            if (tipo !== 'SAÍDA') return;
            const nome = (m.destinatario||'').toString().trim();
            if (!nome) return;
            const key = nome.toUpperCase();
            clientes[key] = clientes[key] || {nome: nome, total:0, pedidos:0, produtos:new Set()};
            clientes[key].total += (Number(m.valor)||0);
            clientes[key].pedidos += 1;
            if (m.produtoId) clientes[key].produtos.add(m.produtoId);
        });

        const clientesArr = Object.values(clientes).map(c => ({nome:c.nome,total:c.total,pedidos:c.pedidos,produtosCount:c.produtos.size}));
        clientesArr.sort((a,b) => b.total - a.total);

        // Top 10
        const topWrap = document.createElement('div');
        topWrap.style.cssText = 'background:#fff;border-radius:10px;padding:12px;box-shadow:0 1px 6px rgba(0,0,0,.06);margin-top:12px';
        topWrap.innerHTML = `<h3 style="margin:0 0 8px 0;font-size:1rem">Top 10 Clientes (por Receita)</h3>`;
        const tableC = document.createElement('table');
        tableC.style.cssText = 'width:100%;border-collapse:collapse;font-size:.9rem;margin-top:8px';
        tableC.innerHTML = `<thead><tr style="background:#1e3a5f;color:#fff"><th style="padding:8px">#</th><th style="padding:8px">Cliente</th><th style="padding:8px;text-align:right">Total</th><th style="padding:8px;text-align:center">Pedidos</th><th style="padding:8px;text-align:center">Produtos distintos</th></tr></thead><tbody></tbody>`;
        const tbc = tableC.querySelector('tbody');
        clientesArr.slice(0,10).forEach((c, idx) => {
            const tr = document.createElement('tr');
            tr.style.background = '#fff';
            tr.innerHTML = `<td style="padding:8px;border:1px solid #eee;text-align:center">${idx+1}</td><td style="padding:8px;border:1px solid #eee">${c.nome}</td><td style="padding:8px;border:1px solid #eee;text-align:right">R$ ${c.total.toLocaleString('pt-BR',{minimumFractionDigits:2})}</td><td style="padding:8px;border:1px solid #eee;text-align:center">${c.pedidos}</td><td style="padding:8px;border:1px solid #eee;text-align:center">${c.produtosCount}</td>`;
            tbc.appendChild(tr);
        });
        topWrap.appendChild(tableC);
        container.appendChild(topWrap);

        // Clientes que compraram múltiplos produtos
        const multi = clientesArr.filter(c => c.produtosCount > 1);
        if (multi.length) {
            const multiWrap = document.createElement('div');
            multiWrap.style.cssText = 'background:#fff;border-radius:10px;padding:12px;box-shadow:0 1px 6px rgba(0,0,0,.06);margin-top:12px';
            multiWrap.innerHTML = `<h3 style="margin:0 0 8px 0;font-size:1rem">Clientes com múltiplos produtos</h3>`;
            const tbl = document.createElement('table');
            tbl.style.cssText = 'width:100%;border-collapse:collapse;font-size:.9rem;margin-top:8px';
            tbl.innerHTML = `<thead><tr style="background:#1e3a5f;color:#fff"><th style="padding:8px">Cliente</th><th style="padding:8px;text-align:right">Total</th><th style="padding:8px;text-align:center">Produtos distintos</th></tr></thead><tbody></tbody>`;
            const tbm = tbl.querySelector('tbody');
            multi.forEach(c => {
                const tr = document.createElement('tr');
                tr.style.background = '#fff';
                tr.innerHTML = `<td style="padding:8px;border:1px solid #eee">${c.nome}</td><td style="padding:8px;border:1px solid #eee;text-align:right">R$ ${c.total.toLocaleString('pt-BR',{minimumFractionDigits:2})}</td><td style="padding:8px;border:1px solid #eee;text-align:center">${c.produtosCount}</td>`;
                tbm.appendChild(tr);
            });
            multiWrap.appendChild(tbl);
            container.appendChild(multiWrap);
        }
    } catch(e) { console.warn('Erro ao gerar análise de clientes', e); }
}

function renderControleImbelEstoque() {
    const data = loadImbel();
    const container = document.getElementById('controleImbelEstoqueContainer');
    container.innerHTML = '';

    const wrap = document.createElement('div');
    wrap.style.cssText = 'overflow-x:auto;background:#fff;border-radius:10px;padding:16px;box-shadow:0 1px 4px rgba(0,0,0,.08)';

    const thStyle = 'padding:8px 12px;border:1px solid #ddd;background:#1e3a5f;color:#fff;font-size:.82rem;white-space:nowrap;text-align:center';
    const tabela = document.createElement('table');
    tabela.style.cssText = 'width:100%;border-collapse:collapse;font-size:.85rem';
    // Observação removida para simplificar a tabela IMBEL
    tabela.innerHTML = `<thead><tr>
        <th style="${thStyle}">ID</th>
        <th style="${thStyle}">Descrição</th>
        <th style="${thStyle}">Quantidade Inicial</th>
        <th style="${thStyle}">Entrada</th>
        <th style="${thStyle}">Saída</th>
        <th style="${thStyle}">Estoque Atual</th>
    </tr></thead><tbody></tbody>`;

    const tbody = tabela.querySelector('tbody');
    const tdBase  = 'padding:8px 12px;border:1px solid #ddd;vertical-align:middle';
    const tdCenter = tdBase + ';text-align:center';

    // Calcular totais de Entrada e Saída por produto
    const totEntrada = {};
    const totSaida   = {};
    (data.movimentacoes||[]).forEach(m => {
        if (!m.produtoId) return;
        const q = Number(m.quantidade) || 0;
        const tipo = (m.tipo || '').toString().toLowerCase();
        if (tipo === 'entrada') totEntrada[m.produtoId] = (totEntrada[m.produtoId]||0) + q;
        else if (tipo === 'saída' || tipo === 'saida') totSaida[m.produtoId] = (totSaida[m.produtoId]||0) + q;
    });

    const produtos = data.produtos || [];

    if (produtos.length === 0) {
        const tr = document.createElement('tr');
        tr.innerHTML = `<td colspan="6" style="${tdBase};text-align:center;color:#999;font-style:italic">Nenhum produto cadastrado. Cadastre na sub-aba "Cadastro de Produtos".</td>`;
        tbody.appendChild(tr);
    }

    produtos.forEach((p, idx) => {
        const entrada = totEntrada[p.id] || 0;
        const saida   = totSaida[p.id]   || 0;
        const inicial = Number(p.quantidadeInicial) || 0;
        const estoqueAtual = inicial + entrada - saida;
        const saldoColor = estoqueAtual < 0 ? '#721c24' : estoqueAtual === 0 ? '#856404' : '#155724';
        const saldoBg    = estoqueAtual < 0 ? '#f8d7da' : estoqueAtual === 0 ? '#fff3cd' : '#d4edda';

        const tr = document.createElement('tr');
        tr.style.background = idx % 2 === 0 ? '#fff' : '#f7f9fc';
        // Observação removida: manter apenas colunas essenciais
        tr.innerHTML = `
            <td style="${tdCenter}">${idx + 1}</td>
            <td style="${tdBase}">${p.nome}${p.codigo ? ' <span style="color:#888;font-size:.75rem">('+p.codigo+')</span>' : ''} <button class="btn btn-outline" data-editprod="${p.id}" style="margin-left:8px;padding:4px 8px;font-size:.8rem">Editar</button></td>
            <td style="${tdCenter};font-weight:600">${inicial}</td>
            <td style="${tdCenter};color:#155724;font-weight:600">${entrada}</td>
            <td style="${tdCenter};color:#721c24;font-weight:600">${saida}</td>
            <td style="${tdCenter};font-weight:700;background:${saldoBg};color:${saldoColor}">${estoqueAtual}</td>
        `;
        tbody.appendChild(tr);
    });

    // Linha de totais gerais
    if (produtos.length > 0) {
        const totalInicial = (produtos||[]).reduce((s,p)=>s + (Number(p.quantidadeInicial)||0),0);
        const totalE = Object.values(totEntrada).reduce((a,b)=>a+b,0);
        const totalS = Object.values(totSaida).reduce((a,b)=>a+b,0);
        const totalSaldo = totalInicial + totalE - totalS;
        const tr = document.createElement('tr');
        tr.style.cssText = 'background:#1e3a5f;color:#fff;font-weight:700';
        tr.innerHTML = `
            <td colspan="2" style="${tdBase};background:#1e3a5f;color:#fff;text-align:right">TOTAL GERAL</td>
            <td style="${tdCenter};background:#1e3a5f;color:#fff">${totalInicial}</td>
            <td style="${tdCenter};background:#1e3a5f;color:#fff">${totalE}</td>
            <td style="${tdCenter};background:#1e3a5f;color:#fff">${totalS}</td>
            <td style="${tdCenter};background:#1e3a5f;color:#fff">${totalSaldo}</td>
        `;
        tbody.appendChild(tr);
    }

    wrap.appendChild(tabela);
    container.appendChild(wrap);

    // Handler: editar produto a partir do Estoque
    tbody.querySelectorAll('button[data-editprod]').forEach(btn => btn.onclick = function(){
        const id = this.getAttribute('data-editprod');
        if (!id) return;
        editarProdutoPorId(id);
    });

    // Atalho para cadastrar produto
    const btnCad = document.createElement('button');
    btnCad.className = 'btn btn-outline';
    btnCad.style.marginTop = '10px';
    btnCad.innerHTML = '<span class="btn-icon">➕</span> Cadastrar Produto';
    btnCad.onclick = () => trocarSubAbaControleImbel('cadastro');
    container.appendChild(btnCad);
}

function renderControleImbelCadastro() {
    const data = loadImbel();
    const container = document.getElementById('controleImbelCadastroContainer');
    container.innerHTML = '';

    const thStyle = 'padding:8px 12px;border:1px solid #ddd;background:#1e3a5f;color:#fff;font-size:.82rem;white-space:nowrap';
    const tdBase  = 'padding:8px 12px;border:1px solid #ddd;vertical-align:middle;font-size:.85rem';

    // Botões de ação no topo
    const topActions = document.createElement('div');
    topActions.style.cssText = 'display:flex;gap:10px;margin-bottom:16px;flex-wrap:wrap';
    const btnAdd = document.createElement('button');
    btnAdd.className = 'btn btn-primary';
    btnAdd.innerHTML = '<span class="btn-icon">➕</span> Adicionar Produto';
    btnAdd.onclick = () => {
        // Limpar campos do modal
        document.getElementById('imbel_prod_nome').value = '';
        document.getElementById('imbel_prod_codigo').value = '';
        document.getElementById('imbel_prod_qtd_inicial').value = '0';
        document.getElementById('imbel_prod_obs').value = '';
        const editField = document.getElementById('imbel_prod_edit_id'); if (editField) editField.value = '';
        document.getElementById('imbel_prod_salvar').textContent = 'Salvar';
        document.getElementById('imbel_prod_nome').focus();
        openImbelProdModal();
    };
    topActions.appendChild(btnAdd);
    container.appendChild(topActions);

    // Tabela de produtos
    const wrap = document.createElement('div');
    wrap.style.cssText = 'overflow-x:auto;background:#fff;border-radius:10px;padding:16px;box-shadow:0 1px 4px rgba(0,0,0,.08)';
    const tabela = document.createElement('table');
    tabela.style.cssText = 'width:100%;border-collapse:collapse;font-size:.85rem';
    tabela.innerHTML = `<thead><tr>
        <th style="${thStyle}">Nome</th>
        <th style="${thStyle}">Código</th>
        <th style="${thStyle}">Quantidade Inicial</th>
        <th style="${thStyle}">Observação</th>
        <th style="${thStyle}">Ações</th>
    </tr></thead><tbody></tbody>`;
    const tbody = tabela.querySelector('tbody');
    (data.produtos||[]).forEach((p, idx) => {
        const tr = document.createElement('tr');
        tr.style.background = idx % 2 === 0 ? '#fff' : '#f7f9fc';
        tr.innerHTML = `
            <td style="${tdBase}">${p.nome}</td>
            <td style="${tdBase}">${p.codigo||'-'}</td>
            <td style="${tdBase};text-align:center">${(p.quantidadeInicial||p.quantidadeInicial===0)?p.quantidadeInicial:'-'}</td>
            <td style="${tdBase}">${p.observacao||'-'}</td>
            <td style="${tdBase}"><button class="btn btn-outline" data-editid="${p.id}" style="margin-right:6px">Editar</button><button class="btn btn-outline" data-id="${p.id}">Remover</button></td>`;
        tbody.appendChild(tr);
    });
    wrap.appendChild(tabela);
    container.appendChild(wrap);

    // handlers
    document.getElementById('imbel_prod_salvar').onclick = function() {
        const editIdField = document.getElementById('imbel_prod_edit_id');
        const editId = editIdField ? editIdField.value : '';
        const nome = document.getElementById('imbel_prod_nome').value.trim().toUpperCase();
        const codigo = document.getElementById('imbel_prod_codigo').value.trim().toUpperCase();
        const quantidadeInicial = parseInt(document.getElementById('imbel_prod_qtd_inicial').value) || 0;
        const observacao = document.getElementById('imbel_prod_obs').value.trim().toUpperCase();
        if (!nome) { alert('Informe o nome do produto'); return; }
        data.produtos = data.produtos || [];
        if (editId) {
            const prod = data.produtos.find(p => p.id === editId);
            if (prod) {
                prod.nome = nome; prod.codigo = codigo; prod.observacao = observacao; prod.quantidadeInicial = quantidadeInicial;
                saveImbel(data);
                // reset edit state
                editIdField.value = '';
                document.getElementById('imbel_prod_salvar').textContent = 'Salvar';
                closeImbelProdModal();
                renderControleImbelCadastro();
                renderControleImbelEstoque();
                return;
            }
        }
        const novo = { id: 'p' + Date.now(), nome, codigo, observacao, quantidadeInicial };
        data.produtos.push(novo);
        saveImbel(data);
        closeImbelProdModal();
        renderControleImbelCadastro();
    };

    tbody.querySelectorAll('button[data-id]').forEach(b => b.onclick = function(){
        const id = this.getAttribute('data-id');
        if (!confirm('Remover produto?')) return;
        data.produtos = (data.produtos||[]).filter(p => p.id !== id);
        // também remover movimentações relacionadas
        data.movimentacoes = (data.movimentacoes||[]).filter(m => m.produtoId !== id);
        saveImbel(data);
        renderControleImbelCadastro();
    });

    tbody.querySelectorAll('button[data-editid]').forEach(b => b.onclick = function(){
        const id = this.getAttribute('data-editid');
        const prod = (data.produtos||[]).find(p => p.id === id);
        if (!prod) return;
        // preencher formulário modal
        document.getElementById('imbel_prod_nome').value = prod.nome || '';
        document.getElementById('imbel_prod_codigo').value = prod.codigo || '';
        document.getElementById('imbel_prod_qtd_inicial').value = prod.quantidadeInicial || 0;
        document.getElementById('imbel_prod_obs').value = prod.observacao || '';
        const editField = document.getElementById('imbel_prod_edit_id'); if (editField) editField.value = id;
        document.getElementById('imbel_prod_salvar').textContent = 'Atualizar';
        openImbelProdModal();
        document.getElementById('imbel_prod_nome').focus();
    });
}

function renderControleImbelMovimentacao() {
    const data = loadImbel();
    const container = document.getElementById('controleImbelMovContainer');
    container.innerHTML = '';

    // ---- Botões de ação (Add Movimentação / Limpar Tabela) ----
    const topActions = document.createElement('div');
    topActions.style.cssText = 'display:flex;gap:10px;margin-bottom:16px;flex-wrap:wrap';
    
    const btnAdd = document.createElement('button');
    btnAdd.className = 'btn btn-primary';
    btnAdd.innerHTML = '<span class="btn-icon">➕</span> Adicionar Movimentação';
    btnAdd.onclick = () => {
        // Limpar campos do modal
        document.getElementById('imbel_mov_prod').value = '';
        document.getElementById('imbel_mov_tipo').value = 'Entrada';
        const hoje = new Date().toISOString().slice(0,10);
        document.getElementById('imbel_mov_data').value = hoje;
        document.getElementById('imbel_mov_qtd').value = '1';
        document.getElementById('imbel_mov_dest').value = '';
        document.getElementById('imbel_mov_cpf').value = '';
        document.getElementById('imbel_mov_valor').value = '';
        document.getElementById('imbel_mov_endereco').value = '';
        document.getElementById('imbel_mov_tel').value = '';
        document.getElementById('imbel_mov_email').value = '';
        document.getElementById('imbel_mov_obs').value = '';
        const editField = document.getElementById('imbel_mov_edit_id'); if (editField) editField.value = '';
        document.getElementById('imbel_mov_salvar').textContent = 'Registrar Movimentação';
        document.getElementById('imbel_mov_prod').focus();
        openImbelMovModal();
    };
    topActions.appendChild(btnAdd);
    
    const btnClearTable = document.createElement('button');
    btnClearTable.className = 'btn btn-outline';
    btnClearTable.innerHTML = '<span class="btn-icon">🧹</span> Limpar Tabela';
    btnClearTable.onclick = function(){
        if (!confirm('Deseja apagar TODAS as movimentações? Esta ação é irreversível.')) return;
        data.movimentacoes = [];
        saveImbel(data);
        mostrarNotificacao('Todas as movimentações foram removidas.', 'success');
        renderControleImbelMovimentacao();
        renderControleImbelEstoque();
        renderControleImbelCadastro();
    };
    topActions.appendChild(btnClearTable);
    container.appendChild(topActions);

    // ---- Filtros ----
    const filtrosWrap = document.createElement('div');
    filtrosWrap.style.cssText = 'display:flex;flex-wrap:wrap;gap:8px;margin:8px 0 12px;align-items:center';
    filtrosWrap.innerHTML = `
        <select id="imbel_filter_prod" style="padding:6px 8px;border:1px solid #ddd;border-radius:6px;font-size:.85rem">
            <option value="">Todos os produtos</option>
        </select>
        <select id="imbel_filter_tipo" style="padding:6px 8px;border:1px solid #ddd;border-radius:6px;font-size:.85rem">
            <option value="">Todos os tipos</option>
            <option value="ENTRADA">Entrada</option>
            <option value="SAIDA">Saída</option>
        </select>
        <label style="font-size:.8rem;color:#444">De</label>
        <input type="date" id="imbel_filter_date_start" style="padding:6px;border:1px solid #ddd;border-radius:6px" />
        <label style="font-size:.8rem;color:#444">Até</label>
        <input type="date" id="imbel_filter_date_end" style="padding:6px;border:1px solid #ddd;border-radius:6px" />
        <input type="text" id="imbel_filter_dest" placeholder="Destinatário" style="padding:6px;border:1px solid #ddd;border-radius:6px;min-width:160px" />
        <input type="text" id="imbel_filter_cpf" placeholder="CPF/CNPJ" style="padding:6px;border:1px solid #ddd;border-radius:6px;min-width:140px" />
        <label style="display:flex;align-items:center;gap:6px;margin-left:auto"><input type="checkbox" id="imbel_filter_pago"/> Somente pagos</label>
        <label style="display:flex;align-items:center;gap:6px"><input type="checkbox" id="imbel_filter_entregue_only"/> Somente entregues</label>
        <label style="display:flex;align-items:center;gap:6px"><input type="checkbox" id="imbel_filter_fi_only"/> Somente FI</label>
        <button id="imbel_filter_reset" class="btn btn-outline" style="padding:6px 10px">Limpar filtros</button>
    `;
    container.appendChild(filtrosWrap);

    // popular opções de produto no filtro
    const selFilterProd = filtrosWrap.querySelector('#imbel_filter_prod');
    (data.produtos||[]).forEach(p=>{
        const o = document.createElement('option'); o.value = p.id; o.textContent = p.nome; selFilterProd.appendChild(o);
    });

    // ---- Tabela histórico ----
    const tabelaWrap = document.createElement('div');
    tabelaWrap.style.cssText = 'overflow-x:auto;background:#fff;border-radius:10px;padding:16px;box-shadow:0 1px 4px rgba(0,0,0,.08)';

    const thStyle = 'padding:8px 10px;border:1px solid #ddd;background:#1e3a5f;color:#fff;font-size:.8rem;white-space:nowrap;text-align:center';
    const tabela = document.createElement('table');
    tabela.style.cssText = 'width:100%;border-collapse:collapse;font-size:.78rem';
    tabela.innerHTML = `<thead><tr>
        <th style="${thStyle}">ID</th>
        <th style="${thStyle}">Descrição</th>
        <th style="${thStyle}">Tipo</th>
        <th style="${thStyle}">Data</th>
        <th style="${thStyle}">Quant.</th>
        <th style="${thStyle}">Destinatário</th>
        <th style="${thStyle}">CPF/CNPJ</th>
        <th style="${thStyle}">Observações</th>
        <th style="${thStyle}">Valor</th>
        <th style="${thStyle}">Endereço</th>
        <th style="${thStyle}">Telefone</th>
        <th style="${thStyle}">E-mail</th>
        <th style="${thStyle}">Pagamento</th>
        <th style="${thStyle}">Entregue</th>
        <th style="${thStyle}">FI</th>
        <th style="${thStyle}">Ações</th>
    </tr></thead><tbody></tbody>`;

    const tbody = tabela.querySelector('tbody');
    const tdStyle = 'padding:6px 8px;border:1px solid #ddd;vertical-align:middle;white-space:normal;word-break:break-word;max-width:260px';
    const tdCenter = tdStyle + ';text-align:center';

    // função para popular tbody com filtros
    function populateTbody() {
        tbody.innerHTML = '';
        const all = (data.movimentacoes||[]).slice().reverse();
        const fProd = (document.getElementById('imbel_filter_prod')||{value:''}).value;
        const fTipo = (document.getElementById('imbel_filter_tipo')||{value:''}).value;
        const fStart = (document.getElementById('imbel_filter_date_start')||{value:''}).value;
        const fEnd = (document.getElementById('imbel_filter_date_end')||{value:''}).value;
        const fDest = (document.getElementById('imbel_filter_dest')||{value:''}).value.trim().toUpperCase();
        const fCpf = (document.getElementById('imbel_filter_cpf')||{value:''}).value.trim().toUpperCase();
        const fPago = !!(document.getElementById('imbel_filter_pago') && document.getElementById('imbel_filter_pago').checked);
        const fEnt = !!(document.getElementById('imbel_filter_entregue_only') && document.getElementById('imbel_filter_entregue_only').checked);
        const fFi = !!(document.getElementById('imbel_filter_fi_only') && document.getElementById('imbel_filter_fi_only').checked);

        let rowIndex = 0;
        all.forEach((m, idx) => {
            if (fProd && String(m.produtoId) !== String(fProd)) return;
            const tipoNorm = (m.tipo||'').toString().toUpperCase();
            if (fTipo) {
                if (fTipo === 'ENTRADA' && tipoNorm !== 'ENTRADA') return;
                if (fTipo === 'SAIDA' && tipoNorm !== 'SAIDA' && tipoNorm !== 'SAÍDA') return;
            }
            const mDateNorm = parseDateToYYYYMMDD(m.data) || '';
            if (fStart && mDateNorm && mDateNorm < fStart) return;
            if (fEnd && mDateNorm && mDateNorm > fEnd) return;
            if (fDest && !(m.destinatario||'').toString().toUpperCase().includes(fDest)) return;
            if (fCpf && !(m.cpfCnpj||'').toString().toUpperCase().includes(fCpf)) return;
            if (fPago && (m.pagamento||'').toString().toUpperCase() !== 'SIM') return;
            if (fEnt && (m.entregue||'').toString().toUpperCase() !== 'SIM') return;
            if (fFi && (m.fi||'').toString().toUpperCase() !== 'SIM') return;

            const produto = (data.produtos||[]).find(p => p.id === m.produtoId) || {nome: m.descricao || '-'};
            const dataFmt = formatDateToDDMMYYYY(m.data);
            const valorFmt = m.valor ? 'R$ ' + Number(m.valor).toLocaleString('pt-BR', {minimumFractionDigits:2}) : '-';
            const seqNum = (data.movimentacoes.length - idx);

            const tr = document.createElement('tr');
            // cor padrão alternada
            let bg = rowIndex % 2 === 0 ? '#fff' : '#f7f9fc';
            // destacar linhas cujo destinatário contenha DVMK (case-insensitive)
            try {
                if ((m.destinatario||'').toString().toUpperCase().includes('DVMK')) {
                    bg = '#fff3cd'; // tom claro amarelo
                }
            } catch(e) {}
            tr.style.background = bg;

            // decidir se mostramos checkboxes (não mostrar para Entradas)
            const tipoUpper = ((m.tipo||'').toString().toUpperCase());
            let pagCell = `<td style="${tdCenter}"><input type=\"checkbox\" class=\"imbel_table_chk_pag\" data-id=\"${m.id}\" ${((m.pagamento||'').toString().toUpperCase()==='SIM')? 'checked' : ''}></td>`;
            let entCell = `<td style="${tdCenter}"><input type=\"checkbox\" class=\"imbel_table_chk_ent\" data-id=\"${m.id}\" ${((m.entregue||'').toString().toUpperCase()==='SIM')? 'checked' : ''}></td>`;
            let fiCell = `<td style="${tdCenter}"><input type=\"checkbox\" class=\"imbel_table_chk_fi\" data-id=\"${m.id}\" ${((m.fi||'').toString().toUpperCase()==='SIM')? 'checked' : ''}></td>`;
            if (tipoUpper === 'ENTRADA' || ((m.destinatario||'').toString().toUpperCase().includes('DVMK'))) {
                pagCell = `<td style="${tdCenter}"></td>`;
                entCell = `<td style="${tdCenter}"></td>`;
                fiCell = `<td style="${tdCenter}"></td>`;
            }

            tr.innerHTML = `
                <td style="${tdCenter};white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:70px">${seqNum}</td>
                <td style="${tdStyle}">${produto.nome}</td>
                <td style="${tdCenter};white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:110px"><span style="padding:2px 8px;border-radius:12px;font-size:.75rem;font-weight:600;color:${(m.tipo||'').toString().toUpperCase()==='ENTRADA' ? '#155724' : ((m.tipo||'').toString().toUpperCase()==='SAIDA' || (m.tipo||'').toString().toUpperCase()==='SAÍDA' ? '#721c24' : '#333')};">${((m.tipo||'').toString().toUpperCase()==='ENTRADA' ? 'Entrada' : (((m.tipo||'').toString().toUpperCase()==='SAIDA' || (m.tipo||'').toString().toUpperCase()==='SAÍDA') ? 'Saída' : (m.tipo||'-')))}</span></td>
                <td style="${tdCenter};white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:110px">${dataFmt}</td>
                <td style="${tdCenter}">${m.quantidade||'-'}</td>
                <td style="${tdStyle}">${m.destinatario||'-'}</td>
                <td style="${tdStyle};white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:140px">${m.cpfCnpj||'-'}</td>
                <td style="${tdStyle}">${m.observacoes||'-'}</td>
                <td style="${tdCenter};white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:120px">${valorFmt}</td>
                <td style="${tdStyle}">${m.endereco||'-'}</td>
                <td style="${tdStyle}">${m.telefone||'-'}</td>
                <td style="${tdStyle}">${m.email||'-'}</td>
                ${pagCell}
                ${entCell}
                ${fiCell}
                <td style="${tdCenter}"><button class="btn btn-outline" style="padding:3px 8px;font-size:.75rem;margin-right:6px" data-editmov="${m.id}">Editar</button><button class="btn btn-outline" style="padding:3px 8px;font-size:.75rem" data-delid="${m.id}">🗑️</button></td>
            `;
            tbody.appendChild(tr);
            rowIndex++;
        });

        // rebind handlers
        tbody.querySelectorAll('button[data-delid]').forEach(btn => {
            btn.onclick = function(){
                const id = this.getAttribute('data-delid');
                if (!confirm('Excluir esta movimentação?')) return;
                data.movimentacoes = (data.movimentacoes||[]).filter(m => m.id !== id);
                saveImbel(data);
                mostrarNotificacao('Movimentação excluída.', 'success');
                populateTbody();
                renderControleImbelEstoque();
            };
        });

        tbody.querySelectorAll('button[data-editmov]').forEach(btn => {
            btn.onclick = function(){
                const id = this.getAttribute('data-editmov');
                const mov = (data.movimentacoes||[]).find(m => m.id === id);
                if (!mov) return;
                const hoje = new Date().toISOString().slice(0,10);
                document.getElementById('imbel_mov_prod').value = mov.produtoId || '';
                const tipoLower = (mov.tipo||'').toString().toLowerCase();
                if (tipoLower === 'entrada') document.getElementById('imbel_mov_tipo').value = 'Entrada';
                else if (tipoLower === 'saída' || tipoLower === 'saida') document.getElementById('imbel_mov_tipo').value = 'Saída';
                else document.getElementById('imbel_mov_tipo').value = mov.tipo || 'Entrada';
                document.getElementById('imbel_mov_data').value = mov.data || hoje;
                document.getElementById('imbel_mov_qtd').value = mov.quantidade || 1;
                document.getElementById('imbel_mov_dest').value = mov.destinatario || '';
                document.getElementById('imbel_mov_cpf').value = formatCpfCnpjMask(mov.cpfCnpj || '');
                document.getElementById('imbel_mov_valor').value = (Number(mov.valor) > 0)
                    ? Number(mov.valor).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 })
                    : '';
                document.getElementById('imbel_mov_endereco').value = mov.endereco || '';
                document.getElementById('imbel_mov_tel').value = formatPhoneMask(mov.telefone || '');
                document.getElementById('imbel_mov_email').value = mov.email || '';
                document.getElementById('imbel_mov_obs').value = mov.observacoes || '';
                const editField = document.getElementById('imbel_mov_edit_id'); if (editField) editField.value = id;
                document.getElementById('imbel_mov_salvar').textContent = 'Atualizar Movimentação';
                openImbelMovModal();
                document.getElementById('imbel_mov_prod').focus();
            };
        });

        // bind checkboxes in table (pagamento / entregue / fi)
        tbody.querySelectorAll('input.imbel_table_chk_pag').forEach(cb => cb.onchange = function(){
            const id = this.getAttribute('data-id');
            const mov = (data.movimentacoes||[]).find(m=>m.id===id);
            if (!mov) return; mov.pagamento = this.checked ? 'SIM' : 'NÃO'; saveImbel(data); mostrarNotificacao('Pagamento atualizado.', 'success'); renderControleImbelEstoque();
        });
        tbody.querySelectorAll('input.imbel_table_chk_ent').forEach(cb => cb.onchange = function(){
            const id = this.getAttribute('data-id');
            const mov = (data.movimentacoes||[]).find(m=>m.id===id);
            if (!mov) return; mov.entregue = this.checked ? 'SIM' : 'NÃO'; saveImbel(data); mostrarNotificacao('Status de entrega atualizado.', 'success'); renderControleImbelEstoque();
        });
        tbody.querySelectorAll('input.imbel_table_chk_fi').forEach(cb => cb.onchange = function(){
            const id = this.getAttribute('data-id');
            const mov = (data.movimentacoes||[]).find(m=>m.id===id);
            if (!mov) return; mov.fi = this.checked ? 'SIM' : 'NÃO'; saveImbel(data); mostrarNotificacao('FI atualizado.', 'success'); renderControleImbelEstoque();
        });
    }

    // conectar eventos dos filtros
    ['imbel_filter_prod','imbel_filter_tipo','imbel_filter_date_start','imbel_filter_date_end','imbel_filter_dest','imbel_filter_cpf','imbel_filter_pago','imbel_filter_entregue_only','imbel_filter_fi_only'].forEach(id=>{
        const el = document.getElementById(id);
        if (el) el.addEventListener('change', populateTbody);
        if (el && el.addEventListener && el.tagName === 'INPUT') el.addEventListener('input', populateTbody);
    });
    const btnResetFiltros = document.getElementById('imbel_filter_reset'); if (btnResetFiltros) btnResetFiltros.onclick = function(){
        document.getElementById('imbel_filter_prod').value = '';
        document.getElementById('imbel_filter_tipo').value = '';
        document.getElementById('imbel_filter_date_start').value = '';
        document.getElementById('imbel_filter_date_end').value = '';
        document.getElementById('imbel_filter_dest').value = '';
        document.getElementById('imbel_filter_cpf').value = '';
        document.getElementById('imbel_filter_pago').checked = false;
        document.getElementById('imbel_filter_entregue_only').checked = false;
        document.getElementById('imbel_filter_fi_only').checked = false;
        populateTbody();
    };

    // preencher tabela inicialmente
    populateTbody();

    // Preencher select de produtos do modal (se houver produtos)
    const selec = document.getElementById('imbel_mov_prod');
    if (selec) {
        selec.innerHTML = '<option value="">— selecione o produto —</option>';
        (data.produtos||[]).forEach(p => {
            const opt = document.createElement('option');
            opt.value = p.id;
            opt.textContent = p.nome + (p.codigo ? ' ('+p.codigo+')' : '');
            selec.appendChild(opt);
        });
    }

    bindImbelMovInputMasks();

    tabelaWrap.appendChild(tabela);
    container.appendChild(tabelaWrap);

    // ---- Handler: Salvar (no modal) ----
    document.getElementById('imbel_mov_salvar').onclick = function(){
        const hoje = new Date().toISOString().slice(0,10);
        const produtoId = document.getElementById('imbel_mov_prod').value;
        const tipo = document.getElementById('imbel_mov_tipo').value;
        const quantidade = Number(document.getElementById('imbel_mov_qtd').value) || 0;
        const dataStr = document.getElementById('imbel_mov_data').value || hoje;
        const destinatario = document.getElementById('imbel_mov_dest').value.trim();
        const cpfCnpj = document.getElementById('imbel_mov_cpf').value.trim();
        const valor = parseCurrencyBRLToNumber(document.getElementById('imbel_mov_valor').value);
        // pagamento/entregue/fi are managed in the table; default to not paid/not delivered/not FI on new records
        const pagamento = false;
        const endereco = document.getElementById('imbel_mov_endereco').value.trim();
        const telefone = document.getElementById('imbel_mov_tel').value.trim();
        const email = document.getElementById('imbel_mov_email').value.trim();
        const entregue = false;
        const fi = false;
        const obs = document.getElementById('imbel_mov_obs').value.trim();
        const editId = (document.getElementById('imbel_mov_edit_id') || {value:''}).value || '';

        if (!produtoId) { mostrarNotificacao('Selecione um produto.', 'warning'); return; }
        if (quantidade <= 0) { mostrarNotificacao('Informe uma quantidade válida.', 'warning'); return; }

        // Bloquear saída se saldo insuficiente (aceita variações maiúsculas/minúsculas/sem acento)
        const tipoNorm = (tipo||'').toString().toLowerCase();
        if (tipoNorm === 'saída' || tipoNorm === 'saida') {
            const saldos = {};
            (data.movimentacoes||[]).forEach(m2 => {
                const q = Number(m2.quantidade)||0;
                const t = (m2.tipo||'').toString().toLowerCase();
                const s = (t === 'entrada') ? 1 : (t === 'saída' || t === 'saida' ? -1 : 0);
                saldos[m2.produtoId] = (saldos[m2.produtoId]||0) + s * q;
            });
            const saldoAtual = saldos[produtoId] || 0;
            if (quantidade > saldoAtual) {
                mostrarNotificacao(`Saldo insuficiente. Saldo atual: ${saldoAtual} un.`, 'error');
                return;
            }
        }

        data.movimentacoes = data.movimentacoes || [];
        if (editId) {
            const mov = data.movimentacoes.find(m => m.id === editId);
            if (mov) {
                mov.produtoId = produtoId;
                mov.tipo = (tipo||'').toString().toUpperCase();
                mov.quantidade = quantidade;
                mov.data = dataStr;
                mov.destinatario = (destinatario||'').toUpperCase();
                mov.cpfCnpj = (cpfCnpj||'').toUpperCase();
                mov.valor = valor;
                mov.endereco = (endereco||'').toUpperCase();
                mov.telefone = (telefone||'').toUpperCase();
                mov.email = (email||'').toUpperCase();
            // preserve existing pagamento/entregue/fi values (these are edited via table checkboxes)
                mov.observacoes = (obs||'').toUpperCase();
                saveImbel(data);
                mostrarNotificacao('Movimentação atualizada!', 'success');
                // reset edit state
                const editField = document.getElementById('imbel_mov_edit_id'); if (editField) editField.value = '';
                document.getElementById('imbel_mov_salvar').textContent = 'Registrar Movimentação';
                closeImbelMovModal();
                renderControleImbelMovimentacao();
                renderControleImbelEstoque();
                return;
            }
        }

        data.movimentacoes.push({
            id: 'm' + Date.now(), produtoId, tipo: (tipo||'').toString().toUpperCase(), quantidade, data: dataStr,
            destinatario: (destinatario||'').toUpperCase(), cpfCnpj: (cpfCnpj||'').toUpperCase(), valor,
            pagamento: 'NÃO', endereco: (endereco||'').toUpperCase(),
            telefone: (telefone||'').toUpperCase(), email: (email||'').toUpperCase(), entregue: 'NÃO',
            fi: 'NÃO', observacoes: (obs||'').toUpperCase()
        });
        saveImbel(data);
        mostrarNotificacao('Movimentação registrada!', 'success');
        closeImbelMovModal();
        renderControleImbelMovimentacao();
        renderControleImbelEstoque();
    };
}

// Inicializar controle IMBEL (não altera outros dados)
try { initControleImbel(); } catch(e) {}

// Abre a sub-aba de Cadastro e preenche o formulário para editar o produto indicado
function editarProdutoPorId(id) {
    try {
        const data = loadImbel();
        const prod = (data.produtos||[]).find(p => p.id === id);
        if (!prod) { mostrarNotificacao('Produto não encontrado para edição.', 'error'); return; }
        // Preencher modal
        document.getElementById('imbel_prod_nome').value = prod.nome || '';
        document.getElementById('imbel_prod_codigo').value = prod.codigo || '';
        document.getElementById('imbel_prod_qtd_inicial').value = prod.quantidadeInicial || 0;
        document.getElementById('imbel_prod_obs').value = prod.observacao || '';
        const editField = document.getElementById('imbel_prod_edit_id'); if (editField) editField.value = id;
        document.getElementById('imbel_prod_salvar').textContent = 'Atualizar';
        openImbelProdModal();
        document.getElementById('imbel_prod_nome').focus();
    } catch (e) {
        console.error('editarProdutoPorId erro:', e);
    }
}

// ========================================
// EXPORTAR / IMPORTAR IMBEL (Produtos + Saldos)
// ========================================

function exportarImbel() {
    const sep = ';';
    const data = loadImbel();
    const produtos = data.produtos || [];
    const movimentacoes = data.movimentacoes || [];

    // Calcular entradas/saidas por produto
    const totE = {};
    const totS = {};
    movimentacoes.forEach(m => {
        if (!m.produtoId) return;
        const q = Number(m.quantidade) || 0;
        const t = (m.tipo||'').toString().toLowerCase();
        if (t === 'entrada') totE[m.produtoId] = (totE[m.produtoId]||0) + q;
        else if (t === 'saída' || t === 'saida') totS[m.produtoId] = (totS[m.produtoId]||0) + q;
    });

    let csv = `ID${sep}DESCRICAO${sep}OBSERVACAO${sep}QUANTIDADE_INICIAL${sep}ENTRADA${sep}SAIDA${sep}ESTOQUE_ATUAL\n`;
    produtos.forEach(p => {
        const entrada = totE[p.id] || 0;
        const saida = totS[p.id] || 0;
        const inicial = Number(p.quantidadeInicial) || 0;
        const saldo = inicial + entrada - saida;
        const linha = `${p.id || ''}${sep}${(p.nome||'').toString().replace(/\n/g,' ')}${sep}${(p.observacao||p.codigo||'').toString().replace(/\n/g,' ')}${sep}${inicial}${sep}${entrada}${sep}${saida}${sep}${saldo}`;
        csv += linha + '\n';
    });

    const blob = new Blob(['\ufeff' + csv], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    const url = URL.createObjectURL(blob);
    const dataAtual = new Date().toISOString().split('T')[0];
    link.setAttribute('href', url);
    link.setAttribute('download', `imbel_estoque_${dataAtual}.csv`);
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    mostrarNotificacao('IMBEL exportado com sucesso!', 'success');
}

function importarImbel(event) {
    const file = event.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const conteudo = e.target.result;
            const linhas = conteudo.split(/\r?\n/).filter(l => l.trim());
            if (linhas.length < 2) { mostrarNotificacao('Arquivo vazio ou sem dados!', 'error'); return; }

            const data = loadImbel();
            data.produtos = data.produtos || [];
            data.movimentacoes = data.movimentacoes || [];

            let importados = 0;
            let erros = [];

            for (let i = 1; i < linhas.length; i++) {
                const linha = linhas[i].trim();
                if (!linha) continue;
                const col = parseCsvLinha(linha);
                // Espera colunas: ID;DESCRICAO;OBSERVACAO;ENTRADA;SAIDA;ESTOQUE_ATUAL
                if (col.length < 3) { erros.push(`Linha ${i+1}: formato inválido`); continue; }

                const descricao = (col[1]||'').trim().toUpperCase();
                if (!descricao) { erros.push(`Linha ${i+1}: descrição vazia`); continue; }

                const observacao = (col[2]||'').trim().toUpperCase();
                const quantidadeInicial = parseInt((col[3]||'').replace(/[^0-9-]/g,'')) || 0;
                const entrada = parseInt((col[4]||'').replace(/[^0-9-]/g,'')) || 0;
                const saida = parseInt((col[5]||'').replace(/[^0-9-]/g,'')) || 0;

                // Verificar se produto já existe (por nome)
                let produto = data.produtos.find(p => (p.nome||'').toString().toUpperCase() === descricao);
                if (!produto) {
                    produto = { id: 'p' + Date.now() + i, nome: descricao, codigo: '', observacao, quantidadeInicial };
                    data.produtos.push(produto);
                } else {
                    // atualizar observacao se houver
                    produto.observacao = produto.observacao || observacao;
                    produto.quantidadeInicial = produto.quantidadeInicial || quantidadeInicial;
                }

                // Criar movimentações de entrada/saida somando valores (se informados)
                const hoje = new Date().toISOString().split('T')[0];
                if (entrada > 0) {
                    data.movimentacoes.push({ id: 'm' + Date.now() + i, produtoId: produto.id, tipo: 'ENTRADA', quantidade: entrada, data: hoje, destinatario: '', cpfCnpj: '', valor: 0, pagamento: '', endereco: '', telefone: '', email: '', entregue: 'NÃO', fi: '', observacoes: 'Importado - Entrada' });
                }
                if (saida > 0) {
                    data.movimentacoes.push({ id: 'm' + Date.now() + i + 1000, produtoId: produto.id, tipo: 'SAÍDA', quantidade: saida, data: hoje, destinatario: '', cpfCnpj: '', valor: 0, pagamento: '', endereco: '', telefone: '', email: '', entregue: 'NÃO', fi: '', observacoes: 'Importado - Saída' });
                }

                importados++;
            }

            if (importados > 0) {
                saveImbel(data);
                renderControleImbelCadastro();
                renderControleImbelMovimentacao();
                renderControleImbelEstoque();
            }

            event.target.value = '';

            if (erros.length > 0 && importados === 0) {
                mostrarNotificacao('Nenhum registro importado. Verifique o formato.', 'error');
                console.debug('Erros de importação IMBEL:', erros);
            } else if (erros.length > 0) {
                mostrarNotificacao(`${importados} registros importados. ${erros.length} linhas com erro.`, 'warning');
                console.debug('Erros de importação IMBEL:', erros);
            } else {
                mostrarNotificacao(`${importados} registros importados com sucesso!`, 'success');
            }

        } catch (err) {
            console.error('Erro ao importar IMBEL:', err);
            mostrarNotificacao('Erro ao processar o arquivo IMBEL. Verifique o formato.', 'error');
        }
    };
    reader.readAsText(file, 'UTF-8');
}

// ========================================
// EXPORT/IMPORT - IMBEL Cadastro (Produtos)
// ========================================

function exportarImbelCadastro() {
    const sep = ';';
    const data = loadImbel();
    const produtos = data.produtos || [];
    let csv = `ID${sep}NOME${sep}CODIGO${sep}OBSERVACAO${sep}QUANTIDADE_INICIAL\n`;
    produtos.forEach(p => {
        const linha = `${p.id||''}${sep}${(p.nome||'').toString().replace(/\n/g,' ')}${sep}${(p.codigo||'').toString().replace(/\n/g,' ')}${sep}${(p.observacao||'').toString().replace(/\n/g,' ')}${sep}${(Number(p.quantidadeInicial)||0)}`;
        csv += linha + '\n';
    });
    const blob = new Blob(['\ufeff' + csv], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    const url = URL.createObjectURL(blob);
    const dataAtual = new Date().toISOString().split('T')[0];
    link.setAttribute('href', url);
    link.setAttribute('download', `imbel_produtos_${dataAtual}.csv`);
    link.style.visibility = 'hidden'; document.body.appendChild(link); link.click(); document.body.removeChild(link);
    mostrarNotificacao('Produtos IMBEL exportados!', 'success');
}

function importarImbelCadastro(event) {
    const file = event.target.files[0]; if (!file) return;
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const conteudo = e.target.result;
            const linhas = conteudo.split(/\r?\n/).filter(l => l.trim());
            if (linhas.length < 2) { mostrarNotificacao('Arquivo vazio ou sem dados!', 'error'); return; }
            const data = loadImbel(); data.produtos = data.produtos || [];
            let importados = 0; let erros = [];
            for (let i=1;i<linhas.length;i++) {
                const linha = linhas[i].trim(); if (!linha) continue;
                const col = parseCsvLinha(linha);
                if (col.length < 2) { erros.push(`Linha ${i+1}: formato inválido`); continue; }
                const nome = (col[1]||'').trim().toUpperCase();
                if (!nome) { erros.push(`Linha ${i+1}: nome vazio`); continue; }
                const codigo = (col[2]||'').trim().toUpperCase();
                const observacao = (col[3]||'').trim().toUpperCase();
                const quantidadeInicial = parseInt((col[4]||'').replace(/[^0-9-]/g,'')) || 0;
                let produto = data.produtos.find(p => (p.nome||'').toString().toUpperCase() === nome);
                if (!produto) {
                    produto = { id: 'p' + Date.now() + i, nome, codigo, observacao, quantidadeInicial };
                    data.produtos.push(produto);
                } else {
                    produto.codigo = produto.codigo || codigo;
                    produto.observacao = produto.observacao || observacao;
                    produto.quantidadeInicial = produto.quantidadeInicial || quantidadeInicial;
                }
                importados++;
            }
            if (importados>0) { saveImbel(data); renderControleImbelCadastro(); renderControleImbelEstoque(); }
            event.target.value = '';
            if (erros.length>0 && importados===0) { mostrarNotificacao('Nenhum produto importado. Verifique o formato.', 'error'); console.debug('Erros IMBEL cadastro:',erros); }
            else if (erros.length>0) { mostrarNotificacao(`${importados} produtos importados. ${erros.length} linhas com erro.`, 'warning'); console.debug('Erros IMBEL cadastro:',erros); }
            else { mostrarNotificacao(`${importados} produtos importados com sucesso!`, 'success'); }
        } catch(err) { console.error('Erro importar cadastro IMBEL',err); mostrarNotificacao('Erro ao processar o arquivo.', 'error'); }
    };
    reader.readAsText(file,'UTF-8');
}

// ========================================
// EXPORT/IMPORT - IMBEL Movimentação
// ========================================

function exportarImbelMovimentacao() {
    const sep = ';';
    const data = loadImbel();
    const movs = data.movimentacoes || [];
    const produtos = (data.produtos||[]).reduce((acc,p)=>{acc[p.id]=p;return acc;},{})
    let csv = `ID${sep}PRODUTO_ID${sep}PRODUTO_NOME${sep}TIPO${sep}DATA${sep}QUANTIDADE${sep}DESTINATARIO${sep}CPF_CNPJ${sep}OBSERVACOES${sep}VALOR${sep}PAGAMENTO${sep}ENDERECO${sep}TELEFONE${sep}EMAIL${sep}ENTREGUE${sep}FI\n`;
    movs.forEach(m => {
        const prod = produtos[m.produtoId] || {};
        const dataCsv = formatDateToDDMMYYYY(m.data);
        const linha = `${m.id||''}${sep}${m.produtoId||''}${sep}${(prod.nome||'').toString().replace(/\n/g,' ')}${sep}${(m.tipo||'').toString()}${sep}${dataCsv}${sep}${m.quantidade||''}${sep}${(m.destinatario||'').toString().replace(/\n/g,' ')}${sep}${(m.cpfCnpj||'').toString()}${sep}${(m.observacoes||'').toString().replace(/\n/g,' ')}${sep}${(m.valor||'')}${sep}${(m.pagamento||'').toString()}${sep}${(m.endereco||'').toString().replace(/\n/g,' ')}${sep}${(m.telefone||'').toString()}${sep}${(m.email||'').toString()}${sep}${(m.entregue||'').toString()}${sep}${(m.fi||'').toString()}`;
        csv += linha + '\n';
    });
    const blob = new Blob(['\ufeff' + csv], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a'); const url = URL.createObjectURL(blob);
    const dataAtual = new Date().toISOString().split('T')[0];
    link.setAttribute('href', url); link.setAttribute('download', `imbel_movimentacoes_${dataAtual}.csv`);
    link.style.visibility='hidden'; document.body.appendChild(link); link.click(); document.body.removeChild(link);
    mostrarNotificacao('Movimentações exportadas!', 'success');
}

function importarImbelMovimentacao(event) {
    const file = event.target.files[0]; if (!file) return;
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const conteudo = e.target.result;
            const linhas = conteudo.split(/\r?\n/).filter(l=>l.trim());
            if (linhas.length < 2) { mostrarNotificacao('Arquivo vazio ou sem dados!', 'error'); return; }
            const data = loadImbel(); data.produtos = data.produtos || []; data.movimentacoes = data.movimentacoes || [];
            let importados = 0; let erros = [];
            for (let i=1;i<linhas.length;i++) {
                const linha = linhas[i].trim(); if (!linha) continue;
                const col = parseCsvLinha(linha);
                // Espera ao menos: PRODUTO_NOME ou PRODUTO_ID e TIPO e QUANTIDADE
                if (col.length < 6) { erros.push(`Linha ${i+1}: formato inválido`); continue; }
                const produtoIdCol = (col[1]||'').trim();
                const produtoNomeCol = (col[2]||'').trim().toUpperCase();
                const tipo = (col[3]||'').trim().toUpperCase() || 'ENTRADA';
                // Normalizar campo DATA: aceitar YYYY-MM-DD ou DD/MM/YYYY e converter para YYYY-MM-DD
                let dataRaw = (col[4]||'').trim();
                let dataStr = '';
                if (!dataRaw) {
                    dataStr = new Date().toISOString().split('T')[0];
                } else {
                    // YYYY-MM-DD
                    if (/^\d{4}-\d{2}-\d{2}$/.test(dataRaw)) {
                        dataStr = dataRaw;
                    } else if (/^\d{2}\/\d{2}\/\d{4}$/.test(dataRaw)) {
                        const parts = dataRaw.split('/');
                        dataStr = `${parts[2]}-${parts[1].padStart(2,'0')}-${parts[0].padStart(2,'0')}`;
                    } else {
                        // tentar parse geral e formatar para YYYY-MM-DD quando possível
                        const d = new Date(dataRaw);
                        if (!isNaN(d.getTime())) {
                            dataStr = d.toISOString().slice(0,10);
                        } else {
                            // fallback para data atual
                            dataStr = new Date().toISOString().split('T')[0];
                        }
                    }
                }
                const quantidade = parseInt((col[5]||'').replace(/[^0-9-]/g,'')) || 0;
                if (quantidade === 0) { erros.push(`Linha ${i+1}: quantidade inválida`); continue; }

                // Encontrar produto: por ID ou por nome
                let produto = null;
                if (produtoIdCol) produto = data.produtos.find(p => p.id === produtoIdCol);
                if (!produto && produtoNomeCol) produto = data.produtos.find(p => (p.nome||'').toString().toUpperCase() === produtoNomeCol);
                if (!produto && produtoNomeCol) {
                    produto = { id: 'p' + Date.now() + i, nome: produtoNomeCol, codigo: '', observacao: '' };
                    data.produtos.push(produto);
                }

                const destinatario = (col[6]||'').trim().toUpperCase();
                const cpfCnpj = (col[7]||'').trim().toUpperCase();
                const observacoes = (col[8]||'').trim().toUpperCase();
                // Normalizar campo VALOR: remover símbolos, tratar milhares (.) e decimais (',')
                let valor = 0;
                try {
                    let vs = (col[9]||'').toString().trim();
                    // remover símbolos de moeda e espaços
                    vs = vs.replace(/[^0-9,\.\-]/g, '');
                    if (vs === '') vs = '0';
                    // se contém ',' como separador decimal (BR), remover pontos de milhares e trocar ',' por '.'
                    if (vs.indexOf(',') !== -1) {
                        vs = vs.replace(/\./g, '');
                        vs = vs.replace(/,/g, '.');
                    } else {
                        // sem vírgula, apenas remover separadores de milhar (caso existam) e keep dot
                        vs = vs.replace(/\.(?=\d{3}(?:\.|$))/g, '');
                    }
                    valor = parseFloat(vs) || 0;
                } catch (errVal) { valor = 0; }
                const pagamento = (col[10]||'').trim().toUpperCase();
                const endereco = (col[11]||'').trim().toUpperCase();
                const telefone = (col[12]||'').trim().toUpperCase();
                const email = (col[13]||'').trim().toUpperCase();
                const entregue = (col[14]||'').trim().toUpperCase() || 'NÃO';
                const fi = (col[15]||'').trim().toUpperCase();

                const mov = { id: 'm' + Date.now() + i, produtoId: produto.id, tipo, quantidade, data: dataStr, destinatario, cpfCnpj, valor, pagamento, endereco, telefone, email, entregue, fi, observacoes };
                data.movimentacoes.push(mov);
                importados++;
            }
            if (importados>0) { saveImbel(data); renderControleImbelMovimentacao(); renderControleImbelEstoque(); renderControleImbelCadastro(); }
            event.target.value = '';
            if (erros.length>0 && importados===0) { mostrarNotificacao('Nenhuma movimentação importada. Verifique o formato.', 'error'); console.debug('Erros IMBEL mov:',erros); }
            else if (erros.length>0) { mostrarNotificacao(`${importados} movimentações importadas. ${erros.length} linhas com erro.`, 'warning'); console.debug('Erros IMBEL mov:',erros); }
            else { mostrarNotificacao(`${importados} movimentações importadas com sucesso!`, 'success'); }
        } catch(err) { console.error('Erro importar movimentações IMBEL',err); mostrarNotificacao('Erro ao processar o arquivo.', 'error'); }
    };
    reader.readAsText(file,'UTF-8');
}


function obterItensVendaNormalizados(venda) {
    if (Array.isArray(venda.items) && venda.items.length > 0) {
        return venda.items.map(it => ({
            produtoNome: it.produtoNome || venda.produtoNome || '',
            quantidade: Number(it.quantidade) || 0,
            valorTotal: typeof it.valorTotal === 'number' ? it.valorTotal : ((Number(it.valorUnitario) || 0) * (Number(it.quantidade) || 0))
        }));
    }
    return [{
        produtoNome: venda.produtoNome || '',
        quantidade: Number(venda.quantidade) || 0,
        valorTotal: typeof venda.valorTotal === 'number' ? venda.valorTotal : ((Number(venda.valorUnitario) || 0) * (Number(venda.quantidade) || 0))
    }];
}

function obterVendasDashboardFiltradas() {
    const preset = document.getElementById('dashboardRangePreset')?.value || '30';
    const startInput = document.getElementById('dashboardDateStart')?.value || '';
    const endInput = document.getElementById('dashboardDateEnd')?.value || '';

    const hoje = new Date();
    let inicio = null;
    let fim = new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate(), 23, 59, 59);

    if (preset === '7' || preset === '30' || preset === '90') {
        const dias = Number(preset);
        inicio = new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate() - (dias - 1), 0, 0, 0);
    } else if (preset === 'custom') {
        if (startInput) inicio = new Date(startInput + 'T00:00:00');
        if (endInput) fim = new Date(endInput + 'T23:59:59');
    }

    return (estoque.registroVendas || []).filter(v => {
        if (!v.data) return false;
        const d = new Date(v.data);
        if (isNaN(d.getTime())) return false;
        if (inicio && d < inicio) return false;
        if (fim && d > fim) return false;
        return true;
    });
}

function atualizarFiltroDashboardPeriodo() {
    const preset = document.getElementById('dashboardRangePreset');
    const startEl = document.getElementById('dashboardDateStart');
    const endEl = document.getElementById('dashboardDateEnd');
    if (!preset || !startEl || !endEl) return;

    const custom = preset.value === 'custom';
    startEl.style.display = custom ? '' : 'none';
    endEl.style.display = custom ? '' : 'none';

    if (!custom) {
        startEl.value = '';
        endEl.value = '';
    }
    renderizarDashboard();
}

function renderizarDashboard() {
    const vendasFiltradas = obterVendasDashboardFiltradas();
    // Aplicar filtro de período (todos / mês / trimestre / ano)
    const periodo = document.getElementById('filtroDashboardPeriodo')?.value || 'todos';
    const agora = new Date();
    const vendasPeriodo = (vendasFiltradas || []).filter(v => {
        if (periodo === 'todos') return true;
        const d = new Date(v.data || 0);
        if (isNaN(d.getTime())) return false;
        if (periodo === 'mes') return d.getMonth() === agora.getMonth() && d.getFullYear() === agora.getFullYear();
        if (periodo === 'trimestre') return Math.floor(d.getMonth()/3) === Math.floor(agora.getMonth()/3) && d.getFullYear() === agora.getFullYear();
        if (periodo === 'ano') return d.getFullYear() === agora.getFullYear();
        return true;
    });
    const repsOrdem = ['ADES', 'FL', 'IMBEL', 'ISA', 'KOLTE', 'LC'];
    const vendasPorRep = {};
    repsOrdem.forEach(rep => { vendasPorRep[rep] = 0; });

    const produtoMap = new Map();
    let totalUnidades = 0;
    let totalFaturamento = 0;

    vendasPeriodo.forEach(venda => {
        const rep = (venda.representante || '').toUpperCase();
        const itens = obterItensVendaNormalizados(venda);
        itens.forEach(it => {
            if (!it.produtoNome) return;
            const registro = produtoMap.get(it.produtoNome) || {
                nome: it.produtoNome,
                quantidade: 0,
                valor: 0,
                vendasPorRep: {}
            };
            registro.quantidade += it.quantidade;
            registro.valor += it.valorTotal;
            registro.vendasPorRep[rep] = (registro.vendasPorRep[rep] || 0) + it.quantidade;
            produtoMap.set(it.produtoNome, registro);

            totalUnidades += it.quantidade;
            totalFaturamento += it.valorTotal;
            if (vendasPorRep[rep] === undefined) vendasPorRep[rep] = 0;
            vendasPorRep[rep] += it.quantidade;
        });
    });

    const dadosVendas = Array.from(produtoMap.values()).sort((a, b) => b.quantidade - a.quantidade);
    const dadosVendasAlpha = [...dadosVendas].sort((a, b) => a.nome.localeCompare(b.nome, 'pt-BR'));

    let melhorRep = '-';
    let maxVendas = 0;
    Object.entries(vendasPorRep).forEach(([rep, qtd]) => {
        if (qtd > maxVendas) {
            melhorRep = rep;
            maxVendas = qtd;
        }
    });

    const produtoTop = dadosVendas.length > 0 && dadosVendas[0].quantidade > 0
        ? dadosVendas[0].nome.substring(0, 25) + (dadosVendas[0].nome.length > 25 ? '...' : '')
        : '-';

    const dashTotalUnidadesEl = document.getElementById('dashTotalUnidades');
    const dashTotalFaturamentoEl = document.getElementById('dashTotalFaturamento');
    const dashMelhorRepEl = document.getElementById('dashMelhorRep');
    const dashProdutoTopEl = document.getElementById('dashProdutoTop');
    if (dashTotalUnidadesEl) dashTotalUnidadesEl.textContent = totalUnidades.toLocaleString('pt-BR');
    if (dashTotalFaturamentoEl) dashTotalFaturamentoEl.textContent = formatarMoedaValor(totalFaturamento);
    if (dashMelhorRepEl) dashMelhorRepEl.textContent = maxVendas > 0 ? `${melhorRep} (${maxVendas})` : '-';
    if (dashProdutoTopEl) dashProdutoTopEl.textContent = produtoTop;

    const tabelaQtd = document.getElementById('tabelaQtdProduto');
    if (tabelaQtd) {
        tabelaQtd.innerHTML = '';
        dadosVendasAlpha.forEach(item => {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td class="produto-nome">${item.nome}</td>
                <td class="cell-qtd">${item.quantidade.toLocaleString('pt-BR')}</td>
            `;
            tabelaQtd.appendChild(tr);
        });
        const trTotalQtd = document.createElement('tr');
        trTotalQtd.className = 'total-row';
        trTotalQtd.innerHTML = `
            <td class="produto-nome"><strong>Total Geral</strong></td>
            <td class="cell-qtd"><strong>${totalUnidades.toLocaleString('pt-BR')}</strong></td>
        `;
        tabelaQtd.appendChild(trTotalQtd);
    }

    const tabelaValor = document.getElementById('tabelaValorProduto');
    if (tabelaValor) {
        tabelaValor.innerHTML = '';
        dadosVendasAlpha.forEach(item => {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td class="produto-nome">${item.nome}</td>
                <td class="cell-valor">${formatarMoedaValor(item.valor)}</td>
            `;
            tabelaValor.appendChild(tr);
        });
        const trTotalValor = document.createElement('tr');
        trTotalValor.className = 'total-row';
        trTotalValor.innerHTML = `
            <td class="produto-nome"><strong>Total Geral</strong></td>
            <td class="cell-valor"><strong>${formatarMoedaValor(totalFaturamento)}</strong></td>
        `;
        tabelaValor.appendChild(trTotalValor);
    }

    const tabelaRep = document.getElementById('tabelaVendasRepBody');
    if (tabelaRep) {
        tabelaRep.innerHTML = '';
        const totaisPorRep = {};
        repsOrdem.forEach(rep => { totaisPorRep[rep] = 0; });

        dadosVendasAlpha.forEach(item => {
            const tr = document.createElement('tr');
            let html = `<td class="produto-nome">${item.nome}</td>`;
            repsOrdem.forEach(rep => {
                const qtd = item.vendasPorRep[rep] || 0;
                totaisPorRep[rep] += qtd;
                html += `<td class="${qtd === 0 ? 'cell-zero' : 'cell-qtd'}">${qtd > 0 ? qtd : '-'}</td>`;
            });
            html += `<td class="geral-venda"><strong>${item.quantidade}</strong></td>`;
            tr.innerHTML = html;
            tabelaRep.appendChild(tr);
        });

        const trTotalRep = document.createElement('tr');
        trTotalRep.className = 'total-row';
        let htmlTotal = `<td class="produto-nome"><strong>Total Geral</strong></td>`;
        repsOrdem.forEach(rep => {
            htmlTotal += `<td class="cell-qtd"><strong>${totaisPorRep[rep]}</strong></td>`;
        });
        htmlTotal += `<td class="geral-venda"><strong>${totalUnidades}</strong></td>`;
        trTotalRep.innerHTML = htmlTotal;
        tabelaRep.appendChild(trTotalRep);
    }
    try {
        if (typeof renderizarGraficoComissoes === 'function') renderizarGraficoComissoes();
    } catch (e) { console.warn('Erro ao renderizar gráfico de comissões', e); }
}

// ========================================
// FORMATAÇÃO
// ========================================

function formatarNumero(num) {
    if (num === 0) return '-';
    return num.toLocaleString('pt-BR');
}

function formatarMoedaValor(valor) {
    return valor.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
}

function formatarMoeda(input) {
    let valor = input.value.replace(/\D/g, '');
    valor = (parseInt(valor) / 100).toFixed(2);
    valor = valor.replace('.', ',');
    valor = valor.replace(/\B(?=(\d{3})+(?!\d))/g, '.');
    input.value = valor;
}

function converterMoedaParaNumero(valor) {
    if (!valor) return 0;
    return parseFloat(valor.replace(/\./g, '').replace(',', '.')) || 0;
}

// ========================================
// FUNÇÕES DOS MODAIS
// ========================================

function abrirModalProduto() {
    if (!requireAdminOrNotify()) return;
    // Abrir em modo de criar novo produto
    produtoEditandoId = null;
    document.getElementById('modalProduto').style.display = 'flex';
    document.getElementById('formProduto').reset();
    const categoriaSel = document.getElementById('produtoCategoria');
    if (categoriaSel) categoriaSel.value = '';
    // Restaurar título e botão padrão
    const header = document.querySelector('#modalProduto .modal-header h2');
    if (header) header.innerHTML = '<span>+</span> Adicionar Novo Produto';
    const submitBtn = document.querySelector('#modalProduto .modal-footer button[type="submit"]');
    if (submitBtn) submitBtn.textContent = 'Salvar Produto';
}

function abrirModalNovoProduto() {
    abrirModalProduto();
}

function abrirModalEditarProduto(produtoId) {
    if (!requireAdminOrNotify()) return;
    const produto = estoque.produtos.find(p => p.id === Number(produtoId));
    if (!produto) {
        mostrarNotificacao('Produto não encontrado!', 'error');
        return;
    }
    produtoEditandoId = produto.id;
    document.getElementById('modalProduto').style.display = 'flex';
    document.getElementById('formProduto').reset();
    document.getElementById('nomeProduto').value = produto.nome || '';
    document.getElementById('estoqueTotal').value = Number(produto.estoqueConsolidado || 0);
    const regra = (precificacao && precificacao[produto.nome]) ? precificacao[produto.nome] : {};
    const cat = (categoriaPorProduto && categoriaPorProduto[produto.nome]) || produto.categoria || '';
    const ncmInput = document.getElementById('produtoNCM');
    const catInput = document.getElementById('produtoCategoria');
    const ciInput = document.getElementById('produtoCI');
    const margemInput = document.getElementById('produtoMargemMin');
    const descInput = document.getElementById('produtoDescMax');
    const obsInput = document.getElementById('produtoObservacoes');
    if (ncmInput) ncmInput.value = produto.ncm || '';
    if (catInput) catInput.value = cat || '';
    if (ciInput) ciInput.value = Number(regra.ci ?? produto.ci ?? 0) || '';
    if (margemInput) margemInput.value = Number(regra.margemMinima ?? produto.margemMinima ?? 0) || '';
    if (descInput) descInput.value = Number(regra.descontoMaximo ?? produto.descontoMaximo ?? 0) || '';
    if (obsInput) obsInput.value = produto.observacoes || '';
    // Ajustar título e botão para modo edição
    const header = document.querySelector('#modalProduto .modal-header h2');
    if (header) header.innerHTML = '<span>✎</span> Editar Produto';
    const submitBtn = document.querySelector('#modalProduto .modal-footer button[type="submit"]');
    if (submitBtn) submitBtn.textContent = 'Salvar Alterações';
}

document.addEventListener('keydown', function(event) {
    if (event.key === 'Escape') {
        document.querySelectorAll('.modal').forEach(modal => {
            modal.style.display = 'none';
        });
        vendaEditandoId = null;
    }
});

// ========================================
// FUNÇÕES DE CRUD
// ========================================

function atualizarSelectsProdutos() {
    const selects = ['produtoDistribuicao', 'produtoVenda', 'produtoDevolucao', 'produtoVendaDet', 'filtroProduto', 'produtoDistDet', 'filtroDistribuicaoProduto'];
    
    selects.forEach(selectId => {
        const select = document.getElementById(selectId);
        if (select) {
            const valorAtual = select.value;
            const primeiraOpcao = (selectId === 'filtroProduto' || selectId === 'filtroDistribuicaoProduto') ? 'Todos' : 'Selecione um produto';
            select.innerHTML = `<option value="">${primeiraOpcao}</option>`;
            // Preencher em ordem alfabética
            const produtosOrdenados = [...estoque.produtos].sort((a, b) => a.nome.localeCompare(b.nome, 'pt-BR'));
            produtosOrdenados.forEach(produto => {
                const option = document.createElement('option');
                option.value = produto.id;
                option.textContent = produto.nome;
                select.appendChild(option);
            });
            
            select.value = valorAtual;
        }
    });

    // Atualizar selects dinâmicos de itens (se existirem)
    try {
        const produtosOrdenados = [...estoque.produtos].sort((a, b) => a.nome.localeCompare(b.nome, 'pt-BR'));
        document.querySelectorAll('.item-produto-dist, .item-produto-dev').forEach(sel => {
            const cur = sel.value;
            sel.innerHTML = '<option value="">Selecione um produto</option>';
            produtosOrdenados.forEach(produto => {
                const opt = document.createElement('option');
                opt.value = produto.id;
                opt.textContent = produto.nome;
                sel.appendChild(opt);
            });
            sel.value = cur;
        });
    } catch (e) {}
}

// Helpers para adicionar/remover linhas nos modais de distribuição/devolução
function adicionarItemDistribuicao(produtoId = null, quantidade = 1) {
    const container = document.getElementById('itensDistribuicaoContainer');
    if (!container) return;
    const row = document.createElement('div');
    row.className = 'item-dist-row';
    row.style.display = 'flex';
    row.style.gap = '8px';

    const sel = document.createElement('select');
    sel.className = 'item-produto-dist';
    sel.style.flex = '1';
    sel.innerHTML = '<option value="">Carregando...</option>';

    const inp = document.createElement('input');
    inp.type = 'number';
    inp.min = '1';
    inp.value = quantidade || 1;
    inp.className = 'item-quantidade-dist';
    inp.style.width = '96px';
    inp.placeholder = 'Qtd';

    const btnRem = document.createElement('button');
    btnRem.type = 'button';
    btnRem.className = 'btn btn-outline btn-sm';
    btnRem.style.width = '40px';
    btnRem.textContent = '✕';
    btnRem.title = 'Remover'
    btnRem.onclick = function() { row.remove(); };

    row.appendChild(sel);
    row.appendChild(inp);
    row.appendChild(btnRem);
    container.appendChild(row);

    // popular opções
    try {
        const produtosOrdenados = [...estoque.produtos].sort((a, b) => a.nome.localeCompare(b.nome, 'pt-BR'));
        sel.innerHTML = '<option value="">Selecione um produto</option>';
        produtosOrdenados.forEach(produto => {
            const opt = document.createElement('option');
            opt.value = produto.id;
            opt.textContent = produto.nome;
            sel.appendChild(opt);
        });
        if (produtoId) sel.value = String(produtoId);
    } catch (e) { sel.innerHTML = '<option value="">Nenhum produto</option>'; }
}

function adicionarItemDevolucao(produtoId = null, quantidade = 1) {
    const container = document.getElementById('itensDevolucaoContainer');
    if (!container) return;
    const row = document.createElement('div');
    row.className = 'item-dev-row';
    row.style.display = 'flex';
    row.style.gap = '8px';

    const sel = document.createElement('select');
    sel.className = 'item-produto-dev';
    sel.style.flex = '1';
    sel.innerHTML = '<option value="">Carregando...</option>';

    const inp = document.createElement('input');
    inp.type = 'number';
    inp.min = '1';
    inp.value = quantidade || 1;
    inp.className = 'item-quantidade-dev';
    inp.style.width = '96px';
    inp.placeholder = 'Qtd';

    const btnRem = document.createElement('button');
    btnRem.type = 'button';
    btnRem.className = 'btn btn-outline btn-sm';
    btnRem.style.width = '40px';
    btnRem.textContent = '✕';
    btnRem.title = 'Remover'
    btnRem.onclick = function() { row.remove(); };

    row.appendChild(sel);
    row.appendChild(inp);
    row.appendChild(btnRem);
    container.appendChild(row);

    // popular opções
    try {
        const produtosOrdenados = [...estoque.produtos].sort((a, b) => a.nome.localeCompare(b.nome, 'pt-BR'));
        sel.innerHTML = '<option value="">Selecione um produto</option>';
        produtosOrdenados.forEach(produto => {
            const opt = document.createElement('option');
            opt.value = produto.id;
            opt.textContent = produto.nome;
            sel.appendChild(opt);
        });
        if (produtoId) sel.value = String(produtoId);
    } catch (e) { sel.innerHTML = '<option value="">Nenhum produto</option>'; }
}

function atualizarSelectsRelatorios() {
    // Popular select de representantes
    const selRep = document.getElementById('filtroRelatoriosRep');
    if (selRep) {
        const atual = selRep.value;
        selRep.innerHTML = '<option value="">Todos</option>';
        estoque.representantes.forEach(rep => {
            const opt = document.createElement('option');
            opt.value = rep;
            opt.textContent = rep;
            selRep.appendChild(opt);
        });
        selRep.value = atual;
    }

    // Popular select de produtos
    const selProd = document.getElementById('filtroRelatoriosProduto');
    if (selProd) {
        const atualP = selProd.value;
        selProd.innerHTML = '<option value="">Todos</option>';
        const produtosOrdenados = [...estoque.produtos].sort((a, b) => a.nome.localeCompare(b.nome, 'pt-BR'));
        produtosOrdenados.forEach(produto => {
            const opt = document.createElement('option');
            opt.value = produto.id;
            opt.textContent = produto.nome;
            selProd.appendChild(opt);
        });
        selProd.value = atualP;
    }
}

function atualizarSelectDistribuicaoProduto() {
    const select = document.getElementById('filtroDistribuicaoProduto');
    if (select) {
        const valorAtual = select.value;
        select.innerHTML = '<option value="">Todos</option>';
        const produtosOrdenados = [...estoque.produtos].sort((a, b) => a.nome.localeCompare(b.nome, 'pt-BR'));
        produtosOrdenados.forEach(produto => {
            const option = document.createElement('option');
            option.value = produto.id;
            option.textContent = produto.nome;
            select.appendChild(option);
        });
        
        select.value = valorAtual;
    }
}

// ========================================
// ENTRADA DE ESTOQUE (IMBEL)
// ========================================

function abrirModalEntradaEstoque() {
    if (!requireAdminOrNotify()) return;
    document.getElementById('modalEntradaEstoque').style.display = 'flex';
    document.getElementById('formEntradaEstoque').reset();
    document.getElementById('estoqueAtualIMBEL').value = '-';
    
    // Atualizar select de produtos
    const select = document.getElementById('produtoEntrada');
    select.innerHTML = '<option value="">Selecione um produto</option>';
    const produtosOrdenados = [...estoque.produtos].sort((a, b) => a.nome.localeCompare(b.nome, 'pt-BR'));
    produtosOrdenados.forEach(produto => {
        const option = document.createElement('option');
        option.value = produto.id;
        option.textContent = produto.nome;
        select.appendChild(option);
    });
}

function abrirModalDevolucao() {
    if (!requireAdminOrNotify()) return;
    const modal = document.getElementById('modalDevolucao');
    if (!modal) return;
    modal.style.display = 'flex';
    document.getElementById('formDevolucao').reset();
    // Popular selects de representantes via helper centralizado
    try { popularSelectRepresentantes('representanteDevolucao', false); } catch (e) {}
    try { popularSelectRepresentantes('destinoDevolucao', true); } catch (e) {}

    // Reset container de itens e adicionar uma linha inicial
    try {
        const container = document.getElementById('itensDevolucaoContainer');
        if (container) {
            container.innerHTML = '';
            adicionarItemDevolucao();
        }
    } catch (e) {}
}

// Fecha um modal e restaura z-index do header fixo se necessário
function fecharModal(modalId) {
    try {
        const modal = document.getElementById(modalId);
        if (modal) modal.style.display = 'none';
    } catch (e) { /* ignore */ }
    // Se o modal fechado for o de produto, restaurar estado de edição padrão
    try {
        if (modalId === 'modalProduto') {
            produtoEditandoId = null;
            const header = document.querySelector('#modalProduto .modal-header h2');
            if (header) header.innerHTML = '<span>+</span> Adicionar Novo Produto';
            const submitBtn = document.querySelector('#modalProduto .modal-footer button[type="submit"]');
            if (submitBtn) submitBtn.textContent = 'Salvar Produto';
        }
    } catch (e) {}
    // Sempre tentar restaurar z-index do header fixo quando um modal fechar
}

function mostrarEstoqueAtual() {
    const produtoId = parseInt(document.getElementById('produtoEntrada').value);
    const produto = estoque.produtos.find(p => p.id === produtoId);
    
    if (produto) {
        const estoqueIMBEL = calcularImbelDisponivel(produto);
        document.getElementById('estoqueAtualIMBEL').value = `${estoqueIMBEL} unidades`;
    } else {
        document.getElementById('estoqueAtualIMBEL').value = '-';
    }
}

function salvarEntradaEstoque(event) {
    if (!requireAdminOrNotify()) return;
    event.preventDefault();
    
    const produtoId = parseInt(document.getElementById('produtoEntrada').value);
    const quantidade = parseInt(document.getElementById('quantidadeEntrada').value);
    const observacao = document.getElementById('observacaoEntrada').value.trim();
    
    const produto = estoque.produtos.find(p => p.id === produtoId);
    
    if (!produto) {
        mostrarNotificacao('Produto não encontrado!', 'error');
        return;
    }
    
    // Adicionar ao estoque da IMBEL
    // Agora adicionamos ao `estoqueConsolidado` (cadastro do total disponível)
    produto.estoqueConsolidado = (Number(produto.estoqueConsolidado) || 0) + quantidade;

    // Registrar entrada histórica para permitir edição/auditoria futura
    try {
        if (!Array.isArray(estoque.registroEntradas)) estoque.registroEntradas = [];
        estoque.registroEntradas.push({
            id: Date.now() + Math.floor(Math.random() * 1000),
            produtoId: produto.id,
            produtoNome: produto.nome,
            quantidade: quantidade,
            data: (new Date()).toISOString().split('T')[0],
            observacoes: observacao
        });
    } catch (e) { /* ignore registro failure */ }
    
    salvarDados();
    renderizarTabela();
    renderizarDashboard();
    fecharModal('modalEntradaEstoque');
    
    const msgObs = observacao ? ` (${observacao})` : '';
    mostrarNotificacao(`Entrada registrada: +${quantidade} "${produto.nome}" no estoque IMBEL${msgObs}`, 'success');
}

// ========================================
// REGISTRO DE VENDAS DETALHADO
// ========================================

// Configuração de Numeração de Contratos (NNN/AAAA)
function carregarConfigContrato() {
    try {
        const cfg = JSON.parse(localStorage.getItem('configContrato') || '{}');
        return {
            ano:     cfg.ano     || new Date().getFullYear(),
            proximo: cfg.proximo || 1,
        };
    } catch(e) {
        return { ano: new Date().getFullYear(), proximo: 1 };
    }
}

function salvarConfigContrato() {
    const ano     = parseInt(document.getElementById('configContratoAno')?.value)     || new Date().getFullYear();
    const proximo = parseInt(document.getElementById('configContratoProximo')?.value)  || 1;
    const cfg = { ano, proximo };
    localStorage.setItem('configContrato', JSON.stringify(cfg));
    try { atualizarPreviaContrato(); } catch(e) {}
    try { mostrarNotificacao('Configuração de contrato salva!', 'success'); } catch(e) {}
}

function atualizarPreviaContrato() {
    const ano     = parseInt(document.getElementById('configContratoAno')?.value)     || new Date().getFullYear();
    const proximo = parseInt(document.getElementById('configContratoProximo')?.value)  || 1;
    const previa  = document.getElementById('configContratoPreviaPrev');
    if (previa) {
        previa.textContent = String(proximo).padStart(3,'0') + '/' + ano;
    }
}

function gerarNumeroContrato() {
    const cfg = carregarConfigContrato();
    const anoAtual = new Date().getFullYear();

    if (cfg.ano !== anoAtual) {
        // new year: reset sequence
        cfg.ano = anoAtual;
        cfg.proximo = 1;
    }

    const numero = String(cfg.proximo).padStart(3, '0') + '/' + cfg.ano;

    // Increment and save
    cfg.proximo++;
    localStorage.setItem('configContrato', JSON.stringify(cfg));

    return numero;
}

function abrirModalVendaDetalhada(vendaId = null, propostaId = null) {
    // vendaId: se fornecido, abre o modal em modo de edição para essa venda
    // propostaId: se fornecido, preenche campos a partir da proposta para revisão
    if (!requireAdminOrNotify()) return;
    const modalEl = document.getElementById('modalVendaDetalhada');
    modalEl.style.display = 'flex';
    document.getElementById('formVendaDetalhada').reset();
    document.getElementById('valorUnitarioVenda').value = '';
    document.getElementById('valorTotalVenda').value = '';
    atualizarSelectsProdutos();
    try { popularSelectRepresentantes('representanteVendaDet', true); } catch (e) {}

    const container = document.getElementById('itensVendaContainer');
    if (!container) return;

    // Garantir listener no campo cliente/loja para atualizar preços quando mudar
    try {
        const lojaEl = document.getElementById('lojaVenda');
        if (lojaEl) {
            lojaEl.oninput = function() {
                try { preencherDadosCliente(lojaEl.value); } catch(e) {}
                try { atualizarPrecosVendaPorCliente(); } catch(e) {}
            };
        }
    } catch (e) {}

    // Pre-fill from proposal if propostaId provided
    if (propostaId && !vendaId) {
        const proposta = (propostas || []).find(p => p.id === propostaId || p.id === String(propostaId));
        if (proposta) {
            vendaEditandoId = null;
            container.innerHTML = '';

            document.getElementById('contratoVenda').value = gerarNumeroContrato();
            document.getElementById('lojaVenda').value = proposta.cliente || '';
            document.getElementById('representanteVendaDet').value = proposta.representante || '';
            document.getElementById('observacoesVenda').value = 'Gerado a partir da proposta ' + proposta.numero + (proposta.observacoes ? '\n' + proposta.observacoes : '');
            try { document.getElementById('dataVenda').value = new Date().toISOString().slice(0, 10); } catch (e) {}

            if (Array.isArray(proposta.itens) && proposta.itens.length > 0) {
                proposta.itens.forEach(it => {
                    const preValor = it.valorUnitario ? it.valorUnitario.toLocaleString('pt-BR', { minimumFractionDigits: 2 }) : '';
                    adicionarItemVendaRow(it.produtoId, it.quantidade, preValor);
                });
            } else {
                adicionarItemVendaRow();
            }
            atualizarTotalVendaDetalhada();
            return;
        }
    }

    // Se estamos criando nova venda, inicializa com uma linha vazia e sugere contrato
    if (!vendaId) {
        vendaEditandoId = null;
        container.innerHTML = '';
        adicionarItemVendaRow();

        document.getElementById('contratoVenda').value = gerarNumeroContrato();
        // Preencher data padrão como hoje
        try { document.getElementById('dataVenda').value = new Date().toISOString().slice(0,10); } catch (e) {}
        return;
    }

    // Modo edição: preencher campos com os dados da venda
    const venda = estoque.registroVendas.find(v => v.id === vendaId);
    if (!venda) {
        mostrarNotificacao('Venda não encontrada para edição', 'error');
        vendaEditandoId = null;
        container.innerHTML = '';
        adicionarItemVendaRow();
        return;
    }

    vendaEditandoId = vendaId;
    container.innerHTML = '';

    document.getElementById('contratoVenda').value = venda.contrato || '';
    document.getElementById('lojaVenda').value = venda.loja || '';
    document.getElementById('representanteVendaDet').value = venda.representante || '';
    document.getElementById('observacoesVenda').value = venda.observacoes || '';
    // Preencher campo de data com valor existente (normalizado para YYYY-MM-DD)
    try { document.getElementById('dataVenda').value = parseDateToYYYYMMDD(venda.data) || ''; } catch (e) {}

    if (Array.isArray(venda.items) && venda.items.length > 0) {
        venda.items.forEach(it => {
            const preValor = it.valorUnitario ? it.valorUnitario.toLocaleString('pt-BR', { minimumFractionDigits: 2 }) : '';
            adicionarItemVendaRow(it.produtoId, it.quantidade, preValor);
        });
    } else {
        // compatibilidade com registro antigo
        const preValor = venda.valorUnitario ? venda.valorUnitario.toLocaleString('pt-BR', { minimumFractionDigits: 2 }) : '';
        adicionarItemVendaRow(venda.produtoId, venda.quantidade || 1, preValor);
    }

    atualizarTotalVendaDetalhada();

}

// Constrói opções de produtos (HTML) para selects dinâmicos
function construirOpcoesProdutos() {
    let html = '<option value="">Selecione um produto</option>';
    const produtosOrdenados = [...estoque.produtos].sort((a, b) => a.nome.localeCompare(b.nome, 'pt-BR'));
    produtosOrdenados.forEach(produto => {
        html += `<option value="${produto.id}">${produto.nome}</option>`;
    });
    return html;
}

function adicionarItemVendaRow(preProdutoId = '', preQuantidade = 1, preValor = '') {
    try {
        const container = document.getElementById('itensVendaContainer');
        if (!container) return;

        const row = document.createElement('div');
        row.className = 'item-venda-row';
        row.style.display = 'flex';
        row.style.gap = '8px';
        row.style.alignItems = 'center';
        row.style.marginBottom = '6px';

        // Construir opções com segurança
        let opcoesHtml = '';
        try {
            opcoesHtml = construirOpcoesProdutos();
        } catch (err) {
            console.error('Erro ao construir opções de produtos:', err);
            opcoesHtml = '<option value="">(nenhum produto)</option>';
        }

        row.innerHTML = `
            <select class="item-produto" onchange="autoPreencherPrecoProduto(this); atualizarItemRow(this)">${opcoesHtml}</select>
            <input type="number" class="item-quantidade" min="1" value="${preQuantidade}" style="width:90px" onchange="atualizarItemRow(this)" />
            <input type="text" class="item-valor" placeholder="Valor unit. (obrigatório)" style="width:140px" oninput="formatarMoeda(this); atualizarItemRow(this)" />
            <div class="item-subtotal" style="min-width:120px">-</div>
            <button type="button" class="btn btn-outline btn-sm" onclick="removerItemRow(this)">Remover</button>
        `;

        container.appendChild(row);

        // Preencher valores se fornecidos
        if (preProdutoId) row.querySelector('.item-produto').value = preProdutoId;
        if (preValor) row.querySelector('.item-valor').value = preValor;

        // Garantir que o input de valor remova o estado "auto-preenchido" quando o usuário editar
        try {
            const valorInp = row.querySelector('.item-valor');
            if (valorInp && !valorInp.dataset._listenerVal) {
                valorInp.addEventListener('input', () => {
                    valorInp.style.background = '';
                    valorInp.removeAttribute('data-autofilled');
                });
                valorInp.dataset._listenerVal = '1';
            }
        } catch (e) {}

        // Se já existe um cliente preenchido no modal, tentar preencher o preço salvo automaticamente
        try {
            const lojaEl = document.getElementById('lojaVenda');
            if (lojaEl && lojaEl.value && (!preValor || preValor === '')) {
                const sel = row.querySelector('.item-produto');
                if (sel) autoPreencherPrecoProduto(sel);
            }
        } catch (e) {}

        // Atualizar visual do item
        atualizarItemRow(row.querySelector('.item-produto'));
    } catch (err) {
        console.error('Erro em adicionarItemVendaRow:', err);
        mostrarNotificacao('Erro ao adicionar item. Veja o console para detalhes.', 'error');
    }
}

function removerItemRow(btn) {
    const row = btn.closest('.item-venda-row');
    if (row) {
        row.remove();
        atualizarTotalVendaDetalhada();
    }
}

function atualizarItemRow(el) {
    const row = el.closest ? el.closest('.item-venda-row') : el.parentElement;
    if (!row) return;

    const quantidade = parseInt(row.querySelector('.item-quantidade').value) || 0;
    const valorInput = row.querySelector('.item-valor').value || '';

    let unit = 0;
    if (valorInput && valorInput.trim() !== '') {
        unit = converterMoedaParaNumero(valorInput);
    }

    const subtotal = unit * quantidade;
    const subtotalEl = row.querySelector('.item-subtotal');
    subtotalEl.textContent = quantidade > 0 ? formatarMoedaValor(subtotal) : '-';

    // Placeholder fixo para reforçar valor informado na venda
    const valorEl = row.querySelector('.item-valor');
    if (valorEl && (!valorEl.value || valorEl.value.trim() === '')) {
        valorEl.placeholder = 'Valor unit. (obrigatório)';
    }

    atualizarTotalVendaDetalhada();
}

function atualizarTotalVendaDetalhada() {
    const container = document.getElementById('itensVendaContainer');
    if (!container) return;

    let total = 0;
    let totalQtd = 0;
    const rows = container.querySelectorAll('.item-venda-row');
    rows.forEach(row => {
        const quantidade = parseInt(row.querySelector('.item-quantidade').value) || 0;
        const valorInput = row.querySelector('.item-valor').value || '';

        let unit = 0;
        if (valorInput && valorInput.trim() !== '') {
            unit = converterMoedaParaNumero(valorInput);
        }

        total += unit * quantidade;
        totalQtd += quantidade;
    });

    document.getElementById('valorTotalVenda').value = total > 0 ? formatarMoedaValor(total) : '';
    document.getElementById('valorUnitarioVenda').value = totalQtd > 0 ? formatarMoedaValor(total / totalQtd) : '';
}

// Layout toggle removed per user request (button removed from modal)

function atualizarPrecoVenda() {
    // Valor de venda não é mais vinculado ao cadastro do produto.
    // Mantemos o total com base apenas nos valores informados nos itens da venda.
    atualizarTotalVendaDetalhada();
}

function validarEstoqueParaVenda(representante, itens) {
    const erros = [];
    (itens || []).forEach(it => {
        let produto = null;
        if (it.produtoId) produto = (estoque.produtos || []).find(p => p.id === it.produtoId);
        if (!produto && it.produtoNome) produto = (estoque.produtos || []).find(p => (p.nome || '').toLowerCase() === (it.produtoNome || '').toLowerCase());
        if (!produto) {
            erros.push({ produto: it.produtoNome || ('ID:' + (it.produtoId || '')), solicitado: it.quantidade || 0, disponivel: 0 });
            return;
        }
        const isImbelRep = (representante || '').toString().toUpperCase() === 'IMBEL';
        const disp = isImbelRep ? calcularImbelDisponivel(produto) : ((produto.distribuicao && produto.distribuicao[representante]) ? produto.distribuicao[representante] : 0);
        const vendido = isImbelRep ? 0 : ((produto.vendas && produto.vendas[representante]) ? produto.vendas[representante] : 0);
        const saldo = disp - vendido;
        if ((it.quantidade || 0) > saldo) {
            erros.push({ produto: produto.nome, solicitado: it.quantidade || 0, disponivel: saldo });
        }
    });
    return erros.length ? { valido: false, erros } : { valido: true };
}

function mostrarConfirmacaoEstoque(mensagem, callbackConfirmar, callbackCancelar) {
    const elMsg = document.getElementById('msgConfirmacaoEstoque');
    const modal = document.getElementById('modalConfirmacaoEstoque');
    if (!elMsg || !modal) {
        if (typeof callbackConfirmar === 'function') callbackConfirmar();
        return;
    }
    elMsg.textContent = mensagem;
    modal.style.display = 'block';
    const btn = document.getElementById('btnConfirmarVendaSemEstoque');
    if (btn) {
        btn.onclick = () => {
            fecharModal('modalConfirmacaoEstoque');
            if (typeof callbackConfirmar === 'function') callbackConfirmar();
        };
    }
    const btnCancel = modal.querySelector('.modal-footer .btn.btn-outline');
    if (btnCancel) {
        btnCancel.onclick = () => {
            fecharModal('modalConfirmacaoEstoque');
            if (typeof callbackCancelar === 'function') callbackCancelar();
        };
    }
}

function finalizarSalvamentoVendaDetalhada(params) {
    const { contrato, loja, representante, observacoes, itens, totalQtd, totalValor, isEditing, vendaAnterior, vendaAnteriorSnapshot, vendaEditandoIdLocal } = params;

    // Aplicar novos valores ao estoque
    (itens || []).forEach(it => {
        const produto = estoque.produtos.find(p => p.id === it.produtoId);
        if (produto) produto.vendas[representante] = (produto.vendas[representante] || 0) + it.quantidade;
    });

    if (isEditing && vendaAnterior) {
        const idx = estoque.registroVendas.findIndex(v => v.id === vendaEditandoIdLocal);
        if (idx !== -1) {
            estoque.registroVendas[idx].contrato = contrato;
            estoque.registroVendas[idx].loja = loja;
            estoque.registroVendas[idx].representante = representante;
            estoque.registroVendas[idx].items = itens;
            estoque.registroVendas[idx].quantidadeTotal = totalQtd;
            estoque.registroVendas[idx].valorTotal = totalValor;
            estoque.registroVendas[idx].observacoes = observacoes;
            const dataInput = document.getElementById('dataVenda') ? document.getElementById('dataVenda').value : '';
            if (dataInput && dataInput !== '') {
                try { estoque.registroVendas[idx].data = new Date(dataInput + 'T00:00:00Z').toISOString(); } catch (e) { estoque.registroVendas[idx].data = new Date().toISOString(); }
            } else {
                estoque.registroVendas[idx].data = new Date().toISOString();
            }
            try {
                registrarAuditoriaVenda(
                    'edicao',
                    vendaAnteriorSnapshot,
                    JSON.parse(JSON.stringify(estoque.registroVendas[idx])),
                    `Contrato ${contrato} atualizado (${totalQtd} itens / ${formatarMoedaValor(totalValor)})`
                );
            } catch (e) {}
        }
        vendaEditandoId = null;

        salvarDados();
        renderizarTabela();
        renderizarDashboard();
        renderizarRegistroVendas();
        fecharModal('modalVendaDetalhada');

        mostrarNotificacao(`Venda atualizada: Contrato ${contrato} - ${totalQtd} itens - ${formatarMoedaValor(totalValor)}`, 'success');
        return;
    }

    // Criar registro de venda com múltiplos itens (novo)
    const novaVenda = {
        id: Date.now(),
        contrato: contrato,
        loja: loja,
        representante: representante,
        items: itens,
        quantidadeTotal: totalQtd,
        valorTotal: totalValor,
        observacoes: observacoes,
        data: (function(){ const di = document.getElementById('dataVenda') ? document.getElementById('dataVenda').value : ''; if (di && di !== '') { try { return new Date(di + 'T00:00:00Z').toISOString(); } catch(e){} } return new Date().toISOString(); })()
    };

    estoque.registroVendas.push(novaVenda);

    try {
        registrarAuditoriaVenda(
            'criacao',
            null,
            JSON.parse(JSON.stringify(novaVenda)),
            `Contrato ${contrato} criado (${totalQtd} itens / ${formatarMoedaValor(totalValor)})`
        );
    } catch (e) {}

    salvarDados();
    renderizarTabela();
    renderizarDashboard();
    renderizarRegistroVendas();
    fecharModal('modalVendaDetalhada');

    mostrarNotificacao(`Venda registrada: Contrato ${contrato} - ${totalQtd} itens - ${formatarMoedaValor(totalValor)}`, 'success');
}

function salvarVendaDetalhada(event) {
    if (!requireAdminOrNotify()) return;
    event.preventDefault();
    const contrato = document.getElementById('contratoVenda').value.trim();
    const loja = document.getElementById('lojaVenda').value.trim().toUpperCase();
    const representante = document.getElementById('representanteVendaDet').value;
    const observacoes = document.getElementById('observacoesVenda').value.trim();

    // Coletar itens
    const container = document.getElementById('itensVendaContainer');
    if (!container) {
        mostrarNotificacao('Erro interno: container de itens não encontrado.', 'error');
        return;
    }

    const rows = Array.from(container.querySelectorAll('.item-venda-row'));
    if (rows.length === 0) {
        mostrarNotificacao('Adicione ao menos um item à venda.', 'error');
        return;
    }

    let itens = [];
    let erros = [];
    let totalQtd = 0;
    let totalValor = 0;

    rows.forEach((row, idx) => {
        const produtoId = parseInt(row.querySelector('.item-produto').value) || null;
        const quantidade = parseInt(row.querySelector('.item-quantidade').value) || 0;
        const valorInput = row.querySelector('.item-valor').value || '';

        if (!produtoId || quantidade <= 0) {
            erros.push(`Item ${idx + 1}: produto ou quantidade inválidos.`);
            return;
        }

        const produto = estoque.produtos.find(p => p.id === produtoId);
        if (!produto) {
            erros.push(`Item ${idx + 1}: produto não encontrado.`);
            return;
        }

        // Determinar preço unitário (obrigatório por item)
        let unit = 0;
        if (valorInput && valorInput.trim() !== '') {
            unit = converterMoedaParaNumero(valorInput);
        } else {
            erros.push(`Item ${idx + 1}: informe o valor unitário.`);
            return;
        }

        if (!unit || unit <= 0) {
            erros.push(`Item ${idx + 1}: valor unitário inválido.`);
            return;
        }

        const valorTotalItem = unit * quantidade;

        itens.push({ produtoId: produtoId, produtoNome: produto.nome, quantidade: quantidade, valorUnitario: unit, valorTotal: valorTotalItem });

        totalQtd += quantidade;
        totalValor += valorTotalItem;
    });

    if (erros.length > 0) {
        mostrarNotificacao(erros.join('\n'), 'error');
        return;
    }

    // Se estivermos editando, primeiro reverter os efeitos da venda anterior
    const isEditing = vendaEditandoId !== null;
    let vendaAnterior = null;
    let vendaAnteriorSnapshot = null;
    if (isEditing) {
        vendaAnterior = estoque.registroVendas.find(v => v.id === vendaEditandoId);
        if (!vendaAnterior) {
            mostrarNotificacao('Venda anterior não encontrada para edição.', 'error');
            vendaEditandoId = null;
            return;
        }
        try { vendaAnteriorSnapshot = JSON.parse(JSON.stringify(vendaAnterior)); } catch (e) { vendaAnteriorSnapshot = null; }

        // Reverter quantidades da venda anterior no representante antigo
        if (Array.isArray(vendaAnterior.items) && vendaAnterior.items.length > 0) {
            vendaAnterior.items.forEach(it => {
                const produto = estoque.produtos.find(p => p.id === it.produtoId);
                if (produto) {
                    produto.vendas[vendaAnterior.representante] = Math.max(0, (produto.vendas[vendaAnterior.representante] || 0) - it.quantidade);
                }
            });
        } else {
            const produto = estoque.produtos.find(p => p.id === vendaAnterior.produtoId);
            if (produto) {
                produto.vendas[vendaAnterior.representante] = Math.max(0, (produto.vendas[vendaAnterior.representante] || 0) - (vendaAnterior.quantidade || 0));
            }
        }
    }

    // Validar estoque antes de salvar
    const validacao = validarEstoqueParaVenda(representante, itens);
    if (!validacao.valido) {
        let msg = '⚠️ Estoque insuficiente para ' + representante + ':\n\n';
        validacao.erros.forEach(e => { msg += `• ${e.produto}: solicitado ${e.solicitado}, disponível ${e.disponivel}\n`; });
        msg += '\nDeseja registrar mesmo assim?';

        const cancelarCallback = () => {
            // re-aplica venda anterior caso o usuário cancele
            if (isEditing && vendaAnterior) {
                if (Array.isArray(vendaAnterior.items) && vendaAnterior.items.length > 0) {
                    vendaAnterior.items.forEach(it => {
                        const produto = estoque.produtos.find(p => p.id === it.produtoId);
                        if (produto) produto.vendas[vendaAnterior.representante] = (produto.vendas[vendaAnterior.representante] || 0) + it.quantidade;
                    });
                } else {
                    const produto = estoque.produtos.find(p => p.id === vendaAnterior.produtoId);
                    if (produto) produto.vendas[vendaAnterior.representante] = (produto.vendas[vendaAnterior.representante] || 0) + (vendaAnterior.quantidade || 0);
                }
                vendaEditandoId = null;
            }
        };

        mostrarConfirmacaoEstoque(msg, () => {
            finalizarSalvamentoVendaDetalhada({ contrato, loja, representante, observacoes, itens, totalQtd, totalValor, isEditing, vendaAnterior, vendaAnteriorSnapshot, vendaEditandoIdLocal: vendaEditandoId });
        }, cancelarCallback);

        return;
    }

    // Se válido, finalizar imediatamente
    finalizarSalvamentoVendaDetalhada({ contrato, loja, representante, observacoes, itens, totalQtd, totalValor, isEditing, vendaAnterior, vendaAnteriorSnapshot, vendaEditandoIdLocal: vendaEditandoId });
}

function renderizarRegistroVendas() {
    const tbody = document.getElementById('tabelaRegistroVendasBody');
    if (!tbody) return;
    
    const filtroRep = document.getElementById('filtroRepresentante')?.value || '';
    const filtroProduto = document.getElementById('filtroProduto')?.value || '';
    const filtroProdutoId = filtroProduto ? parseInt(filtroProduto) : null;
    const filtroDataInicio = document.getElementById('filtroVendasDataInicio')?.value || '';
    const filtroDataFim = document.getElementById('filtroVendasDataFim')?.value || '';
    
    let vendasFiltradas = estoque.registroVendas || [];
    
    if (filtroRep) {
        vendasFiltradas = vendasFiltradas.filter(v => v.representante === filtroRep);
    }

    // Expandir vendas para linhas de item (cada produto vira uma linha)
    const linhas = [];
    vendasFiltradas.forEach(venda => {
        const dataNorm = parseDateToYYYYMMDD(venda.data) || '';
        if (filtroDataInicio && dataNorm && dataNorm < filtroDataInicio) return;
        if (filtroDataFim && dataNorm && dataNorm > filtroDataFim) return;
        if ((filtroDataInicio || filtroDataFim) && !dataNorm) return;

        const rawContrato = (venda.contrato ?? '').toString().normalize('NFKC');
        const contratoClean = rawContrato.replace(/[\u200B-\u200D\uFEFF\s]+/g, '');
        const somenteDigitos = contratoClean.replace(/\D+/g, '');
        const contratoKey = somenteDigitos ? String(parseInt(somenteDigitos, 10)) : contratoClean.toUpperCase();

        if (Array.isArray(venda.items) && venda.items.length > 0) {
            venda.items.forEach(it => {
                if (filtroProdutoId && it.produtoId !== filtroProdutoId) return;
                const qtd = Number(it.quantidade || 0);
                const valorUnNum = Number(it.valorUnitario || 0);
                const valorTotNum = Number(it.valorTotal || (valorUnNum * qtd) || 0);
                linhas.push({
                    vendaId: venda.id,
                    contratoKey,
                    contratoRaw: venda.contrato,
                    loja: venda.loja || '-',
                    representante: venda.representante || '-',
                    dataNorm,
                    observacoes: venda.observacoes || '-',
                    produtoNome: it.produtoNome || '-',
                    quantidade: qtd,
                    valorUnitario: valorUnNum,
                    valorTotal: valorTotNum
                });
            });
        } else {
            if (filtroProdutoId && venda.produtoId !== filtroProdutoId) return;
            const qtd = Number(venda.quantidade || 0);
            const valorUnNum = Number(venda.valorUnitario || 0);
            const valorTotNum = Number(venda.valorTotal || (valorUnNum * qtd) || 0);
            linhas.push({
                vendaId: venda.id,
                contratoKey,
                contratoRaw: venda.contrato,
                loja: venda.loja || '-',
                representante: venda.representante || '-',
                dataNorm,
                observacoes: venda.observacoes || '-',
                produtoNome: venda.produtoNome || '-',
                quantidade: qtd,
                valorUnitario: valorUnNum,
                valorTotal: valorTotNum
            });
        }
    });

    // Ordenar conforme seleção do usuário (_ordenVendas)
    linhas.sort((a, b) => {
        const campo = _ordenVendas.campo || 'contrato';
        const dire = _ordenVendas.direcao === 'asc' ? 1 : -1;
        try {
            if (campo === 'contrato') {
                const na = parseInt(a.contratoKey);
                const nb = parseInt(b.contratoKey);
                if (!isNaN(na) && !isNaN(nb)) return dire * (na - nb);
                return dire * a.contratoKey.localeCompare(b.contratoKey);
            }
            if (campo === 'valorTotal') {
                return dire * ((a.valorTotal || 0) - (b.valorTotal || 0));
            }
            if (campo === 'data') {
                const da = a.dataNorm || '';
                const db = b.dataNorm || '';
                if (da === db) return 0;
                return dire * (da < db ? -1 : 1);
            }
            const mapaCampo = { loja: 'loja', representante: 'representante' };
            const chave = mapaCampo[campo] || 'contratoKey';
            const aa = (a[chave] || '').toString().toLowerCase();
            const bb = (b[chave] || '').toString().toLowerCase();
            if (aa < bb) return -1 * dire;
            if (aa > bb) return 1 * dire;
            return 0;
        } catch (e) { return 0; }
    });
    
    tbody.innerHTML = '';
    
    if (linhas.length === 0) {
        tbody.innerHTML = `
            <tr>
                <td colspan="10" class="empty-state">
                    <div class="empty-icon">📋</div>
                    <div class="empty-text">Nenhuma venda registrada</div>
                    <div class="empty-hint">Clique em "Nova Venda" para adicionar o primeiro registro</div>
                </td>
            </tr>
        `;
        atualizarTotaisVendas(0, 0);
        return;
    }
    
    let totalQtd = 0;
    let totalValor = 0;

    const grupos = {};
    linhas.forEach(linha => {
        if (!grupos[linha.contratoKey]) grupos[linha.contratoKey] = [];
        grupos[linha.contratoKey].push(linha);
    });

    // Ordenar chaves por contrato numérico
    const chavesOrdenadas = Object.keys(grupos).sort((a,b) => {
        const na = parseInt(a);
        const nb = parseInt(b);
        if (!isNaN(na) && !isNaN(nb)) return na - nb;
        return a.localeCompare(b);
    });

    // Remover overlay de debug temporário se existir
    const dbg = document.getElementById('debug-grupos-vendas');
    if (dbg) dbg.remove();

    chavesOrdenadas.forEach(contratoKey => {
        const grupo = grupos[contratoKey] || [];
        const linhasDoContrato = grupo.length;
        if (!linhasDoContrato) return;

        const totalContrato = grupo.reduce((sum, linha) => sum + (Number(linha.valorTotal) || 0), 0);
        const totalQtdContrato = grupo.reduce((sum, linha) => sum + (Number(linha.quantidade) || 0), 0);
        const primeira = grupo[0];
        const repClass = (primeira.representante || '').toLowerCase();
        const obsGrupo = grupo.find(g => g.observacoes && g.observacoes !== '-')?.observacoes || primeira.observacoes || '-';
        const minData = grupo.map(g => g.dataNorm).filter(Boolean).sort()[0] || null;
        const maxData = grupo.map(g => g.dataNorm).filter(Boolean).sort().slice(-1)[0] || null;
        const dataDisplay = minData
            ? (maxData && maxData !== minData
                ? `${formatDateToDDMMYYYY(minData)} até ${formatDateToDDMMYYYY(maxData)}`
                : formatDateToDDMMYYYY(minData))
            : '-';

        const expandido = !!_contratosExpandidos[contratoKey];

        // verificar se todas as vendas deste contrato estão canceladas
        const vendasDoContrato = (estoque.registroVendas || []).filter(v => normalizarContratoKey(v.contrato) === contratoKey);
        const contratoCancelado = vendasDoContrato.length > 0 && vendasDoContrato.every(v => !!v.cancelado);
        const cancelBadgeHtml = contratoCancelado ? '<span class="badge-cancelado">CANCELADO</span>' : '';

        const resumo = document.createElement('tr');
        resumo.className = 'row-contrato-resumo' + (contratoCancelado ? ' contrato-cancelado' : '');
        let actionsHtml = '';
        if (contratoCancelado) {
            actionsHtml = '<span class="badge-cancelado">CANCELADO</span>';
        } else {
            actionsHtml = `<button class="btn-action btn-edit" onclick="abrirModalVendaDetalhada(${primeira.vendaId})" title="Editar venda">✎</button>` +
                          `<button class="btn-action btn-delete" onclick="excluirVenda(${primeira.vendaId})" title="Excluir venda">🗑</button>` +
                          `<button class="btn-action" onclick="abrirHistoricoContrato('${contratoKey}')" title="Histórico do Contrato">🕘</button>` +
                          `<button class="btn-action btn-cancel" onclick="cancelarContrato('${contratoKey}')" title="Cancelar contrato">✖</button>`;
        }

        resumo.innerHTML = `
            <td class="col-contrato">${contratoKey || '-'} ${cancelBadgeHtml}</td>
            <td class="col-loja" title="${primeira.loja}">${primeira.loja}</td>
            <td class="col-representante"><span class="badge-rep ${repClass}">${primeira.representante}</span></td>
            <td class="col-produto-venda"><button class="btn-expand-contrato" onclick="toggleContratoExpandido('${contratoKey}')">${expandido ? '▾' : '▸'} ${linhasDoContrato} item(ns)</button></td>
            <td class="col-qtd">${totalQtdContrato}</td>
            <td class="col-valor-un">-</td>
            <td class="col-valor-total">${formatarMoedaValor(totalContrato)}</td>
            <td class="col-data">${dataDisplay}</td>
            <td class="col-obs" title="${obsGrupo}">${obsGrupo}</td>
            <td class="col-acoes">${actionsHtml}</td>
        `;
        tbody.appendChild(resumo);

        grupo.forEach((linha) => {
            const tr = document.createElement('tr');
            tr.className = `row-contrato-detalhe ${expandido ? '' : 'hidden-row'}`;
            const valorUn = linha.valorUnitario ? formatarMoedaValor(linha.valorUnitario) : '-';
            const valorTot = linha.valorTotal || 0;
            totalQtd += linha.quantidade || 0;
            totalValor += valorTot || 0;

            // verificar se a venda específica está cancelada
            const vendaObj = estoque.registroVendas.find(v => v.id === linha.vendaId);
            const isLinhaCancelada = vendaObj && vendaObj.cancelado;
            const detalheAcoesHtml = isLinhaCancelada
                ? '<span class="badge-cancelado">CANCELADO</span>'
                : `<button class="btn-action btn-edit" onclick="abrirModalVendaDetalhada(${linha.vendaId})" title="Editar venda">✎</button><button class="btn-action btn-delete" onclick="excluirVenda(${linha.vendaId})" title="Excluir venda">🗑</button>`;

            tr.innerHTML = `
                <td class="col-contrato detalhe-vazio"></td>
                <td class="col-loja detalhe-vazio"></td>
                <td class="col-representante detalhe-vazio"></td>
                <td class="col-produto-venda" title="${linha.produtoNome}">↳ ${linha.produtoNome}</td>
                <td class="col-qtd">${linha.quantidade}</td>
                <td class="col-valor-un">${valorUn}</td>
                <td class="col-valor-total">${valorTot > 0 ? formatarMoedaValor(valorTot) : '-'}</td>
                <td class="col-data">${linha.dataNorm ? formatDateToDDMMYYYY(linha.dataNorm) : '-'}</td>
                <td class="col-obs" title="${linha.observacoes || '-'}">${linha.observacoes || '-'}</td>
                <td class="col-acoes">${detalheAcoesHtml}</td>
            `;

            tbody.appendChild(tr);
        });
    });
    
    atualizarTotaisVendas(totalQtd, totalValor);
}

function atualizarTotaisVendas(totalQtd, totalValor) {
    const spanQtd = document.getElementById('totalQtdVendas');
    const spanValor = document.getElementById('totalValorVendas');
    
    if (spanQtd) spanQtd.innerHTML = `<strong>${totalQtd.toLocaleString('pt-BR')}</strong>`;
    if (spanValor) spanValor.innerHTML = `<strong>${formatarMoedaValor(totalValor)}</strong>`;
}

function toggleContratoExpandido(contratoKey) {
    _contratosExpandidos[contratoKey] = !_contratosExpandidos[contratoKey];
    renderizarRegistroVendas();
}

function abrirHistoricoContrato(contratoInformado = '') {
    const contrato = (contratoInformado || prompt('Informe o contrato para visualizar o histórico:') || '').trim();
    if (!contrato) return;

    const lista = obterAuditoriaPorContrato(contrato);
    const container = document.getElementById('historicoConteudo');
    if (!container) return;

    if (lista.length === 0) {
        container.innerHTML = `<p style="text-align:center;color:var(--text-secondary);padding:20px">Nenhum histórico encontrado para o contrato ${contrato}.</p>`;
    } else {
        container.innerHTML = lista.map(h => {
            const dt = h.quando ? new Date(h.quando).toLocaleString('pt-BR') : '-';
            const quem = h.quem || 'Usuário';
            const acao = (h.acao || '-').toUpperCase();
            const desc = h.detalhes || '-';
            return `<div class="historico-item">
                <span class="hist-data">${dt}<br><small>${quem}</small></span>
                <span class="hist-tipo venda">${acao}</span>
                <span class="hist-descricao">${desc}</span>
            </div>`;
        }).join('');
    }

    document.getElementById('modalHistorico').style.display = 'flex';
}

function filtrarVendas() {
    renderizarRegistroVendas();
}

function limparFiltrosVendas() {
    const filtroRep = document.getElementById('filtroRepresentante');
    const filtroProduto = document.getElementById('filtroProduto');
    
    if (filtroRep) filtroRep.value = '';
    if (filtroProduto) filtroProduto.value = '';
    
    renderizarRegistroVendas();
}

function excluirVenda(vendaId) {
    if (!requireAdminOrNotify()) return;
    const venda = estoque.registroVendas.find(v => v.id === vendaId);
    let vendaSnapshot = null;
    try { vendaSnapshot = venda ? JSON.parse(JSON.stringify(venda)) : null; } catch (e) { vendaSnapshot = null; }

    if (!venda) {
        mostrarNotificacao('Venda não encontrada!', 'error');
        return;
    }

    // Mensagem resumo para confirmação
    let resumo = '';
    if (Array.isArray(venda.items) && venda.items.length > 0) {
        resumo = venda.items.map(it => `${it.produtoNome} x ${it.quantidade}`).join('\n');
    } else {
        resumo = `${venda.produtoNome || '-'} x ${venda.quantidade || 0}`;
    }

    if (!confirm(`Deseja excluir a venda do contrato ${venda.contrato}?\n\n${resumo}\n\nATENÇÃO: As quantidades serão devolvidas ao estoque do representante.`)) {
        return;
    }

    // Devolver ao estoque: lidar com vendas multi-itens
    if (Array.isArray(venda.items) && venda.items.length > 0) {
        venda.items.forEach(it => {
            const produto = estoque.produtos.find(p => p.id === it.produtoId);
            if (produto) {
                produto.vendas[venda.representante] = Math.max(0, (produto.vendas[venda.representante] || 0) - it.quantidade);
            }
        });
    } else {
        const produto = estoque.produtos.find(p => p.id === venda.produtoId);
        if (produto) {
            produto.vendas[venda.representante] = Math.max(0, (produto.vendas[venda.representante] || 0) - venda.quantidade);
        }
    }

    // Remover do registro
    estoque.registroVendas = estoque.registroVendas.filter(v => v.id !== vendaId);
    
    // Remover do controle de envio se este era o último contrato
    const contratoRestante = estoque.registroVendas.some(v => v.contrato === venda.contrato);
    if (!contratoRestante && estoque.controleEnvio[venda.contrato]) {
        delete estoque.controleEnvio[venda.contrato];
    }

    salvarDados();
    renderizarTabela();
    renderizarDashboard();
    renderizarRegistroVendas();
    renderizarControleEnvio();

    mostrarNotificacao(`Venda do contrato ${venda.contrato} excluída com sucesso!`, 'success');

    try {
        registrarAuditoriaVenda(
            'exclusao',
            vendaSnapshot,
            null,
            `Contrato ${venda.contrato} excluído`
        );
    } catch (e) {}
}

function cancelarContrato(contratoKey) {
    if (!requireAdminOrNotify()) return;
    if (!contratoKey) return mostrarNotificacao('Contrato inválido para cancelamento.', 'error');
    const ok = confirm(`Confirma cancelamento do contrato ${contratoKey}?\n\nAs quantidades serão devolvidas ao representante e a comissão não será considerada.`);
    if (!ok) return;

    const vendasAfetadas = (estoque.registroVendas || []).filter(v => normalizarContratoKey(v.contrato) === contratoKey && !v.cancelado);
    if (vendasAfetadas.length === 0) {
        mostrarNotificacao(`Nenhuma venda ativa encontrada para o contrato ${contratoKey}.`, 'warning');
        return;
    }

    vendasAfetadas.forEach(venda => {
        let snapshot = null;
        try { snapshot = JSON.parse(JSON.stringify(venda)); } catch (e) { snapshot = null; }

        // Reverter quantidades vendidas no representante
        if (Array.isArray(venda.items) && venda.items.length > 0) {
            venda.items.forEach(it => {
                const produto = estoque.produtos.find(p => p.id === it.produtoId);
                if (produto) produto.vendas[venda.representante] = Math.max(0, (produto.vendas[venda.representante] || 0) - it.quantidade);
            });
        } else {
            const produto = estoque.produtos.find(p => p.id === venda.produtoId);
            if (produto) produto.vendas[venda.representante] = Math.max(0, (produto.vendas[venda.representante] || 0) - (venda.quantidade || 0));
        }

        venda.cancelado = true;
        venda.canceladoEm = new Date().toISOString();
        venda.canceladoPor = getUsuarioAtual();

        try {
            registrarAuditoriaVenda('cancelamento', snapshot, JSON.parse(JSON.stringify(venda)), `Contrato ${contratoKey} cancelado`);
        } catch (e) {}
    });

    salvarDados();
    renderizarTabela();
    renderizarDashboard();
    renderizarRegistroVendas();
    renderizarControleEnvio();

    mostrarNotificacao(`Contrato ${contratoKey} cancelado. Quantidades devolvidas ao representante.`, 'success');
}

function exportarVendas() {
    const vendas = estoque.registroVendas || [];
    
    // Ordenar por contrato
    const vendasOrdenadas = [...vendas].sort((a, b) => {
        const contratoA = parseInt(a.contrato) || 0;
        const contratoB = parseInt(b.contrato) || 0;
        return contratoA - contratoB;
    });
    
    // Usar ponto-e-vírgula como separador (padrão Excel PT-BR)
    const sep = ';';
    
    // Cabeçalho
    let csv = `CONTRATO${sep}LOJA/CLIENTE${sep}REPRESENTANTE${sep}PRODUTO${sep}QUANTIDADE${sep}VALOR UNITÁRIO${sep}VALOR TOTAL${sep}OBSERVAÇÕES${sep}DATA\n`;
    
    if (vendasOrdenadas.length > 0) {
        // Percorrer vendas; suportar vendas com múltiplos itens
        vendasOrdenadas.forEach(venda => {
            const data = venda.data ? new Date(venda.data).toLocaleDateString('pt-BR') : '';
            if (Array.isArray(venda.items) && venda.items.length > 0) {
                venda.items.forEach(it => {
                    const valorUnit = typeof it.valorUnitario === 'number' ? it.valorUnitario : 0;
                    const valorTot = typeof it.valorTotal === 'number' ? it.valorTotal : (valorUnit * (it.quantidade || 0));
                    csv += `${venda.contrato}${sep}${venda.loja}${sep}${venda.representante}${sep}${it.produtoNome}${sep}${it.quantidade}${sep}${valorUnit.toFixed(2).replace('.', ',')}${sep}${valorTot.toFixed(2).replace('.', ',')}${sep}${venda.observacoes || ''}${sep}${data}\n`;
                });
            } else {
                // venda no formato antigo
                const produtoNome = venda.produtoNome || '';
                const quantidade = venda.quantidade || 0;
                const valorUnit = (typeof venda.valorUnitario === 'number') ? venda.valorUnitario : 0;
                const valorTot = (typeof venda.valorTotal === 'number') ? venda.valorTotal : 0;
                csv += `${venda.contrato}${sep}${venda.loja}${sep}${venda.representante}${sep}${produtoNome}${sep}${quantidade}${sep}${valorUnit.toFixed(2).replace('.', ',')}${sep}${valorTot.toFixed(2).replace('.', ',')}${sep}${venda.observacoes || ''}${sep}${data}\n`;
            }
        });

        // Linha de total (somar corretamente considerando itens)
        const totalQtd = vendas.reduce((sum, v) => {
            if (Array.isArray(v.items) && v.items.length > 0) return sum + v.items.reduce((s, it) => s + (it.quantidade || 0), 0);
            return sum + (v.quantidade || 0);
        }, 0);
        const totalValor = vendas.reduce((sum, v) => {
            if (Array.isArray(v.items) && v.items.length > 0) return sum + v.items.reduce((s, it) => s + (typeof it.valorTotal === 'number' ? it.valorTotal : ((it.valorUnitario||0) * (it.quantidade||0))), 0);
            return sum + (typeof v.valorTotal === 'number' ? v.valorTotal : 0);
        }, 0);
        csv += `${sep}${sep}${sep}TOTAL${sep}${totalQtd}${sep}${sep}${totalValor.toFixed(2).replace('.', ',')}${sep}${sep}\n`;
    }
    
    const blob = new Blob(['\ufeff' + csv], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    const url = URL.createObjectURL(blob);
    
    const dataAtual = new Date().toISOString().split('T')[0];
    
    link.setAttribute('href', url);
    link.setAttribute('download', `registro_vendas_${dataAtual}.csv`);
    link.style.visibility = 'hidden';
    
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    
    if (vendas.length === 0) {
        mostrarNotificacao('Modelo exportado (sem vendas registradas)', 'info');
    } else {
        mostrarNotificacao('Registro de vendas exportado com sucesso!', 'success');
    }
}

// ========================================
// REGISTRO DE DISTRIBUIÇÃO
// ========================================

function abrirModalNovaDistribuicao() {
    if (!requireAdminOrNotify()) return;
    document.getElementById('modalNovaDistribuicao').style.display = 'flex';
    document.getElementById('formNovaDistribuicao').reset();
    atualizarSelectsProdutos();
    try { popularSelectRepresentantes('representanteDistDet', false); } catch (e) {}
    // Reset container de itens e adicionar uma linha inicial
    try {
        const container = document.getElementById('itensDistribuicaoContainer');
        if (container) {
            container.innerHTML = '';
            adicionarItemDistribuicao();
        }
    } catch (e) {}

    // Definir data atual
    const hoje = new Date().toISOString().split('T')[0];
    document.getElementById('dataDistDet').value = hoje;
}

function salvarNovaDistribuicao(event) {
    if (!requireAdminOrNotify()) return;
    event.preventDefault();
    const representante = document.getElementById('representanteDistDet').value;
    const data = document.getElementById('dataDistDet').value || new Date().toISOString().split('T')[0];
    const observacoes = document.getElementById('observacoesDistDet').value.trim();

    const container = document.getElementById('itensDistribuicaoContainer');
    if (!container) {
        mostrarNotificacao('Container de itens não encontrado.', 'error');
        return;
    }

    const rows = Array.from(container.querySelectorAll('.item-dist-row'));
    if (rows.length === 0) {
        mostrarNotificacao('Adicione ao menos um produto para distribuir.', 'error');
        return;
    }

    // Agregar solicitações por produto (para validar estoque IMBEL consolidado)
    const solicitadoMap = new Map();
    const detalhesLinhas = [];
    const erros = [];

    rows.forEach((row, idx) => {
        const prodId = parseInt(row.querySelector('.item-produto-dist')?.value || '0');
        const qtd = parseInt(row.querySelector('.item-quantidade-dist')?.value || '0');
        if (!prodId || qtd <= 0) {
            erros.push(`Linha ${idx + 1}: produto ou quantidade inválidos.`);
            return;
        }
        detalhesLinhas.push({ produtoId: prodId, quantidade: qtd });
        solicitadoMap.set(prodId, (solicitadoMap.get(prodId) || 0) + qtd);
    });

    if (erros.length > 0) {
        mostrarNotificacao(erros.join('\n'), 'error');
        return;
    }

    // Validar estoque IMBEL para cada produto agregado usando registros como fonte de verdade
    const insuficientes = [];
    solicitadoMap.forEach((totalSolicitado, produtoId) => {
        const produto = estoque.produtos.find(p => p.id === produtoId);
        if (!produto) {
            insuficientes.push({ produtoId, nome: '(produto não encontrado)', disponivel: 0, solicitado: totalSolicitado });
            return;
        }
        const disponivel = calcularEstoqueIMBEL(produto.nome);
        if (totalSolicitado > disponivel) {
            insuficientes.push({ produtoId, nome: produto.nome, disponivel, solicitado: totalSolicitado });
        }
    });

    if (insuficientes.length > 0) {
        let msg = '⚠️ Estoque insuficiente na IMBEL para os seguintes produtos:\n\n';
        insuficientes.forEach(i => { msg += `• ${i.nome}: solicitado ${i.solicitado}, disponível ${i.disponivel}\n`; });
        mostrarNotificacao(msg, 'error');
        return;
    }

    // Processar cada linha
    const registrosCriados = [];
    detalhesLinhas.forEach(item => {
        const produto = estoque.produtos.find(p => p.id === item.produtoId);
        if (!produto) return;

        // Não mutamos `produto.distribuicao` diretamente aqui; persistimos o registro.

        // Criar registro individual para cada item
        const novaDistribuicao = {
            id: Date.now() + Math.floor(Math.random() * 1000),
            representante: representante,
            produtoId: item.produtoId,
            produtoNome: produto.nome,
            quantidade: item.quantidade,
            data: data,
            observacoes: observacoes
        };
        estoque.registroDistribuicao.push(novaDistribuicao);
        registrosCriados.push(novaDistribuicao);
    });
    // Reconstruir agregados a partir dos registros e salvar
    try { reconstruirDistribuicaoAPartirDeRegistros(); } catch (e) {}
    salvarDados();
    renderizarTabela();
    renderizarDashboard();
    renderizarRegistroDistribuicao();
    fecharModal('modalNovaDistribuicao');

    try { verificarAlertasEstoque(); } catch (e) {}

    mostrarNotificacao(`Distribuição registrada: ${registrosCriados.length} item(s) para ${representante}`, 'success');
}

function renderizarRegistroDistribuicao() {
    const tbody = document.getElementById('tabelaRegistroDistribuicaoBody');
    if (!tbody) return;
    
    const filtroRep = document.getElementById('filtroDistribuicaoRep')?.value || '';
    const filtroProduto = document.getElementById('filtroDistribuicaoProduto')?.value || '';
    
    // Filtrar distribuições
    let distribuicoesFiltradas = estoque.registroDistribuicao || [];
    
    if (filtroRep) {
        distribuicoesFiltradas = distribuicoesFiltradas.filter(d => d.representante === filtroRep);
    }
    
    if (filtroProduto) {
        distribuicoesFiltradas = distribuicoesFiltradas.filter(d => d.produtoId === parseInt(filtroProduto));
    }
    
    // Ordenar por data (mais recente primeiro)
    distribuicoesFiltradas.sort((a, b) => new Date(b.data) - new Date(a.data));
    
    tbody.innerHTML = '';
    
    if (distribuicoesFiltradas.length === 0) {
        tbody.innerHTML = `
            <tr>
                <td colspan="7" class="empty-state">
                    <div class="empty-icon">🚚</div>
                    <div class="empty-text">Nenhuma distribuição registrada</div>
                    <div class="empty-hint">Clique em "Nova Distribuição" para adicionar o primeiro registro</div>
                </td>
            </tr>
        `;
        atualizarTotaisDistribuicao(0);
        return;
    }
    
    let totalQtd = 0;
    let numero = distribuicoesFiltradas.length;
    
    distribuicoesFiltradas.forEach(dist => {
        totalQtd += dist.quantidade;
        
        const repClass = dist.representante.toLowerCase();
        const dataFormatada = dist.data ? new Date(dist.data + 'T00:00:00').toLocaleDateString('pt-BR') : '-';
        
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td class="col-contrato">${numero--}</td>
            <td class="col-loja"><span class="badge-rep ${repClass}">${dist.representante}</span></td>
            <td class="col-produto-venda" title="${dist.produtoNome}">${dist.produtoNome}</td>
            <td class="col-qtd">${dist.quantidade}</td>
            <td>${dataFormatada}</td>
            <td class="col-obs" title="${dist.observacoes || '-'}">${dist.observacoes || '-'}</td>
            <td class="col-acoes">
                <button class="btn-action btn-delete" onclick="excluirDistribuicao(${dist.id})" title="Excluir distribuição">🗑</button>
            </td>
        `;
        tbody.appendChild(tr);
    });
    
    atualizarTotaisDistribuicao(totalQtd);
}

function atualizarTotaisDistribuicao(totalQtd) {
    const spanQtd = document.getElementById('totalQtdDistribuicao');
    if (spanQtd) spanQtd.innerHTML = `<strong>${totalQtd.toLocaleString('pt-BR')}</strong>`;
}

function limparFiltrosDistribuicao() {
    const filtroRep = document.getElementById('filtroDistribuicaoRep');
    const filtroProduto = document.getElementById('filtroDistribuicaoProduto');
    
    if (filtroRep) filtroRep.value = '';
    if (filtroProduto) filtroProduto.value = '';
    
    renderizarRegistroDistribuicao();
}

function excluirDistribuicao(distId) {
    if (!requireAdminOrNotify()) return;
    const dist = estoque.registroDistribuicao.find(d => d.id === distId);
    
    if (!dist) {
        mostrarNotificacao('Distribuição não encontrada!', 'error');
        return;
    }
    
    if (!confirm(`Deseja excluir esta distribuição?\n\nRepresentante: ${dist.representante}\nProduto: ${dist.produtoNome}\nQuantidade: ${dist.quantidade}\n\nATENÇÃO: A quantidade será devolvida ao estoque da IMBEL.`)) {
        return;
    }
    
    // Devolver ao estoque da IMBEL e remover do representante
    const produto = estoque.produtos.find(p => p.id === dist.produtoId);
    if (produto) {
        produto.distribuicao[dist.representante] = Math.max(0, (produto.distribuicao[dist.representante] || 0) - dist.quantidade);
        // Nota: não alterar `estoqueConsolidado` aqui — o consolidado é cadastro.
    }
    
    // Remover do registro
    estoque.registroDistribuicao = estoque.registroDistribuicao.filter(d => d.id !== distId);
    
    salvarDados();
    renderizarTabela();
    renderizarDashboard();
    renderizarRegistroDistribuicao();
    
    mostrarNotificacao(`Distribuição excluída! ${dist.quantidade} unidades devolvidas à IMBEL.`, 'success');
}

function excluirDevolucao(devId) {
    if (!requireAdminOrNotify()) return;
    const dev = (estoque.registroDevolucoes || []).find(d => d.id === devId);
    if (!dev) {
        mostrarNotificacao('Devolução não encontrada!', 'error');
        return;
    }

    if (!confirm(`Deseja excluir esta devolução?\n\nOrigem: ${dev.origem}\nProduto: ${dev.produtoNome}\nQuantidade: ${dev.quantidade}\n\nATENÇÃO: Esta operação irá reverter a devolução (movendo as unidades de volta ao representante).`)) {
        return;
    }

    // Reverter a devolução: subtrair do destino e recolocar no representante de origem
    const produto = estoque.produtos.find(p => p.id === dev.produtoId);
    if (produto) {
        if ((dev.destino || '').toUpperCase() === 'IMBEL') {
            // Se a devolução foi para o consolidado (IMBEL), reduzir o consolidado
            produto.estoqueConsolidado = Math.max(0, (Number(produto.estoqueConsolidado) || 0) - dev.quantidade);
            produto.distribuicao[dev.origem] = (produto.distribuicao[dev.origem] || 0) + dev.quantidade;
        } else {
            produto.distribuicao[dev.destino] = Math.max(0, (produto.distribuicao[dev.destino] || 0) - dev.quantidade);
            produto.distribuicao[dev.origem] = (produto.distribuicao[dev.origem] || 0) + dev.quantidade;
        }
    }

    // Remover registro de devolução
    estoque.registroDevolucoes = (estoque.registroDevolucoes || []).filter(d => d.id !== devId);

    salvarDados();
    renderizarTabela();
    renderizarDashboard();
    renderizarRegistroDistribuicao();

    mostrarNotificacao(`Devolução excluída e estoque revertido (${dev.quantidade} unidades).`, 'success');
}

function exportarDistribuicao() {
    const distribuicoes = estoque.registroDistribuicao || [];
    
    // Ordenar por data
    const distribuicoesOrdenadas = [...distribuicoes].sort((a, b) => new Date(a.data) - new Date(b.data));
    
    // Usar ponto-e-vírgula como separador (padrão Excel PT-BR)
    const sep = ';';
    
    // Cabeçalho
    let csv = `REPRESENTANTE${sep}PRODUTO${sep}QUANTIDADE${sep}DATA${sep}OBSERVAÇÕES\n`;
    
    if (distribuicoesOrdenadas.length > 0) {
        distribuicoesOrdenadas.forEach(dist => {
            const dataFormatada = dist.data ? new Date(dist.data + 'T00:00:00').toLocaleDateString('pt-BR') : '';
            csv += `${dist.representante}${sep}${dist.produtoNome}${sep}${dist.quantidade}${sep}${dataFormatada}${sep}${dist.observacoes || ''}\n`;
        });
        
        // Linha de total
        const totalQtd = distribuicoes.reduce((sum, d) => sum + d.quantidade, 0);
        csv += `${sep}TOTAL${sep}${totalQtd}${sep}${sep}\n`;
    }
    
    const blob = new Blob(['\ufeff' + csv], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    const url = URL.createObjectURL(blob);
    
    const dataAtual = new Date().toISOString().split('T')[0];
    
    link.setAttribute('href', url);
    link.setAttribute('download', `registro_distribuicao_${dataAtual}.csv`);
    link.style.visibility = 'hidden';
    
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    
    if (distribuicoes.length === 0) {
        mostrarNotificacao('Modelo exportado (sem distribuições registradas)', 'info');
    } else {
        mostrarNotificacao('Registro de distribuição exportado com sucesso!', 'success');
    }
}

function importarDistribuicao(event) {
    const file = event.target.files[0];
    if (!file) return;
    
    const reader = new FileReader();
    
    reader.onload = function(e) {
        try {
            const conteudo = e.target.result;
            const linhas = conteudo.split(/\r?\n/).filter(l => l.trim());
            
            if (linhas.length < 2) {
                mostrarNotificacao('Arquivo vazio ou sem dados!', 'error');
                return;
            }
            
            let distribuicoesImportadas = 0;
            let erros = [];
            
            // Processar cada linha (começando da segunda - pular cabeçalho)
            for (let i = 1; i < linhas.length; i++) {
                const linha = linhas[i].trim();
                if (!linha || linha.toLowerCase().includes('total')) continue;
                
                const colunas = parseCsvLinha(linha);
                
                // Pular linha se primeira coluna estiver vazia
                if (!colunas[0] || colunas[0].trim() === '') continue;
                
                if (colunas.length < 3) {
                    erros.push(`Linha ${i + 1}: formato inválido`);
                    continue;
                }
                
                const representante = colunas[0]?.trim().toUpperCase();
                const produtoNome = colunas[1]?.trim().toUpperCase();
                const quantidade = parseInt(colunas[2]?.trim()) || 0;
                const dataStr = colunas[3]?.trim() || '';
                const observacoes = colunas[4]?.trim() || '';
                
                if (!representante || !produtoNome || quantidade <= 0) {
                    erros.push(`Linha ${i + 1}: dados obrigatórios faltando`);
                    continue;
                }
                
                // Verificar se representante é válido
                const repsValidos = ['KOLTE', 'ISA', 'LC', 'ADES', 'FL'];
                if (!repsValidos.includes(representante)) {
                    erros.push(`Linha ${i + 1}: representante inválido (${representante})`);
                    continue;
                }
                
                // Buscar produto pelo nome
                let produto = estoque.produtos.find(p => 
                    p.nome.toUpperCase() === produtoNome || 
                    p.nome.toUpperCase().includes(produtoNome) ||
                    produtoNome.includes(p.nome.toUpperCase())
                );
                
                if (!produto) {
                    erros.push(`Linha ${i + 1}: produto não encontrado (${produtoNome})`);
                    continue;
                }
                
                // Converter data
                let data = new Date().toISOString().split('T')[0];
                if (dataStr) {
                    // Tentar converter dd/mm/yyyy para yyyy-mm-dd
                    const partes = dataStr.split('/');
                    if (partes.length === 3) {
                        data = `${partes[2]}-${partes[1].padStart(2, '0')}-${partes[0].padStart(2, '0')}`;
                    }
                }
                
                // Verificar estoque disponível na IMBEL usando o novo modelo
                const estoqueIMBEL = calcularImbelDisponivel(produto);
                if (quantidade > estoqueIMBEL) {
                    erros.push(`Linha ${i + 1}: estoque insuficiente na IMBEL para ${produto.nome} (disponível: ${estoqueIMBEL})`);
                    continue;
                }
                
                // Criar registro da distribuição
                const novaDistribuicao = {
                    id: Date.now() + i,
                    representante: representante,
                    produtoId: produto.id,
                    produtoNome: produto.nome,
                    quantidade: quantidade,
                    data: data,
                    observacoes: observacoes
                };
                
                // Alocar para o representante (IMBEL é derivado do consolidado)
                produto.distribuicao[representante] = (produto.distribuicao[representante] || 0) + quantidade;
                
                // Adicionar ao registro
                estoque.registroDistribuicao.push(novaDistribuicao);
                distribuicoesImportadas++;
            }
            
            if (distribuicoesImportadas > 0) {
                salvarDados();
                renderizarTabela();
                renderizarDashboard();
                renderizarRegistroDistribuicao();
            }
            
            // Limpar input
            event.target.value = '';
            
            // Mostrar resultado
            if (erros.length > 0 && distribuicoesImportadas === 0) {
                mostrarNotificacao(`Nenhuma distribuição importada. Verifique o formato.`, 'error');
                console.debug('Erros de importação:', erros);
            } else if (erros.length > 0) {
                mostrarNotificacao(`${distribuicoesImportadas} distribuições importadas. ${erros.length} linhas com erro.`, 'warning');
                console.debug('Erros de importação:', erros);
            } else {
                mostrarNotificacao(`${distribuicoesImportadas} distribuições importadas com sucesso!`, 'success');
            }
            
        } catch (error) {
            console.error('Erro ao importar:', error);
            mostrarNotificacao('Erro ao processar o arquivo. Verifique o formato.', 'error');
        }
    };
    
    reader.readAsText(file, 'UTF-8');
}

function salvarProduto(event) {
    if (!requireAdminOrNotify()) return;
    event.preventDefault();
    
    const nome = document.getElementById('nomeProduto').value.trim().toUpperCase();
    const estoqueTotal = parseInt(document.getElementById('estoqueTotal').value) || 0;
    const ncm = (document.getElementById('produtoNCM')?.value || '').trim();
    const categoria = (document.getElementById('produtoCategoria')?.value || '').trim();
    const ci = Number(document.getElementById('produtoCI')?.value || 0) || 0;
    const margemMinima = Number(document.getElementById('produtoMargemMin')?.value || 0) || 0;
    const descontoMaximo = Number(document.getElementById('produtoDescMax')?.value || 0) || 0;
    const observacoes = (document.getElementById('produtoObservacoes')?.value || '').trim();

    if (!nome) {
        mostrarNotificacao('Informe o nome do produto.', 'error');
        return;
    }

    if (!precificacao || typeof precificacao !== 'object') precificacao = {};
    if (!categoriaPorProduto || typeof categoriaPorProduto !== 'object') categoriaPorProduto = {};

    // Se estamos editando um produto existente
    if (produtoEditandoId !== null) {
        const idx = estoque.produtos.findIndex(p => p.id === produtoEditandoId);
        if (idx === -1) {
            mostrarNotificacao('Produto não encontrado para edição.', 'error');
            produtoEditandoId = null;
            fecharModal('modalProduto');
            return;
        }
        // Verificar duplicidade de nome em outro produto
        if (estoque.produtos.some(p => p.nome === nome && p.id !== produtoEditandoId)) {
            mostrarNotificacao('Outro produto com este nome já existe!', 'error');
            return;
        }

        const produto = estoque.produtos[idx];
        const nomeAnterior = produto.nome;
        produto.nome = nome;
        produto.estoqueConsolidado = Number(estoqueTotal) || 0;
        produto.quantidadeInicial = Number(estoqueTotal) || 0;
        produto.ncm = ncm;
        produto.categoria = categoria;
        produto.ci = ci;
        produto.margemMinima = margemMinima;
        produto.margemMin = margemMinima;
        produto.descontoMaximo = descontoMaximo;
        produto.descMax = descontoMaximo;
        produto.observacoes = observacoes;
        produto.atualizadoEm = new Date().toISOString();
        produto.dataAtualizacao = Date.now();
        produto.distribuicao = produto.distribuicao || { KOLTE: 0, ISA: 0, LC: 0, ADES: 0, FL: 0, IMBEL: 0 };
        produto.vendas = produto.vendas || { KOLTE: 0, ISA: 0, LC: 0, ADES: 0, FL: 0, IMBEL: 0 };

        if (nomeAnterior && nomeAnterior !== nome) {
            if (precificacao[nomeAnterior] && !precificacao[nome]) {
                precificacao[nome] = { ...precificacao[nomeAnterior] };
            }
            delete precificacao[nomeAnterior];

            if (categoriaPorProduto[nomeAnterior] && !categoriaPorProduto[nome]) {
                categoriaPorProduto[nome] = categoriaPorProduto[nomeAnterior];
            }
            delete categoriaPorProduto[nomeAnterior];

            if (tabelaAliquotas[nomeAnterior] && !tabelaAliquotas[nome]) {
                tabelaAliquotas[nome] = { ...tabelaAliquotas[nomeAnterior] };
            }
            delete tabelaAliquotas[nomeAnterior];
        }

        precificacao[nome] = {
            ...(precificacao[nome] || {}),
            ci,
            margemMinima,
            descontoMaximo
        };
        if (categoria) categoriaPorProduto[nome] = categoria;

        atualizarSelectsProdutos();
        salvarDados();
        renderizarTabela();
        renderizarDashboard();
        renderizarCadastroProdutos();
        fecharModal('modalProduto');
        mostrarNotificacao(`Produto "${nome}" atualizado com sucesso!`, 'success');
        produtoEditandoId = null;
        return;
    }

    // Criar novo produto
    if (estoque.produtos.some(p => p.nome === nome)) {
        mostrarNotificacao('Este produto já existe no sistema!', 'error');
        return;
    }

    const novoProduto = {
        id: Date.now(),
        nome: nome,
        estoqueConsolidado: Number(estoqueTotal) || 0,
        quantidadeInicial: Number(estoqueTotal) || 0,
        ncm,
        categoria,
        ci,
        margemMinima,
        margemMin: margemMinima,
        descontoMaximo,
        descMax: descontoMaximo,
        observacoes,
        criadoEm: new Date().toISOString(),
        atualizadoEm: new Date().toISOString(),
        dataAtualizacao: Date.now(),
        distribuicao: { KOLTE: 0, ISA: 0, LC: 0, ADES: 0, FL: 0, IMBEL: 0 },
        vendas: { KOLTE: 0, ISA: 0, LC: 0, ADES: 0, FL: 0, IMBEL: 0 }
    };

    precificacao[nome] = {
        ...(precificacao[nome] || {}),
        ci,
        margemMinima,
        margemMin: margemMinima,
        descontoMaximo,
        descMax: descontoMaximo
    };
    if (categoria) categoriaPorProduto[nome] = categoria;

    estoque.produtos.push(novoProduto);
    // Atualizar selects imediatamente para refletir o novo produto em qualquer modal aberto
    atualizarSelectsProdutos();
    salvarDados();
    renderizarTabela();
    renderizarDashboard();
    renderizarCadastroProdutos();
    fecharModal('modalProduto');

    mostrarNotificacao(`Produto "${nome}" adicionado com sucesso!`, 'success');
}

function salvarDistribuicao(event) {
    if (!requireAdminOrNotify()) return;
    event.preventDefault();
    
    const produtoId = parseInt(document.getElementById('produtoDistribuicao').value);
    const representante = document.getElementById('representanteDistribuicao').value;
    const quantidade = parseInt(document.getElementById('quantidadeDistribuicao').value);
    
    const produto = estoque.produtos.find(p => p.id === produtoId);
    
    if (!produto) {
        mostrarNotificacao('Produto não encontrado!', 'error');
        return;
    }
    
    const saldoIMBEL = calcularEstoqueIMBEL(produto.nome);
    if (quantidade > saldoIMBEL) {
        mostrarNotificacao(`Estoque insuficiente na IMBEL! Saldo disponível: ${saldoIMBEL} unidades`, 'error');
        return;
    }

    // Criar registro da distribuição (não mutamos mais produto.distribuicao diretamente)
    const novaDistribuicao = {
        id: Date.now(),
        representante: representante,
        produtoId: produto.id,
        produtoNome: produto.nome,
        quantidade: quantidade,
        data: new Date().toISOString().split('T')[0]
    };
    estoque.registroDistribuicao.push(novaDistribuicao);

    try { reconstruirDistribuicaoAPartirDeRegistros(); } catch (e) {}
    salvarDados();
    renderizarTabela();
    renderizarDashboard();
    try { verificarAlertasEstoque(); } catch(e) {}
    fecharModal('modalDistribuicao');

    mostrarNotificacao(`${quantidade} unidades distribuídas para ${representante}!`, 'success');
}

function salvarVenda(event) {
    if (!requireAdminOrNotify()) return;
    event.preventDefault();
    const produtoId = parseInt(document.getElementById('produtoVenda').value);
    const vendedor = document.getElementById('vendedorVenda').value;
    const quantidade = parseInt(document.getElementById('quantidadeVenda').value);

    const produto = estoque.produtos.find(p => p.id === produtoId);
    if (!produto) {
        mostrarNotificacao('Produto não encontrado!', 'error');
        return;
    }

    // Usar a validação compartilhada
    const itens = [{ produtoId: produtoId, produtoNome: produto.nome, quantidade }];
    const validacao = typeof validarEstoqueParaVenda === 'function' ? validarEstoqueParaVenda(vendedor, itens) : { valido: true };
    if (!validacao.valido) {
        let msg = '⚠️ Estoque insuficiente para ' + vendedor + ':\n\n';
        validacao.erros.forEach(e => { msg += `• ${e.produto}: solicitado ${e.solicitado}, disponível ${e.disponivel}\n`; });
        msg += '\nDeseja registrar mesmo assim?';

        mostrarConfirmacaoEstoque(msg, () => {
            produto.vendas[vendedor] = (produto.vendas[vendedor] || 0) + quantidade;
            salvarDados();
            renderizarTabela();
            renderizarDashboard();
            fecharModal('modalVenda');
            mostrarNotificacao(`Venda registrada: ${quantidade}x "${produto.nome}"`, 'success');
        });
        return;
    }

    // Sem problemas de estoque — registra normalmente
    produto.vendas[vendedor] = (produto.vendas[vendedor] || 0) + quantidade;
    salvarDados();
    renderizarTabela();
    renderizarDashboard();
    fecharModal('modalVenda');
    mostrarNotificacao(`Venda registrada: ${quantidade}x "${produto.nome}"`, 'success');
}

function salvarDevolucao(event) {
    event.preventDefault();

    const representante = document.getElementById('representanteDevolucao').value;
    const destino = document.getElementById('destinoDevolucao') ? document.getElementById('destinoDevolucao').value : 'IMBEL';
    const container = document.getElementById('itensDevolucaoContainer');
    if (!container) {
        mostrarNotificacao('Container de itens não encontrado.', 'error');
        return;
    }

    const rows = Array.from(container.querySelectorAll('.item-dev-row'));
    if (rows.length === 0) {
        mostrarNotificacao('Adicione ao menos um produto para devolução.', 'error');
        return;
    }

    // Agregar por produto para validar saldos no representante de origem
    const solicitadoMap = new Map();
    const detalhes = [];
    const erros = [];

    rows.forEach((row, idx) => {
        const prodId = parseInt(row.querySelector('.item-produto-dev')?.value || '0');
        const qtd = parseInt(row.querySelector('.item-quantidade-dev')?.value || '0');
        if (!prodId || qtd <= 0) {
            erros.push(`Linha ${idx + 1}: produto ou quantidade inválidos.`);
            return;
        }
        detalhes.push({ produtoId: prodId, quantidade: qtd });
        solicitadoMap.set(prodId, (solicitadoMap.get(prodId) || 0) + qtd);
    });

    if (erros.length > 0) {
        mostrarNotificacao(erros.join('\n'), 'error');
        return;
    }

    const insuficientes = [];
    solicitadoMap.forEach((totalSolicitado, produtoId) => {
        const produto = estoque.produtos.find(p => p.id === produtoId);
        if (!produto) { insuficientes.push({ produtoId, nome: '(produto não encontrado)', disponivel: 0, solicitado: totalSolicitado }); return; }
        const disp = produto.distribuicao[representante] || 0;
        const vendido = produto.vendas[representante] || 0;
        const saldo = disp - vendido;
        if (totalSolicitado > saldo) insuficientes.push({ produtoId, nome: produto.nome, disponivel: saldo, solicitado: totalSolicitado });
    });

    if (insuficientes.length > 0) {
        let msg = '⚠️ Saldo insuficiente no representante para:\n\n';
        insuficientes.forEach(i => { msg += `• ${i.nome}: solicitado ${i.solicitado}, disponível ${i.disponivel}\n`; });
        mostrarNotificacao(msg, 'error');
        return;
    }

    const registros = [];
    detalhes.forEach(item => {
        const produto = estoque.produtos.find(p => p.id === item.produtoId);
        if (!produto) return;

        if (destino === representante) return; // pular (não faz sentido)

        // Subtrair do representante de origem
        produto.distribuicao[representante] = (produto.distribuicao[representante] || 0) - item.quantidade;
        if (produto.distribuicao[representante] < 0) produto.distribuicao[representante] = 0;

        // Garantir chave de destino: se destino for IMBEL, incrementar o consolidado
        if ((destino || '').toString().toUpperCase() === 'IMBEL') {
            produto.estoqueConsolidado = (Number(produto.estoqueConsolidado) || 0) + item.quantidade;
        } else {
            produto.distribuicao[destino] = (produto.distribuicao[destino] || 0) + item.quantidade;
        }

        // Registrar devolução
        try {
            if (!Array.isArray(estoque.registroDevolucoes)) estoque.registroDevolucoes = [];
            const registro = {
                id: Date.now() + Math.floor(Math.random() * 1000),
                origem: representante,
                destino: destino,
                produtoId: item.produtoId,
                produtoNome: produto.nome,
                quantidade: item.quantidade,
                data: (new Date()).toISOString().split('T')[0],
                observacoes: ''
            };
            estoque.registroDevolucoes.push(registro);
            registros.push(registro);
        } catch (e) { console.warn('Falha ao registrar devolução:', e); }
    });

    salvarDados();
    renderizarTabela();
    renderizarDashboard();
    atualizarSelectsProdutos();

    fecharModal('modalDevolucao');

    try { verificarAlertasEstoque(); } catch(e) {}

    mostrarNotificacao(`${registros.length} item(s) devolvidos de ${representante} para ${destino}!`, 'success');
}

function limparFiltros() {
    renderizarTabela();
    renderizarDashboard();
    atualizarEstatisticas();
    mostrarNotificacao('Dados atualizados!', 'info');
}

// ========================================
// SISTEMA DE NOTIFICAÇÕES
// ========================================

function mostrarNotificacao(mensagem, tipo = 'info') {
    const notificacaoExistente = document.querySelector('.notificacao');
    if (notificacaoExistente) {
        notificacaoExistente.remove();
    }

    const cores = {
        success: { bg: '#d4edda', border: '#28a745', text: '#155724', icon: '✓' },
        error: { bg: '#f8d7da', border: '#dc3545', text: '#721c24', icon: '✕' },
        warning: { bg: '#fff3cd', border: '#ffc107', text: '#856404', icon: '⚠' },
        info: { bg: '#d1ecf1', border: '#17a2b8', text: '#0c5460', icon: 'ℹ' }
    };

    const cor = cores[tipo] || cores.info;

    const notificacao = document.createElement('div');
    notificacao.className = 'notificacao';
    notificacao.style.cssText = `
        position: fixed;
        top: 20px;
        right: 20px;
        padding: 15px 20px;
        background: ${cor.bg};
        border: 1px solid ${cor.border};
        border-left: 4px solid ${cor.border};
        border-radius: 6px;
        color: ${cor.text};
        font-size: 0.9rem;
        font-weight: 500;
        box-shadow: 0 4px 12px rgba(0,0,0,0.15);
        z-index: 9999;
        display: flex;
        align-items: center;
        gap: 10px;
        animation: slideInRight 0.3s ease;
        max-width: 400px;
    `;
    
    notificacao.innerHTML = `<span style="font-size: 1.2rem;">${cor.icon}</span> ${mensagem}`;
    
    const style = document.createElement('style');
    style.textContent = `
        @keyframes slideInRight {
            from { transform: translateX(100%); opacity: 0; }
            to { transform: translateX(0); opacity: 1; }
        }
        @keyframes slideOutRight {
            from { transform: translateX(0); opacity: 1; }
            to { transform: translateX(100%); opacity: 0; }
        }
    `;
    document.head.appendChild(style);
    
    document.body.appendChild(notificacao);
    
    setTimeout(() => {
        notificacao.style.animation = 'slideOutRight 0.3s ease forwards';
        setTimeout(() => notificacao.remove(), 300);
    }, 4000);
}

// ========================================
// EXPORTAÇÃO DE DADOS
// ========================================

function exportarDados() {
    let csv = 'PRODUTO;';

    const valorPorProduto = {};
    (estoque.registroVendas || []).forEach(venda => {
        if (Array.isArray(venda.items) && venda.items.length > 0) {
            venda.items.forEach(it => {
                const nome = it.produtoNome || '';
                if (!nome) return;
                const valorItem = typeof it.valorTotal === 'number'
                    ? it.valorTotal
                    : ((Number(it.valorUnitario) || 0) * (Number(it.quantidade) || 0));
                valorPorProduto[nome] = (valorPorProduto[nome] || 0) + valorItem;
            });
        } else {
            const nome = venda.produtoNome || '';
            if (!nome) return;
            const valorLegacy = typeof venda.valorTotal === 'number'
                ? venda.valorTotal
                : ((Number(venda.valorUnitario) || 0) * (Number(venda.quantidade) || 0));
            valorPorProduto[nome] = (valorPorProduto[nome] || 0) + valorLegacy;
        }
    });
    
    estoque.representantes.forEach(rep => {
        csv += `${rep} Disp;${rep} Venda;${rep} Saldo;`;
    });
    csv += 'GERAL Disp;GERAL Venda;GERAL Saldo;VALOR TOTAL VENDAS\n';
    
    estoque.produtos.forEach(produto => {
        csv += `"${produto.nome}";`;
        
        let geralDisp = 0, geralVenda = 0;
        
        estoque.representantes.forEach(rep => {
            const disp = produto.distribuicao[rep] || 0;
            const venda = produto.vendas[rep] || 0;
            const saldo = disp - venda;
            
            geralDisp += disp;
            geralVenda += venda;
            
            csv += `${disp};${venda};${saldo};`;
        });
        
        const valorVendas = valorPorProduto[produto.nome] || 0;
        csv += `${geralDisp};${geralVenda};${geralDisp - geralVenda};${valorVendas.toFixed(2).replace('.', ',')}\n`;
    });
    
    const blob = new Blob(['\ufeff' + csv], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    const url = URL.createObjectURL(blob);
    
    const dataAtual = new Date().toISOString().split('T')[0];
    
    link.setAttribute('href', url);
    link.setAttribute('download', `estoque_material_belico_${dataAtual}.csv`);
    link.style.visibility = 'hidden';
    
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    
    mostrarNotificacao('Arquivo exportado com sucesso!', 'success');
}

// ========================================
// EXPORTAR/IMPORTAR ESTOQUE COMPLETO
// ========================================

function exportarEstoqueCompleto() {
    const sep = ';';
    
    // Cabeçalho simples: Produto, Quantidade Total
    let csv = `PRODUTO${sep}QUANTIDADE_TOTAL\n`;
    
    // Dados
    estoque.produtos.forEach(produto => {
        // Quantidade total em estoque baseada no saldo consolidado cadastrado
        const totalEstoque = Number(produto.estoqueConsolidado) || 0;
        
        csv += `${produto.nome}${sep}${totalEstoque}\n`;
    });
    
    const blob = new Blob(['\ufeff' + csv], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    const url = URL.createObjectURL(blob);
    
    const dataAtual = new Date().toISOString().split('T')[0];
    
    link.setAttribute('href', url);
    link.setAttribute('download', `estoque_${dataAtual}.csv`);
    link.style.visibility = 'hidden';
    
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    
    mostrarNotificacao('Estoque exportado com sucesso!', 'success');
}

// ========================================
// EXPORTAR / IMPORTAR SISTEMA COMPLETO (JSON)
// ========================================

function exportarSistema() {
    try {
        const payload = {
            ...estoque,
            precificacao,
            tabelaAliquotas,
            tabelaICMS,
            categoriaPorProduto,
            impostosEditaveis: impostosEditaveis || {},
            icmsEditavelPJ: icmsEditavelPJ || {},
            icmsEditavelPF: icmsEditavelPF || {}
        };
        const dataStr = JSON.stringify(payload, null, 2);
        const blob = new Blob([dataStr], { type: 'application/json' });
        const link = document.createElement('a');
        const url = URL.createObjectURL(blob);
        const dataAtual = new Date().toISOString().split('T')[0];

        link.setAttribute('href', url);
        link.setAttribute('download', `controle_estoque_full_${dataAtual}.json`);
        link.style.visibility = 'hidden';

        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        mostrarNotificacao('Exportação do sistema concluída!', 'success');
    } catch (error) {
        console.error('Erro ao exportar sistema:', error);
        mostrarNotificacao('Erro ao exportar o sistema.', 'error');
    }
}

function importarSistema(event) {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const conteudo = e.target.result;
            const obj = JSON.parse(conteudo);

            // Validações básicas
            if (!obj || !Array.isArray(obj.produtos) || !Array.isArray(obj.representantes)) {
                mostrarNotificacao('Arquivo inválido: formato JSON inesperado.', 'error');
                event.target.value = '';
                return;
            }

            if (!confirm('⚠️ Importar o arquivo substituirá TODO o estado do sistema atual (produtos, distribuições e vendas). Deseja continuar?')) {
                event.target.value = '';
                return;
            }

            // Substitui o estado em memória e persiste
            estoque = obj;
            if (!Array.isArray(estoque.clientes)) estoque.clientes = [];
            clientes = estoque.clientes;
            if (!Array.isArray(estoque.propostas)) estoque.propostas = [];
            propostas = estoque.propostas;
            if (estoque.precificacao && typeof estoque.precificacao === 'object') precificacao = estoque.precificacao;
            else { estoque.precificacao = {}; precificacao = estoque.precificacao; }

            precificacao = (obj.precificacao && typeof obj.precificacao === 'object')
                ? obj.precificacao
                : (estoque.precificacao || {});
            tabelaAliquotas = (obj.tabelaAliquotas && typeof obj.tabelaAliquotas === 'object')
                ? obj.tabelaAliquotas
                : (estoque.tabelaAliquotas || {});
            tabelaICMS = Array.isArray(obj.tabelaICMS)
                ? obj.tabelaICMS
                : (Array.isArray(estoque.tabelaICMS) ? estoque.tabelaICMS : []);
            categoriaPorProduto = (obj.categoriaPorProduto && typeof obj.categoriaPorProduto === 'object')
                ? obj.categoriaPorProduto
                : (estoque.categoriaPorProduto || {});

            // Restaurar tabelas de impostos editáveis do arquivo importado
            impostosEditaveis = obj.impostosEditaveis || {};
            icmsEditavelPJ = obj.icmsEditavelPJ || {};
            icmsEditavelPF = obj.icmsEditavelPF || {};
            try { inicializarImpostosEditaveis(); } catch (e) {}
            try { inicializarICMSEditavel(); } catch (e) {}

            salvarDados();

            // Re-renderizar tudo
            renderizarTabela();
            renderizarDashboard();
            renderizarRegistroVendas();
            renderizarRegistroDistribuicao();
            renderizarClientes();
            atualizarKPIsClientes();
            atualizarDatalistClientes();
            renderizarPropostas();
            atualizarKPIsPropostas();
            renderizarPrecificacao();
            atualizarSelectsProdutos();
            atualizarEstatisticas();

            mostrarNotificacao('Importação do sistema concluída com sucesso!', 'success');
        } catch (error) {
            console.error('Erro ao importar sistema:', error);
            mostrarNotificacao('Erro ao processar o arquivo JSON. Verifique o formato.', 'error');
        } finally {
            event.target.value = '';
        }
    };

    reader.readAsText(file, 'UTF-8');
}


function exportarBackupCompleto() {
    try {
        const backup = {
            versao: '2.0',
            data: new Date().toISOString(),
            dados: estoque,
            precificacoesCliente: precificacoesCliente,
            impostosEditaveis: impostosEditaveis,
            icmsEditavelPJ: icmsEditavelPJ,
            icmsEditavelPF: icmsEditavelPF
        };
        const blob = new Blob([JSON.stringify(backup, null, 2)], { type: 'application/json' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'backup_FI_' + new Date().toISOString().split('T')[0] + '.json';
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
        mostrarNotificacao('Backup exportado com sucesso!', 'success');
    } catch (e) {
        console.error('Erro ao exportar backup completo:', e);
        mostrarNotificacao('Erro ao exportar backup: ' + (e.message || e), 'error');
    }
}

function importarBackup(event) {
    const file = event.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = e => {
        try {
            const backup = JSON.parse(e.target.result);
            if (!backup || !backup.dados || !backup.versao) throw new Error('Formato inválido');
            if (!confirm('Importar backup de ' + new Date(backup.data).toLocaleDateString('pt-BR') + '?\nIsso substituirá TODOS os dados atuais.')) {
                event.target.value = ''; 
                return;
            }
            estoque = backup.dados;
            precificacoesCliente = backup.precificacoesCliente || [];
            impostosEditaveis    = backup.impostosEditaveis    || {};
            icmsEditavelPJ       = backup.icmsEditavelPJ       || {};
            icmsEditavelPF       = backup.icmsEditavelPF       || {};
            try { salvarDados(); } catch (e) {}
            location.reload();
        } catch (err) {
            console.error('Erro ao importar backup:', err);
            mostrarNotificacao('Erro ao importar backup: ' + (err.message || err), 'error');
        } finally {
            event.target.value = '';
        }
    };
    reader.readAsText(file);
}


function importarEstoque(event) {
    const file = event.target.files[0];
    if (!file) return;
    
    const reader = new FileReader();
    
    reader.onload = function(e) {
        try {
            const conteudo = e.target.result;
            const linhas = conteudo.split(/\r?\n/).filter(l => l.trim());
            
            if (linhas.length < 2) {
                mostrarNotificacao('Arquivo vazio ou sem dados!', 'error');
                return;
            }
            
            // Ler cabeçalho para identificar colunas
            const cabecalho = parseCsvLinha(linhas[0]);
            console.debug('Cabeçalho:', cabecalho);
            
            let produtosAtualizados = 0;
            let erros = [];
            
            // Processar cada linha
            for (let i = 1; i < linhas.length; i++) {
                const linha = linhas[i].trim();
                if (!linha) continue;
                
                const colunas = parseCsvLinha(linha);
                
                if (colunas.length < 2) {
                    erros.push(`Linha ${i + 1}: formato inválido`);
                    continue;
                }
                
                const produtoNome = colunas[0]?.trim().toUpperCase();
                
                if (!produtoNome) continue;
                
                // Buscar produto
                let produto = estoque.produtos.find(p => 
                    p.nome.toUpperCase() === produtoNome ||
                    p.nome.toUpperCase().includes(produtoNome) ||
                    produtoNome.includes(p.nome.toUpperCase())
                );
                
                if (!produto) {
                    erros.push(`Linha ${i + 1}: produto não encontrado (${produtoNome})`);
                    continue;
                }
                
                // Atualizar quantidade total no estoque IMBEL
                // Formato novo: PRODUTO;QUANTIDADE_TOTAL
                // Compatibilidade: se vier formato antigo (PRODUTO;PRECO;QUANTIDADE_TOTAL), usa a 3a coluna
                const qtdCol = (colunas.length >= 3) ? colunas[2] : colunas[1];
                if (qtdCol !== undefined) {
                    const quantidade = parseInt(qtdCol) || 0;
                    // Atualizar o saldo consolidado (campo de cadastro)
                    produto.estoqueConsolidado = quantidade;
                }
                
                produtosAtualizados++;
            }
            
            if (produtosAtualizados > 0) {
                salvarDados();
                renderizarTabela();
                renderizarDashboard();
                renderizarRegistroVendas();
                renderizarRegistroDistribuicao();
                atualizarEstatisticas();
            }
            
            // Limpar input
            event.target.value = '';
            
            // Mostrar resultado
            if (erros.length > 0 && produtosAtualizados === 0) {
                mostrarNotificacao(`Nenhum produto atualizado. Verifique o formato.`, 'error');
                console.debug('Erros de importação:', erros);
            } else if (erros.length > 0) {
                mostrarNotificacao(`${produtosAtualizados} produtos atualizados. ${erros.length} erros.`, 'warning');
                console.debug('Erros de importação:', erros);
            } else {
                mostrarNotificacao(`${produtosAtualizados} produtos atualizados com sucesso!`, 'success');
            }
            
        } catch (error) {
            console.error('Erro ao importar:', error);
            mostrarNotificacao('Erro ao processar o arquivo. Verifique o formato.', 'error');
        }
    };
    
    reader.readAsText(file, 'UTF-8');
}

// ========================================
// IMPORTAÇÃO DE VENDAS
// ========================================

function importarVendas(event) {
    const file = event.target.files[0];
    if (!file) return;
    
    const reader = new FileReader();
    
    reader.onload = function(e) {
        try {
            const conteudo = e.target.result;
            console.debug('Conteúdo do arquivo:', conteudo); // Debug
            
            const linhas = conteudo.split(/\r?\n/).filter(l => l.trim());
            
            console.debug('Linhas encontradas:', linhas.length); // Debug
            console.debug('Primeira linha:', linhas[0]); // Debug
            if (linhas[1]) console.debug('Segunda linha:', linhas[1]); // Debug
            
            // Pular cabeçalho
            if (linhas.length < 2) {
                mostrarNotificacao('Arquivo vazio ou sem dados!', 'error');
                return;
            }
            
            let vendasImportadas = 0;
            let erros = [];
            
            // Processar cada linha (começando da segunda - pular cabeçalho)
            for (let i = 1; i < linhas.length; i++) {
                const linha = linhas[i].trim();
                if (!linha || linha.toLowerCase().includes('total')) continue;
                
                // Parse do CSV com suporte a aspas
                const colunas = parseCsvLinha(linha);
                
                console.debug(`Linha ${i + 1} - Colunas:`, colunas); // Debug
                
                // Pular linha se primeira coluna estiver vazia (linha de total ou vazia)
                if (!colunas[0] || colunas[0].trim() === '') continue;
                
                if (colunas.length < 5) {
                    erros.push(`Linha ${i + 1}: formato inválido (${colunas.length} colunas encontradas)`);
                    continue;
                }
                
                const contrato = colunas[0]?.trim();
                const loja = colunas[1]?.trim().replace(/"/g, '').toUpperCase();
                const representante = colunas[2]?.trim().toUpperCase();
                const produtoNome = colunas[3]?.trim().replace(/"/g, '').toUpperCase();
                const quantidade = parseInt(colunas[4]?.trim()) || 0;
                const observacoes = colunas[7]?.trim()?.replace(/"/g, '') || '';
                
                console.debug(`Dados: contrato=${contrato}, loja=${loja}, rep=${representante}, produto=${produtoNome}, qtd=${quantidade}`); // Debug
                
                if (!contrato || !loja || !representante || !produtoNome || quantidade <= 0) {
                    erros.push(`Linha ${i + 1}: dados obrigatórios faltando (contrato=${contrato}, loja=${loja}, rep=${representante}, produto=${produtoNome}, qtd=${quantidade})`);
                    continue;
                }
                
                // Verificar se representante é válido
                const repsValidos = ['KOLTE', 'ISA', 'LC', 'ADES', 'FL', 'IMBEL'];
                if (!repsValidos.includes(representante)) {
                    erros.push(`Linha ${i + 1}: representante inválido (${representante})`);
                    continue;
                }
                
                // Buscar produto pelo nome (correspondência parcial)
                let produto = estoque.produtos.find(p => 
                    p.nome.toUpperCase() === produtoNome || 
                    p.nome.toUpperCase().includes(produtoNome) ||
                    produtoNome.includes(p.nome.toUpperCase())
                );
                
                if (!produto) {
                    erros.push(`Linha ${i + 1}: produto não encontrado (${produtoNome})`);
                    continue;
                }
                
                // Valor deve vir do registro da venda (arquivo), sem vínculo com cadastro de produto
                const valorUnitario = parseFloat((colunas[5] || '0').toString().replace(/\./g, '').replace(',', '.')) || 0;
                const valorTotalArquivo = parseFloat((colunas[6] || '0').toString().replace(/\./g, '').replace(',', '.')) || 0;
                const valorTotal = valorTotalArquivo > 0 ? valorTotalArquivo : (valorUnitario * quantidade);
                
                // Criar registro da venda
                const novaVenda = {
                    id: Date.now() + i,
                    contrato: contrato,
                    loja: loja,
                    representante: representante,
                    produtoId: produto.id,
                    produtoNome: produto.nome,
                    quantidade: quantidade,
                    valorUnitario: valorUnitario,
                    valorTotal: valorTotal,
                    observacoes: observacoes,
                    data: new Date().toISOString()
                };
                
                // Apenas registrar a venda (NÃO mexer na distribuição)
                // A distribuição deve ser feita separadamente na aba Distribuição
                produto.vendas[representante] = (produto.vendas[representante] || 0) + quantidade;
                
                // Adicionar ao registro
                estoque.registroVendas.push(novaVenda);
                vendasImportadas++;
            }
            
            if (vendasImportadas > 0) {
                salvarDados();
                renderizarTabela();
                renderizarDashboard();
                renderizarRegistroVendas();
            }
            
            // Limpar input
            event.target.value = '';
            
            // Mostrar resultado
            if (erros.length > 0 && vendasImportadas === 0) {
                mostrarNotificacao(`Nenhuma venda importada. Verifique o formato do arquivo.`, 'error');
                console.debug('Erros de importação:', erros);
            } else if (erros.length > 0) {
                mostrarNotificacao(`${vendasImportadas} vendas importadas. ${erros.length} linhas com erro.`, 'warning');
                console.debug('Erros de importação:', erros);
            } else {
                mostrarNotificacao(`${vendasImportadas} vendas importadas com sucesso!`, 'success');
            }
            
        } catch (error) {
            console.error('Erro ao importar:', error);
            mostrarNotificacao('Erro ao processar o arquivo. Verifique o formato.', 'error');
        }
    };
    
    reader.readAsText(file, 'UTF-8');
}

function parseCsvLinha(linha) {
    const resultado = [];
    let atual = '';
    let dentroAspas = false;
    
    // Detectar separador (TAB, ponto-e-vírgula ou vírgula)
    let separador = '\t';
    if (linha.includes('\t')) {
        separador = '\t';
    } else if (linha.includes(';')) {
        separador = ';';
    } else if (linha.includes(',')) {
        separador = ',';
    }
    
    for (let i = 0; i < linha.length; i++) {
        const char = linha[i];
        
        if (char === '"') {
            dentroAspas = !dentroAspas;
        } else if (char === separador && !dentroAspas) {
            resultado.push(atual.trim());
            atual = '';
        } else {
            atual += char;
        }
    }
    resultado.push(atual.trim());
    
    return resultado;
}

// ========================================
// LIMPAR DADOS DO SISTEMA
// ========================================

function limparTodosDados() {
    if (!confirm('⚠️ ATENÇÃO!\n\nEsta ação irá APAGAR TODOS os dados do sistema:\n- Produtos\n- Distribuições\n- Vendas\n- Registro de vendas\n- Registro de distribuição\n\nOs dados serão resetados para os valores iniciais.\n\nDeseja continuar?')) {
        return;
    }
    
    if (!confirm('ÚLTIMA CONFIRMAÇÃO:\n\nVocê tem certeza absoluta? Esta ação não pode ser desfeita!')) {
        return;
    }
    
    // Remover do localStorage
    localStorage.removeItem('estoqueArmasV2');
    
    // Recarregar dados iniciais
    estoque.produtos = dadosIniciais.map((item, index) => ({
        id: index + 1,
        nome: item.nome,
        preco: item.preco,
        distribuicao: { KOLTE: 0, ISA: 0, LC: 0, ADES: 0, FL: 0, IMBEL: 0 },
        vendas: { KOLTE: 0, ISA: 0, LC: 0, ADES: 0, FL: 0, IMBEL: 0 }
    }));
    estoque.registroVendas = [];
    estoque.registroDistribuicao = [];
    estoque.controleEnvio = {};
    estoque.clientes = [];
    clientes = estoque.clientes;
    estoque.propostas = [];
    propostas = estoque.propostas;
    estoque.precificacao = {};
    precificacoesCliente = [];
    estoque.precificacoesCliente = [];
    estoque.precificacaoConfig = null;
    precificacao = estoque.precificacao;
    tabelaAliquotas = {};
    tabelaICMS = [];
    categoriaPorProduto = {};
    // Limpar tabelas de impostos editáveis
    impostosEditaveis = {};
    icmsEditavelPJ = {};
    icmsEditavelPF = {};
    
    salvarDados();
    renderizarTabela();
    renderizarDashboard();
    renderizarRegistroVendas();
    renderizarRegistroDistribuicao();
    renderizarControleEnvio();
    renderizarClientes();
    atualizarKPIsClientes();
    atualizarDatalistClientes();
    renderizarPropostas();
    atualizarKPIsPropostas();
    renderizarPrecificacao();
    atualizarSelectsProdutos();
    atualizarEstatisticas();
    
    mostrarNotificacao('Todos os dados foram apagados!', 'success');
}

// ========================================
// CONTROLE DE ENVIO DE CONTRATOS
// ========================================

function campoMarcado(valor) {
    return valor === true || valor === 'Sim' || valor === 'SAP' || valor === 'Outro' || valor === 'true';
}

function renderizarControleEnvio() {
    const tbody = document.getElementById('tabelaControleEnvioBody');
    if (!tbody) return;

    const filtroRep = document.getElementById('filtroControleEnvioRep')?.value || '';
    const filtroSistema = document.getElementById('filtroControleEnvioSistema')?.value || '';
    const filtroAssinado = document.getElementById('filtroControleEnvioAssinado')?.value || '';
    const filtroEnviado = document.getElementById('filtroControleEnvioEnviado')?.value || '';

    // Agrupa vendas por contrato (pega a primeira ocorrência de cada contrato)
    const contratoMap = {};
    
    estoque.registroVendas.forEach(venda => {
        if (!contratoMap[venda.contrato]) {
            contratoMap[venda.contrato] = {
                contrato: venda.contrato,
                loja: venda.loja,
                representante: venda.representante,
                id: venda.id
            };
        }
    });

    let contratos = Object.values(contratoMap);

    if (filtroRep) {
        contratos = contratos.filter(c => c.representante === filtroRep);
    }

    if (filtroSistema || filtroAssinado || filtroEnviado) {
        contratos = contratos.filter(c => {
            const envio = estoque.controleEnvio[c.contrato] || {};
            const sistemaMarcado = campoMarcado(envio.sistema);
            const assinadoMarcado = campoMarcado(envio.assinado);
            const enviadoMarcado = campoMarcado(envio.enviado);

            if (filtroSistema === 'sim' && !sistemaMarcado) return false;
            if (filtroSistema === 'nao' && sistemaMarcado) return false;
            if (filtroAssinado === 'sim' && !assinadoMarcado) return false;
            if (filtroAssinado === 'nao' && assinadoMarcado) return false;
            if (filtroEnviado === 'sim' && !enviadoMarcado) return false;
            if (filtroEnviado === 'nao' && enviadoMarcado) return false;

            return true;
        });
    }

    contratos = contratos.sort((a, b) => {
        const contratoA = parseInt(a.contrato) || 0;
        const contratoB = parseInt(b.contrato) || 0;
        return contratoA - contratoB;
    });

    tbody.innerHTML = '';

    if (contratos.length === 0) {
        tbody.innerHTML = `
            <tr>
                <td colspan="8" class="empty-state">
                    <div class="empty-icon">📮</div>
                    <div class="empty-text">Nenhum contrato registrado</div>
                    <div class="empty-hint">Registre vendas na aba "Registro de Vendas" para aparecerem aqui</div>
                </td>
            </tr>
        `;
        return;
    }

    contratos.forEach(contrato => {
        const envio = estoque.controleEnvio[contrato.contrato] || {};
        const sistemaMarcado = campoMarcado(envio.sistema);
        const assinadoMarcado = campoMarcado(envio.assinado);
        const enviadoMarcado = campoMarcado(envio.enviado);
        const repClass = (contrato.representante || '').toLowerCase();
        const concluidos = Number(sistemaMarcado) + Number(assinadoMarcado) + Number(enviadoMarcado);

        const statusBtn = (checked, campo) => `
            <button type="button" class="status-indicator ${checked ? 'checked' : ''}" onclick="salvarControleEnvio('${contrato.contrato}', '${campo}', ${!checked})" title="${campo}">
                <svg viewBox="0 0 12 12" aria-hidden="true"><path fill="white" d="M4.7 9.2 1.9 6.4l1.1-1.1 1.7 1.7 4.2-4.2L10 4z"/></svg>
            </button>
        `;

        const tr = document.createElement('tr');
        if (concluidos === 3) tr.classList.add('row-envio-completo');
        tr.innerHTML = `
            <td class="col-contrato">
                <div class="ctr-cell">
                    <span>${contrato.contrato}</span>
                    <span class="ctr-progress">${concluidos}/3</span>
                </div>
            </td>
            <td class="col-loja" title="${contrato.loja}">${contrato.loja}</td>
            <td class="col-representante"><span class="badge-rep ${repClass}">${contrato.representante}</span></td>
            <td class="col-sistema">
                ${statusBtn(sistemaMarcado, 'sistema')}
            </td>
            <td class="col-assinado">
                ${statusBtn(assinadoMarcado, 'assinado')}
            </td>
            <td class="col-enviado">
                ${statusBtn(enviadoMarcado, 'enviado')}
            </td>
            <td class="col-solicitacao">
                <input type="text" class="campo-editavel" value="${envio.solicitacao || ''}" placeholder="Data ou observação" onchange="salvarControleEnvio('${contrato.contrato}', 'solicitacao', this.value)">
            </td>
            <td class="col-acoes">
                <button class="btn-action btn-delete" onclick="limparControleEnvio('${contrato.contrato}')" title="Limpar dados">🗑</button>
            </td>
        `;
        tbody.appendChild(tr);
    });
}

function salvarControleEnvio(contrato, campo, valor) {
    if (!estoque.controleEnvio[contrato]) {
        estoque.controleEnvio[contrato] = {};
    }

    if (campo === 'sistema' || campo === 'assinado' || campo === 'enviado') {
        valor = Boolean(valor);
    }
    
    estoque.controleEnvio[contrato][campo] = valor;
    salvarDados();
    // Re-renderizar a tabela para refletir imediatamente a alteração visual
    try { renderizarControleEnvio(); } catch (e) {}
}

function limparFiltrosControleEnvio() {
    const filtroRep = document.getElementById('filtroControleEnvioRep');
    const filtroSistema = document.getElementById('filtroControleEnvioSistema');
    const filtroAssinado = document.getElementById('filtroControleEnvioAssinado');
    const filtroEnviado = document.getElementById('filtroControleEnvioEnviado');

    if (filtroRep) filtroRep.value = '';
    if (filtroSistema) filtroSistema.value = '';
    if (filtroAssinado) filtroAssinado.value = '';
    if (filtroEnviado) filtroEnviado.value = '';

    renderizarControleEnvio();
}

function limparControleEnvio(contrato) {
    if (confirm(`Deseja limpar os dados de envio do contrato ${contrato}?`)) {
        delete estoque.controleEnvio[contrato];
        salvarDados();
        renderizarControleEnvio();
        mostrarNotificacao(`Dados de envio do contrato ${contrato} removidos`, 'success');
    }
}

function exportarControleEnvio() {
    const contratoMap = {};
    
    estoque.registroVendas.forEach(venda => {
        if (!contratoMap[venda.contrato]) {
            contratoMap[venda.contrato] = {
                contrato: venda.contrato,
                loja: venda.loja,
                representante: venda.representante
            };
        }
    });

    const contratos = Object.values(contratoMap).sort((a, b) => {
        const contratoA = parseInt(a.contrato) || 0;
        const contratoB = parseInt(b.contrato) || 0;
        return contratoA - contratoB;
    });

    const dados = contratos.map(c => {
        const envio = estoque.controleEnvio[c.contrato] || {};
        const sistemaMarcado = campoMarcado(envio.sistema);
        return {
            'CTR': c.contrato,
            'NOME': c.loja,
            'REPRESENTANTE': c.representante,
            'SISTEMA': sistemaMarcado ? 'Sim' : 'Não',
            'ASSINADO': envio.assinado ? 'Sim' : 'Não',
            'ENVIADO': envio.enviado ? 'Sim' : 'Não',
            'SOLICITAÇÃO': envio.solicitacao || ''
        };
    });

    const csv = gerarCSV(dados);
    const blob = new Blob(['\ufeff' + csv], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = `Controle_Envio_${new Date().toLocaleDateString('pt-BR').replace(/\//g, '-')}.csv`;
    link.click();
}

function gerarCSV(dados) {
    if (!dados || dados.length === 0) return '';
    
    const headers = Object.keys(dados[0]);
    const csv = [
        headers.join(';'),
        ...dados.map(row => headers.map(header => `"${(row[header] || '').toString().replace(/"/g, '""')}"`).join(';'))
    ].join('\n');
    
    return csv;
}

// ========================================
// INICIALIZAÇÃO
// ========================================

document.addEventListener('DOMContentLoaded', inicializar);

// ========================================
// BUSCA / FILTRO NA TABELA DE ESTOQUE
// ========================================

function filtrarTabelaEstoque(termo) {
    const tbody = document.getElementById('corpoTabela');
    if (!tbody) return;
    const rows = tbody.querySelectorAll('tr:not(.total-row)');
    const termoLower = (termo || '').toLowerCase().trim();

    rows.forEach(row => {
        const nome = (row.querySelector('.produto-nome')?.textContent || '').toLowerCase();
        row.style.display = (!termoLower || nome.includes(termoLower)) ? '' : 'none';
    });
}

// ========================================
// BUSCA GLOBAL
// ========================================

let _buscaGlobalTimer = null;
function executarBuscaGlobal(termo) {
    if (_buscaGlobalTimer) clearTimeout(_buscaGlobalTimer);
    _buscaGlobalTimer = setTimeout(() => {
        _executarBuscaGlobalReal(termo);
    }, 400);
}

function _executarBuscaGlobalReal(termo) {
    if (!termo || termo.trim().length < 2) return;
    const t = termo.toLowerCase().trim();
    const resultados = { produtos: [], vendas: [], distribuicoes: [], contratos: [] };

    // Buscar em produtos
    estoque.produtos.forEach(p => {
        if (p.nome.toLowerCase().includes(t)) {
            resultados.produtos.push(p);
        }
    });

    // Buscar em vendas
    (estoque.registroVendas || []).forEach(v => {
        const match = (v.contrato || '').toLowerCase().includes(t) ||
                      (v.loja || '').toLowerCase().includes(t) ||
                      (v.representante || '').toLowerCase().includes(t);
        if (match) resultados.vendas.push(v);
    });

    // Buscar em distribuições
    (estoque.registroDistribuicao || []).forEach(d => {
        const match = (d.produtoNome || '').toLowerCase().includes(t) ||
                      (d.representante || '').toLowerCase().includes(t);
        if (match) resultados.distribuicoes.push(d);
    });

    const container = document.getElementById('resultadosBuscaGlobal');
    if (!container) return;

    let html = '';
    const total = resultados.produtos.length + resultados.vendas.length + resultados.distribuicoes.length;

    if (total === 0) {
        html = '<p style="text-align:center;color:var(--text-secondary);padding:20px">Nenhum resultado encontrado.</p>';
    } else {
        if (resultados.produtos.length > 0) {
            html += '<div class="busca-categoria"><h4>📦 Produtos (' + resultados.produtos.length + ')</h4>';
            resultados.produtos.forEach(p => {
                html += `<div class="busca-resultado-item" onclick="fecharModal('modalBuscaGlobal');trocarAba('estoque')"><span>${p.nome}</span><span class="resultado-aba">Estoque</span></div>`;
            });
            html += '</div>';
        }
        if (resultados.vendas.length > 0) {
            html += '<div class="busca-categoria"><h4>📝 Vendas (' + resultados.vendas.length + ')</h4>';
            resultados.vendas.slice(0, 20).forEach(v => {
                html += `<div class="busca-resultado-item" onclick="fecharModal('modalBuscaGlobal');trocarAba('vendas')"><span>CTR ${v.contrato} — ${v.loja} (${v.representante})</span><span class="resultado-aba">Vendas</span></div>`;
            });
            html += '</div>';
        }
        if (resultados.distribuicoes.length > 0) {
            html += '<div class="busca-categoria"><h4>🚚 Distribuições (' + resultados.distribuicoes.length + ')</h4>';
            resultados.distribuicoes.slice(0, 20).forEach(d => {
                html += `<div class="busca-resultado-item" onclick="fecharModal('modalBuscaGlobal');trocarAba('distribuicao')"><span>${d.produtoNome} → ${d.representante} (${d.quantidade})</span><span class="resultado-aba">Distribuição</span></div>`;
            });
            html += '</div>';
        }
    }

    container.innerHTML = html;
    document.getElementById('modalBuscaGlobal').style.display = 'flex';
}

// ========================================
// SIDEBAR RESPONSIVA (DESKTOP + MOBILE)
// ========================================

function toggleSidebarExpanded() {
    document.body.classList.toggle('sidebar-expanded');
}

function toggleMobileSidebar(forceOpen) {
    const shouldOpen = typeof forceOpen === 'boolean'
        ? forceOpen
        : !document.body.classList.contains('mobile-sidebar-open');
    document.body.classList.toggle('mobile-sidebar-open', shouldOpen);
}

// Compatibilidade com chamadas antigas
function toggleMenuMobile() {
    toggleMobileSidebar();
}

// Fechar drawer ao trocar aba (mobile)
const _trocarAbaOriginal = trocarAba;
trocarAba = function(aba) {
    _trocarAbaOriginal(aba);
    try {
        document.body.classList.remove('mobile-sidebar-open');
    } catch(e) {}
};

// ========================================
// ORDENAÇÃO CLICÁVEL NAS COLUNAS
// ========================================

let _ordenVendas = { campo: 'contrato', direcao: 'asc' };
let _ordenDistribuicao = { campo: 'data', direcao: 'desc' };
let _contratosExpandidos = {};

function ordenarVendas(campo) {
    if (_ordenVendas.campo === campo) {
        _ordenVendas.direcao = _ordenVendas.direcao === 'asc' ? 'desc' : 'asc';
    } else {
        _ordenVendas.campo = campo;
        _ordenVendas.direcao = 'asc';
    }
    // Atualizar ícones
    document.querySelectorAll('#tabelaRegistroVendas th.sortable').forEach(th => {
        th.classList.remove('sort-asc', 'sort-desc');
        if (th.dataset.sort === campo) th.classList.add('sort-' + _ordenVendas.direcao);
    });
    renderizarRegistroVendas();
}

function ordenarDistribuicao(campo) {
    if (_ordenDistribuicao.campo === campo) {
        _ordenDistribuicao.direcao = _ordenDistribuicao.direcao === 'asc' ? 'desc' : 'asc';
    } else {
        _ordenDistribuicao.campo = campo;
        _ordenDistribuicao.direcao = 'asc';
    }
    document.querySelectorAll('#tabelaRegistroDistribuicao th.sortable').forEach(th => {
        th.classList.remove('sort-asc', 'sort-desc');
        if (th.dataset.sort === campo) th.classList.add('sort-' + _ordenDistribuicao.direcao);
    });
    renderizarRegistroDistribuicao();
}

// ========================================
// PAGINAÇÃO
// ========================================

const ITENS_POR_PAGINA_OPCOES = [15, 30, 50, 100];
let _paginaVendas = 1;
let _itensPorPaginaVendas = 30;
let _paginaDistribuicao = 1;
let _itensPorPaginaDistribuicao = 30;

function renderizarPaginacao(containerId, paginaAtual, totalItens, itensPorPagina, onChangePage, onChangePerPage) {
    const container = document.getElementById(containerId);
    if (!container) return;

    const totalPaginas = Math.max(1, Math.ceil(totalItens / itensPorPagina));
    if (totalItens <= itensPorPagina) {
        container.innerHTML = '';
        return;
    }

    let html = '';
    html += `<button class="page-btn" ${paginaAtual <= 1 ? 'disabled' : ''} onclick="${onChangePage}(${paginaAtual - 1})">‹</button>`;

    const maxBtns = 5;
    let start = Math.max(1, paginaAtual - Math.floor(maxBtns / 2));
    let end = Math.min(totalPaginas, start + maxBtns - 1);
    if (end - start < maxBtns - 1) start = Math.max(1, end - maxBtns + 1);

    if (start > 1) html += `<button class="page-btn" onclick="${onChangePage}(1)">1</button><span class="page-info">...</span>`;

    for (let i = start; i <= end; i++) {
        html += `<button class="page-btn ${i === paginaAtual ? 'active' : ''}" onclick="${onChangePage}(${i})">${i}</button>`;
    }

    if (end < totalPaginas) html += `<span class="page-info">...</span><button class="page-btn" onclick="${onChangePage}(${totalPaginas})">${totalPaginas}</button>`;

    html += `<button class="page-btn" ${paginaAtual >= totalPaginas ? 'disabled' : ''} onclick="${onChangePage}(${paginaAtual + 1})">›</button>`;
    html += `<span class="page-info">${totalItens} registros</span>`;
    html += `<select onchange="${onChangePerPage}(parseInt(this.value))">`;
    ITENS_POR_PAGINA_OPCOES.forEach(n => {
        html += `<option value="${n}" ${n === itensPorPagina ? 'selected' : ''}>${n} por pág.</option>`;
    });
    html += '</select>';

    container.innerHTML = html;
}

function mudarPaginaVendas(p) { _paginaVendas = p; renderizarRegistroVendas(); }
function mudarItensPaginaVendas(n) { _itensPorPaginaVendas = n; _paginaVendas = 1; renderizarRegistroVendas(); }
function mudarPaginaDistribuicao(p) { _paginaDistribuicao = p; renderizarRegistroDistribuicao(); }
function mudarItensPaginaDistribuicao(n) { _itensPorPaginaDistribuicao = n; _paginaDistribuicao = 1; renderizarRegistroDistribuicao(); }

// ========================================
// NOTIFICAÇÕES DE ESTOQUE BAIXO
// ========================================

const LIMITE_ESTOQUE_BAIXO = 3;

function verificarEstoqueBaixo() {
    try {
        return verificarAlertasEstoque();
    } catch (e) {
        console.warn('verificarAlertasEstoque falhou', e);
        return 0;
    }
}

function salvarConfigAlertas() {
    try {
        const limEl = document.getElementById('limiteAlertaEstoque');
        const ativoEl = document.getElementById('alertasAtivos');
        const limite = limEl ? parseInt(limEl.value || '0', 10) : configAlertas.limite;
        const ativo = ativoEl ? !!ativoEl.checked : !!configAlertas.ativo;
        configAlertas.limite = isNaN(limite) ? configAlertas.limite : limite;
        configAlertas.ativo = ativo;
        try { localStorage.setItem('configAlertas', JSON.stringify(configAlertas)); } catch (e) {}
        mostrarNotificacao('Configurações de alertas salvas.', 'success');
        verificarAlertasEstoque();
    } catch (e) {
        console.warn('salvarConfigAlertas erro', e);
        mostrarNotificacao('Erro ao salvar configurações de alertas.', 'error');
    }
}

function togglePainelAlertas() {
    try {
        const painel = document.getElementById('painelAlertas');
        if (!painel) return;
        if (painel.style.display === 'block') painel.style.display = 'none';
        else { painel.style.display = 'block'; verificarAlertasEstoque(); }
    } catch (e) { console.warn('togglePainelAlertas erro', e); }
}

function fecharPainelAlertas() {
    const p = document.getElementById('painelAlertas');
    if (p) p.style.display = 'none';
}

function verificarAlertasEstoque() {
    try {
        const limite = (configAlertas && typeof configAlertas.limite === 'number') ? configAlertas.limite : LIMITE_ESTOQUE_BAIXO;
        const ativo = (configAlertas && typeof configAlertas.ativo !== 'undefined') ? !!configAlertas.ativo : true;
        const alertas = [];
        if (!Array.isArray(estoque.produtos) || !Array.isArray(estoque.representantes)) {
            const badge = document.getElementById('badgeAlertas'); if (badge) badge.style.display = 'none';
            const el = document.getElementById('alertaEstoqueBaixo'); if (el) el.style.display = 'none';
            const lista = document.getElementById('listaPainelAlertas'); if (lista) lista.innerHTML = '';
            return 0;
        }

        estoque.produtos.forEach(produto => {
            let totalDisp = 0, totalVenda = 0;
            const repsZerados = [];
            estoque.representantes.forEach(rep => {
                const disp = Number((produto.distribuicao && produto.distribuicao[rep]) || 0);
                const venda = Number((produto.vendas && produto.vendas[rep]) || 0);
                totalDisp += disp;
                totalVenda += venda;
                const saldoRep = disp - venda;
                if (saldoRep === 0) repsZerados.push(rep);
            });
            const saldoConsolidado = totalDisp - totalVenda;
            if (saldoConsolidado <= limite) {
                alertas.push({ id: produto.id, nome: produto.nome, saldoConsolidado, repsZerados });
            }
        });

        // Precificações expirando em até 3 dias
        const precifExpirando = (precificacoesCliente || []).filter(p => {
            if (!p || p.status !== 'ativa' || !p.dataExpiracao) return false;
            const dias = Math.ceil((new Date(p.dataExpiracao) - new Date()) / 86400000);
            return dias >= 0 && dias <= 3;
        });

        // Total de alertas inclui produtos + precificações expirando
        const totalCount = alertas.length + (precifExpirando ? precifExpirando.length : 0);

        // atualizar badge (produtos em alerta + precificações expirando)
        const badge = document.getElementById('badgeAlertas');
        if (badge) {
            if (totalCount > 0) { badge.style.display = 'inline-flex'; badge.textContent = String(totalCount); }
            else { badge.style.display = 'none'; }
        }

        const totalAlertasMsg = document.getElementById('totalAlertasMsg');
        if (totalAlertasMsg) {
            if (totalCount > 0) {
                totalAlertasMsg.textContent = `${totalCount} alerta(s) no total`;
            } else {
                totalAlertasMsg.textContent = 'Nenhum alerta no momento';
            }
        }

        // atualizar lista do painel
        const lista = document.getElementById('listaPainelAlertas');
        if (lista) {
            if (alertas.length === 0 && (!precifExpirando || precifExpirando.length === 0)) {
                lista.innerHTML = '<div style="color:#64748b">Nenhum alerta no momento.</div>';
            } else {
                let html = '';
                alertas.forEach(a => {
                    const label = (a.saldoConsolidado === 0) ? 'ZERADO' : ((a.saldoConsolidado <= Math.max(1, Math.floor(limite/2))) ? 'CRÍTICO' : 'BAIXO');
                    const color = label === 'ZERADO' ? '#ef4444' : (label === 'CRÍTICO' ? '#f59e0b' : '#fb923c');
                    const bg = label === 'ZERADO' ? '#fff1f2' : '#fff7ed';
                    html += `<div style="padding:10px; margin-bottom:8px; border-radius:8px; border-left:4px solid ${color}; background:${bg}">` +
                                `<div style="font-weight:700; color:#0f172a">${a.nome}</div>` +
                                `<div style="font-size:0.85rem; color:#334155; margin-top:6px">Saldo consolidado: <strong>${a.saldoConsolidado}</strong> un.</div>` +
                                (a.repsZerados && a.repsZerados.length ? `<div style="font-size:0.8rem; color:#64748b; margin-top:6px">Representantes sem saldo: ${a.repsZerados.join(', ')}</div>` : '') +
                              `</div>`;
                });
                lista.innerHTML = html;
                                // anexar alertas de precificações expirando/vence em até 3 dias
                if (precifExpirando && precifExpirando.length) {
                    precifExpirando.forEach(p => {
                        const dias = Math.ceil((new Date(p.dataExpiracao) - new Date()) / 86400000);
                        lista.insertAdjacentHTML('beforeend', `
                                                <div style="padding:10px; margin-bottom:8px; border-radius:8px; border-left:3px solid #f59e0b; background:#fffbeb">
                          <div style="font-weight:600; font-size:0.88rem; color:#1e293b">
                                                        ⚠️ Precificação expirando — ${p.clienteNome || ''}
                          </div>
                                                    <div style="font-size:0.78rem; color:#64748b; margin-top:3px">
                                                        v${p.versao || 1} expira em ${dias} dia(s)
                          </div>
                          <button onclick="trocarSubabaPrecif('porcliente')"
                                                                    style="font-size:0.75rem; color:#f59e0b; background:none; border:none; cursor:pointer; text-decoration:underline; padding:4px 0 0">
                            Ver precificação →
                          </button>
                        </div>
                      `);
                    });
                }
            }
        }

        // atualizar alerta pequeno na página
        const el = document.getElementById('alertaEstoqueBaixo');
        if (el) {
            if (!ativo || alertas.length === 0) {
                el.style.display = 'none';
                el.innerHTML = '';
            } else {
                const shortHtml = `<div style="padding:12px; background:#fff7ed; border-left:4px solid #fb923c; border-radius:8px;">` +
                                  `⚠️ <strong>${alertas.length}</strong> produto(s) abaixo do limite. <a href="#" onclick="togglePainelAlertas(); return false;">Ver detalhes</a>` +
                                  `</div>`;
                el.innerHTML = shortHtml;
                el.style.display = 'block';
            }
        }

        const agora = new Date();
        const alertasPropostas = [];
        let mudouStatusProposta = false;

        (propostas || []).forEach(proposta => {
            if (!proposta.dataExpiracao) return;
            if (!['enviada', 'rascunho'].includes(proposta.status)) return;

            const exp = new Date(proposta.dataExpiracao);
            const dias = Math.ceil((exp - agora) / 86400000);

            if (dias < 0 && proposta.status === 'enviada') {
                proposta.status = 'expirada';
                mudouStatusProposta = true;
                alertasPropostas.push({
                    proposta, cor: '#94a3b8', urgencia: 'EXPIRADA', tipo: 'expirada'
                });
            } else if (dias >= 0 && dias <= 3) {
                alertasPropostas.push({
                    proposta, diasRestantes: dias,
                    cor: dias === 0 ? '#dc2626' : '#d97706',
                    urgencia: dias === 0 ? 'VENCE HOJE'
                        : dias === 1 ? 'VENCE AMANHÃ'
                        : `Vence em ${dias} dias`,
                    tipo: 'expirando'
                });
            }
        });

        (propostas || []).filter(p => p.aguardandoAprovacao || p.status === 'aguardando_aprovacao')
            .forEach(p => alertasPropostas.push({
                proposta: p, cor: '#7c3aed', urgencia: 'AGUARD. APROVAÇÃO', tipo: 'aprovacao'
            }));

        // Propostas recusadas recentemente (até 7 dias) — avisar time para follow-up
        (propostas || []).forEach(p => {
            if (p && p.status === 'recusada' && p.dataRecusa) {
                const diasDesde = Math.ceil((agora - new Date(p.dataRecusa)) / 86400000);
                if (diasDesde >= 0 && diasDesde <= 7) {
                    alertasPropostas.push({ proposta: p, cor: '#ef4444', urgencia: 'RECUSADA', tipo: 'recusa' });
                }
            }
        });

        if (mudouStatusProposta) {
            try { estoque.propostas = propostas; } catch (e) {}
            try { salvarDados(); } catch (e) {}
        }

        if (alertasPropostas.length > 0) {
            const lista = document.getElementById('listaPainelAlertas');
            if (lista) {
                if ((lista.textContent || '').includes('Nenhum alerta no momento')) lista.innerHTML = '';

                const header = document.createElement('div');
                header.style.cssText = 'font-size:0.7rem;font-weight:700;color:#64748b;' +
                    'text-transform:uppercase;letter-spacing:0.8px;padding:8px 0 4px;' +
                    'margin-top:8px;border-top:1px solid #f1f5f9';
                header.textContent = `Propostas (${alertasPropostas.length})`;
                lista.appendChild(header);

                alertasPropostas.forEach(alerta => {
                    const p = alerta.proposta;
                    const fmt = v => 'R$' + (v || 0).toLocaleString('pt-BR', { minimumFractionDigits: 2 });
                    const div = document.createElement('div');
                    div.style.cssText = `padding:10px;margin-bottom:8px;border-radius:8px; border-left:3px solid ${alerta.cor};background:${alerta.cor}10`;
                                        div.innerHTML = `
                    <div style="display:flex;justify-content:space-between;align-items:flex-start">
                        <div style="font-weight:600;font-size:0.88rem;color:#1e293b">
                            ${alerta.tipo === 'aprovacao' ? '⏳' : (alerta.tipo === 'recusa' ? '❌' : '📋')} Proposta ${p.numero}
                        </div>
                        <span style="font-size:0.68rem;font-weight:700;color:${alerta.cor};
                                                 background:${alerta.cor}20;padding:1px 7px;
                                                 border-radius:20px;white-space:nowrap">
                            ${alerta.urgencia}
                        </span>
                    </div>
                    <div style="font-size:0.78rem;color:#64748b;margin-top:3px">
                        ${p.cliente} · ${fmt(p.valorTotal)}
                    </div>
                    ${alerta.tipo === 'aprovacao' && p.motivoAprovacao
                                                        ? `<div style="font-size:0.75rem;color:#7c3aed;margin-top:3px">Motivo: ${_escapeHtml(String(p.motivoAprovacao))}</div>` : ''}
                    ${alerta.tipo === 'recusa' && p.motivoRecusa
                                                        ? `<div style="font-size:0.75rem;color:#dc2626;margin-top:3px">Motivo: ${_escapeHtml(String(p.motivoRecusa))}</div>` : ''}
                    <button onclick="trocarAba('propostas');fecharPainelAlertas();"
                                    style="font-size:0.75rem;color:${alerta.cor};background:none;
                                                 border:none;cursor:pointer;text-decoration:underline;
                                                 padding:4px 0 0">
                        Ver proposta →
                    </button>`;
                    lista.appendChild(div);
                });
            }
        }

        // Add proposal alerts to total badge count
        const countSaldoZero = alertas.length;
        const totalAlertas = totalCount;
        const totalComPropostas = (totalAlertas || 0) + alertasPropostas.length;
        const badge2 = document.getElementById('badgeAlertas');
        if (badge2) {
            badge2.textContent = totalComPropostas;
            badge2.style.display = totalComPropostas > 0 ? 'flex' : 'none';
        }
        const totalMsg = document.getElementById('totalAlertasMsg');
        if (totalMsg) {
            const partes = [];
            if ((countSaldoZero || 0) > 0) partes.push(`${countSaldoZero} produto(s) em estoque baixo`);
            if (alertasPropostas.length > 0) partes.push(`${alertasPropostas.length} alerta(s) de proposta`);
            totalMsg.textContent = partes.join(' · ') || 'Nenhum alerta';
        }

        return totalComPropostas;
    } catch (e) {
        console.warn('verificarAlertasEstoque erro', e);
        return 0;
    }
}

// ========================================
// DASHBOARD COM GRÁFICOS (Chart.js)
// ========================================

let _chartVendasRep = null;
let _chartTopProdutos = null;
let _chartComissoesRep = null;

function renderizarGraficos() {
    if (typeof Chart === 'undefined') return;

    const reps = ['KOLTE', 'ISA', 'LC', 'ADES', 'FL', 'IMBEL'];
    const coresReps = ['#79c0ff', '#7ee787', '#58a6ff', '#ffa657', '#d2a8ff', '#ff7b72'];
    const vendasFiltradas = obterVendasDashboardFiltradas();
    const vendasPorRepMap = {};
    reps.forEach(rep => { vendasPorRepMap[rep] = 0; });

    const produtoTotals = new Map();
    vendasFiltradas.forEach(v => {
        const rep = (v.representante || '').toUpperCase();
        const itens = obterItensVendaNormalizados(v);
        itens.forEach(it => {
            if (vendasPorRepMap[rep] === undefined) vendasPorRepMap[rep] = 0;
            vendasPorRepMap[rep] += Number(it.quantidade) || 0;
            produtoTotals.set(it.produtoNome, (produtoTotals.get(it.produtoNome) || 0) + (Number(it.quantidade) || 0));
        });
    });
    const vendasPorRep = reps.map(rep => vendasPorRepMap[rep] || 0);

    // Chart 1: Vendas por Representante (bar)
    const ctx1 = document.getElementById('chartVendasRep');
    if (ctx1) {
        if (_chartVendasRep) _chartVendasRep.destroy();
        _chartVendasRep = new Chart(ctx1, {
            type: 'bar',
            data: {
                labels: reps,
                datasets: [{
                    label: 'Unidades Vendidas',
                    data: vendasPorRep,
                    backgroundColor: coresReps,
                    borderColor: coresReps,
                    borderWidth: 1,
                    borderRadius: 6
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: true,
                plugins: { legend: { display: false } },
                scales: { y: { beginAtZero: true, ticks: { precision: 0 } } }
            }
        });
    }

    // Chart 2: Top Produtos (doughnut)
    const dadosProdutos = Array.from(produtoTotals.entries())
        .map(([nome, total]) => ({ nome: (nome || '').substring(0, 20), total }))
        .filter(p => p.total > 0)
        .sort((a, b) => b.total - a.total)
        .slice(0, 6);

    const ctx2 = document.getElementById('chartTopProdutos');
    if (ctx2) {
        if (_chartTopProdutos) _chartTopProdutos.destroy();
        const palette = ['#79c0ff', '#7ee787', '#58a6ff', '#ffa657', '#d2a8ff', '#ff7b72'];
        _chartTopProdutos = new Chart(ctx2, {
            type: 'doughnut',
            data: {
                labels: dadosProdutos.map(p => p.nome),
                datasets: [{
                    data: dadosProdutos.map(p => p.total),
                    backgroundColor: palette.slice(0, dadosProdutos.length),
                    borderWidth: 2,
                    borderColor: '#fff'
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: true,
                plugins: {
                    legend: { position: 'bottom', labels: { font: { size: 11 }, padding: 10 } }
                }
            }
        });
    }
}

// ========================================
// MÓDULO DE CLIENTES
// ========================================

let clientes = [];

function abrirModalCliente(id = null) {
    document.getElementById('clienteEditId').value = '';
    document.getElementById('clienteNome').value = '';
    document.getElementById('clienteCnpj').value = '';
    document.getElementById('clienteEndereco').value = '';
    document.getElementById('clienteCidade').value = '';
    document.getElementById('clienteUf').value = '';
    document.getElementById('clienteTelefone').value = '';
    document.getElementById('clienteEmail').value = '';
    document.getElementById('clienteContato').value = '';
    try { popularSelectRepresentantes('clienteRepresentante', true); } catch (e) {}
    document.getElementById('clienteRepresentante').value = '';
    document.getElementById('clienteObservacoes').value = '';
    document.getElementById('modalClienteTitulo').textContent = 'Novo Cliente';
    try { if (typeof selecionarTipoPessoa === 'function') selecionarTipoPessoa('PJ'); } catch(e) {}

    if (id) {
        const cliente = clientes.find(c => c.id === id);
        if (cliente) {
            document.getElementById('clienteEditId').value = cliente.id;
            document.getElementById('clienteNome').value = cliente.nome || '';
            document.getElementById('clienteCnpj').value = cliente.cnpj || '';
            document.getElementById('clienteEndereco').value = cliente.endereco || '';
            document.getElementById('clienteCidade').value = cliente.cidade || '';
            document.getElementById('clienteUf').value = cliente.uf || '';
            document.getElementById('clienteTelefone').value = cliente.telefone || '';
            document.getElementById('clienteEmail').value = cliente.email || '';
            document.getElementById('clienteContato').value = cliente.contato || '';
            document.getElementById('clienteRepresentante').value = cliente.representante || '';
            document.getElementById('clienteObservacoes').value = cliente.observacoes || '';
            document.getElementById('modalClienteTitulo').textContent = 'Editar Cliente';
            // determinar tipo de pessoa a partir do registro salvo ou do formato do documento
            try {
                const tipoPessoa = cliente.tipoPessoa || (((cliente.cnpj||'').replace(/\D/g,'').length <= 11) ? 'PF' : 'PJ');
                if (typeof selecionarTipoPessoa === 'function') selecionarTipoPessoa(tipoPessoa);
            } catch(e) {}
        }
    }

    document.getElementById('modalCliente').style.display = 'block';
}

function fecharModalCliente() {
    fecharModal('modalCliente');
}

// Seleciona tipo de pessoa e ajusta rótulo/placeholder/styles
function selecionarTipoPessoa(tipo) {
    const hiddenInput = document.getElementById('clienteTipoPessoa');
    const label = document.getElementById('labelClienteDocumento');
    const input = document.getElementById('clienteCnpj');
    const btnPJ = document.getElementById('btnTipoPJ');
    const btnPF = document.getElementById('btnTipoPF');

    if (hiddenInput) hiddenInput.value = tipo;

    if (tipo === 'PJ') {
        if (label) label.textContent = 'CNPJ';
        if (input) {
            input.placeholder = '00.000.000/0001-00';
            input.maxLength = 18;
        }
        if (btnPJ) { btnPJ.style.background = '#1e3a5f'; btnPJ.style.color = '#fff'; }
        if (btnPF) { btnPF.style.background = '#f1f5f9'; btnPF.style.color = '#64748b'; }
    } else {
        if (label) label.textContent = 'CPF';
        if (input) {
            input.placeholder = '000.000.000-00';
            input.maxLength = 14;
        }
        if (btnPJ) { btnPJ.style.background = '#f1f5f9'; btnPJ.style.color = '#64748b'; }
        if (btnPF) { btnPF.style.background = '#1e3a5f'; btnPF.style.color = '#fff'; }
    }

    // Clear the field when switching type
    if (input) input.value = '';
}

function aplicarMascaraDocumento(input) {
    // Uses the existing formatCpfCnpjMask() function
    const tipo = document.getElementById('clienteTipoPessoa')?.value || 'PJ';
    const digits = (input.value || '').replace(/\D/g, '');

    if (tipo === 'PF') {
        // Force CPF mask (max 11 digits)
        const cpfDigits = digits.slice(0, 11);
        input.value = formatCpfCnpjMask(cpfDigits + 'xxx');
        input.value = formatCpfCnpjMask(cpfDigits);
    } else {
        // CNPJ mask (max 14 digits)
        input.value = formatCpfCnpjMask(digits.slice(0, 14));
    }
}

function salvarCliente(event) {
    event.preventDefault();

    const editId = document.getElementById('clienteEditId').value;
    const dados = {
        nome: document.getElementById('clienteNome').value.trim(),
        cnpj: document.getElementById('clienteCnpj').value.trim(),
        tipoPessoa: document.getElementById('clienteTipoPessoa')?.value || 'PJ',
        endereco: document.getElementById('clienteEndereco').value.trim(),
        cidade: document.getElementById('clienteCidade').value.trim(),
        uf: document.getElementById('clienteUf').value.trim().toUpperCase(),
        telefone: document.getElementById('clienteTelefone').value.trim(),
        email: document.getElementById('clienteEmail').value.trim(),
        contato: document.getElementById('clienteContato').value.trim(),
        representante: document.getElementById('clienteRepresentante').value,
        observacoes: document.getElementById('clienteObservacoes').value.trim()
    };

    if (editId) {
        const idx = clientes.findIndex(c => c.id === editId);
        if (idx !== -1) {
            clientes[idx] = { ...clientes[idx], ...dados };
            registrarHistorico('edição', `Cliente editado: ${dados.nome}`);
        }
    } else {
        const novo = {
            id: Date.now().toString(),
            ...dados,
            dataCadastro: new Date().toISOString()
        };
        clientes.push(novo);
        registrarHistorico('cadastro', `Cliente cadastrado: ${dados.nome}`);
    }

    renderizarClientes();
    fecharModal('modalCliente');
    atualizarKPIsClientes();
    atualizarDatalistClientes();
    estoque.clientes = clientes;
    salvarDados();
    salvarNoCloud().catch(e => console.error('Auto-save clientes falhou:', e));
    // Se a sub-aba de precificação por cliente estiver aberta, atualizar o dropdown
    try {
        const sub = document.getElementById('subaba-precif-porcliente');
        if (sub && sub.style.display === 'block') {
            popularSelectClientesPrecif();
        }
    } catch (e) {}
}

function excluirCliente(id) {
    const cliente = clientes.find(c => c.id === id);
    if (!cliente) return;
    if (!confirm(`Deseja excluir este cliente?\n${cliente.nome}`)) return;
    clientes = clientes.filter(c => c.id !== id);
    estoque.clientes = clientes;
    registrarHistorico('exclusão', `Cliente excluído: ${cliente.nome}`);
    renderizarClientes();
    atualizarKPIsClientes();
    atualizarDatalistClientes();
    salvarDados();
    salvarNoCloud().catch(e => console.error('Auto-save clientes falhou:', e));
}

function renderizarClientes(filtro = '') {
    const tbody = document.getElementById('tabelaClientesBody');
    if (!tbody) return;

    const termo = (filtro || '').toLowerCase();
    const clientesFiltrados = (clientes || []).filter(c =>
        !termo ||
        (c.nome || '').toLowerCase().includes(termo) ||
        (c.cnpj || '').includes(termo) ||
        (c.uf || '').toLowerCase().includes(termo) ||
        (c.contato || '').toLowerCase().includes(termo)
    );

    if (clientesFiltrados.length === 0) {
        tbody.innerHTML = '<tr><td colspan="10" style="text-align:center;padding:24px;color:var(--text-secondary)">Nenhum cliente encontrado.</td></tr>';
        return;
    }

    const vendas = estoque.registroVendas || [];

    tbody.innerHTML = clientesFiltrados.map(c => {
        const repClass = (c.representante || '').toLowerCase();
        const repBadge = c.representante
            ? `<span class="badge-rep ${repClass}">${c.representante}</span>`
            : '-';

        const tipo = c.tipoPessoa || (((c.cnpj||'').replace(/\D/g,'').length <= 11) ? 'PF' : 'PJ');
        const tipoBadge = `
            <span style="font-size:0.72rem;font-weight:700;padding:2px 7px;
                         border-radius:20px;
                         background:${tipo==='PJ'?"#eff6ff":"#faf5ff"};
                         color:${tipo==='PJ'?"#1d4ed8":"#7c3aed"}">
              ${tipo}
            </span>`;

        // Contar vendas que referenciam este cliente pelo nome
        const totalCompras = vendas.filter(v =>
            (v.loja || '').toLowerCase() === (c.nome || '').toLowerCase()
        ).length;

        const cidadeUf = [c.cidade, c.uf].filter(Boolean).join(' / ');

        return `<tr>
            <td>${c.nome || '-'}</td>
            <td style="text-align:center">${tipoBadge}</td>
            <td>${c.cnpj || '-'}</td>
            <td>${cidadeUf || '-'}</td>
            <td>${c.telefone || '-'}</td>
            <td>${c.email || '-'}</td>
            <td>${c.contato || '-'}</td>
            <td>${repBadge}</td>
            <td style="text-align:center">${totalCompras}</td>
            <td class="col-acoes">
                <button class="btn-action" onclick="abrirHistoricoCliente('${c.id}')" title="Ver histórico de compras"
                    style="background:#eff6ff;color:#1d4ed8;border:1px solid #bfdbfe;border-radius:6px;padding:3px 8px;cursor:pointer;font-size:0.8rem">
                    📋 Histórico
                </button>
                <button class="btn-action btn-edit" data-admin="true" onclick="abrirModalCliente('${c.id}')" title="Editar cliente">✎</button>
                <button class="btn-action btn-delete" data-admin="true" onclick="excluirCliente('${c.id}')" title="Excluir cliente">🗑</button>
            </td>
        </tr>`;
    }).join('');

    atualizarDatalistClientes();
}

function filtrarClientes(valor) {
    renderizarClientes(valor);
}

function atualizarKPIsClientes() {
    const totalEl = document.getElementById('kpiTotalClientes');
    const ativosEl = document.getElementById('kpiClientesAtivos');
    const ticketEl = document.getElementById('kpiTicketMedio');

    if (totalEl) totalEl.textContent = clientes.length;

    const vendas = estoque.registroVendas || [];

    // Clientes com compras e totais por cliente
    let clientesComCompras = 0;
    let somaTotal = 0;

    clientes.forEach(c => {
        const vendasCliente = vendas.filter(v =>
            (v.loja || '').toLowerCase() === (c.nome || '').toLowerCase()
        );
        if (vendasCliente.length > 0) {
            clientesComCompras++;
            vendasCliente.forEach(v => {
                if (Array.isArray(v.items)) {
                    v.items.forEach(it => { somaTotal += Number(it.valorTotal || 0); });
                } else {
                    somaTotal += Number(v.valorTotal || 0);
                }
            });
        }
    });

    if (ativosEl) ativosEl.textContent = clientesComCompras;
    if (ticketEl) {
        const ticketMedio = clientesComCompras > 0 ? somaTotal / clientesComCompras : 0;
        ticketEl.textContent = formatarMoedaValor(ticketMedio);
    }
}

function getClienteNomes() {
    return clientes.map(c => c.nome);
}

function atualizarDatalistClientes() {
    const datalist = document.getElementById('clientesDatalist');
    if (!datalist) return;
    datalist.innerHTML = clientes.map(c =>
        `<option value="${(c.nome || '').replace(/"/g, '&quot;')}">`
    ).join('');
}

function imprimirClientes() {
    const tabEl = document.getElementById('tab-clientes');
    if (!tabEl) return;
    const html = tabEl.querySelector('.table-container')?.innerHTML || '';
    const win = window.open('', '_blank');
    if (!win) { mostrarNotificacao('Pop-up bloqueado pelo navegador.', 'error'); return; }
    win.document.write(`<!DOCTYPE html><html><head><meta charset="UTF-8"><title>Clientes</title>
        <style>body{font-family:Inter,Arial,sans-serif;padding:20px}table{border-collapse:collapse;width:100%}th,td{border:1px solid #ddd;padding:6px 8px;font-size:12px}th{background:#f5f5f5;font-weight:600}.col-acoes{display:none}</style>
        <script>window.onload=function(){ setTimeout(function(){ window.print(); },200); }<\/script>
    </head><body><h2>Cadastro de Clientes</h2>${html}</body></html>`);
    win.document.close();
}

let _clienteHistoricoAtual = null;

function abrirHistoricoCliente(clienteId) {
    const cliente = (clientes||[]).find(c => String(c.id) === String(clienteId));
    if (!cliente) return;
    _clienteHistoricoAtual = cliente;

    const vendas = (estoque.registroVendas||[])
        .filter(v => (v.loja||'').toLowerCase() === (cliente.nome||'').toLowerCase())
        .sort((a,b) => new Date(b.data||0) - new Date(a.data||0));

    const propostasC = (propostas||[])
        .filter(p => p.cliente === cliente.nome)
        .sort((a,b) => new Date(b.dataCriacao||b.data||0) - new Date(a.dataCriacao||a.data||0));

    const precifs = (precificacoesCliente||[])
        .filter(p => String(p.clienteId) === String(clienteId))
        .sort((a,b) => new Date(b.dataCriacao||0) - new Date(a.dataCriacao||0));

    const fmt = v => 'R$ ' + (v||0).toLocaleString('pt-BR',{minimumFractionDigits:2});
    const totalFaturado = vendas.reduce((s,v) => s+(v.valorTotal||0), 0);
    const ticketMedio = vendas.length > 0 ? totalFaturado / vendas.length : 0;
    const taxaConversao = propostasC.length > 0
        ? Math.round(propostasC.filter(p=>p.status==='aceita').length / propostasC.length * 100)
        : 0;
    const ultimaCompra = vendas[0]?.data
        ? new Date(vendas[0].data).toLocaleDateString('pt-BR')
        : '—';

    // Header cards
    const headerEl = document.getElementById('historicoClienteHeader');
    if (headerEl) headerEl.innerHTML = `
        <div style="background:#eff6ff;border-radius:10px;padding:14px;
                                border:1px solid #bfdbfe;text-align:center">
            <div style="font-size:1.8rem;font-weight:800;color:#1d4ed8">
                ${vendas.length}
            </div>
            <div style="font-size:0.78rem;color:#64748b;text-transform:uppercase;
                                    letter-spacing:0.5px;margin-top:2px">Contratos</div>
        </div>
        <div style="background:#f0fdf4;border-radius:10px;padding:14px;
                                border:1px solid #86efac;text-align:center">
            <div style="font-size:1.4rem;font-weight:800;color:#16a34a">
                ${fmt(totalFaturado)}
            </div>
            <div style="font-size:0.78rem;color:#64748b;text-transform:uppercase;
                                    letter-spacing:0.5px;margin-top:2px">Total Faturado</div>
        </div>
        <div style="background:#fff7ed;border-radius:10px;padding:14px;
                                border:1px solid #fdba74;text-align:center">
            <div style="font-size:1.4rem;font-weight:800;color:#d97706">
                ${fmt(ticketMedio)}
            </div>
            <div style="font-size:0.78rem;color:#64748b;text-transform:uppercase;
                                    letter-spacing:0.5px;margin-top:2px">Ticket Médio</div>
        </div>
        <div style="background:#faf5ff;border-radius:10px;padding:14px;
                                border:1px solid #d8b4fe;text-align:center">
            <div style="font-size:1.8rem;font-weight:800;color:#7c3aed">
                ${taxaConversao}%
            </div>
            <div style="font-size:0.78rem;color:#64748b;text-transform:uppercase;
                                    letter-spacing:0.5px;margin-top:2px">Taxa Conversão</div>
        </div>
        <div style="background:#f1f5f9;border-radius:10px;padding:14px;
                                border:1px solid #cbd5e1;text-align:center">
            <div style="font-size:1rem;font-weight:700;color:#1e293b">
                ${ultimaCompra}
            </div>
            <div style="font-size:0.78rem;color:#64748b;text-transform:uppercase;
                                    letter-spacing:0.5px;margin-top:2px">Última Compra</div>
        </div>
    `;

    // Contratos tab
    const elContratos = document.getElementById('hcContent-contratos');
    if (elContratos) {
        elContratos.innerHTML = vendas.length === 0
            ? `<div style="text-align:center;padding:40px;color:#94a3b8">Nenhum contrato registrado para este cliente</div>`
            : `<table class="dashboard-table" style="font-size:0.85rem">
                    <thead>
                        <tr>
                            <th>Contrato</th>
                            <th>Data</th>
                            <th>Representante</th>
                            <th style="text-align:center">Itens</th>
                            <th style="text-align:right">Valor Total</th>
                            <th>Observações</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${vendas.map(v => `
                            <tr>
                                <td style="font-weight:700;color:#1e3a5f">#${v.contrato}</td>
                                <td>${v.data ? new Date(v.data).toLocaleDateString('pt-BR') : '—'}</td>
                                <td>${v.representante || '—'}</td>
                                <td style="text-align:center">${(v.items||v.itens||[]).length}</td>
                                <td style="text-align:right;font-weight:700;color:#16a34a">${fmt(v.valorTotal)}</td>
                                <td style="font-size:0.78rem;color:#64748b;max-width:200px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${v.observacoes||''}">${v.observacoes || '—'}</td>
                            </tr>`).join('')}
                    </tbody>
                    <tfoot>
                        <tr style="background:#f8fafc;font-weight:700">
                            <td colspan="4" style="padding:10px 12px;color:#64748b">Total (${vendas.length} contrato(s))</td>
                            <td style="text-align:right;color:#16a34a;font-size:1rem">${fmt(totalFaturado)}</td>
                            <td></td>
                        </tr>
                    </tfoot>
                </table>`;
    }

    // Propostas tab
    const elPropostas = document.getElementById('hcContent-propostas');
    const statusColor = { rascunho:'#64748b', enviada:'#1d4ed8', aceita:'#16a34a', recusada:'#dc2626', expirada:'#d97706', convertida:'#0ea5e9' };
    if (elPropostas) {
        elPropostas.innerHTML = propostasC.length === 0
            ? `<div style="text-align:center;padding:40px;color:#94a3b8">Nenhuma proposta para este cliente</div>`
            : `<table class="dashboard-table" style="font-size:0.85rem">
                    <thead>
                        <tr>
                            <th>Proposta</th>
                            <th>Data</th>
                            <th>Status</th>
                            <th style="text-align:center">Itens</th>
                            <th style="text-align:right">Valor</th>
                            <th>Contrato</th>
                            <th>Motivo Recusa</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${propostasC.map(p => {
                            const sc = statusColor[p.status]||'#64748b';
                            return `<tr>
                                <td style="font-weight:700;color:#1e3a5f">${_escapeHtml(p.numero)}</td>
                                <td>${new Date(p.dataCriacao||p.data||new Date()).toLocaleDateString('pt-BR')}</td>
                                <td><span style="background:${sc}20;color:${sc};font-size:0.72rem;font-weight:600;padding:2px 8px;border-radius:20px">${p.status}</span></td>
                                <td style="text-align:center">${(p.itens||[]).length}</td>
                                <td style="text-align:right;font-weight:600;color:#16a34a">${fmt(p.valorTotal)}</td>
                                <td style="font-weight:600;color:#1e3a5f">${p.contratoNumero ? '#'+p.contratoNumero : '—'}</td>
                                <td style="font-size:0.78rem;color:#dc2626;max-width:180px;overflow:hidden;text-overflow:ellipsis" title="${p.motivoRecusa||''}">${p.motivoRecusa || '—'}</td>
                            </tr>`;
                        }).join('')}
                    </tbody>
                </table>`;
    }

    // Precificações tab
    const elPrec = document.getElementById('hcContent-precificacoes');
    if (elPrec) {
        elPrec.innerHTML = precifs.length === 0
            ? `<div style="text-align:center;padding:40px;color:#94a3b8">Nenhuma precificação salva para este cliente</div>`
            : `<table class="dashboard-table" style="font-size:0.85rem">
                    <thead>
                        <tr>
                            <th>Versão</th>
                            <th>Data</th>
                            <th>Descrição</th>
                            <th>Taxa</th>
                            <th>ROI</th>
                            <th style="text-align:center">Produtos</th>
                            <th>Status</th>
                            <th>Proposta</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${precifs.map(p => `
                            <tr>
                                <td style="font-weight:700;color:#1e3a5f">v${p.versao||1}</td>
                                <td>${new Date(p.dataCriacao||new Date()).toLocaleDateString('pt-BR')}</td>
                                <td style="color:#475569">${_escapeHtml(p.descricao||'—')}</td>
                                <td style="text-align:center">${p.taxa||0}%</td>
                                <td style="text-align:center">${p.roi||0}%</td>
                                <td style="text-align:center">${(p.itens||[]).length}</td>
                                <td><span style="font-size:0.72rem;font-weight:600;padding:2px 8px;border-radius:20px;background:${p.status==='ativa'?'#f0fdf4':p.status==='convertida'?'#eff6ff':'#f1f5f9'};color:${p.status==='ativa'?'#16a34a':p.status==='convertida'?'#1d4ed8':'#64748b'}">${p.status}</span></td>
                                <td style="color:#0ea5e9;font-weight:600">${p.propostaId ? ((propostas||[]).find(x=>x.id===p.propostaId)?.numero||'—') : '—'}</td>
                            </tr>`).join('')}
                    </tbody>
                </table>`;
    }

    // Show first tab
    trocarAbaHistoricoCliente('contratos');
    const modal = document.getElementById('modalHistoricoCliente');
    if (modal) modal.style.display = 'flex';
}

function trocarAbaHistoricoCliente(aba) {
    ['contratos','propostas','precificacoes'].forEach(t => {
        const content = document.getElementById('hcContent-' + t);
        const btn = document.getElementById('hcTab-' + t);
        if (content) content.style.display = t === aba ? 'block' : 'none';
        if (btn) {
            btn.style.color = t === aba ? '#1e3a5f' : '#64748b';
            btn.style.borderBottomColor = t === aba ? '#1e3a5f' : 'transparent';
        }
    });
}

// ========================================
// MÓDULO DE PRECIFICAÇÃO
// ========================================

const ESTADOS_BR = [
    'AC', 'AL', 'AM', 'AP', 'BA', 'CE', 'DF', 'ES', 'GO', 'MA', 'MG', 'MS', 'MT',
    'PA', 'PB', 'PE', 'PI', 'PR', 'RJ', 'RN', 'RO', 'RR', 'RS', 'SC', 'SE', 'SP', 'TO'
];

function _nomeProdutoId(nome) {
    return String(nome || '').replace(/[^a-zA-Z0-9]/g, '_');
}

function _escapeHtml(texto) {
    return String(texto || '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}

function _escapeJsString(texto) {
    return String(texto || '').replace(/\\/g, '\\\\').replace(/'/g, "\\'");
}

function _fmtMoeda(v) {
    return Number(v || 0).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function detectarNCM(nomeProduto) {
    if (!nomeProduto) return null;
    const nomeUpper = nomeProduto.toString().toUpperCase();
    for (const [keyword, ncm] of Object.entries(NCM_POR_CATEGORIA)) {
        if (nomeUpper.includes(keyword)) return ncm;
    }
    return null;
}

function inicializarImpostosPreDefinidos() {
    try {
        // 1. Detectar NCM por nome e preencher tabelaAliquotas com impostos federais quando ausentes
        (estoque.produtos || []).forEach(produto => {
            const nomeUpper = (produto.nome || '').toString().toUpperCase();
            let ncmDetectado = null;
            for (const [keyword, ncm] of Object.entries(NCM_POR_CATEGORIA)) {
                if (nomeUpper.includes(keyword)) { ncmDetectado = ncm; break; }
            }
            if (!ncmDetectado) return;
            produto.ncm = produto.ncm || ncmDetectado;

            if (!tabelaAliquotas[produto.nome]) tabelaAliquotas[produto.nome] = {};
            const aliq = tabelaAliquotas[produto.nome];
            const fed = IMPOSTOS_FEDERAIS_POR_NCM[ncmDetectado];
            if (fed) {
                if (!aliq.pis) aliq.pis = fed.pis;
                if (!aliq.cofins) aliq.cofins = fed.cofins;
                if (!aliq.ipi) aliq.ipi = fed.ipi;
            }
        });

        // 2. Preencher tabelaICMS com regras predefinidas por NCM/UF/PF-PJ
        const regraExiste = (ncm, estado, tipoPessoa) =>
            tabelaICMS.some(r => String(r.ncm || '').toUpperCase() === String(ncm || '').toUpperCase()
                && String(r.estado).toUpperCase() === String(estado).toUpperCase()
                && String(r.tipoPessoa).toUpperCase() === String(tipoPessoa).toUpperCase()
            );

        const ESTADOS = [
            'AC','AL','AP','AM','BA','CE','DF','ES','GO',
            'MA','MT','MS','PA','PB','PR','PE','PI',
            'RJ','RN','RS','RO','RR','SC','SP','SE','TO','MG'
        ];

        Object.entries(ICMS_PJ_POR_NCM).forEach(([ncm, estadoMap]) => {
            ESTADOS.forEach(uf => {
                if (estadoMap[uf] !== undefined && !regraExiste(ncm, uf, 'PJ')) {
                    tabelaICMS.push({
                        id: `predef_${ncm}_${uf}_PJ`,
                        ncm,
                        estado: uf,
                        tipoPessoa: 'PJ',
                        categoriaProduto: 'Todos',
                        aliquota: estadoMap[uf]
                    });
                }
            });
        });

        Object.entries(ICMS_PF_POR_NCM).forEach(([ncm, estadoMap]) => {
            ESTADOS.forEach(uf => {
                if (estadoMap[uf] !== undefined && !regraExiste(ncm, uf, 'PF')) {
                    tabelaICMS.push({
                        id: `predef_${ncm}_${uf}_PF`,
                        ncm,
                        estado: uf,
                        tipoPessoa: 'PF',
                        categoriaProduto: 'Todos',
                        aliquota: estadoMap[uf]
                    });
                }
            });
        });

        // Persistir mudanças mínimas localmente
        salvarDados();
    } catch (e) {
        console.warn('inicializarImpostosPreDefinidos falhou:', e);
    }
}

function buscarAliquotaICMS(estado, tipoPessoa, nomeProduto) {
    // tentar obter produto e seu NCM
    const produto = (estoque.produtos || []).find(p => p.nome === nomeProduto);
    const ncm = produto?.ncm || detectarNCM(nomeProduto);

    // 1) tentar tabela predefinida por NCM (PJ/PF)
    if (ncm) {
        const tabela = (tipoPessoa === 'PF') ? ICMS_PF_POR_NCM : ICMS_PJ_POR_NCM;
        if (tabela[ncm] && tabela[ncm][estado] !== undefined) {
            return tabela[ncm][estado];
        }
    }

    // 2) regras usuário (tabelaICMS) com scoring incluindo NCM
    const categoria = categoriaPorProduto[nomeProduto] || 'Outro';
    const score = (rule) => {
        let s = 0;
        if (rule.ncm && ncm && rule.ncm === ncm) s += 8;
        else if (rule.ncm && rule.ncm !== 'Todos') return -1;
        if (rule.estado === estado) s += 4;
        else if (rule.estado !== 'Todos') return -1;
        if (rule.tipoPessoa === tipoPessoa) s += 2;
        else if (rule.tipoPessoa !== 'Todos') return -1;
        if (rule.categoriaProduto === categoria) s += 1;
        else if (rule.categoriaProduto !== 'Todos') return -1;
        return s;
    };

    const match = tabelaICMS
        .map(r => ({ rule: r, score: score(r) }))
        .filter(x => x.score >= 0)
        .sort((a, b) => b.score - a.score)[0];

    if (match) return match.rule.aliquota;

    // fallback: icmsBase na tabelaAliquotas do produto
    return parseFloat(tabelaAliquotas[nomeProduto]?.icmsBase) || 0;
}

function calcularPreco(nomeProduto, estado = null, tipoPessoa = null) {
    const prec = precificacao[nomeProduto] || {};
    const aliq = tabelaAliquotas[nomeProduto] || {};

    const ci = parseFloat(prec.ci) || 0;
    // `taxa` and `roi` are provided as percentages (e.g. 1 = 1%).
    // Read stored value or fallback to defaults, then convert to multipliers below.
    let taxaPct = (prec.taxa !== null && prec.taxa !== undefined && prec.taxa !== '')
        ? parseFloat(prec.taxa)
        : parseFloat(document.getElementById('taxaPadrao')?.value);
    if (!Number.isFinite(taxaPct)) taxaPct = 1;
    let roiPct = (prec.roi !== null && prec.roi !== undefined && prec.roi !== '')
        ? parseFloat(prec.roi)
        : parseFloat(document.getElementById('roiPadrao')?.value);
    if (!Number.isFinite(roiPct)) roiPct = 1;
    const comissao = parseFloat(prec.comissao)
        || parseFloat(document.getElementById('comissaoPadrao')?.value)
        || 5;
    const pis = parseFloat(aliq.pis)
        || parseFloat(document.getElementById('pisPadrao')?.value)
        || 0;
    const cofins = parseFloat(aliq.cofins)
        || parseFloat(document.getElementById('cofinsPadrao')?.value)
        || 0;
    const ipi = parseFloat(aliq.ipi) || 0;

    const icms = (estado && tipoPessoa)
        ? buscarAliquotaICMS(estado, tipoPessoa, nomeProduto)
        : (parseFloat(aliq.icmsBase) || 0);

    if (ci === 0) return null;

    // Step 1
    // Convert percent values to multiplicative factors and apply sequentially
    const taxaFactor = 1 + (Number(taxaPct) / 100);
    const roiFactor = 1 + (Number(roiPct) / 100);
    const valorBase = ci * taxaFactor * roiFactor;

    // Step 2
    const icmsR = valorBase * icms / 100;
    const pisR = valorBase * pis / 100;
    const cofinsR = valorBase * cofins / 100;
    const valorImpostos = valorBase + icmsR + pisR + cofinsR;

    // Step 3
    const ipiR = valorImpostos * ipi / 100;
    const valorTotal = valorImpostos + ipiR;

    // Step 4
    const comissaoR = valorBase * comissao / 100;
    const precoFinal = (prec.precoFinalManual !== null && prec.precoFinalManual !== undefined && prec.precoFinalManual !== '')
        ? parseFloat(prec.precoFinalManual)
        : valorTotal + comissaoR;

    return {
        ci, taxa: taxaPct, roi: roiPct,
        valorBase,
        icms, icmsR,
        pis, pisR,
        cofins, cofinsR,
        valorImpostos,
        ipi, ipiR,
        valorTotal,
        comissao, comissaoR,
        precoFinal,
        isManual: !!(prec.precoFinalManual !== null && prec.precoFinalManual !== undefined && prec.precoFinalManual !== '')
    };
}

function renderizarPrecificacao() {
    try {
        const tbody = document.getElementById('tabelaPrecificacaoBody');
        const resumo = document.getElementById('precificacaoResumoCI');
        if (!tbody) return;

        const produtos = estoque?.produtos || [];
        if (!produtos.length) {
            tbody.innerHTML = '<tr><td colspan="6" style="text-align:center;color:#94a3b8;padding:20px">Nenhum produto cadastrado</td></tr>';
            if (resumo) resumo.innerHTML = '';
            return;
        }

        tbody.innerHTML = produtos.map((produto) => {
            const prec = precificacao[produto.nome] || {};
            const ncm = produto.ncm || detectarNCM(produto.nome) || '—';
            const nomeId = produto.nome.replace(/[^a-zA-Z0-9]/g, '_');
            const nomeJs = _escapeJsString(produto.nome);
            const categoriaAtual = categoriaPorProduto[produto.nome] || '';
            const ci = prec.ci || '';
            const ultimaAtt = prec.ciAtualizadoEm
                ? new Date(prec.ciAtualizadoEm).toLocaleDateString('pt-BR')
                : '—';
            const temCI = parseFloat(ci) > 0;

            return `
                <tr>
                    <td style="text-align:left; padding-left:15px; font-weight:500; position:sticky; left:0; background:#fff; z-index:1; border-right:1px solid #e2e8f0">
                        ${_escapeHtml(produto.nome)}
                        <div style="font-size:0.72rem; font-family:monospace; color:#94a3b8; margin-top:2px">${_escapeHtml(ncm)}</div>
                    </td>
                    <td style="text-align:center">
                        <span style="font-size:0.75rem; font-family:monospace; color:#64748b; background:#f1f5f9; padding:2px 8px; border-radius:4px">${_escapeHtml(ncm)}</span>
                    </td>
                    <td style="text-align:center">
                        <select onchange="salvarCategoriaProduto('${nomeJs}', this.value)" style="border:1px solid #e2e8f0; border-radius:6px; padding:5px 8px; font-size:0.82rem; background:#fff; min-width:110px">
                            <option value="">Selecione...</option>
                            ${CATEGORIAS_PRODUTO.map(c => `<option value="${c}" ${categoriaAtual === c ? 'selected' : ''}>${c}</option>`).join('')}
                        </select>
                    </td>
                    <td style="text-align:center">
                        <div style="display:flex; align-items:center; justify-content:center; gap:6px">
                            <span style="color:#64748b; font-size:0.9rem">R$</span>
                            <input type="number" step="0.01" min="0" id="ci_${nomeId}" value="${ci}" placeholder="0,00" onchange="(typeof salvarCI === 'function') && salvarCI('${nomeJs}', this.value)" style="width:120px; border:1px solid ${temCI ? '#22c55e' : '#e2e8f0'}; border-radius:6px; padding:6px 8px; text-align:right; font-size:0.9rem; font-weight:${temCI ? '600' : '400'}; color:${temCI ? '#16a34a' : '#1e293b'}">
                        </div>
                    </td>
                    <td style="text-align:center">
                        <input type="number" step="0.1" min="0" max="100"
                               value="${precificacao[produto.nome]?.margemMinima ?? ''}"
                               placeholder="—"
                               onchange="salvarMargemMinima('${nomeJs}', this.value)"
                               style="width:75px;border:1px solid #e2e8f0;border-radius:6px;
                                      padding:5px 8px;text-align:center;font-size:0.85rem">
                    </td>
                    <td id="ci_att_${nomeId}" style="text-align:center; font-size:0.8rem; color:#94a3b8">${ultimaAtt}</td>
                </tr>
            `;
        }).join('');

        atualizarResumoPrecificacaoCI(produtos);
    } catch (e) {
        console.error('renderizarPrecificacao error:', e);
    }
}

function atualizarResumoPrecificacaoCI(produtos = estoque.produtos || []) {
    const resumo = document.getElementById('precificacaoResumoCI');
    if (!resumo) return;

    const total = produtos.length;
    const comCI = produtos.filter(produto => (parseFloat(precificacao[produto.nome]?.ci) || 0) > 0).length;
    const semCI = total - comCI;

    resumo.innerHTML = `
        <div style="font-size:0.85rem; color:#475569">Total produtos: <strong style="color:#1e3a5f">${total}</strong></div>
        <div style="font-size:0.85rem; color:#475569">Com CI: <strong style="color:#16a34a">${comCI}</strong></div>
        <div style="font-size:0.85rem; color:#475569">Sem CI: <strong style="color:#d97706">${semCI}</strong></div>
    `;
}

// Função de salvamento de CI removida: gestão de CI passa a ser feita pelo Cadastro de Produtos (modal `salvarProduto`).

function salvarMargemMinima(nomeProduto, valor) {
    if (!precificacao[nomeProduto]) precificacao[nomeProduto] = {};
    precificacao[nomeProduto].margemMinima = parseFloat(valor) || null;
    estoque.precificacao = precificacao;
    try { localStorage.setItem('estoqueArmasV2', JSON.stringify(estoque)); } catch (e) {}
}

function atualizarLinhaPrecificacao(nomeProduto) {
    const nomeId = _nomeProdutoId(nomeProduto);
    const ciEl = document.getElementById(`ci_${nomeId}`);
    const taxaEl = document.getElementById(`taxa_${nomeId}`);
    const roiEl = document.getElementById(`roi_${nomeId}`);
    const pisEl = document.getElementById(`pis_${nomeId}`);
    const cofinsEl = document.getElementById(`cofins_${nomeId}`);
    const icmsEl = document.getElementById(`icms_${nomeId}`);
    const ipiEl = document.getElementById(`ipi_${nomeId}`);
    const comissaoEl = document.getElementById(`comissao_${nomeId}`);

    if (!precificacao[nomeProduto]) precificacao[nomeProduto] = { ci: 0, taxa: null, roi: null, comissao: null, precoFinalManual: null };
    if (!tabelaAliquotas[nomeProduto]) tabelaAliquotas[nomeProduto] = { pis: null, cofins: null, ipi: null, icmsBase: null };

    const ciRaw = ciEl?.value ?? '';
    const taxaRaw = taxaEl?.value ?? '';
    const roiRaw = roiEl?.value ?? '';
    const pisRaw = pisEl?.value ?? '';
    const cofinsRaw = cofinsEl?.value ?? '';
    const icmsRaw = icmsEl?.value ?? '';
    const ipiRaw = ipiEl?.value ?? '';
    const comissaoRaw = comissaoEl?.value ?? '';

    const taxaDefault = parseFloat(document.getElementById('taxaPadrao')?.value) || 1;
    const roiDefault = parseFloat(document.getElementById('roiPadrao')?.value) || 1;
    const pisDefault = parseFloat(document.getElementById('pisPadrao')?.value) || 0;
    const cofinsDefault = parseFloat(document.getElementById('cofinsPadrao')?.value) || 0;
    const comissaoDefault = parseFloat(document.getElementById('comissaoPadrao')?.value) || 5;

    precificacao[nomeProduto].ci = parseFloat(ciRaw) || 0;
    precificacao[nomeProduto].taxa = taxaRaw === '' ? null : (parseFloat(taxaRaw) || taxaDefault);
    precificacao[nomeProduto].roi = roiRaw === '' ? null : (parseFloat(roiRaw) || roiDefault);
    precificacao[nomeProduto].comissao = comissaoRaw === '' ? null : (parseFloat(comissaoRaw) || comissaoDefault);

    tabelaAliquotas[nomeProduto].pis = pisRaw === '' ? null : (parseFloat(pisRaw) || pisDefault);
    tabelaAliquotas[nomeProduto].cofins = cofinsRaw === '' ? null : (parseFloat(cofinsRaw) || cofinsDefault);
    tabelaAliquotas[nomeProduto].icmsBase = icmsRaw === '' ? null : (parseFloat(icmsRaw) || 0);
    tabelaAliquotas[nomeProduto].ipi = ipiRaw === '' ? null : (parseFloat(ipiRaw) || 0);

    const r = calcularPreco(nomeProduto);

    const vbaseEl = document.getElementById(`vbase_${nomeId}`);
    const vimpEl = document.getElementById(`vimp_${nomeId}`);
    const vcomEl = document.getElementById(`vcom_${nomeId}`);
    const vpfEl = document.getElementById(`vpf_${nomeId}`);

    if (!r) {
        if (vbaseEl) vbaseEl.textContent = '-';
        if (vimpEl) vimpEl.textContent = '-';
        if (vcomEl) vcomEl.textContent = '-';
        if (vpfEl && !(precificacao[nomeProduto].precoFinalManual !== null && precificacao[nomeProduto].precoFinalManual !== undefined && precificacao[nomeProduto].precoFinalManual !== '')) {
            vpfEl.value = '';
        }
        if (vpfEl) vpfEl.style.border = '2px solid #22c55e';
        salvarDados();
        return;
    }

    if (vbaseEl) vbaseEl.textContent = _fmtMoeda(r.valorBase);
    if (vimpEl) vimpEl.textContent = _fmtMoeda(r.valorImpostos);
    if (vcomEl) vcomEl.textContent = _fmtMoeda(r.comissaoR);

    const isManual = !!(precificacao[nomeProduto].precoFinalManual !== null && precificacao[nomeProduto].precoFinalManual !== undefined && precificacao[nomeProduto].precoFinalManual !== '');
    if (vpfEl && !isManual) {
        vpfEl.value = Number(r.precoFinal).toFixed(2);
    }
    if (vpfEl) {
        vpfEl.style.border = isManual ? '2px solid #f59e0b' : '2px solid #22c55e';
    }

    salvarDados();
}

function salvarCategoriaProduto(nomeProduto, categoria) {
    categoriaPorProduto[nomeProduto] = categoria;
    salvarDados();
}

function salvarPrecoFinalManual(nomeProduto, valor) {
    if (!precificacao[nomeProduto]) {
        precificacao[nomeProduto] = { ci: 0, taxa: null, roi: null, comissao: null, precoFinalManual: null };
    }

    const manual = parseFloat(valor);
    if (!Number.isFinite(manual)) {
        precificacao[nomeProduto].precoFinalManual = null;
        atualizarLinhaPrecificacao(nomeProduto);
        return;
    }

    const backup = precificacao[nomeProduto].precoFinalManual;
    precificacao[nomeProduto].precoFinalManual = null;
    const calculado = calcularPreco(nomeProduto);
    const precoAuto = calculado ? Number(calculado.precoFinal) : 0;
    const vpf = document.getElementById(`vpf_${_nomeProdutoId(nomeProduto)}`);

    if (Math.abs(manual - precoAuto) > 0.01) {
        precificacao[nomeProduto].precoFinalManual = manual;
        if (vpf) vpf.style.border = '2px solid #f59e0b';
    } else {
        precificacao[nomeProduto].precoFinalManual = null;
        if (vpf) vpf.style.border = '2px solid #22c55e';
    }

    if (backup !== precificacao[nomeProduto].precoFinalManual) {
        salvarDados();
    }
}

function resetarPrecoManual(nomeProduto) {
    if (!precificacao[nomeProduto]) {
        precificacao[nomeProduto] = { ci: 0, taxa: null, roi: null, comissao: null, precoFinalManual: null };
    }
    precificacao[nomeProduto].precoFinalManual = null;
    atualizarLinhaPrecificacao(nomeProduto);
    const vpf = document.getElementById(`vpf_${_nomeProdutoId(nomeProduto)}`);
    if (vpf) vpf.style.border = '2px solid #22c55e';
    salvarDados();
}

// ── SUB-TAB NAVIGATION FOR PRECIFICAÇÃO ─────────────────────────
function trocarSubabaPrecif(subaba) {
    ['produtos','federais','icms','porcliente','comparativo','consulta','rastreabilidade'].forEach(s => {
                const el = document.getElementById('subaba-precif-' + s);
                if (el) el.style.display = (s === subaba) ? 'block' : 'none';
                const btn = document.getElementById('sbtn-' + s);
                if (btn) {
                        btn.style.color = (s === subaba) ? '#1e3a5f' : '#64748b';
                        btn.style.borderBottomColor = (s === subaba) ? '#1e3a5f' : 'transparent';
                }
        });
        if (subaba === 'federais') renderizarImpostosFederais();
        if (subaba === 'icms') renderizarICMSPorEstado();
        if (subaba === 'comparativo') {
            try { popularSelectsComparativo(); } catch (e) {}
            try { calcularComparativo(); } catch (e) {}
        }
        if (subaba === 'consulta') {
            try { renderizarConsultaPrecificacao(); } catch (e) {}
        }
        if (subaba === 'rastreabilidade') {
            try { renderizarRastreabilidade(); } catch (e) {}
        }
        if (subaba === 'porcliente') {
                // preparar dropdown e estado inicial da sub-aba
                try {
                try { verificarExpiracaoPrecificacoes(); } catch (e) {}
                popularSelectClientesPrecif();
                } catch (e) {}
                const res = document.getElementById('precifClienteResultado');
                const empty = document.getElementById('precifClienteEmpty');
                const banner = document.getElementById('precifClienteBanner');
                if (res) res.style.display = 'none';
                if (empty) empty.style.display = 'block';
                if (banner) banner.style.display = 'none';
        }
}

function popularSelectsComparativo() {
        const lista = (clientes || []).slice().sort((a, b) => (a.nome || '').localeCompare(b.nome || ''));
        [1, 2, 3].forEach(n => {
                const sel = document.getElementById('compCliente' + n);
                if (!sel) return;
                const valorAtual = sel.value;
                sel.innerHTML = '<option value="">— não selecionado —</option>'
                        + lista.map(c => {
                                const cnpjLimpo = (c.cnpj || '').replace(/\D/g, '');
                                const tipo = cnpjLimpo.length === 14 ? 'PJ' : 'PF';
                                return `<option value="${c.id}">${_escapeHtml(c.nome || '')} (${_escapeHtml(c.uf || '??')} / ${tipo})</option>`;
                        }).join('');
                if (valorAtual) sel.value = valorAtual;
        });
}

function renderizarConsultaPrecificacao() {
    // Populate selects
    const selCliente = document.getElementById('consultaFiltroCliente');
    if (selCliente && selCliente.options.length <= 1) {
        (clientes||[]).slice().sort((a,b)=> (a.nome||'').localeCompare(b.nome||'')).forEach(c => {
            const opt = document.createElement('option');
            opt.value = c.id;
            opt.textContent = (c.nome || '') + ' (' + (c.uf||'??') + ')';
            selCliente.appendChild(opt);
        });
    }

    const selProd = document.getElementById('consultaFiltroProduto');
    if (selProd && selProd.options.length <= 1) {
        (estoque.produtos||[]).slice().sort((a,b)=> (a.nome||'').localeCompare(b.nome||'')).forEach(p => {
            const opt = document.createElement('option');
            opt.value = p.nome;
            opt.textContent = p.nome;
            selProd.appendChild(opt);
        });
    }

    const filtroClienteId = document.getElementById('consultaFiltroCliente')?.value || '';
    const filtroStatus    = document.getElementById('consultaFiltroStatus')?.value || '';
    const filtroProduto   = document.getElementById('consultaFiltroProduto')?.value || '';

    const fmt = v => 'R$ ' + parseFloat(v||0).toLocaleString('pt-BR',{minimumFractionDigits:2,maximumFractionDigits:2});
    const pct = v => parseFloat(v||0).toFixed(2) + '%';
    const corMargem = m => m >= 30 ? '#16a34a' : m >= 15 ? '#d97706' : '#dc2626';

    const statusColors = {
        ativa:      { bg:'#f0fdf4', text:'#16a34a' },
        expirada:   { bg:'#fff8f0', text:'#d97706' },
        convertida: { bg:'#eff6ff', text:'#1d4ed8' },
        arquivada:  { bg:'#f1f5f9', text:'#94a3b8' },
    };

    const rows = [];

    (precificacoesCliente||[])
        .filter(p => {
            if (filtroClienteId && String(p.clienteId) !== String(filtroClienteId)) return false;
            if (filtroStatus && p.status !== filtroStatus) return false;
            return true;
        })
        .sort((a,b) => new Date(b.dataCriacao||0) - new Date(a.dataCriacao||0))
        .forEach(precif => {
            const cliente = (clientes||[]).find(c => String(c.id) === String(precif.clienteId));
            const cnpjLimpo = (cliente?.cnpj||'').replace(/\D/g,'');
            const tipo = cnpjLimpo.length === 14 ? 'PJ' : 'PF';
            const sc = statusColors[precif.status] || { bg:'#f1f5f9', text:'#64748b' };
            const proposta = precif.propostaId ? (propostas||[]).find(p => p.id === precif.propostaId) : null;

            const itensFiltrados = (precif.itens||[]).filter(it => !filtroProduto || it.produto === filtroProduto);
            if (!itensFiltrados.length && filtroProduto) return;

            const itensParaExibir = itensFiltrados.length ? itensFiltrados : (precif.itens||[]);

            itensParaExibir.forEach((item, idx) => {
                const margem = item.margem ?? (
                    item.precoFinal > 0 ? ((item.precoFinal - item.ci) / item.precoFinal * 100) : 0
                );

                rows.push({
                    clienteNome: precif.clienteNome || cliente?.nome || '—',
                    uf: precif.clienteUF || cliente?.uf || '—',
                    tipo,
                    versao: 'v' + (precif.versao || 1),
                    data: new Date(precif.dataCriacao||new Date()).toLocaleDateString('pt-BR'),
                    status: precif.status,
                    sc,
                    produto: item.produto || '—',
                    ncm: item.ncm || '—',
                    ci: item.ci || 0,
                    taxa: item.taxa || 0,
                    roi: item.roi || 0,
                    valorBase: item.valorBase || 0,
                    icms: item.icms || 0, icmsR: item.icmsR || 0,
                    pis: item.pis || 0, pisR: item.pisR || 0,
                    cofins: item.cofins || 0, cofinsR: item.cofinsR || 0,
                    ipi: item.ipi || 0, ipiR: item.ipiR || 0,
                    valorImpostos: item.valorImpostos || 0,
                    comissao: item.comissao || 0, comissaoR: item.comissaoR || 0,
                    precoFinal: item.precoFinal || 0,
                    margem,
                    propostaNum: proposta?.numero || (precif.propostaId ? '—' : ''),
                });
            });
        });

    const tbody = document.getElementById('tabelaConsultaPrecifBody');
    if (!tbody) return;

    if (!rows.length) {
        tbody.innerHTML = `
            <tr>
                <td colspan="26" style="text-align:center;color:#94a3b8;padding:60px;font-size:0.9rem">
                    <div style="font-size:2rem;margin-bottom:8px">🔍</div>
                    Nenhuma precificação encontrada com os filtros selecionados
                </td>
            </tr>`;
        return;
    }

    tbody.innerHTML = rows.map((r, i) => `
        <tr style="background:${i%2===0?'#fff':'#f8fafc'};border-bottom:1px solid #f1f5f9">
            <td style="font-weight:600;color:#1e3a5f;padding:8px 12px;position:sticky;left:0;background:${i%2===0?'#fff':'#f8fafc'};z-index:1;border-right:1px solid #e2e8f0">${r.clienteNome}</td>
            <td style="text-align:center;font-size:0.8rem">${r.uf}</td>
            <td style="text-align:center;font-size:0.75rem;font-weight:600;color:${r.tipo==='PJ'?'#1d4ed8':'#7c3aed'}">${r.tipo}</td>
            <td style="text-align:center;font-weight:700;color:#1e3a5f">${r.versao}</td>
            <td style="font-size:0.78rem;color:#64748b">${r.data}</td>
            <td style="text-align:center"><span style="background:${r.sc.bg};color:${r.sc.text};font-size:0.72rem;font-weight:600;padding:2px 8px;border-radius:20px">${r.status}</span></td>
            <td style="font-weight:500;max-width:180px;overflow:hidden;text-overflow:ellipsis" title="${r.produto}">${r.produto}</td>
            <td style="font-size:0.75rem;color:#64748b">${r.ncm}</td>
            <td style="text-align:right;font-weight:600;color:#1e3a5f">${fmt(r.ci)}</td>
            <td style="text-align:center">${pct(r.taxa)}</td>
            <td style="text-align:center">${pct(r.roi)}</td>
            <td style="text-align:right;color:#475569">${fmt(r.valorBase)}</td>
            <td style="text-align:center;color:#0369a1;font-weight:600">${pct(r.icms)}</td>
            <td style="text-align:right;color:#0369a1">${fmt(r.icmsR)}</td>
            <td style="text-align:center;color:#dc2626;font-weight:600">${pct(r.pis)}</td>
            <td style="text-align:right;color:#dc2626">${fmt(r.pisR)}</td>
            <td style="text-align:center;color:#dc2626;font-weight:600">${pct(r.cofins)}</td>
            <td style="text-align:right;color:#dc2626">${fmt(r.cofinsR)}</td>
            <td style="text-align:center;color:#7c3aed;font-weight:600">${pct(r.ipi)}</td>
            <td style="text-align:right;color:#7c3aed">${fmt(r.ipiR)}</td>
            <td style="text-align:right;font-weight:600;color:#1e293b">${fmt(r.valorImpostos)}</td>
            <td style="text-align:center;color:#d97706;font-weight:600">${pct(r.comissao)}</td>
            <td style="text-align:right;color:#d97706">${fmt(r.comissaoR)}</td>
            <td style="text-align:right;font-weight:800;color:#16a34a;font-size:0.9rem">${fmt(r.precoFinal)}</td>
            <td style="text-align:center;font-weight:700;color:${corMargem(r.margem)}">${r.margem.toFixed(1)}%</td>
            <td style="text-align:center;font-size:0.8rem">${r.propostaNum ? `<span style="color:#1d4ed8;font-weight:600">${r.propostaNum}</span>` : '<span style="color:#94a3b8">—</span>'}</td>
        </tr>
    `).join('');
}

function exportarConsultaPrecificacao() {
    const rows = [];
    (precificacoesCliente||[]).forEach(precif => {
        const cliente = (clientes||[]).find(c => String(c.id) === String(precif.clienteId));
        const cnpjLimpo = (cliente?.cnpj||'').replace(/\D/g,'');
        const tipo = cnpjLimpo.length === 14 ? 'PJ' : 'PF';
        const proposta = precif.propostaId ? (propostas||[]).find(p => p.id === precif.propostaId) : null;

        (precif.itens||[]).forEach(item => {
            rows.push({
                'Cliente':       precif.clienteNome || cliente?.nome || '',
                'UF':            precif.clienteUF   || cliente?.uf   || '',
                'Tipo':          tipo,
                'Versão':        'v' + (precif.versao||1),
                'Data':          new Date(precif.dataCriacao||new Date()).toLocaleDateString('pt-BR'),
                'Status':        precif.status,
                'Proposta':      proposta?.numero || '',
                'Produto':       item.produto   || '',
                'NCM':           item.ncm       || '',
                'CI (R$)':       item.ci        || 0,
                'Taxa (%)':      item.taxa      || 0,
                'ROI (%)':       item.roi       || 0,
                'Valor Base':    item.valorBase || 0,
                'ICMS (%)':      item.icms      || 0,
                'ICMS (R$)':     item.icmsR     || 0,
                'PIS (%)':       item.pis       || 0,
                'PIS (R$)':      item.pisR      || 0,
                'COFINS (%)':    item.cofins    || 0,
                'COFINS (R$)':   item.cofinsR   || 0,
                'IPI (%)':       item.ipi       || 0,
                'IPI (R$)':      item.ipiR      || 0,
                'c/ Impostos':   item.valorImpostos || 0,
                'Comissão (%)':  item.comissao  || 0,
                'Comissão (R$)': item.comissaoR || 0,
                'Preço Final':   item.precoFinal || 0,
                'Margem (%)':    item.margem ? item.margem.toFixed(2) : item.precoFinal > 0 ? (((item.precoFinal - item.ci) / item.precoFinal)*100).toFixed(2) : 0,
            });
        });
    });

    if (!rows.length) {
        mostrarNotificacao('Nenhuma precificação para exportar.', 'warning');
        return;
    }

    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Consulta Precificação');
    XLSX.writeFile(wb, 'consulta_precificacao_' + new Date().toISOString().split('T')[0] + '.xlsx');
    mostrarNotificacao('Exportado com sucesso!', 'success');
}

function _obterCenarioComparativo(n) {
        const id = document.getElementById('compCliente' + n)?.value;
        if (!id) return null;
        const cliente = (clientes || []).find(x => String(x.id) === String(id));
        if (!cliente) return null;
        const uf = (cliente.uf || '').toUpperCase();
        const cnpjLimpo = (cliente.cnpj || '').replace(/\D/g, '');
        const tipo = cnpjLimpo.length === 14 ? 'PJ' : 'PF';
        const taxa = parseFloat(document.getElementById('compTaxa' + n)?.value)
                || parseFloat(document.getElementById('taxaPadrao')?.value)
                || 20;
        const roi = parseFloat(document.getElementById('compROI' + n)?.value)
                || parseFloat(document.getElementById('roiPadrao')?.value)
                || 30;
        const comissao = parseFloat(document.getElementById('comissaoPadrao')?.value) || 5;
        return { id, nome: cliente.nome, uf, tipo, taxa, roi, comissao };
}

function _calcularPrecoComparativo(nomeProduto, estado, tipoPessoa, taxa, roi, comissao) {
        const prec = precificacao[nomeProduto] || {};
        const aliq = tabelaAliquotas[nomeProduto] || {};
        const ci = parseFloat(prec.ci) || 0;
        if (ci === 0) return null;

        const taxaPct = Number.isFinite(Number(taxa)) ? Number(taxa) : 20;
        const roiPct = Number.isFinite(Number(roi)) ? Number(roi) : 30;
        const comissaoPct = Number.isFinite(Number(comissao)) ? Number(comissao) : 5;
        const pis = parseFloat(aliq.pis) || parseFloat(document.getElementById('pisPadrao')?.value) || 0;
        const cofins = parseFloat(aliq.cofins) || parseFloat(document.getElementById('cofinsPadrao')?.value) || 0;
        const ipi = parseFloat(aliq.ipi) || 0;
        const icms = buscarAliquotaICMS(estado, tipoPessoa, nomeProduto);

        const valorBase = ci * (1 + taxaPct / 100) * (1 + roiPct / 100);
        const icmsR = valorBase * icms / 100;
        const pisR = valorBase * pis / 100;
        const cofinsR = valorBase * cofins / 100;
        const valorImpostos = valorBase + icmsR + pisR + cofinsR;
        const ipiR = valorImpostos * ipi / 100;
        const valorTotal = valorImpostos + ipiR;
        const comissaoR = valorBase * comissaoPct / 100;
        const precoFinal = valorTotal + comissaoR;
        const margem = precoFinal > 0 ? ((precoFinal - ci) / precoFinal) * 100 : 0;

        return {
                precoFinal,
                margem,
                taxa: taxaPct,
                roi: roiPct,
                comissao: comissaoPct
        };
}

function calcularComparativo() {
        const resultado = document.getElementById('compResultado');
        const empty = document.getElementById('compEmpty');
        const header = document.getElementById('compHeader');
        const body = document.getElementById('compBody');
        if (!resultado || !empty || !header || !body) return;

        const fmt = v => 'R$ ' + parseFloat(v).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
        const fmtPct = v => parseFloat(v).toFixed(1) + '%';
        const corMargem = m => m >= 30 ? '#16a34a' : m >= 15 ? '#d97706' : '#dc2626';

        const clientesSel = [1, 2, 3].map(_obterCenarioComparativo).filter(Boolean);

        if (clientesSel.length < 2) {
                resultado.style.display = 'none';
                empty.style.display = 'block';
                return;
        }

        const nCols = clientesSel.length;
        header.innerHTML = `
        <tr>
            <th style="text-align:left;padding-left:15px;min-width:220px;
                                 position:sticky;left:0;background:#1e3a5f;z-index:2">
                Produto
            </th>
            ${clientesSel.map((c, i) => `
                <th colspan="2" style="text-align:center;
                        background:${['#1e3a5f','#2d5a8b','#3d7ab5'][i]};min-width:200px">
                    ${_escapeHtml(c.nome)}<br>
                    <span style="font-size:0.7rem;font-weight:400;color:rgba(255,255,255,0.75)">
                        ${_escapeHtml(c.uf)} / ${c.tipo} | Taxa ${c.taxa}% | ROI ${c.roi}%
                    </span>
                </th>`).join('')}
            <th style="min-width:90px;text-align:center">Diferença</th>
        </tr>
        <tr>
            <th style="position:sticky;left:0;background:#161b22"></th>
            ${clientesSel.map(() => `
                <th style="font-size:0.7rem;font-weight:400;background:#161b22;color:#8b949e">
                    Preço Final
                </th>
                <th style="font-size:0.7rem;font-weight:400;background:#161b22;color:#8b949e">
                    Margem
                </th>`).join('')}
            <th style="background:#161b22"></th>
        </tr>`;

        const produtos = estoque.produtos || [];
        body.innerHTML = produtos.map(produto => {
                const precos = clientesSel.map(c => _calcularPrecoComparativo(produto.nome, c.uf, c.tipo, c.taxa, c.roi, c.comissao));

                if (precos.every(r => !r)) {
                        return `<tr style="opacity:0.4">
                <td style="text-align:left;padding-left:15px;position:sticky;
                                     left:0;background:#fff;z-index:1">${_escapeHtml(produto.nome)}</td>
                <td colspan="${nCols * 2 + 1}" style="text-align:center;color:#94a3b8;
                                                                                    font-size:0.8rem">CI não configurado</td>
            </tr>`;
                }

                const valores = precos.map(r => r?.precoFinal || 0).filter(v => v > 0);
                const maxP = valores.length ? Math.max(...valores) : 0;
                const minP = valores.length ? Math.min(...valores) : 0;
                const diff = maxP - minP;
                const diffPct = minP > 0 ? (diff / minP * 100) : 0;

                return `<tr>
            <td style="text-align:left;padding-left:15px;font-weight:500;
                                 position:sticky;left:0;background:#fff;z-index:1;
                                 border-right:1px solid #e2e8f0">${_escapeHtml(produto.nome)}</td>
            ${precos.map(r => {
                if (!r) return '<td colspan="2" style="text-align:center;color:#94a3b8">—</td>';
                const isMax = r.precoFinal === maxP && nCols > 1;
                const isMin = r.precoFinal === minP && nCols > 1 && maxP !== minP;
                return `
                    <td style="text-align:right;padding-right:8px;font-weight:700;
                                         color:${isMin ? '#16a34a' : isMax ? '#dc2626' : '#1e3a5f'};
                                         background:${isMin ? '#f0fdf4' : isMax ? '#fef2f2' : 'transparent'}">
                        ${fmt(r.precoFinal)}
                        ${isMin ? '<span style="font-size:0.65rem">▼min</span>' : ''}
                        ${isMax ? '<span style="font-size:0.65rem">▲max</span>' : ''}
                    </td>
                    <td style="text-align:center;font-size:0.82rem;font-weight:600;
                                         color:${corMargem(r.margem)}">
                        ${fmtPct(r.margem)}
                    </td>`;
            }).join('')}
            <td style="text-align:center;font-weight:600;color:#d97706">
                ${diff > 0 ? fmt(diff) + '<br><span style="font-size:0.72rem">(' + fmtPct(diffPct) + ')</span>' : '—'}
            </td>
        </tr>`;
        }).join('');

        resultado.style.display = 'block';
        empty.style.display = 'none';
}

function exportarComparativo() {
        const clientesSel = [1, 2, 3].map(_obterCenarioComparativo).filter(Boolean);
        if (clientesSel.length < 2) {
                alert('Selecione ao menos 2 clientes para exportar o comparativo.');
                return;
        }

        const rows = (estoque.produtos || []).map(p => {
                const row = { 'Produto': p.nome };
                clientesSel.forEach(c => {
                        const r = _calcularPrecoComparativo(p.nome, c.uf, c.tipo, c.taxa, c.roi, c.comissao);
                        row[c.nome + ' - Preço'] = r?.precoFinal || '';
                        row[c.nome + ' - Margem%'] = r?.margem?.toFixed(1) || '';
                });
                return row;
        });

        const ws = XLSX.utils.json_to_sheet(rows);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Comparativo');
        XLSX.writeFile(wb, 'comparativo_precos_' + new Date().toISOString().split('T')[0] + '.xlsx');
}

function renderizarRastreabilidade() {
        const selCliente = document.getElementById('rastreaCliente');
        if (selCliente) {
                const currentVal = selCliente.value;
                selCliente.innerHTML = '<option value="">Todos os clientes</option>'
                        + (clientes || []).slice().sort((a, b) => (a.nome || '').localeCompare(b.nome || ''))
                                .map(c => `<option value="${c.id}" ${String(c.id) === String(currentVal) ? 'selected' : ''}>${_escapeHtml(c.nome || '')}</option>`).join('');
        }

        const filtroId = selCliente?.value || '';
        const periodo = document.getElementById('rastreaPeriodo')?.value || 'todos';
        const agora = new Date();

        const inPeriod = iso => {
                if (!iso || periodo === 'todos') return true;
                const d = new Date(iso);
                if (periodo === 'mes') return d.getMonth() === agora.getMonth() && d.getFullYear() === agora.getFullYear();
                if (periodo === 'trimestre') return Math.floor(d.getMonth() / 3) === Math.floor(agora.getMonth() / 3) && d.getFullYear() === agora.getFullYear();
                if (periodo === 'ano') return d.getFullYear() === agora.getFullYear();
                return true;
        };

        const clientesFilt = (clientes || []).filter(c => !filtroId || String(c.id) === String(filtroId));
        const fmt = v => 'R$ ' + parseFloat(v || 0).toLocaleString('pt-BR', { minimumFractionDigits: 2 });

        const statusColor = {
                rascunho: '#64748b', enviada: '#1d4ed8', aceita: '#166534',
                recusada: '#dc2626', expirada: '#d97706', convertida: '#0ea5e9',
                aguardando_aprovacao: '#7c3aed'
        };

        const container = document.getElementById('painelRastreabilidade');
        if (!container) return;

        const html = clientesFilt.map(cliente => {
                const precifs = (precificacoesCliente || [])
                        .filter(p => String(p.clienteId) === String(cliente.id) && inPeriod(p.dataCriacao))
                        .sort((a, b) => new Date(b.dataCriacao) - new Date(a.dataCriacao));
                const propostasC = (propostas || [])
                        .filter(p => p.cliente === cliente.nome && inPeriod(p.dataCriacao || p.data))
                        .sort((a, b) => new Date(b.dataCriacao || b.data || 0) - new Date(a.dataCriacao || a.data || 0));
                const vendasC = (estoque.registroVendas || [])
                        .filter(v => v.loja === cliente.nome && inPeriod(v.data))
                        .sort((a, b) => new Date(b.data || 0) - new Date(a.data || 0));

                if (!precifs.length && !propostasC.length && !vendasC.length) return '';

                const totalFat = vendasC.reduce((s, v) => s + (v.valorTotal || 0), 0);
                const eventos = [
                        ...precifs.map(p => ({ ...p, _tipo: 'precif', _date: p.dataCriacao })),
                        ...propostasC.map(p => ({ ...p, _tipo: 'proposta', _date: p.dataCriacao || p.data })),
                        ...vendasC.map(v => ({ ...v, _tipo: 'contrato', _date: v.data }))
                ].sort((a, b) => new Date(b._date || 0) - new Date(a._date || 0));

                return `
            <div style="background:#fff;border:1px solid #e2e8f0;border-radius:12px;
                                    margin-bottom:20px;overflow:hidden">
                <div style="background:linear-gradient(135deg,#1e3a5f,#2d5a8b);
                                        padding:14px 20px;display:flex;justify-content:space-between;align-items:center">
                    <div>
                        <div style="font-size:1rem;font-weight:700;color:#fff">${_escapeHtml(cliente.nome || '')}</div>
                        <div style="font-size:0.8rem;color:rgba(255,255,255,0.7);margin-top:2px">
                            ${_escapeHtml(cliente.uf || '—')} · ${((cliente.cnpj || '').replace(/\D/g, '').length === 14 ? 'PJ' : 'PF')}
                        </div>
                    </div>
                    <div style="text-align:right">
                        <div style="font-size:0.75rem;color:rgba(255,255,255,0.6)">Total faturado</div>
                        <div style="font-size:1.2rem;font-weight:800;color:#7ee787">${fmt(totalFat)}</div>
                    </div>
                </div>
                <div style="display:flex;border-bottom:1px solid #f1f5f9;background:#f8fafc">
                    ${[
                            { icon: '💰', label: 'Precificações', val: precifs.length, cor: '#c9a227' },
                            { icon: '📋', label: 'Propostas', val: propostasC.length, cor: '#1d4ed8' },
                            { icon: '✅', label: 'Aceitas', val: propostasC.filter(p => p.status === 'aceita').length, cor: '#16a34a' },
                            { icon: '📄', label: 'Contratos', val: vendasC.length, cor: '#0ea5e9' },
                    ].map(s => `<div style="flex:1;padding:10px;text-align:center;border-right:1px solid #f1f5f9">
                        <div>${s.icon}</div>
                        <div style="font-size:1.2rem;font-weight:700;color:${s.cor}">${s.val}</div>
                        <div style="font-size:0.72rem;color:#64748b">${s.label}</div>
                    </div>`).join('')}
                </div>
                <div style="padding:16px 20px">
                    <div style="position:relative;padding-left:24px">
                        <div style="position:absolute;left:8px;top:0;bottom:0;width:2px;background:#e2e8f0"></div>
                        ${eventos.map(item => {
                                const d = new Date(item._date || new Date()).toLocaleDateString('pt-BR');
                                if (item._tipo === 'precif') return `
                                <div style="position:relative;margin-bottom:14px">
                                    <div style="position:absolute;left:-19px;top:4px;width:12px;height:12px;
                                                            border-radius:50%;background:#c9a227;border:2px solid #fff;
                                                            box-shadow:0 0 0 2px #c9a227"></div>
                                    <div style="background:#fffbf0;border:1px solid #fed7aa;border-radius:8px;padding:10px 14px">
                                        <div style="display:flex;justify-content:space-between">
                                            <span style="font-weight:600;color:#92400e;font-size:0.85rem">
                                                💰 Precificação v${item.versao || '—'}
                                            </span>
                                            <span style="font-size:0.78rem;color:#94a3b8">${d}</span>
                                        </div>
                                        <div style="font-size:0.78rem;color:#64748b;margin-top:3px">
                                            Taxa: ${item.taxa || '—'}% · ROI: ${item.roi || '—'}% ·
                                            ${(item.itens || []).length} produto(s)
                                            ${item.descricao ? ` · &quot;${_escapeHtml(item.descricao)}&quot;` : ''}
                                        </div>
                                        ${item.propostaId
                                                ? `<div style="font-size:0.72rem;color:#0ea5e9;margin-top:3px">
                                                     → Gerou proposta ${_escapeHtml((propostas || []).find(p => p.id === item.propostaId)?.numero || '—')}
                                                 </div>` : ''}
                                    </div>
                                </div>`;
                                if (item._tipo === 'proposta') {
                                        const sc = statusColor[item.status] || '#64748b';
                                        return `
                                    <div style="position:relative;margin-bottom:14px">
                                        <div style="position:absolute;left:-19px;top:4px;width:12px;height:12px;
                                                                border-radius:50%;background:${sc};border:2px solid #fff;
                                                                box-shadow:0 0 0 2px ${sc}"></div>
                                        <div style="background:#f0f9ff;border:1px solid #bae6fd;border-radius:8px;padding:10px 14px">
                                            <div style="display:flex;justify-content:space-between;align-items:center">
                                                <span style="font-weight:600;color:#0369a1;font-size:0.85rem">
                                                    📋 Proposta ${_escapeHtml(item.numero || '')}
                                                </span>
                                                <div style="display:flex;gap:6px;align-items:center">
                                                    <span style="background:${sc}20;color:${sc};font-size:0.7rem;
                                                                             font-weight:600;padding:1px 8px;border-radius:20px">
                                                        ${_escapeHtml(item.status || '')}
                                                    </span>
                                                    <span style="font-size:0.78rem;color:#94a3b8">${d}</span>
                                                </div>
                                            </div>
                                            <div style="display:flex;justify-content:space-between;margin-top:4px">
                                                <span style="font-size:0.78rem;color:#64748b">
                                                    ${(item.itens || []).length} item(ns)
                                                </span>
                                                <span style="font-size:0.85rem;font-weight:700;color:#16a34a">
                                                    ${fmt(item.valorTotal)}
                                                </span>
                                            </div>
                                            ${item.contratoNumero
                                                ? `<div style="font-size:0.72rem;color:#0ea5e9;margin-top:3px">
                                                         → Contrato #${_escapeHtml(item.contratoNumero)}
                                                     </div>` : ''}
                                        </div>
                                    </div>`;
                                }
                                if (item._tipo === 'contrato') return `
                                <div style="position:relative;margin-bottom:14px">
                                    <div style="position:absolute;left:-19px;top:4px;width:12px;height:12px;
                                                            border-radius:50%;background:#16a34a;border:2px solid #fff;
                                                            box-shadow:0 0 0 2px #16a34a"></div>
                                    <div style="background:#f0fdf4;border:1px solid #86efac;border-radius:8px;padding:10px 14px">
                                        <div style="display:flex;justify-content:space-between;align-items:center">
                                            <span style="font-weight:700;color:#166534;font-size:0.9rem">
                                                📄 Contrato #${_escapeHtml(item.contrato || '')}
                                            </span>
                                            <span style="font-size:0.78rem;color:#94a3b8">${d}</span>
                                        </div>
                                        <div style="display:flex;justify-content:space-between;margin-top:4px">
                                            <span style="font-size:0.78rem;color:#64748b">
                                                Rep: ${_escapeHtml(item.representante || '')} · ${(item.itens || []).length} item(ns)
                                            </span>
                                            <span style="font-size:1rem;font-weight:800;color:#16a34a">
                                                ${fmt(item.valorTotal)}
                                            </span>
                                        </div>
                                    </div>
                                </div>`;
                                return '';
                        }).join('')}
                    </div>
                </div>
            </div>`;
        }).filter(Boolean).join('');

        container.innerHTML = html || `
        <div style="text-align:center;padding:60px;color:#94a3b8">
            <div style="font-size:3rem;margin-bottom:12px">🔗</div>
            <div>Nenhuma rastreabilidade encontrada</div>
        </div>`;
}

function exportarRastreabilidade() {
        const rows = [];
        (clientes || []).forEach(c => {
                (precificacoesCliente || []).filter(p => String(p.clienteId) === String(c.id)).forEach(p => rows.push({
                        'Cliente': c.nome,
                        'UF': c.uf || '',
                        'Tipo': ((c.cnpj || '').replace(/\D/g, '').length === 14 ? 'PJ' : 'PF'),
                        'Evento': 'Precificação',
                        'Data': new Date(p.dataCriacao).toLocaleDateString('pt-BR'),
                        'Versão/Nº': 'v' + (p.versao || 1),
                        'Valor': '',
                        'Status': p.status,
                        'Ligação': p.propostaId ? ((propostas || []).find(x => x.id === p.propostaId)?.numero || '') : ''
                }));
                (propostas || []).filter(p => p.cliente === c.nome).forEach(p => rows.push({
                        'Cliente': c.nome,
                        'UF': c.uf || '',
                        'Tipo': ((c.cnpj || '').replace(/\D/g, '').length === 14 ? 'PJ' : 'PF'),
                        'Evento': 'Proposta',
                        'Data': new Date(p.dataCriacao || p.data || new Date()).toLocaleDateString('pt-BR'),
                        'Versão/Nº': p.numero,
                        'Valor': p.valorTotal || 0,
                        'Status': p.status,
                        'Ligação': p.contratoNumero || ''
                }));
                (estoque.registroVendas || []).filter(v => v.loja === c.nome).forEach(v => rows.push({
                        'Cliente': c.nome,
                        'UF': c.uf || '',
                        'Tipo': ((c.cnpj || '').replace(/\D/g, '').length === 14 ? 'PJ' : 'PF'),
                        'Evento': 'Contrato',
                        'Data': new Date(v.data || new Date()).toLocaleDateString('pt-BR'),
                        'Versão/Nº': '#' + (v.contrato || ''),
                        'Valor': v.valorTotal || 0,
                        'Status': 'fechado',
                        'Ligação': ''
                }));
        });
        const ws = XLSX.utils.json_to_sheet(rows);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Rastreabilidade');
        XLSX.writeFile(wb, 'rastreabilidade_' + new Date().toISOString().split('T')[0] + '.xlsx');
}

// ── Precificação por Cliente: utilitários e calculadora (em tempo real, não persiste dados) ──
function popularSelectClientesPrecif() {
    const select = document.getElementById('precifClienteSelect');
    if (!select) return;
    const valorAtual = select.value;
    const list = (clientes || []).slice().sort((a,b) => (a.nome||'').localeCompare(b.nome||''));
    const optionsHtml = list.map(c => {
        const tipo = c.tipoPessoa || (((c.cnpj||'').replace(/\D/g,''))).length === 14 ? 'PJ' : 'PF';
        const ultima = (propostas || [])
            .filter(p => p.cliente === c.nome)
            .sort((a,b) => new Date(b.dataCriacao||0) - new Date(a.dataCriacao||0))[0];
        const badge = ultima ? ` [${String(ultima.status || '').toUpperCase()}]` : '';
        return `<option value="${c.id}">${_escapeHtml(c.nome||'-')} — ${_escapeHtml(c.uf||'??')} (${tipo})${badge}</option>`;
    }).join('');
    select.innerHTML = '<option value="">Selecione um cliente...</option>' + optionsHtml;
    if (valorAtual) select.value = valorAtual;

    // Atualizar labels de padrão
    const taxaPad = document.getElementById('taxaPadrao')?.value || '—';
    const roiPad  = document.getElementById('roiPadrao')?.value  || '—';
    const comPad  = document.getElementById('comissaoPadrao')?.value || '—';
    const elTaxa = document.getElementById('precifTaxaPadraoLabel'); if (elTaxa) elTaxa.textContent = taxaPad;
    const elROI  = document.getElementById('precifROIPadraoLabel'); if (elROI) elROI.textContent = roiPad;
    const elCom  = document.getElementById('precifComissaoPadraoLabel'); if (elCom) elCom.textContent = comPad;
}

function obterUltimaPrecificacaoCliente(clienteId) {
    const registros = (precificacoesCliente || []).filter(p => String(p.clienteId) === String(clienteId));
    if (!registros.length) return null;
    return registros.sort((a, b) => new Date(b.dataCriacao || 0) - new Date(a.dataCriacao || 0))[0];
}

function limparAvisoCI() {
    const oldAviso = document.getElementById('avisoCI');
    if (oldAviso) oldAviso.remove();
}

function atualizarAvisoCIDivergente(precifSalva) {
    if (!precifSalva) {
        limparAvisoCI();
        return;
    }

    const alertaCI = [];
    (precifSalva.itens || []).forEach(item => {
        const ciAtual = parseFloat(precificacao[item.produto]?.ci) || 0;
        const ciSalvo = parseFloat(item.ci) || 0;
        if (ciAtual !== ciSalvo && ciSalvo > 0) {
            alertaCI.push({
                produto: item.produto,
                ciSalvo,
                ciAtual,
                diff: ((ciAtual - ciSalvo) / ciSalvo * 100).toFixed(1)
            });
        }
    });

    if (alertaCI.length > 0) {
        const avisoEl = document.createElement('div');
        avisoEl.style.cssText = 'background:#fff8f0; border:1px solid #fed7aa; border-radius:8px; padding:10px 14px; margin-top:10px; font-size:0.82rem; color:#92400e';
        avisoEl.innerHTML = `
            ⚠️ <strong>${alertaCI.length} produto(s)</strong> tiveram o CI alterado desde que esta precificação foi salva:<br>
            <ul style="margin:6px 0 0 16px; padding:0">
                ${alertaCI.map(a => `
                    <li>
                        ${_escapeHtml(a.produto)}:
                        salvo R$ ${a.ciSalvo.toFixed(2)} →
                        atual R$ ${a.ciAtual.toFixed(2)}
                        <span style="color:${parseFloat(a.diff) > 0 ? '#dc2626' : '#16a34a'}">
                            (${parseFloat(a.diff) > 0 ? '+' : ''}${a.diff}%)
                        </span>
                    </li>
                `).join('')}
            </ul>
            <div style="margin-top:8px">
                A precificação salva usa os valores originais.
                Para recalcular com o CI atual, clique em <strong>Calcular</strong> e salve novamente.
            </div>
        `;
        const banner = document.getElementById('precifClienteBanner');
        if (banner && banner.parentNode) {
            limparAvisoCI();
            avisoEl.id = 'avisoCI';
            banner.parentNode.insertBefore(avisoEl, banner.nextSibling);
        }
    } else {
        limparAvisoCI();
    }
}

function aplicarEstadoPrecificacaoSalva(registro) {
    if (!registro) return;

    try {
        retidPorProduto = registro.retidPorProduto || {};
        beneficioFiscalAtivo = registro.beneficioFiscalAtivo || false;
        beneficiosPorProduto = registro.beneficiosPorProduto || {};
    } catch (e) {}

    exibindoPrecifSalva = true;
    precifSalvaCarregada = registro;
    ultimaPrecificacaoCalculada = registro;
    try { ultimaVersaoSalva = registro.versao || null; } catch (e) {}

    try {
        const cliente = (clientes || []).find(c => String(c.id) === String(registro.clienteId));
        const select = document.getElementById('precifClienteSelect');
        if (select) select.value = registro.clienteId || '';
        const ufEl = document.getElementById('precifClienteUF'); if (ufEl) ufEl.value = registro.uf || '—';
        const tipoEl = document.getElementById('precifClienteTipo'); if (tipoEl) tipoEl.value = registro.tipoPessoa || '—';
        const banner = document.getElementById('precifClienteBanner'); if (banner) banner.style.display = 'flex';
        const name = document.getElementById('precifBannerNome'); if (name) name.textContent = cliente?.nome || registro.clienteNome || '—';
        const doc = document.getElementById('precifBannerDoc'); if (doc) doc.textContent = cliente?.cnpj || cliente?.cpf || '—';
        const ufBanner = document.getElementById('precifBannerUF'); if (ufBanner) ufBanner.textContent = registro.uf || cliente?.uf || cliente?.estado || '—';
        const tipoBanner = document.getElementById('precifBannerTipo'); if (tipoBanner) tipoBanner.textContent = registro.tipoPessoa || cliente?.tipoPessoa || '—';
        const contato = document.getElementById('precifBannerContato'); if (contato) contato.textContent = cliente?.contato || cliente?.email || '—';
        const taxaOverride = document.getElementById('precifTaxaOverride'); if (taxaOverride) taxaOverride.value = registro.taxa ?? '';
        const roiOverride = document.getElementById('precifROIOverride'); if (roiOverride) roiOverride.value = registro.roi ?? '';
        const comOverride = document.getElementById('precifComissaoOverride'); if (comOverride) comOverride.value = registro.comissao ?? '';
        const validadeEl = document.getElementById('precifValidadeDias'); if (validadeEl) validadeEl.value = registro.validadeDias ?? 30;

        const filtros = registro.filtros || {};
        const filtroProduto = document.getElementById('precifFiltroProduto'); if (filtroProduto) filtroProduto.value = filtros.filtroProduto || '';
        const filtroNCM = document.getElementById('precifFiltroNCM'); if (filtroNCM) filtroNCM.value = filtros.filtroNCM || '';
        const filtroCI = document.getElementById('precifFiltroCI'); if (filtroCI) filtroCI.value = filtros.filtroCI || 'todos';

        const cb = document.getElementById('precifBeneficioFiscal'); if (cb) cb.checked = !!beneficioFiscalAtivo;
        const painel = document.getElementById('painelBeneficioFiscal'); if (painel) painel.style.display = beneficioFiscalAtivo ? 'block' : 'none';
        if (beneficioFiscalAtivo) renderizarTabelaBeneficios();
    } catch (e) {}

    calcularPrecificacaoPorCliente();
}

function selecionarClientePrecif() {
        const select = document.getElementById('precifClienteSelect');
        if (!select) return;
    const clienteId = select.value;
    exibindoPrecifSalva = false;
    precifSalvaCarregada = null;
    // Resetar benefícios/RETID ao trocar de cliente
    beneficioFiscalAtivo = false;
    beneficiosPorProduto = {};
    retidPorProduto = {};
    try {
        const cbBenef = document.getElementById('precifBeneficioFiscal'); if (cbBenef) cbBenef.checked = false;
        const painel = document.getElementById('painelBeneficioFiscal'); if (painel) painel.style.display = 'none';
    } catch (e) {}
        const cliente = (clientes || []).find(c => String(c.id) === String(clienteId));
        if (!cliente) {
                const banner = document.getElementById('precifClienteBanner'); if (banner) banner.style.display = 'none';
                const resultado = document.getElementById('precifClienteResultado'); if (resultado) resultado.style.display = 'none';
                const empty = document.getElementById('precifClienteEmpty'); if (empty) empty.style.display = 'block';
            try { atualizarStatusPropostaNaPrecif(''); } catch (e) {}
            limparAvisoCI();
                return;
        }

        const uf = (cliente.uf || cliente.estado || '').toUpperCase().trim();
        const cnpjLimpo = (cliente.cnpj || '').replace(/\D/g,'');
        const tipoPessoa = cliente.tipoPessoa || (cnpjLimpo.length === 14 ? 'PJ' : 'PF');

        const ufEl = document.getElementById('precifClienteUF'); if (ufEl) ufEl.value = uf || '—';
        const tipoEl = document.getElementById('precifClienteTipo'); if (tipoEl) tipoEl.value = tipoPessoa;

        const banner = document.getElementById('precifClienteBanner');
        if (banner) banner.style.display = 'flex';
        const name = document.getElementById('precifBannerNome'); if (name) name.textContent = cliente.nome || '—';
        const doc = document.getElementById('precifBannerDoc'); if (doc) doc.textContent = cliente.cnpj || cliente.cpf || '—';
        const ufB = document.getElementById('precifBannerUF'); if (ufB) ufB.textContent = uf || '—';
        const tipoB = document.getElementById('precifBannerTipo'); if (tipoB) tipoB.textContent = tipoPessoa;
        const contato = document.getElementById('precifBannerContato'); if (contato) contato.textContent = cliente.contato || cliente.email || '—';

        calcularPrecificacaoPorCliente({ forcarAtual: true });

        try { renderizarHistoricoPrecif(cliente.id); } catch (e) {}

        // Mostrar se existe precificação salva para este cliente
        try {
            const precifSalva = obterUltimaPrecificacaoCliente(cliente.id);
            const infoEl = document.getElementById('precifSalvoInfo');
            if (precifSalva && infoEl) {
                const dt = precifSalva.dataCriacao ? new Date(precifSalva.dataCriacao).toLocaleString('pt-BR') : '';
                infoEl.innerHTML = `<span style="color:#16a34a; font-size:0.82rem">✅ Precificação salva em ${dt}</span> <button onclick="carregarPrecificacaoSalva('${cliente.id}')" class="btn btn-outline btn-sm" style="margin-left:8px">↩ Carregar salva</button>`;
                atualizarAvisoCIDivergente(precifSalva);
            } else {
                const infoEl2 = document.getElementById('precifSalvoInfo'); if (infoEl2) infoEl2.innerHTML = '';
                limparAvisoCI();
            }
        } catch (e) {}
        try { atualizarStatusPropostaNaPrecif(cliente.id); } catch (e) {}
}

function calcularPrecificacaoPorCliente(opcoes = {}) {
    if (opcoes.forcarAtual) {
        exibindoPrecifSalva = false;
        precifSalvaCarregada = null;
    }

    const select = document.getElementById('precifClienteSelect');
    const clienteId = select?.value;
    const cliente = (clientes || []).find(c => String(c.id) === String(clienteId));
    if (!cliente) {
        const resultado = document.getElementById('precifClienteResultado'); if (resultado) resultado.style.display = 'none';
        const empty = document.getElementById('precifClienteEmpty'); if (empty) empty.style.display = 'block';
        const banner = document.getElementById('precifClienteBanner'); if (banner) banner.style.display = 'none';
        ultimaPrecificacaoCalculada = null;
        limparAvisoCI();
        return;
    }

    const uf = (cliente.uf || cliente.estado || '').toUpperCase().trim();
    const cnpjLimpo = (cliente.cnpj || '').replace(/\D/g,'');
    const tipoPessoa = cliente.tipoPessoa || (cnpjLimpo.length === 14 ? 'PJ' : 'PF');

    try { document.getElementById('precifClienteUF').value = uf || '—'; } catch (e) {}
    try { document.getElementById('precifClienteTipo').value = tipoPessoa; } catch (e) {}

    const taxaOverride = parseFloat(document.getElementById('precifTaxaOverride')?.value);
    const roiOverride = parseFloat(document.getElementById('precifROIOverride')?.value);
    const comOverride = parseFloat(document.getElementById('precifComissaoOverride')?.value);

    const taxaGlobal = parseFloat(document.getElementById('taxaPadrao')?.value) || 20;
    const roiGlobal = parseFloat(document.getElementById('roiPadrao')?.value) || 30;
    const comGlobal = parseFloat(document.getElementById('comissaoPadrao')?.value) || 5;

    const taxaFinal = !isNaN(taxaOverride) ? taxaOverride : taxaGlobal;
    const roiFinal  = !isNaN(roiOverride)  ? roiOverride  : roiGlobal;
    const comFinal  = !isNaN(comOverride)  ? comOverride  : comGlobal;

    // popular NCM select se necessário
    try {
        const selectNCM = document.getElementById('precifFiltroNCM');
        if (selectNCM && selectNCM.options.length <= 1) {
            const keys = Object.keys(impostosEditaveis || {}).sort();
            keys.forEach(ncm => {
                const opt = document.createElement('option');
                opt.value = ncm;
                opt.textContent = ncm + (impostosEditaveis[ncm] && impostosEditaveis[ncm].descricao ? ' — ' + impostosEditaveis[ncm].descricao : '');
                selectNCM.appendChild(opt);
            });
        }
    } catch(e) {}

    // filtros
    const filtroProduto = (document.getElementById('precifFiltroProduto')?.value || '').toLowerCase().trim();
    const filtroNCM = document.getElementById('precifFiltroNCM')?.value || '';
    const filtroCI = document.getElementById('precifFiltroCI')?.value || 'todos';

    const produtos = estoque.produtos || [];
    const fmt = v => 'R$ ' + _fmtMoeda(v);
    const pct = v => (Number.isFinite(Number(v)) ? Number(v).toFixed(2) : '0.00') + '%';
    const corMargem = m => m >= 30 ? '#16a34a' : m >= 15 ? '#d97706' : '#dc2626';

    let totalFaturamento = 0;
    let produtosSemCI = 0;
    let produtosCalculados = 0;
    let abaixoCount = 0;

    const produtosFiltrados = produtos.filter(produto => {
        try {
            if (filtroProduto && !(produto.nome || '').toLowerCase().includes(filtroProduto)) return false;
            const pNcm = produto.ncm || detectarNCM(produto.nome) || '';
            if (filtroNCM && pNcm !== filtroNCM) return false;
            const itemSalvo = exibindoPrecifSalva && precifSalvaCarregada
                ? (precifSalvaCarregada.itens || []).find(i => i.produto === produto.nome)
                : null;
            const ci = itemSalvo ? (parseFloat(itemSalvo.ci) || 0) : (parseFloat((precificacao[produto.nome] || {}).ci) || 0);
            if (filtroCI === 'comCI' && ci === 0) return false;
            if (filtroCI === 'semCI' && ci > 0) return false;
            return true;
        } catch (e) { return false; }
    });

    const itensCalculados = [];
    const rows = produtosFiltrados.map(produto => {
        const prec = precificacao[produto.nome] || {};
        const aliq = tabelaAliquotas[produto.nome] || {};
        const ncm = produto.ncm || detectarNCM(produto.nome) || '—';

        let ci;
        if (exibindoPrecifSalva && precifSalvaCarregada) {
            const itemSalvo = (precifSalvaCarregada.itens || []).find(i => i.produto === produto.nome);
            ci = itemSalvo ? parseFloat(itemSalvo.ci) : (parseFloat(prec.ci) || 0);
        } else {
            ci = parseFloat(prec.ci) || 0;
        }
        if (ci === 0) {
            produtosSemCI++;
            const nomeId = (produto.nome || '').replace(/[^a-z0-9]/gi, '_');
            return `
                <tr id="precif_row_${nomeId}" style="opacity:0.45">
                    <td style="text-align:left; padding-left:15px; font-weight:500; position:sticky; left:0; background:#fff; z-index:1">${_escapeHtml(produto.nome)}<span style="font-size:0.7rem; color:#94a3b8; margin-left:6px">sem CI</span></td>
                    <td colspan="14" style="text-align:center; color:#94a3b8; font-size:0.85rem">CI não configurado — acesse a aba Produtos para definir</td>
                </tr>
            `;
        }

        const fedImpostos = impostosEditaveis[ncm] || {};
        const pisPadrao = parseFloat(aliq.pis ?? fedImpostos.pis ?? document.getElementById('pisPadrao')?.value) || 1.65;
        const cofinsPadrao = parseFloat(aliq.cofins ?? fedImpostos.cofins ?? document.getElementById('cofinsPadrao')?.value) || 7.6;
        const ipiPadrao = parseFloat(aliq.ipi ?? fedImpostos.ipi ?? 0) || 0;

        const icmsPadrao = uf ? buscarAliquotaICMS(uf, tipoPessoa, produto.nome) : (parseFloat(aliq.icmsBase) || 0);

        const taxaProd = (prec.taxa !== null && prec.taxa !== undefined && prec.taxa !== '') ? parseFloat(prec.taxa) : taxaFinal;
        const roiProd = (prec.roi !== null && prec.roi !== undefined && prec.roi !== '') ? parseFloat(prec.roi) : roiFinal;
        const comissaoProd = (prec.comissao !== null && prec.comissao !== undefined && prec.comissao !== '') ? parseFloat(prec.comissao) : comFinal;

        // resolver alíquotas com benefícios/RETID
        const pisEfetivo = resolverAliquota(produto.nome, 'pis', pisPadrao);
        const cofinsEfetivo = resolverAliquota(produto.nome, 'cofins', cofinsPadrao);
        const ipiEfetivo = resolverAliquota(produto.nome, 'ipi', ipiPadrao);
        const icmsEfetivo = resolverAliquota(produto.nome, 'icms', icmsPadrao);

        const retidAtivo = !!retidPorProduto[produto.nome];

        const valorBase = ci * (1 + taxaProd/100) * (1 + roiProd/100);
        const icmsR = valorBase * icmsEfetivo / 100;
        const pisR = valorBase * pisEfetivo / 100;
        const cofinsR = valorBase * cofinsEfetivo / 100;
        const valorImpostos = valorBase + icmsR + pisR + cofinsR;
        const ipiR = valorImpostos * ipiEfetivo / 100;
        const valorTotal = valorImpostos + ipiR;
        const comissaoR = valorBase * comissaoProd / 100;
        const precoFinal = valorTotal + comissaoR;

        const margem = precoFinal > 0 ? ((precoFinal - ci) / precoFinal) * 100 : 0;
        const margemMinima = parseFloat(precificacao[produto.nome]?.margemMinima) || null;
        const abaixo = margemMinima !== null && margem < margemMinima;
        if (abaixo) abaixoCount++;

        totalFaturamento += precoFinal;
        produtosCalculados++;

        itensCalculados.push({ produto: produto.nome, produtoId: produto.id || null, ncm, ci, taxa: taxaProd, roi: roiProd, valorBase, pis: pisEfetivo, pisR, cofins: cofinsEfetivo, cofinsR, icms: icmsEfetivo, icmsR, ipi: ipiEfetivo, ipiR, valorImpostos, comissao: comissaoProd, comissaoR, precoFinal, margem });

        const icmsBg = tipoPessoa === 'PF' ? '#fef2f2' : '#f0fdf4';
        const icmsColor = tipoPessoa === 'PF' ? '#dc2626' : '#16a34a';

        const nomeId = (produto.nome || '').replace(/[^a-z0-9]/gi, '_');
        const safeNome = (produto.nome || '').replace(/'/g, "\\'");
        const ciBadge = exibindoPrecifSalva
            ? '<span style="font-size:0.65rem; background:#dbeafe; color:#1d4ed8; padding:1px 5px; border-radius:10px; margin-left:4px; font-weight:600">🔒 salvo</span>'
            : '';

        // product badge
        let prodBadge = '';
        if (retidAtivo) prodBadge = '<span style="font-size:0.68rem; background:#dbeafe; color:#1d4ed8; padding:1px 6px; border-radius:10px; margin-left:6px; font-weight:700">RETID</span>';
        else if ((beneficiosPorProduto[produto.nome] && beneficiosPorProduto[produto.nome].isentoTotal)) prodBadge = '<span style="font-size:0.68rem; background:#dcfce7; color:#166534; padding:1px 6px; border-radius:10px; margin-left:6px; font-weight:700">ISENTO</span>';

        // RETID cell
        const retidCell = `<td style="text-align:center">
          <label style="display:flex; flex-direction:column; align-items:center; gap:4px; cursor:pointer">
            <input type="checkbox" id="retid_${nomeId}" ${retidAtivo ? 'checked' : ''} onchange="toggleRetid('${safeNome}', this.checked)" style="width:16px; height:16px; cursor:pointer; accent-color:#1e3a5f">
            <span style="font-size:0.65rem; color:${retidAtivo ? '#16a34a' : '#94a3b8'}; font-weight:${retidAtivo ? '700' : '400'}">${retidAtivo ? 'Ativo' : '—'}</span>
          </label>
        </td>`;

        // PIS / COFINS / IPI cells with visual indicators
        const pisZero = pisEfetivo === 0;
        const pisOver = pisEfetivo !== pisPadrao && !pisZero;
        let pisCell = '';
        if (retidAtivo) pisCell = `<td style="text-decoration:line-through; color:#94a3b8; text-align:center; font-size:0.85rem">0% (RETID)</td>`;
        else if (pisZero) pisCell = `<td style="text-align:center; background:#f0f9ff; color:#1e3a5f; font-weight:700">${Number(pisEfetivo).toFixed(2)}%<div style="font-size:0.65rem; background:#dcfce7; color:#166534; padding:1px 5px; border-radius:10px; margin-top:2px">ISENTO</div></td>`;
        else if (pisOver) pisCell = `<td style="text-align:center; background:#f0fdf4; color:#16a34a; font-weight:700">${Number(pisEfetivo).toFixed(2)}%<div style="font-size:0.65rem; background:#dcfce7; color:#166534; padding:1px 5px; border-radius:10px; margin-top:2px">BENEFÍCIO</div></td>`;
        else pisCell = `<td style="text-align:center; color:#dc2626; font-size:0.85rem">${Number(pisEfetivo).toFixed(2)}%</td>`;

        const cofZero = cofinsEfetivo === 0;
        const cofOver = cofinsEfetivo !== cofinsPadrao && !cofZero;
        let cofCell = '';
        if (retidAtivo) cofCell = `<td style="text-decoration:line-through; color:#94a3b8; text-align:center; font-size:0.85rem">0% (RETID)</td>`;
        else if (cofZero) cofCell = `<td style="text-align:center; background:#f0f9ff; color:#1e3a5f; font-weight:700">${Number(cofinsEfetivo).toFixed(2)}%<div style="font-size:0.65rem; background:#dcfce7; color:#166534; padding:1px 5px; border-radius:10px; margin-top:2px">ISENTO</div></td>`;
        else if (cofOver) cofCell = `<td style="text-align:center; background:#f0fdf4; color:#16a34a; font-weight:700">${Number(cofinsEfetivo).toFixed(2)}%<div style="font-size:0.65rem; background:#dcfce7; color:#166534; padding:1px 5px; border-radius:10px; margin-top:2px">BENEFÍCIO</div></td>`;
        else cofCell = `<td style="text-align:center; color:#dc2626; font-size:0.85rem">${Number(cofinsEfetivo).toFixed(2)}%</td>`;

        const ipiZero = ipiEfetivo === 0;
        const ipiOver = ipiEfetivo !== ipiPadrao && !ipiZero;
        let ipiCell = '';
        if (retidAtivo) ipiCell = `<td style="text-decoration:line-through; color:#94a3b8; text-align:center; font-size:0.85rem">0% (RETID)</td>`;
        else if (ipiZero) ipiCell = `<td style="text-align:center; background:#f0f9ff; color:#1e3a5f; font-weight:700">${Number(ipiEfetivo).toFixed(2)}%<div style="font-size:0.65rem; background:#dcfce7; color:#166534; padding:1px 5px; border-radius:10px; margin-top:2px">ISENTO</div></td>`;
        else if (ipiOver) ipiCell = `<td style="text-align:center; background:#f0fdf4; color:#16a34a; font-weight:700">${Number(ipiEfetivo).toFixed(2)}%<div style="font-size:0.65rem; background:#dcfce7; color:#166534; padding:1px 5px; border-radius:10px; margin-top:2px">BENEFÍCIO</div></td>`;
        else ipiCell = `<td style="text-align:center; color:#dc2626; font-size:0.85rem">${Number(ipiEfetivo).toFixed(2)}%</td>`;

        // ICMS cell visuals
        const icmsZero = icmsEfetivo === 0;
        const icmsOver = icmsEfetivo !== icmsPadrao && !icmsZero;
        let icmsCell = '';
        if (icmsZero) icmsCell = `<td style="text-align:center; background:#f0f9ff; color:#1e3a5f; font-weight:700">${Number(icmsEfetivo).toFixed(2)}%<div style="font-size:0.65rem; background:#dcfce7; color:#166534; padding:1px 5px; border-radius:10px; margin-top:2px">ISENTO</div><div style="font-size:0.65rem; font-weight:400; color:#94a3b8">${uf} / ${tipoPessoa}</div></td>`;
        else if (icmsOver) icmsCell = `<td style="text-align:center; background:#f0fdf4; color:#16a34a; font-weight:700">${Number(icmsEfetivo).toFixed(2)}%<div style="font-size:0.65rem; background:#dcfce7; color:#166534; padding:1px 5px; border-radius:10px; margin-top:2px">BENEFÍCIO</div><div style="font-size:0.65rem; font-weight:400; color:#94a3b8">${uf} / ${tipoPessoa}</div></td>`;
        else icmsCell = `<td style="text-align:center; font-weight:700; background:${icmsBg}; color:${icmsColor}">${Number(icmsEfetivo).toFixed(2)}%<div style="font-size:0.65rem; font-weight:400; color:#94a3b8">${uf} / ${tipoPessoa}</div></td>`;

        // row left border if margin below minimum (priority) or RETID active
        const rowStyle = abaixo ? 'border-left:3px solid #dc2626' : (retidAtivo ? 'border-left:3px solid #1e3a5f' : '');
        const delta = abaixo ? ((margemMinima - margem) / 100 * precoFinal) : 0;
        const precoFinalCell = `
            <td style="font-weight:800; color:#c9a227; background:#fffbf0; font-size:1rem">
                ${fmt(precoFinal)}
                ${abaixo ? `<div style="font-size:0.68rem;color:#dc2626;margin-top:2px">↑ R$ ${Math.abs(delta).toFixed(2)} abaixo do mín.</div>` : ''}
            </td>
        `;
        const margemCell = abaixo
            ? `<td style="font-weight:700;color:#dc2626;background:#fef2f2">
                   ${margem.toFixed(1)}%
                   <div style="font-size:0.65rem;background:#fecaca;color:#991b1b;
                               padding:1px 5px;border-radius:10px;margin-top:2px">
                     Mín: ${margemMinima}%
                   </div>
               </td>`
            : `<td style="font-weight:700; color:${corMargem(margem)}">${margem.toFixed(1)}%</td>`;

        return `
            <tr id="precif_row_${nomeId}" style="${rowStyle}">
                <td style="text-align:left; padding-left:15px; font-weight:500; position:sticky; left:0; background:#fff; z-index:1; border-right:1px solid #e2e8f0">${_escapeHtml(produto.nome)} ${prodBadge}</td>
                ${retidCell}
                <td style="font-family:monospace; font-size:0.75rem; color:#64748b; text-align:center">${_escapeHtml(ncm)}</td>
                <td style="font-weight:600; color:#1e3a5f">${fmt(ci)} ${ciBadge}</td>
                <td style="text-align:center; color:#475569">${(Number(taxaProd)).toFixed(2)}%</td>
                <td style="text-align:center; color:#475569">${(Number(roiProd)).toFixed(2)}%</td>
                <td style="font-weight:600; color:#1e3a5f">${fmt(valorBase)}</td>
                ${pisCell}
                ${cofCell}
                ${icmsCell}
                ${ipiCell}
                <td style="font-weight:600; color:#475569">${fmt(valorImpostos)}</td>
                <td style="text-align:center; color:#d97706; font-size:0.85rem">${fmt(comissaoR)}</td>
                ${precoFinalCell}
                ${margemCell}
            </tr>
        `;
    });

    const tbody = document.getElementById('tabelaPrecifClienteBody');
    if (tbody) tbody.innerHTML = rows.join('');

    // Summary com contagem filtrada
    const totalProdutos = produtos.length;
    const filtrados = produtosFiltrados.length;
    let summaryHTML = `
        <div>
            <div style="font-size:0.75rem; color:#64748b; text-transform:uppercase; letter-spacing:0.5px">Exibindo</div>
            <div style="font-size:1.1rem; font-weight:700; color:#1e3a5f">${filtrados} de ${totalProdutos}</div>
        </div>
        <div>
            <div style="font-size:0.75rem; color:#64748b; text-transform:uppercase; letter-spacing:0.5px">Sem CI definido</div>
            <div style="font-size:1.2rem; font-weight:700; color:#94a3b8">${produtosSemCI}</div>
        </div>
        <div>
            <div style="font-size:0.75rem; color:#64748b; text-transform:uppercase; letter-spacing:0.5px">ICMS aplicado</div>
            <div style="font-size:1rem; font-weight:700; color:${tipoPessoa==='PF'?'#dc2626':'#16a34a'}">${uf} / ${tipoPessoa}</div>
        </div>
        <div style="margin-left:auto; text-align:right">
            <div style="font-size:0.75rem; color:#64748b; text-transform:uppercase; letter-spacing:0.5px">Total da tabela</div>
            <div style="font-size:1.3rem; font-weight:800; color:#16a34a">${fmt(totalFaturamento)}</div>
        </div>
    `;
    if (abaixoCount > 0) {
        summaryHTML += `<div style="color:#dc2626;font-weight:600">⚠️ ${abaixoCount} produto(s) abaixo da margem mínima</div>`;
    }
    document.getElementById('precifClienteSummary').innerHTML = summaryHTML;

    document.getElementById('precifClienteResultado').style.display = 'block';
    document.getElementById('precifClienteEmpty').style.display = 'none';

    // guardar última precificação calculada para salvar/gerar proposta
    ultimaPrecificacaoCalculada = {
        clienteId: cliente.id,
        clienteNome: cliente.nome,
        uf,
        tipoPessoa,
        taxa: taxaFinal,
        roi: roiFinal,
        comissao: comFinal,
        filtros: { filtroProduto, filtroNCM, filtroCI },
        dataCriacao: new Date().toISOString(),
        itens: itensCalculados,
        totalFaturamento
    };
    atualizarAvisoCIDivergente(exibindoPrecifSalva ? precifSalvaCarregada : obterUltimaPrecificacaoCliente(cliente.id));
    try { atualizarStatusPropostaNaPrecif(ultimaPrecificacaoCalculada.clienteId); } catch (e) {}
}

function exportarPrecifCliente() {
        const select = document.getElementById('precifClienteSelect');
        const clienteId = select?.value;
        const cliente = (clientes || []).find(c => String(c.id) === String(clienteId));
        if (!cliente) { alert('Selecione um cliente antes de exportar.'); return; }

        const rows = [];
        const cabecalho = ['Produto','RETID','NCM','CI (R$)','Taxa %','ROI %','Valor Base','PIS %','COFINS %','ICMS % (UF/Tipo)','IPI %','c/ Impostos','Comissão R$','Preço Final','Margem%'];
        rows.push([`Cliente: ${cliente.nome}`, `UF: ${cliente.uf || cliente.estado || '—'}`, `Tipo: ${cliente.tipoPessoa || (((cliente.cnpj||'').replace(/\D/g,'')).length===14?'PJ':'PF')}`]);
        rows.push([]);
        rows.push(cabecalho);

        // Ler linhas da tabela DOM (para preservar a ordem e valores formatados)
        const trs = Array.from(document.querySelectorAll('#tabelaPrecifClienteBody tr'));
        trs.forEach(tr => {
            const cols = Array.from(tr.querySelectorAll('td'));
            if (!cols || cols.length < 8) return; // linha de aviso
            const produto = cols[0]?.textContent?.trim() || '';
            const retid = cols[1]?.textContent?.trim() || '';
            const ncm = cols[2]?.textContent?.trim() || '';
            const ci = cols[3]?.textContent?.trim() || '';
            const taxa = cols[4]?.textContent?.trim() || '';
            const roi = cols[5]?.textContent?.trim() || '';
            const vbase = cols[6]?.textContent?.trim() || '';
            const pis = cols[7]?.textContent?.trim() || '';
            const cof = cols[8]?.textContent?.trim() || '';
            const icms = cols[9]?.textContent?.trim() || '';
            const ipi = cols[10]?.textContent?.trim() || '';
            const vimp = cols[11]?.textContent?.trim() || '';
            const vcom = cols[12]?.textContent?.trim() || '';
            const vfinal = cols[13]?.textContent?.trim() || '';
            const margem = cols[14]?.textContent?.trim() || '';
            rows.push([produto,retid,ncm,ci,taxa,roi,vbase,pis,cof,icms,ipi,vimp,vcom,vfinal,margem]);
        });

        const ws = XLSX.utils.aoa_to_sheet(rows);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Preços Cliente');
        const nomeArquivo = `precos_${(cliente.nome||'cliente').replace(/[^a-z0-9\-]/gi,'_')}_${(cliente.uf||'XX')}_${(cliente.tipoPessoa||'PF')}_${new Date().toISOString().slice(0,10)}.xlsx`;
        XLSX.writeFile(wb, nomeArquivo);
}

function limparFiltrosPrecifCliente() {
    try {
        const fp = document.getElementById('precifFiltroProduto'); if (fp) fp.value = '';
        const fn = document.getElementById('precifFiltroNCM'); if (fn) fn.value = '';
        const fci = document.getElementById('precifFiltroCI'); if (fci) fci.value = 'todos';
        calcularPrecificacaoPorCliente();
    } catch (e) { console.warn('limparFiltrosPrecifCliente', e); }
}

function salvarPrecificacaoCliente() {
    if (!ultimaPrecificacaoCalculada) {
        calcularPrecificacaoPorCliente();
        if (!ultimaPrecificacaoCalculada) { alert('Calcule a precificação antes de salvar.'); return; }
    }
    const cId = ultimaPrecificacaoCalculada.clienteId;
    // calcular próxima versão para este cliente
    const versoesDoCliente = (precificacoesCliente || []).filter(p => String(p.clienteId) === String(cId));
    const proximaVersao = versoesDoCliente.length + 1;
    const registro = JSON.parse(JSON.stringify(ultimaPrecificacaoCalculada));
    registro.id = Date.now().toString();
    registro.dataCriacao = new Date().toISOString();
    registro.versao = proximaVersao;
    const descricao = prompt(
        `Versão ${proximaVersao} — adicione uma descrição opcional:\n` +
        `(Ex: "Negociação inicial", "Após desconto", "Preços reajustados")`
    ) || '';
    registro.descricao = descricao;
    registro.status = 'ativa';
    registro.propostaId = null;
    /*
    IMPORTANT: itens[].ci is the CI value frozen at save time.
    Future changes to precificacao[nome].ci do NOT affect this saved pricing.
    To recalculate with updated CI, the user must run a new calculation
    and save again — this will create a new version, not update the old one.
    */
    // Snapshot de RETID e benefícios no momento do salvamento
    try { registro.retidPorProduto = JSON.parse(JSON.stringify(retidPorProduto || {})); } catch (e) { registro.retidPorProduto = {}; }
    registro.beneficioFiscalAtivo = !!beneficioFiscalAtivo;
    try { registro.beneficiosPorProduto = JSON.parse(JSON.stringify(beneficiosPorProduto || {})); } catch (e) { registro.beneficiosPorProduto = {}; }
    // validade da precificação (dias) e data de expiração
    const validadeBruta = parseInt(document.getElementById('precifValidadeDias')?.value || 30, 10);
    const validadeDias = Number.isFinite(validadeBruta) ? Math.max(1, Math.min(365, validadeBruta)) : 30;
    registro.validadeDias = validadeDias;
    registro.dataExpiracao = new Date(Date.now() + validadeDias * 86400000).toISOString();
    precificacoesCliente.push(registro);
    ultimaVersaoSalva = proximaVersao;
    try { estoque.precificacoesCliente = precificacoesCliente; } catch (e) {}
    salvarDados();
    const infoEl = document.getElementById('precifSalvoInfo');
    if (infoEl) infoEl.innerHTML = `<span style="color:#16a34a; font-size:0.82rem">✅ Precificação salva v${registro.versao} em ${new Date(registro.dataCriacao).toLocaleString('pt-BR')}</span> <button onclick="carregarVersaoPrecif('${registro.id}')" class="btn btn-outline btn-sm" style="margin-left:8px">↩ Carregar versão</button> <button onclick="renderizarHistoricoPrecif('${cId}')" class="btn btn-outline btn-sm" style="margin-left:8px">🕘 Histórico</button>`;
    mostrarNotificacao('Precificação salva para o cliente.', 'success');
    try { renderizarHistoricoPrecif(cId); } catch (e) {}
}

function carregarPrecificacaoSalva(clienteId) {
    try {
        const registro = obterUltimaPrecificacaoCliente(clienteId);
        if (!registro) { mostrarNotificacao('Nenhuma precificação salva para este cliente.', 'warning'); return; }
        aplicarEstadoPrecificacaoSalva(registro);
        const infoEl = document.getElementById('precifSalvoInfo'); if (infoEl) infoEl.innerHTML = `<span style="color:#16a34a; font-size:0.82rem">Usando precificação salva v${registro.versao} em ${new Date(registro.dataCriacao).toLocaleString('pt-BR')}</span>`;
        mostrarNotificacao('Precificação carregada.', 'success');
    } catch (e) { console.error('carregarPrecificacaoSalva', e); mostrarNotificacao('Erro ao carregar precificação salva.', 'error'); }
}

function criarPropostaDaPrecificacao() {
    // Avisar se a precificação ativa para este cliente estiver expirada
    try {
        const precifAtiva = (precificacoesCliente || []).find(p =>
            String(p.clienteId) === String(ultimaPrecificacaoCalculada?.clienteId) &&
            p.status === 'ativa'
        );
        if (precifAtiva?.dataExpiracao) {
            const exp = new Date(precifAtiva.dataExpiracao);
            if (exp < new Date()) {
                const ok = confirm(
                    `⚠️ A precificação salva expirou em ${exp.toLocaleDateString('pt-BR')}.\n\n` +
                    `Deseja gerar a proposta mesmo assim com valores desatualizados?\n` +
                    `(Recomendado: cancele e recalcule)`
                );
                if (!ok) return;
            }
        }
    } catch (e) {}

    if (!ultimaPrecificacaoCalculada) {
        alert('Calcule ou carregue uma precificação antes de gerar proposta.');
        return;
    }
    const abaixoMinimo = (ultimaPrecificacaoCalculada?.itens || []).filter(item => {
        const mm = precificacao[item.produto]?.margemMinima;
        return mm && item.margem < mm;
    });
    if (abaixoMinimo.length > 0) {
        const lista = abaixoMinimo
            .map(i => `• ${i.produto}: ${i.margem.toFixed(1)}% (mín: ${precificacao[i.produto].margemMinima}%)`)
            .join('\n');
        if (!confirm(
            `⚠️ ${abaixoMinimo.length} produto(s) abaixo da margem mínima:\n\n${lista}\n\n` +
            `Deseja gerar a proposta mesmo assim?`
        )) return;
    }
    const reg = ultimaPrecificacaoCalculada;
    const clienteObj = (clientes || []).find(c => String(c.id) === String(reg.clienteId));
    const retidMap = (reg && reg.retidPorProduto && Object.keys(reg.retidPorProduto).length) ? reg.retidPorProduto : (retidPorProduto || {});
    const benefMap = (reg && reg.beneficiosPorProduto) ? reg.beneficiosPorProduto : (beneficiosPorProduto || {});
    const benefAtivo = (reg && reg.beneficioFiscalAtivo) || !!beneficioFiscalAtivo;
    const items = (reg.itens || []).map(it => ({ produto: it.produto, produtoId: it.produtoId, quantidade: 1, valorUnitario: Number(it.precoFinal || 0), retid: !!(retidMap[it.produto]) }));
    const valorTotal = items.reduce((s, it) => s + (Number(it.valorUnitario || 0) * (it.quantidade || 1)), 0);
    const nova = {
        id: 'prop-' + Date.now(),
        numero: gerarNumeroProposta(),
        cliente: clienteObj ? clienteObj.nome : (reg.clienteNome || ''),
        representante: clienteObj ? (clienteObj.representante || '') : '',
        data: new Date().toISOString(),
        validade: 30,
        status: 'rascunho',
        itens: items.map(it => ({ produtoId: it.produtoId, produto: it.produto, quantidade: it.quantidade, valorUnitario: it.valorUnitario, retid: !!it.retid })),
        valorTotal
    };
    // adicionar resumo de benefícios na observação da proposta
    try {
        const benefSummary = [];
        if (Object.values(retidMap || {}).some(v => v)) benefSummary.push('RETID aplicado em produtos selecionados');
        if (benefAtivo && Object.keys(benefMap || {}).length) benefSummary.push('Benefícios fiscais aplicados');
        if (benefSummary.length) nova.observacoes = (nova.observacoes || '') + ' | ' + benefSummary.join(', ');
    } catch (e) {}
    propostas.push(nova);
    estoque.propostas = propostas;
    salvarDados();
    renderizarPropostas();
    atualizarKPIsPropostas();
    mostrarNotificacao('Proposta criada a partir da precificação. Abra para editar.', 'success');
    abrirModalProposta(nova.id);
    try {
        if (ultimaVersaoSalva) {
            const precifAtual = (precificacoesCliente || []).find(p => String(p.clienteId) === String(reg.clienteId) && Number(p.versao) === Number(ultimaVersaoSalva));
            if (precifAtual) {
                precifAtual.status = 'convertida';
                precifAtual.propostaId = nova.id;
                try { estoque.precificacoesCliente = precificacoesCliente; } catch (e) {}
                salvarDados();
                try { renderizarHistoricoPrecif(reg.clienteId); } catch(e) {}
            }
        }
    } catch (e) {}
    try { atualizarStatusPropostaNaPrecif(reg.clienteId); } catch (e) {}
}

// ====== Helpers: RETID e Benefícios Fiscais ======

// ===== Histórico de precificações por cliente =====
function renderizarHistoricoPrecif(clienteId) {
        const versoes = (precificacoesCliente || [])
                .filter(p => String(p.clienteId) === String(clienteId))
                .sort((a,b) => new Date(b.dataCriacao) - new Date(a.dataCriacao));
        if (!versoes.length) {
                const painel = document.getElementById('painelHistoricoPrecif'); if (painel) painel.style.display = 'none';
                return;
        }
        const statusColor = { ativa:'#16a34a', arquivada:'#94a3b8', convertida:'#0ea5e9', expirada:'#dc2626' };
        const statusLabel = { ativa:'Ativa', arquivada:'Arquivada', convertida:'Convertida', expirada:'Expirada' };
        const listaEl = document.getElementById('listaHistoricoPrecif');
        if (!listaEl) return;
        const rowsHtml = (versoes || []).map(v => {
            const agora = new Date();
            const exp = v.dataExpiracao ? new Date(v.dataExpiracao) : null;
            const dias = exp ? Math.ceil((exp - agora) / 86400000) : null;
            let expCell = '—';
            if (exp) {
                if (dias < 0) {
                    expCell = `<span style="color:#dc2626;font-size:0.78rem;font-weight:600">Expirada</span>`;
                } else if (dias <= 5) {
                    expCell = `<span style="color:#d97706;font-size:0.78rem;font-weight:600">⚠️ ${dias}d restantes</span>`;
                } else {
                    expCell = `<span style="color:#64748b;font-size:0.78rem">${exp.toLocaleDateString('pt-BR')} (${dias}d)</span>`;
                }
            }
            return `
                        <tr>
                            <td style="font-weight:700; color:#1e3a5f; text-align:center">v${v.versao}</td>
                            <td>${new Date(v.dataCriacao).toLocaleDateString('pt-BR')}</td>
                            <td style="color:#475569">${v.descricao || '<span style="color:#94a3b8">—</span>'}</td>
                            <td style="text-align:center">${v.taxa}%</td>
                            <td style="text-align:center">${v.roi}%</td>
                            <td style="text-align:center">${(v.itens||[]).length}</td>
                            <td style="text-align:center"><span style="background:${statusColor[v.status]}20; color:${statusColor[v.status]}; font-size:0.75rem; font-weight:600; padding:2px 8px; border-radius:20px">${statusLabel[v.status] || v.status}</span></td>
                            <td style="text-align:center">${expCell}</td>
                            <td style="text-align:center">${v.propostaId ? `<span style="color:#0ea5e9; font-size:0.75rem">${(propostas.find(p=>p.id===v.propostaId)?.numero)||'—'}</span>` : '—'}</td>
                            <td style="text-align:center">
                                <button onclick="carregarVersaoPrecif('${v.id}')" class="btn btn-outline btn-sm" title="Carregar esta versão">↩ Carregar</button>
                                <button onclick="arquivarVersaoPrecif('${v.id}')" style="background:none;border:none;cursor:pointer;color:#94a3b8;font-size:0.9rem" title="Arquivar">📦</button>
                                <button onclick="excluirVersaoPrecif('${v.id}')" style="background:none;border:none;cursor:pointer;color:#dc2626;font-size:0.9rem" title="Excluir">🗑️</button>
                            </td>
                        </tr>
                    `;
        }).join('');
        listaEl.innerHTML = `
        <div class="table-wrapper">
            <table class="dashboard-table" style="font-size:0.82rem">
                <thead>
                    <tr>
                        <th>Versão</th>
                        <th>Data</th>
                        <th>Descrição</th>
                        <th>Taxa</th>
                        <th>ROI</th>
                        <th>Produtos</th>
                        <th>Status</th>
                        <th>Expira em</th>
                        <th>Proposta</th>
                        <th>Ações</th>
                    </tr>
                </thead>
                <tbody>
                    ${rowsHtml}
                </tbody>
            </table>
        </div>
        `;
        const painel = document.getElementById('painelHistoricoPrecif'); if (painel) painel.style.display = 'block';
}

function toggleHistoricoPrecif() {
        const lista = document.getElementById('listaHistoricoPrecif');
        if (!lista) return;
        lista.style.display = lista.style.display === 'none' ? 'block' : 'none';
}

function carregarVersaoPrecif(id) {
        const v = (precificacoesCliente || []).find(p => p.id === id);
        if (!v) return;
    aplicarEstadoPrecificacaoSalva(v);
    const infoEl = document.getElementById('precifSalvoInfo'); if (infoEl) infoEl.innerHTML = `<span style="color:#16a34a; font-size:0.82rem">Usando versão v${v.versao} de ${new Date(v.dataCriacao).toLocaleString('pt-BR')}</span>`;
    mostrarNotificacao('Versão carregada.', 'success');
                try { atualizarStatusPropostaNaPrecif(v.clienteId); } catch(e) { console.warn('Erro ao atualizar status da proposta:', e); }
}


function atualizarStatusPropostaNaPrecif(clienteId) {
        try {
                const clienteNome = (clientes || []).find(c => String(c.id) === String(clienteId))?.nome;
        const propostasDoCliente = (propostas || [])
            .filter(p => p.cliente === clienteNome)
            .sort((a,b) => new Date((b.dataCriacao||b.data||0)) - new Date((a.dataCriacao||a.data||0)));
                const container = document.getElementById('precifStatusProposta');
                if (!container) return;
        if (!propostasDoCliente.length) { container.innerHTML = ''; return; }
                const statusColor = {
                        rascunho:  { bg:'#f1f5f9', text:'#64748b', icon:'📝' },
                        enviada:   { bg:'#eff6ff', text:'#1d4ed8', icon:'📤' },
                        aceita:    { bg:'#f0fdf4', text:'#166534', icon:'✅' },
                        recusada:  { bg:'#fef2f2', text:'#dc2626', icon:'❌' },
                        expirada:  { bg:'#fff8f0', text:'#d97706', icon:'⏰' }
                };
                const ultima = propostasDoCliente[0];
                const sc = statusColor[ultima.status] || statusColor.rascunho;
                const fmt = v => 'R$ ' + (v || 0).toLocaleString('pt-BR',{minimumFractionDigits:2});
                container.innerHTML = `
        <div style="background:${sc.bg};border-radius:8px;padding:10px 14px;
                    display:flex;align-items:center;gap:12px;flex-wrap:wrap">
            <span style="font-size:1.1rem">${sc.icon}</span>
            <div>
                <div style="font-size:0.75rem;color:#64748b;text-transform:uppercase;
                            letter-spacing:0.5px">Última proposta</div>
                <div style="font-weight:700;color:${sc.text};font-size:0.9rem">
                    ${ultima.numero} — ${String(ultima.status || '').toUpperCase()}
                </div>
            </div>
            <div>
                <div style="font-size:0.75rem;color:#64748b">Data</div>
                <div style="font-size:0.85rem;color:#1e293b">
                    ${new Date(ultima.dataCriacao||ultima.data||new Date()).toLocaleDateString('pt-BR')}
                </div>
            </div>
            <div>
                <div style="font-size:0.75rem;color:#64748b">Valor</div>
                <div style="font-size:0.85rem;font-weight:600;color:#16a34a">
                    ${fmt(ultima.valorTotal)}
                </div>
            </div>
            ${propostasDoCliente.length > 1
                ? `<div style="font-size:0.78rem;color:#64748b">
                     +${propostasDoCliente.length-1} proposta(s) anterior(es)
                   </div>` : ''}
            <div style="margin-left:auto;display:flex;gap:6px">
                <button class="btn btn-outline btn-sm" onclick="trocarAba('propostas')">
                    Ver propostas →
                </button>
                ${ultima.status === 'aceita'
                    ? `<button class="btn btn-success btn-sm"
                               onclick="converterPropostaEmVenda('${ultima.id}')">
                         📄 Gerar Contrato
                       </button>` : ''}
            </div>
        </div>
        ${propostasDoCliente.length > 1 ? `
            <details style="margin-top:6px">
                <summary style="font-size:0.8rem;color:#64748b;cursor:pointer;padding:4px 0">
                    Ver histórico completo (${propostasDoCliente.length})
                </summary>
                <div style="margin-top:8px;display:flex;flex-direction:column;gap:4px">
                    ${propostasDoCliente.map(p => {
                        const s = statusColor[p.status] || statusColor.rascunho;
                        return `
                            <div style="display:flex;align-items:center;gap:10px;
                                        padding:6px 10px;background:#f8fafc;
                                        border-radius:6px;font-size:0.82rem">
                                <span>${s.icon}</span>
                                <span style="font-weight:600; color:#1e3a5f">${p.numero}</span>
                                <span style="color:${s.text}; font-weight:500">${p.status}</span>
                                <span style="color:#64748b">
                                    ${new Date(p.dataCriacao||p.data||new Date()).toLocaleDateString('pt-BR')}
                                </span>
                                <span style="color:#16a34a;margin-left:auto;font-weight:600">
                                    ${fmt(p.valorTotal)}
                                </span>
                            </div>
                        `;
                    }).join('')}
                </div>
            </details>
        ` : ''}
    `;
        } catch (e) { console.warn('atualizarStatusPropostaNaPrecif', e); }
}
// ====== Helpers: RETID e Benefícios Fiscais ======
function resolverAliquota(nomeProduto, campo, valorPadrao) {
    try {
        // 1) isentoTotal em benefícios
        if (beneficioFiscalAtivo && beneficiosPorProduto && beneficiosPorProduto[nomeProduto] && beneficiosPorProduto[nomeProduto].isentoTotal) return 0;
        // 2) RETID zera apenas impostos federais (não ICMS)
        if (retidPorProduto && retidPorProduto[nomeProduto] && campo !== 'icms') return 0;
        // 3) override explícito por produto
        if (beneficioFiscalAtivo && beneficiosPorProduto && beneficiosPorProduto[nomeProduto]) {
            const override = beneficiosPorProduto[nomeProduto][campo];
            if (override !== null && override !== undefined && override !== '') return Number(override);
        }
    } catch (e) {}
    return Number(valorPadrao || 0);
}

function toggleRetid(nomeProduto, ativo) {
    try {
        retidPorProduto = retidPorProduto || {};
        retidPorProduto[nomeProduto] = !!ativo;
        const nomeId = (nomeProduto || '').replace(/[^a-z0-9]/gi, '_');
        const cb = document.getElementById('retid_' + nomeId); if (cb) cb.checked = !!ativo;
    } catch (e) { console.warn('toggleRetid', e); }
    try { recalcularLinhaPrecifCliente(nomeProduto); } catch (e) { calcularPrecificacaoPorCliente(); }
}

 

function toggleBeneficioFiscal(ativo) {
    beneficioFiscalAtivo = !!ativo;
    const painel = document.getElementById('painelBeneficioFiscal'); if (painel) painel.style.display = beneficioFiscalAtivo ? 'block' : 'none';
    if (beneficioFiscalAtivo) renderizarTabelaBeneficios();
}

function renderizarTabelaBeneficios() {
    const tbody = document.getElementById('tabelaBeneficioFiscalBody');
    if (!tbody) return;
    const produtos = estoque.produtos || [];
    tbody.innerHTML = produtos.map(p => {
        const nome = p.nome || '';
        const nomeId = (nome || '').replace(/[^a-z0-9]/gi, '_');
        const b = beneficiosPorProduto[nome] || { pis:null, cofins:null, ipi:null, icms:null, isentoTotal:false };
        const retChecked = retidPorProduto[nome] ? 'checked' : '';
        const isentoChecked = b.isentoTotal ? 'checked' : '';
        return `
            <tr>
              <td style="padding:6px 10px; font-weight:500">${_escapeHtml(nome)}</td>
              <td style="padding:4px 6px; text-align:center"><input type="number" step="0.01" min="0" max="100" id="bpis_${nomeId}" value="${b.pis ?? ''}" placeholder="padrão" onchange="atualizarBeneficio('${nome}','pis',this.value)" style="width:70px; border:1px solid #fed7aa; border-radius:4px; padding:3px 4px; text-align:center; font-size:0.82rem"></td>
              <td style="padding:4px 6px; text-align:center"><input type="number" step="0.01" min="0" max="100" id="bcofins_${nomeId}" value="${b.cofins ?? ''}" placeholder="padrão" onchange="atualizarBeneficio('${nome}','cofins',this.value)" style="width:70px; border:1px solid #fed7aa; border-radius:4px; padding:3px 4px; text-align:center; font-size:0.82rem"></td>
              <td style="padding:4px 6px; text-align:center"><input type="number" step="0.01" min="0" max="100" id="bipi_${nomeId}" value="${b.ipi ?? ''}" placeholder="padrão" onchange="atualizarBeneficio('${nome}','ipi',this.value)" style="width:70px; border:1px solid #fed7aa; border-radius:4px; padding:3px 4px; text-align:center; font-size:0.82rem"></td>
              <td style="padding:4px 6px; text-align:center"><input type="number" step="0.01" min="0" max="100" id="bicms_${nomeId}" value="${b.icms ?? ''}" placeholder="padrão" onchange="atualizarBeneficio('${nome}','icms',this.value)" style="width:70px; border:1px solid #fed7aa; border-radius:4px; padding:3px 4px; text-align:center; font-size:0.82rem"></td>
              <td style="padding:4px 6px; text-align:center"><input type="checkbox" id="bretid_${nomeId}" ${retChecked} onchange="atualizarBeneficio('${nome}','retid',this.checked)" style="width:16px; height:16px; accent-color:#1e3a5f"></td>
              <td style="padding:4px 6px; text-align:center"><input type="checkbox" id="bisento_${nomeId}" ${isentoChecked} onchange="atualizarBeneficio('${nome}','isentoTotal',this.checked)" style="width:16px; height:16px; accent-color:#dc2626"></td>
            </tr>
        `;
    }).join('');
}

function atualizarBeneficio(nomeProduto, campo, valor) {
    if (!beneficiosPorProduto[nomeProduto]) {
        beneficiosPorProduto[nomeProduto] = { pis:null, cofins:null, ipi:null, icms:null, isentoTotal:false };
    }
    if (campo === 'retid') {
        retidPorProduto[nomeProduto] = !!valor;
        try { recalcularLinhaPrecifCliente(nomeProduto); } catch (e) { calcularPrecificacaoPorCliente(); }
        return;
    }
    if (campo === 'isentoTotal') {
        beneficiosPorProduto[nomeProduto].isentoTotal = !!valor;
        if (valor) {
            ['pis','cofins','ipi','icms'].forEach(t => {
                beneficiosPorProduto[nomeProduto][t] = 0;
                const el = document.getElementById('b' + t + '_' + (nomeProduto || '').replace(/[^a-z0-9]/gi, '_'));
                if (el) { el.value = '0'; el.disabled = true; }
            });
        } else {
            ['pis','cofins','ipi','icms'].forEach(t => {
                beneficiosPorProduto[nomeProduto][t] = null;
                const el = document.getElementById('b' + t + '_' + (nomeProduto || '').replace(/[^a-z0-9]/gi, '_'));
                if (el) { el.value = ''; el.disabled = false; }
            });
        }
        try { recalcularLinhaPrecifCliente(nomeProduto); } catch (e) { calcularPrecificacaoPorCliente(); }
        return;
    }
    // numeric overrides
    beneficiosPorProduto[nomeProduto][campo] = (valor === '' || valor === null) ? null : parseFloat(valor);
    try { recalcularLinhaPrecifCliente(nomeProduto); } catch (e) { calcularPrecificacaoPorCliente(); }
}

function isentarTodosProdutos() {
    const produtos = estoque.produtos || [];
    produtos.forEach(p => {
        beneficiosPorProduto[p.nome] = { pis:0, cofins:0, ipi:0, icms:0, isentoTotal:true };
        retidPorProduto[p.nome] = true;
    });
    renderizarTabelaBeneficios();
}

function limparBeneficios() {
    beneficiosPorProduto = {};
    retidPorProduto = {};
    renderizarTabelaBeneficios();
}

// ── FEDERAL TAXES EDITING (mutable, persisted) ──────────────────
let impostosEditaveis = {};

function inicializarImpostosEditaveis() {
    const defaults = {
        "9301.90.00": { descricao: "Fuzil de assalto IMBEL", pis:1.65, cofins:7.60, ipi:55.00 },
        "9305.91.00": { descricao: "Partes de armas de guerra 93.01", pis:1.65, cofins:7.60, ipi:0.00 },
        "9302.00.00": { descricao: "Revólveres e pistolas", pis:1.65, cofins:7.60, ipi:55.00 },
        "9305.10.00": { descricao: "Partes de revólveres ou pistolas", pis:1.65, cofins:7.60, ipi:29.25 },
        "8211.10.00": { descricao: "Faca", pis:1.65, cofins:7.60, ipi:7.80 },
        "8201.40.00": { descricao: "Machadinha", pis:1.65, cofins:7.60, ipi:0.00 }
    };
    Object.entries(defaults).forEach(([ncm, vals]) => {
        if (!impostosEditaveis[ncm]) impostosEditaveis[ncm] = { ...vals };
    });
}

function renderizarImpostosFederais() {
    inicializarImpostosEditaveis();
    const tbody = document.getElementById('tabelaImpostosFederaisBody');
    if (!tbody) return;
    const produtos = estoque.produtos || [];
    const entries = Object.entries(impostosEditaveis);
    tbody.innerHTML = entries.map(([ncm, imp]) => {
        const vinculados = produtos.filter(p => (p.ncm || detectarNCM(p.nome)) === ncm).length;
        const idSafe = ncm.replace(/\./g,'_');
        return `
            <tr id="row-fed-${idSafe}">
                <td style="text-align:left; padding-left:15px; font-weight:600; font-family:monospace; color:#1e3a5f">${ncm}</td>
                <td style="text-align:left">
                    <input type="text" value="${(imp.descricao||'')}" onchange="editarImpostoFederal('${ncm}','descricao',this.value)" style="width:100%; border:1px solid transparent; border-radius:4px; padding:4px 6px; font-size:0.85rem; background:transparent" onfocus="this.style.borderColor='#1e3a5f'" onblur="this.style.borderColor='transparent'">
                </td>
                <td>
                    <input type="number" step="0.01" min="0" max="100" value="${imp.pis}" onchange="editarImpostoFederal('${ncm}','pis',this.value)" style="width:70px; border:1px solid #e2e8f0; border-radius:4px; padding:4px 6px; text-align:center; font-size:0.85rem">
                </td>
                <td>
                    <input type="number" step="0.01" min="0" max="100" value="${imp.cofins}" onchange="editarImpostoFederal('${ncm}','cofins',this.value)" style="width:70px; border:1px solid #e2e8f0; border-radius:4px; padding:4px 6px; text-align:center; font-size:0.85rem">
                </td>
                <td>
                    <input type="number" step="0.01" min="0" max="100" value="${imp.ipi}" onchange="editarImpostoFederal('${ncm}','ipi',this.value)" style="width:70px; border:1px solid #e2e8f0; border-radius:4px; padding:4px 6px; text-align:center; font-size:0.85rem">
                </td>
                <td style="text-align:center">
                    <span style="background:#e0f2fe; color:#0369a1; font-size:0.75rem; font-weight:600; padding:2px 8px; border-radius:20px">${vinculados} produto(s)</span>
                </td>
                <td style="text-align:center">
                    <button onclick="excluirNCM('${ncm}')" style="background:none; border:none; cursor:pointer; color:#dc2626; font-size:1rem" title="Excluir NCM">🗑️</button>
                </td>
            </tr>
        `;
    }).join('');
}

function editarImpostoFederal(ncm, campo, valor) {
    if (!impostosEditaveis[ncm]) return;
    impostosEditaveis[ncm][campo] = campo === 'descricao' ? String(valor) : (parseFloat(valor) || 0);
    // Propagate to tabelaAliquotas for all products with this NCM
    const produtos = estoque.produtos || [];
    produtos.forEach(p => {
        if ((p.ncm || detectarNCM(p.nome)) === ncm) {
            if (!tabelaAliquotas[p.nome]) tabelaAliquotas[p.nome] = {};
            if (campo !== 'descricao') tabelaAliquotas[p.nome][campo] = impostosEditaveis[ncm][campo];
        }
    });
    salvarDados();
}

function adicionarNCM() {
    const ncm = prompt('Digite o código NCM (ex: 9302.00.00):');
    if (!ncm || ncm.trim() === '') return;
    const ncmLimpo = ncm.trim();
    if (impostosEditaveis[ncmLimpo]) { alert('NCM já existe.'); return; }
    impostosEditaveis[ncmLimpo] = { descricao: 'Novo NCM', pis:1.65, cofins:7.60, ipi:0 };
    renderizarImpostosFederais();
    salvarDados();
}

function excluirNCM(ncm) {
    const produtos = estoque.produtos || [];
    const vinculados = produtos.filter(p => (p.ncm || detectarNCM(p.nome)) === ncm).length;
    const msg = vinculados > 0 ? `Este NCM está vinculado a ${vinculados} produto(s). Deseja excluir mesmo assim?` : `Excluir NCM ${ncm}?`;
    if (!confirm(msg)) return;
    delete impostosEditaveis[ncm];
    renderizarImpostosFederais();
    salvarDados();
}

function exportarImpostosFederais() {
    const rows = Object.entries(impostosEditaveis).map(([ncm, imp]) => ({ 'NCM': ncm, 'Descrição': imp.descricao, 'PIS (%)': imp.pis, 'COFINS (%)': imp.cofins, 'IPI (%)': imp.ipi }));
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Impostos Federais');
    XLSX.writeFile(wb, 'impostos_federais.xlsx');
}

function importarImpostosFederais() { document.getElementById('inputImportarFederais').click(); }

function importarImpostosFederaisArquivo(event) {
    const file = event.target.files[0]; if (!file) return;
    const reader = new FileReader();
    reader.onload = (e) => {
        const wb = XLSX.read(e.target.result, { type: 'binary' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws);
        let importados = 0, atualizados = 0;
        rows.forEach(row => {
            const ncm = (row['NCM'] || row['ncm'] || '').toString().trim();
            if (!ncm) return;
            const novo = { descricao: row['Descrição'] || row['Descricao'] || row['descricao'] || '', pis: parseFloat(row['PIS (%)'] || row['PIS'] || 0), cofins: parseFloat(row['COFINS (%)'] || row['COFINS'] || 0), ipi: parseFloat(row['IPI (%)'] || row['IPI'] || 0) };
            if (impostosEditaveis[ncm]) atualizados++; else importados++;
            impostosEditaveis[ncm] = novo;
        });
        renderizarImpostosFederais();
        alert(`✅ ${importados} NCMs importados, ${atualizados} atualizados.`);
        event.target.value = '';
        salvarDados();
    };
    reader.readAsBinaryString(file);
}

// ── ICMS BY STATE MATRIX (editable) ────────────────────────────
let icmsEditavelPJ = {};
let icmsEditavelPF = {};
const ESTADOS_LISTA = [...ESTADOS_BR];

function inicializarICMSEditavel() {
    Object.entries(ICMS_PJ_POR_NCM).forEach(([ncm, mapa]) => { if (!icmsEditavelPJ[ncm]) icmsEditavelPJ[ncm] = { ...mapa }; });
    Object.entries(ICMS_PF_POR_NCM).forEach(([ncm, mapa]) => { if (!icmsEditavelPF[ncm]) icmsEditavelPF[ncm] = { ...mapa }; });
}

function renderizarICMSPorEstado() {
    inicializarICMSEditavel();
    inicializarImpostosEditaveis();
    const tipo = document.getElementById('filtroICMS_Tipo')?.value || 'PJ';
    const filtroNCM = document.getElementById('filtroICMS_NCM')?.value || '';
    const tabela = tipo === 'PF' ? icmsEditavelPF : icmsEditavelPJ;
    const selectNCM = document.getElementById('filtroICMS_NCM');
    if (selectNCM) {
        const ncms = Object.keys(impostosEditaveis);
        selectNCM.innerHTML = '<option value="">Todos os NCMs</option>' + ncms.map(ncm => `<option value="${ncm}" ${ncm===filtroNCM?'selected':''}>${ncm} — ${impostosEditaveis[ncm]?.descricao || ''}</option>`).join('');
    }
    const ncmsParaExibir = filtroNCM ? [filtroNCM] : Object.keys(tabela);
    document.getElementById('icmsEstadosHeader').innerHTML = `
        <th style="text-align:left; padding-left:10px; min-width:120px; position:sticky; left:0; background:#1e3a5f; z-index:2">NCM</th>
        <th style="text-align:left; min-width:180px; position:sticky; left:120px; background:#1e3a5f; z-index:2">Descrição</th>
        ${ESTADOS_LISTA.map(uf => `<th style="min-width:52px; text-align:center">${uf}</th>`).join('')}
    `;
    document.getElementById('tabelaICMSEstadosBody').innerHTML = ncmsParaExibir.map(ncm => {
        const mapa = tabela[ncm] || {};
        const desc = impostosEditaveis[ncm]?.descricao || ncm;
        return `
            <tr>
                <td style="text-align:left; padding-left:10px; font-weight:600; font-family:monospace; color:#1e3a5f; font-size:0.78rem; position:sticky; left:0; background:#fff; z-index:1; border-right:1px solid #e2e8f0">${ncm}</td>
                <td style="text-align:left; font-size:0.82rem; color:#475569; position:sticky; left:120px; background:#fff; z-index:1; border-right:2px solid #e2e8f0; max-width:180px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap" title="${desc}">${desc}</td>
                ${ESTADOS_LISTA.map(uf => {
                    const val = mapa[uf] !== undefined ? mapa[uf] : '';
                    return `<td style="padding:2px 3px; text-align:center"><input type="number" step="0.5" min="0" max="100" value="${val}" placeholder="—" title="${ncm} / ${uf} / ${tipo}: ${val}%" onchange="editarICMSEstado('${ncm}','${uf}','${tipo}',this.value)" style="width:46px; border:1px solid #e2e8f0; border-radius:4px; padding:3px 2px; text-align:center; font-size:0.8rem; background:${val===''?'#f8fafc':'#fff'}"></td>`;
                }).join('')}
            </tr>
        `;
    }).join('');
}

function editarICMSEstado(ncm, estado, tipoPessoa, valor) {
    const tabela = tipoPessoa === 'PF' ? icmsEditavelPF : icmsEditavelPJ;
    if (!tabela[ncm]) tabela[ncm] = {};
    tabela[ncm][estado] = parseFloat(valor) || 0;
    const idRegra = `predef_${ncm}_${estado}_${tipoPessoa}`;
    const idx = tabelaICMS.findIndex(r => r.id === idRegra);
    if (idx >= 0) {
        tabelaICMS[idx].aliquota = parseFloat(valor) || 0;
    } else {
        tabelaICMS.push({ id: idRegra, ncm, estado, tipoPessoa, categoriaProduto: 'Todos', aliquota: parseFloat(valor) || 0 });
    }
    salvarDados();
}

function exportarICMSEstados() {
    const tipo = document.getElementById('filtroICMS_Tipo')?.value || 'PJ';
    const tabela = tipo === 'PF' ? icmsEditavelPF : icmsEditavelPJ;
    const rows = Object.entries(tabela).map(([ncm, mapa]) => {
        const row = { 'NCM': ncm, 'Descrição': impostosEditaveis[ncm]?.descricao || '', 'Tipo': tipo };
        ESTADOS_LISTA.forEach(uf => { row[uf] = mapa[uf] ?? ''; });
        return row;
    });
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, `ICMS_${tipo}`);
    XLSX.writeFile(wb, `icms_${tipo.toLowerCase()}_${new Date().toISOString().split('T')[0]}.xlsx`);
}

function importarICMSEstados() { document.getElementById('inputImportarICMSEstados').click(); }

function importarICMSEstadosArquivo(event) {
    const file = event.target.files[0]; if (!file) return; const reader = new FileReader();
    reader.onload = (e) => {
        const wb = XLSX.read(e.target.result, { type: 'binary' });
        const ws = wb.Sheets[wb.SheetNames[0]]; const rows = XLSX.utils.sheet_to_json(ws);
        let count = 0; rows.forEach(row => {
            const ncm = (row['NCM'] || '').toString().trim(); const tipo = (row['Tipo'] || 'PJ').toString().trim().toUpperCase();
            if (!ncm) return; const tabela = tipo === 'PF' ? icmsEditavelPF : icmsEditavelPJ; if (!tabela[ncm]) tabela[ncm] = {};
            ESTADOS_LISTA.forEach(uf => { if (row[uf] !== undefined && row[uf] !== '') { tabela[ncm][uf] = parseFloat(row[uf]) || 0; editarICMSEstado(ncm, uf, tipo, tabela[ncm][uf]); } });
            count++;
        });
        renderizarICMSPorEstado(); alert(`✅ ${count} linha(s) de ICMS importadas.`); event.target.value = ''; salvarDados();
    };
    reader.readAsBinaryString(file);
}

function recalcularTodosProdutos() {
    renderizarPrecificacao();
    popularSelectClientesPrecif();
    calcularPrecificacaoPorCliente({ forcarAtual: true });
}

function recalcularTodos() {
    recalcularTodosProdutos();
}

function aplicarMargemGlobal() {
    recalcularTodosProdutos();
}

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
            <td>${_escapeHtml(rule.tipoPessoa)}</td>
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
                        <option value="PJ">PJ</option>
                        <option value="PF">PF</option>
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
                <option value="PJ" ${rule.tipoPessoa === 'PJ' ? 'selected' : ''}>PJ</option>
                <option value="PF" ${rule.tipoPessoa === 'PF' ? 'selected' : ''}>PF</option>
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

function abrirSimuladorICMS() {
    const select = document.getElementById('simProduto');
    if (!select) return;
    const produtos = estoque.produtos || [];

    select.innerHTML = '<option value="">Selecione...</option>'
      + produtos
          .filter(p => (precificacao[p.nome]?.ci || 0) > 0)
          .map(p => `<option value="${_escapeHtml(p.nome)}">${_escapeHtml(p.nome)}</option>`)
          .join('');

    document.getElementById('simResultado').style.display = 'none';
    document.getElementById('modalSimuladorICMS').style.display = 'block';
}

function rodarSimulador() {
    const nome = document.getElementById('simProduto').value;
    const estado = document.getElementById('simEstado').value;
    const tipoPessoa = document.getElementById('simTipoPessoa').value;
    if (!nome || !estado) return;

    const r = calcularPreco(nome, estado, tipoPessoa);
    if (!r) {
        document.getElementById('simResultado').style.display = 'none';
        return;
    }

    const fmt = v => 'R$ ' + Number(v || 0).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    const pct = v => Number(v || 0).toFixed(2) + '%';

    const row = (label, pctTxt, valor, color = '#1e293b', bold = false) => `
      <tr>
        <td style="padding:6px 4px;color:#475569">${label}</td>
        <td style="text-align:center;padding:6px 4px;color:#94a3b8;font-size:0.85rem">
          ${pctTxt}
        </td>
        <td style="text-align:right;padding:6px 4px;font-weight:${bold ? 700 : 500};color:${color}">${valor}</td>
      </tr>`;

    const sep = () => `<tr><td colspan="3" style="border-top:1px solid #e2e8f0;padding:2px 0"></td></tr>`;

    document.getElementById('simDetalhamento').innerHTML = `
      ${row('Custo de Industrialização (CI)', '—', fmt(r.ci))}
      ${row('Taxa (diretoria)', '× ' + r.taxa, '—')}
      ${row('ROI', '× ' + r.roi, '—')}
      ${row('Valor Base (CI × Taxa × ROI)', '—', fmt(r.valorBase), '#1e3a5f', true)}
      ${sep()}
      ${row('ICMS (' + estado + ' / ' + tipoPessoa + ')', pct(r.icms), fmt(r.icmsR), '#dc2626')}
      ${row('PIS', pct(r.pis), fmt(r.pisR), '#dc2626')}
      ${row('COFINS', pct(r.cofins), fmt(r.cofinsR), '#dc2626')}
      ${row('Valor c/ ICMS + PIS + COFINS', '—', fmt(r.valorImpostos), '#1e3a5f', true)}
      ${sep()}
      ${row('IPI', pct(r.ipi), fmt(r.ipiR), '#dc2626')}
      ${row('Valor c/ IPI (Valor Total)', '—', fmt(r.valorTotal), '#1e3a5f', true)}
      ${sep()}
      ${row('Comissão (% s/ Valor Base)', pct(r.comissao), fmt(r.comissaoR), '#d97706')}
    `;

    document.getElementById('simPrecoFinal').textContent = fmt(r.precoFinal);
    document.getElementById('simResultado').style.display = 'block';
}

function sincronizarPrecoNasVendas(nomeProduto) {
    const container = document.getElementById('itensVendaContainer');
    if (!container) return;
    const rows = container.querySelectorAll('.item-venda-row');
    rows.forEach(row => {
        const select = row.querySelector('.item-produto');
        if (!select) return;
        const produtoId = parseInt(select.value);
        const produto = (estoque.produtos || []).find(p => p.id === produtoId);
        if (produto && produto.nome === nomeProduto) {
            const valorInput = row.querySelector('.item-valor');
            const calc = calcularPreco(nomeProduto);
            if (valorInput && calc && (!valorInput.value || valorInput.value.trim() === '')) {
                valorInput.value = Number(calc.precoFinal).toFixed(2).replace('.', ',');
            }
        }
    });
}

function autoPreencherPrecoProduto(selectEl) {
    const row = selectEl.closest('.item-venda-row');
    if (!row) return;
    const valorInput = row.querySelector('.item-valor');
    if (!valorInput) return;
    if (valorInput.value && valorInput.value.trim() !== '') return;

    const produtoId = parseInt(selectEl.value);
    const produto = (estoque.produtos || []).find(p => p.id === produtoId);
    if (!produto) return;

    // Primeiro: tentar obter preço da precificação salva para o cliente selecionado
    const lojaNome = (document.getElementById('lojaVenda')?.value || '').trim();
    let precoSalvo = null;
    try { precoSalvo = obterPrecoFinalSalvoParaClienteProduto(lojaNome, produtoId); } catch (e) { precoSalvo = null; }

    if (precoSalvo !== null && !isNaN(precoSalvo) && Number(precoSalvo) > 0) {
        valorInput.value = Number(precoSalvo).toFixed(2).replace('.', ',');
        valorInput.style.background = '#ecfdf5';
        valorInput.setAttribute('data-autofilled', '1');
        return;
    }

    // Se não houver precificação salva, tentar usar a última precificação calculada para este cliente (se existir)
    try {
        const clienteObj = (clientes || []).find(c => (c.nome||'').toLowerCase() === (lojaNome||'').toLowerCase());
        if (clienteObj && ultimaPrecificacaoCalculada && String(ultimaPrecificacaoCalculada.clienteId) === String(clienteObj.id)) {
            const item = (ultimaPrecificacaoCalculada.itens || []).find(it => Number(it.produtoId) === Number(produtoId) || (it.produto && it.produto.toLowerCase() === (produto.nome||'').toLowerCase()));
            if (item && Number(item.precoFinal) > 0) {
                valorInput.value = Number(item.precoFinal).toFixed(2).replace('.', ',');
                valorInput.style.background = '#ecfdf5';
                valorInput.setAttribute('data-autofilled', '1');
                return;
            }
        }
    } catch (e) {}

    // Caso não haja preço salvo nem versão calculada para este cliente, manter em branco (usuário preencherá manualmente)
}

// Retorna o preço final salvo (number) para um produto em uma precificação do cliente, ou null
function obterPrecoFinalSalvoParaClienteProduto(lojaNome, produtoId) {
    if (!lojaNome) return null;
    const clienteObj = (clientes || []).find(c => (c.nome||'').toLowerCase() === (lojaNome||'').toLowerCase());
    if (!clienteObj) return null;
    let registro = null;
    try { registro = obterUltimaPrecificacaoCliente(clienteObj.id); } catch (e) { registro = null; }
    if (!registro) return null;
    const itens = registro.itens || [];
    const item = itens.find(it => Number(it.produtoId) === Number(produtoId) || (it.produto && (it.produto.toLowerCase() === ((estoque.produtos.find(p=>p.id===produtoId)||{}).nome||'').toLowerCase())));
    if (!item) return null;
    return Number(item.precoFinal || item.preco || item.valorUnitario || 0) || null;
}

// Atualiza preços de todas as linhas do modal de venda de acordo com o cliente preenchido
function atualizarPrecosVendaPorCliente() {
    try {
        const lojaNome = (document.getElementById('lojaVenda')?.value || '').trim();
        if (!lojaNome) return;
        const container = document.getElementById('itensVendaContainer');
        if (!container) return;
        container.querySelectorAll('.item-venda-row').forEach(row => {
            const sel = row.querySelector('.item-produto');
            if (!sel) return;
            autoPreencherPrecoProduto(sel);
        });
    } catch (e) { console.error('atualizarPrecosVendaPorCliente erro', e); }
}

function exportarPrecificacaoExcel() {
    try {
        const federaisEl = document.getElementById('subaba-precif-federais');
        const icmsEl = document.getElementById('subaba-precif-icms');
        const porclienteEl = document.getElementById('subaba-precif-porcliente');
        const comparativoEl = document.getElementById('subaba-precif-comparativo');
        const rastreabilidadeEl = document.getElementById('subaba-precif-rastreabilidade');

        if (federaisEl && federaisEl.style.display === 'block') { exportarImpostosFederais(); return; }
        if (icmsEl && icmsEl.style.display === 'block') { exportarICMSEstados(); return; }
        if (porclienteEl && porclienteEl.style.display === 'block') { exportarPrecifCliente(); return; }
        if (comparativoEl && comparativoEl.style.display === 'block') { exportarComparativo(); return; }
        if (rastreabilidadeEl && rastreabilidadeEl.style.display === 'block') { exportarRastreabilidade(); return; }

        // fallback: exportar impostos federais
        exportarImpostosFederais();
    } catch (e) {
        console.error('exportarPrecificacaoExcel error:', e);
        mostrarNotificacao('Erro ao exportar precificação.', 'error');
    }
}

function imprimirPrecificacao() {
    const tabEl = document.getElementById('tab-precificacao');
    if (!tabEl) return;
    const html = tabEl.querySelector('.table-container')?.innerHTML || '';
    const win = window.open('', '_blank');
    if (!win) { mostrarNotificacao('Pop-up bloqueado pelo navegador.', 'error'); return; }
    win.document.write(`<!DOCTYPE html><html><head><meta charset="UTF-8"><title>Precificação</title>
        <style>body{font-family:Inter,Arial,sans-serif;padding:20px}table{border-collapse:collapse;width:100%}th,td{border:1px solid #ddd;padding:6px 8px;font-size:11px}th{background:#f5f5f5;font-weight:600}input{border:none;background:transparent;text-align:right;font-size:11px}</style>
        <script>window.onload=function(){ setTimeout(function(){ window.print(); },200); }<\/script>
    </head><body><h2>Precificação de Produtos</h2>${html}</body></html>`);
    win.document.close();
}

// ========================================
// MÓDULO: PROPOSTAS COMERCIAIS
// ========================================

let propostas = [];

function gerarNumeroProposta() {
    const existentes = propostas.map(p => {
        const m = (p.numero || '').match(/P-(\d+)/);
        return m ? parseInt(m[1]) : 0;
    });
    const proximo = existentes.length ? Math.max(...existentes) + 1 : 1;
    return 'P-' + String(proximo).padStart(3, '0');
}

function abrirModalProposta(id = null) {
    if (!requireAdminOrNotify()) return;
    const modalEl = document.getElementById('modalProposta');
    modalEl.style.display = 'flex';
    document.getElementById('formProposta').reset();
    document.getElementById('propostaEditId').value = '';
    const container = document.getElementById('itensPropostaContainer');
    if (container) container.innerHTML = '';
    try { popularSelectRepresentantes('propostaRepresentante', true); } catch (e) {}

    try { document.getElementById('propostaData').value = new Date().toISOString().slice(0, 10); } catch (e) {}

    if (!id) {
        document.getElementById('modalPropostaTitulo').textContent = 'Nova Proposta';
        document.getElementById('propostaNumero').value = gerarNumeroProposta();
        document.getElementById('propostaValidade').value = 30;
        document.getElementById('propostaValorTotal').value = '';
        adicionarItemPropostaRow();
        return;
    }

    const proposta = propostas.find(p => p.id === id);
    if (!proposta) { mostrarNotificacao('Proposta não encontrada.', 'error'); return; }

    document.getElementById('modalPropostaTitulo').textContent = 'Editar Proposta ' + proposta.numero;
    document.getElementById('propostaEditId').value = proposta.id;
    document.getElementById('propostaNumero').value = proposta.numero || '';
    document.getElementById('propostaCliente').value = proposta.cliente || '';
    document.getElementById('propostaRepresentante').value = proposta.representante || '';
    document.getElementById('propostaStatus').value = proposta.status || 'rascunho';
    document.getElementById('propostaValidade').value = proposta.validade || 30;
    document.getElementById('propostaObservacoes').value = proposta.observacoes || '';
    try { document.getElementById('propostaData').value = proposta.data ? proposta.data.split('T')[0] : ''; } catch (e) {}

    if (Array.isArray(proposta.itens) && proposta.itens.length > 0) {
        proposta.itens.forEach(it => {
            const preValor = it.valorUnitario ? it.valorUnitario.toLocaleString('pt-BR', { minimumFractionDigits: 2 }) : '';
            adicionarItemPropostaRow({ produtoId: it.produtoId, quantidade: it.quantidade, valorUnit: preValor });
        });
    } else {
        adicionarItemPropostaRow();
    }

    calcularTotalProposta();
}

function adicionarItemPropostaRow(item = null) {
    const container = document.getElementById('itensPropostaContainer');
    if (!container) return;

    const row = document.createElement('div');
    row.className = 'item-venda-row';
    row.style.display = 'flex';
    row.style.gap = '8px';
    row.style.alignItems = 'center';
    row.style.marginBottom = '6px';

    let opcoesHtml = '';
    try { opcoesHtml = construirOpcoesProdutos(); } catch (e) { opcoesHtml = '<option value="">(nenhum produto)</option>'; }

    const preQtd = item ? (item.quantidade || 1) : 1;
    const preValor = item ? (item.valorUnit || '') : '';

    row.innerHTML = `
        <select class="item-produto" onchange="autoPreencherPrecoItemProposta(this); atualizarItemPropostaRow(this)">${opcoesHtml}</select>
        <input type="number" class="item-quantidade" min="1" value="${preQtd}" style="width:90px" onchange="atualizarItemPropostaRow(this)" />
        <input type="text" class="item-valor" placeholder="Valor unit." style="width:140px" oninput="formatarMoeda(this); atualizarItemPropostaRow(this)" value="${preValor}" />
        <div class="item-subtotal" style="min-width:120px">-</div>
        <button type="button" class="btn btn-outline btn-sm" onclick="removerItemPropostaRow(this)">Remover</button>
    `;

    container.appendChild(row);

    if (item && item.produtoId) row.querySelector('.item-produto').value = item.produtoId;

    atualizarItemPropostaRow(row.querySelector('.item-produto'));
}

function autoPreencherPrecoItemProposta(selectEl) {
    const row = selectEl.closest('.item-venda-row');
    if (!row) return;
    const valorInput = row.querySelector('.item-valor');
    if (!valorInput) return;
    if (valorInput.value && valorInput.value.trim() !== '') return;
    const produtoId = parseInt(selectEl.value);
    const produto = (estoque.produtos || []).find(p => p.id === produtoId);
    if (produto) {
        const calc = calcularPreco(produto.nome);
        if (calc && calc.precoFinal > 0) {
            valorInput.value = Number(calc.precoFinal).toFixed(2).replace('.', ',');
        }
    }
}

function atualizarItemPropostaRow(el) {
    const row = el.closest('.item-venda-row');
    if (!row) return;
    const qtdInput = row.querySelector('.item-quantidade');
    const valorInput = row.querySelector('.item-valor');
    const subtotalDiv = row.querySelector('.item-subtotal');

    const qtd = parseInt(qtdInput?.value) || 0;
    const valorStr = (valorInput?.value || '').replace(/\./g, '').replace(',', '.');
    const valor = parseFloat(valorStr) || 0;
    const sub = qtd * valor;

    if (subtotalDiv) subtotalDiv.textContent = sub > 0 ? formatarMoedaValor(sub) : '-';
    calcularTotalProposta();
}

function removerItemPropostaRow(btnEl) {
    const row = btnEl.closest('.item-venda-row');
    if (row) row.remove();
    calcularTotalProposta();
}

function calcularTotalProposta() {
    const container = document.getElementById('itensPropostaContainer');
    if (!container) return;
    let total = 0;
    container.querySelectorAll('.item-venda-row').forEach(row => {
        const qtd = parseInt(row.querySelector('.item-quantidade')?.value) || 0;
        const valorStr = (row.querySelector('.item-valor')?.value || '').replace(/\./g, '').replace(',', '.');
        total += qtd * (parseFloat(valorStr) || 0);
    });
    const el = document.getElementById('propostaValorTotal');
    if (el) el.value = formatarMoedaValor(total);
}

function _coletarItensProposta() {
    const container = document.getElementById('itensPropostaContainer');
    if (!container) return [];
    const itens = [];
    container.querySelectorAll('.item-venda-row').forEach(row => {
        const select = row.querySelector('.item-produto');
        const produtoId = parseInt(select?.value) || 0;
        if (!produtoId) return;
        const produto = (estoque.produtos || []).find(p => p.id === produtoId);
        const quantidade = parseInt(row.querySelector('.item-quantidade')?.value) || 0;
        const valorStr = (row.querySelector('.item-valor')?.value || '').replace(/\./g, '').replace(',', '.');
        const valorUnitario = parseFloat(valorStr) || 0;
        itens.push({
            produtoId,
            produtoNome: produto ? produto.nome : '',
            quantidade,
            valorUnitario,
            valorTotal: quantidade * valorUnitario
        });
    });
    return itens;
}

function salvarProposta(event) {
    if (event) event.preventDefault();
    const editId = document.getElementById('propostaEditId').value;
    const numero = document.getElementById('propostaNumero').value;
    const cliente = document.getElementById('propostaCliente').value.trim();
    const representante = document.getElementById('propostaRepresentante').value;
    const status = document.getElementById('propostaStatus').value;
    const dataStr = document.getElementById('propostaData').value;
    const validade = parseInt(document.getElementById('propostaValidade').value) || 30;
    const observacoes = document.getElementById('propostaObservacoes').value.trim();
    const itens = _coletarItensProposta();

    if (!cliente) { mostrarNotificacao('Informe o cliente.', 'error'); return; }
    if (!representante) { mostrarNotificacao('Selecione o representante.', 'error'); return; }
    if (itens.length === 0) { mostrarNotificacao('Adicione pelo menos um item.', 'error'); return; }

    const valorTotal = itens.reduce((s, i) => s + i.valorTotal, 0);
    const dataISO = dataStr ? new Date(dataStr + 'T00:00:00').toISOString() : new Date().toISOString();
    const dataExp = new Date(dataISO);
    dataExp.setDate(dataExp.getDate() + validade);

    if (editId) {
        const idx = propostas.findIndex(p => p.id === editId);
        if (idx !== -1) {
            propostas[idx].cliente = cliente;
            propostas[idx].representante = representante;
            propostas[idx].status = status;
            propostas[idx].data = dataISO;
            propostas[idx].validade = validade;
            propostas[idx].dataExpiracao = dataExp.toISOString();
            propostas[idx].itens = itens;
            propostas[idx].valorTotal = valorTotal;
            propostas[idx].observacoes = observacoes;
        }
    } else {
        const novaProposta = {
            id: Date.now().toString(),
            numero: numero,
            cliente: cliente,
            representante: representante,
            data: dataISO,
            validade: validade,
            dataExpiracao: dataExp.toISOString(),
            status: status,
            itens: itens,
            valorTotal: valorTotal,
            observacoes: observacoes,
            contratoNumero: null,
            vendaId: null,
            dataCriacao: new Date().toISOString()
        };
        propostas.push(novaProposta);
    }

    estoque.propostas = propostas;
    salvarDados();
    renderizarPropostas();
    atualizarKPIsPropostas();
    fecharModal('modalProposta');
    mostrarNotificacao(editId ? 'Proposta atualizada!' : 'Proposta criada: ' + numero, 'success');
}

function salvarPropostaRascunho() {
    document.getElementById('propostaStatus').value = 'rascunho';
    salvarProposta(null);
}

let _propostaParaConverter = null;

function converterPropostaEmVenda(propostaId) {
    const proposta = (propostas||[]).find(p => p.id === propostaId);
    if (!proposta) { mostrarNotificacao('Proposta não encontrada.', 'error'); return; }

    _propostaParaConverter = proposta;

    // Calculate next contract number (format NNN/AAAA)
    const nextNum = gerarNumeroContrato();

    // Populate modal
    const contratoEl = document.getElementById('confirmarVendaContrato');
    if (contratoEl) contratoEl.value = nextNum;
    const dataEl = document.getElementById('confirmarVendaData');
    if (dataEl) dataEl.value = new Date().toISOString().split('T')[0];

    // Set representante from proposta
    const selRep = document.getElementById('confirmarVendaRepresentante');
    if (selRep && proposta.representante) selRep.value = proposta.representante;

    // Pre-fill observations
    const obsEl = document.getElementById('confirmarVendaObs');
    if (obsEl) obsEl.value = proposta.observacoes
        ? 'Gerado da proposta ' + proposta.numero + '\n' + proposta.observacoes
        : 'Gerado da proposta ' + proposta.numero;

    // Info banner
    const fmt = v => 'R$ ' + (v||0).toLocaleString('pt-BR',{minimumFractionDigits:2});
    const infoEl = document.getElementById('confirmarVendaInfo');
    if (infoEl) infoEl.innerHTML = `
        <div style="display:flex;justify-content:space-between;flex-wrap:wrap;gap:10px">
            <div>
                <div style="font-size:0.75rem;color:#64748b">Proposta</div>
                <div style="font-weight:700;color:#1e3a5f;font-size:1rem">${proposta.numero}</div>
            </div>
            <div>
                <div style="font-size:0.75rem;color:#64748b">Cliente</div>
                <div style="font-weight:600;color:#1e293b">${proposta.cliente}</div>
            </div>
            <div>
                <div style="font-size:0.75rem;color:#64748b">Valor Total</div>
                <div style="font-weight:800;color:#16a34a;font-size:1.1rem">
                    ${fmt(proposta.valorTotal)}
                </div>
            </div>
        </div>
    `;

    // Items table
    const itensEl = document.getElementById('confirmarVendaItens');
    if (itensEl) itensEl.innerHTML = `
        <table style="width:100%;border-collapse:collapse;font-size:0.85rem">
            <thead>
                <tr style="background:#1e3a5f;color:#fff">
                    <th style="padding:8px 12px;text-align:left">Produto</th>
                    <th style="padding:8px 12px;text-align:center">Qtd</th>
                    <th style="padding:8px 12px;text-align:right">Valor Unit.</th>
                    <th style="padding:8px 12px;text-align:right">Total</th>
                </tr>
            </thead>
            <tbody>
                ${(proposta.itens||[]).map((it,i) => `
                    <tr style="background:${i%2===0?'#fff':'#f8fafc'}">
                        <td style="padding:8px 12px;font-weight:500">
                            ${it.produtoNome || it.produto || '—'}
                        </td>
                        <td style="padding:8px 12px;text-align:center">${it.quantidade}</td>
                        <td style="padding:8px 12px;text-align:right">
                            ${fmt(it.valorUnitario || it.valorUnit || 0)}
                        </td>
                        <td style="padding:8px 12px;text-align:right;font-weight:600;color:#16a34a">
                            ${fmt(it.valorTotal || (it.quantidade*(it.valorUnitario||it.valorUnit||0)))}
                        </td>
                    </tr>
                `).join('')}
            </tbody>
        </table>
    `;

    document.getElementById('modalConfirmarVenda').style.display = 'flex';
}

function confirmarConversaoVenda() {
    const proposta = _propostaParaConverter;
    if (!proposta) return;

    const contrato = document.getElementById('confirmarVendaContrato').value?.trim();
    const data     = document.getElementById('confirmarVendaData').value;
    const rep      = document.getElementById('confirmarVendaRepresentante').value;
    const obs      = document.getElementById('confirmarVendaObs').value;

    if (!contrato) {
        mostrarNotificacao('Informe o número do contrato.', 'warning');
        document.getElementById('confirmarVendaContrato').focus();
        return;
    }

    // Check if contract number already exists
    const jaExiste = (estoque.registroVendas||[])
        .some(v => v.contrato === contrato.toString());
    if (jaExiste) {
        if (!confirm(`Contrato nº ${contrato} já existe.\nDeseja usar este número mesmo assim?`))
            return;
    }

    const itensVenda = (proposta.itens||[]).map(it => ({
        produtoId:    it.produtoId   || null,
        produtoNome:  it.produtoNome || it.produto || '',
        quantidade:   it.quantidade  || 0,
        valorUnitario: it.valorUnitario || it.valorUnit || 0,
        valorTotal:   it.valorTotal  || 0,
    }));

    const novaVenda = {
        id:             Date.now(),
        contrato:       contrato.toString(),
        loja:           proposta.cliente,
        representante:  rep || proposta.representante || '',
        data:           data || new Date().toISOString().split('T')[0],
        items:          itensVenda,
        itens:          itensVenda,
        quantidadeTotal: itensVenda.reduce((s,i) => s+i.quantidade, 0),
        valorTotal:     proposta.valorTotal,
        observacoes:    obs,
        propostaId:     proposta.id,
        propostaNumero: proposta.numero,
    };

    estoque.registroVendas.push(novaVenda);
    proposta.status         = 'aceita';
    proposta.contratoNumero = contrato;
    proposta.vendaId        = novaVenda.id;
    estoque.propostas = propostas;

    salvarDados();
    fecharModal('modalConfirmarVenda');
    _propostaParaConverter = null;

    try { renderizarRegistroVendas(); } catch(e) {}
    try { renderizarPropostas(); } catch(e) {}
    try { atualizarKPIsPropostas(); } catch(e) {}
    try { renderizarTabela(); } catch(e) {}
    try { renderizarDashboard(); } catch(e) {}
    try {
        const c = (clientes||[]).find(x => x.nome === proposta.cliente);
        if (c) atualizarStatusPropostaNaPrecif(c.id);
    } catch(e) {}

    mostrarNotificacao(
        `✅ Contrato nº ${contrato} criado! Venda registrada com sucesso.`,
        'success'
    );
    trocarAba('vendas');
}

function excluirProposta(id) {
    if (!requireAdminOrNotify()) return;
    const proposta = propostas.find(p => p.id === id);
    if (!proposta) return;
    if (!confirm('Excluir proposta ' + proposta.numero + '?')) return;

    propostas = propostas.filter(p => p.id !== id);
    estoque.propostas = propostas;
    salvarDados();
    renderizarPropostas();
    atualizarKPIsPropostas();
    mostrarNotificacao('Proposta excluída.', 'success');
}

function recusarProposta(id) {
    if (!requireAdminOrNotify()) return;
    const p = (propostas || []).find(x => x.id === id);
    if (!p) return;
    const motivo = prompt(
        `Recusar proposta ${p.numero}?\n\n` +
        `Informe o motivo da recusa (aparecerá no histórico do cliente):\n` +
        `(Ex: "Preço acima do orçamento", "Concorrente venceu", ` +
        `"Cliente desistiu", "Prazo de entrega")`
    );
    if (motivo === null) return; // user cancelled

    p.status = 'recusada';
    p.motivoRecusa = motivo || 'Sem motivo informado';
    p.dataRecusa = new Date().toISOString();
    try { estoque.propostas = propostas; } catch (e) {}
    salvarDados();
    try { renderizarPropostas(); } catch (e) {}
    try { atualizarKPIsPropostas(); } catch (e) {}
    mostrarNotificacao(`Proposta ${p.numero} marcada como recusada.`, 'warning');
    try { registrarHistorico('proposta_recusada', `Proposta ${p.numero} recusada: ${p.motivoRecusa}`); } catch (e) {}
}

function renderizarPropostas(filtro, statusFiltro) {
    const tbody = document.getElementById('tabelaPropostasBody');
    if (!tbody) return;

    if (filtro === undefined) filtro = (document.getElementById('filtroProposta')?.value || '').trim().toLowerCase();
    else filtro = (filtro || '').trim().toLowerCase();
    if (statusFiltro === undefined) statusFiltro = document.getElementById('filtroStatusProposta')?.value || '';

    const agora = new Date();
    propostas.forEach(p => {
        if (p.status === 'enviada' && p.dataExpiracao && new Date(p.dataExpiracao) < agora) {
            p.status = 'expirada';
        }
    });

    let lista = propostas;
    if (filtro) {
        lista = lista.filter(p =>
            (p.numero || '').toLowerCase().includes(filtro) ||
            (p.cliente || '').toLowerCase().includes(filtro) ||
            (p.representante || '').toLowerCase().includes(filtro)
        );
    }
    if (statusFiltro) {
        lista = lista.filter(p => p.status === statusFiltro);
    }

    const statusLabels = {
        rascunho: { label: 'Rascunho', bg: '#6b7280' },
        enviada:  { label: 'Enviada',  bg: '#0ea5e9' },
        aceita:   { label: 'Aceita',   bg: '#22c55e' },
        recusada: { label: 'Recusada', bg: '#ef4444' },
        expirada: { label: 'Expirada', bg: '#f59e0b' }
    };

    tbody.innerHTML = lista.map(p => {
        const repClass = (p.representante || '').toLowerCase();
        const statusConf = statusLabels[p.status] || statusLabels.rascunho;
        const dataProposta = p.data ? new Date(p.data).toLocaleDateString('pt-BR') : '-';
        const dataValidade = p.dataExpiracao ? new Date(p.dataExpiracao).toLocaleDateString('pt-BR') : '-';
        const validadeExpirada = p.dataExpiracao && new Date(p.dataExpiracao) < agora;
        const contratoDisplay = p.contratoNumero ? p.contratoNumero : '-';

        const podeConverter = p.status === 'rascunho' || p.status === 'enviada';

        const motivoRecusaEsc = p.motivoRecusa ? _escapeHtml(String(p.motivoRecusa)) : '';
        const motivoRecusaSmall = (p.status === 'recusada' && p.motivoRecusa) ? `<div style="font-size:0.75rem;color:#dc2626;margin-top:2px;max-width:200px;overflow:hidden;text-overflow:ellipsis" title="${motivoRecusaEsc}">↳ ${motivoRecusaEsc}</div>` : '';

        return `<tr>
            <td><span style="background:#0ea5e9; color:#fff; padding:2px 10px; border-radius:12px; font-size:0.82rem; font-weight:600;">${_escapeHtml(String(p.numero || ''))}</span></td>
            <td style="text-align:left">${_escapeHtml(p.cliente || '-')}</td>
            <td><span class="badge-rep ${repClass}">${_escapeHtml(p.representante || '-')}</span></td>
            <td style="color:#16a34a; font-weight:600">${formatarMoedaValor(p.valorTotal || 0)}</td>
            <td>${dataProposta}</td>
            <td style="${validadeExpirada ? 'color:#ef4444; font-weight:600' : ''}">${dataValidade}</td>
            <td><span style="background:${statusConf.bg}; color:#fff; padding:2px 10px; border-radius:12px; font-size:0.8rem;">${statusConf.label}</span>${motivoRecusaSmall}</td>
            <td style="font-weight:600">${_escapeHtml(String(contratoDisplay))}</td>
            <td style="font-size:0.78rem;color:#dc2626;max-width:160px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${motivoRecusaEsc}">${p.status === 'recusada' && p.motivoRecusa ? motivoRecusaEsc : '—'}</td>
            <td>
                                <div style="position:relative;display:inline-block">
                                    <button class="btn btn-outline btn-sm" onclick="toggleMenuPDF('${p.id}')" id="btnPDF_${p.id}">
                                        📄 PDF ▾
                                    </button>
                                    <div id="menuPDF_${p.id}" style="display:none;position:absolute;right:0;top:100%;z-index:100;background:#fff;border:1px solid #e2e8f0;border-radius:8px;box-shadow:0 4px 12px rgba(0,0,0,0.12);min-width:180px;padding:4px 0">
                                        <button onclick="gerarPdfProposta('${p.id}','simples')" style="display:block;width:100%;text-align:left;padding:8px 14px;border:none;background:none;cursor:pointer;font-size:0.85rem">
                                            📋 Proposta simples
                                        </button>
                                        <button onclick="gerarPdfProposta('${p.id}','fiscal')" style="display:block;width:100%;text-align:left;padding:8px 14px;border:none;background:none;cursor:pointer;font-size:0.85rem">
                                            🧾 Com detalhamento fiscal
                                        </button>
                                    </div>
                                </div>
                <button class="btn btn-outline btn-sm" data-admin="true" onclick="abrirModalProposta('${p.id}')" title="Editar">✏️</button>
                ${podeConverter ? `<button class="btn btn-success btn-sm" data-admin="true" onclick="converterPropostaEmVenda('${p.id}')" title="Converter em Venda" style="font-size:0.78rem;">🔄</button>` : ''}
                ${p.status === 'enviada' ? `<button class="btn btn-outline btn-sm" data-admin="true" onclick="recusarProposta('${p.id}')" title="Recusar" style="color:#dc2626">❌</button>` : ''}
                <button class="btn btn-outline btn-sm" data-admin="true" onclick="excluirProposta('${p.id}')" title="Excluir" style="color:#ef4444">🗑️</button>
            </td>
        </tr>`;
    }).join('');
}

function toggleMenuPDF(id) {
        document.querySelectorAll('[id^="menuPDF_"]').forEach(m => {
                m.style.display = 'none';
        });
        const menu = document.getElementById('menuPDF_' + id);
        if (menu) menu.style.display = 'block';
        setTimeout(() => {
                document.addEventListener('click', function close(e) {
                        if (menu && !menu.contains(e.target)) {
                                menu.style.display = 'none';
                                document.removeEventListener('click', close);
                        }
                });
        }, 0);
}

function filtrarPropostas(valor) {
    const statusVal = document.getElementById('filtroStatusProposta')?.value || '';
    renderizarPropostas(valor || '', statusVal);
}

function atualizarKPIsPropostas() {
    const abertas = propostas.filter(p => p.status === 'rascunho' || p.status === 'enviada').length;
    const aceitas = propostas.filter(p => p.status === 'aceita').length;
    const recusadas = propostas.filter(p => p.status === 'recusada').length;
    const totalNaoRascunho = propostas.filter(p => p.status !== 'rascunho').length;
    const taxa = totalNaoRascunho > 0 ? ((aceitas / totalNaoRascunho) * 100).toFixed(1) : '0';

    const elAbertas = document.getElementById('kpiPropostasAbertas');
    const elAceitas = document.getElementById('kpiPropostasAceitas');
    const elRecusadas = document.getElementById('kpiPropostasRecusadas');
    const elTaxa = document.getElementById('kpiTaxaConversao');
    if (elAbertas) elAbertas.textContent = abertas;
    if (elAceitas) elAceitas.textContent = aceitas;
    if (elRecusadas) elRecusadas.textContent = recusadas;
    if (elTaxa) elTaxa.textContent = taxa + '%';
}

function preencherDadosCliente(nomeCliente) {
    if (!nomeCliente) return;
    const repSelect = document.getElementById('propostaRepresentante');
    if (!repSelect || repSelect.value) return;
    const cliente = (clientes || []).find(c => (c.nome || '').toLowerCase() === nomeCliente.toLowerCase());
    if (cliente && cliente.representante) {
        repSelect.value = cliente.representante;
    }
}

function imprimirPropostas() {
    const tabEl = document.getElementById('tab-propostas');
    if (!tabEl) return;
    const html = tabEl.querySelector('.table-container')?.innerHTML || '';
    const win = window.open('', '_blank');
    if (!win) { mostrarNotificacao('Pop-up bloqueado pelo navegador.', 'error'); return; }
    win.document.write(`<!DOCTYPE html><html><head><meta charset="UTF-8"><title>Propostas Comerciais</title>
        <style>body{font-family:Inter,Arial,sans-serif;padding:20px}table{border-collapse:collapse;width:100%}th,td{border:1px solid #ddd;padding:6px 8px;font-size:11px}th{background:#f5f5f5;font-weight:600}.badge-rep,.btn{font-size:10px}span[style*="background"]{padding:2px 8px;border-radius:10px;font-size:10px}</style>
        <script>window.onload=function(){ setTimeout(function(){ window.print(); },200); }<\/script>
    </head><body><h2>Propostas Comerciais</h2>${html}</body></html>`);
    win.document.close();
}

// ========================================
// PROGRESS BAR (CLOUD OPERATIONS)
// ========================================

function showProgressBar(text) {
    const el = document.getElementById('progressBar');
    const fill = document.getElementById('progressBarFill');
    const txt = document.getElementById('progressBarText');
    if (!el) return;
    el.style.display = 'block';
    fill.style.width = '10%';
    txt.textContent = text || 'Processando...';
    // Simulate progress
    let pct = 10;
    window._progressInterval = setInterval(() => {
        pct = Math.min(pct + Math.random() * 15, 90);
        fill.style.width = pct + '%';
    }, 300);
}

function hideProgressBar() {
    const el = document.getElementById('progressBar');
    const fill = document.getElementById('progressBarFill');
    if (!el) return;
    clearInterval(window._progressInterval);
    fill.style.width = '100%';
    setTimeout(() => {
        el.style.display = 'none';
        fill.style.width = '0%';
    }, 500);
}

// Override cloud UI functions to use progress bar
const _salvarNoCloudUI_original = salvarNoCloudUI;
salvarNoCloudUI = async function() {
    if (!requireAdminOrNotify()) return false;
    showProgressBar('Salvando no Cloud...');
    try {
        return await _salvarNoCloudUI_original();
    } finally {
        hideProgressBar();
    }
};

const _carregarDoCloudUI_original = carregarDoCloudUI;
carregarDoCloudUI = async function() {
    if (!requireAdminOrNotify()) return false;
    showProgressBar('Carregando do Cloud...');
    try {
        return await _carregarDoCloudUI_original();
    } finally {
        hideProgressBar();
    }
};

// ========================================
// HISTÓRICO DE ALTERAÇÕES (AUDIT LOG)
// ========================================

function getHistorico() {
    try {
        return JSON.parse(localStorage.getItem('estoqueHistorico') || '[]');
    } catch(e) { return []; }
}

function salvarHistorico(hist) {
    // Manter últimos 200 registros
    const trimmed = hist.slice(-200);
    localStorage.setItem('estoqueHistorico', JSON.stringify(trimmed));
}

function registrarHistorico(tipo, descricao) {
    const hist = getHistorico();
    hist.push({
        data: new Date().toISOString(),
        tipo: tipo,
        descricao: descricao
    });
    salvarHistorico(hist);
}

function abrirHistorico() {
    const hist = getHistorico().reverse();
    const container = document.getElementById('historicoConteudo');
    if (!container) return;

    if (hist.length === 0) {
        container.innerHTML = '<p style="text-align:center;color:var(--text-secondary);padding:20px">Nenhuma alteração registrada.</p>';
    } else {
        container.innerHTML = hist.map(h => {
            const dt = new Date(h.data).toLocaleString('pt-BR');
            return `<div class="historico-item">
                <span class="hist-data">${dt}</span>
                <span class="hist-tipo ${h.tipo}">${h.tipo}</span>
                <span class="hist-descricao">${h.descricao}</span>
            </div>`;
        }).join('');
    }

    document.getElementById('modalHistorico').style.display = 'flex';
}

function limparHistorico() {
    if (!confirm('Deseja limpar todo o histórico de alterações?')) return;
    localStorage.removeItem('estoqueHistorico');
    mostrarNotificacao('Histórico limpo com sucesso!', 'success');
}

// ========================================
// VALIDAÇÃO ROBUSTA
// ========================================

function validarContratoUnico(contrato, vendaIdEditando) {
    const existente = estoque.registroVendas.find(v =>
        v.contrato === contrato && v.id !== vendaIdEditando
    );
    return !existente;
}

// ========================================
// RELATÓRIO DE DISTRIBUIÇÃO
// ========================================

function prepararRelatorioDistribuicao() {
    const preview = document.getElementById('relatoriosPreview');
    if (!preview) return;

    const filtroRep = document.getElementById('filtroRelatoriosRep')?.value || '';
    const dataInicio = document.getElementById('filtroRelatoriosDataInicio')?.value || '';
    const dataFim = document.getElementById('filtroRelatoriosDataFim')?.value || '';

    let distribuicoes = [...(estoque.registroDistribuicao || [])];
    let devolucoes = [...(estoque.registroDevolucoes || [])];

    if (filtroRep) {
        distribuicoes = distribuicoes.filter(d => d.representante === filtroRep);
        devolucoes = devolucoes.filter(d => d.origem === filtroRep);
    }

    // Filtrar por data (comparação por DATA YYYY-MM-DD para evitar timezone/formato)
    const aplicarFiltroData = (arr) => {
        if ((!dataInicio || dataInicio === '') && (!dataFim || dataFim === '')) return arr;
        return arr.filter(d => {
            if (!d.data) return false;
            const registroDateStr = parseDateToYYYYMMDD(d.data);
            if (!registroDateStr) return false;
            if (dataInicio && dataInicio !== '' && registroDateStr < dataInicio) return false;
            if (dataFim && dataFim !== '' && registroDateStr > dataFim) return false;
            return true;
        });
    };

    distribuicoes = aplicarFiltroData(distribuicoes);
    devolucoes = aplicarFiltroData(devolucoes);

    // Ordenar por data (mais recente primeiro)
    distribuicoes.sort((a, b) => new Date(b.data) - new Date(a.data));
    devolucoes.sort((a, b) => new Date(b.data) - new Date(a.data));

    // Agrupar por representante (considerar reps presentes em distribuições E devoluções)
    const porRep = {};
    distribuicoes.forEach(d => {
        if (!porRep[d.representante]) porRep[d.representante] = { distrib: [], devol: [] };
        porRep[d.representante].distrib.push(d);
    });
    devolucoes.forEach(d => {
        const rep = d.origem || 'IMBEL';
        if (!porRep[rep]) porRep[rep] = { distrib: [], devol: [] };
        porRep[rep].devol.push(d);
    });

    const container = document.createElement('div');
    container.className = 'report-distribuicao';

    let totalGeral = 0;

    Object.keys(porRep).sort().forEach(rep => {
        const grupo = porRep[rep];
        const titulo = document.createElement('h3');
        titulo.textContent = `Representante: ${rep}`;
        titulo.style.margin = '12px 0 6px 0';
        container.appendChild(titulo);

        const table = document.createElement('table');
        table.style.width = '100%';
        table.style.borderCollapse = 'collapse';
        table.innerHTML = `<thead><tr>
            <th style="padding:6px;border:1px solid #ddd;text-align:left">Produto</th>
            <th style="padding:6px;border:1px solid #ddd;text-align:center">Qtd</th>
            <th style="padding:6px;border:1px solid #ddd;text-align:center">Data</th>
            <th style="padding:6px;border:1px solid #ddd;text-align:left">Obs</th>
        </tr></thead><tbody></tbody>`;

        const tbody = table.querySelector('tbody');
        // combinar listas preservando tipo
        const items = [];
        (grupo.distrib || []).forEach(d => items.push(Object.assign({ tipo: 'dist' }, d)));
        (grupo.devol || []).forEach(d => items.push(Object.assign({ tipo: 'dev' }, d)));

        // ordenar por data desc
        items.sort((a, b) => new Date(b.data) - new Date(a.data));

        let subtotalDist = 0;
        let subtotalDev = 0;
        items.forEach(d => {
            if (d.tipo === 'dist') {
                subtotalDist += d.quantidade || 0;
                totalGeral += d.quantidade || 0;
            } else {
                subtotalDev += d.quantidade || 0;
                totalGeral -= d.quantidade || 0;
            }

            const parsed = parseDateToYYYYMMDD(d.data);
            const dataFmt = parsed ? new Date(parsed).toLocaleDateString('pt-BR') : '-';
            const tr = document.createElement('tr');
            const qtdDisplay = d.tipo === 'dev' ? `-${d.quantidade}` : `${d.quantidade}`;
            const obs = d.tipo === 'dev' ? (d.observacoes ? `Devolução: ${d.observacoes}` : 'Devolução') : (d.observacoes || '-');
            tr.innerHTML = `
                <td style="padding:6px;border:1px solid #ddd">${d.produtoNome}</td>
                <td style="padding:6px;border:1px solid #ddd;text-align:center">${qtdDisplay}</td>
                <td style="padding:6px;border:1px solid #ddd;text-align:center">${dataFmt}</td>
                <td style="padding:6px;border:1px solid #ddd">${obs}</td>`;
            if (d.tipo === 'dev') tr.style.background = '#fff7f7';
            tbody.appendChild(tr);
        });

        const trSubtotal = document.createElement('tr');
        trSubtotal.innerHTML = `<td colspan="1" style="padding:6px;border:1px solid #ddd;text-align:right"><strong>Distribuído: ${subtotalDist} — Devolvido: ${subtotalDev}</strong></td>
            <td style="padding:6px;border:1px solid #ddd;text-align:center"><strong>Saldo: ${subtotalDist - subtotalDev}</strong></td>
            <td colspan="2" style="padding:6px;border:1px solid #ddd"></td>`;
        tbody.appendChild(trSubtotal);
        container.appendChild(table);
    });

    const resumo = document.createElement('div');
    resumo.style.cssText = 'margin:12px 0;font-size:1rem;font-weight:700';
    resumo.textContent = `Total Geral Distribuído: ${totalGeral} unidades`;
    container.insertBefore(resumo, container.firstChild);

    preview.innerHTML = '';
    const wrapper = document.createElement('div');
    wrapper.className = 'report-printable';
    wrapper.appendChild(container);
    preview.appendChild(wrapper);
}

// ========================================
// EXPORTAR PDF (jsPDF)
// ========================================

function exportarRelatorioPDF() {
    if (typeof jspdf === 'undefined' && typeof window.jspdf === 'undefined') {
        mostrarNotificacao('Biblioteca jsPDF não carregada. Tente novamente.', 'error');
        return;
    }

    const { jsPDF } = window.jspdf;
    const tipo = document.getElementById('filtroRelatoriosTipo')?.value || 'inventario';
    const orient = document.getElementById('filtroRelatoriosOrientacao')?.value || 'landscape';

    const doc = new jsPDF({ orientation: orient, unit: 'mm', format: 'a4' });
    doc.setFont('helvetica');

    const dataAgora = new Date().toLocaleString('pt-BR');
    let titulo = 'Relatório';

    if (tipo === 'inventario') {
        titulo = 'Inventário de Produtos';
        // Build table data from estoque
        const headers = [['Produto', 'Disp', 'Venda', 'Saldo']];
        const data = estoque.produtos.map(p => {
            let d = 0, v = 0;
            estoque.representantes.forEach(r => { d += (p.distribuicao[r]||0); v += (p.vendas[r]||0); });
            return [p.nome, d.toString(), v.toString(), (d-v).toString()];
        });
        doc.setFontSize(14);
        doc.text(titulo, 14, 15);
        doc.setFontSize(9);
        doc.text(`Data: ${dataAgora}`, 14, 22);
        doc.autoTable({ head: headers, body: data, startY: 26, styles: { fontSize: 8 } });
    } else if (tipo === 'comissoes') {
        titulo = 'Relatório de Comissões (5%)';
        const vendas = (estoque.registroVendas || []).filter(v => (v.representante||'').toUpperCase() !== 'IMBEL');
        const obterValorVenda = (venda) => {
            if (typeof venda.valorTotal === 'number') return venda.valorTotal;
            if (Array.isArray(venda.items) && venda.items.length > 0) {
                return venda.items.reduce((s, it) => s + (Number(it.valorTotal) || ((Number(it.valorUnitario) || 0) * (Number(it.quantidade) || 0))), 0);
            }
            return ((Number(venda.valorUnitario) || 0) * (Number(venda.quantidade) || 0));
        };
        const normalizarContrato = (valor) => {
            const bruto = (valor ?? '').toString().normalize('NFKC').replace(/[\u200B-\u200D\uFEFF\s]+/g, '');
            const digitos = bruto.replace(/\D+/g, '');
            return digitos ? String(parseInt(digitos, 10)) : bruto.toUpperCase();
        };
        const contratosMap = new Map();
        vendas.forEach(v => {
            const contratoKey = normalizarContrato(v.contrato);
            if (!contratoKey) return;
            const mapKey = `${v.representante || ''}||${contratoKey}`;
            const dataNorm = parseDateToYYYYMMDD(v.data);
            const atual = contratosMap.get(mapKey) || { representante: v.representante || '', contrato: contratoKey, loja: v.loja || '', valor: 0, dataMin: null, dataMax: null };
            atual.valor += obterValorVenda(v);
            if (!atual.loja && v.loja) atual.loja = v.loja;
            if (dataNorm) {
                if (!atual.dataMin || dataNorm < atual.dataMin) atual.dataMin = dataNorm;
                if (!atual.dataMax || dataNorm > atual.dataMax) atual.dataMax = dataNorm;
            }
            contratosMap.set(mapKey, atual);
        });
        const contratos = Array.from(contratosMap.values());
        const headers = [['Rep', 'Contrato', 'Cliente', 'Data', 'Valor', 'Comissão 5%']];
        const data = contratos.map(c => {
            const valor = c.valor || 0;
            const dataTexto = c.dataMin
                ? (c.dataMax && c.dataMax !== c.dataMin
                    ? `${new Date(c.dataMin + 'T00:00:00').toLocaleDateString('pt-BR')} até ${new Date(c.dataMax + 'T00:00:00').toLocaleDateString('pt-BR')}`
                    : new Date(c.dataMin + 'T00:00:00').toLocaleDateString('pt-BR'))
                : '-';
            return [c.representante, c.contrato, c.loja, dataTexto, formatarMoedaValor(valor), formatarMoedaValor(Math.round(valor*0.05*100)/100)];
        });
        doc.setFontSize(14);
        doc.text(titulo, 14, 15);
        doc.setFontSize(9);
        doc.text(`Data: ${dataAgora}`, 14, 22);
        doc.autoTable({ head: headers, body: data, startY: 26, styles: { fontSize: 8 } });
    } else if (tipo === 'distribuicao') {
        titulo = 'Relatório de Distribuição';
        const headers = [['Rep', 'Produto', 'Qtd', 'Data', 'Obs']];
        // combinar distribuições e devoluções (devoluções aparecem como quantidade negativa)
        const distrib = (estoque.registroDistribuicao || []).map(d => ({ rep: d.representante, produto: d.produtoNome, qtd: Number(d.quantidade||0), data: d.data || '', obs: d.observacoes || '', tipo: 'D' }));
        const devol = (estoque.registroDevolucoes || []).map(d => ({ rep: d.origem || '', produto: d.produtoNome, qtd: -(Number(d.quantidade||0)), data: d.data || '', obs: d.observacoes || '', tipo: 'R' }));
        const combined = [...distrib, ...devol].sort((a,b) => new Date(b.data||0) - new Date(a.data||0));
        const data = combined.map(d => [d.rep, d.produto, d.qtd.toString(), d.data ? new Date(d.data+'T00:00:00').toLocaleDateString('pt-BR') : '-', (d.tipo==='R'?'Devolução: ':'') + (d.obs || '-')]);
        doc.setFontSize(14);
        doc.text(titulo, 14, 15);
        doc.setFontSize(9);
        doc.text(`Data: ${dataAgora}`, 14, 22);
        doc.autoTable({ head: headers, body: data, startY: 26, styles: { fontSize: 8 } });
    }

    doc.save(`${titulo.replace(/\s+/g, '_')}_${new Date().toISOString().slice(0,10)}.pdf`);
    mostrarNotificacao('PDF exportado com sucesso!', 'success');
}

// ========================================
// VISUALIZAR RELATÓRIO (atualizado com distribuição)
// ========================================

// Override visualizarRelatorioSelecionado
const _visualizarRelatorioOriginal = visualizarRelatorioSelecionado;
visualizarRelatorioSelecionado = function() {
    const tipo = document.getElementById('filtroRelatoriosTipo')?.value || 'inventario';
    if (tipo === 'distribuicao') {
        prepararRelatorioDistribuicao();
    } else if (tipo === 'comissoes') {
        prepararRelatorioComissoes();
    } else {
        prepararRelatorioInventario();
    }
};

// ========================================
// HOOK: REGISTRAR HISTÓRICO EM OPERAÇÕES
// ========================================

// Wrap salvarVendaDetalhada
const _salvarVendaDetalhadaOriginal = salvarVendaDetalhada;
// Note: can't easily wrap form submit handlers, so we hook into salvarDados
const _salvarDadosOriginal = salvarDados;

// Hook renderizarDashboard to include charts
const _renderizarDashboardOriginal = renderizarDashboard;
renderizarDashboard = function() {
    _renderizarDashboardOriginal();
    try { renderizarGraficos(); } catch(e) { console.warn('Erro renderizando gráficos:', e); }
};

// Hook renderizarRegistroVendas to include pagination and date filters
const _renderizarRegistroVendasOriginal = renderizarRegistroVendas;
renderizarRegistroVendas = function() {
    // Delega para a implementação principal corrigida (com agrupamento por contrato e colunas alinhadas)
    _renderizarRegistroVendasOriginal();

    // Limpar paginação antiga para não confundir a interface
    const pag = document.getElementById('paginacaoVendas');
    if (pag) pag.innerHTML = '';
};

// Hook renderizarRegistroDistribuicao to include pagination, date filters and sorting
const _renderizarRegistroDistribuicaoOriginal = renderizarRegistroDistribuicao;
renderizarRegistroDistribuicao = function() {
    const tbody = document.getElementById('tabelaRegistroDistribuicaoBody');
    if (!tbody) return;

    const filtroRep = document.getElementById('filtroDistribuicaoRep')?.value || '';
    const filtroProduto = document.getElementById('filtroDistribuicaoProduto')?.value || '';
    const dataInicio = document.getElementById('filtroDistribuicaoDataInicio')?.value || '';
    const dataFim = document.getElementById('filtroDistribuicaoDataFim')?.value || '';

    // Combina distribuições e devoluções para a visualização; devoluções serão marcadas com tipo 'dev'
    const distrib = (estoque.registroDistribuicao || []).map(d => Object.assign({}, d, { tipo: 'dist' }));
    const devol = (estoque.registroDevolucoes || []).map(d => Object.assign({}, d, { tipo: 'dev', representante: d.origem, produtoNome: d.produtoNome, produtoId: d.produtoId }));

    let combinado = [...distrib, ...devol];

    if (filtroRep) combinado = combinado.filter(d => (d.representante || '') === filtroRep);
    if (filtroProduto) combinado = combinado.filter(d => (d.produtoId || 0) === parseInt(filtroProduto));

    // Filtro por data com boundaries
    if (dataInicio || dataFim) {
        const start = dataInicio ? new Date(dataInicio + 'T00:00:00').getTime() : null;
        const end = dataFim ? new Date(dataFim + 'T23:59:59').getTime() : null;
        combinado = combinado.filter(d => {
            if (!d.data) return false;
            const t = new Date(d.data + (d.data.length===10 ? 'T00:00:00' : '')).getTime();
            if (start && t < start) return false;
            if (end && t > end) return false;
            return true;
        });
    }

    // Ordenação com suporte a campos e tipos
    const campo = _ordenDistribuicao.campo;
    const dir = _ordenDistribuicao.direcao === 'asc' ? 1 : -1;
    combinado.sort((a, b) => {
        let va, vb;
        if (campo === 'representante') { va = (a.representante||'').toString(); vb = (b.representante||'').toString(); }
        else if (campo === 'produtoNome') { va = (a.produtoNome||'').toString(); vb = (b.produtoNome||'').toString(); }
        else if (campo === 'quantidade') { va = (a.tipo === 'dev' ? -(a.quantidade||0) : (a.quantidade||0)); vb = (b.tipo === 'dev' ? -(b.quantidade||0) : (b.quantidade||0)); }
        else if (campo === 'data') { va = a.data || ''; vb = b.data || ''; }
        else { va = a.data || ''; vb = b.data || ''; }
        if (va < vb) return -1 * dir;
        if (va > vb) return 1 * dir;
        return 0;
    });

    tbody.innerHTML = '';

    if (combinado.length === 0) {
        tbody.innerHTML = `<tr><td colspan="7" class="empty-state"><div class="empty-icon">🚚</div><div class="empty-text">Nenhuma distribuição ou devolução registrada</div><div class="empty-hint">Clique em "Nova Distribuição" ou registre uma devolução</div></td></tr>`;
        atualizarTotaisDistribuicao(0);
        renderizarPaginacao('paginacaoDistribuicao', 1, 0, _itensPorPaginaDistribuicao, 'mudarPaginaDistribuicao', 'mudarItensPaginaDistribuicao');
        return;
    }

    const totalLinhas = combinado.length;
    const totalPaginas = Math.max(1, Math.ceil(totalLinhas / _itensPorPaginaDistribuicao));
    if (_paginaDistribuicao > totalPaginas) _paginaDistribuicao = totalPaginas;
    const inicio = (_paginaDistribuicao - 1) * _itensPorPaginaDistribuicao;
    const pagina = combinado.slice(inicio, inicio + _itensPorPaginaDistribuicao);

    let totalQtd = 0;
    combinado.forEach(d => { totalQtd += (d.tipo === 'dev' ? -(d.quantidade||0) : (d.quantidade||0)); });

    let numero = totalLinhas - inicio;

    pagina.forEach(item => {
        const repClass = (item.representante || '').toLowerCase();
        const dataFormatada = item.data ? new Date(item.data + 'T00:00:00').toLocaleDateString('pt-BR') : '-';
        const tr = document.createElement('tr');
        const qtdDisplay = item.tipo === 'dev' ? `-${(item.quantidade||0)}` : `${(item.quantidade||0)}`;
        const obs = item.tipo === 'dev' ? (item.observacoes ? `Devolução: ${item.observacoes}` : 'Devolução') : (item.observacoes || '-');
        const acaoBtn = item.tipo === 'dev'
            ? `<button class="btn-action btn-delete" onclick="excluirDevolucao(${item.id})" title="Excluir devolução">🗑</button>`
            : `<button class="btn-action btn-delete" onclick="excluirDistribuicao(${item.id})" title="Excluir">🗑</button>`;

        tr.innerHTML = `
            <td class="col-contrato">${numero--}</td>
            <td class="col-loja"><span class="badge-rep ${repClass}">${item.representante || '-'}</span></td>
            <td class="col-produto-venda" title="${item.produtoNome || '-'}">${item.produtoNome || '-'}</td>
            <td class="col-qtd">${qtdDisplay}</td>
            <td>${dataFormatada}</td>
            <td class="col-obs" title="${obs}">${obs}</td>
            <td class="col-acoes">${acaoBtn}</td>`;
        if (item.tipo === 'dev') tr.style.background = '#fff7f7';
        tbody.appendChild(tr);
    });

    atualizarTotaisDistribuicao(totalQtd);
    renderizarPaginacao('paginacaoDistribuicao', _paginaDistribuicao, totalLinhas, _itensPorPaginaDistribuicao, 'mudarPaginaDistribuicao', 'mudarItensPaginaDistribuicao');
};

// Hook limparFiltrosVendas to clear date fields
const _limparFiltrosVendasOriginal = limparFiltrosVendas;
limparFiltrosVendas = function() {
    const filtroRep = document.getElementById('filtroRepresentante');
    const filtroProduto = document.getElementById('filtroProduto');
    const dataInicio = document.getElementById('filtroVendasDataInicio');
    const dataFim = document.getElementById('filtroVendasDataFim');

    if (filtroRep) filtroRep.value = '';
    if (filtroProduto) filtroProduto.value = '';
    if (dataInicio) dataInicio.value = '';
    if (dataFim) dataFim.value = '';

    _paginaVendas = 1;
    renderizarRegistroVendas();
};

// Hook limparFiltrosDistribuicao to clear date fields
const _limparFiltrosDistribuicaoOriginal = limparFiltrosDistribuicao;
limparFiltrosDistribuicao = function() {
    const filtroRep = document.getElementById('filtroDistribuicaoRep');
    const filtroProduto = document.getElementById('filtroDistribuicaoProduto');
    const dataInicio = document.getElementById('filtroDistribuicaoDataInicio');
    const dataFim = document.getElementById('filtroDistribuicaoDataFim');

    if (filtroRep) filtroRep.value = '';
    if (filtroProduto) filtroProduto.value = '';
    if (dataInicio) dataInicio.value = '';
    if (dataFim) dataFim.value = '';

    _paginaDistribuicao = 1;
    renderizarRegistroDistribuicao();
};

// Hook inicializar to check for low stock and integrate audit log
const _inicializarOriginal = inicializar;
inicializar = async function() {
    await _inicializarOriginal();
    try { verificarAlertasEstoque(); } catch(e) {}
};

// Hook salvarDados to check low stock and register audit
const _salvarDadosHook = salvarDados;
// We can't easily re-declare salvarDados since it's used everywhere,
// but we do check low stock after renders
const _renderizarTabelaOriginal = renderizarTabela;
renderizarTabela = function() {
    _renderizarTabelaOriginal();
    try { verificarAlertasEstoque(); } catch(e) {}
};

// Hooks for audit log on key operations
const _salvarNovaDistribuicaoOriginal = salvarNovaDistribuicao;
salvarNovaDistribuicao = function(event) {
    _salvarNovaDistribuicaoOriginal(event);
    try {
        const rep = document.getElementById('representanteDistDet')?.value || '';
        const container = document.getElementById('itensDistribuicaoContainer');
        if (container) {
            const rows = Array.from(container.querySelectorAll('.item-dist-row'));
            const resumo = rows.map(r => {
                const prod = r.querySelector('.item-produto-dist')?.selectedOptions[0]?.text || '';
                const qtd = r.querySelector('.item-quantidade-dist')?.value || '';
                return `${qtd}x ${prod}`;
            }).join(', ');
            registrarHistorico('distribuicao', `${resumo} → ${rep}`);
        } else {
            const prod = document.getElementById('produtoDistDet')?.selectedOptions[0]?.text || '';
            const qtd = document.getElementById('quantidadeDistDet')?.value || '';
            registrarHistorico('distribuicao', `${qtd}x ${prod} → ${rep}`);
        }
    } catch(e) {}
};

const _salvarEntradaEstoqueOriginal = salvarEntradaEstoque;
salvarEntradaEstoque = function(event) {
    const prodEl = document.getElementById('produtoEntrada');
    const qtdEl = document.getElementById('quantidadeEntrada');
    const prodNome = prodEl?.selectedOptions[0]?.text || '';
    const qtd = qtdEl?.value || '';
    _salvarEntradaEstoqueOriginal(event);
    try { registrarHistorico('entrada', `+${qtd} ${prodNome} (IMBEL)`); } catch(e) {}
    try { verificarAlertasEstoque(); } catch(e) {}
};

const _excluirVendaOriginal = excluirVenda;
excluirVenda = function(vendaId) {
    const venda = estoque.registroVendas.find(v => v.id === vendaId);
    _excluirVendaOriginal(vendaId);
    if (venda) {
        try { registrarHistorico('exclusao', `Venda CTR ${venda.contrato} excluída`); } catch(e) {}
    }
};

const _excluirDistribuicaoOriginal = excluirDistribuicao;
excluirDistribuicao = function(distId) {
    const dist = estoque.registroDistribuicao.find(d => d.id === distId);
    _excluirDistribuicaoOriginal(distId);
    if (dist) {
        try { registrarHistorico('exclusao', `Distribuição ${dist.produtoNome} x${dist.quantidade} (${dist.representante}) excluída`); } catch(e) {}
    }
};

// ---------------------------
// Autenticação (Firebase Auth - client)
// ---------------------------

// Realiza login com email/senha usando Firebase Auth (compat)
async function signIn() {
    const emailEl = document.getElementById('authEmail');
    const passEl = document.getElementById('authPassword');
    if (!emailEl || !passEl) return;
    const email = emailEl.value.trim();
    const password = passEl.value;
    try {
        await firebase.auth().signInWithEmailAndPassword(email, password);
        mostrarNotificacao('Login efetuado', 'success');
    } catch (err) {
        console.error('Erro signIn', err);
        mostrarNotificacao('Falha no login: ' + (err.message || err), 'error');
    }
}

// Desloga o usuário
async function signOut() {
    try {
        await firebase.auth().signOut();
        mostrarNotificacao('Sessão encerrada', 'info');
    } catch (err) {
        console.error('Erro signOut', err);
        mostrarNotificacao('Erro ao sair: ' + (err.message || err), 'error');
    }
}

// Atualiza UI conforme estado de autenticação (guardado caso Firebase esteja bloqueado)
if (window.firebase && firebase.auth) {
    try {
        firebase.auth().onAuthStateChanged(async function(user) {
            const formEl = document.getElementById('authPanelForm');
            const signedEl = document.getElementById('authSignedIn');
            const userDisplay = document.getElementById('authUserDisplay');
            const loggedEmailEl = document.getElementById('loggedAccountEmail');
            const loggedBadgeEl = document.getElementById('loggedAccountBadge');
            if (user) {
                if (formEl) formEl.style.display = 'none';
                if (signedEl) signedEl.style.display = 'flex';
                if (userDisplay) userDisplay.textContent = user.email || user.uid;

                // Verifica claims para habilitar controles de admin
                let isAdmin = false;
                try {
                    const idt = await user.getIdTokenResult();
                    isAdmin = !!idt.claims && !!idt.claims.admin;
                } catch (e) { /* ignore */ }

                // Fallback por email (temporário) — remove se preferir depender apenas da claim
                if (!isAdmin && user.email === 'joffre.ribeiro@gmail.com') isAdmin = true;

                if (isAdmin) {
                    document.body.classList.add('is-admin');
                    if (loggedBadgeEl) loggedBadgeEl.style.display = 'inline-block';
                } else {
                    document.body.classList.remove('is-admin');
                    if (loggedBadgeEl) loggedBadgeEl.style.display = 'none';
                }
                // Toggle UI controls marked as admin-only
                try {
                    const adminControls = document.querySelectorAll('[data-admin="true"]');
                    adminControls.forEach(el => {
                        if (isAdmin) {
                            el.removeAttribute('disabled');
                            el.style.pointerEvents = '';
                            el.style.opacity = '';
                        } else {
                            el.setAttribute('disabled','disabled');
                            el.style.pointerEvents = 'none';
                            el.style.opacity = '0.55';
                        }
                    });
                } catch(e) {}
                if (loggedEmailEl) loggedEmailEl.textContent = user.email || '';

                // Persist a short record of current user in localStorage for quick reference
                try {
                    localStorage.setItem('currentUser', JSON.stringify({ email: user.email || null, uid: user.uid || null, isAdmin }));
                } catch(e) {}

                // Após autenticar, atualizar status do cloud e tentar auto-load 1x por usuário
                try {
                    if (window.firestoreDB) {
                        try {
                            const doc = await window.firestoreDB.collection('app_data').doc('latest').get();
                            if (doc && doc.exists) {
                                const data = doc.data();
                                const updatedAt = data && data.updatedAt ? data.updatedAt.toDate() : null;
                                updateFirestoreStatus(true, updatedAt, 'Cloud: pronto');
                            } else {
                                updateFirestoreStatus(true, null, 'Cloud: pronto (sem backup)');
                            }
                        } catch (e) {
                            updateFirestoreStatus(true, null, 'Cloud: sem permissão de leitura');
                        }

                        // Auto-load somente uma vez por usuário autenticado
                        if (window.__cloudAutoLoadDoneForUid !== user.uid) {
                            const autoLoaded = await carregarDoCloudAuto();
                            window.__cloudAutoLoadDoneForUid = user.uid;
                            if (autoLoaded) {
                                try { mostrarNotificacao('Dados carregados automaticamente do Cloud (remoto mais recente).', 'success'); } catch (e) {}
                            }
                        }
                    }
                } catch (e) { /* ignore */ }

            } else {
                if (formEl) formEl.style.display = 'flex';
                if (signedEl) signedEl.style.display = 'none';
                if (userDisplay) userDisplay.textContent = '';
                document.body.classList.remove('is-admin');
                if (loggedEmailEl) loggedEmailEl.textContent = '';
                if (loggedBadgeEl) loggedBadgeEl.style.display = 'none';
                try { localStorage.removeItem('currentUser'); } catch(e) {}
                try { window.__cloudAutoLoadDoneForUid = null; } catch (e) {}
                try { updateFirestoreStatus(true, null, 'Cloud: aguardando login'); } catch (e) {}
            }
        });
    } catch (err) {
        console.warn('firebase.auth() hook failed:', err);
        window.__showRuntimeErrorOverlay && window.__showRuntimeErrorOverlay('firebase.auth() hook failed: ' + (err && err.message ? err.message : err));
    }
} else {
    console.warn('Firebase SDK não disponível; pulando onAuthStateChanged.');
}

// Helper to check admin state synchronously from localStorage/cache
function isCurrentUserAdmin() {
    try {
        const raw = localStorage.getItem('currentUser');
        if (!raw) return false;
        const parsed = JSON.parse(raw);
        return !!parsed && !!parsed.isAdmin;
    } catch(e) { return false; }
}

// Verifica se o usuário atual é admin; caso contrário exibe notificação e mostra o painel de login
function requireAdminOrNotify() {
    try {
        if (isCurrentUserAdmin()) return true;
        mostrarNotificacao('Ação restrita: somente administradores podem realizar esta operação.', 'warning');
        const formEl = document.getElementById('authPanelForm');
        if (formEl) formEl.style.display = 'flex';
    } catch(e) { /* ignore */ }
    return false;
}

// Forçar chamada inicial para ajustar UI caso o listener já tenha ocorrido
try { if (firebase && firebase.auth) firebase.auth().currentUser; } catch(e) {}