// ========================================
// SISTEMA DE CONTROLE DE ESTOQUE
// Material Bélico - v2.0
// ========================================

// Estrutura de dados principal
let estoque = {
    produtos: [],
    representantes: ['KOLTE', 'ISA', 'LC', 'ADES', 'FL', 'IMBEL'],
    registroVendas: [],
    registroDistribuicao: [],
    controleEnvio: {},
    auditoriaVendas: [],
    fechamentosComissoes: []
};

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

async function inicializar() {
    carregarDados();

    // Sync automático com cloud ocorre após autenticação (onAuthStateChanged)

    renderizarTabela();
    renderizarDashboard();
    renderizarRegistroVendas();
    renderizarRegistroDistribuicao();
    renderizarControleEnvio();
    atualizarSelectsProdutos();
    atualizarSelectsRelatorios();
    atualizarEstatisticas();
    atualizarData();

    // Reativar auto-save: habilita salvamento automático (debounced + periódico)
    try { window.__AUTO_SAVE_CLOUD.enabled = true; } catch (e) {}
    try { iniciarAutoSaveCloud(); } catch (e) { console.warn('Falha ao iniciar auto-save:', e); }
}

function carregarDados() {
    const dadosSalvos = localStorage.getItem('estoqueArmasV2');
    if (dadosSalvos) {
        estoque = JSON.parse(dadosSalvos);
        // Garantir que registroVendas existe
        if (!estoque.registroVendas) {
            estoque.registroVendas = [];
        }
        // Garantir que registroDistribuicao existe
        if (!estoque.registroDistribuicao) {
            estoque.registroDistribuicao = [];
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
        estoque.controleEnvio = {};
        estoque.auditoriaVendas = [];
        estoque.fechamentosComissoes = [];
        salvarDados();
    }
}

function salvarDados() {
    // marca hora local de atualização para comparação com o remoto
    try { estoque._localUpdatedAt = new Date().toISOString(); } catch (e) {}
    localStorage.setItem('estoqueArmasV2', JSON.stringify(estoque));
    atualizarEstatisticas();

    // agendar salvamento no cloud (debounced) se habilitado
    try { scheduleCloudSaveDebounced(); } catch (e) {}
}

// =============================
// Funções para salvar/carregar no Firestore
// =============================
async function salvarNoCloud() {
    if (!window.firestoreDB) {
        console.warn('Firestore não inicializado. Impossível salvar no cloud.');
        return false;
    }
    try {
        const docRef = window.firestoreDB.collection('app_data').doc('latest');
        await docRef.set({
            estado: estoque,
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
        console.log('Dados salvos no Firestore (coleção app_data / doc latest)');
        return true;
    } catch (e) {
        console.error('Erro salvando no Firestore:', e);
        updateFirestoreStatus(false, null, 'Cloud: erro ao salvar');
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
        if (!Array.isArray(estoque.auditoriaVendas)) estoque.auditoriaVendas = [];
        if (!Array.isArray(estoque.fechamentosComissoes)) estoque.fechamentosComissoes = [];
        salvarDados();
        renderizarTabela();
        renderizarDashboard();
        renderizarRegistroVendas();
        renderizarRegistroDistribuicao();
        renderizarControleEnvio();
        atualizarSelectsProdutos();
        atualizarSelectsRelatorios();
        console.log('Dados carregados do Firestore com sucesso.');
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
            if (!Array.isArray(estoque.auditoriaVendas)) estoque.auditoriaVendas = [];
            if (!Array.isArray(estoque.fechamentosComissoes)) estoque.fechamentosComissoes = [];
            salvarDados();
            renderizarTabela();
            renderizarDashboard();
            renderizarRegistroVendas();
            renderizarRegistroDistribuicao();
            renderizarControleEnvio();
            atualizarSelectsProdutos();
            atualizarSelectsRelatorios();
            console.log('Dados carregados automaticamente do Firestore (remoto mais recente).');
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
    document.getElementById('dataAtual').textContent = dataFormatada;
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
            if (Array.isArray(venda.items) && venda.items.length > 0) {
                venda.items.forEach(it => {
                    valorTotalVendas += it.valorTotal || 0;
                });
            } else {
                valorTotalVendas += venda.valorTotal || 0;
            }
        });
    }

    document.getElementById('totalProdutos').textContent = estoque.produtos.length;
    document.getElementById('totalEstoque').textContent = totalEstoque.toLocaleString('pt-BR');
    document.getElementById('totalVendas').textContent = totalVendas.toLocaleString('pt-BR');
    document.getElementById('valorTotalVendas').textContent = formatarMoedaValor(valorTotalVendas);
    // Calcular total de comissões (5%) excluindo vendas da IMBEL
    let totalComissoes = 0;
    if (Array.isArray(estoque.registroVendas)) {
        estoque.registroVendas.forEach(venda => {
            const rep = (venda.representante || '').toString().trim().toUpperCase();
            if (rep === 'IMBEL') return; // sem comissão
            const valor = typeof venda.valorTotal === 'number' ? venda.valorTotal : 0;
            totalComissoes += (Math.round((valor * 0.05) * 100) / 100);
        });
    }
    try { document.getElementById('totalComissoes').textContent = formatarMoedaValor(totalComissoes); } catch (e) {}
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
// NAVEGAÇÃO POR ABAS
// ========================================

function trocarAba(aba) {
    // Atualizar botões — usa data-tab para robustez
    document.querySelectorAll('.tabs-navigation .tab-btn').forEach(btn => {
        btn.classList.remove('active');
    });
    const activeBtn = document.querySelector(`.tabs-navigation .tab-btn[data-tab="${aba}"]`);
    if (activeBtn) activeBtn.classList.add('active');
    
    // Atualizar conteúdo
    document.querySelectorAll('.tab-content').forEach(content => content.classList.remove('active'));
    document.getElementById(`tab-${aba}`).classList.add('active');
    
    // Mostrar/ocultar barra de ações
    const acoesEstoque = document.getElementById('acoesEstoque');
    if (aba === 'estoque') {
        acoesEstoque.style.display = 'flex';
    } else {
        acoesEstoque.style.display = 'none';
    }
    
    // Renderizar conteúdo da aba
    if (aba === 'dashboard') {
        renderizarDashboard();
    } else if (aba === 'vendas') {
        renderizarRegistroVendas();
    } else if (aba === 'distribuicao') {
        renderizarRegistroDistribuicao();
        atualizarSelectDistribuicaoProduto();
    } else if (aba === 'relatorios') {
        prepararRelatorioInventario();
    } else if (aba === 'controleenvio') {
        renderizarControleEnvio();
    } else if (aba === 'controleimbel') {
        trocarSubAbaControleImbel('estoque');
    }
}

// ========================================
// RENDERIZAÇÃO DA TABELA DE ESTOQUE
// ========================================

function renderizarTabela() {
    const tbody = document.getElementById('corpoTabela');
    tbody.innerHTML = '';

    let totais = {
        KOLTE: { disp: 0, venda: 0, saldo: 0 },
        ISA: { disp: 0, venda: 0, saldo: 0 },
        LC: { disp: 0, venda: 0, saldo: 0 },
        ADES: { disp: 0, venda: 0, saldo: 0 },
        FL: { disp: 0, venda: 0, saldo: 0 },
        IMBEL: { disp: 0, venda: 0, saldo: 0 },
        GERAL: { disp: 0, venda: 0, saldo: 0 }
    };

    // Ordenar produtos alfabeticamente por nome para exibição
    const produtosOrdenados = [...estoque.produtos].sort((a, b) => a.nome.localeCompare(b.nome, 'pt-BR'));
    produtosOrdenados.forEach(produto => {
        const tr = document.createElement('tr');
        tr.dataset.id = produto.id;
        tr.innerHTML = `<td class="produto-nome col-produto" title="${produto.nome}">${produto.nome}</td>`;

        let geralDisp = 0;
        let geralVenda = 0;

        estoque.representantes.forEach(rep => {
            const disp = produto.distribuicao[rep] || 0;
            const venda = produto.vendas[rep] || 0;
            const saldo = disp - venda;

            geralDisp += disp;
            geralVenda += venda;

            totais[rep].disp += disp;
            totais[rep].venda += venda;
            totais[rep].saldo += saldo;

            const saldoClass = saldo < 0 ? 'negativo' : (saldo > 0 && saldo <= 5 ? 'baixo' : '');
            const animateClass = saldo < 0 ? 'animate-negativo' : '';

            tr.innerHTML += `
                <td class="cell-disp ${disp === 0 ? 'cell-zero' : ''}">${formatarNumero(disp)}</td>
                <td class="cell-venda ${venda === 0 ? 'cell-zero' : ''}">${formatarNumero(venda)}</td>
                <td class="cell-saldo ${saldoClass} ${saldo === 0 ? 'cell-zero' : ''} ${animateClass}">${formatarNumero(saldo)}</td>
            `;
        });

        const geralSaldo = geralDisp - geralVenda;
        totais.GERAL.disp += geralDisp;
        totais.GERAL.venda += geralVenda;
        totais.GERAL.saldo += geralSaldo;

        const saldoGeralClass = geralSaldo < 0 ? 'negativo' : (geralSaldo > 0 && geralSaldo <= 10 ? 'baixo' : '');
        const animateGeral = geralSaldo < 0 ? 'animate-negativo' : '';

        // Marcação vermelha quando saldo consolidado === 0
        if (geralSaldo === 0) {
            tr.classList.add('row-saldo-zero');
        }

        tr.innerHTML += `
            <td class="geral-disp">${formatarNumero(geralDisp)}</td>
            <td class="geral-venda">${formatarNumero(geralVenda)}</td>
            <td class="geral-saldo ${saldoGeralClass} ${animateGeral}">${formatarNumero(geralSaldo)}</td>
        `;

        tbody.appendChild(tr);
    });

    // Linha de totais
    const trTotal = document.createElement('tr');
    trTotal.className = 'total-row';
    trTotal.innerHTML = `<td class="produto-nome col-produto"><strong>TOTAL GERAL</strong></td>`;

    produtosOrdenados.forEach(repProd => {}); // placeholder to keep lint tools happy
    estoque.representantes.forEach(rep => {
        const saldoRep = totais[rep].saldo;
        const saldoRepClass = saldoRep < 0 ? 'negativo' : (saldoRep > 0 && saldoRep <= 5 ? 'baixo' : '');
        const animateRep = saldoRep < 0 ? 'animate-negativo' : '';
        trTotal.innerHTML += `
            <td class="cell-disp"><strong>${formatarNumero(totais[rep].disp)}</strong></td>
            <td class="cell-venda"><strong>${formatarNumero(totais[rep].venda)}</strong></td>
            <td class="cell-saldo ${saldoRepClass} ${animateRep}"><strong>${formatarNumero(totais[rep].saldo)}</strong></td>
        `;
    });

    const saldoGeralTotal = totais.GERAL.saldo;
    const saldoGeralTotalClass = saldoGeralTotal < 0 ? 'negativo' : (saldoGeralTotal > 0 && saldoGeralTotal <= 10 ? 'baixo' : '');
    const animateGeralTotal = saldoGeralTotal < 0 ? 'animate-negativo' : '';

    trTotal.innerHTML += `
        <td class="geral-disp"><strong>${formatarNumero(totais.GERAL.disp)}</strong></td>
        <td class="geral-venda"><strong>${formatarNumero(totais.GERAL.venda)}</strong></td>
        <td class="geral-saldo ${saldoGeralTotalClass} ${animateGeralTotal}"><strong>${formatarNumero(totais.GERAL.saldo)}</strong></td>
    `;

    tbody.appendChild(trTotal);

    // Ajustar posição sticky da segunda linha do header
    ajustarStickyHeader();
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
        ['Disp', 'Venda', 'Saldo', 'Disp', 'Venda', 'Saldo'].forEach(text => {
            const t = document.createElement('th');
            t.className = 'sub-header';
            t.textContent = text;
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
    const vendas = Array.isArray(estoque.registroVendas) ? [...estoque.registroVendas] : [];
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
    const vendas = Array.isArray(estoque.registroVendas) ? [...estoque.registroVendas] : [];
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
    const vendasRaw = Array.isArray(estoque.registroVendas) ? [...estoque.registroVendas] : [];
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

    const snapshot = {
        id: Date.now(),
        chave,
        competencia,
        filtroRep: filtroRep || '',
        criadoEm: new Date().toISOString(),
        criadoPor: getUsuarioAtual(),
        linhas: contratos.map(c => ({
            contrato: c.contrato,
            loja: c.loja,
            representantes: Array.from(c.representantes || []),
            dataMin: c.dataMin || null,
            dataMax: c.dataMax || null,
            valorContrato: c.valorContrato,
            comissao: c.comissao
        })),
        totalComissoes
    };

    if (idxExistente !== -1) estoque.fechamentosComissoes[idxExistente] = snapshot;
    else estoque.fechamentosComissoes.push(snapshot);

    salvarDados();
    mostrarNotificacao(`Fechamento ${competencia} salvo com ${snapshot.linhas.length} contrato(s).`, 'success');
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

    produtos.forEach(p => {
        const estoqueAtual = calcSaldo(p.id);
        const ponto = (p.pontoReposicao !== undefined) ? Number(p.pontoReposicao) : Math.max(1, Math.floor((Number(p.quantidadeInicial)||0) * 0.2));
        let status = 'OK';
        let color = '#28a745';
        if (estoqueAtual <= 0) { status = 'CRÍTICO'; color = '#dc3545'; }
        else if (estoqueAtual <= ponto) { status = 'ATENÇÃO'; color = '#ffc107'; }

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
    tabela.innerHTML = `<thead><tr>
        <th style="${thStyle}">ID</th>
        <th style="${thStyle}">Descrição</th>
        <th style="${thStyle}">Observação</th>
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
        tr.innerHTML = `<td colspan="7" style="${tdBase};text-align:center;color:#999;font-style:italic">Nenhum produto cadastrado. Cadastre na sub-aba "Cadastro de Produtos".</td>`;
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
        tr.innerHTML = `
            <td style="${tdCenter}">${idx + 1}</td>
            <td style="${tdBase}">${p.nome}${p.codigo ? ' <span style="color:#888;font-size:.75rem">('+p.codigo+')</span>' : ''} <button class="btn btn-outline" data-editprod="${p.id}" style="margin-left:8px;padding:4px 8px;font-size:.8rem">Editar</button></td>
            <td style="${tdBase}">${p.observacao || '-'}</td>
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
            <td colspan="3" style="${tdBase};background:#1e3a5f;color:#fff;text-align:right">TOTAL GERAL</td>
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

    const formWrap = document.createElement('div');
    formWrap.style.cssText = 'background:#fff;border-radius:10px;padding:16px;margin-bottom:16px;box-shadow:0 1px 4px rgba(0,0,0,.08)';
    formWrap.innerHTML = `
        <input type="hidden" id="imbel_prod_edit_id" />
        <div style="display:grid;grid-template-columns:2fr 1fr 1fr 2fr auto;gap:10px;align-items:end">
            <div><label style="font-size:.82rem;font-weight:600;color:#555;display:block;margin-bottom:4px">Nome do Produto</label>
                <input type="text" id="imbel_prod_nome" placeholder="Nome" style="width:100%;padding:7px 8px;border:1px solid #ddd;border-radius:6px;font-size:.85rem"/></div>
            <div><label style="font-size:.82rem;font-weight:600;color:#555;display:block;margin-bottom:4px">Código</label>
                <input type="text" id="imbel_prod_codigo" placeholder="Cód." style="width:100%;padding:7px 8px;border:1px solid #ddd;border-radius:6px;font-size:.85rem"/></div>
            <div><label style="font-size:.82rem;font-weight:600;color:#555;display:block;margin-bottom:4px">Quantidade Inicial</label>
                <input type="number" id="imbel_prod_qtd_inicial" value="0" min="0" style="width:100%;padding:7px 8px;border:1px solid #ddd;border-radius:6px;font-size:.85rem"/></div>
            <div><label style="font-size:.82rem;font-weight:600;color:#555;display:block;margin-bottom:4px">Observação</label>
                <input type="text" id="imbel_prod_obs" placeholder="Observação" style="width:100%;padding:7px 8px;border:1px solid #ddd;border-radius:6px;font-size:.85rem"/></div>
            <div><button type="button" id="imbel_prod_salvar" class="btn btn-primary">Salvar</button></div>
        </div>
    `;
    container.appendChild(formWrap);

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
                renderControleImbelCadastro();
                renderControleImbelEstoque();
                return;
            }
        }
        const novo = { id: 'p' + Date.now(), nome, codigo, observacao, quantidadeInicial };
        data.produtos.push(novo);
        saveImbel(data);
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
        // preencher formulário
        document.getElementById('imbel_prod_nome').value = prod.nome || '';
        document.getElementById('imbel_prod_codigo').value = prod.codigo || '';
        document.getElementById('imbel_prod_qtd_inicial').value = prod.quantidadeInicial || 0;
        document.getElementById('imbel_prod_obs').value = prod.observacao || '';
        const editField = document.getElementById('imbel_prod_edit_id'); if (editField) editField.value = id;
        document.getElementById('imbel_prod_salvar').textContent = 'Atualizar';
        // focus form
        document.getElementById('imbel_prod_nome').focus();
    });
}

function renderControleImbelMovimentacao() {
    const data = loadImbel();
    const container = document.getElementById('controleImbelMovContainer');
    container.innerHTML = '';

    // ---- Formulário ----
    const formWrap = document.createElement('div');
    formWrap.style.cssText = 'background:#fff;border-radius:10px;padding:20px;margin-bottom:20px;box-shadow:0 1px 4px rgba(0,0,0,.08)';
    formWrap.innerHTML = `
        <input type="hidden" id="imbel_mov_edit_id" />
        <div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(200px,1fr));gap:12px;align-items:end">
            <div>
                <label style="font-size:.82rem;font-weight:600;color:#555;display:block;margin-bottom:4px">Descrição (Produto)</label>
                <select id="imbel_mov_prod" style="width:100%;padding:7px 8px;border:1px solid #ddd;border-radius:6px;font-size:.85rem"></select>
            </div>
            <div>
                <label style="font-size:.82rem;font-weight:600;color:#555;display:block;margin-bottom:4px">Tipo</label>
                <select id="imbel_mov_tipo" style="width:100%;padding:7px 8px;border:1px solid #ddd;border-radius:6px;font-size:.85rem">
                    <option value="Entrada">Entrada</option>
                    <option value="Saída">Saída</option>
                </select>
            </div>
            <div>
                <label style="font-size:.82rem;font-weight:600;color:#555;display:block;margin-bottom:4px">Data</label>
                <input type="date" id="imbel_mov_data" style="width:100%;padding:7px 8px;border:1px solid #ddd;border-radius:6px;font-size:.85rem"/>
            </div>
            <div>
                <label style="font-size:.82rem;font-weight:600;color:#555;display:block;margin-bottom:4px">Quantidade</label>
                <input type="number" id="imbel_mov_qtd" value="1" min="1" style="width:100%;padding:7px 8px;border:1px solid #ddd;border-radius:6px;font-size:.85rem"/>
            </div>
            <div>
                <label style="font-size:.82rem;font-weight:600;color:#555;display:block;margin-bottom:4px">Destinatário</label>
                <input type="text" id="imbel_mov_dest" placeholder="Nome" style="width:100%;padding:7px 8px;border:1px solid #ddd;border-radius:6px;font-size:.85rem"/>
            </div>
            <div>
                <label style="font-size:.82rem;font-weight:600;color:#555;display:block;margin-bottom:4px">CPF / CNPJ</label>
                <input type="text" id="imbel_mov_cpf" placeholder="000.000.000-00" style="width:100%;padding:7px 8px;border:1px solid #ddd;border-radius:6px;font-size:.85rem"/>
            </div>
            <div>
                <label style="font-size:.82rem;font-weight:600;color:#555;display:block;margin-bottom:4px">Valor (R$)</label>
                <input type="number" id="imbel_mov_valor" min="0" step="0.01" placeholder="0,00" style="width:100%;padding:7px 8px;border:1px solid #ddd;border-radius:6px;font-size:.85rem"/>
            </div>
            <div style="grid-column:span 2">
                <label style="font-size:.82rem;font-weight:600;color:#555;display:block;margin-bottom:4px">Endereço</label>
                <input type="text" id="imbel_mov_endereco" placeholder="Rua, número, cidade" style="width:100%;padding:7px 8px;border:1px solid #ddd;border-radius:6px;font-size:.85rem"/>
            </div>
            <div>
                <label style="font-size:.82rem;font-weight:600;color:#555;display:block;margin-bottom:4px">Telefone</label>
                <input type="text" id="imbel_mov_tel" placeholder="(00) 00000-0000" style="width:100%;padding:7px 8px;border:1px solid #ddd;border-radius:6px;font-size:.85rem"/>
            </div>
            <div>
                <label style="font-size:.82rem;font-weight:600;color:#555;display:block;margin-bottom:4px">E-mail</label>
                <input type="email" id="imbel_mov_email" placeholder="email@exemplo.com" style="width:100%;padding:7px 8px;border:1px solid #ddd;border-radius:6px;font-size:.85rem"/>
            </div>
            <!-- Pagamento / Entregue / FI serão controlados via checkboxes na TABELA, não no cadastro -->
            <div style="grid-column:1/-1">
                <label style="font-size:.82rem;font-weight:600;color:#555;display:block;margin-bottom:4px">Observações</label>
                <input type="text" id="imbel_mov_obs" placeholder="Observações adicionais" style="width:100%;padding:7px 8px;border:1px solid #ddd;border-radius:6px;font-size:.85rem"/>
            </div>
            <div style="grid-column:1/-1;display:flex;gap:8px;justify-content:flex-end;margin-top:4px">
                <button type="button" id="imbel_mov_limpar" class="btn btn-outline">
                    <span class="btn-icon">🗑️</span> Limpar
                </button>
                <button type="button" id="imbel_mov_clear_table" class="btn btn-outline">
                    <span class="btn-icon">🧹</span> Limpar Tabela
                </button>
                <button type="button" id="imbel_mov_salvar" class="btn btn-primary">
                    <span class="btn-icon">💾</span> Registrar Movimentação
                </button>
            </div>
        </div>
    `;
    container.appendChild(formWrap);

    // Popular select de produtos
    const selec = document.getElementById('imbel_mov_prod');
    selec.innerHTML = '<option value="">— selecione o produto —</option>';
    (data.produtos||[]).forEach(p => {
        const opt = document.createElement('option');
        opt.value = p.id;
        opt.textContent = p.nome + (p.codigo ? ' ('+p.codigo+')' : '');
        selec.appendChild(opt);
    });

    // Data padrão = hoje
    const hoje = new Date().toISOString().slice(0,10);
    document.getElementById('imbel_mov_data').value = hoje;

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
            tr.style.background = rowIndex % 2 === 0 ? '#fff' : '#f7f9fc';
            tr.innerHTML = `
                <td style="${tdCenter}">${seqNum}</td>
                <td style="${tdStyle}">${produto.nome}</td>
                <td style="${tdCenter}"><span style="padding:2px 8px;border-radius:12px;font-size:.75rem;font-weight:600;color:${(m.tipo||'').toString().toUpperCase()==='ENTRADA' ? '#155724' : ((m.tipo||'').toString().toUpperCase()==='SAIDA' || (m.tipo||'').toString().toUpperCase()==='SAÍDA' ? '#721c24' : '#333')};">${((m.tipo||'').toString().toUpperCase()==='ENTRADA' ? 'Entrada' : (((m.tipo||'').toString().toUpperCase()==='SAIDA' || (m.tipo||'').toString().toUpperCase()==='SAÍDA') ? 'Saída' : (m.tipo||'-')))}</span></td>
                <td style="${tdCenter}">${dataFmt}</td>
                <td style="${tdCenter}">${m.quantidade||'-'}</td>
                <td style="${tdStyle}">${m.destinatario||'-'}</td>
                <td style="${tdStyle}">${m.cpfCnpj||'-'}</td>
                <td style="${tdStyle}">${m.observacoes||'-'}</td>
                <td style="${tdCenter}">${valorFmt}</td>
                <td style="${tdStyle}">${m.endereco||'-'}</td>
                <td style="${tdStyle}">${m.telefone||'-'}</td>
                <td style="${tdStyle}">${m.email||'-'}</td>
                <td style="${tdCenter}"><input type=\"checkbox\" class=\"imbel_table_chk_pag\" data-id=\"${m.id}\" ${((m.pagamento||'').toString().toUpperCase()==='SIM')? 'checked' : ''}></td>
                <td style="${tdCenter}"><input type=\"checkbox\" class=\"imbel_table_chk_ent\" data-id=\"${m.id}\" ${((m.entregue||'').toString().toUpperCase()==='SIM')? 'checked' : ''}></td>
                <td style="${tdCenter}"><input type=\"checkbox\" class=\"imbel_table_chk_fi\" data-id=\"${m.id}\" ${((m.fi||'').toString().toUpperCase()==='SIM')? 'checked' : ''}></td>
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
                document.getElementById('imbel_mov_prod').value = mov.produtoId || '';
                const tipoLower = (mov.tipo||'').toString().toLowerCase();
                if (tipoLower === 'entrada') document.getElementById('imbel_mov_tipo').value = 'Entrada';
                else if (tipoLower === 'saída' || tipoLower === 'saida') document.getElementById('imbel_mov_tipo').value = 'Saída';
                else document.getElementById('imbel_mov_tipo').value = mov.tipo || 'Entrada';
                document.getElementById('imbel_mov_data').value = mov.data || hoje;
                document.getElementById('imbel_mov_qtd').value = mov.quantidade || 1;
                document.getElementById('imbel_mov_dest').value = mov.destinatario || '';
                document.getElementById('imbel_mov_cpf').value = mov.cpfCnpj || '';
                document.getElementById('imbel_mov_valor').value = mov.valor || '';
                document.getElementById('imbel_mov_endereco').value = mov.endereco || '';
                document.getElementById('imbel_mov_tel').value = mov.telefone || '';
                document.getElementById('imbel_mov_email').value = mov.email || '';
                // pagamento/entregue/fi são gerenciados na tabela como checkboxes
                document.getElementById('imbel_mov_obs').value = mov.observacoes || '';
                const editField = document.getElementById('imbel_mov_edit_id'); if (editField) editField.value = id;
                document.getElementById('imbel_mov_salvar').textContent = 'Atualizar Movimentação';
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

    tabelaWrap.appendChild(tabela);
    container.appendChild(tabelaWrap);

    // ---- Handler: Salvar ----
    document.getElementById('imbel_mov_salvar').onclick = function(){
        const produtoId = document.getElementById('imbel_mov_prod').value;
        const tipo = document.getElementById('imbel_mov_tipo').value;
        const quantidade = Number(document.getElementById('imbel_mov_qtd').value) || 0;
        const dataStr = document.getElementById('imbel_mov_data').value || hoje;
        const destinatario = document.getElementById('imbel_mov_dest').value.trim();
        const cpfCnpj = document.getElementById('imbel_mov_cpf').value.trim();
        const valor = parseFloat(document.getElementById('imbel_mov_valor').value) || 0;
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
        renderControleImbelMovimentacao();
        renderControleImbelEstoque();
    };

    // ---- Handler: Limpar ----
    document.getElementById('imbel_mov_limpar').onclick = function(){
        document.getElementById('imbel_mov_prod').value = '';
        document.getElementById('imbel_mov_qtd').value = 1;
        document.getElementById('imbel_mov_data').value = hoje;
        document.getElementById('imbel_mov_dest').value = '';
        document.getElementById('imbel_mov_cpf').value = '';
        document.getElementById('imbel_mov_valor').value = '';
        document.getElementById('imbel_mov_endereco').value = '';
        document.getElementById('imbel_mov_tel').value = '';
        document.getElementById('imbel_mov_email').value = '';
        document.getElementById('imbel_mov_obs').value = '';
        // reset edit state
        const editField = document.getElementById('imbel_mov_edit_id'); if (editField) editField.value = '';
        document.getElementById('imbel_mov_salvar').textContent = 'Registrar Movimentação';
    };

    // ---- Handler: Limpar Tabela ----
    const btnClearTable = document.getElementById('imbel_mov_clear_table');
    if (btnClearTable) {
        btnClearTable.onclick = function(){
            if (!confirm('Deseja apagar TODAS as movimentações? Esta ação é irreversible.')) return;
            data.movimentacoes = [];
            saveImbel(data);
            mostrarNotificacao('Todas as movimentações foram removidas.', 'success');
            renderControleImbelMovimentacao();
            renderControleImbelEstoque();
            renderControleImbelCadastro();
        };
    }

    // ---- Handler: Deletar ----
    tbody.querySelectorAll('button[data-delid]').forEach(btn => {
        btn.onclick = function(){
            const id = this.getAttribute('data-delid');
            if (!confirm('Excluir esta movimentação?')) return;
            data.movimentacoes = (data.movimentacoes||[]).filter(m => m.id !== id);
            saveImbel(data);
            mostrarNotificacao('Movimentação excluída.', 'success');
            renderControleImbelMovimentacao();
            renderControleImbelEstoque();
        };
    });

    // ---- Handler: Editar Movimentação ----
    tbody.querySelectorAll('button[data-editmov]').forEach(btn => {
        btn.onclick = function(){
            const id = this.getAttribute('data-editmov');
            const mov = (data.movimentacoes||[]).find(m => m.id === id);
            if (!mov) return;
            document.getElementById('imbel_mov_prod').value = mov.produtoId || '';
            const tipoLower = (mov.tipo||'').toString().toLowerCase();
            if (tipoLower === 'entrada') document.getElementById('imbel_mov_tipo').value = 'Entrada';
            else if (tipoLower === 'saída' || tipoLower === 'saida') document.getElementById('imbel_mov_tipo').value = 'Saída';
            else document.getElementById('imbel_mov_tipo').value = mov.tipo || 'Entrada';
            document.getElementById('imbel_mov_data').value = mov.data || hoje;
            document.getElementById('imbel_mov_qtd').value = mov.quantidade || 1;
            document.getElementById('imbel_mov_dest').value = mov.destinatario || '';
            document.getElementById('imbel_mov_cpf').value = mov.cpfCnpj || '';
            document.getElementById('imbel_mov_valor').value = mov.valor || '';
            document.getElementById('imbel_mov_endereco').value = mov.endereco || '';
            document.getElementById('imbel_mov_tel').value = mov.telefone || '';
            document.getElementById('imbel_mov_email').value = mov.email || '';
            // pagamento/entregue/fi são gerenciados na tabela como checkboxes
            document.getElementById('imbel_mov_obs').value = mov.observacoes || '';
            const editField = document.getElementById('imbel_mov_edit_id'); if (editField) editField.value = id;
            document.getElementById('imbel_mov_salvar').textContent = 'Atualizar Movimentação';
            document.getElementById('imbel_mov_prod').focus();
        };
    });
}

// Inicializar controle IMBEL (não altera outros dados)
try { initControleImbel(); } catch(e) {}

// Abre a sub-aba de Cadastro e preenche o formulário para editar o produto indicado
function editarProdutoPorId(id) {
    try {
        const data = loadImbel();
        const prod = (data.produtos||[]).find(p => p.id === id);
        if (!prod) { mostrarNotificacao('Produto não encontrado para edição.', 'error'); return; }
        trocarSubAbaControleImbel('cadastro');
        // aguardar renderização do formulário e então preencher
        setTimeout(() => {
            const nomeEl = document.getElementById('imbel_prod_nome');
            if (!nomeEl) return;
            document.getElementById('imbel_prod_nome').value = prod.nome || '';
            document.getElementById('imbel_prod_codigo').value = prod.codigo || '';
            document.getElementById('imbel_prod_qtd_inicial').value = prod.quantidadeInicial || 0;
            document.getElementById('imbel_prod_obs').value = prod.observacao || '';
            const editField = document.getElementById('imbel_prod_edit_id'); if (editField) editField.value = id;
            document.getElementById('imbel_prod_salvar').textContent = 'Atualizar';
            document.getElementById('imbel_prod_nome').focus();
        }, 60);
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
                console.log('Erros de importação IMBEL:', erros);
            } else if (erros.length > 0) {
                mostrarNotificacao(`${importados} registros importados. ${erros.length} linhas com erro.`, 'warning');
                console.log('Erros de importação IMBEL:', erros);
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
            if (erros.length>0 && importados===0) { mostrarNotificacao('Nenhum produto importado. Verifique o formato.', 'error'); console.log('Erros IMBEL cadastro:',erros); }
            else if (erros.length>0) { mostrarNotificacao(`${importados} produtos importados. ${erros.length} linhas com erro.`, 'warning'); console.log('Erros IMBEL cadastro:',erros); }
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
            if (erros.length>0 && importados===0) { mostrarNotificacao('Nenhuma movimentação importada. Verifique o formato.', 'error'); console.log('Erros IMBEL mov:',erros); }
            else if (erros.length>0) { mostrarNotificacao(`${importados} movimentações importadas. ${erros.length} linhas com erro.`, 'warning'); console.log('Erros IMBEL mov:',erros); }
            else { mostrarNotificacao(`${importados} movimentações importadas com sucesso!`, 'success'); }
        } catch(err) { console.error('Erro importar movimentações IMBEL',err); mostrarNotificacao('Erro ao processar o arquivo.', 'error'); }
    };
    reader.readAsText(file,'UTF-8');
}


function renderizarDashboard() {
    // Calcular dados
    let dadosVendas = [];
    let vendasPorRep = { ADES: 0, FL: 0, IMBEL: 0, ISA: 0, KOLTE: 0, LC: 0 };
    let totalUnidades = 0;
    let totalFaturamento = 0;

    // Valor por produto deve vir apenas das vendas registradas
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

    estoque.produtos.forEach(produto => {
        let totalProduto = 0;
        let vendasProdutoPorRep = {};
        
        estoque.representantes.forEach(rep => {
            const venda = produto.vendas[rep] || 0;
            totalProduto += venda;
            vendasPorRep[rep] = (vendasPorRep[rep] || 0) + venda;
            vendasProdutoPorRep[rep] = venda;
        });

        const valorTotal = valorPorProduto[produto.nome] || 0;
        
        dadosVendas.push({
            nome: produto.nome,
            quantidade: totalProduto,
            valor: valorTotal,
            vendasPorRep: vendasProdutoPorRep
        });

        totalUnidades += totalProduto;
        totalFaturamento += valorTotal;
    });

    // Ordenar por quantidade
    dadosVendas.sort((a, b) => b.quantidade - a.quantidade);

    // Encontrar melhor representante
    let melhorRep = '-';
    let maxVendas = 0;
    Object.entries(vendasPorRep).forEach(([rep, vendas]) => {
        if (vendas > maxVendas) {
            maxVendas = vendas;
            melhorRep = rep;
        }
    });

    // Produto mais vendido
    const produtoTop = dadosVendas.length > 0 && dadosVendas[0].quantidade > 0 
        ? dadosVendas[0].nome.substring(0, 25) + (dadosVendas[0].nome.length > 25 ? '...' : '')
        : '-';

    // Criar versão ordenada alfabeticamente para exibição nas tabelas do dashboard
    const dadosVendasAlpha = [...dadosVendas].sort((a, b) => a.nome.localeCompare(b.nome, 'pt-BR'));

    // Atualizar cards
    document.getElementById('dashTotalUnidades').textContent = totalUnidades.toLocaleString('pt-BR');
    document.getElementById('dashTotalFaturamento').textContent = formatarMoedaValor(totalFaturamento);
    document.getElementById('dashMelhorRep').textContent = melhorRep + ` (${maxVendas})`;
    document.getElementById('dashProdutoTop').textContent = produtoTop;

    // Tabela: Quantidade por Produto
    const tabelaQtd = document.getElementById('tabelaQtdProduto');
    tabelaQtd.innerHTML = '';
    
    dadosVendasAlpha.forEach(item => {
        if (item.quantidade > 0) {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td class="produto-nome">${item.nome}</td>
                <td class="cell-qtd">${item.quantidade.toLocaleString('pt-BR')}</td>
            `;
            tabelaQtd.appendChild(tr);
        }
    });

    // Linha total
    const trTotalQtd = document.createElement('tr');
    trTotalQtd.className = 'total-row';
    trTotalQtd.innerHTML = `
        <td class="produto-nome"><strong>Total Geral</strong></td>
        <td class="cell-qtd"><strong>${totalUnidades.toLocaleString('pt-BR')}</strong></td>
    `;
    tabelaQtd.appendChild(trTotalQtd);

    // Tabela: Valor por Produto
    const tabelaValor = document.getElementById('tabelaValorProduto');
    tabelaValor.innerHTML = '';
    
    // Exibir lista de produtos em ordem alfabética por nome (valor exibido permanece)
    dadosVendasAlpha.forEach(item => {
        if (item.valor > 0) {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td class="produto-nome">${item.nome}</td>
                <td class="cell-valor">${formatarMoedaValor(item.valor)}</td>
            `;
            tabelaValor.appendChild(tr);
        }
    });

    // Linha total
    const trTotalValor = document.createElement('tr');
    trTotalValor.className = 'total-row';
    trTotalValor.innerHTML = `
        <td class="produto-nome"><strong>Total Geral</strong></td>
        <td class="cell-valor"><strong>${formatarMoedaValor(totalFaturamento)}</strong></td>
    `;
    tabelaValor.appendChild(trTotalValor);

    // Tabela: Vendas por Representante
    const tabelaRep = document.getElementById('tabelaVendasRepBody');
    tabelaRep.innerHTML = '';
    
    const repsOrdem = ['ADES', 'FL', 'IMBEL', 'ISA', 'KOLTE', 'LC'];
    let totaisPorRep = { ADES: 0, FL: 0, IMBEL: 0, ISA: 0, KOLTE: 0, LC: 0 };

    dadosVendasAlpha.forEach(item => {
        if (item.quantidade > 0) {
            const tr = document.createElement('tr');
            let html = `<td class="produto-nome">${item.nome}</td>`;

            repsOrdem.forEach(rep => {
                const venda = item.vendasPorRep[rep] || 0;
                totaisPorRep[rep] += venda;
                html += `<td class="${venda === 0 ? 'cell-zero' : 'cell-qtd'}">${venda > 0 ? venda : '-'}</td>`;
            });

            html += `<td class="geral-venda"><strong>${item.quantidade}</strong></td>`;
            tr.innerHTML = html;
            tabelaRep.appendChild(tr);
        }
    });

    // Linha total
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
    document.getElementById('modalProduto').style.display = 'flex';
    document.getElementById('formProduto').reset();
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
    // Popular select de produtos
    const selectProduto = document.getElementById('produtoDevolucao');
    if (selectProduto) {
        selectProduto.innerHTML = '<option value="">Selecione um produto</option>';
        const produtosOrdenados = [...estoque.produtos].sort((a, b) => a.nome.localeCompare(b.nome, 'pt-BR'));
        produtosOrdenados.forEach(produto => {
            const option = document.createElement('option');
            option.value = produto.id;
            option.textContent = produto.nome;
            selectProduto.appendChild(option);
        });
    }

    // Popular select de representantes (origem)
    const selectRep = document.getElementById('representanteDevolucao');
    if (selectRep) {
        selectRep.innerHTML = '<option value="">Selecione o representante</option>';
        estoque.representantes.forEach(rep => {
            const opt = document.createElement('option');
            opt.value = rep;
            opt.textContent = rep;
            selectRep.appendChild(opt);
        });
    }

    // Popular select destino (IMBEL por padrão, mas permitir redistribuir para qualquer rep)
    const selectDestino = document.getElementById('destinoDevolucao');
    if (selectDestino) {
        selectDestino.innerHTML = '<option value="IMBEL">IMBEL (Retornar ao estoque central)</option>';
        estoque.representantes.forEach(rep => {
            if (rep !== 'IMBEL') {
                const opt = document.createElement('option');
                opt.value = rep;
                opt.textContent = rep;
                selectDestino.appendChild(opt);
            }
        });
    }
}

// Fecha um modal e restaura z-index do header fixo se necessário
function fecharModal(modalId) {
    try {
        const modal = document.getElementById(modalId);
        if (modal) modal.style.display = 'none';
    } catch (e) { /* ignore */ }
    // Sempre tentar restaurar z-index do header fixo quando um modal fechar
}

function mostrarEstoqueAtual() {
    const produtoId = parseInt(document.getElementById('produtoEntrada').value);
    const produto = estoque.produtos.find(p => p.id === produtoId);
    
    if (produto) {
        const estoqueIMBEL = (produto.distribuicao.IMBEL || 0) - (produto.vendas.IMBEL || 0);
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
    produto.distribuicao.IMBEL = (produto.distribuicao.IMBEL || 0) + quantidade;
    
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

function abrirModalVendaDetalhada(vendaId = null) {
    // vendaId: se fornecido, abre o modal em modo de edição para essa venda
    if (!requireAdminOrNotify()) return;
    const modalEl = document.getElementById('modalVendaDetalhada');
    modalEl.style.display = 'flex';
    document.getElementById('formVendaDetalhada').reset();
    document.getElementById('valorUnitarioVenda').value = '';
    document.getElementById('valorTotalVenda').value = '';
    atualizarSelectsProdutos();

    const container = document.getElementById('itensVendaContainer');
    if (!container) return;

    // Se estamos criando nova venda, inicializa com uma linha vazia e sugere contrato
    if (!vendaId) {
        vendaEditandoId = null;
        container.innerHTML = '';
        adicionarItemVendaRow();

        const ultimoContrato = estoque.registroVendas.length > 0 
            ? Math.max(...estoque.registroVendas.map(v => parseInt(v.contrato) || 0)) 
            : 0;
        document.getElementById('contratoVenda').value = ultimoContrato + 1;
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
            <select class="item-produto" onchange="atualizarItemRow(this)">${opcoesHtml}</select>
            <input type="number" class="item-quantidade" min="1" value="${preQuantidade}" style="width:90px" onchange="atualizarItemRow(this)" />
            <input type="text" class="item-valor" placeholder="Valor unit. (obrigatório)" style="width:140px" oninput="formatarMoeda(this); atualizarItemRow(this)" />
            <div class="item-subtotal" style="min-width:120px">-</div>
            <button type="button" class="btn btn-outline btn-sm" onclick="removerItemRow(this)">Remover</button>
        `;

        container.appendChild(row);

        // Preencher valores se fornecidos
        if (preProdutoId) row.querySelector('.item-produto').value = preProdutoId;
        if (preValor) row.querySelector('.item-valor').value = preValor;

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

    // Validar estoque para cada item (no representante selecionado)
    let falta = [];
    itens.forEach(it => {
        const produto = estoque.produtos.find(p => p.id === it.produtoId);
        const disp = produto.distribuicao[representante] || 0;
        const vendido = produto.vendas[representante] || 0;
        const saldo = disp - vendido;
        if (it.quantidade > saldo) {
            falta.push(`${produto.nome}: disponível ${saldo}, solicitado ${it.quantidade}`);
        }
    });

    if (falta.length > 0) {
        // Restaurar venda anterior caso haja erro na validação
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

        mostrarNotificacao('Estoque insuficiente:\n' + falta.join('\n'), 'error');
        return;
    }

    // Aplicar novos valores ao estoque
    itens.forEach(it => {
        const produto = estoque.produtos.find(p => p.id === it.produtoId);
        produto.vendas[representante] = (produto.vendas[representante] || 0) + it.quantidade;
    });

    if (isEditing && vendaAnterior) {
        // Atualizar o registro existente
        const idx = estoque.registroVendas.findIndex(v => v.id === vendaEditandoId);
        if (idx !== -1) {
            estoque.registroVendas[idx].contrato = contrato;
            estoque.registroVendas[idx].loja = loja;
            estoque.registroVendas[idx].representante = representante;
            estoque.registroVendas[idx].items = itens;
            estoque.registroVendas[idx].quantidadeTotal = totalQtd;
            estoque.registroVendas[idx].valorTotal = totalValor;
            estoque.registroVendas[idx].observacoes = observacoes;
            // usar data informada pelo usuário, se presente; caso contrário, registrar timestamp atual
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
        // usar data informada pelo usuário (YYYY-MM-DD) convertida para ISO, ou timestamp atual
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
                <td colspan="11" class="empty-state">
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
        const resumo = document.createElement('tr');
        resumo.className = 'row-contrato-resumo';
        resumo.innerHTML = `
            <td class="col-contrato">${contratoKey || '-'}</td>
            <td class="col-loja" title="${primeira.loja}">${primeira.loja}</td>
            <td class="col-representante"><span class="badge-rep ${repClass}">${primeira.representante}</span></td>
            <td class="col-produto-venda"><button class="btn-expand-contrato" onclick="toggleContratoExpandido('${contratoKey}')">${expandido ? '▾' : '▸'} ${linhasDoContrato} item(ns)</button></td>
            <td class="col-qtd">${totalQtdContrato}</td>
            <td class="col-valor-un">-</td>
            <td class="col-valor-total">${formatarMoedaValor(totalContrato)}</td>
            <td class="col-data">${dataDisplay}</td>
            <td class="col-total-contrato">${formatarMoedaValor(totalContrato)}</td>
            <td class="col-obs" title="${obsGrupo}">${obsGrupo}</td>
            <td class="col-acoes">
                <button class="btn-action btn-edit" onclick="abrirModalVendaDetalhada(${primeira.vendaId})" title="Editar venda">✎</button>
                <button class="btn-action btn-delete" onclick="excluirVenda(${primeira.vendaId})" title="Excluir venda">🗑</button>
                <button class="btn-action" onclick="abrirHistoricoContrato('${contratoKey}')" title="Histórico do Contrato">🕘</button>
            </td>
        `;
        tbody.appendChild(resumo);

        grupo.forEach((linha) => {
            const tr = document.createElement('tr');
            tr.className = `row-contrato-detalhe ${expandido ? '' : 'hidden-row'}`;
            const valorUn = linha.valorUnitario ? formatarMoedaValor(linha.valorUnitario) : '-';
            const valorTot = linha.valorTotal || 0;
            totalQtd += linha.quantidade || 0;
            totalValor += valorTot || 0;

            tr.innerHTML = `
                <td class="col-contrato detalhe-vazio"></td>
                <td class="col-loja detalhe-vazio"></td>
                <td class="col-representante detalhe-vazio"></td>
                <td class="col-produto-venda" title="${linha.produtoNome}">↳ ${linha.produtoNome}</td>
                <td class="col-qtd">${linha.quantidade}</td>
                <td class="col-valor-un">${valorUn}</td>
                <td class="col-valor-total">${valorTot > 0 ? formatarMoedaValor(valorTot) : '-'}</td>
                <td class="col-data">${linha.dataNorm ? formatDateToDDMMYYYY(linha.dataNorm) : '-'}</td>
                <td class="col-total-contrato">-</td>
                <td class="col-obs" title="${linha.observacoes || '-'}">${linha.observacoes || '-'}</td>
                <td class="col-acoes">
                    <button class="btn-action btn-edit" onclick="abrirModalVendaDetalhada(${linha.vendaId})" title="Editar venda">✎</button>
                    <button class="btn-action btn-delete" onclick="excluirVenda(${linha.vendaId})" title="Excluir venda">🗑</button>
                </td>
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
    
    // Definir data atual
    const hoje = new Date().toISOString().split('T')[0];
    document.getElementById('dataDistDet').value = hoje;
}

function salvarNovaDistribuicao(event) {
    if (!requireAdminOrNotify()) return;
    event.preventDefault();
    
    const representante = document.getElementById('representanteDistDet').value;
    const produtoId = parseInt(document.getElementById('produtoDistDet').value);
    const quantidade = parseInt(document.getElementById('quantidadeDistDet').value);
    const data = document.getElementById('dataDistDet').value || new Date().toISOString().split('T')[0];
    const observacoes = document.getElementById('observacoesDistDet').value.trim();
    
    const produto = estoque.produtos.find(p => p.id === produtoId);
    
    if (!produto) {
        mostrarNotificacao('Produto não encontrado!', 'error');
        return;
    }
    
    // Verificar estoque disponível na IMBEL
    const estoqueIMBEL = (produto.distribuicao.IMBEL || 0) - (produto.vendas.IMBEL || 0);
    if (quantidade > estoqueIMBEL) {
        mostrarNotificacao(`Estoque insuficiente na IMBEL! Disponível: ${estoqueIMBEL} unidades`, 'error');
        return;
    }
    
    // Criar registro da distribuição
    const novaDistribuicao = {
        id: Date.now(),
        representante: representante,
        produtoId: produtoId,
        produtoNome: produto.nome,
        quantidade: quantidade,
        data: data,
        observacoes: observacoes
    };
    
    // Retirar da IMBEL e adicionar ao representante
    produto.distribuicao.IMBEL = (produto.distribuicao.IMBEL || 0) - quantidade;
    produto.distribuicao[representante] = (produto.distribuicao[representante] || 0) + quantidade;
    
    // Adicionar ao registro
    estoque.registroDistribuicao.push(novaDistribuicao);
    
    salvarDados();
    renderizarTabela();
    renderizarDashboard();
    renderizarRegistroDistribuicao();
    fecharModal('modalNovaDistribuicao');
    
    mostrarNotificacao(`Distribuição registrada: ${quantidade}x "${produto.nome}" para ${representante}`, 'success');
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
        produto.distribuicao.IMBEL = (produto.distribuicao.IMBEL || 0) + dist.quantidade;
    }
    
    // Remover do registro
    estoque.registroDistribuicao = estoque.registroDistribuicao.filter(d => d.id !== distId);
    
    salvarDados();
    renderizarTabela();
    renderizarDashboard();
    renderizarRegistroDistribuicao();
    
    mostrarNotificacao(`Distribuição excluída! ${dist.quantidade} unidades devolvidas à IMBEL.`, 'success');
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
                
                // Verificar estoque disponível na IMBEL
                const estoqueIMBEL = (produto.distribuicao.IMBEL || 0) - (produto.vendas.IMBEL || 0);
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
                
                // Retirar da IMBEL e adicionar ao representante
                produto.distribuicao.IMBEL = (produto.distribuicao.IMBEL || 0) - quantidade;
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
                console.log('Erros de importação:', erros);
            } else if (erros.length > 0) {
                mostrarNotificacao(`${distribuicoesImportadas} distribuições importadas. ${erros.length} linhas com erro.`, 'warning');
                console.log('Erros de importação:', erros);
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
    const estoqueTotal = parseInt(document.getElementById('estoqueTotal').value);
    
    if (estoque.produtos.some(p => p.nome === nome)) {
        mostrarNotificacao('Este produto já existe no sistema!', 'error');
        return;
    }
    
    const novoProduto = {
        id: Date.now(),
        nome: nome,
        distribuicao: { KOLTE: 0, ISA: 0, LC: 0, ADES: 0, FL: 0, IMBEL: estoqueTotal },
        vendas: { KOLTE: 0, ISA: 0, LC: 0, ADES: 0, FL: 0, IMBEL: 0 }
    };
    
    estoque.produtos.push(novoProduto);
    // Atualizar selects imediatamente para refletir o novo produto em qualquer modal aberto
    atualizarSelectsProdutos();
    salvarDados();
    renderizarTabela();
    renderizarDashboard();
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
    
    const saldoIMBEL = produto.distribuicao.IMBEL - produto.vendas.IMBEL;
    
    if (quantidade > saldoIMBEL) {
        mostrarNotificacao(`Estoque insuficiente na IMBEL! Saldo disponível: ${saldoIMBEL} unidades`, 'error');
        return;
    }
    
    produto.distribuicao.IMBEL -= quantidade;
    produto.distribuicao[representante] = (produto.distribuicao[representante] || 0) + quantidade;
    
    salvarDados();
    renderizarTabela();
    renderizarDashboard();
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
    
    const disp = produto.distribuicao[vendedor] || 0;
    const vendido = produto.vendas[vendedor] || 0;
    const saldo = disp - vendido;
    
    if (quantidade > saldo) {
        mostrarNotificacao(`Estoque insuficiente! ${vendedor} possui apenas ${saldo} unidades disponíveis.`, 'error');
        return;
    }
    
    produto.vendas[vendedor] = (produto.vendas[vendedor] || 0) + quantidade;
    
    salvarDados();
    renderizarTabela();
    renderizarDashboard();
    fecharModal('modalVenda');
    
    mostrarNotificacao(`Venda registrada: ${quantidade}x "${produto.nome}"`, 'success');
}

function salvarDevolucao(event) {
    event.preventDefault();
    
    const produtoId = parseInt(document.getElementById('produtoDevolucao').value);
    const representante = document.getElementById('representanteDevolucao').value;
    const quantidade = parseInt(document.getElementById('quantidadeDevolucao').value);
    const destino = document.getElementById('destinoDevolucao') ? document.getElementById('destinoDevolucao').value : 'IMBEL';
    
    const produto = estoque.produtos.find(p => p.id === produtoId);
    
    if (!produto) {
        mostrarNotificacao('Produto não encontrado!', 'error');
        return;
    }
    
    const disp = produto.distribuicao[representante] || 0;
    const vendido = produto.vendas[representante] || 0;
    const saldo = disp - vendido;
    
    if (quantidade > saldo) {
        mostrarNotificacao(`Quantidade inválida! ${representante} possui apenas ${saldo} unidades em saldo.`, 'error');
        return;
    }
    
    if (destino === representante) {
        mostrarNotificacao('Destino selecionado é o mesmo representante de origem. Selecione outro destino.', 'error');
        return;
    }

    // Subtrair do representante de origem
    produto.distribuicao[representante] = (produto.distribuicao[representante] || 0) - quantidade;
    if (produto.distribuicao[representante] < 0) produto.distribuicao[representante] = 0;

    // Garantir chave de destino
    produto.distribuicao[destino] = (produto.distribuicao[destino] || 0) + quantidade;

    salvarDados();
    renderizarTabela();
    renderizarDashboard();
    atualizarSelectsProdutos();

    fecharModal('modalDevolucao');

    mostrarNotificacao(`${quantidade} unidades movidas de ${representante} para ${destino}!`, 'success');
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
        // Calcular quantidade total em estoque (saldo de todos os representantes + IMBEL)
        let totalEstoque = 0;
        estoque.representantes.forEach(rep => {
            const disp = produto.distribuicao[rep] || 0;
            const venda = produto.vendas[rep] || 0;
            if (rep === 'IMBEL') {
                totalEstoque += disp; // IMBEL é estoque direto
            } else {
                totalEstoque += (disp - venda); // Saldo = Disp - Venda
            }
        });
        
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
        const dataStr = JSON.stringify(estoque, null, 2);
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
            salvarDados();

            // Re-renderizar tudo
            renderizarTabela();
            renderizarDashboard();
            renderizarRegistroVendas();
            renderizarRegistroDistribuicao();
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
            console.log('Cabeçalho:', cabecalho);
            
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
                    produto.distribuicao.IMBEL = quantidade;
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
                console.log('Erros de importação:', erros);
            } else if (erros.length > 0) {
                mostrarNotificacao(`${produtosAtualizados} produtos atualizados. ${erros.length} erros.`, 'warning');
                console.log('Erros de importação:', erros);
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
            console.log('Conteúdo do arquivo:', conteudo); // Debug
            
            const linhas = conteudo.split(/\r?\n/).filter(l => l.trim());
            
            console.log('Linhas encontradas:', linhas.length); // Debug
            console.log('Primeira linha:', linhas[0]); // Debug
            if (linhas[1]) console.log('Segunda linha:', linhas[1]); // Debug
            
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
                
                console.log(`Linha ${i + 1} - Colunas:`, colunas); // Debug
                
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
                
                console.log(`Dados: contrato=${contrato}, loja=${loja}, rep=${representante}, produto=${produtoNome}, qtd=${quantidade}`); // Debug
                
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
                console.log('Erros de importação:', erros);
            } else if (erros.length > 0) {
                mostrarNotificacao(`${vendasImportadas} vendas importadas. ${erros.length} linhas com erro.`, 'warning');
                console.log('Erros de importação:', erros);
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
    
    salvarDados();
    renderizarTabela();
    renderizarDashboard();
    renderizarRegistroVendas();
    renderizarRegistroDistribuicao();
    renderizarControleEnvio();
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
        const repClass = (contrato.representante || '').toLowerCase();

        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td class="col-contrato">${contrato.contrato}</td>
            <td class="col-loja" title="${contrato.loja}">${contrato.loja}</td>
            <td class="col-representante"><span class="badge-rep ${repClass}">${contrato.representante}</span></td>
            <td class="col-sistema">
                <input type="checkbox" class="checkbox-campo" ${sistemaMarcado ? 'checked' : ''} onchange="salvarControleEnvio('${contrato.contrato}', 'sistema', this.checked)">
            </td>
            <td class="col-assinado">
                <input type="checkbox" class="checkbox-campo" ${envio.assinado ? 'checked' : ''} onchange="salvarControleEnvio('${contrato.contrato}', 'assinado', this.checked)">
            </td>
            <td class="col-enviado">
                <input type="checkbox" class="checkbox-campo" ${envio.enviado ? 'checked' : ''} onchange="salvarControleEnvio('${contrato.contrato}', 'enviado', this.checked)">
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
// MENU HAMBÚRGUER (MOBILE)
// ========================================

function toggleMenuMobile() {
    const btn = document.getElementById('hamburgerBtn');
    const nav = document.getElementById('tabsNav');
    btn.classList.toggle('active');
    nav.classList.toggle('mobile-open');
}

// Fechar menu ao trocar aba (mobile)
const _trocarAbaOriginal = trocarAba;
trocarAba = function(aba) {
    _trocarAbaOriginal(aba);
    try {
        document.getElementById('hamburgerBtn')?.classList.remove('active');
        document.getElementById('tabsNav')?.classList.remove('mobile-open');
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
    const alertas = [];
    estoque.produtos.forEach(produto => {
        let totalDisp = 0, totalVenda = 0;
        estoque.representantes.forEach(rep => {
            totalDisp += (produto.distribuicao[rep] || 0);
            totalVenda += (produto.vendas[rep] || 0);
        });
        const saldo = totalDisp - totalVenda;
        if (saldo > 0 && saldo <= LIMITE_ESTOQUE_BAIXO) {
            alertas.push({ nome: produto.nome, saldo: saldo });
        }
    });

    const el = document.getElementById('alertaEstoqueBaixo');
    if (!el) return;

    if (alertas.length === 0) {
        el.style.display = 'none';
        return;
    }

    let html = '<div class="alerta-titulo"><span>⚠️ Estoque Baixo</span><span class="alerta-close" onclick="this.closest(\'.alerta-estoque-baixo\').style.display=\'none\'">✕</span></div>';
    alertas.forEach(a => {
        html += `<div class="alerta-item"><span>${a.nome}</span><strong>${a.saldo} un.</strong></div>`;
    });
    el.innerHTML = html;
    el.style.display = 'block';
}

// ========================================
// DASHBOARD COM GRÁFICOS (Chart.js)
// ========================================

let _chartVendasRep = null;
let _chartTopProdutos = null;

function renderizarGraficos() {
    if (typeof Chart === 'undefined') return;

    // Dados por representante
    const reps = ['KOLTE', 'ISA', 'LC', 'ADES', 'FL', 'IMBEL'];
    const coresReps = ['#3d5a80', '#5c4d7d', '#2d6a4f', '#9c4a1a', '#7b2d26', '#1e3a5f'];
    const vendasPorRep = reps.map(rep => {
        let total = 0;
        estoque.produtos.forEach(p => { total += (p.vendas[rep] || 0); });
        return total;
    });

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
                    backgroundColor: coresReps.map(c => c + 'cc'),
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
    const dadosProdutos = estoque.produtos.map(p => {
        let total = 0;
        estoque.representantes.forEach(r => { total += (p.vendas[r] || 0); });
        return { nome: p.nome.substring(0, 20), total };
    }).filter(p => p.total > 0).sort((a, b) => b.total - a.total).slice(0, 6);

    const ctx2 = document.getElementById('chartTopProdutos');
    if (ctx2) {
        if (_chartTopProdutos) _chartTopProdutos.destroy();
        const palette = ['#3b82f6', '#22c55e', '#f59e0b', '#ef4444', '#8b5cf6', '#06b6d4'];
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

    if (filtroRep) distribuicoes = distribuicoes.filter(d => d.representante === filtroRep);

    // Filtrar por data (comparação por DATA YYYY-MM-DD para evitar timezone/formato)
    if ((dataInicio && dataInicio !== '') || (dataFim && dataFim !== '')) {
        distribuicoes = distribuicoes.filter(d => {
            if (!d.data) return false;
            const registroDateStr = parseDateToYYYYMMDD(d.data);
            if (!registroDateStr) return false;
            if (dataInicio && dataInicio !== '' && registroDateStr < dataInicio) return false;
            if (dataFim && dataFim !== '' && registroDateStr > dataFim) return false;
            return true;
        });
    }

    distribuicoes.sort((a, b) => {
        const da = parseDateToYYYYMMDD(a.data);
        const db = parseDateToYYYYMMDD(b.data);
        const ta = da ? new Date(da).getTime() : 0;
        const tb = db ? new Date(db).getTime() : 0;
        return tb - ta;
    });

    // Agrupar por representante
    const porRep = {};
    distribuicoes.forEach(d => {
        if (!porRep[d.representante]) porRep[d.representante] = [];
        porRep[d.representante].push(d);
    });

    const container = document.createElement('div');
    container.className = 'report-distribuicao';

    let totalGeral = 0;

    Object.keys(porRep).sort().forEach(rep => {
        const items = porRep[rep];
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
        let subtotal = 0;
        items.forEach(d => {
            subtotal += d.quantidade;
            totalGeral += d.quantidade;
            const parsed = parseDateToYYYYMMDD(d.data);
            const dataFmt = parsed ? new Date(parsed).toLocaleDateString('pt-BR') : '-';
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td style="padding:6px;border:1px solid #ddd">${d.produtoNome}</td>
                <td style="padding:6px;border:1px solid #ddd;text-align:center">${d.quantidade}</td>
                <td style="padding:6px;border:1px solid #ddd;text-align:center">${dataFmt}</td>
                <td style="padding:6px;border:1px solid #ddd">${d.observacoes || '-'}</td>`;
            tbody.appendChild(tr);
        });

        const trTotal = document.createElement('tr');
        trTotal.innerHTML = `<td colspan="1" style="padding:6px;border:1px solid #ddd;text-align:right"><strong>Subtotal ${rep}</strong></td>
            <td style="padding:6px;border:1px solid #ddd;text-align:center"><strong>${subtotal}</strong></td>
            <td colspan="2" style="padding:6px;border:1px solid #ddd"></td>`;
        tbody.appendChild(trTotal);
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
        const data = (estoque.registroDistribuicao || []).map(d => {
            return [d.representante, d.produtoNome, d.quantidade.toString(), d.data ? new Date(d.data+'T00:00:00').toLocaleDateString('pt-BR') : '-', d.observacoes || '-'];
        });
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

    let distribuicoesFiltradas = [...(estoque.registroDistribuicao || [])];

    if (filtroRep) distribuicoesFiltradas = distribuicoesFiltradas.filter(d => d.representante === filtroRep);
    if (filtroProduto) distribuicoesFiltradas = distribuicoesFiltradas.filter(d => d.produtoId === parseInt(filtroProduto));

    // Filtro por data
    if (dataInicio || dataFim) {
        const start = dataInicio ? new Date(dataInicio + 'T00:00:00').getTime() : null;
        const end = dataFim ? new Date(dataFim + 'T23:59:59').getTime() : null;
        distribuicoesFiltradas = distribuicoesFiltradas.filter(d => {
            if (!d.data) return false;
            const t = new Date(d.data).getTime();
            if (start && t < start) return false;
            if (end && t > end) return false;
            return true;
        });
    }

    // Ordenação
    const campo = _ordenDistribuicao.campo;
    const dir = _ordenDistribuicao.direcao === 'asc' ? 1 : -1;
    distribuicoesFiltradas.sort((a, b) => {
        let va, vb;
        if (campo === 'representante') { va = a.representante || ''; vb = b.representante || ''; }
        else if (campo === 'produtoNome') { va = a.produtoNome || ''; vb = b.produtoNome || ''; }
        else if (campo === 'quantidade') { va = a.quantidade || 0; vb = b.quantidade || 0; }
        else if (campo === 'data') { va = a.data || ''; vb = b.data || ''; }
        else { va = a.data || ''; vb = b.data || ''; }
        if (va < vb) return -1 * dir;
        if (va > vb) return 1 * dir;
        return 0;
    });

    tbody.innerHTML = '';

    if (distribuicoesFiltradas.length === 0) {
        tbody.innerHTML = `<tr><td colspan="7" class="empty-state"><div class="empty-icon">🚚</div><div class="empty-text">Nenhuma distribuição registrada</div><div class="empty-hint">Clique em "Nova Distribuição"</div></td></tr>`;
        atualizarTotaisDistribuicao(0);
        renderizarPaginacao('paginacaoDistribuicao', 1, 0, _itensPorPaginaDistribuicao, 'mudarPaginaDistribuicao', 'mudarItensPaginaDistribuicao');
        return;
    }

    const totalLinhas = distribuicoesFiltradas.length;
    const totalPaginas = Math.max(1, Math.ceil(totalLinhas / _itensPorPaginaDistribuicao));
    if (_paginaDistribuicao > totalPaginas) _paginaDistribuicao = totalPaginas;
    const inicio = (_paginaDistribuicao - 1) * _itensPorPaginaDistribuicao;
    const pagina = distribuicoesFiltradas.slice(inicio, inicio + _itensPorPaginaDistribuicao);

    let totalQtd = 0;
    distribuicoesFiltradas.forEach(d => { totalQtd += d.quantidade; });

    let numero = totalLinhas - inicio;

    pagina.forEach(dist => {
        const repClass = dist.representante.toLowerCase();
        const dataFormatada = dist.data ? new Date(dist.data + 'T00:00:00').toLocaleDateString('pt-BR') : '-';
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td class="col-contrato">${numero--}</td>
            <td class="col-loja"><span class="badge-rep ${repClass}">${dist.representante}</span></td>
            <td class="col-produto-venda" title="${dist.produtoNome}">${dist.produtoNome}</td>
            <td class="col-qtd">${dist.quantidade}</td>
            <td>${dataFormatada}</td>
            <td class="col-obs" title="${dist.observacoes||'-'}">${dist.observacoes||'-'}</td>
            <td class="col-acoes">
                <button class="btn-action btn-delete" onclick="excluirDistribuicao(${dist.id})" title="Excluir">🗑</button>
            </td>`;
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
    try { verificarEstoqueBaixo(); } catch(e) {}
};

// Hook salvarDados to check low stock and register audit
const _salvarDadosHook = salvarDados;
// We can't easily re-declare salvarDados since it's used everywhere,
// but we do check low stock after renders
const _renderizarTabelaOriginal = renderizarTabela;
renderizarTabela = function() {
    _renderizarTabelaOriginal();
    try { verificarEstoqueBaixo(); } catch(e) {}
};

// Hooks for audit log on key operations
const _salvarNovaDistribuicaoOriginal = salvarNovaDistribuicao;
salvarNovaDistribuicao = function(event) {
    _salvarNovaDistribuicaoOriginal(event);
    try {
        const rep = document.getElementById('representanteDistDet')?.value || '';
        const prod = document.getElementById('produtoDistDet')?.selectedOptions[0]?.text || '';
        const qtd = document.getElementById('quantidadeDistDet')?.value || '';
        registrarHistorico('distribuicao', `${qtd}x ${prod} → ${rep}`);
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

// Atualiza UI conforme estado de autenticação
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
