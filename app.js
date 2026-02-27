// ========================================
// SISTEMA DE CONTROLE DE ESTOQUE
// Material Bélico - v2.0
// ========================================

// Estrutura de dados principal
let estoque = {
    produtos: [],
    representantes: ['KOLTE', 'ISA', 'LC', 'ADES', 'FL', 'IMBEL'],
    registroVendas: []
};

// Dados iniciais com PREÇOS baseados na planilha
const dadosIniciais = [
    {
        nome: 'CARABINA IA2 5,56',
        preco: 10420.75, // 125.049,04 / 12 = ~10.420,75
        distribuicao: { KOLTE: 20, ISA: 20, LC: 20, ADES: 20, FL: 20, IMBEL: 241 },
        vendas: { KOLTE: 0, ISA: 8, LC: 0, ADES: 4, FL: 0, IMBEL: 0 }
    },
    {
        nome: 'CARABINA IA2 7,62',
        preco: 12690.21, // 101.521,70 / 8 = ~12.690,21
        distribuicao: { KOLTE: 4, ISA: 4, LC: 4, ADES: 4, FL: 4, IMBEL: 1 },
        vendas: { KOLTE: 2, ISA: 2, LC: 0, ADES: 4, FL: 0, IMBEL: 1 }
    },
    {
        nome: 'FACA CAMPANHA AMZ',
        preco: 360.00, // 74.880,00 / 208 = 360,00
        distribuicao: { KOLTE: 100, ISA: 100, LC: 100, ADES: 100, FL: 100, IMBEL: 525 },
        vendas: { KOLTE: 100, ISA: 100, LC: 0, ADES: 0, FL: 0, IMBEL: 8 }
    },
    {
        nome: 'FACA POLICIAL AMZ',
        preco: 352.45, // 18.680,00 / 53 = ~352,45
        distribuicao: { KOLTE: 14, ISA: 20, LC: 0, ADES: 0, FL: 0, IMBEL: 8 },
        vendas: { KOLTE: 20, ISA: 20, LC: 0, ADES: 5, FL: 0, IMBEL: 8 }
    },
    {
        nome: 'FACA POLICIAL IA2',
        preco: 380.00,
        distribuicao: { KOLTE: 20, ISA: 0, LC: 20, ADES: 0, FL: 20, IMBEL: 0 },
        vendas: { KOLTE: 0, ISA: 0, LC: 0, ADES: 0, FL: 0, IMBEL: 0 }
    },
    {
        nome: 'FUZIL DE ALTA PRECISÃO IMBEL 308 AGLC (COMPLETO)',
        preco: 13500.00, // 27.000,00 / 2 = 13.500,00
        distribuicao: { KOLTE: 0, ISA: 0, LC: 0, ADES: 0, FL: 0, IMBEL: 2 },
        vendas: { KOLTE: 0, ISA: 0, LC: 0, ADES: 0, FL: 0, IMBEL: 2 }
    },
    {
        nome: 'PISTOLA .40 GC MD7 C/ ADC',
        preco: 5159.71, // 30.958,26 / 6 = ~5.159,71
        distribuicao: { KOLTE: 2, ISA: 0, LC: 0, ADES: 4, FL: 0, IMBEL: 193 },
        vendas: { KOLTE: 2, ISA: 0, LC: 0, ADES: 4, FL: 0, IMBEL: 0 }
    },
    {
        nome: 'PISTOLA 380 GC MD1 C/ ADC',
        preco: 5219.54, // 46.975,85 / 9 = ~5.219,54
        distribuicao: { KOLTE: 8, ISA: 8, LC: 8, ADES: 8, FL: 8, IMBEL: 10 },
        vendas: { KOLTE: 2, ISA: 2, LC: 0, ADES: 5, FL: 0, IMBEL: 0 }
    },
    {
        nome: 'PISTOLA 380 GC MD1 S/ ADC',
        preco: 5406.19, // 16.218,56 / 3 = ~5.406,19
        distribuicao: { KOLTE: 8, ISA: 8, LC: 8, ADES: 8, FL: 8, IMBEL: 13 },
        vendas: { KOLTE: 0, ISA: 3, LC: 0, ADES: 0, FL: 0, IMBEL: 0 }
    },
    {
        nome: 'PISTOLA 380 GC MD2 C/ ADC',
        preco: 5162.57, // 56.788,27 / 11 = ~5.162,57
        distribuicao: { KOLTE: 0, ISA: 0, LC: 0, ADES: 5, FL: 0, IMBEL: 60 },
        vendas: { KOLTE: 0, ISA: 0, LC: 0, ADES: 5, FL: 0, IMBEL: 6 }
    },
    {
        nome: 'PISTOLA 380 GC MD2 S/ ADC',
        preco: 5207.87, // 692.646,71 / 133 = ~5.207,87
        distribuicao: { KOLTE: 50, ISA: 50, LC: 50, ADES: 50, FL: 50, IMBEL: 50 },
        vendas: { KOLTE: 38, ISA: 45, LC: 0, ADES: 40, FL: 0, IMBEL: 10 }
    },
    {
        nome: 'PISTOLA 9 GC MD1 S/ ADC',
        preco: 5236.30, // 141.379,97 / 27 = ~5.236,30
        distribuicao: { KOLTE: 30, ISA: 30, LC: 30, ADES: 30, FL: 30, IMBEL: 46 },
        vendas: { KOLTE: 2, ISA: 16, LC: 0, ADES: 9, FL: 0, IMBEL: 0 }
    }
];

// ========================================
// FUNÇÕES DE INICIALIZAÇÃO
// ========================================

function inicializar() {
    carregarDados();
    renderizarTabela();
    renderizarDashboard();
    renderizarRegistroVendas();
    atualizarSelectsProdutos();
    atualizarEstatisticas();
    atualizarData();
}

function carregarDados() {
    const dadosSalvos = localStorage.getItem('estoqueArmasV2');
    if (dadosSalvos) {
        estoque = JSON.parse(dadosSalvos);
        // Garantir que registroVendas existe
        if (!estoque.registroVendas) {
            estoque.registroVendas = [];
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
        salvarDados();
    }
}

function salvarDados() {
    localStorage.setItem('estoqueArmasV2', JSON.stringify(estoque));
    atualizarEstatisticas();
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
            valorTotalVendas += venda * (produto.preco || 0);
        });
    });

    document.getElementById('totalProdutos').textContent = estoque.produtos.length;
    document.getElementById('totalEstoque').textContent = totalEstoque.toLocaleString('pt-BR');
    document.getElementById('totalVendas').textContent = totalVendas.toLocaleString('pt-BR');
    document.getElementById('valorTotalVendas').textContent = formatarMoedaValor(valorTotalVendas);
}

// ========================================
// NAVEGAÇÃO POR ABAS
// ========================================

function trocarAba(aba) {
    // Atualizar botões
    document.querySelectorAll('.tab-btn').forEach(btn => btn.classList.remove('active'));
    event.target.closest('.tab-btn').classList.add('active');
    
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

    estoque.produtos.forEach(produto => {
        const tr = document.createElement('tr');
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
            
            tr.innerHTML += `
                <td class="cell-disp ${disp === 0 ? 'cell-zero' : ''}">${formatarNumero(disp)}</td>
                <td class="cell-venda ${venda === 0 ? 'cell-zero' : ''}">${formatarNumero(venda)}</td>
                <td class="cell-saldo ${saldoClass} ${saldo === 0 ? 'cell-zero' : ''}">${formatarNumero(saldo)}</td>
            `;
        });

        const geralSaldo = geralDisp - geralVenda;
        totais.GERAL.disp += geralDisp;
        totais.GERAL.venda += geralVenda;
        totais.GERAL.saldo += geralSaldo;

        const saldoGeralClass = geralSaldo < 0 ? 'negativo' : (geralSaldo > 0 && geralSaldo <= 10 ? 'baixo' : '');

        tr.innerHTML += `
            <td class="geral-disp">${formatarNumero(geralDisp)}</td>
            <td class="geral-venda">${formatarNumero(geralVenda)}</td>
            <td class="geral-saldo ${saldoGeralClass}">${formatarNumero(geralSaldo)}</td>
        `;

        tbody.appendChild(tr);
    });

    // Linha de totais
    const trTotal = document.createElement('tr');
    trTotal.className = 'total-row';
    trTotal.innerHTML = `<td class="produto-nome col-produto"><strong>TOTAL GERAL</strong></td>`;

    estoque.representantes.forEach(rep => {
        trTotal.innerHTML += `
            <td class="cell-disp"><strong>${formatarNumero(totais[rep].disp)}</strong></td>
            <td class="cell-venda"><strong>${formatarNumero(totais[rep].venda)}</strong></td>
            <td class="cell-saldo"><strong>${formatarNumero(totais[rep].saldo)}</strong></td>
        `;
    });

    trTotal.innerHTML += `
        <td class="geral-disp"><strong>${formatarNumero(totais.GERAL.disp)}</strong></td>
        <td class="geral-venda"><strong>${formatarNumero(totais.GERAL.venda)}</strong></td>
        <td class="geral-saldo"><strong>${formatarNumero(totais.GERAL.saldo)}</strong></td>
    `;

    tbody.appendChild(trTotal);
}

// ========================================
// RENDERIZAÇÃO DO DASHBOARD
// ========================================

function renderizarDashboard() {
    // Calcular dados
    let dadosVendas = [];
    let vendasPorRep = { ADES: 0, FL: 0, IMBEL: 0, ISA: 0, KOLTE: 0, LC: 0 };
    let totalUnidades = 0;
    let totalFaturamento = 0;

    estoque.produtos.forEach(produto => {
        let totalProduto = 0;
        let vendasProdutoPorRep = {};
        
        estoque.representantes.forEach(rep => {
            const venda = produto.vendas[rep] || 0;
            totalProduto += venda;
            vendasPorRep[rep] = (vendasPorRep[rep] || 0) + venda;
            vendasProdutoPorRep[rep] = venda;
        });

        const valorTotal = totalProduto * (produto.preco || 0);
        
        dadosVendas.push({
            nome: produto.nome,
            quantidade: totalProduto,
            valor: valorTotal,
            preco: produto.preco || 0,
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

    // Atualizar cards
    document.getElementById('dashTotalUnidades').textContent = totalUnidades.toLocaleString('pt-BR');
    document.getElementById('dashTotalFaturamento').textContent = formatarMoedaValor(totalFaturamento);
    document.getElementById('dashMelhorRep').textContent = melhorRep + ` (${maxVendas})`;
    document.getElementById('dashProdutoTop').textContent = produtoTop;

    // Tabela: Quantidade por Produto
    const tabelaQtd = document.getElementById('tabelaQtdProduto');
    tabelaQtd.innerHTML = '';
    
    dadosVendas.forEach(item => {
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
    
    // Ordenar por valor
    const dadosPorValor = [...dadosVendas].sort((a, b) => b.valor - a.valor);
    
    dadosPorValor.forEach(item => {
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

    dadosVendas.forEach(item => {
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
    document.getElementById('modalProduto').style.display = 'block';
    document.getElementById('formProduto').reset();
}

function abrirModalDistribuicao() {
    document.getElementById('modalDistribuicao').style.display = 'block';
    document.getElementById('formDistribuicao').reset();
    atualizarSelectsProdutos();
}

function abrirModalVenda() {
    document.getElementById('modalVenda').style.display = 'block';
    document.getElementById('formVenda').reset();
    atualizarSelectsProdutos();
}

function abrirModalDevolucao() {
    document.getElementById('modalDevolucao').style.display = 'block';
    document.getElementById('formDevolucao').reset();
    atualizarSelectsProdutos();
}

function fecharModal(modalId) {
    document.getElementById(modalId).style.display = 'none';
}

window.onclick = function(event) {
    if (event.target.classList.contains('modal')) {
        event.target.style.display = 'none';
    }
}

document.addEventListener('keydown', function(event) {
    if (event.key === 'Escape') {
        document.querySelectorAll('.modal').forEach(modal => {
            modal.style.display = 'none';
        });
    }
});

// ========================================
// FUNÇÕES DE CRUD
// ========================================

function atualizarSelectsProdutos() {
    const selects = ['produtoDistribuicao', 'produtoVenda', 'produtoDevolucao', 'produtoVendaDet', 'filtroVendaProduto'];
    
    selects.forEach(selectId => {
        const select = document.getElementById(selectId);
        if (select) {
            const valorAtual = select.value;
            const primeiraOpcao = selectId === 'filtroVendaProduto' ? 'Todos os produtos' : 'Selecione um produto';
            select.innerHTML = `<option value="">${primeiraOpcao}</option>`;
            
            estoque.produtos.forEach(produto => {
                const option = document.createElement('option');
                option.value = produto.id;
                option.textContent = produto.nome;
                select.appendChild(option);
            });
            
            select.value = valorAtual;
        }
    });
}

// ========================================
// REGISTRO DE VENDAS DETALHADO
// ========================================

function abrirModalVendaDetalhada() {
    document.getElementById('modalVendaDetalhada').style.display = 'block';
    document.getElementById('formVendaDetalhada').reset();
    document.getElementById('valorUnitarioVenda').value = '';
    document.getElementById('valorTotalVenda').value = '';
    atualizarSelectsProdutos();
    
    // Sugerir próximo número de contrato
    const ultimoContrato = estoque.registroVendas.length > 0 
        ? Math.max(...estoque.registroVendas.map(v => parseInt(v.contrato) || 0)) 
        : 0;
    document.getElementById('contratoVenda').value = ultimoContrato + 1;
}

function atualizarPrecoVenda() {
    const produtoId = parseInt(document.getElementById('produtoVendaDet').value);
    const quantidade = parseInt(document.getElementById('quantidadeVendaDet').value) || 0;
    
    const produto = estoque.produtos.find(p => p.id === produtoId);
    
    if (produto && produto.preco) {
        const valorUnitario = produto.preco;
        const valorTotal = valorUnitario * quantidade;
        
        document.getElementById('valorUnitarioVenda').value = formatarMoedaValor(valorUnitario);
        document.getElementById('valorTotalVenda').value = quantidade > 0 ? formatarMoedaValor(valorTotal) : '';
    } else {
        document.getElementById('valorUnitarioVenda').value = '';
        document.getElementById('valorTotalVenda').value = '';
    }
}

function salvarVendaDetalhada(event) {
    event.preventDefault();
    
    const contrato = document.getElementById('contratoVenda').value.trim();
    const loja = document.getElementById('lojaVenda').value.trim().toUpperCase();
    const representante = document.getElementById('representanteVendaDet').value;
    const produtoId = parseInt(document.getElementById('produtoVendaDet').value);
    const quantidade = parseInt(document.getElementById('quantidadeVendaDet').value);
    const observacoes = document.getElementById('observacoesVenda').value.trim();
    
    const produto = estoque.produtos.find(p => p.id === produtoId);
    
    if (!produto) {
        mostrarNotificacao('Produto não encontrado!', 'error');
        return;
    }
    
    // Verificar estoque disponível
    const disp = produto.distribuicao[representante] || 0;
    const vendido = produto.vendas[representante] || 0;
    const saldo = disp - vendido;
    
    if (quantidade > saldo) {
        mostrarNotificacao(`Estoque insuficiente! ${representante} possui apenas ${saldo} unidades disponíveis.`, 'error');
        return;
    }
    
    const valorUnitario = produto.preco || 0;
    const valorTotal = valorUnitario * quantidade;
    
    // Criar registro da venda
    const novaVenda = {
        id: Date.now(),
        contrato: contrato,
        loja: loja,
        representante: representante,
        produtoId: produtoId,
        produtoNome: produto.nome,
        quantidade: quantidade,
        valorUnitario: valorUnitario,
        valorTotal: valorTotal,
        observacoes: observacoes,
        data: new Date().toISOString()
    };
    
    // Atualizar vendas no estoque
    produto.vendas[representante] = (produto.vendas[representante] || 0) + quantidade;
    
    // Adicionar ao registro
    estoque.registroVendas.push(novaVenda);
    
    salvarDados();
    renderizarTabela();
    renderizarDashboard();
    renderizarRegistroVendas();
    fecharModal('modalVendaDetalhada');
    
    mostrarNotificacao(`Venda registrada: Contrato ${contrato} - ${quantidade}x "${produto.nome}" - ${formatarMoedaValor(valorTotal)}`, 'success');
}

function renderizarRegistroVendas() {
    const tbody = document.getElementById('tabelaRegistroVendasBody');
    if (!tbody) return;
    
    const filtroRep = document.getElementById('filtroRepresentante')?.value || '';
    const filtroProduto = document.getElementById('filtroProduto')?.value || '';
    
    // Filtrar vendas
    let vendasFiltradas = estoque.registroVendas || [];
    
    if (filtroRep) {
        vendasFiltradas = vendasFiltradas.filter(v => v.representante === filtroRep);
    }
    
    if (filtroProduto) {
        vendasFiltradas = vendasFiltradas.filter(v => v.produtoId === parseInt(filtroProduto));
    }
    
    // Ordenar por contrato
    vendasFiltradas.sort((a, b) => {
        const contratoA = parseInt(a.contrato) || 0;
        const contratoB = parseInt(b.contrato) || 0;
        return contratoA - contratoB;
    });
    
    tbody.innerHTML = '';
    
    if (vendasFiltradas.length === 0) {
        tbody.innerHTML = `
            <tr>
                <td colspan="9" class="empty-state">
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
    
    vendasFiltradas.forEach(venda => {
        totalQtd += venda.quantidade;
        totalValor += venda.valorTotal;
        
        const repClass = venda.representante.toLowerCase();
        
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td class="col-contrato">${venda.contrato}</td>
            <td class="col-loja" title="${venda.loja}">${venda.loja}</td>
            <td class="col-representante"><span class="badge-rep ${repClass}">${venda.representante}</span></td>
            <td class="col-produto-venda" title="${venda.produtoNome}">${venda.produtoNome}</td>
            <td class="col-qtd">${venda.quantidade}</td>
            <td class="col-valor-un">${formatarMoedaValor(venda.valorUnitario)}</td>
            <td class="col-valor-total">${formatarMoedaValor(venda.valorTotal)}</td>
            <td class="col-obs" title="${venda.observacoes || '-'}">${venda.observacoes || '-'}</td>
            <td class="col-acoes">
                <button class="btn-action btn-delete" onclick="excluirVenda(${venda.id})" title="Excluir venda">🗑</button>
            </td>
        `;
        tbody.appendChild(tr);
    });
    
    atualizarTotaisVendas(totalQtd, totalValor);
}

function atualizarTotaisVendas(totalQtd, totalValor) {
    const spanQtd = document.getElementById('totalQtdVendas');
    const spanValor = document.getElementById('totalValorVendas');
    
    if (spanQtd) spanQtd.innerHTML = `<strong>${totalQtd.toLocaleString('pt-BR')}</strong>`;
    if (spanValor) spanValor.innerHTML = `<strong>${formatarMoedaValor(totalValor)}</strong>`;
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
    const venda = estoque.registroVendas.find(v => v.id === vendaId);
    
    if (!venda) {
        mostrarNotificacao('Venda não encontrada!', 'error');
        return;
    }
    
    if (!confirm(`Deseja excluir a venda do contrato ${venda.contrato}?\n\nProduto: ${venda.produtoNome}\nQuantidade: ${venda.quantidade}\nValor: ${formatarMoedaValor(venda.valorTotal)}\n\nATENÇÃO: A quantidade será devolvida ao estoque do representante.`)) {
        return;
    }
    
    // Devolver ao estoque
    const produto = estoque.produtos.find(p => p.id === venda.produtoId);
    if (produto) {
        produto.vendas[venda.representante] = Math.max(0, (produto.vendas[venda.representante] || 0) - venda.quantidade);
    }
    
    // Remover do registro
    estoque.registroVendas = estoque.registroVendas.filter(v => v.id !== vendaId);
    
    salvarDados();
    renderizarTabela();
    renderizarDashboard();
    renderizarRegistroVendas();
    
    mostrarNotificacao(`Venda do contrato ${venda.contrato} excluída com sucesso!`, 'success');
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
        vendasOrdenadas.forEach(venda => {
            const data = new Date(venda.data).toLocaleDateString('pt-BR');
            csv += `${venda.contrato}${sep}${venda.loja}${sep}${venda.representante}${sep}${venda.produtoNome}${sep}${venda.quantidade}${sep}${venda.valorUnitario.toFixed(2).replace('.', ',')}${sep}${venda.valorTotal.toFixed(2).replace('.', ',')}${sep}${venda.observacoes || ''}${sep}${data}\n`;
        });
        
        // Linha de total
        const totalQtd = vendas.reduce((sum, v) => sum + v.quantidade, 0);
        const totalValor = vendas.reduce((sum, v) => sum + v.valorTotal, 0);
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

function salvarProduto(event) {
    event.preventDefault();
    
    const nome = document.getElementById('nomeProduto').value.trim().toUpperCase();
    const estoqueTotal = parseInt(document.getElementById('estoqueTotal').value);
    const preco = converterMoedaParaNumero(document.getElementById('precoProduto').value);
    
    if (estoque.produtos.some(p => p.nome === nome)) {
        mostrarNotificacao('Este produto já existe no sistema!', 'error');
        return;
    }
    
    const novoProduto = {
        id: Date.now(),
        nome: nome,
        preco: preco,
        distribuicao: { KOLTE: 0, ISA: 0, LC: 0, ADES: 0, FL: 0, IMBEL: estoqueTotal },
        vendas: { KOLTE: 0, ISA: 0, LC: 0, ADES: 0, FL: 0, IMBEL: 0 }
    };
    
    estoque.produtos.push(novoProduto);
    salvarDados();
    renderizarTabela();
    renderizarDashboard();
    fecharModal('modalProduto');
    
    mostrarNotificacao(`Produto "${nome}" adicionado com sucesso!`, 'success');
}

function salvarDistribuicao(event) {
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
    
    const valorVenda = quantidade * (produto.preco || 0);
    mostrarNotificacao(`Venda registrada: ${quantidade}x "${produto.nome}" - ${formatarMoedaValor(valorVenda)}`, 'success');
}

function salvarDevolucao(event) {
    event.preventDefault();
    
    const produtoId = parseInt(document.getElementById('produtoDevolucao').value);
    const representante = document.getElementById('representanteDevolucao').value;
    const quantidade = parseInt(document.getElementById('quantidadeDevolucao').value);
    
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
    
    produto.distribuicao[representante] -= quantidade;
    produto.distribuicao.IMBEL += quantidade;
    
    salvarDados();
    renderizarTabela();
    renderizarDashboard();
    fecharModal('modalDevolucao');
    
    mostrarNotificacao(`${quantidade} unidades devolvidas de ${representante} para IMBEL!`, 'success');
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
    let csv = 'PRODUTO;PREÇO UNITÁRIO;';
    
    estoque.representantes.forEach(rep => {
        csv += `${rep} Disp;${rep} Venda;${rep} Saldo;`;
    });
    csv += 'GERAL Disp;GERAL Venda;GERAL Saldo;VALOR TOTAL VENDAS\n';
    
    estoque.produtos.forEach(produto => {
        csv += `"${produto.nome}";${(produto.preco || 0).toFixed(2).replace('.', ',')};`;
        
        let geralDisp = 0, geralVenda = 0;
        
        estoque.representantes.forEach(rep => {
            const disp = produto.distribuicao[rep] || 0;
            const venda = produto.vendas[rep] || 0;
            const saldo = disp - venda;
            
            geralDisp += disp;
            geralVenda += venda;
            
            csv += `${disp};${venda};${saldo};`;
        });
        
        const valorVendas = geralVenda * (produto.preco || 0);
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
                
                const valorUnitario = produto.preco || 0;
                const valorTotal = valorUnitario * quantidade;
                
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
                
                // Atualizar distribuição e vendas no estoque (para refletir que houve venda)
                produto.distribuicao[representante] = (produto.distribuicao[representante] || 0) + quantidade;
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
    if (!confirm('⚠️ ATENÇÃO!\n\nEsta ação irá APAGAR TODOS os dados do sistema:\n- Produtos\n- Distribuições\n- Vendas\n- Registro de vendas\n\nOs dados serão resetados para os valores iniciais.\n\nDeseja continuar?')) {
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
    
    salvarDados();
    renderizarTabela();
    renderizarDashboard();
    renderizarRegistroVendas();
    atualizarSelectsProdutos();
    atualizarEstatisticas();
    
    mostrarNotificacao('Todos os dados foram apagados!', 'success');
}

// ========================================
// INICIALIZAÇÃO
// ========================================

document.addEventListener('DOMContentLoaded', inicializar);
