/**
 * crm-kanban.js - Renderização do board Kanban do CRM (Controle-Estoque)
 * Depende de: CrmModel, CrmCalculos, CrmStore, Utils.escapeHtml (crm-compat.js)
 */
(function () {
    function esc(v) { return window.Utils.escapeHtml(v); }

    function nomeCliente(negocio) {
        if (!negocio.clienteId) return '';
        var c = CrmStore.getCliente(negocio.clienteId);
        return c ? c.nome : '';
    }

    function iniciais(nome) {
        var partes = String(nome || '').trim().split(/\s+/).filter(Boolean);
        if (!partes.length) return '?';
        return (partes[0][0] + (partes[1] ? partes[1][0] : '')).toUpperCase();
    }

    function renderizarCard(n, mostrarValor, atividades) {
        var semAtividade = !CrmCalculos.temAtividadePendente(atividades || [], n.id);
        var cliente = nomeCliente(n);
        var qtdItens = (n.itens || []).length;

        return '' +
            '<div class="crm-card" draggable="true" data-crm-negocio-id="' + esc(n.id) + '" data-crm-action="abrirDetalhe" data-id="' + esc(n.id) + '">' +
                '<div class="crm-card-topo">' +
                    '<span class="crm-card-titulo">' + esc(n.titulo || '(sem título)') + '</span>' +
                    (semAtividade ? '<span class="crm-alerta-atividade" title="Sem atividade agendada">⚠️</span>' : '') +
                '</div>' +
                (cliente ? '<div class="crm-card-vinculos">' + esc(cliente) + '</div>' : '') +
                (n.origem ? '<div class="crm-card-vinculos">' + esc(n.origem) + '</div>' : '') +
                (qtdItens ? '<div class="crm-card-vinculos">' + qtdItens + ' item' + (qtdItens > 1 ? 's' : '') + '</div>' : '') +
                '<div class="crm-card-rodape">' +
                    (n.responsavel ? '<span class="crm-card-avatar" title="' + esc(n.responsavel) + '">' + esc(iniciais(n.responsavel)) + '</span>' : '') +
                    (mostrarValor && n.valor !== null && n.valor !== undefined
                        ? '<span class="crm-card-valor">' + esc(CrmCalculos.formatarMoeda(n.valor, n.moeda)) + '</span>' : '') +
                '</div>' +
            '</div>';
    }

    function renderizarBoard(funil, negocios, opcoes) {
        var op = opcoes || {};
        var atividades = op.atividades || [];
        var agrupado = CrmCalculos.agruparPorEtapa(negocios, funil.etapas);

        var colunas = funil.etapas.map(function (etapa) {
            var itens = agrupado[etapa.id] || [];
            var soma = funil.mostrarValor ? CrmCalculos.somarValor(itens) : null;
            var subtitulo = funil.mostrarValor
                ? (esc(CrmCalculos.formatarMoeda(soma)) + ' · ' + itens.length + ' negócio' + (itens.length !== 1 ? 's' : ''))
                : (itens.length + ' negócio' + (itens.length !== 1 ? 's' : ''));

            var cardsHtml = itens.map(function (n) { return renderizarCard(n, funil.mostrarValor, atividades); }).join('');

            return '' +
                '<div class="crm-coluna" data-crm-etapa-id="' + esc(etapa.id) + '">' +
                    '<div class="crm-coluna-header">' +
                        '<div class="crm-coluna-nome">' + esc(etapa.nome) + '</div>' +
                        '<div class="crm-col-soma">' + subtitulo + '</div>' +
                    '</div>' +
                    '<div class="crm-coluna-corpo kanban-list" data-crm-etapa-id="' + esc(etapa.id) + '">' +
                        (cardsHtml || '<div class="crm-coluna-vazia">Nenhum negócio</div>') +
                    '</div>' +
                '</div>';
        }).join('');

        return '<div class="crm-board">' + colunas + '</div>';
    }

    window.CrmKanban = {
        renderizarBoard: renderizarBoard,
        renderizarCard: renderizarCard,
        nomeCliente: nomeCliente
    };
})();
