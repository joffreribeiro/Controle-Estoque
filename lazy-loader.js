/**
 * lazy-loader.js - Renderização em lotes para tabelas grandes (Controle-Estoque)
 *
 * Utilitário genérico: em vez de inserir centenas/milhares de linhas de uma vez
 * (custo de layout/paint proporcional ao total), renderiza um lote pequeno e só
 * renderiza o próximo quando o usuário rola até perto do fim — via
 * IntersectionObserver observando uma linha-sentinela.
 *
 * Não assume formato de item nem de elemento: quem chama decide o que cada
 * "item" representa e como desenhá-lo (`renderItem`), inclusive quando um único
 * item corresponde a várias linhas de tabela (ex.: produto + peças filhas).
 * Sem dependências externas, sem build step — script solto como os demais do projeto.
 */
(function () {
    /**
     * @param {Element} container - onde as linhas/elementos e a sentinela são inseridos (ex.: um <tbody>)
     * @param {Array} items - lista completa já filtrada/ordenada
     * @param {Function} renderItem - (item, indice) => void; deve inserir o(s) elemento(s) no `container`
     * @param {Object} [opts]
     * @param {number} [opts.chunkSize=50]
     * @param {string} [opts.rootMargin='300px'] - antecipa o carregamento antes da sentinela realmente aparecer
     * @param {Function} [opts.criarSentinela] - (restantes) => Element; padrão cria um <tr><td>...</td></tr>
     */
    function render(container, items, renderItem, opts) {
        opts = opts || {};
        var chunkSize = opts.chunkSize > 0 ? opts.chunkSize : 50;
        var criarSentinela = opts.criarSentinela || criarSentinelaPadrao;
        var indice = 0;
        var sentinela = null;
        var observer = null;

        function renderizarLote() {
            var fim = Math.min(indice + chunkSize, items.length);
            for (; indice < fim; indice++) {
                renderItem(items[indice], indice);
            }
            if (sentinela && sentinela.parentNode) sentinela.parentNode.removeChild(sentinela);
            sentinela = null;

            if (indice < items.length) {
                sentinela = criarSentinela(items.length - indice);
                container.appendChild(sentinela);
                if (observer) observer.observe(sentinela);
            } else if (observer) {
                observer.disconnect();
            }
        }

        if (!items.length) return;

        if (typeof IntersectionObserver === 'function') {
            observer = new IntersectionObserver(function (entries) {
                entries.forEach(function (entry) {
                    if (entry.isIntersecting) renderizarLote();
                });
            }, { rootMargin: opts.rootMargin || '300px' });
        }

        renderizarLote();

        // Navegador sem IntersectionObserver: renderiza o resto de uma vez (comportamento antigo, sem quebrar).
        if (!observer) {
            while (indice < items.length) renderizarLote();
        }
    }

    function criarSentinelaPadrao(restantes) {
        var div = document.createElement('div');
        div.className = 'lazy-sentinela';
        div.style.cssText = 'padding:14px;text-align:center;color:#94a3b8;font-size:0.8rem';
        div.textContent = 'Carregando mais ' + restantes + ' item(ns)...';
        return div;
    }

    window.LazyLoader = { render: render };
})();
