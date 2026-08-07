/**
 * debounce.js - Debounce genérico para buscas/filtros (Controle-Estoque)
 *
 * Atrasa a execução de `fn` até `delayMs` depois da última chamada — evita
 * re-renderizar tabelas grandes a cada tecla digitada num campo de busca.
 * Sem dependências externas, sem build step.
 */
(function () {
    function debounce(fn, delayMs) {
        var timer = null;
        return function () {
            var args = arguments;
            var contexto = this;
            clearTimeout(timer);
            timer = setTimeout(function () {
                fn.apply(contexto, args);
            }, delayMs > 0 ? delayMs : 250);
        };
    }

    window.Debounce = { criar: debounce };
})();
