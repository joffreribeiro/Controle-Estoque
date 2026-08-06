/**
 * estoque-calculos.js - Cálculos puros do módulo Operação/Estoque do Controle-Estoque
 * Sem DOM, sem estado global — testável isoladamente com Vitest.
 *
 * O módulo Operação ainda não tem uma separação lógica/renderização como o
 * CRM (crm-calculos.js) ou o Ponto (ponto-calculos.js); a maior parte do
 * código vive em app2.js, que acessa `document`/`window` no escopo de
 * carregamento e por isso não pode ser importado com segurança pelo Vitest
 * (ambiente `node`, sem jsdom — ver vitest.config.js). Este arquivo existe
 * para abrigar funções puras do módulo Operação, seguindo o mesmo padrão dos
 * demais módulos, à medida que forem extraídas de app2.js.
 */

/**
 * Situação do custo cadastrado de um produto, para o Relatório de
 * Rentabilidade: 'sem_custo' quando o custo total é nulo, indefinido ou zero
 * (a margem calculada nesse caso não é confiável — vira ~100% por ausência de
 * dado, não por resultado real); 'ok' quando há custo cadastrado maior que zero.
 */
function calcularStatusCusto(produto) {
    const custo = produto && produto.custoTotal;
    const num = Number(custo);
    return (Number.isFinite(num) && num > 0) ? 'ok' : 'sem_custo';
}

const EstoqueCalculos = {
    calcularStatusCusto
};

if (typeof window !== 'undefined') {
    window.EstoqueCalculos = EstoqueCalculos;
}

if (typeof module !== 'undefined' && module.exports) {
    module.exports = EstoqueCalculos;
}
