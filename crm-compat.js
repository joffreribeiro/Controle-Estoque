/**
 * crm-compat.js - Camada de compatibilidade do CRM com o Controle-Estoque.
 *
 * Os módulos crm-kanban.js e crm-ui.js foram portados do sistema Ponto, onde
 * dependiam de `Utils.escapeHtml`, `Notifications` e `DateUtils`. O Estoque não
 * tem esses namespaces, mas tem funções equivalentes globais. Este arquivo cria
 * SOMENTE o que falta, mapeando para o que já existe no app2.js — sem sobrescrever
 * nada que porventura já esteja definido.
 *
 * Deve carregar depois de app2.js (que define mostrarNotificacao,
 * mostrarConfirmacaoEstoque e formatDateToDDMMYYYY) e antes de crm-kanban/crm-ui.
 */
(function () {
    // ── Utils.escapeHtml (o Estoque não tem; é essencial contra XSS em texto livre) ──
    if (typeof window.Utils === 'undefined') window.Utils = {};
    if (typeof window.Utils.escapeHtml !== 'function') {
        window.Utils.escapeHtml = function (valor) {
            if (valor === null || valor === undefined) return '';
            return String(valor)
                .replace(/&/g, '&amp;')
                .replace(/</g, '&lt;')
                .replace(/>/g, '&gt;')
                .replace(/"/g, '&quot;')
                .replace(/'/g, '&#39;');
        };
    }

    // ── Notifications (mapeia para as funções não-bloqueantes do Estoque) ──
    if (typeof window.Notifications === 'undefined') {
        window.Notifications = {
            error: function (msg) {
                if (typeof window.mostrarNotificacao === 'function') window.mostrarNotificacao(msg, 'error');
                else console.error('CRM:', msg);
            },
            success: function (msg) {
                if (typeof window.mostrarNotificacao === 'function') window.mostrarNotificacao(msg, 'success');
            },
            info: function (msg) {
                if (typeof window.mostrarNotificacao === 'function') window.mostrarNotificacao(msg, 'info');
            },
            // confirm(mensagem, aoConfirmar) — assinatura usada pelo crm-ui.js
            confirm: function (msg, aoConfirmar) {
                if (typeof window.mostrarConfirmacaoEstoque === 'function') {
                    window.mostrarConfirmacaoEstoque(msg, aoConfirmar);
                } else if (typeof aoConfirmar === 'function') {
                    aoConfirmar();
                }
            }
        };
    }

    // ── DateUtils.formatBR (mapeia para o formatador de data do Estoque) ──
    if (typeof window.DateUtils === 'undefined') window.DateUtils = {};
    if (typeof window.DateUtils.formatBR !== 'function') {
        window.DateUtils.formatBR = function (iso) {
            if (!iso) return '';
            if (typeof window.formatDateToDDMMYYYY === 'function') {
                try { return window.formatDateToDDMMYYYY(iso); } catch (_) { /* cai no fallback */ }
            }
            // Fallback: YYYY-MM-DD → DD/MM/YYYY
            var m = String(iso).slice(0, 10).match(/^(\d{4})-(\d{2})-(\d{2})$/);
            return m ? (m[3] + '/' + m[2] + '/' + m[1]) : String(iso);
        };
    }
})();
