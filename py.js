// ==UserScript==
// @name         Zé Delivery - V22.0 (Turbo: Impressão Rápida)
// @match        https://seu.ze.delivery/poc-orders*
// @version      22.0
// @grant        none
// @run-at       document-start
// ==/UserScript==

(function() {
    'use strict';

    const originalPrint = window.print;
    let executandoFluxo = false;

    // Bloqueia impressão nativa
    window.print = function(forced) {
        if (forced === true) originalPrint();
    };

    // --- 1. Controle de IDs ---
    function marcarComoConcluido(id) {
        let concluidos = JSON.parse(localStorage.getItem('ze_ids_concluidos') || '[]');
        if (!concluidos.includes(id)) {
            concluidos.push(id);
            if (concluidos.length > 50) concluidos.shift();
            localStorage.setItem('ze_ids_concluidos', JSON.stringify(concluidos));
        }
    }

    function jaFoiProcessado(id) {
        let concluidos = JSON.parse(localStorage.getItem('ze_ids_concluidos') || '[]');
        return concluidos.includes(id);
    }

    // --- 2. Fecha popups ---
    function clicarOkProfundo(root = document) {
        const elementos = root.querySelectorAll('*');
        for (let el of elementos) {
            if (el.shadowRoot) clicarOkProfundo(el.shadowRoot);
            if (el.tagName === 'SPAN' && el.textContent.trim() === "Ok") {
                const btn = el.closest('button');
                if (btn && btn.offsetHeight > 0) btn.click();
            }
        }
    }

    // --- 3. Busca botão Aceitar ---
    function buscarBotaoAceitar() {
        let btnEncontrado = null;
        document.querySelectorAll('hexa-v2-button').forEach(hexa => {
            if (hexa.shadowRoot) {
                const span = hexa.shadowRoot.querySelector('[data-testid="text"]');
                const botao = hexa.shadowRoot.querySelector('button');
                if (span && span.innerText.trim() === "Aceitar" && !botao.disabled) {
                    btnEncontrado = botao;
                }
            }
        });
        return btnEncontrado;
    }

    // --- 4. Fluxo Rápido ---
    function monitorarPedidos() {
        clicarOkProfundo();
        if (executandoFluxo) return;

        const colunaNovos = document.querySelector('[data-testid="kanban-column-body-new-orders"]');
        if (!colunaNovos) return;

        const cards = Array.from(colunaNovos.querySelectorAll('a[data-testid^="link-to-order-"]'));
        const pendentes = cards.filter(c => !jaFoiProcessado(c.getAttribute('data-testid')));

        if (cards.length > 0 && pendentes.length === 0) {
            setTimeout(() => { location.reload(); }, 5000);
            executandoFluxo = true;
            return;
        }

        if (pendentes.length === 0) return;
    
        // INÍCIO
        const proximoCard = pendentes[0];
        const idPedido = proximoCard.getAttribute('data-testid');
        executandoFluxo = true;

        // 1. Abre Pedido
        proximoCard.click();

        // Espera carregar o modal (reduzido para 1.2s)
        setTimeout(() => {
            const btnAceitar = buscarBotaoAceitar();

            if (btnAceitar) {

                // ADIÇÃO: Nova lógica de detecção exata usando a tag HTML fornecida
                let isRetirada = false;
                

                // Procura na tela por todas as tags h6 com o testid específico
                const elementosH6 = document.querySelectorAll('h6[data-testid="hexa-v2-text"]');
                for (let h6 of elementosH6) {
                    // Verifica se o texto exato dentro dessa tag é "Retirada"
                    if (h6.textContent.trim() === "Retirada") {
                        isRetirada = true;
                        break; // Para de procurar assim que encontrar
                    }
                }

                // Fallback de segurança: caso a tag mude ou não carregue a tempo, ainda checa o card
                if (!isRetirada && proximoCard.innerText.toLowerCase().includes("retirada")) {
                    isRetirada = true;
                }

                // 2. Aceita
                btnAceitar.click();

                // Separa o fluxo entre pedido de retirada e pedido de entrega
                if (isRetirada && proximoCard.innerText.toLowerCase().includes("retirada")) {
                    console.log("Pedido de Retirada detectado (Tag H6 encontrada) - Apenas aceitando, sem imprimir.");

                    // Finaliza o fluxo imediatamente sem chamar o window.print()
                    marcarComoConcluido(idPedido);
                    setTimeout(() => { executandoFluxo = false; }, 800);
                } else {
                    // 3. Sequência Turbo de Impressão
                    // Espera só 0.5s após aceitar
                    setTimeout(() => {

                        // Via 1
                        console.log("Via 1");
                        window.print(true);

                        // Espera só 0.5s entre as vias (MUITO RÁPIDO)
                        setTimeout(() => {

                            // Via 2
                            console.log("Via 2");
                            window.print(true);

                            // Finaliza
                            marcarComoConcluido(idPedido);
                            setTimeout(() => { executandoFluxo = false; }, 800);

                        }, 500); // <-- Tempo entre as impressões

                    }, 500); // <-- Tempo após aceitar
                }

            } else {
                marcarComoConcluido(idPedido);
                executandoFluxo = false;
            }
        }, 1200);
    }

    setInterval(monitorarPedidos, 2500); // Verifica novos pedidos um pouco mais rápido

})();