/* DITO_LORDICON_MENU_V1
 * Animações sob hover/foco; respeita reduced-motion; Bootstrap Icons como fallback.
 */
(function () {
    'use strict';

    // DITO_LORDICON_TRACOS_REAIS_V3: INICIO
    function reforcarTracosNoMenu(dados) {
        var pilha = [dados];
        var fator = 1.35; // DITO_LORDICON_AJUSTE_V4: traco 15 -> 20,25; icone 21px
        var contagem = 0;
        function ajustarNumero(valor) {
            // Largura zero pode ser parte de uma animacao: nunca mudar.
            // Evita inflar efeitos especiais com larguras muito grandes.
            return typeof valor === 'number' && valor > 0 && valor <= 25
                ? Math.round(valor * fator * 100) / 100
                : valor;
        }
        while (pilha.length) {
            var atual = pilha.pop();
            if (!atual || typeof atual !== 'object') continue;
            if (Array.isArray(atual)) {
                for (var i = 0; i < atual.length; i++) pilha.push(atual[i]);
                continue;
            }
            if (atual.ty === 'st' && atual.w && typeof atual.w === 'object') {
                var chave = atual.w.k;
                if (typeof chave === 'number') {
                    atual.w.k = ajustarNumero(chave);
                    if (atual.w.k !== chave) contagem++;
                } else if (Array.isArray(chave)) {
                    for (var n = 0; n < chave.length; n++) {
                        var frame = chave[n];
                        if (!frame || typeof frame !== 'object') continue;
                        ['s', 'e'].forEach(function (campo) {
                            if (!Array.isArray(frame[campo])) return;
                            for (var j = 0; j < frame[campo].length; j++) {
                                frame[campo][j] = ajustarNumero(frame[campo][j]);
                            }
                        });
                    }
                    contagem++;
                }
            }
            Object.keys(atual).forEach(function (chave) { pilha.push(atual[chave]); });
        }
        return contagem;
    }
    // DITO_LORDICON_TRACOS_REAIS_V3: FIM


    function iniciarLordiconDito() {
        if (!window.lottie || typeof window.lottie.loadAnimation !== 'function') {
            // Caso a biblioteca não carregue (ex.: CDN indisponível), mantém Bootstrap Icons.
            return;
        }
        if (window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches) {
            return;
        }

        document.querySelectorAll('.dito-lordicon-slot[data-lordicon-src]').forEach(function (slot) {
            var botao = slot.closest('a, button');
            var palco = slot.querySelector('.dito-lordicon-stage');
            if (!botao || !palco || slot.dataset.lordiconReady === '1') return;
            slot.dataset.lordiconReady = '1';

            fetch(slot.getAttribute('data-lordicon-src'), { credentials: 'same-origin' })
                .then(function (res) {
                    if (!res.ok) throw new Error('Lottie HTTP ' + res.status);
                    return res.json();
                })
                .then(function (dados) {
                    var marcadores = Array.isArray(dados.markers) ? dados.markers : [];
                    var marcadorParado = marcadores.find(function (m) {
                        return m.cm === 'default:hover-pinch';
                    }) || marcadores[0];
                    var marcoPedido = slot.getAttribute('data-lordicon-marker') || 'default:hover-pinch';
                    var marcador = marcadores.find(function (m) { return m.cm === marcoPedido; }) || marcadorParado;
                    // DITO_LORDICON_AJUSTE_V4: quadro visível para Usuários e Painel Adm.
                    var arquivoIcone = slot.getAttribute('data-lordicon-src') || '';
                    var precisaRevelacao = /(?:314-avatars-group|457-shield-check|39-cog|381-arrow-right-sign-out)\.json(?:\?|$)/.test(arquivoIcone);
                    var marcadorRevelacao = marcadores.find(function (m) { return m.cm === 'in-reveal'; });
                    var quadroParado = (precisaRevelacao && marcadorRevelacao)
                        ? Number(marcadorRevelacao.tm) + Number(marcadorRevelacao.dr) - 1
                        : (marcadorParado ? Number(marcadorParado.tm) : 0);
                    var primeiro = marcador ? Number(marcador.tm) : 0;
                    var duracao = marcador ? Number(marcador.dr) : Math.min(60, dados.op || 60);
                    var ultimo = Math.min(Math.max(primeiro + 1, primeiro + duracao - 1), (dados.op || primeiro + duracao) - 1);
                    if (![quadroParado, primeiro, ultimo].every(Number.isFinite)) {
                        throw new Error('Marcadores Lottie inválidos');
                    }

                    // Aplica somente no menu; nao altera Configuracoes/Sair.
                    if (slot.closest('.dito-sidebar-nav') || slot.closest('.dito-header-right .dito-conta-acao')) {
                        // DITO_LORDICON_CABECALHO_V5: mesmo fator 1.35 do menu V4.
                        reforcarTracosNoMenu(dados);
                    }

                    var animacao = window.lottie.loadAnimation({
                        container: palco,
                        renderer: 'svg',
                        loop: false,
                        autoplay: false,
                        animationData: dados,
                        rendererSettings: { preserveAspectRatio: 'xMidYMid meet' }
                    });

                    var entrou = false;
                    function tocar() {
                        if (!entrou) return;
                        animacao.stop();
                        animacao.playSegments([primeiro, ultimo], true);
                    }
                    function parar() {
                        if (entrou) animacao.goToAndStop(quadroParado, true);
                    }

                    function finalizarCarregamento() {
                        if (entrou) return;
                        entrou = true;
                        animacao.goToAndStop(quadroParado, true);
                        slot.classList.add('dito-lordicon-carregado');
                        if (botao.matches(':hover, :focus-visible')) tocar();
                    }
                    // DITO_LORDICON_AJUSTE_V4: nunca deixar a animação no quadro final vazio.
                    animacao.addEventListener('complete', function () {
                        if (entrou) animacao.goToAndStop(quadroParado, true);
                    });
                    animacao.addEventListener('DOMLoaded', finalizarCarregamento);
                    // Protege também a situação em que o SVG já foi montado
                    // antes de o listener ser cadastrado.
                    if (animacao.isLoaded && palco.querySelector('svg')) {
                        finalizarCarregamento();
                    }
                    animacao.addEventListener('data_failed', function () {
                        slot.classList.remove('dito-lordicon-carregado');
                    });
                    botao.addEventListener('mouseenter', tocar);
                    botao.addEventListener('mouseleave', parar);
                    botao.addEventListener('focus', tocar);
                    botao.addEventListener('blur', parar);
                })
                .catch(function (erro) {
                    // A interface segue normal mesmo sem animação.
                    if (window.console && console.debug) console.debug('Dito Lordicon fallback:', erro);
                });
        });
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', iniciarLordiconDito);
    } else {
        iniciarLordiconDito();
    }
}());
