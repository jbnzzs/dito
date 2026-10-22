/* Filtros no cabecalho do Relatorio Dito! — independentes da visibilidade.
   GET server-side: pagina, XLSX e CSV consultam o mesmo conjunto filtrado. */
(function () {
  'use strict';

  function iniciar() {
    const tabela = document.getElementById('dito-relatorio-tabela');
    if (!tabela || tabela.dataset.ditoFiltrosReady === '1') return;
    tabela.dataset.ditoFiltrosReady = '1';

    const pesquisa = new URLSearchParams(window.location.search);
    const datas = new Set([
      'prevista', 'efetiva', 'prazo', 'cadastrada', 'atualizada',
      'inicio-desc', 'entregue-desc', 'inicio-rev', 'entregue-rev',
      'data-final', 'primeira-lote-desc', 'primeira-lote-rev'
    ]);
    const escolhas = {
      projeto: 'filtro-projeto', obra: 'filtro-obra',
      componente: 'filtro-componente', lote: 'filtro-lote',
      status: 'filtro-status', responsavel: 'filtro-responsavel',
      'pag-descritor': 'filtro-pagamento-descritor',
      'pag-revisor': 'filtro-pagamento-revisor'
    };
    const antigos = {
      retranca: ['retranca'], projeto: ['projeto'], obra: ['obra'],
      componente: ['componente'], lote: ['lote'], status: ['status'],
      responsavel: ['responsavel'],
      'pag-descritor': ['pagamento_descritor'],
      'pag-revisor': ['pagamento_revisor'],
      cadastrada: ['data_inicio', 'data_fim']
    };

    // V7: filtros extras ficam no proprio formulario superior.
    // Preenche com a URL (inclusive filtros usados pelo funil do cabecalho)
    // e remove inputs hidden com o mesmo name para permitir limpar campos.
    const formularioSuperior = document.getElementById('dito-relatorio-filtros-legados');
    if (formularioSuperior) {
      formularioSuperior.querySelectorAll('[data-dito-filtro-novo]').forEach(function (campo) {
        campo.value = pesquisa.get(campo.name) || '';
        formularioSuperior.querySelectorAll('input[type="hidden"]').forEach(function (oculto) {
          if (oculto.name === campo.name) oculto.remove();
        });
      });
      formularioSuperior.addEventListener('submit', function () {
        // Campos vazios nao geram dezenas de parametros ?cf_...= na URL.
        formularioSuperior.querySelectorAll('[data-dito-filtro-novo]').forEach(function (campo) {
          campo.disabled = !campo.value.trim();
        });
      });
    }

    // V8: reabrir o bloco caso existam filtros avancados ativos.
    // Os campos permanecem no DOM e sao submetidos mesmo com <details> fechado.
    const avancados = document.getElementById('dito-relatorio-avancados');
    if (avancados) {
      const indicador = document.getElementById('dito-relatorio-avancados-contador');
      const blocos = Array.from(avancados.querySelectorAll('.dito-relatorio-avancados-campos > div'));
      function contarAvancados() {
        const quantidade = blocos.filter(function (bloco) {
          return Array.from(bloco.querySelectorAll('input[name], select[name]')).some(function (campo) {
            return !campo.disabled && String(campo.value || '').trim() !== '';
          });
        }).length;
        if (indicador) {
          indicador.hidden = quantidade === 0;
          indicador.textContent = quantidade + (quantidade === 1 ? ' ativo' : ' ativos');
        }
        return quantidade;
      }
      if (contarAvancados() > 0) avancados.open = true;
      avancados.addEventListener('input', contarAvancados);
      avancados.addEventListener('change', contarAvancados);
    }

    // O formulario antigo continua utilizavel; quando alterar a mesma coluna,
    // sua escolha substitui o filtro do cabecalho, sem conflito de valores.
    const formularioAntigo = document.getElementById('dito-relatorio-filtros-legados');
    if (formularioAntigo) {
      formularioAntigo.addEventListener('submit', function () {
        Object.entries(antigos).forEach(function ([chave, nomes]) {
          const temNovoValor = nomes.some(function (nome) {
            const campo = formularioAntigo.elements.namedItem(nome);
            return campo && typeof campo.value === 'string' && campo.value.trim();
          });
          if (!temNovoValor) return;
          const camposHeader = datas.has(chave)
            ? ['cf_' + chave + '_de', 'cf_' + chave + '_ate']
            : ['cf_' + chave];
          camposHeader.forEach(function (nome) {
            formularioAntigo.querySelectorAll('input[type="hidden"][name="' + nome + '"]')
              .forEach(function (campo) { campo.remove(); });
          });
        });
      });
    }

    let menu = null;
    let ancora = null;

    function campos(chave) {
      if (datas.has(chave)) return ['cf_' + chave + '_de', 'cf_' + chave + '_ate'];
      return ['cf_' + chave];
    }

    function filtrada(chave) {
      return [...campos(chave), ...(antigos[chave] || [])].some(
        nome => (pesquisa.get(nome) || '').trim() !== ''
      );
    }

    function fechar() {
      if (menu) menu.remove();
      if (ancora) ancora.setAttribute('aria-expanded', 'false');
      menu = null;
      ancora = null;
    }

    function posicionar() {
      if (!menu || !ancora) return;
      const r = ancora.getBoundingClientRect();
      const largura = menu.offsetWidth;
      const esquerda = Math.min(
        Math.max(8, r.left), Math.max(8, window.innerWidth - largura - 8)
      );
      menu.style.left = esquerda + 'px';
      const topo = r.bottom + 7;
      if (topo + menu.offsetHeight <= window.innerHeight - 8 || r.top < menu.offsetHeight + 8) {
        menu.style.top = Math.max(8, Math.min(topo, window.innerHeight - menu.offsetHeight - 8)) + 'px';
      } else {
        menu.style.top = Math.max(8, r.top - menu.offsetHeight - 7) + 'px';
      }
    }

    function criarInput(tipo, label, valor) {
      const rotulo = document.createElement('label');
      rotulo.className = 'dito-filtro-campo';
      const titulo = document.createElement('span');
      titulo.textContent = label;
      rotulo.appendChild(titulo);
      const campo = document.createElement('input');
      campo.className = 'form-control form-control-sm';
      campo.type = tipo;
      campo.value = valor || '';
      if (tipo === 'search') {
        campo.placeholder = 'Digite para filtrar...';
        campo.autocomplete = 'off';
      }
      rotulo.appendChild(campo);
      return { rotulo, campo };
    }

    function criarSelect(chave, valor) {
      const select = document.createElement('select');
      select.className = 'form-select form-select-sm';
      const todas = document.createElement('option');
      todas.value = '';
      todas.textContent = 'Todos os valores';
      select.appendChild(todas);

      if (chave === 'etapa') {
        const elemento = document.getElementById('dito-relatorio-etapas');
        try {
          const etapas = JSON.parse(elemento.textContent);
          etapas.forEach(function (item) {
            const opt = new Option(item[1], item[0]);
            select.appendChild(opt);
          });
        } catch (_) { /* O campo continua disponivel para todos. */ }
      } else {
        const original = document.getElementById(escolhas[chave]);
        if (original) {
          Array.from(original.options).forEach(function (option) {
            if (!option.value || option.disabled) return;
            select.appendChild(new Option(option.textContent.trim(), option.value));
          });
        }
      }
      select.value = valor || '';
      return select;
    }

    function navegar(chave, valores) {
      const qs = new URLSearchParams(window.location.search);
      qs.delete('pagina');
      campos(chave).forEach(nome => qs.delete(nome));
      (antigos[chave] || []).forEach(nome => qs.delete(nome));
      Object.entries(valores).forEach(([nome, valor]) => {
        if (valor) qs.set(nome, valor);
      });
      const url = new URL(window.location.href);
      url.search = qs.toString();
      window.location.assign(url.toString());
    }

    function abrir(chave, titulo, botao) {
      if (ancora === botao) { fechar(); return; }
      fechar();
      ancora = botao;
      botao.setAttribute('aria-expanded', 'true');
      menu = document.createElement('div');
      menu.className = 'dito-filtro-menu-coluna';
      menu.setAttribute('role', 'dialog');
      menu.setAttribute('aria-label', 'Filtrar ' + titulo);
      const cabecalho = document.createElement('strong');
      cabecalho.className = 'dito-filtro-menu-titulo';
      cabecalho.textContent = 'Filtrar: ' + titulo;
      menu.appendChild(cabecalho);
      const formulario = document.createElement('form');
      formulario.className = 'dito-filtro-menu-form';
      const valores = {};

      if (datas.has(chave)) {
        const [de, ate] = campos(chave);
        const campoDe = criarInput('date', 'De (inclusive)', pesquisa.get(de) || (chave === 'cadastrada' ? pesquisa.get('data_inicio') : ''));
        const campoAte = criarInput('date', 'Até (inclusive)', pesquisa.get(ate) || (chave === 'cadastrada' ? pesquisa.get('data_fim') : ''));
        formulario.append(campoDe.rotulo, campoAte.rotulo);
        valores[de] = campoDe.campo;
        valores[ate] = campoAte.campo;
      } else {
        const nome = campos(chave)[0];
        const legado = (antigos[chave] || []).find(x => (pesquisa.get(x) || '').trim());
        const valor = pesquisa.get(nome) || (legado ? pesquisa.get(legado) : '') || '';
        const rotulo = document.createElement('label');
        rotulo.className = 'dito-filtro-campo';
        const span = document.createElement('span');
        span.textContent = escolhas[chave] || chave === 'etapa' ? 'Selecione um valor' : 'Contém o texto';
        rotulo.appendChild(span);
        const campo = (escolhas[chave] || chave === 'etapa') ? criarSelect(chave, valor) : criarInput('search', '', valor).campo;
        rotulo.appendChild(campo);
        formulario.appendChild(rotulo);
        valores[nome] = campo;
      }

      const acoes = document.createElement('div');
      acoes.className = 'dito-filtro-menu-acoes';
      const limpar = document.createElement('button');
      limpar.type = 'button';
      limpar.className = 'btn btn-sm btn-outline-secondary';
      limpar.textContent = 'Limpar coluna';
      limpar.addEventListener('click', () => navegar(chave, {}));
      const aplicar = document.createElement('button');
      aplicar.type = 'submit';
      aplicar.className = 'btn btn-sm btn-primary';
      aplicar.textContent = 'Aplicar';
      acoes.append(limpar, aplicar);
      formulario.appendChild(acoes);
      formulario.addEventListener('submit', function (evento) {
        evento.preventDefault();
        const entrada = {};
        Object.entries(valores).forEach(([nome, campo]) => {
          entrada[nome] = campo.value.trim();
        });
        navegar(chave, entrada);
      });
      menu.appendChild(formulario);
      document.body.appendChild(menu);
      posicionar();
      const primeiro = menu.querySelector('input, select');
      if (primeiro) primeiro.focus();
    }

    tabela.querySelectorAll('thead th[data-coluna]').forEach(function (th) {
      const chave = th.dataset.coluna;
      const titulo = th.textContent.trim();
      const ativo = filtrada(chave);
      const botao = document.createElement('button');
      botao.type = 'button';
      botao.className = 'dito-filtro-botao' + (ativo ? ' dito-filtro-ativo' : '');
      botao.setAttribute('aria-label', 'Filtrar coluna ' + titulo);
      botao.setAttribute('title', 'Filtrar ' + titulo);
      botao.setAttribute('aria-expanded', 'false');
      const icone = document.createElement('i');
      icone.className = 'bi bi-funnel' + (ativo ? '-fill' : '');
      icone.setAttribute('aria-hidden', 'true');
      botao.appendChild(icone);
      botao.addEventListener('click', () => abrir(chave, titulo, botao));
      th.appendChild(botao);

    });

    document.addEventListener('pointerdown', function (evento) {
      if (menu && !menu.contains(evento.target) && (!ancora || !ancora.contains(evento.target))) fechar();
    });
    document.addEventListener('keydown', function (evento) {
      if (evento.key === 'Escape') fechar();
    });
    window.addEventListener('resize', posicionar);
    document.addEventListener('scroll', function () {
      if (menu) posicionar();
    }, true);
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', iniciar);
  } else {
    iniciar();
  }
})();
