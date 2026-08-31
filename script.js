// ============================================================================
// Estado em memória (somente do que já veio da API — a lógica de busca no
// backend não muda; aqui só filtramos/ordenamos o que já foi recebido).
// ============================================================================
let ultimosResultados = [];
let ultimosAvisos = [];

// Caminhos marcados manualmente como "já verificados" pelo usuário.
// Fica em memória (não é salvo em disco) e dura enquanto a página estiver
// aberta — não interfere em nada da busca/lógica do backend.
const caminhosVerificados = new Set();

// ============================================================================
// Elementos
// ============================================================================
const form = document.getElementById('form-busca');
const inputCaminho = document.getElementById('caminho');
const btnBuscar = document.getElementById('btn-buscar');
const btnBuscarSpinner = btnBuscar.querySelector('.btn-spinner');
const btnBuscarLabel = btnBuscar.querySelector('.btn-label');

const statusArea = document.getElementById('status-area');
const emptyState = document.getElementById('empty-state');
const errorState = document.getElementById('error-state');

const resultadosSection = document.getElementById('resultados-section');
const resultsCount = document.getElementById('results-count');
const listaResultados = document.getElementById('lista-resultados');
const avisosArea = document.getElementById('avisos-area');
const emptyFiltered = document.getElementById('empty-filtered');

const filtroTexto = document.getElementById('filtro-texto');
const filtroSeveridade = document.getElementById('filtro-severidade');
const filtroOrdenacao = document.getElementById('filtro-ordenacao');

const btnCopiarTodos = document.getElementById('btn-copiar-todos');
const btnCopiarTodosLabel = document.getElementById('btn-copiar-todos-label');
const btnExportarExcel = document.getElementById('btn-exportar-excel');
const btnSalvarLog = document.getElementById('btn-salvar-log');

const toast = document.getElementById('toast');

// ============================================================================
// Busca (chamada à API — mesma lógica de antes)
// ============================================================================
form.addEventListener('submit', async function (evento) {
  evento.preventDefault();

  const caminho = inputCaminho.value.trim();
  if (!caminho) return;

  limparEstados();
  definirCarregando(true);
  statusArea.innerHTML = `
    <div class="status-loading">
      <span class="dot-spinner" aria-hidden="true"></span>
      Buscando pastas vencidas em "${escapeHtml(caminho)}"...
    </div>`;

  try {
    const resposta = await fetch('/api/scan?path=' + encodeURIComponent(caminho));
    const dados = await resposta.json();

    statusArea.innerHTML = '';

    if (!dados.sucesso) {
      mostrarErro(dados.erro);
      return;
    }

    ultimosResultados = dados.resultados || [];
    ultimosAvisos = dados.avisos || [];

    if (ultimosResultados.length === 0) {
      mostrarVazio();
    } else {
      resultadosSection.hidden = false;
      renderizarAvisos();
      aplicarFiltros();
    }
  } catch (erro) {
    statusArea.innerHTML = '';
    mostrarErro('Não foi possível se comunicar com o servidor. Verifique se o sistema está em execução.');
  } finally {
    definirCarregando(false);
  }
});

function definirCarregando(carregando) {
  btnBuscar.disabled = carregando;
  btnBuscarSpinner.hidden = !carregando;
  btnBuscarLabel.textContent = carregando ? 'Buscando...' : 'Buscar pastas';
}

function limparEstados() {
  emptyState.hidden = true;
  errorState.hidden = true;
  resultadosSection.hidden = true;
  listaResultados.innerHTML = '';
  avisosArea.innerHTML = '';
  emptyFiltered.hidden = true;
}

function mostrarErro(mensagem) {
  errorState.hidden = false;
  errorState.innerHTML = `
    <svg viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg">
      <circle cx="10" cy="10" r="7.2" stroke="currentColor" stroke-width="1.5"/>
      <path d="M10 6.5V10.5" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"/>
      <circle cx="10" cy="13.3" r="0.9" fill="currentColor"/>
    </svg>
    <div>
      <strong>Não foi possível concluir a busca</strong>
      ${escapeHtml(mensagem)}
    </div>`;
}

function mostrarVazio() {
  emptyState.hidden = false;
  emptyState.innerHTML = `
    <strong>Nenhuma pasta vencida encontrada</strong>
    Não há subpastas "Val..." vencidas há 3 meses ou mais nesse caminho.`;
}

// ============================================================================
// Avisos de acesso (permissão negada em alguma subpasta, etc.)
// ============================================================================
// Quantos avisos individuais mostrar antes de resumir o restante. Sem esse
// limite, um caminho de rede com muitas subpastas sem permissão gera uma
// lista enorme que polui a tela.
const LIMITE_AVISOS_EXIBIDOS = 12;

function renderizarAvisos() {
  if (!ultimosAvisos.length) {
    avisosArea.innerHTML = '';
    return;
  }

  const visiveis = ultimosAvisos.slice(0, LIMITE_AVISOS_EXIBIDOS);
  const restante = ultimosAvisos.length - visiveis.length;

  avisosArea.innerHTML = `
    <div class="avisos">
      <strong>Avisos durante a leitura (${ultimosAvisos.length})</strong>
      <ul>${visiveis.map(a => `<li>${escapeHtml(a)}</li>`).join('')}</ul>
      ${restante > 0 ? `<p class="avisos-mais">+ ${restante} outro${restante === 1 ? '' : 's'} caminho${restante === 1 ? '' : 's'} sem permissão de leitura.</p>` : ''}
    </div>`;
}

// ============================================================================
// Filtros, ordenação e renderização da lista
// ============================================================================
[filtroTexto, filtroSeveridade, filtroOrdenacao].forEach(el => {
  el.addEventListener('input', aplicarFiltros);
  el.addEventListener('change', aplicarFiltros);
});

function classificarSeveridade(mesesVencida) {
  if (mesesVencida >= 12) return 'critica';
  if (mesesVencida >= 6) return 'vencida';
  return 'atencao';
}

const ROTULO_SEVERIDADE = {
  atencao: 'Atenção',
  vencida: 'Vencida',
  critica: 'Crítica',
};

function aplicarFiltros() {
  const texto = filtroTexto.value.trim().toLowerCase();
  const severidade = filtroSeveridade.value;
  const ordenacao = filtroOrdenacao.value;

  let filtrados = ultimosResultados.filter(item => {
    const bateTexto = !texto ||
      item.nome.toLowerCase().includes(texto) ||
      item.caminho.toLowerCase().includes(texto);
    const bateSeveridade = severidade === 'todas' ||
      classificarSeveridade(item.meses_vencida) === severidade;
    return bateTexto && bateSeveridade;
  });

  filtrados = filtrados.slice().sort((a, b) => {
    if (ordenacao === 'antigas') return b.meses_vencida - a.meses_vencida;
    if (ordenacao === 'recentes') return a.meses_vencida - b.meses_vencida;
    return a.nome.localeCompare(b.nome, 'pt-BR');
  });

  renderizarLista(filtrados);
  atualizarContagem(filtrados.length, ultimosResultados.length);
}

function atualizarContagem(visiveis, total) {
  const palavraPasta = total === 1 ? 'pasta vencida' : 'pastas vencidas';
  let texto = `<span class="num">${total}</span> ${palavraPasta} encontrada${total === 1 ? '' : 's'}`;
  if (visiveis !== total) {
    texto += ` &middot; ${visiveis} exibida${visiveis === 1 ? '' : 's'} com os filtros atuais`;
  }
  resultsCount.innerHTML = texto;

  btnCopiarTodosLabel.textContent = visiveis === total
    ? `Copiar todos os caminhos (${total})`
    : `Copiar caminhos filtrados (${visiveis})`;
}

function renderizarLista(itens) {
  if (itens.length === 0) {
    listaResultados.innerHTML = '';
    emptyFiltered.hidden = false;
    return;
  }
  emptyFiltered.hidden = true;

  listaResultados.innerHTML = itens.map((item, indice) => {
    const sev = classificarSeveridade(item.meses_vencida);
    const verificada = caminhosVerificados.has(item.caminho);
    return `
      <li class="result-item${verificada ? ' verificada' : ''}" data-sev="${sev}">
        <span class="result-accent" aria-hidden="true"></span>
        <label class="result-check" title="Marcar como verificada manualmente">
          <input
            type="checkbox"
            class="checkbox-verificada"
            data-indice="${indice}"
            ${verificada ? 'checked' : ''}
            aria-label="Marcar ${escapeHtml(item.nome)} como verificada"
          >
        </label>
        <div class="result-body">
          <div class="result-top">
            <span class="result-nome">${escapeHtml(item.nome)}</span>
            <span class="badge-sev">${ROTULO_SEVERIDADE[sev]} &middot; ${item.meses_vencida} meses</span>
          </div>
          <div class="result-meta">Validade: ${escapeHtml(item.validade)}</div>
          <div class="result-path-row">
            <code class="result-path">${escapeHtml(item.caminho)}</code>
            <button type="button" class="btn-copiar-caminho" data-indice="${indice}">
              <svg viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">
                <rect x="7" y="7" width="10" height="10" rx="1.5" stroke="currentColor" stroke-width="1.4"/>
                <path d="M13 7V5.5C13 4.67 12.33 4 11.5 4H4.5C3.67 4 3 4.67 3 5.5V12.5C3 13.33 3.67 14 4.5 14H7" stroke="currentColor" stroke-width="1.4"/>
              </svg>
              <span>Copiar</span>
            </button>
          </div>
        </div>
      </li>`;
  }).join('');

  // Associa o clique de cada botão "Copiar" ao caminho correspondente
  listaResultados.querySelectorAll('.btn-copiar-caminho').forEach(botao => {
    botao.addEventListener('click', () => {
      const item = itens[Number(botao.dataset.indice)];
      copiarTexto(item.caminho, botao, 'Copiado!', 'Copiar');
    });
  });

  // Associa cada checkbox de verificação manual ao caminho correspondente.
  // O estado (verificado/não verificado) fica salvo em `caminhosVerificados`
  // e é reaplicado sempre que a lista for re-renderizada (filtro, ordenação
  // ou nova busca com os mesmos resultados), então o destaque não some.
  listaResultados.querySelectorAll('.checkbox-verificada').forEach(checkbox => {
    checkbox.addEventListener('change', () => {
      const item = itens[Number(checkbox.dataset.indice)];
      const linha = checkbox.closest('.result-item');
      if (checkbox.checked) {
        caminhosVerificados.add(item.caminho);
        linha.classList.add('verificada');
      } else {
        caminhosVerificados.delete(item.caminho);
        linha.classList.remove('verificada');
      }
    });
  });
}

// ============================================================================
// Copiar todos os caminhos (respeita os filtros atualmente aplicados)
// ============================================================================
btnCopiarTodos.addEventListener('click', () => {
  const linhas = Array.from(listaResultados.querySelectorAll('.result-path'))
    .map(el => el.textContent);

  if (linhas.length === 0) return;

  copiarTexto(linhas.join('\n'), btnCopiarTodos, null, null, () => {
    const original = btnCopiarTodosLabel.textContent;
    btnCopiarTodosLabel.textContent = 'Copiado!';
    setTimeout(() => { btnCopiarTodosLabel.textContent = original; }, 1600);
  });
  mostrarToast(`${linhas.length} caminho${linhas.length === 1 ? '' : 's'} copiado${linhas.length === 1 ? '' : 's'}`);
});

// ============================================================================
// Exportar para Excel (.xlsx)
// ============================================================================
// Gera a planilha inteiramente no navegador, a partir dos dados já recebidos
// da API na última busca — nenhuma chamada nova ao backend é feita.
btnExportarExcel.addEventListener('click', () => {
  if (!ultimosResultados.length) {
    mostrarToast('Nenhum resultado para exportar');
    return;
  }

  const linhas = ultimosResultados.map(item => ({
    'Nome da pasta': item.nome,
    'Validade': item.validade,
    'Meses vencida': item.meses_vencida,
    'Caminho completo': item.caminho,
    'Status': caminhosVerificados.has(item.caminho) ? 'Verificada' : 'Não verificada',
  }));

  const planilha = XLSX.utils.json_to_sheet(linhas);
  planilha['!cols'] = [
    { wch: 26 }, // Nome da pasta
    { wch: 12 }, // Validade
    { wch: 14 }, // Meses vencida
    { wch: 60 }, // Caminho completo
    { wch: 16 }, // Status
  ];

  const livro = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(livro, planilha, 'Pastas vencidas');

  const dataArquivo = new Date().toISOString().slice(0, 10);
  XLSX.writeFile(livro, `pastas-vencidas_${dataArquivo}.xlsx`);

  mostrarToast('Excel exportado com sucesso');
});

// ============================================================================
// Salvar log (.txt) — gerado no navegador a partir dos dados já recebidos
// da última busca, sem nenhuma chamada nova ao backend. O log é sempre um
// arquivo além do que já é mostrado na tela, não substitui a exibição.
// ============================================================================
btnSalvarLog.addEventListener('click', () => {
  if (!ultimosResultados.length) {
    mostrarToast('Nenhum resultado para salvar em log');
    return;
  }

  const agora = new Date();
  const carimboArquivo = agora.toISOString().slice(0, 10);
  const carimboCabecalho = agora.toLocaleString('pt-BR');

  const linhas = [
    `Log de pastas vencidas — gerado em ${carimboCabecalho}`,
    `Total de pastas vencidas encontradas: ${ultimosResultados.length}`,
    '='.repeat(60),
    '',
    ...ultimosResultados.map(item => (
      `${item.nome}\n` +
      `  Validade: ${item.validade}  |  Vencida há: ${item.meses_vencida} mes${item.meses_vencida === 1 ? '' : 'es'}\n` +
      `  Status: ${caminhosVerificados.has(item.caminho) ? 'Verificada' : 'Não verificada'}\n` +
      `  Caminho: ${item.caminho}\n`
    )),
  ];

  const blob = new Blob([linhas.join('\n')], { type: 'text/plain;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = `pastas-vencidas_${carimboArquivo}.log.txt`;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);

  mostrarToast('Log salvo com sucesso');
});

// ============================================================================
// Utilitário de cópia com fallback e feedback visual
// ============================================================================
function copiarTexto(texto, botao, textoTemporario, textoOriginal, aoConcluir) {
  const aplicarFeedback = () => {
    if (botao && textoTemporario) {
      const span = botao.querySelector('span:last-child') || botao;
      const original = span.textContent;
      span.textContent = textoTemporario;
      botao.classList.add('copiado');
      setTimeout(() => {
        span.textContent = textoOriginal || original;
        botao.classList.remove('copiado');
      }, 1600);
    }
    if (botao && !textoTemporario) {
      botao.classList.add('copiado');
      setTimeout(() => botao.classList.remove('copiado'), 1600);
    }
    if (aoConcluir) aoConcluir();
  };

  if (navigator.clipboard && window.isSecureContext) {
    navigator.clipboard.writeText(texto).then(aplicarFeedback).catch(() => {
      copiarComFallback(texto);
      aplicarFeedback();
    });
  } else {
    copiarComFallback(texto);
    aplicarFeedback();
  }
}

function copiarComFallback(texto) {
  const area = document.createElement('textarea');
  area.value = texto;
  area.style.position = 'fixed';
  area.style.opacity = '0';
  document.body.appendChild(area);
  area.focus();
  area.select();
  try {
    document.execCommand('copy');
  } catch (erro) {
    // Se nem o fallback funcionar, não há mais nada a fazer silenciosamente.
  }
  document.body.removeChild(area);
}

let toastTimeout = null;
function mostrarToast(mensagem) {
  toast.textContent = mensagem;
  toast.classList.add('visivel');
  clearTimeout(toastTimeout);
  toastTimeout = setTimeout(() => toast.classList.remove('visivel'), 1800);
}

// ============================================================================
// Auxiliar
// ============================================================================
function escapeHtml(texto) {
  const div = document.createElement('div');
  div.textContent = texto;
  return div.innerHTML;
}
