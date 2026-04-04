// ═══════════════════════════════════════════════════════════════════════════
// ASTIR – Sistema de Tombamento (localStorage)
// Permissões:
//   TI         → tombar, dar baixa, excluir, transferir tudo
//   Setores    → apenas transferir itens que estão NO SEU setor
// ═══════════════════════════════════════════════════════════════════════════

const STORAGE_KEY = 'tombamentos_db';
const USERS_KEY = 'tombamento_users';
const SESSION_KEY = 'tombamento_session';
const HISTORICO_KEY = 'tombamento_historico';
const MARCAS_KEY = 'tombamento_marcas';
const PRODUTOS_KEY = 'tombamento_produtos';
const LOGS_KEY = 'tombamento_logs';
const TOMBAMENTO_INICIO = 500;

// Catálogos padrão
const MARCAS_DEFAULT = ['Samsung', 'LG', 'Dell', 'HP', 'Lenovo', 'TP-Link', 'Intelbras', 'AOC', 'Multilaser', 'Positivo'];
const PRODUTOS_DEFAULT = ['Monitor', 'Nobreak', 'CPU', 'DVR', 'Switch', 'Microfone', 'Câmera', 'Rack'];

if (!localStorage.getItem(MARCAS_KEY)) {
  localStorage.setItem(MARCAS_KEY, JSON.stringify(MARCAS_DEFAULT));
}
if (!localStorage.getItem(PRODUTOS_KEY)) {
  localStorage.setItem(PRODUTOS_KEY, JSON.stringify(PRODUTOS_DEFAULT));
}

function getMarcas() {
  return JSON.parse(localStorage.getItem(MARCAS_KEY) || '[]').sort((a, b) => a.localeCompare(b));
}

function saveMarcas(lista) {
  localStorage.setItem(MARCAS_KEY, JSON.stringify(lista));
}

function getProdutos() {
  return JSON.parse(localStorage.getItem(PRODUTOS_KEY) || '[]').sort((a, b) => a.localeCompare(b));
}

function saveProdutos(lista) {
  localStorage.setItem(PRODUTOS_KEY, JSON.stringify(lista));
}

// Setores e senhas padrão
const SETORES_DEFAULT = {
  'TI':           'ti123',
  'Guias':        'guias123',
  'DIREX':        'direx123',
  'Faturamento':  'faturamento123',
  'Cadastro':     'cadastro123',
  'Ambulatório':  'ambulatorio123',
  'Financeiro':   'financeiro123',
  'Compras':      'compras123',
  'Jurídico':     'juridico123',
  'RH':           'rh123',
  'Almoxarifado': 'almoxarifado123'
};

if (!localStorage.getItem(USERS_KEY)) {
  localStorage.setItem(USERS_KEY, JSON.stringify(SETORES_DEFAULT));
} else {
  // Garante que TI sempre exista nas senhas salvas
  const saved = JSON.parse(localStorage.getItem(USERS_KEY));
  let updated = false;
  for (const setor in SETORES_DEFAULT) {
    if (!(setor in saved)) {
      saved[setor] = SETORES_DEFAULT[setor];
      updated = true;
    }
  }
  if (updated) localStorage.setItem(USERS_KEY, JSON.stringify(saved));
}

// ── Banco ────────────────────────────────────────────────────────────────
function getTombamentos() {
  const data = localStorage.getItem(STORAGE_KEY);
  return data ? JSON.parse(data) : [];
}

function saveTombamentos(lista) {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(lista));
}

function getHistorico() {
  const data = localStorage.getItem(HISTORICO_KEY);
  return data ? JSON.parse(data) : [];
}

function addHistorico(entry) {
  const hist = getHistorico();
  hist.unshift(entry);
  localStorage.setItem(HISTORICO_KEY, JSON.stringify(hist));
}

// ── Logs de Atividade ────────────────────────────────────────────────────
function getLogs() {
  const data = localStorage.getItem(LOGS_KEY);
  return data ? JSON.parse(data) : [];
}

function addLog(acao, detalhes, nivel) {
  const logs = getLogs();
  logs.unshift({
    id: Date.now(),
    data: new Date().toISOString(),
    setor: meuSetor() || 'Sistema',
    acao: acao,
    detalhes: detalhes || '',
    nivel: nivel || 'info'
  });
  // Manter no máximo 500 registros
  if (logs.length > 500) logs.length = 500;
  localStorage.setItem(LOGS_KEY, JSON.stringify(logs));
}

function proximoTombamento() {
  const lista = getTombamentos();
  if (lista.length === 0) return TOMBAMENTO_INICIO + 1;
  return Math.max(...lista.map(t => t.numero_tombamento)) + 1;
}

function gerarId() {
  const lista = getTombamentos();
  if (lista.length === 0) return 1;
  return Math.max(...lista.map(t => t.id)) + 1;
}

function escapeHtml(text) {
  const div = document.createElement('div');
  div.appendChild(document.createTextNode(text));
  return div.innerHTML;
}

// ── Sessão ───────────────────────────────────────────────────────────────
function getSession() {
  const s = sessionStorage.getItem(SESSION_KEY);
  return s ? JSON.parse(s) : null;
}

function setSession(setor) {
  sessionStorage.setItem(SESSION_KEY, JSON.stringify({ setor }));
}

function clearSession() {
  sessionStorage.removeItem(SESSION_KEY);
}

function isTI() {
  const s = getSession();
  return s && s.setor === 'TI';
}

function meuSetor() {
  const s = getSession();
  return s ? s.setor : null;
}

// ═════════════════════════════════════════════════════════════════════════
// INIT
// ═════════════════════════════════════════════════════════════════════════
document.addEventListener('DOMContentLoaded', () => {

  // ── MODO ESCURO ────────────────────────────────────────────────────────
  const THEME_KEY = 'tombamento_theme';
  const btnDark = document.getElementById('btn-dark-mode');

  function aplicarTema(tema) {
    document.documentElement.setAttribute('data-theme', tema);
    localStorage.setItem(THEME_KEY, tema);
    const icon = btnDark.querySelector('i');
    const label = btnDark.querySelector('span');
    if (tema === 'dark') {
      icon.className = 'fas fa-sun';
      label.textContent = 'Modo Claro';
    } else {
      icon.className = 'fas fa-moon';
      label.textContent = 'Modo Escuro';
    }
  }

  // Carregar tema salvo
  const temaSalvo = localStorage.getItem(THEME_KEY) || 'light';
  aplicarTema(temaSalvo);

  btnDark.addEventListener('click', () => {
    const temaAtual = document.documentElement.getAttribute('data-theme');
    aplicarTema(temaAtual === 'dark' ? 'light' : 'dark');
  });

  const telaLogin = document.getElementById('tela-login');
  const appPrincipal = document.getElementById('app-principal');

  const sessao = getSession();
  if (sessao) mostrarApp(sessao.setor);

  // ── LOGIN ──────────────────────────────────────────────────────────────
  const formLogin = document.getElementById('form-login');
  const loginError = document.getElementById('login-error');

  formLogin.addEventListener('submit', (e) => {
    e.preventDefault();
    const setor = document.getElementById('login-setor').value;
    const senha = document.getElementById('login-senha').value;
    const users = JSON.parse(localStorage.getItem(USERS_KEY));

    if (!setor || !users[setor]) {
      loginError.textContent = 'Selecione um setor válido.';
      return;
    }
    if (senha !== users[setor]) {
      loginError.textContent = 'Senha incorreta.';
      addLog('Login falhou', `Tentativa de login no setor ${setor} — senha incorreta`, 'erro');
      return;
    }

    loginError.textContent = '';
    setSession(setor);
    addLog('Login', `Setor ${setor} entrou no sistema`, 'info');
    mostrarApp(setor);
  });

  function mostrarApp(setor) {
    telaLogin.style.display = 'none';
    appPrincipal.style.display = 'flex';
    document.getElementById('user-setor').textContent = setor;

    // Controle de visibilidade por permissão
    const btnAdd = document.getElementById('btn-adicionar');
    const navConfig = document.getElementById('nav-config');
    const navLogs = document.getElementById('nav-logs');
    if (isTI()) {
      btnAdd.style.display = '';
      navConfig.style.display = '';
      navLogs.style.display = '';
    } else {
      btnAdd.style.display = 'none';
      navConfig.style.display = 'none';
      navLogs.style.display = 'none';
    }

    popularSelectsProdutos();
    popularDatalistMarcas();
    carregarDashboard();
  }

  // Logout
  document.getElementById('btn-logout').addEventListener('click', () => {
    addLog('Logout', `Setor ${meuSetor()} saiu do sistema`, 'info');
    clearSession();
    appPrincipal.style.display = 'none';
    telaLogin.style.display = 'flex';
    formLogin.reset();
    loginError.textContent = '';
  });

  // ── NAVEGAÇÃO ──────────────────────────────────────────────────────────
  const navItems = document.querySelectorAll('.nav-item');
  const sections = document.querySelectorAll('.section');

  navItems.forEach(item => {
    item.addEventListener('click', (e) => {
      e.preventDefault();
      const target = item.dataset.section;
      navItems.forEach(n => n.classList.remove('active'));
      item.classList.add('active');
      sections.forEach(s => s.classList.remove('active'));
      document.getElementById(`section-${target}`).classList.add('active');
      if (target === 'dashboard') carregarDashboard();
      if (target === 'listagem') carregarListagem();
      if (target === 'configuracoes') carregarConfiguracoes();
      if (target === 'logs') carregarLogs();
    });
  });

  // ── MODAL CADASTRO (só TI) ────────────────────────────────────────────
  const modalCadastro = document.getElementById('modal-cadastro');
  const btnAdicionar = document.getElementById('btn-adicionar');
  const btnFecharCad = document.getElementById('btn-fechar-cadastro');
  const btnCancelarCad = document.getElementById('btn-cancelar-cadastro');

  function abrirModalCadastro() {
    if (!isTI()) return;
    document.getElementById('tombamento-preview').innerHTML =
      `<span class="tombamento-numero">#${proximoTombamento()}</span>`;
    modalCadastro.classList.add('show');
  }

  function fecharModalCadastro() {
    modalCadastro.classList.remove('show');
    document.getElementById('form-cadastro').reset();
  }

  btnAdicionar.addEventListener('click', abrirModalCadastro);
  btnFecharCad.addEventListener('click', fecharModalCadastro);
  btnCancelarCad.addEventListener('click', fecharModalCadastro);
  modalCadastro.addEventListener('click', (e) => {
    if (e.target === modalCadastro) fecharModalCadastro();
  });

  // ── CADASTRO ───────────────────────────────────────────────────────────
  const formCadastro = document.getElementById('form-cadastro');
  formCadastro.addEventListener('submit', (e) => {
    e.preventDefault();
    if (!isTI()) return;

    const produto = document.getElementById('produto').value;
    const marca = document.getElementById('marca').value.trim();
    const numero_serie = document.getElementById('numero_serie').value.trim();
    if (!produto || !marca || !numero_serie) return;

    const lista = getTombamentos();
    const num = proximoTombamento();
    const novo = {
      id: gerarId(),
      numero_tombamento: num,
      produto,
      marca,
      numero_serie,
      status: 'Ativo',
      setor: 'TI',
      data_cadastro: new Date().toISOString()
    };

    lista.push(novo);
    saveTombamentos(lista);

    addHistorico({
      tombamento: num,
      tipo: 'Cadastro',
      de: '—',
      para: 'TI',
      por: 'TI',
      data: new Date().toISOString()
    });

    addLog('Cadastro', `Tombamento #${num} — ${produto} ${marca} (Série: ${numero_serie})`, 'sucesso');
    showToast(`Tombamento #${num} cadastrado com sucesso!`);
    fecharModalCadastro();
    carregarListagem();
  });

  // ── TOAST ──────────────────────────────────────────────────────────────
  function showToast(msg, isError = false) {
    const toast = document.getElementById('toast-sucesso');
    const toastMsg = document.getElementById('toast-msg');
    toastMsg.textContent = msg;
    toast.style.background = isError ? '#ef4444' : '#10b981';
    toast.querySelector('i').className = isError
      ? 'fas fa-exclamation-circle' : 'fas fa-check-circle';
    toast.classList.add('show');
    setTimeout(() => toast.classList.remove('show'), 3000);
  }

  // ── LISTAGEM ───────────────────────────────────────────────────────────
  const tabelaBody = document.getElementById('tabela-body');
  const emptyState = document.getElementById('empty-state');

  function carregarListagem() {
    const busca = document.getElementById('filtro-busca').value.toLowerCase();
    const filtroProduto = document.getElementById('filtro-produto').value;
    const filtroStatus = document.getElementById('filtro-status').value;
    const filtroSetor = document.getElementById('filtro-setor').value;
    const setor = meuSetor();
    const tiUser = isTI();

    let lista = getTombamentos();

    if (filtroProduto !== 'Todos') {
      lista = lista.filter(t => t.produto === filtroProduto);
    }
    if (filtroStatus !== 'Todos') {
      lista = lista.filter(t => t.status === filtroStatus);
    }
    if (filtroSetor !== 'Todos') {
      lista = lista.filter(t => t.setor === filtroSetor);
    }
    if (busca) {
      lista = lista.filter(t =>
        t.marca.toLowerCase().includes(busca) ||
        t.numero_serie.toLowerCase().includes(busca) ||
        String(t.numero_tombamento).includes(busca) ||
        (t.setor && t.setor.toLowerCase().includes(busca))
      );
    }

    // Resumo do setor filtrado
    const resumoEl = document.getElementById('resumo-setor');
    if (filtroSetor !== 'Todos') {
      const totalSetor = lista.length;
      const ativosSetor = lista.filter(t => t.status === 'Ativo').length;
      const baixaSetor = lista.filter(t => t.status === 'Em Baixa').length;
      const porProduto = {};
      const marcas = new Set();
      const tombamentos = [];
      lista.forEach(t => {
        porProduto[t.produto] = (porProduto[t.produto] || 0) + 1;
        marcas.add(t.marca);
        tombamentos.push(`#${t.numero_tombamento}`);
      });

      const produtosList = Object.entries(porProduto)
        .sort((a, b) => b[1] - a[1])
        .map(([p, q]) => `<span class="resumo-tag">${escapeHtml(p)}: <strong>${q}</strong></span>`)
        .join('');

      const marcasList = [...marcas].map(m => `<span class="resumo-tag marca-tag">${escapeHtml(m)}</span>`).join('');

      resumoEl.innerHTML = `
        <div class="resumo-header">
          <i class="fas fa-building"></i>
          <strong>${escapeHtml(filtroSetor)}</strong>
          <span class="resumo-total">${totalSetor} equipamento${totalSetor !== 1 ? 's' : ''}</span>
          <span class="resumo-ativo"><i class="fas fa-check-circle"></i> ${ativosSetor} ativos</span>
          <span class="resumo-baixa"><i class="fas fa-arrow-down"></i> ${baixaSetor} em baixa</span>
        </div>
        <div class="resumo-detalhe">
          <div class="resumo-linha"><span class="resumo-label">Produtos:</span>${produtosList}</div>
          <div class="resumo-linha"><span class="resumo-label">Marcas:</span>${marcasList}</div>
          <div class="resumo-linha"><span class="resumo-label">Tombamentos:</span><span class="resumo-tombs">${tombamentos.join(', ')}</span></div>
        </div>`;
      resumoEl.style.display = '';
    } else {
      resumoEl.style.display = 'none';
    }

    lista.sort((a, b) => b.numero_tombamento - a.numero_tombamento);

    if (lista.length === 0) {
      tabelaBody.innerHTML = '';
      emptyState.style.display = 'block';
      return;
    }

    emptyState.style.display = 'none';

    tabelaBody.innerHTML = lista.map(item => {
      const isAtivo = item.status === 'Ativo';
      const isMeuSetor = item.setor === setor;

      // Botão toggle (junto ao status)
      let toggleHtml = '';
      if (tiUser) {
        toggleHtml = `<button class="btn-status-toggle ${isAtivo ? 'btn-status-baixa' : 'btn-status-ativar'}" title="${isAtivo ? 'Dar Baixa' : 'Ativar'}" data-toggle-id="${item.id}">
            <i class="fas fa-${isAtivo ? 'arrow-down' : 'arrow-up'}"></i>
          </button>`;
      }

      // Monta botões de ação conforme permissão
      let acoesHtml = '';

      if (tiUser) {
        // TI: pode tudo (sem o toggle que foi pro status)
        acoesHtml = `
          <button class="btn-icon btn-transfer" title="Transferir" data-transfer-id="${item.id}">
            <i class="fas fa-share"></i>
          </button>
          <button class="btn-icon btn-history" title="Histórico" data-hist-id="${item.id}" data-hist-num="${item.numero_tombamento}">
            <i class="fas fa-history"></i>
          </button>
          <button class="btn-icon btn-delete" title="Excluir" data-delete-id="${item.id}" data-delete-num="${item.numero_tombamento}">
            <i class="fas fa-trash-alt"></i>
          </button>`;
      } else if (isMeuSetor && isAtivo) {
        // Setor dono: pode transferir apenas
        acoesHtml = `
          <button class="btn-icon btn-transfer" title="Transferir para outro setor" data-transfer-id="${item.id}">
            <i class="fas fa-share"></i>
          </button>
          <button class="btn-icon btn-history" title="Histórico" data-hist-id="${item.id}" data-hist-num="${item.numero_tombamento}">
            <i class="fas fa-history"></i>
          </button>`;
      } else {
        // Não é do meu setor: sem ações
        acoesHtml = `
          <button class="btn-icon btn-history" title="Histórico" data-hist-id="${item.id}" data-hist-num="${item.numero_tombamento}">
            <i class="fas fa-history"></i>
          </button>
          <span class="sem-acao">—</span>`;
      }

      return `
        <tr class="${isMeuSetor ? 'row-meu-setor' : ''}">
          <td class="tombamento-cell">#${item.numero_tombamento}</td>
          <td>
            <div class="produto-setor-cell">
              <span class="produto-nome">${escapeHtml(item.produto)}</span>
              <span class="produto-setor-tag"><i class="fas fa-building"></i> ${escapeHtml(item.setor || '—')}</span>
            </div>
          </td>
          <td>${escapeHtml(item.marca)}</td>
          <td>${escapeHtml(item.numero_serie)}</td>
          <td>
            <div class="status-cell">
              <span class="badge ${isAtivo ? 'badge-ativo' : 'badge-baixa'}">
                <i class="fas fa-${isAtivo ? 'check-circle' : 'arrow-down'}"></i>
                ${isAtivo ? 'Ativo' : 'Em Baixa'}
              </span>
              ${toggleHtml}
            </div>
          </td>
          <td>${formatarData(item.data_cadastro)}</td>
          <td><div class="acoes-cell">${acoesHtml}</div></td>
        </tr>`;
    }).join('');

    // Event listeners
    tabelaBody.querySelectorAll('.btn-status-toggle').forEach(btn => {
      btn.addEventListener('click', () => toggleStatus(Number(btn.dataset.toggleId)));
    });
    tabelaBody.querySelectorAll('.btn-delete').forEach(btn => {
      btn.addEventListener('click', () =>
        confirmarExclusao(Number(btn.dataset.deleteId), btn.dataset.deleteNum));
    });
    tabelaBody.querySelectorAll('.btn-transfer').forEach(btn => {
      btn.addEventListener('click', () => abrirModalTransferir(Number(btn.dataset.transferId)));
    });
    tabelaBody.querySelectorAll('.btn-history').forEach(btn => {
      btn.addEventListener('click', () => abrirHistorico(Number(btn.dataset.histNum)));
    });
  }

  function formatarData(dateStr) {
    if (!dateStr) return '—';
    const d = new Date(dateStr);
    return d.toLocaleDateString('pt-BR') + ' ' +
      d.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });
  }

  // Filtros
  document.getElementById('filtro-busca').addEventListener('input', debounce(carregarListagem, 300));
  document.getElementById('filtro-produto').addEventListener('change', carregarListagem);
  document.getElementById('filtro-status').addEventListener('change', carregarListagem);
  document.getElementById('filtro-setor').addEventListener('change', carregarListagem);

  // Exportar Excel
  document.getElementById('btn-exportar').addEventListener('click', exportarExcel);

  function debounce(fn, delay) {
    let timer;
    return (...args) => { clearTimeout(timer); timer = setTimeout(() => fn(...args), delay); };
  }

  // ── TOGGLE STATUS (só TI) ─────────────────────────────────────────────
  function toggleStatus(id) {
    if (!isTI()) { showToast('Apenas a TI pode alterar o status.', true); return; }
    const lista = getTombamentos();
    const item = lista.find(t => t.id === id);
    if (!item) return;

    const statusAntigo = item.status;
    item.status = item.status === 'Ativo' ? 'Em Baixa' : 'Ativo';
    saveTombamentos(lista);

    addHistorico({
      tombamento: item.numero_tombamento,
      tipo: item.status === 'Em Baixa' ? 'Baixa' : 'Reativação',
      de: statusAntigo,
      para: item.status,
      por: 'TI',
      data: new Date().toISOString()
    });

    addLog('Status alterado', `#${item.numero_tombamento} — ${statusAntigo} → ${item.status}`, item.status === 'Em Baixa' ? 'alerta' : 'sucesso');
    carregarListagem();
    showToast(`Status alterado para ${item.status}!`);
  }

  // ── EXCLUIR (só TI) ───────────────────────────────────────────────────
  const modalOverlay = document.getElementById('modal-overlay');
  const modalMsg = document.getElementById('modal-msg');
  const modalConfirmar = document.getElementById('modal-confirmar');
  const modalCancelarBtn = document.getElementById('modal-cancelar');
  let excluirId = null;

  function confirmarExclusao(id, numero) {
    if (!isTI()) { showToast('Apenas a TI pode excluir.', true); return; }
    excluirId = id;
    modalMsg.textContent = `Deseja realmente excluir o tombamento #${numero}? Esta ação não pode ser desfeita.`;
    modalOverlay.classList.add('show');
  }

  modalCancelarBtn.addEventListener('click', () => {
    modalOverlay.classList.remove('show');
    excluirId = null;
  });

  modalOverlay.addEventListener('click', (e) => {
    if (e.target === modalOverlay) { modalOverlay.classList.remove('show'); excluirId = null; }
  });

  modalConfirmar.addEventListener('click', () => {
    if (!excluirId || !isTI()) return;
    let lista = getTombamentos();
    const itemExcluir = lista.find(t => t.id === excluirId);
    lista = lista.filter(t => t.id !== excluirId);
    saveTombamentos(lista);
    if (itemExcluir) addLog('Exclusão', `#${itemExcluir.numero_tombamento} — ${itemExcluir.produto} ${itemExcluir.marca} excluído`, 'erro');
    showToast('Tombamento excluído com sucesso!');
    carregarListagem();
    modalOverlay.classList.remove('show');
    excluirId = null;
  });

  // ── TRANSFERIR ─────────────────────────────────────────────────────────
  const modalTransferir = document.getElementById('modal-transferir');
  const transferirInfo = document.getElementById('transferir-info');
  const transferirDestino = document.getElementById('transferir-destino');
  const btnFecharTransf = document.getElementById('btn-fechar-transferir');
  const btnCancelarTransf = document.getElementById('btn-cancelar-transferir');
  const btnConfirmarTransf = document.getElementById('btn-confirmar-transferir');
  let transferirId = null;

  function abrirModalTransferir(id) {
    const lista = getTombamentos();
    const item = lista.find(t => t.id === id);
    if (!item) return;

    const setor = meuSetor();
    // Setores só podem transferir do próprio setor
    if (!isTI() && item.setor !== setor) {
      showToast('Você só pode transferir equipamentos do seu setor.', true);
      return;
    }

    if (item.status !== 'Ativo') {
      showToast('Não é possível transferir equipamento em baixa.', true);
      return;
    }

    transferirId = id;
    transferirInfo.innerHTML = `
      <div class="transfer-detail">
        <strong>#${item.numero_tombamento}</strong> — ${escapeHtml(item.produto)} ${escapeHtml(item.marca)}
        <br><small>Atualmente em: <strong>${escapeHtml(item.setor)}</strong></small>
      </div>`;
    transferirDestino.value = '';
    document.getElementById('transferir-justificativa').value = '';

    // Remove opção do setor atual do destino
    Array.from(transferirDestino.options).forEach(opt => {
      opt.hidden = (opt.value === item.setor);
    });

    modalTransferir.classList.add('show');
  }

  function fecharModalTransferir() {
    modalTransferir.classList.remove('show');
    transferirId = null;
    transferirDestino.value = '';
    document.getElementById('transferir-justificativa').value = '';
  }

  btnFecharTransf.addEventListener('click', fecharModalTransferir);
  btnCancelarTransf.addEventListener('click', fecharModalTransferir);
  modalTransferir.addEventListener('click', (e) => {
    if (e.target === modalTransferir) fecharModalTransferir();
  });

  btnConfirmarTransf.addEventListener('click', () => {
    if (!transferirId) return;
    const destino = transferirDestino.value;
    if (!destino) { showToast('Selecione o setor de destino.', true); return; }
    const justificativa = document.getElementById('transferir-justificativa').value.trim();
    if (!justificativa) { showToast('Informe a justificativa da transferência.', true); return; }

    const lista = getTombamentos();
    const item = lista.find(t => t.id === transferirId);
    if (!item) return;

    const setor = meuSetor();
    if (!isTI() && item.setor !== setor) {
      showToast('Você só pode transferir equipamentos do seu setor.', true);
      fecharModalTransferir();
      return;
    }

    const origem = item.setor;
    item.setor = destino;
    saveTombamentos(lista);

    addHistorico({
      tombamento: item.numero_tombamento,
      tipo: 'Transferência',
      de: origem,
      para: destino,
      por: setor,
      justificativa: justificativa,
      data: new Date().toISOString()
    });

    addLog('Transferência', `#${item.numero_tombamento} — ${origem} → ${destino} (${justificativa})`, 'alerta');
    showToast(`#${item.numero_tombamento} transferido de ${origem} → ${destino}`);
    fecharModalTransferir();
    carregarListagem();
  });

  // ── HISTÓRICO ──────────────────────────────────────────────────────────
  const modalHistorico = document.getElementById('modal-historico');
  const historicoBody = document.getElementById('historico-body');
  const btnFecharHist = document.getElementById('btn-fechar-historico');

  function abrirHistorico(numTombamento) {
    const hist = getHistorico().filter(h => h.tombamento === numTombamento);

    if (hist.length === 0) {
      historicoBody.innerHTML = '<p class="empty-hist">Nenhuma movimentação registrada.</p>';
    } else {
      historicoBody.innerHTML = `
        <div class="hist-titulo">Tombamento <strong>#${numTombamento}</strong></div>
        ${hist.map(h => `
          <div class="hist-item">
            <div class="hist-icon ${h.tipo === 'Transferência' ? 'hist-transfer' : h.tipo === 'Baixa' ? 'hist-baixa' : h.tipo === 'Reativação' ? 'hist-ativo' : 'hist-cadastro'}">
              <i class="fas fa-${h.tipo === 'Transferência' ? 'share' : h.tipo === 'Baixa' ? 'arrow-down' : h.tipo === 'Reativação' ? 'arrow-up' : 'plus'}"></i>
            </div>
            <div class="hist-content">
              <strong>${escapeHtml(h.tipo)}</strong>
              ${h.tipo === 'Transferência'
                ? `<span>de <strong>${escapeHtml(h.de)}</strong> para <strong>${escapeHtml(h.para)}</strong></span>
                   ${h.justificativa ? `<span class="hist-justificativa"><i class="fas fa-comment-dots"></i> ${escapeHtml(h.justificativa)}</span>` : ''}`
                : h.tipo === 'Cadastro'
                  ? `<span>Tombado no setor <strong>${escapeHtml(h.para)}</strong></span>`
                  : `<span>${escapeHtml(h.de)} → ${escapeHtml(h.para)}</span>`
              }
              <small>Por: ${escapeHtml(h.por)} • ${formatarData(h.data)}</small>
            </div>
          </div>
        `).join('')}`;
    }

    modalHistorico.classList.add('show');
  }

  btnFecharHist.addEventListener('click', () => modalHistorico.classList.remove('show'));
  modalHistorico.addEventListener('click', (e) => {
    if (e.target === modalHistorico) modalHistorico.classList.remove('show');
  });

  // ── DASHBOARD ──────────────────────────────────────────────────────────
  function carregarDashboard() {
    const lista = getTombamentos();
    const total = lista.length;
    const ativos = lista.filter(t => t.status === 'Ativo').length;
    const baixa = lista.filter(t => t.status === 'Em Baixa').length;

    document.getElementById('stat-total').textContent = total;
    document.getElementById('stat-ativos').textContent = ativos;
    document.getElementById('stat-baixa').textContent = baixa;
    document.getElementById('stat-proximo').textContent = `#${proximoTombamento()}`;

    const contagem = {};
    lista.forEach(t => { contagem[t.produto] = (contagem[t.produto] || 0) + 1; });

    const chartEl = document.getElementById('chart-produtos');
    const produtos = Object.entries(contagem).sort((a, b) => b[1] - a[1]);

    if (produtos.length === 0) {
      chartEl.innerHTML = '<p style="color:#94a3b8;text-align:center;padding:20px;">Nenhum equipamento cadastrado ainda.</p>';
      return;
    }

    const iconesProduto = {
      'Monitor': 'fa-desktop',
      'Nobreak': 'fa-car-battery',
      'CPU': 'fa-microchip',
      'DVR': 'fa-video',
      'Switch': 'fa-network-wired',
      'Microfone': 'fa-microphone',
      'Câmera': 'fa-camera',
      'Rack': 'fa-server',
      'Impressora': 'fa-print',
      'Roteador': 'fa-wifi',
      'Teclado': 'fa-keyboard',
      'Mouse': 'fa-computer-mouse'
    };

    chartEl.innerHTML = produtos.map(([nome, qtd]) => {
      const pct = total > 0 ? Math.round((qtd / total) * 100) : 0;
      const icone = iconesProduto[nome] || 'fa-cube';
      return `
        <div class="produto-card">
          <div class="produto-card-icon"><i class="fas ${icone}"></i></div>
          <div class="produto-card-info">
            <span class="produto-card-nome">${escapeHtml(nome)}</span>
            <span class="produto-card-qtd">${qtd} <small>un.</small></span>
          </div>
          <div class="produto-card-pct">${pct}%</div>
        </div>`;
    }).join('');

    // ── Relatório por setor (só TI) ──────────────────────────────────
    const cardSetores = document.getElementById('card-setores');
    const gridSetores = document.getElementById('grid-setores');

    if (isTI()) {
      cardSetores.style.display = '';

      // Agrupar por setor
      const porSetor = {};
      lista.forEach(t => {
        if (!porSetor[t.setor]) porSetor[t.setor] = [];
        porSetor[t.setor].push(t);
      });

      const setoresOrdenados = Object.entries(porSetor).sort((a, b) => b[1].length - a[1].length);

      if (setoresOrdenados.length === 0) {
        gridSetores.innerHTML = '<p style="color:#94a3b8;text-align:center;padding:20px;">Nenhum dado.</p>';
      } else {
        gridSetores.innerHTML = setoresOrdenados.map(([setor, items]) => {
          const ativos = items.filter(i => i.status === 'Ativo').length;
          const baixas = items.filter(i => i.status === 'Em Baixa').length;

          // Detalhe por produto
          const detalheProduto = {};
          items.forEach(i => {
            detalheProduto[i.produto] = (detalheProduto[i.produto] || 0) + 1;
          });
          const detalhes = Object.entries(detalheProduto)
            .sort((a, b) => b[1] - a[1])
            .map(([prod, qtd]) => `<div class="tooltip-row"><span>${escapeHtml(prod)}</span><strong>${qtd}</strong></div>`)
            .join('');

          return `
            <div class="setor-card">
              <div class="setor-card-header">
                <div class="setor-nome">
                  <i class="fas fa-building"></i>
                  <span>${escapeHtml(setor)}</span>
                </div>
                <span class="setor-total">${items.length}</span>
              </div>
              <div class="setor-card-status">
                <span class="setor-ativo"><i class="fas fa-check-circle"></i> ${ativos}</span>
                <span class="setor-baixa"><i class="fas fa-arrow-down"></i> ${baixas}</span>
              </div>
              <div class="setor-tooltip">
                <div class="tooltip-title">Detalhamento — ${escapeHtml(setor)}</div>
                ${detalhes}
              </div>
            </div>`;
        }).join('');
      }
    } else {
      cardSetores.style.display = 'none';
    }
  }

  // ── EXPORTAR EXCEL (.xlsx com ExcelJS) ─────────────────────────────────
  async function exportarExcel() {
    const filtroProduto = document.getElementById('filtro-produto').value;
    const filtroStatus = document.getElementById('filtro-status').value;
    const filtroSetor = document.getElementById('filtro-setor').value;
    const busca = document.getElementById('filtro-busca').value.toLowerCase();

    let lista = getTombamentos();

    if (filtroProduto !== 'Todos') lista = lista.filter(t => t.produto === filtroProduto);
    if (filtroStatus !== 'Todos') lista = lista.filter(t => t.status === filtroStatus);
    if (filtroSetor !== 'Todos') lista = lista.filter(t => t.setor === filtroSetor);
    if (busca) {
      lista = lista.filter(t =>
        t.marca.toLowerCase().includes(busca) ||
        t.numero_serie.toLowerCase().includes(busca) ||
        String(t.numero_tombamento).includes(busca) ||
        (t.setor && t.setor.toLowerCase().includes(busca))
      );
    }

    if (lista.length === 0) {
      showToast('Nenhum dado para exportar.', true);
      return;
    }

    lista.sort((a, b) => a.numero_tombamento - b.numero_tombamento);

    const wb = new ExcelJS.Workbook();
    wb.creator = 'ASTIR - Sistema de Tombamento';
    wb.created = new Date();
    const ws = wb.addWorksheet('Tombamentos', {
      properties: { defaultRowHeight: 22 },
      views: [{ state: 'frozen', ySplit: 4 }]
    });

    // ── Cores
    const AZUL_ESCURO = '0A1628';
    const AZUL_MEDIO  = '2563EB';
    const AZUL_CLARO  = 'DBEAFE';
    const VERDE_BG    = 'D1FAE5';
    const VERDE_TXT   = '065F46';
    const VERM_BG     = 'FEE2E2';
    const VERM_TXT    = '991B1B';
    const CINZA_BG    = 'F8FAFC';
    const CINZA_BORDA = 'E2E8F0';

    // ── Colunas (larguras espaçosas)
    ws.columns = [
      { key: 'tomb',  width: 16 },
      { key: 'prod',  width: 20 },
      { key: 'marca', width: 22 },
      { key: 'serie', width: 28 },
      { key: 'setor', width: 20 },
      { key: 'status',width: 16 }
    ];

    // ── Linha 1: Título grande
    ws.mergeCells('A1:F1');
    const cellTitulo = ws.getCell('A1');
    const subtitulo = filtroSetor !== 'Todos' ? ` — ${filtroSetor}` : '';
    cellTitulo.value = `ASTIR — Relatório de Tombamentos${subtitulo}`;
    cellTitulo.font = { name: 'Calibri', size: 18, bold: true, color: { argb: 'FF' + AZUL_ESCURO } };
    cellTitulo.alignment = { horizontal: 'center', vertical: 'middle' };
    cellTitulo.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + AZUL_CLARO } };
    ws.getRow(1).height = 42;

    // ── Linha 2: Subtítulo com data e totais
    ws.mergeCells('A2:F2');
    const cellSub = ws.getCell('A2');
    const agora = new Date();
    const dataStr = agora.toLocaleDateString('pt-BR') + ' às ' + agora.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });
    const totalAtivos = lista.filter(t => t.status === 'Ativo').length;
    const totalBaixa = lista.filter(t => t.status === 'Em Baixa').length;
    cellSub.value = `Gerado em ${dataStr}   |   Total: ${lista.length}   |   Ativos: ${totalAtivos}   |   Em Baixa: ${totalBaixa}`;
    cellSub.font = { name: 'Calibri', size: 10, italic: true, color: { argb: 'FF475569' } };
    cellSub.alignment = { horizontal: 'center', vertical: 'middle' };
    ws.getRow(2).height = 24;

    // ── Linha 3: espaçador
    ws.getRow(3).height = 8;

    // ── Linha 4: Cabeçalhos
    const headers = ['Tombamento', 'Produto', 'Marca', 'Número de Série', 'Setor', 'Status'];
    const headerRow = ws.getRow(4);
    headers.forEach((h, i) => {
      const cell = headerRow.getCell(i + 1);
      cell.value = h;
      cell.font = { name: 'Calibri', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + AZUL_MEDIO } };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      cell.border = {
        bottom: { style: 'medium', color: { argb: 'FF' + AZUL_ESCURO } }
      };
    });
    headerRow.height = 30;

    // ── Filtros automáticos
    ws.autoFilter = { from: 'A4', to: 'F4' };

    // ── Dados
    lista.forEach((t, idx) => {
      const r = ws.getRow(5 + idx);
      r.getCell(1).value = t.numero_tombamento;
      r.getCell(2).value = t.produto;
      r.getCell(3).value = t.marca;
      r.getCell(4).value = t.numero_serie;
      r.getCell(5).value = t.setor || '';
      r.getCell(6).value = t.status;

      // Alinhamento
      r.getCell(1).alignment = { horizontal: 'center' };
      r.getCell(6).alignment = { horizontal: 'center' };

      // Zebra (linhas alternadas)
      const bgColor = idx % 2 === 0 ? 'FFFFFFFF' : 'FF' + CINZA_BG;
      for (let c = 1; c <= 6; c++) {
        const cell = r.getCell(c);
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgColor } };
        cell.font = { name: 'Calibri', size: 10.5, color: { argb: 'FF334155' } };
        cell.border = {
          bottom: { style: 'thin', color: { argb: 'FF' + CINZA_BORDA } }
        };
      }

      // Status colorido
      const cellStatus = r.getCell(6);
      if (t.status === 'Ativo') {
        cellStatus.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + VERDE_BG } };
        cellStatus.font = { name: 'Calibri', size: 10.5, bold: true, color: { argb: 'FF' + VERDE_TXT } };
      } else {
        cellStatus.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + VERM_BG } };
        cellStatus.font = { name: 'Calibri', size: 10.5, bold: true, color: { argb: 'FF' + VERM_TXT } };
      }

      r.height = 24;
    });

    // ── Borda externa na área do cabeçalho
    for (let c = 1; c <= 6; c++) {
      const cell = ws.getRow(4).getCell(c);
      cell.border = {
        top: { style: 'medium', color: { argb: 'FF' + AZUL_MEDIO } },
        bottom: { style: 'medium', color: { argb: 'FF' + AZUL_ESCURO } },
        left: c === 1 ? { style: 'medium', color: { argb: 'FF' + AZUL_MEDIO } } : undefined,
        right: c === 6 ? { style: 'medium', color: { argb: 'FF' + AZUL_MEDIO } } : undefined
      };
    }

    // ── Gerar e baixar
    const buffer = await wb.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    const nomeArq = filtroSetor !== 'Todos'
      ? `tombamentos_${filtroSetor.replace(/\s+/g, '_')}.xlsx`
      : 'tombamentos_ASTIR.xlsx';
    a.download = nomeArq;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    showToast('Exportação realizada com sucesso!');
  }

  // ── CATÁLOGOS (popular selects e datalists dinâmicos) ──────────────────
  function popularSelectsProdutos() {
    const produtos = getProdutos();

    // Select do modal de cadastro
    const selectProduto = document.getElementById('produto');
    const valAtual = selectProduto.value;
    selectProduto.innerHTML = '<option value="">Selecione o produto...</option>' +
      produtos.map(p => `<option value="${escapeHtml(p)}">${escapeHtml(p)}</option>`).join('');
    if (valAtual) selectProduto.value = valAtual;

    // Select do filtro na listagem
    const filtroProduto = document.getElementById('filtro-produto');
    const valFiltro = filtroProduto.value;
    filtroProduto.innerHTML = '<option value="Todos">Todos os Produtos</option>' +
      produtos.map(p => `<option value="${escapeHtml(p)}">${escapeHtml(p)}</option>`).join('');
    filtroProduto.value = valFiltro || 'Todos';
  }

  function popularDatalistMarcas() {
    // Combobox customizado
    const marcas = getMarcas();
    const input = document.getElementById('marca');
    const dropdown = document.getElementById('marca-dropdown');
    const arrow = document.getElementById('marca-arrow');

    function renderOptions(filtro) {
      const filtered = filtro
        ? marcas.filter(m => m.toLowerCase().includes(filtro.toLowerCase()))
        : marcas;

      if (filtered.length === 0) {
        dropdown.innerHTML = '<div class="combobox-no-result">Nenhuma marca encontrada</div>';
      } else {
        dropdown.innerHTML = filtered.map(m =>
          `<div class="combobox-option">${escapeHtml(m)}</div>`
        ).join('');
      }

      dropdown.querySelectorAll('.combobox-option').forEach(opt => {
        opt.addEventListener('mousedown', (e) => {
          e.preventDefault();
          input.value = opt.textContent;
          dropdown.classList.remove('open');
        });
      });
    }

    // Abrir/fechar com a seta
    arrow.addEventListener('click', (e) => {
      e.preventDefault();
      e.stopPropagation();
      if (dropdown.classList.contains('open')) {
        dropdown.classList.remove('open');
      } else {
        renderOptions(input.value);
        dropdown.classList.add('open');
        input.focus();
      }
    });

    // Filtrar ao digitar
    input.addEventListener('input', () => {
      renderOptions(input.value);
      dropdown.classList.add('open');
    });

    // Abrir ao focar
    input.addEventListener('focus', () => {
      renderOptions(input.value);
      dropdown.classList.add('open');
    });

    // Fechar ao clicar fora
    document.addEventListener('click', (e) => {
      if (!e.target.closest('.combobox-wrapper')) {
        dropdown.classList.remove('open');
      }
    });
  }

  // ── CONFIGURAÇÕES (só TI) ─────────────────────────────────────────────
  function carregarConfiguracoes() {
    renderMarcas();
    renderProdutos();
  }

  function renderMarcas() {
    const marcas = getMarcas();
    const container = document.getElementById('lista-marcas');
    if (marcas.length === 0) {
      container.innerHTML = '<p class="config-empty">Nenhuma marca cadastrada.</p>';
      return;
    }
    container.innerHTML = marcas.map(m => `
      <div class="config-tag">
        <i class="fas fa-tag" style="font-size:0.75rem; opacity:0.5;"></i>
        ${escapeHtml(m)}
        <button class="tag-remove" data-marca="${escapeHtml(m)}" title="Remover">
          <i class="fas fa-times"></i>
        </button>
      </div>
    `).join('');

    container.querySelectorAll('.tag-remove[data-marca]').forEach(btn => {
      btn.addEventListener('click', () => {
        const nome = btn.dataset.marca;
        let lista = getMarcas();
        lista = lista.filter(m => m !== nome);
        saveMarcas(lista);
        renderMarcas();
        popularDatalistMarcas();
        showToast(`Marca "${nome}" removida.`);
      });
    });
  }

  function renderProdutos() {
    const produtos = getProdutos();
    const container = document.getElementById('lista-produtos');
    if (produtos.length === 0) {
      container.innerHTML = '<p class="config-empty">Nenhum produto cadastrado.</p>';
      return;
    }
    container.innerHTML = produtos.map(p => `
      <div class="config-tag">
        <i class="fas fa-desktop" style="font-size:0.75rem; opacity:0.5;"></i>
        ${escapeHtml(p)}
        <button class="tag-remove" data-produto="${escapeHtml(p)}" title="Remover">
          <i class="fas fa-times"></i>
        </button>
      </div>
    `).join('');

    container.querySelectorAll('.tag-remove[data-produto]').forEach(btn => {
      btn.addEventListener('click', () => {
        const nome = btn.dataset.produto;
        let lista = getProdutos();
        lista = lista.filter(p => p !== nome);
        saveProdutos(lista);
        renderProdutos();
        popularSelectsProdutos();
        showToast(`Produto "${nome}" removido.`);
      });
    });
  }

  // Adicionar marca
  document.getElementById('btn-add-marca').addEventListener('click', () => {
    const input = document.getElementById('input-nova-marca');
    const nome = input.value.trim();
    if (!nome) { showToast('Digite o nome da marca.', true); return; }
    const lista = getMarcas();
    if (lista.some(m => m.toLowerCase() === nome.toLowerCase())) {
      showToast('Essa marca já existe.', true); return;
    }
    lista.push(nome);
    saveMarcas(lista);
    input.value = '';
    renderMarcas();
    popularDatalistMarcas();
    showToast(`Marca "${nome}" adicionada!`);
  });

  document.getElementById('input-nova-marca').addEventListener('keydown', (e) => {
    if (e.key === 'Enter') { e.preventDefault(); document.getElementById('btn-add-marca').click(); }
  });

  // Adicionar produto
  document.getElementById('btn-add-produto').addEventListener('click', () => {
    const input = document.getElementById('input-novo-produto');
    const nome = input.value.trim();
    if (!nome) { showToast('Digite o nome do produto.', true); return; }
    const lista = getProdutos();
    if (lista.some(p => p.toLowerCase() === nome.toLowerCase())) {
      showToast('Esse produto já existe.', true); return;
    }
    lista.push(nome);
    saveProdutos(lista);
    input.value = '';
    renderProdutos();
    popularSelectsProdutos();
    showToast(`Produto "${nome}" adicionado!`);
  });

  document.getElementById('input-novo-produto').addEventListener('keydown', (e) => {
    if (e.key === 'Enter') { e.preventDefault(); document.getElementById('btn-add-produto').click(); }
  });

  // ── IMPORTAR EXCEL (.xlsx) — só TI ──────────────────────────────────────
  let dadosImportados = [];

  const inputImport = document.getElementById('input-import-excel');
  const importFileName = document.getElementById('import-file-name');
  const importPreview = document.getElementById('import-preview');
  const importCount = document.getElementById('import-count');
  const importTableBody = document.getElementById('import-table-body');
  const btnConfirmarImport = document.getElementById('btn-confirmar-import');
  const btnCancelarImport = document.getElementById('btn-cancelar-import');
  const btnModeloExcel = document.getElementById('btn-modelo-excel');

  inputImport.addEventListener('change', async (e) => {
    if (!isTI()) { showToast('Apenas a TI pode importar.', true); return; }
    const file = e.target.files[0];
    if (!file) return;

    importFileName.textContent = file.name;

    try {
      const buffer = await file.arrayBuffer();
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.load(buffer);
      const ws = wb.worksheets[0];
      if (!ws) { showToast('Planilha vazia.', true); return; }

      dadosImportados = [];
      let headerRow = null;

      // Encontrar linha de cabeçalho
      ws.eachRow((row, rowNum) => {
        if (headerRow) return;
        const valores = [];
        row.eachCell((cell) => {
          valores.push(String(cell.value || '').toLowerCase().trim());
        });
        const textoLinha = valores.join(' ');
        if (textoLinha.includes('tombamento') || textoLinha.includes('produto')) {
          headerRow = rowNum;
        }
      });

      if (!headerRow) headerRow = 1;

      // Mapear colunas pelo cabeçalho
      const header = [];
      ws.getRow(headerRow).eachCell((cell, colNum) => {
        const val = String(cell.value || '').toLowerCase().trim();
        header[colNum] = val;
      });

      function findCol(...keywords) {
        for (let col = 1; col < header.length; col++) {
          if (!header[col]) continue;
          for (const kw of keywords) {
            if (header[col].includes(kw)) return col;
          }
        }
        return null;
      }

      const colTomb = findCol('tombamento', 'tomb', 'nº tomb', 'numero tomb');
      const colProd = findCol('produto', 'tipo', 'equipamento');
      const colMarca = findCol('marca');
      const colSerie = findCol('serie', 'série', 'numero de serie', 'número de série');
      const colSetor = findCol('setor', 'departamento');
      const colStatus = findCol('status', 'situação', 'situacao');

      if (!colProd && !colTomb) {
        showToast('Não foi possível identificar as colunas. Verifique o cabeçalho.', true);
        return;
      }

      // Ler dados a partir da linha seguinte ao cabeçalho
      for (let r = headerRow + 1; r <= ws.rowCount; r++) {
        const row = ws.getRow(r);
        const tomb = colTomb ? String(row.getCell(colTomb).value || '').replace(/[^0-9]/g, '') : '';
        const prod = colProd ? String(row.getCell(colProd).value || '').trim() : '';
        const marca = colMarca ? String(row.getCell(colMarca).value || '').trim() : '';
        const serie = colSerie ? String(row.getCell(colSerie).value || '').trim() : '';
        const setor = colSetor ? String(row.getCell(colSetor).value || '').trim() : 'TI';
        const status = colStatus ? String(row.getCell(colStatus).value || '').trim() : 'Ativo';

        // Pular linhas vazias
        if (!prod && !marca && !serie && !tomb) continue;

        dadosImportados.push({
          tombamento: tomb ? Number(tomb) : 0,
          produto: prod,
          marca: marca,
          serie: serie,
          setor: setor || 'TI',
          status: (status.toLowerCase().includes('baixa') ? 'Em Baixa' : 'Ativo')
        });
      }

      if (dadosImportados.length === 0) {
        showToast('Nenhum dado encontrado na planilha.', true);
        return;
      }

      // Mostrar preview
      importCount.textContent = dadosImportados.length;
      importTableBody.innerHTML = dadosImportados.map(d => `
        <tr>
          <td>${d.tombamento || '(auto)'}</td>
          <td>${escapeHtml(d.produto)}</td>
          <td>${escapeHtml(d.marca)}</td>
          <td>${escapeHtml(d.serie)}</td>
          <td>${escapeHtml(d.setor)}</td>
          <td><span class="badge ${d.status === 'Ativo' ? 'badge-ativo' : 'badge-baixa'}">${d.status}</span></td>
        </tr>`).join('');
      importPreview.style.display = '';

    } catch (err) {
      showToast('Erro ao ler o arquivo: ' + err.message, true);
    }

    // Reset input para poder selecionar o mesmo arquivo de novo
    inputImport.value = '';
  });

  btnCancelarImport.addEventListener('click', () => {
    dadosImportados = [];
    importPreview.style.display = 'none';
    importFileName.textContent = 'Nenhum arquivo selecionado';
  });

  btnConfirmarImport.addEventListener('click', () => {
    if (!isTI()) { showToast('Apenas a TI pode importar.', true); return; }
    if (dadosImportados.length === 0) return;

    const lista = getTombamentos();
    let nextId = lista.length > 0 ? Math.max(...lista.map(t => t.id)) + 1 : 1;
    let nextTomb = proximoTombamento();
    let importados = 0;
    let duplicados = 0;

    const seriesExistentes = new Set(lista.map(t => t.numero_serie.toLowerCase()));

    dadosImportados.forEach(d => {
      // Verificar duplicado por número de série
      if (d.serie && seriesExistentes.has(d.serie.toLowerCase())) {
        duplicados++;
        return;
      }

      const numTomb = d.tombamento > 0 ? d.tombamento : nextTomb++;
      // Se tombamento manual, atualizar nextTomb
      if (d.tombamento > 0 && d.tombamento >= nextTomb) {
        nextTomb = d.tombamento + 1;
      }

      lista.push({
        id: nextId++,
        numero_tombamento: numTomb,
        produto: d.produto,
        marca: d.marca,
        numero_serie: d.serie,
        status: d.status,
        setor: d.setor,
        data_cadastro: new Date().toISOString()
      });

      if (d.serie) seriesExistentes.add(d.serie.toLowerCase());

      addHistorico({
        tombamento: numTomb,
        tipo: 'Importação',
        de: '—',
        para: d.setor,
        por: 'TI',
        data: new Date().toISOString()
      });

      importados++;
    });

    saveTombamentos(lista);

    let msg = `${importados} tombamento${importados !== 1 ? 's' : ''} importado${importados !== 1 ? 's' : ''} com sucesso!`;
    if (duplicados > 0) msg += ` (${duplicados} duplicado${duplicados !== 1 ? 's' : ''} ignorado${duplicados !== 1 ? 's' : ''})`;
    showToast(msg);

    dadosImportados = [];
    importPreview.style.display = 'none';
    importFileName.textContent = 'Nenhum arquivo selecionado';
    popularSelectsProdutos();
    carregarDashboard();
  });

  // Baixar modelo de planilha
  btnModeloExcel.addEventListener('click', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Modelo');

    ws.columns = [
      { header: 'Tombamento', key: 'tomb', width: 16 },
      { header: 'Produto', key: 'prod', width: 20 },
      { header: 'Marca', key: 'marca', width: 22 },
      { header: 'Número de Série', key: 'serie', width: 28 },
      { header: 'Setor', key: 'setor', width: 20 },
      { header: 'Status', key: 'status', width: 16 }
    ];

    // Estilo do cabeçalho
    ws.getRow(1).eachCell(cell => {
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2563EB' } };
      cell.alignment = { horizontal: 'center' };
    });

    // Linha exemplo
    ws.addRow({ tomb: 501, prod: 'Monitor', marca: 'Samsung', serie: 'SN123456', setor: 'TI', status: 'Ativo' });
    ws.addRow({ tomb: 502, prod: 'CPU', marca: 'Dell', serie: 'DL789012', setor: 'Guias', status: 'Ativo' });

    const buffer = await wb.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'modelo_tombamentos.xlsx';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  });

  // ═════════════════════════════════════════════════════════════════════════
  // LOGS DE ATIVIDADE (só TI)
  // ═════════════════════════════════════════════════════════════════════════

  const NIVEL_CONFIG = {
    info:    { icone: 'fa-circle-info',    cor: '#3b82f6', bg: '#eff6ff',  label: 'Info' },
    sucesso: { icone: 'fa-circle-check',   cor: '#10b981', bg: '#ecfdf5',  label: 'Sucesso' },
    alerta:  { icone: 'fa-triangle-exclamation', cor: '#f59e0b', bg: '#fffbeb', label: 'Alerta' },
    erro:    { icone: 'fa-circle-xmark',   cor: '#ef4444', bg: '#fef2f2',  label: 'Erro' }
  };

  function carregarLogs() {
    const logs = getLogs();
    const logsLista = document.getElementById('logs-lista');
    const logsStats = document.getElementById('logs-stats');
    const filtroBusca = document.getElementById('filtro-log-busca');
    const filtroNivel = document.getElementById('filtro-log-nivel');
    const filtroSetor = document.getElementById('filtro-log-setor');

    // Popular filtro de setores
    const setores = [...new Set(logs.map(l => l.setor))].sort();
    filtroSetor.innerHTML = '<option value="Todos">Todos os setores</option>' +
      setores.map(s => `<option value="${escapeHtml(s)}">${escapeHtml(s)}</option>`).join('');

    function renderLogs() {
      const busca = filtroBusca.value.toLowerCase();
      const nivel = filtroNivel.value;
      const setor = filtroSetor.value;

      let filtrados = logs.filter(l => {
        if (nivel !== 'Todos' && l.nivel !== nivel) return false;
        if (setor !== 'Todos' && l.setor !== setor) return false;
        if (busca && !(l.acao + l.detalhes + l.setor).toLowerCase().includes(busca)) return false;
        return true;
      });

      // Stats
      const contagem = { info: 0, sucesso: 0, alerta: 0, erro: 0 };
      filtrados.forEach(l => { if (contagem[l.nivel] !== undefined) contagem[l.nivel]++; });

      logsStats.innerHTML = Object.entries(NIVEL_CONFIG).map(([key, cfg]) => `
        <div class="log-stat-chip" style="background:${cfg.bg};color:${cfg.cor};border:1.5px solid ${cfg.cor}20;">
          <i class="fas ${cfg.icone}"></i>
          <span>${contagem[key]}</span>
          <small>${cfg.label}</small>
        </div>
      `).join('');

      if (filtrados.length === 0) {
        logsLista.innerHTML = `
          <div style="text-align:center;padding:40px;color:var(--cinza-400);">
            <i class="fas fa-inbox" style="font-size:2rem;margin-bottom:10px;display:block;"></i>
            <p>Nenhum log encontrado.</p>
          </div>`;
        return;
      }

      // Agrupar por data
      const grupos = {};
      filtrados.forEach(l => {
        const d = new Date(l.data);
        const chave = d.toLocaleDateString('pt-BR', { weekday: 'long', day: '2-digit', month: 'long', year: 'numeric' });
        if (!grupos[chave]) grupos[chave] = [];
        grupos[chave].push(l);
      });

      logsLista.innerHTML = Object.entries(grupos).map(([data, items]) => `
        <div class="log-grupo">
          <div class="log-grupo-data"><i class="fas fa-calendar-day"></i> ${data}</div>
          ${items.map(l => {
            const cfg = NIVEL_CONFIG[l.nivel] || NIVEL_CONFIG.info;
            const hora = new Date(l.data).toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit', second: '2-digit' });
            return `
              <div class="log-item" style="border-left:3px solid ${cfg.cor};">
                <div class="log-item-icon" style="background:${cfg.bg};color:${cfg.cor};">
                  <i class="fas ${cfg.icone}"></i>
                </div>
                <div class="log-item-body">
                  <div class="log-item-acao">${escapeHtml(l.acao)}</div>
                  <div class="log-item-detalhes">${escapeHtml(l.detalhes)}</div>
                </div>
                <div class="log-item-meta">
                  <span class="log-item-setor"><i class="fas fa-building"></i> ${escapeHtml(l.setor)}</span>
                  <span class="log-item-hora"><i class="fas fa-clock"></i> ${hora}</span>
                </div>
              </div>`;
          }).join('')}
        </div>
      `).join('');
    }

    filtroBusca.addEventListener('input', debounce(renderLogs, 300));
    filtroNivel.addEventListener('change', renderLogs);
    filtroSetor.addEventListener('change', renderLogs);
    renderLogs();
  }

  // Limpar logs
  document.getElementById('btn-limpar-logs').addEventListener('click', () => {
    if (!confirm('Deseja limpar todos os logs de atividade?')) return;
    localStorage.removeItem(LOGS_KEY);
    addLog('Logs limpos', 'Todos os logs anteriores foram apagados', 'alerta');
    carregarLogs();
    showToast('Logs limpos com sucesso!');
  });
});
