// ═══════════════════════════════════════════════════════════════════════════
// ASTIR – Sistema de Tombamento
// Modos:
//   file://  → localStorage (padrão, sem servidor)
//   http://  → API Flask + SQLite (window.MODO_API = true injetado pelo server.py)
// Permissões:
//   TI       → tombar, dar baixa, excluir, transferir tudo
//   Setores  → apenas transferir itens que estão NO SEU setor
// ═══════════════════════════════════════════════════════════════════════════

const STORAGE_KEY = 'tombamentos_db';
const USERS_KEY = 'tombamento_users';
const SESSION_KEY = 'tombamento_session';
const HISTORICO_KEY = 'tombamento_historico';
const MARCAS_KEY = 'tombamento_marcas';
const MATERIAIS_KEY = 'tombamento_materiais';
const LOGS_KEY = 'tombamento_logs';
const TOMBAMENTO_INICIO = 0;

// Catálogos padrão
const MARCAS_DEFAULT = ['Samsung', 'LG', 'Dell', 'HP', 'Lenovo', 'TP-Link', 'Intelbras', 'AOC', 'Multilaser', 'Positivo'];
const MATERIAIS_DEFAULT = ['COMPUTADOR', 'MONITOR', 'NOBREAK', 'IMPRESSORA', 'DVR', 'SWITCH', 'CÂMERA', 'RACK', 'ROTEADOR', 'AR CONDICIONADO'];

if (!localStorage.getItem(MARCAS_KEY)) {
  localStorage.setItem(MARCAS_KEY, JSON.stringify(MARCAS_DEFAULT));
}
if (!localStorage.getItem(MATERIAIS_KEY)) {
  localStorage.setItem(MATERIAIS_KEY, JSON.stringify(MATERIAIS_DEFAULT));
}

// ── Migração: produto → material ──────────────────────────────────────
(function migrateToMaterial() {
  const tombs = JSON.parse(localStorage.getItem(STORAGE_KEY) || '[]');
  let changed = false;
  tombs.forEach(t => {
    if (t.produto && !t.material) { t.material = t.produto; delete t.produto; changed = true; }
    if (t.nome && !t.descricao) { t.descricao = t.nome; delete t.nome; changed = true; }
    if (t.detalhes && !t.descricao) { t.descricao = t.detalhes; delete t.detalhes; changed = true; }
    if (!t.cor) { t.cor = ''; changed = true; }
    if (!t.modelo) { t.modelo = ''; changed = true; }
    if (!t.processo) { t.processo = ''; changed = true; }
    if (!t.pat) { t.pat = ''; changed = true; }
  });
  if (changed) localStorage.setItem(STORAGE_KEY, JSON.stringify(tombs));
  // Migrar PRODUTOS_KEY → MATERIAIS_KEY
  if (localStorage.getItem('tombamento_produtos') && !localStorage.getItem(MATERIAIS_KEY)) {
    localStorage.setItem(MATERIAIS_KEY, localStorage.getItem('tombamento_produtos'));
    localStorage.removeItem('tombamento_produtos');
  }
})();

function getMarcas() {
  return JSON.parse(localStorage.getItem(MARCAS_KEY) || '[]').sort((a, b) => a.localeCompare(b));
}

function saveMarcas(lista) {
  localStorage.setItem(MARCAS_KEY, JSON.stringify(lista));
  _syncBackground('/api/sync/marcas', lista);
}

function getMateriais() {
  return JSON.parse(localStorage.getItem(MATERIAIS_KEY) || '[]').sort((a, b) => a.localeCompare(b));
}

function saveMateriais(lista) {
  localStorage.setItem(MATERIAIS_KEY, JSON.stringify(lista));
  _syncBackground('/api/sync/materiais', lista);
}

// Setores e senhas padrão
const SETORES_DEFAULT = {
  'TI':                    'ti123',
  'Almoxarifado':          'almoxarifado123',
  'Ambulatorio':           'ambulatorio123',
  'ARRECADAÇÃO':           'arrecadacao123',
  'AUDITORIA':             'auditoria123',
  'Audiologia':            'audiologia123',
  'Cadastro':              'cadastro123',
  'CCIH':                  'ccih123',
  'Compras':               'compras123',
  'Cozinha':               'cozinha123',
  'Descartado':            'descartado123',
  'DIREX':                 'direx123',
  'DIRETOR':               'diretor123',
  'Enfermagem':            'enfermagem123',
  'Farmacia':              'farmacia123',
  'Faturamento':           'faturamento123',
  'Financeiro':            'financeiro123',
  'Fisioterapia':          'fisioterapia123',
  'Guias':                 'guias123',
  'INTERNAÇÃO':            'internacao123',
  'Juridico':              'juridico123',
  'Odontologia':           'odontologia123',
  'Polos':                 'polos123',
  'Psicologia':            'psicologia123',
  'Recepcao':              'recepcao123',
  'RH':                    'rh123',
  'SECONF':                'seconf123',
  'Segurança do Trabalho': 'seguranca123',
  'Servico Social':        'servicosocial123',
  'SPA':                   'spa123',
  'Não Catalogado':        'naocatalogado123'
};

if (!localStorage.getItem(USERS_KEY)) {
  console.log('[TOMBAMENTO] Inicializando setores padrão no localStorage.');
  localStorage.setItem(USERS_KEY, JSON.stringify(SETORES_DEFAULT));
} else {
  // Garante que TI sempre exista nas senhas salvas
  const saved = JSON.parse(localStorage.getItem(USERS_KEY));
  if (!('TI' in saved)) {
    saved['TI'] = SETORES_DEFAULT['TI'];
    localStorage.setItem(USERS_KEY, JSON.stringify(saved));
  }
}

// ── Banco ────────────────────────────────────────────────────────────────
// Cache em memória — evita JSON.parse repetido a cada chamada
let _tombamentosCache = null;
function getTombamentos() {
  if (_tombamentosCache) return _tombamentosCache;
  const data = localStorage.getItem(STORAGE_KEY);
  _tombamentosCache = data ? JSON.parse(data) : [];
  return _tombamentosCache;
}

function saveTombamentos(lista) {
  _tombamentosCache = lista;
  localStorage.setItem(STORAGE_KEY, JSON.stringify(lista));
  _syncBackground('/api/sync/tombamentos', lista);
}

function getHistorico() {
  const data = localStorage.getItem(HISTORICO_KEY);
  return data ? JSON.parse(data) : [];
}

function addHistorico(entry) {
  const hist = getHistorico();
  hist.unshift(entry);
  localStorage.setItem(HISTORICO_KEY, JSON.stringify(hist));
  _syncBackground('/api/historico', entry);
}

// ── Logs de Atividade ────────────────────────────────────────────────────
function getLogs() {
  const data = localStorage.getItem(LOGS_KEY);
  return data ? JSON.parse(data) : [];
}

function addLog(acao, detalhes, nivel) {
  const entry = {
    id: Date.now(),
    data: new Date().toISOString(),
    setor: meuSetor() || 'Sistema',
    acao: acao,
    detalhes: detalhes || '',
    nivel: nivel || 'info'
  };
  const logs = getLogs();
  logs.unshift(entry);
  if (logs.length > 500) logs.length = 500;
  localStorage.setItem(LOGS_KEY, JSON.stringify(logs));
  _syncBackground('/api/logs', entry);
}

// ── Modo API: sincronização em background com Flask/SQLite ──────────────
function _syncBackground(url, data) {
  if (!window.MODO_API) return;
  fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(data)
  }).catch(err => console.warn('[SYNC]', url, err.message));
}

async function loadFromServer() {
  const [tombamentos, setoresArr, marcas, materiais, historicoArr] = await Promise.all([
    fetch('/api/tombamentos').then(r => r.json()),
    fetch('/api/setores').then(r => r.json()),
    fetch('/api/marcas').then(r => r.json()),
    fetch('/api/materiais').then(r => r.json()),
    fetch('/api/historico').then(r => r.json())
  ]);
  _tombamentosCache = null; // invalida cache antes de reescrever
  localStorage.setItem(STORAGE_KEY, JSON.stringify(tombamentos));
  const usersObj = Object.fromEntries(setoresArr.map(s => [s.nome, s.senha]));
  localStorage.setItem(USERS_KEY, JSON.stringify(usersObj));
  localStorage.setItem(MARCAS_KEY, JSON.stringify(marcas));
  localStorage.setItem(MATERIAIS_KEY, JSON.stringify(materiais));
  localStorage.setItem(HISTORICO_KEY, JSON.stringify(historicoArr));
  console.log('[ASTIR] Dados carregados do servidor.');
}

function proximoTombamento() {
  const lista = getTombamentos();
  if (lista.length === 0) return 1;
  return lista.reduce((m, t) => t.numero_tombamento > m ? t.numero_tombamento : m, 0) + 1;
}

function gerarId() {
  const lista = getTombamentos();
  if (lista.length === 0) return 1;
  return lista.reduce((m, t) => t.id > m ? t.id : m, 0) + 1;
}

// escapeHtml via regex — sem criar elemento DOM a cada chamada
const _escMap = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' };
function escapeHtml(text) {
  if (text == null) return '';
  return String(text).replace(/[&<>"']/g, c => _escMap[c]);
}

// ── Sessão ───────────────────────────────────────────────────────────────
function getSession() {
  const s = sessionStorage.getItem(SESSION_KEY);
  return s ? JSON.parse(s) : null;
}

function setSession(setor) {
  sessionStorage.setItem(SESSION_KEY, JSON.stringify({ setor, loginTime: new Date().toISOString() }));
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
  if (sessao) {
    if (window.MODO_API) {
      loadFromServer()
        .then(() => mostrarApp(sessao.setor))
        .catch(() => mostrarApp(sessao.setor)); // usa dados locais se servidor indisponível
    } else {
      mostrarApp(sessao.setor);
    }
  }

  // ── LOGIN ──────────────────────────────────────────────────────────────
  const formLogin = document.getElementById('form-login');
  const loginError = document.getElementById('login-error');

  formLogin.addEventListener('submit', async (e) => {
    e.preventDefault();
    const setor = document.getElementById('login-setor').value;
    const senha = document.getElementById('login-senha').value;

    if (window.MODO_API) {
      loginError.textContent = 'Verificando...';
      try {
        const res = await fetch('/api/login', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ setor, senha })
        });
        if (!res.ok) {
          const err = await res.json().catch(() => ({}));
          loginError.textContent = err.error || 'Setor ou senha incorretos.';
          return;
        }
        await loadFromServer();
        loginError.textContent = '';
        setSession(setor);
        addLog('Login', `Setor ${setor} entrou no sistema`, 'info');
        mostrarApp(setor);
      } catch(err) {
        loginError.textContent = 'Erro ao conectar ao servidor.';
      }
      return;
    }

    // modo localStorage
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
    const filtroSetorGroup = document.getElementById('filtro-setor').closest('.filtro-group');
    if (isTI()) {
      btnAdd.style.display = '';
      navConfig.style.display = '';
      navLogs.style.display = '';
      if (filtroSetorGroup) filtroSetorGroup.style.display = '';
    } else {
      btnAdd.style.display = 'none';
      navConfig.style.display = 'none';
      navLogs.style.display = 'none';
      // Não-TI: esconder filtro de setor (já filtra automaticamente)
      if (filtroSetorGroup) filtroSetorGroup.style.display = 'none';
    }

    popularSelectsMateriais();
    popularDatalistMarcas();
    carregarDashboard();
  }

  // Popular login setor — combobox com autocomplete
  function popularLoginSetores() {
    const users = JSON.parse(localStorage.getItem(USERS_KEY) || '{}');
    const setores = Object.keys(users).sort((a, b) => a.localeCompare(b, 'pt-BR'));
    const input    = document.getElementById('login-setor');
    const dropdown = document.getElementById('login-setor-dropdown');
    const arrow    = document.getElementById('login-setor-arrow');

    function renderOpts(filtro) {
      const f = filtro.toLowerCase();
      const filtered = f ? setores.filter(s => s.toLowerCase().includes(f)) : setores;
      dropdown.innerHTML = filtered.length === 0
        ? '<div class="combobox-no-result">Nenhum setor encontrado</div>'
        : filtered.map(s => `<div class="combobox-option">${escapeHtml(s)}</div>`).join('');
      dropdown.querySelectorAll('.combobox-option').forEach(opt => {
        opt.addEventListener('mousedown', e => {
          e.preventDefault();
          input.value = opt.textContent;
          dropdown.classList.remove('open');
        });
      });
    }

    arrow.onclick = e => {
      e.preventDefault(); e.stopPropagation();
      dropdown.classList.contains('open')
        ? dropdown.classList.remove('open')
        : (renderOpts(input.value), dropdown.classList.add('open'), input.focus());
    };
    input.oninput  = () => { renderOpts(input.value); dropdown.classList.add('open'); };
    input.onfocus  = () => { renderOpts(input.value); dropdown.classList.add('open'); };
    document.addEventListener('click', e => {
      if (!input.closest('.combobox-wrapper').contains(e.target)) dropdown.classList.remove('open');
    }, { capture: true });
  }
  popularLoginSetores();

  // Logout
  document.getElementById('btn-logout').addEventListener('click', () => {
    addLog('Logout', `Setor ${meuSetor()} saiu do sistema`, 'info');
    clearSession();
    if (window.MODO_API) {
      fetch('/api/logout', { method: 'POST' }).catch(() => {});
    }
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

  // ── ATALHOS DASHBOARD (stat-cards clicáveis) ─────────────────────────
  window.irParaListagem = function(filtroStatus) {
    // Navega para listagem e aplica filtro de status
    navItems.forEach(n => n.classList.remove('active'));
    sections.forEach(s => s.classList.remove('active'));
    const navListagem = document.querySelector('[data-section="listagem"]');
    if (navListagem) navListagem.classList.add('active');
    document.getElementById('section-listagem').classList.add('active');
    const selectStatus = document.getElementById('filtro-status');
    if (selectStatus) selectStatus.value = filtroStatus === 'Todos' ? 'Todos' : filtroStatus;
    carregarListagem();
  };

  window.abrirCadastroRapido = function() {
    if (!isTI()) return;
    // Navega para listagem e abre o modal de cadastro
    navItems.forEach(n => n.classList.remove('active'));
    sections.forEach(s => s.classList.remove('active'));
    const navListagem = document.querySelector('[data-section="listagem"]');
    if (navListagem) navListagem.classList.add('active');
    document.getElementById('section-listagem').classList.add('active');
    carregarListagem();
    abrirModalCadastro();
  };

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

    const material = document.getElementById('cad-material').value;
    const marca = document.getElementById('cad-marca').value.trim();
    const modelo = document.getElementById('cad-modelo').value.trim();
    const cor = document.getElementById('cad-cor').value.trim();
    const numero_serie = document.getElementById('cad-numero_serie').value.trim();
    const processo = document.getElementById('cad-processo').value.trim();
    const descricao = document.getElementById('cad-descricao').value.trim();
    const setorDestino = document.getElementById('cad-setor').value || 'TI';
    if (!material || !marca) return;

    const lista = getTombamentos();
    const num = proximoTombamento();
    const novo = {
      id: gerarId(),
      numero_tombamento: num,
      pat: '',
      material,
      cor: cor || '',
      descricao: descricao || '',
      marca,
      modelo: modelo || '',
      setor: setorDestino,
      numero_serie: numero_serie || '',
      status: 'Ativo',
      processo: processo || '',
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

    addLog('Cadastro', `Tombamento #${num} — ${material} ${marca} (Série: ${numero_serie})`, 'sucesso');
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

  // Event delegation — 1 listener na tabela inteira em vez de 1 por botão
  tabelaBody.addEventListener('click', e => {
    const btn = e.target.closest('button[data-toggle-id], button[data-edit-id], button[data-delete-id], button[data-transfer-id], button[data-hist-id]');
    if (!btn) return;
    if (btn.dataset.toggleId)   return toggleStatus(Number(btn.dataset.toggleId));
    if (btn.dataset.editId)     return abrirModalEditar(Number(btn.dataset.editId));
    if (btn.dataset.deleteId)   return confirmarExclusao(Number(btn.dataset.deleteId), btn.dataset.deleteNum);
    if (btn.dataset.transferId) return abrirModalTransferir(Number(btn.dataset.transferId));
    if (btn.dataset.histId)     return abrirHistorico(Number(btn.dataset.histNum));
  });

  let ordemAsc = true; // true = crescente (#1→#305), false = decrescente (#305→#1)

  window.toggleOrdem = function() {
    ordemAsc = !ordemAsc;
    const icone = document.getElementById('icone-ordem');
    if (icone) {
      icone.className = ordemAsc ? 'fas fa-arrow-up-short-wide' : 'fas fa-arrow-down-short-wide';
    }
    carregarListagem();
  };

  function carregarListagem() {
    // Atualizar contagens nos filtros
    popularSelectsMateriais();

    const busca = document.getElementById('filtro-busca').value.toLowerCase();
    const filtroMaterial = document.getElementById('filtro-material').value;
    const filtroStatus = document.getElementById('filtro-status').value;
    const filtroSetor = document.getElementById('filtro-setor').value;
    const setor = meuSetor();
    const tiUser = isTI();

    let lista = getTombamentos();

    // Não-TI: só vê seu próprio setor
    if (!tiUser) {
      lista = lista.filter(t => t.setor === setor);
    } else if (filtroSetor !== 'Todos') {
      lista = lista.filter(t => t.setor === filtroSetor);
    }

    if (filtroMaterial !== 'Todos') {
      lista = lista.filter(t => t.material === filtroMaterial);
    }
    if (filtroStatus !== 'Todos') {
      if (filtroStatus === 'Recentes') {
        const limite = new Date();
        limite.setDate(limite.getDate() - 30);
        lista = lista.filter(t => t.data_cadastro && new Date(t.data_cadastro) >= limite);
      } else {
        lista = lista.filter(t => t.status === filtroStatus);
      }
    }
    if (busca) {
      lista = lista.filter(t =>
        (t.material || '').toLowerCase().includes(busca) ||
        (t.marca || '').toLowerCase().includes(busca) ||
        (t.numero_serie || '').toLowerCase().includes(busca) ||
        (t.modelo || '').toLowerCase().includes(busca) ||
        String(t.numero_tombamento).includes(busca) ||
        (t.setor && t.setor.toLowerCase().includes(busca)) ||
        (t.pat || '').toLowerCase().includes(busca)
      );
    }

    // Resumo do setor filtrado
    const resumoEl = document.getElementById('resumo-setor');
    if (filtroSetor !== 'Todos') {
      const totalSetor = lista.length;
      const ativosSetor = lista.filter(t => t.status === 'Ativo').length;
      const baixaSetor = lista.filter(t => t.status === 'Em Baixa').length;
      const porMaterial = {};
      const marcas = new Set();
      const tombamentos = [];
      lista.forEach(t => {
        porMaterial[t.material] = (porMaterial[t.material] || 0) + 1;
        marcas.add(t.marca);
        tombamentos.push(`#${t.numero_tombamento}`);
      });

      const materiaisList = Object.entries(porMaterial)
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
          <div class="resumo-linha"><span class="resumo-label">Materiais:</span>${materiaisList}</div>
          <div class="resumo-linha"><span class="resumo-label">Marcas:</span>${marcasList}</div>
          <div class="resumo-linha"><span class="resumo-label">Tombamentos:</span><span class="resumo-tombs">${tombamentos.join(', ')}</span></div>
        </div>`;
      resumoEl.style.display = '';
    } else {
      resumoEl.style.display = 'none';
    }

    lista.sort((a, b) => a.numero_tombamento - b.numero_tombamento);

    if (lista.length === 0) {
      tabelaBody.innerHTML = '';
      emptyState.style.display = 'block';
      // Remover botão "carregar mais" anterior
      const oldBtn = document.getElementById('btn-carregar-mais');
      if (oldBtn) oldBtn.remove();
      return;
    }

    emptyState.style.display = 'none';

    // ── PAGINAÇÃO — renderizar em lotes de PAGE_SIZE ──────────────
    const PAGE_SIZE = 50;
    let paginaAtual = 0;
    const totalPaginas = Math.ceil(lista.length / PAGE_SIZE);

    function renderBatch(pagina) {
      const inicio = pagina * PAGE_SIZE;
      const fim = Math.min(inicio + PAGE_SIZE, lista.length);
      const lote = lista.slice(inicio, fim);

      const html = lote.map(item => {
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
          <button class="btn-icon btn-edit" title="Editar" data-edit-id="${item.id}">
            <i class="fas fa-pen"></i>
          </button>
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
          <td class="tombamento-cell">${item.pat ? escapeHtml(item.pat) : '#' + item.numero_tombamento}</td>
          <td>
            <div class="produto-setor-cell">
              <span class="produto-nome">${escapeHtml(item.material)}</span>
              ${item.descricao ? `<span class="produto-apelido" title="${escapeHtml(item.descricao)}"><i class="fas fa-circle-info" style="font-size:.7rem;"></i> ${escapeHtml(item.descricao.substring(0, 50))}${item.descricao.length > 50 ? '...' : ''}</span>` : ''}
              <span class="produto-setor-tag"><i class="fas fa-building"></i> ${escapeHtml(item.setor || '—')}</span>
            </div>
          </td>
          <td>
            <div>
              <span>${escapeHtml(item.marca)}</span>
              ${item.modelo ? `<div style="font-size:.75rem;color:var(--cinza-400);">${escapeHtml(item.modelo)}</div>` : ''}
            </div>
          </td>
          <td>${escapeHtml(item.numero_serie || '—')}</td>
          <td>${escapeHtml(item.cor || '—')}</td>
          <td>
            <div class="status-cell">
              <span class="badge ${isAtivo ? 'badge-ativo' : 'badge-baixa'}">
                <i class="fas fa-${isAtivo ? 'check-circle' : 'arrow-down'}"></i>
                ${isAtivo ? 'Ativo' : 'Em Baixa'}
              </span>
              ${toggleHtml}
            </div>
          </td>
          <td><div class="acoes-cell">${acoesHtml}</div></td>
        </tr>`;
    }).join('');

      // Primeira página: substituir; demais: concatenar
      if (pagina === 0) {
        tabelaBody.innerHTML = html;
      } else {
        tabelaBody.insertAdjacentHTML('beforeend', html);
      }

      // (event delegation configurado na inicialização — sem bindings por linha)

      // Atualizar botão "carregar mais"
      let btnMais = document.getElementById('btn-carregar-mais');
      if (pagina + 1 < totalPaginas) {
        const restante = lista.length - fim;
        if (!btnMais) {
          btnMais = document.createElement('div');
          btnMais.id = 'btn-carregar-mais';
          btnMais.className = 'carregar-mais-container';
          tabelaBody.closest('.tabela-container').after(btnMais);
        }
        btnMais.innerHTML = `<button class="btn-carregar-mais" onclick="document.dispatchEvent(new Event('carregar-mais'))">
          <i class="fas fa-chevron-down"></i> Carregar mais ${Math.min(PAGE_SIZE, restante)} de ${restante} restantes
          <small>(mostrando ${fim} de ${lista.length})</small>
        </button>`;
        btnMais.style.display = '';
      } else if (btnMais) {
        btnMais.style.display = 'none';
      }
    }

    // Renderizar primeira página
    renderBatch(0);
    paginaAtual = 0;

    // Listener para "carregar mais"
    const carregarMaisHandler = () => {
      paginaAtual++;
      if (paginaAtual < totalPaginas) renderBatch(paginaAtual);
    };
    document.removeEventListener('carregar-mais', carregarMaisHandler);
    // Guardar referência para remover depois
    if (window._carregarMaisHandler) document.removeEventListener('carregar-mais', window._carregarMaisHandler);
    window._carregarMaisHandler = carregarMaisHandler;
    document.addEventListener('carregar-mais', carregarMaisHandler);
  }

  function formatarData(dateStr) {
    if (!dateStr) return '—';
    const d = new Date(dateStr);
    return d.toLocaleDateString('pt-BR') + ' ' +
      d.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });
  }

  // Filtros
  document.getElementById('filtro-busca').addEventListener('input', debounce(carregarListagem, 300));
  document.getElementById('filtro-material').addEventListener('change', carregarListagem);
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
  let excluirSetorNome = null;

  function confirmarExclusao(id, numero) {
    if (!isTI()) { showToast('Apenas a TI pode excluir.', true); return; }
    excluirId = id;
    modalMsg.textContent = `Deseja realmente excluir o tombamento #${numero}? Esta ação não pode ser desfeita.`;
    modalOverlay.classList.add('show');
  }

  function fecharModalOverlay() {
    modalOverlay.classList.remove('show');
    excluirId = null;
    excluirSetorNome = null;
    document.getElementById('modal-titulo').textContent = 'Confirmar Ação';
    modalConfirmar.className = 'btn-danger';
  }

  modalCancelarBtn.addEventListener('click', fecharModalOverlay);

  modalOverlay.addEventListener('click', (e) => {
    if (e.target === modalOverlay) fecharModalOverlay();
  });

  modalConfirmar.addEventListener('click', () => {
    if (excluirSetorNome) {
      // Excluir setor
      const users = getUsers();
      delete users[excluirSetorNome];
      saveUsers(users);
      const nomeRemovido = excluirSetorNome;
      fecharModalOverlay();
      renderSetores();
      popularSelectsSetores();
      if (typeof popularLoginSetores === 'function') popularLoginSetores();
      showToast(`Setor "${nomeRemovido}" removido.`);
      addLog('Config', `Setor ${nomeRemovido} removido do sistema`, 'alerta');
      return;
    }
    if (!excluirId || !isTI()) return;
    let lista = getTombamentos();
    const itemExcluir = lista.find(t => t.id === excluirId);
    lista = lista.filter(t => t.id !== excluirId);
    saveTombamentos(lista);
    if (itemExcluir) addLog('Exclusão', `#${itemExcluir.numero_tombamento} — ${itemExcluir.material} ${itemExcluir.marca} excluído`, 'erro');
    showToast('Tombamento excluído com sucesso!');
    carregarListagem();
    fecharModalOverlay();
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
        <strong>#${item.numero_tombamento}</strong> — ${escapeHtml(item.material)} ${escapeHtml(item.marca)}
        <br><small>Atualmente em: <strong>${escapeHtml(item.setor)}</strong></small>
      </div>`;
    transferirDestino.value = '';
    document.getElementById('transferir-justificativa').value = '';

    // TI: transferência instantânea sem justificativa
    if (isTI()) {
      document.getElementById('transferir-justif-group').style.display = 'none';
      document.getElementById('transferir-ti-note').style.display = 'flex';
    } else {
      document.getElementById('transferir-justif-group').style.display = '';
      document.getElementById('transferir-ti-note').style.display = 'none';
    }

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
    document.getElementById('transferir-justif-group').style.display = '';
    document.getElementById('transferir-ti-note').style.display = 'none';
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
    if (!isTI() && !justificativa) { showToast('Informe a justificativa da transferência.', true); return; }

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

  // ── EDITAR TOMBAMENTO (só TI) ─────────────────────────────────────────
  const modalEditar = document.getElementById('modal-editar');
  const formEditar = document.getElementById('form-editar');
  const btnFecharEditar = document.getElementById('btn-fechar-editar');
  const btnCancelarEditar = document.getElementById('btn-cancelar-editar');

  function setupEditCombobox() {
    const marcas = getMarcas();
    const input = document.getElementById('editar-marca');
    const dropdown = document.getElementById('editar-marca-dropdown');
    const arrow = document.getElementById('editar-marca-arrow');

    function renderOpts(filtro) {
      const filtered = filtro ? marcas.filter(m => m.toLowerCase().includes(filtro.toLowerCase())) : marcas;
      dropdown.innerHTML = filtered.length === 0
        ? '<div class="combobox-no-result">Nenhuma marca encontrada</div>'
        : filtered.map(m => `<div class="combobox-option">${escapeHtml(m)}</div>`).join('');
      dropdown.querySelectorAll('.combobox-option').forEach(opt => {
        opt.addEventListener('mousedown', (e) => { e.preventDefault(); input.value = opt.textContent; dropdown.classList.remove('open'); });
      });
    }

    arrow.onclick = (e) => { e.preventDefault(); e.stopPropagation();
      dropdown.classList.contains('open') ? dropdown.classList.remove('open') : (renderOpts(input.value), dropdown.classList.add('open'), input.focus());
    };
    input.oninput = () => { renderOpts(input.value); dropdown.classList.add('open'); };
    input.onfocus = () => { renderOpts(input.value); dropdown.classList.add('open'); };
  }

  function abrirModalEditar(id) {
    if (!isTI()) return;
    const lista = getTombamentos();
    const item = lista.find(t => t.id === id);
    if (!item) return;

    // Popular select de materiais
    const selectMat = document.getElementById('editar-material');
    const materiais = getMateriais();
    selectMat.innerHTML = '<option value="">Selecione o material...</option>' +
      materiais.map(p => `<option value="${escapeHtml(p)}"${p === item.material ? ' selected' : ''}>${escapeHtml(p)}</option>`).join('');
    // Se material não está na lista atual, adicionar como opção
    if (item.material && !materiais.includes(item.material)) {
      selectMat.innerHTML += `<option value="${escapeHtml(item.material)}" selected>${escapeHtml(item.material)}</option>`;
    }

    document.getElementById('editar-id').value = item.id;
    document.getElementById('editar-marca').value = item.marca || '';
    document.getElementById('editar-serie').value = item.numero_serie || '';
    document.getElementById('editar-setor').value = item.setor || 'TI';
    document.getElementById('editar-modelo').value = item.modelo || '';
    document.getElementById('editar-cor').value = item.cor || '';
    document.getElementById('editar-processo').value = item.processo || '';
    document.getElementById('editar-descricao').value = item.descricao || '';

    setupEditCombobox();
    modalEditar.classList.add('show');
  }

  function fecharModalEditar() {
    modalEditar.classList.remove('show');
    formEditar.reset();
    document.getElementById('editar-marca-dropdown').classList.remove('open');
  }

  btnFecharEditar.addEventListener('click', fecharModalEditar);
  btnCancelarEditar.addEventListener('click', fecharModalEditar);
  modalEditar.addEventListener('click', (e) => { if (e.target === modalEditar) fecharModalEditar(); });

  formEditar.addEventListener('submit', (e) => {
    e.preventDefault();
    if (!isTI()) return;
    const id = Number(document.getElementById('editar-id').value);
    const lista = getTombamentos();
    const item = lista.find(t => t.id === id);
    if (!item) return;

    const material = document.getElementById('editar-material').value;
    const marca = document.getElementById('editar-marca').value.trim();
    const serie = document.getElementById('editar-serie').value.trim();
    const setor = document.getElementById('editar-setor').value;
    const modelo = document.getElementById('editar-modelo').value.trim();
    const cor = document.getElementById('editar-cor').value.trim();
    const processo = document.getElementById('editar-processo').value.trim();
    const descricao = document.getElementById('editar-descricao').value.trim();

    if (!material || !marca || !setor) { showToast('Preencha os campos obrigatórios.', true); return; }

    item.material = material;
    item.marca = marca;
    item.numero_serie = serie;
    item.setor = setor;
    item.modelo = modelo;
    item.cor = cor;
    item.processo = processo;
    item.descricao = descricao;

    saveTombamentos(lista);
    addLog('Edição', `#${item.numero_tombamento} — ${material} ${marca} editado`, 'sucesso');
    showToast(`Tombamento #${item.numero_tombamento} atualizado!`);
    fecharModalEditar();
    carregarListagem();
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
    lista.forEach(t => { contagem[t.material] = (contagem[t.material] || 0) + 1; });

    const chartEl = document.getElementById('chart-produtos');
    const materiais = Object.entries(contagem).sort((a, b) => b[1] - a[1]);

    if (materiais.length === 0) {
      chartEl.innerHTML = '<p style="color:#94a3b8;text-align:center;padding:20px;">Nenhum equipamento cadastrado ainda.</p>';
      return;
    }

    const iconesMaterial = {
      'MONITOR': 'fa-desktop',
      'NOBREAK': 'fa-car-battery',
      'COMPUTADOR': 'fa-computer',
      'DVR': 'fa-video',
      'SWITCH': 'fa-network-wired',
      'MICROFONE': 'fa-microphone',
      'CÂMERA': 'fa-camera',
      'RACK': 'fa-server',
      'IMPRESSORA': 'fa-print',
      'ROTEADOR': 'fa-wifi',
      'TECLADO': 'fa-keyboard',
      'MOUSE': 'fa-computer-mouse',
      'AR CONDICIONADO': 'fa-snowflake',
      'TELEFONE': 'fa-phone',
      'TELEVISOR': 'fa-tv',
      'ESTABILIZADOR': 'fa-bolt'
    };

    chartEl.innerHTML = materiais.map(([nome, qtd]) => {
      const pct = total > 0 ? Math.round((qtd / total) * 100) : 0;
      const icone = iconesMaterial[nome.toUpperCase()] || 'fa-cube';
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

          // Detalhe por material
          const detalheMaterial = {};
          items.forEach(i => {
            detalheMaterial[i.material] = (detalheMaterial[i.material] || 0) + 1;
          });
          const detalhes = Object.entries(detalheMaterial)
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
    const filtroMaterial = document.getElementById('filtro-material').value;
    const filtroStatus = document.getElementById('filtro-status').value;
    const filtroSetor = document.getElementById('filtro-setor').value;
    const busca = document.getElementById('filtro-busca').value.toLowerCase();
    const tiUser = isTI();
    const setor = meuSetor();

    let lista = getTombamentos();

    // Não-TI: só exporta seu setor
    if (!tiUser) {
      lista = lista.filter(t => t.setor === setor);
    } else if (filtroSetor !== 'Todos') {
      lista = lista.filter(t => t.setor === filtroSetor);
    }

    if (filtroMaterial !== 'Todos') lista = lista.filter(t => t.material === filtroMaterial);
    if (filtroStatus !== 'Todos') {
      if (filtroStatus === 'Recentes') {
        const limite = new Date(); limite.setDate(limite.getDate() - 30);
        lista = lista.filter(t => t.data_cadastro && new Date(t.data_cadastro) >= limite);
      } else {
        lista = lista.filter(t => t.status === filtroStatus);
      }
    }
    if (busca) {
      lista = lista.filter(t =>
        (t.marca || '').toLowerCase().includes(busca) ||
        (t.numero_serie || '').toLowerCase().includes(busca) ||
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
      { key: 'pat',    width: 14 },
      { key: 'mat',    width: 20 },
      { key: 'marca',  width: 22 },
      { key: 'modelo', width: 28 },
      { key: 'serie',  width: 24 },
      { key: 'cor',    width: 14 },
      { key: 'setor',  width: 20 },
      { key: 'status', width: 16 },
      { key: 'processo', width: 16 }
    ];

    // ── Linha 1: Título grande
    ws.mergeCells('A1:I1');
    const cellTitulo = ws.getCell('A1');
    const subtitulo = filtroSetor !== 'Todos' ? ` — ${filtroSetor}` : '';
    cellTitulo.value = `ASTIR — Relatório de Tombamentos${subtitulo}`;
    cellTitulo.font = { name: 'Calibri', size: 18, bold: true, color: { argb: 'FF' + AZUL_ESCURO } };
    cellTitulo.alignment = { horizontal: 'center', vertical: 'middle' };
    cellTitulo.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + AZUL_CLARO } };
    ws.getRow(1).height = 42;

    // ── Linha 2: Subtítulo com data e totais
    ws.mergeCells('A2:I2');
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
    const headers = ['PAT', 'Material', 'Marca', 'Modelo', 'Nº Série', 'Cor', 'Setor', 'Status', 'Processo'];
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
    ws.autoFilter = { from: 'A4', to: 'I4' };

    // ── Dados
    lista.forEach((t, idx) => {
      const r = ws.getRow(5 + idx);
      r.getCell(1).value = t.pat || t.numero_tombamento;
      r.getCell(2).value = t.material;
      r.getCell(3).value = t.marca;
      r.getCell(4).value = t.modelo || '';
      r.getCell(5).value = t.numero_serie || '';
      r.getCell(6).value = t.cor || '';
      r.getCell(7).value = t.setor || '';
      r.getCell(8).value = t.status;
      r.getCell(9).value = t.processo || '';

      r.getCell(1).alignment = { horizontal: 'center' };
      r.getCell(8).alignment = { horizontal: 'center' };

      const bgColor = idx % 2 === 0 ? 'FFFFFFFF' : 'FF' + CINZA_BG;
      for (let c = 1; c <= 9; c++) {
        const cell = r.getCell(c);
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgColor } };
        cell.font = { name: 'Calibri', size: 10.5, color: { argb: 'FF334155' } };
        cell.border = {
          bottom: { style: 'thin', color: { argb: 'FF' + CINZA_BORDA } }
        };
      }

      // Status colorido
      const cellStatus = r.getCell(8);
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
    for (let c = 1; c <= 9; c++) {
      const cell = ws.getRow(4).getCell(c);
      cell.border = {
        top: { style: 'medium', color: { argb: 'FF' + AZUL_MEDIO } },
        bottom: { style: 'medium', color: { argb: 'FF' + AZUL_ESCURO } },
        left: c === 1 ? { style: 'medium', color: { argb: 'FF' + AZUL_MEDIO } } : undefined,
        right: c === 9 ? { style: 'medium', color: { argb: 'FF' + AZUL_MEDIO } } : undefined
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
  function popularSelectsMateriais() {
    const materiais = getMateriais();
    const todosGlobal = getTombamentos();
    const tiUser = isTI();
    const setor = meuSetor();

    // Não-TI: contagens só do seu setor
    const todos = tiUser ? todosGlobal : todosGlobal.filter(t => t.setor === setor);

    // Contar por material
    const contagemMat = {};
    todos.forEach(t => { contagemMat[t.material] = (contagemMat[t.material] || 0) + 1; });

    // Contar por status
    const contagemStatus = { 'Ativo': 0, 'Em Baixa': 0 };
    todos.forEach(t => { if (contagemStatus[t.status] !== undefined) contagemStatus[t.status]++; });

    // Select do modal de cadastro
    const selectMaterial = document.getElementById('cad-material');
    const valAtual = selectMaterial.value;
    selectMaterial.innerHTML = '<option value="">Selecione o material...</option>' +
      materiais.map(p => `<option value="${escapeHtml(p)}">${escapeHtml(p)}</option>`).join('');
    if (valAtual) selectMaterial.value = valAtual;

    // Select do filtro na listagem — com contagem
    const filtroMaterial = document.getElementById('filtro-material');
    const valFiltro = filtroMaterial.value;
    const matsComContagem = materiais
      .filter(p => contagemMat[p])  // só mostrar materiais que existem
      .sort((a, b) => (contagemMat[b] || 0) - (contagemMat[a] || 0));
    // Também incluir materiais nos tombamentos que não estão no catálogo
    Object.keys(contagemMat).forEach(m => {
      if (!matsComContagem.includes(m)) matsComContagem.push(m);
    });
    filtroMaterial.innerHTML = `<option value="Todos">Todos os Materiais (${todos.length})</option>` +
      matsComContagem.map(p => `<option value="${escapeHtml(p)}">${escapeHtml(p)} (${contagemMat[p] || 0})</option>`).join('');
    filtroMaterial.value = valFiltro || 'Todos';

    // Select do filtro de status — com contagem
    const filtroStatus = document.getElementById('filtro-status');
    const valStatus = filtroStatus.value;
    filtroStatus.innerHTML = `<option value="Todos">Todos os Status (${todos.length})</option>
      <option value="Ativo">Ativo (${contagemStatus['Ativo']})</option>
      <option value="Em Baixa">Em Baixa (${contagemStatus['Em Baixa']})</option>
      <option value="Recentes">+ Recentes (30 dias)</option>`;
    filtroStatus.value = valStatus || 'Todos';

    // Atualizar selects de setores dinamicamente
    popularSelectsSetores();
  }

  function popularSelectsSetores() {
    const users = JSON.parse(localStorage.getItem(USERS_KEY) || '{}');
    const setores = Object.keys(users).sort((a, b) => a.localeCompare(b, 'pt-BR'));

    // Filtro setor (listagem)
    const filtroSetor = document.getElementById('filtro-setor');
    const valFiltroSetor = filtroSetor.value;
    filtroSetor.innerHTML = '<option value="Todos">Todos os Setores</option>' +
      setores.map(s => `<option value="${escapeHtml(s)}">${escapeHtml(s)}</option>`).join('');
    filtroSetor.value = valFiltroSetor || 'Todos';

    // Editar setor
    const editarSetor = document.getElementById('editar-setor');
    const valEditar = editarSetor.value;
    editarSetor.innerHTML = setores.map(s => `<option value="${escapeHtml(s)}">${escapeHtml(s)}</option>`).join('');
    if (valEditar) editarSetor.value = valEditar;

    // Transferir destino
    const transDestino = document.getElementById('transferir-destino');
    const valTrans = transDestino.value;
    transDestino.innerHTML = '<option value="">Selecione o destino...</option>' +
      setores.map(s => `<option value="${escapeHtml(s)}">${escapeHtml(s)}</option>`).join('');
    if (valTrans) transDestino.value = valTrans;
  }

  function popularDatalistMarcas() {
    // Combobox customizado
    const marcas = getMarcas();
    const input = document.getElementById('cad-marca');
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

    // Fechar ao clicar fora (global — cobre todos os comboboxes)
    document.addEventListener('click', (e) => {
      if (!e.target.closest('.combobox-wrapper')) {
        document.querySelectorAll('.combobox-dropdown.open').forEach(d => d.classList.remove('open'));
      }
    });
  }

  // ── CONFIGURAÇÕES (só TI) ─────────────────────────────────────────────
  function carregarConfiguracoes() {
    renderMarcas();
    renderMateriais();
    renderSetores();
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

  function renderMateriais() {
    const materiais = getMateriais();
    const container = document.getElementById('lista-materiais');
    if (materiais.length === 0) {
      container.innerHTML = '<p class="config-empty">Nenhum material cadastrado.</p>';
      return;
    }
    container.innerHTML = materiais.map(p => `
      <div class="config-tag">
        <i class="fas fa-desktop" style="font-size:0.75rem; opacity:0.5;"></i>
        ${escapeHtml(p)}
        <button class="tag-remove" data-material="${escapeHtml(p)}" title="Remover">
          <i class="fas fa-times"></i>
        </button>
      </div>
    `).join('');

    container.querySelectorAll('.tag-remove[data-material]').forEach(btn => {
      btn.addEventListener('click', () => {
        const nome = btn.dataset.material;
        let lista = getMateriais();
        lista = lista.filter(p => p !== nome);
        saveMateriais(lista);
        renderMateriais();
        popularSelectsMateriais();
        showToast(`Material "${nome}" removido.`);
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

  // Adicionar material
  document.getElementById('btn-add-material').addEventListener('click', () => {
    const input = document.getElementById('input-novo-material');
    const nome = input.value.trim();
    if (!nome) { showToast('Digite o nome do material.', true); return; }
    const lista = getMateriais();
    if (lista.some(p => p.toLowerCase() === nome.toLowerCase())) {
      showToast('Esse material já existe.', true); return;
    }
    lista.push(nome);
    saveMateriais(lista);
    input.value = '';
    renderMateriais();
    popularSelectsMateriais();
    showToast(`Material "${nome}" adicionado!`);
  });

  document.getElementById('input-novo-material').addEventListener('keydown', (e) => {
    if (e.key === 'Enter') { e.preventDefault(); document.getElementById('btn-add-material').click(); }
  });

  // ── GERENCIAMENTO DE SETORES E SENHAS ─────────────────────────────────
  function getUsers() {
    return JSON.parse(localStorage.getItem(USERS_KEY) || '{}');
  }
  function saveUsers(users) {
    localStorage.setItem(USERS_KEY, JSON.stringify(users));
    _syncBackground('/api/sync/setores', users);
  }

  function renderSetores() {
    const users = getUsers();
    const container = document.getElementById('lista-setores');
    if (!container) {
      console.error('[TOMBAMENTO] Elemento #lista-setores não encontrado no DOM.');
      return;
    }
    const setores = Object.keys(users).sort((a, b) => a.localeCompare(b, 'pt-BR'));

    if (setores.length === 0) {
      container.innerHTML = '<p class="config-empty">Nenhum setor cadastrado.</p>';
      return;
    }

    container.innerHTML = setores.map(setor => {
      const isTIsetor = setor === 'TI';
      return `
        <div class="config-setor-row ${isTIsetor ? 'setor-ti' : ''}">
          <div class="setor-info">
            <i class="fas fa-building" style="opacity:0.5;"></i>
            <span class="setor-nome">${escapeHtml(setor)}</span>
            ${isTIsetor ? '<span class="setor-badge-admin">ADMIN</span>' : ''}
          </div>
          <div class="setor-actions">
            <div class="setor-senha-group">
              <input type="password" class="setor-senha-input" value="${escapeHtml(users[setor])}" data-setor="${escapeHtml(setor)}" readonly>
              <button class="btn-icon setor-toggle-senha" data-setor="${escapeHtml(setor)}" title="Mostrar/ocultar senha">
                <i class="fas fa-eye"></i>
              </button>
            </div>
            <button class="btn-icon setor-editar-senha" data-setor="${escapeHtml(setor)}" title="Editar senha">
              <i class="fas fa-pen"></i>
            </button>
            ${!isTIsetor ? `<button class="btn-icon setor-remover" data-setor="${escapeHtml(setor)}" title="Remover setor">
              <i class="fas fa-trash" style="color:#ef4444;"></i>
            </button>` : ''}
          </div>
        </div>`;
    }).join('');

    // Toggle senha visível
    container.querySelectorAll('.setor-toggle-senha').forEach(btn => {
      btn.addEventListener('click', () => {
        const setor = btn.dataset.setor;
        const input = container.querySelector(`.setor-senha-input[data-setor="${setor}"]`);
        const icon = btn.querySelector('i');
        if (input.type === 'password') {
          input.type = 'text';
          icon.className = 'fas fa-eye-slash';
        } else {
          input.type = 'password';
          icon.className = 'fas fa-eye';
        }
      });
    });

    // Editar senha
    container.querySelectorAll('.setor-editar-senha').forEach(btn => {
      btn.addEventListener('click', () => {
        const setor = btn.dataset.setor;
        const input = container.querySelector(`.setor-senha-input[data-setor="${setor}"]`);
        if (input.readOnly) {
          input.readOnly = false;
          input.type = 'text';
          input.focus();
          input.select();
          btn.innerHTML = '<i class="fas fa-check" style="color:#10b981;"></i>';
          btn.title = 'Salvar senha';
        } else {
          const novaSenha = input.value.trim();
          if (!novaSenha) { showToast('A senha não pode ficar vazia.', true); return; }
          const users = getUsers();
          users[setor] = novaSenha;
          saveUsers(users);
          input.readOnly = true;
          input.type = 'password';
          btn.innerHTML = '<i class="fas fa-pen"></i>';
          btn.title = 'Editar senha';
          container.querySelector(`.setor-toggle-senha[data-setor="${setor}"] i`).className = 'fas fa-eye';
          showToast(`Senha do setor "${setor}" atualizada!`);
          addLog('Config', `Senha do setor ${setor} alterada`, 'alerta');
        }
      });
    });

    // Remover setor — usa modal de confirmação (confirm() bloqueado em file://)
    container.querySelectorAll('.setor-remover').forEach(btn => {
      btn.addEventListener('click', () => {
        const setor = btn.dataset.setor;
        excluirSetorNome = setor;
        modalMsg.textContent = `Remover o setor "${setor}"? Os tombamentos desse setor NÃO serão apagados.`;
        document.getElementById('modal-titulo').textContent = 'Remover Setor';
        modalConfirmar.className = 'btn-danger';
        modalOverlay.classList.add('show');
      });
    });
  }

  // Adicionar setor
  document.getElementById('btn-add-setor').addEventListener('click', () => {
    const inputNome = document.getElementById('input-novo-setor');
    const inputSenha = document.getElementById('input-nova-senha-setor');
    const nome = inputNome.value.trim();
    const senha = inputSenha.value.trim();
    if (!nome) { showToast('Digite o nome do setor.', true); return; }
    if (!senha) { showToast('Digite a senha do setor.', true); return; }
    const users = getUsers();
    if (nome in users) { showToast('Esse setor já existe.', true); return; }
    users[nome] = senha;
    saveUsers(users);
    inputNome.value = '';
    inputSenha.value = '';
    renderSetores();
    popularSelectsSetores();
    if (typeof popularLoginSetores === 'function') popularLoginSetores();
    showToast(`Setor "${nome}" adicionado!`);
    addLog('Config', `Novo setor "${nome}" criado`, 'sucesso');
  });

  document.getElementById('input-novo-setor').addEventListener('keydown', (e) => {
    if (e.key === 'Enter') { e.preventDefault(); document.getElementById('btn-add-setor').click(); }
  });
  document.getElementById('input-nova-senha-setor').addEventListener('keydown', (e) => {
    if (e.key === 'Enter') { e.preventDefault(); document.getElementById('btn-add-setor').click(); }
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

      const colTomb = findCol('tombamento', 'tomb', 'nº tomb', 'numero tomb', 'pat');
      const colMat = findCol('material', 'produto', 'tipo', 'equipamento');
      const colMarca = findCol('marca');
      const colModelo = findCol('modelo');
      const colSerie = findCol('serie', 'série', 'numero de serie', 'número de série');
      const colCor = findCol('cor');
      const colSetor = findCol('setor', 'departamento');
      const colStatus = findCol('status', 'situação', 'situacao');
      const colProcesso = findCol('processo');

      if (!colMat && !colTomb) {
        showToast('Não foi possível identificar as colunas. Verifique o cabeçalho.', true);
        return;
      }

      // Ler dados a partir da linha seguinte ao cabeçalho
      for (let r = headerRow + 1; r <= ws.rowCount; r++) {
        const row = ws.getRow(r);
        const tomb = colTomb ? String(row.getCell(colTomb).value || '').trim() : '';
        const mat = colMat ? String(row.getCell(colMat).value || '').trim() : '';
        const marca = colMarca ? String(row.getCell(colMarca).value || '').trim() : '';
        const modelo = colModelo ? String(row.getCell(colModelo).value || '').trim() : '';
        const serie = colSerie ? String(row.getCell(colSerie).value || '').trim() : '';
        const cor = colCor ? String(row.getCell(colCor).value || '').trim() : '';
        const setor = colSetor ? String(row.getCell(colSetor).value || '').trim() : 'TI';
        const status = colStatus ? String(row.getCell(colStatus).value || '').trim() : 'Ativo';
        const processo = colProcesso ? String(row.getCell(colProcesso).value || '').trim() : '';

        // Pular linhas vazias
        if (!mat && !marca && !serie && !tomb) continue;

        dadosImportados.push({
          pat: tomb,
          tombamento: tomb ? Number(tomb.replace(/[^0-9]/g, '')) : 0,
          material: mat,
          marca: marca,
          modelo: modelo,
          serie: serie,
          cor: cor,
          setor: setor || 'TI',
          status: (status.toLowerCase().includes('baixa') ? 'Em Baixa' : 'Ativo'),
          processo: processo
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
          <td>${escapeHtml(d.pat) || '(auto)'}</td>
          <td>${escapeHtml(d.material)}</td>
          <td>${escapeHtml(d.marca)}</td>
          <td>${escapeHtml(d.modelo)}</td>
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
        pat: d.pat || '',
        material: d.material,
        cor: d.cor || '',
        descricao: '',
        marca: d.marca,
        modelo: d.modelo || '',
        numero_serie: d.serie || '',
        status: d.status,
        setor: d.setor,
        processo: d.processo || '',
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
    popularSelectsMateriais();
    carregarDashboard();
  });

  // Baixar modelo de planilha
  btnModeloExcel.addEventListener('click', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Modelo');

    ws.columns = [
      { header: 'PAT', key: 'pat', width: 14 },
      { header: 'Material', key: 'mat', width: 20 },
      { header: 'Marca', key: 'marca', width: 22 },
      { header: 'Modelo', key: 'modelo', width: 28 },
      { header: 'Nº Série', key: 'serie', width: 24 },
      { header: 'Cor', key: 'cor', width: 14 },
      { header: 'Setor', key: 'setor', width: 20 },
      { header: 'Status', key: 'status', width: 16 },
      { header: 'Processo', key: 'processo', width: 16 }
    ];

    // Estilo do cabeçalho
    ws.getRow(1).eachCell(cell => {
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2563EB' } };
      cell.alignment = { horizontal: 'center' };
    });

    // Linha exemplo
    ws.addRow({ pat: '25/0001', mat: 'COMPUTADOR', marca: 'LENOVO', modelo: 'V530S', serie: 'SN123456', cor: 'PRETO', setor: 'TI', status: 'Ativo', processo: '' });
    ws.addRow({ pat: '25/0002', mat: 'MONITOR', marca: 'LG', modelo: '20N37AA', serie: 'DL789012', cor: 'PRETO', setor: 'Guias', status: 'Ativo', processo: '' });

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
    const onlinePanel = document.getElementById('logs-online-panel');

    // ── Sessão Ativa & Histórico de Acessos ──────────────────────────
    const sessao = getSession();
    let loginDuracao = '';
    if (sessao && sessao.loginTime) {
      const diff = Date.now() - new Date(sessao.loginTime).getTime();
      const mins = Math.floor(diff / 60000);
      const hrs = Math.floor(mins / 60);
      loginDuracao = hrs > 0 ? `${hrs}h ${mins % 60}min` : `${mins}min`;
    }

    // Últimos logins de cada setor
    const acessos = {};
    logs.forEach(l => {
      if (l.acao === 'Login' && !acessos[l.setor]) {
        acessos[l.setor] = l.data;
      }
    });

    const setoresOnline = Object.entries(acessos).map(([setor, data]) => {
      const diff = Date.now() - new Date(data).getTime();
      const mins = Math.floor(diff / 60000);
      const recente = mins < 10;
      const horaLogin = new Date(data).toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });
      const dataLogin = new Date(data).toLocaleDateString('pt-BR', { day: '2-digit', month: '2-digit' });
      return { setor, data, mins, recente, horaLogin, dataLogin };
    }).sort((a, b) => new Date(b.data) - new Date(a.data));

    if (onlinePanel) {
      onlinePanel.innerHTML = `
        <!-- Sessão Atual -->
        <div class="online-sessao-card">
          <div class="online-sessao-header">
            <i class="fas fa-circle" style="color:#10b981;font-size:.6rem;animation:pulse-dot 2s infinite;"></i>
            <strong>Sua Sessão</strong>
          </div>
          <div class="online-sessao-info">
            <span><i class="fas fa-building"></i> ${escapeHtml(sessao ? sessao.setor : '—')}</span>
            <span><i class="fas fa-clock"></i> Online há ${loginDuracao || '—'}</span>
          </div>
        </div>
        <!-- Últimos Acessos por Setor -->
        <div class="online-acessos-card">
          <div class="online-acessos-header">
            <i class="fas fa-users"></i> <strong>Últimos Acessos por Setor</strong>
          </div>
          <div class="online-acessos-list">
            ${setoresOnline.length === 0 ? '<div style="padding:12px;color:var(--cinza-400);font-size:.82rem;">Nenhum acesso registrado</div>' :
              setoresOnline.map(s => `
                <div class="online-acesso-item">
                  <div class="online-acesso-dot" style="background:${s.recente ? '#10b981' : '#94a3b8'};"></div>
                  <div class="online-acesso-nome">${escapeHtml(s.setor)}</div>
                  <div class="online-acesso-tempo">
                    ${s.recente ? '<span class="badge-online">ativo</span>' : ''}
                    ${s.dataLogin} às ${s.horaLogin}
                  </div>
                </div>
              `).join('')}
          </div>
        </div>
      `;
    }

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
    excluirSetorNome = null;
    excluirId = null;
    modalMsg.textContent = 'Deseja limpar todos os logs de atividade? Esta ação não pode ser desfeita.';
    document.getElementById('modal-titulo').textContent = 'Limpar Logs';
    modalConfirmar.className = 'btn-danger';
    // Usa callback temporário
    const onConfirm = () => {
      localStorage.removeItem(LOGS_KEY);
      if (window.MODO_API) {
        fetch('/api/logs', { method: 'DELETE' }).catch(() => {});
      }
      fecharModalOverlay();
      addLog('Logs limpos', 'Todos os logs anteriores foram apagados', 'alerta');
      carregarLogs();
      showToast('Logs limpos com sucesso!');
      modalConfirmar.removeEventListener('click', onConfirm);
    };
    modalConfirmar.addEventListener('click', onConfirm);
    modalOverlay.classList.add('show');
  });
});

// ── Toggle colapso das seções do dashboard ──────────────────────────────────
function toggleCardSection(header) {
  const el = header.nextElementSibling;
  const isCollapsed = el.classList.contains('collapsed');
  if (isCollapsed) {
    // Expandir: define altura real antes de remover classe
    el.style.height = el.scrollHeight + 'px';
    el.classList.remove('collapsed');
    header.classList.remove('collapsed');
    el.addEventListener('transitionend', () => { el.style.height = ''; }, { once: true });
  } else {
    // Colapsar: fixa altura atual e transiciona para 0
    el.style.height = el.scrollHeight + 'px';
    requestAnimationFrame(() => {
      el.classList.add('collapsed');
      header.classList.add('collapsed');
    });
  }
}
