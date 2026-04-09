/**
 * ASTIR – Tombamento
 * api.js — Camada de acesso ao backend Flask/SQLite
 *
 * Inclua ANTES de app.js no index.html quando usar o server.py:
 *   <script src="js/api.js"></script>
 *   <script src="js/app.js"></script>
 */
(function () {
  'use strict';

  // ── helpers ──────────────────────────────────────────────────────────────

  async function req(method, url, body) {
    const opts = { method, headers: { 'Content-Type': 'application/json' } };
    if (body !== undefined) opts.body = JSON.stringify(body);
    const res = await fetch(url, opts);
    const json = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error(json.error || 'Erro na requisição');
    return json;
  }

  const get  = (url)       => req('GET',    url);
  const post = (url, body) => req('POST',   url, body);
  const put  = (url, body) => req('PUT',    url, body);
  const del  = (url)       => req('DELETE', url);
  const patch = (url, body) => req('PATCH', url, body);

  // ── Expõe a API globalmente ───────────────────────────────────────────────
  window.API = {

    // ── Sessão ───────────────────────────────────────────────────────────
    async getSession()       { return get('/api/session'); },
    async login(setor, senha){ return post('/api/login', { setor, senha }); },
    async logout()           { return post('/api/logout'); },

    // ── Setores ──────────────────────────────────────────────────────────
    async getSetores()              { return get('/api/setores'); },
    async addSetor(nome, senha)     { return post('/api/setores', { nome, senha }); },
    async updateSetor(nome, senha)  { return put(`/api/setores/${encodeURIComponent(nome)}`, { senha }); },
    async deleteSetor(nome)         { return del(`/api/setores/${encodeURIComponent(nome)}`); },

    // ── Tombamentos ──────────────────────────────────────────────────────
    async getTombamentos()    { return get('/api/tombamentos'); },
    async addTombamento(obj)  { return post('/api/tombamentos', obj); },
    async updateTombamento(id, obj) { return put(`/api/tombamentos/${id}`, obj); },
    async deleteTombamento(id)      { return del(`/api/tombamentos/${id}`); },
    async setStatus(id, status)     { return patch(`/api/tombamentos/${id}/status`, { status }); },
    async transferir(id, destino, justificativa) {
      return post(`/api/tombamentos/${id}/transferir`, { destino, justificativa });
    },
    async importar(lista) { return post('/api/tombamentos/importar', lista); },

    // ── Histórico ────────────────────────────────────────────────────────
    async getHistorico(tombamento) {
      const q = tombamento ? `?tombamento=${tombamento}` : '';
      return get('/api/historico' + q);
    },
    async addHistorico(obj) { return post('/api/historico', obj); },

    // ── Logs ─────────────────────────────────────────────────────────────
    async getLogs()        { return get('/api/logs'); },
    async addLog(acao, detalhes, nivel) {
      return post('/api/logs', { acao, detalhes: detalhes || '', nivel: nivel || 'info' });
    },
    async clearLogs()      { return del('/api/logs'); },

    // ── Marcas ───────────────────────────────────────────────────────────
    async getMarcas()       { return get('/api/marcas'); },
    async addMarca(nome)    { return post('/api/marcas', { nome }); },
    async deleteMarca(nome) { return del(`/api/marcas/${encodeURIComponent(nome)}`); },

    // ── Materiais ────────────────────────────────────────────────────────
    async getMateriais()        { return get('/api/materiais'); },
    async addMaterial(nome)     { return post('/api/materiais', { nome }); },
    async deleteMaterial(nome)  { return del(`/api/materiais/${encodeURIComponent(nome)}`); },
  };

  // Indica ao app.js que o modo API está ativo
  window.MODO_API = true;
  console.log('[ASTIR] Modo API ativo — usando Flask + SQLite');
})();
