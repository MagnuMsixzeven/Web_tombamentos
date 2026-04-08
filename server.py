"""
ASTIR – Sistema de Tombamento
Backend Flask + SQLite

Uso:
  pip install flask
  python server.py
  Acesse: http://localhost:5050
"""

import os
import sqlite3
import json
from datetime import datetime
from flask import Flask, request, jsonify, send_from_directory, session

app = Flask(__name__, static_folder='public', static_url_path='')
app.secret_key = 'astir-tombamento-secret-2026-xK9!pL2'

DB_PATH = os.path.join(os.path.dirname(__file__), 'tombamento.db')

# ─── Banco de Dados ────────────────────────────────────────────────────────────

def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    return conn


def init_db():
    conn = get_db()
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS setores (
            nome  TEXT PRIMARY KEY,
            senha TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS tombamentos (
            id                 INTEGER PRIMARY KEY AUTOINCREMENT,
            numero_tombamento  INTEGER NOT NULL,
            pat                TEXT    DEFAULT '',
            material           TEXT    NOT NULL,
            cor                TEXT    DEFAULT '',
            descricao          TEXT    DEFAULT '',
            marca              TEXT    NOT NULL,
            modelo             TEXT    DEFAULT '',
            numero_serie       TEXT    DEFAULT '',
            status             TEXT    DEFAULT 'Ativo',
            setor              TEXT    NOT NULL,
            processo           TEXT    DEFAULT '',
            data_cadastro      TEXT    NOT NULL
        );

        CREATE TABLE IF NOT EXISTS historico (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            tombamento  INTEGER NOT NULL,
            tipo        TEXT    NOT NULL,
            de          TEXT    DEFAULT '',
            para        TEXT    NOT NULL,
            por         TEXT    NOT NULL,
            justificativa TEXT  DEFAULT '',
            data        TEXT    NOT NULL
        );

        CREATE TABLE IF NOT EXISTS logs (
            id       INTEGER PRIMARY KEY AUTOINCREMENT,
            data     TEXT NOT NULL,
            setor    TEXT NOT NULL,
            acao     TEXT NOT NULL,
            detalhes TEXT DEFAULT '',
            nivel    TEXT DEFAULT 'info'
        );

        CREATE TABLE IF NOT EXISTS marcas (
            nome TEXT PRIMARY KEY
        );

        CREATE TABLE IF NOT EXISTS materiais (
            nome TEXT PRIMARY KEY
        );
    """)

    # Seed inicial de setores
    if conn.execute("SELECT COUNT(*) FROM setores").fetchone()[0] == 0:
        setores = [
            ('TI', 'ti123'),
            ('Almoxarifado', 'almoxarifado123'),
            ('Ambulatorio', 'ambulatorio123'),
            ('ARRECADAÇÃO', 'arrecadacao123'),
            ('AUDITORIA', 'auditoria123'),
            ('Audiologia', 'audiologia123'),
            ('Cadastro', 'cadastro123'),
            ('CCIH', 'ccih123'),
            ('Compras', 'compras123'),
            ('Cozinha', 'cozinha123'),
            ('Descartado', 'descartado123'),
            ('DIREX', 'direx123'),
            ('DIRETOR', 'diretor123'),
            ('Enfermagem', 'enfermagem123'),
            ('Farmacia', 'farmacia123'),
            ('Faturamento', 'faturamento123'),
            ('Financeiro', 'financeiro123'),
            ('Fisioterapia', 'fisioterapia123'),
            ('Guias', 'guias123'),
            ('INTERNAÇÃO', 'internacao123'),
            ('Juridico', 'juridico123'),
            ('Odontologia', 'odontologia123'),
            ('Polos', 'polos123'),
            ('Psicologia', 'psicologia123'),
            ('Recepcao', 'recepcao123'),
            ('RH', 'rh123'),
            ('SECONF', 'seconf123'),
            ('Segurança do Trabalho', 'seguranca123'),
            ('Servico Social', 'servicosocial123'),
            ('SPA', 'spa123'),
            ('Não Catalogado', 'naocatalogado123'),
        ]
        conn.executemany("INSERT OR IGNORE INTO setores VALUES (?,?)", setores)

    # Seed inicial de marcas
    if conn.execute("SELECT COUNT(*) FROM marcas").fetchone()[0] == 0:
        marcas = [('Samsung',),('LG',),('Dell',),('HP',),('Lenovo',),
                  ('TP-Link',),('Intelbras',),('AOC',),('Multilaser',),('Positivo',)]
        conn.executemany("INSERT OR IGNORE INTO marcas VALUES (?)", marcas)

    # Seed inicial de materiais
    if conn.execute("SELECT COUNT(*) FROM materiais").fetchone()[0] == 0:
        materiais = [('COMPUTADOR',),('MONITOR',),('NOBREAK',),('IMPRESSORA',),
                     ('DVR',),('SWITCH',),('CÂMERA',),('RACK',),('ROTEADOR',),
                     ('AR CONDICIONADO',)]
        conn.executemany("INSERT OR IGNORE INTO materiais VALUES (?)", materiais)

    conn.commit()
    conn.close()


# ─── Helpers ───────────────────────────────────────────────────────────────────

def row_to_dict(row):
    return dict(row)

def rows_to_list(rows):
    return [dict(r) for r in rows]

def is_ti():
    return session.get('setor') == 'TI'

def require_ti():
    if not is_ti():
        return jsonify({'error': 'Apenas TI tem permissão.'}), 403
    return None

def require_login():
    if 'setor' not in session:
        return jsonify({'error': 'Não autenticado.'}), 401
    return None


# ─── Rota principal ────────────────────────────────────────────────────────────

@app.route('/')
def index():
    """Serve o index.html injetando window.MODO_API = true para usar a API."""
    html_path = os.path.join(os.path.dirname(__file__), 'public', 'index.html')
    with open(html_path, 'r', encoding='utf-8') as f:
        content = f.read()
    inject = (
        '<script>window.MODO_API = true; console.log("[ASTIR] Modo API ativo — Flask + SQLite");</script>\n'
        '  <script src="js/api.js"></script>\n'
    )
    content = content.replace('<script src="js/app.js"></script>', inject + '  <script src="js/app.js"></script>')
    from flask import Response
    return Response(content, mimetype='text/html; charset=utf-8')


# ─── Autenticação ──────────────────────────────────────────────────────────────

@app.route('/api/login', methods=['POST'])
def login():
    data = request.get_json()
    setor = (data.get('setor') or '').strip()
    senha = (data.get('senha') or '').strip()
    if not setor or not senha:
        return jsonify({'error': 'Setor e senha são obrigatórios.'}), 400

    conn = get_db()
    row = conn.execute("SELECT senha FROM setores WHERE nome = ?", (setor,)).fetchone()
    conn.close()

    if not row or row['senha'] != senha:
        return jsonify({'error': 'Setor ou senha incorretos.'}), 401

    session['setor'] = setor
    session['login_time'] = datetime.now().isoformat()
    return jsonify({'ok': True, 'setor': setor})


@app.route('/api/logout', methods=['POST'])
def logout():
    session.clear()
    return jsonify({'ok': True})


@app.route('/api/session')
def get_session():
    if 'setor' in session:
        return jsonify({'setor': session['setor'], 'loginTime': session.get('login_time')})
    return jsonify({'setor': None})


# ─── Setores ───────────────────────────────────────────────────────────────────

@app.route('/api/setores', methods=['GET'])
def listar_setores():
    err = require_login(); 
    if err: return err
    conn = get_db()
    rows = conn.execute("SELECT nome, senha FROM setores ORDER BY nome COLLATE NOCASE").fetchall()
    conn.close()
    return jsonify(rows_to_list(rows))


@app.route('/api/setores', methods=['POST'])
def criar_setor():
    err = require_ti()
    if err: return err
    data = request.get_json()
    nome  = (data.get('nome') or '').strip()
    senha = (data.get('senha') or '').strip()
    if not nome or not senha:
        return jsonify({'error': 'Nome e senha são obrigatórios.'}), 400

    conn = get_db()
    try:
        conn.execute("INSERT INTO setores VALUES (?,?)", (nome, senha))
        conn.commit()
    except sqlite3.IntegrityError:
        conn.close()
        return jsonify({'error': 'Setor já existe.'}), 409
    conn.close()
    return jsonify({'ok': True}), 201


@app.route('/api/setores/<nome>', methods=['PUT'])
def atualizar_setor(nome):
    err = require_ti()
    if err: return err
    if nome == 'TI':
        return jsonify({'error': 'Não é permitido modificar o setor TI via API.'}), 403
    data = request.get_json()
    senha = (data.get('senha') or '').strip()
    if not senha:
        return jsonify({'error': 'Senha não pode ser vazia.'}), 400
    conn = get_db()
    conn.execute("UPDATE setores SET senha=? WHERE nome=?", (senha, nome))
    conn.commit()
    conn.close()
    return jsonify({'ok': True})


@app.route('/api/setores/<nome>', methods=['DELETE'])
def remover_setor(nome):
    err = require_ti()
    if err: return err
    if nome == 'TI':
        return jsonify({'error': 'O setor TI não pode ser removido.'}), 403
    conn = get_db()
    conn.execute("DELETE FROM setores WHERE nome=?", (nome,))
    conn.commit()
    conn.close()
    return jsonify({'ok': True})


# ─── Tombamentos ───────────────────────────────────────────────────────────────

@app.route('/api/tombamentos', methods=['GET'])
def listar_tombamentos():
    err = require_login()
    if err: return err

    conn = get_db()
    setor_usuario = session['setor']
    if setor_usuario == 'TI':
        rows = conn.execute("SELECT * FROM tombamentos ORDER BY numero_tombamento").fetchall()
    else:
        rows = conn.execute(
            "SELECT * FROM tombamentos WHERE setor=? ORDER BY numero_tombamento",
            (setor_usuario,)
        ).fetchall()
    conn.close()
    return jsonify(rows_to_list(rows))


@app.route('/api/tombamentos', methods=['POST'])
def criar_tombamento():
    err = require_ti()
    if err: return err
    data = request.get_json()

    conn = get_db()
    # Próximo número de tombamento
    row = conn.execute("SELECT MAX(numero_tombamento) as m FROM tombamentos").fetchone()
    proximo = (row['m'] or 0) + 1

    conn.execute("""
        INSERT INTO tombamentos
            (numero_tombamento, pat, material, cor, descricao, marca, modelo,
             numero_serie, status, setor, processo, data_cadastro)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?)
    """, (
        proximo,
        data.get('pat', ''),
        data.get('material', ''),
        data.get('cor', ''),
        data.get('descricao', ''),
        data.get('marca', ''),
        data.get('modelo', ''),
        data.get('numero_serie', ''),
        data.get('status', 'Ativo'),
        data.get('setor', 'TI'),
        data.get('processo', ''),
        datetime.now().isoformat()
    ))
    tomb_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
    conn.commit()
    conn.close()
    return jsonify({'ok': True, 'id': tomb_id, 'numero_tombamento': proximo}), 201


@app.route('/api/tombamentos/<int:tomb_id>', methods=['PUT'])
def atualizar_tombamento(tomb_id):
    err = require_ti()
    if err: return err
    data = request.get_json()
    conn = get_db()
    conn.execute("""
        UPDATE tombamentos SET
            material=?, cor=?, descricao=?, marca=?, modelo=?,
            numero_serie=?, status=?, setor=?, processo=?
        WHERE id=?
    """, (
        data.get('material', ''),
        data.get('cor', ''),
        data.get('descricao', ''),
        data.get('marca', ''),
        data.get('modelo', ''),
        data.get('numero_serie', ''),
        data.get('status', 'Ativo'),
        data.get('setor', ''),
        data.get('processo', ''),
        tomb_id
    ))
    conn.commit()
    conn.close()
    return jsonify({'ok': True})


@app.route('/api/tombamentos/<int:tomb_id>/status', methods=['PATCH'])
def alterar_status(tomb_id):
    err = require_login()
    if err: return err
    data = request.get_json()
    novo_status = data.get('status', '')
    if novo_status not in ('Ativo', 'Em Baixa'):
        return jsonify({'error': 'Status inválido.'}), 400

    conn = get_db()
    row = conn.execute("SELECT setor FROM tombamentos WHERE id=?", (tomb_id,)).fetchone()
    if not row:
        conn.close()
        return jsonify({'error': 'Tombamento não encontrado.'}), 404

    setor_usuario = session['setor']
    if setor_usuario != 'TI' and row['setor'] != setor_usuario:
        conn.close()
        return jsonify({'error': 'Sem permissão.'}), 403

    conn.execute("UPDATE tombamentos SET status=? WHERE id=?", (novo_status, tomb_id))
    conn.commit()
    conn.close()
    return jsonify({'ok': True})


@app.route('/api/tombamentos/<int:tomb_id>', methods=['DELETE'])
def excluir_tombamento(tomb_id):
    err = require_ti()
    if err: return err
    conn = get_db()
    conn.execute("DELETE FROM tombamentos WHERE id=?", (tomb_id,))
    conn.commit()
    conn.close()
    return jsonify({'ok': True})


@app.route('/api/tombamentos/importar', methods=['POST'])
def importar_tombamentos():
    err = require_ti()
    if err: return err
    lista = request.get_json()
    if not isinstance(lista, list):
        return jsonify({'error': 'Dados inválidos.'}), 400

    conn = get_db()
    row = conn.execute("SELECT MAX(numero_tombamento) as m FROM tombamentos").fetchone()
    proximo = (row['m'] or 0) + 1

    series_existentes = {
        r['numero_serie'].lower()
        for r in conn.execute("SELECT numero_serie FROM tombamentos WHERE numero_serie != ''").fetchall()
    }

    importados = 0
    duplicados = 0
    for d in lista:
        serie = (d.get('numero_serie') or '').strip().lower()
        if serie and serie in series_existentes:
            duplicados += 1
            continue

        num_tomb = d.get('numero_tombamento') or proximo
        if num_tomb >= proximo:
            proximo = num_tomb + 1

        conn.execute("""
            INSERT INTO tombamentos
                (numero_tombamento, pat, material, cor, descricao, marca, modelo,
                 numero_serie, status, setor, processo, data_cadastro)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?)
        """, (
            num_tomb,
            d.get('pat', ''),
            d.get('material', ''),
            d.get('cor', ''),
            d.get('descricao', ''),
            d.get('marca', ''),
            d.get('modelo', ''),
            d.get('numero_serie', ''),
            d.get('status', 'Ativo'),
            d.get('setor', 'TI'),
            d.get('processo', ''),
            datetime.now().isoformat()
        ))
        if serie:
            series_existentes.add(serie)
        importados += 1

    conn.commit()
    conn.close()
    return jsonify({'ok': True, 'importados': importados, 'duplicados': duplicados})


# ─── Transferência ─────────────────────────────────────────────────────────────

@app.route('/api/tombamentos/<int:tomb_id>/transferir', methods=['POST'])
def transferir(tomb_id):
    err = require_login()
    if err: return err
    data = request.get_json()
    destino = (data.get('destino') or '').strip()
    justificativa = (data.get('justificativa') or '').strip()
    setor_usuario = session['setor']

    if not destino:
        return jsonify({'error': 'Destino obrigatório.'}), 400
    if setor_usuario != 'TI' and not justificativa:
        return jsonify({'error': 'Justificativa obrigatória.'}), 400

    conn = get_db()
    row = conn.execute("SELECT * FROM tombamentos WHERE id=?", (tomb_id,)).fetchone()
    if not row:
        conn.close()
        return jsonify({'error': 'Tombamento não encontrado.'}), 404

    if setor_usuario != 'TI' and row['setor'] != setor_usuario:
        conn.close()
        return jsonify({'error': 'Sem permissão.'}), 403

    if row['status'] != 'Ativo':
        conn.close()
        return jsonify({'error': 'Não é possível transferir equipamento em baixa.'}), 400

    origem = row['setor']
    conn.execute("UPDATE tombamentos SET setor=? WHERE id=?", (destino, tomb_id))
    conn.execute("""
        INSERT INTO historico (tombamento, tipo, de, para, por, justificativa, data)
        VALUES (?,?,?,?,?,?,?)
    """, (row['numero_tombamento'], 'Transferência', origem, destino,
          setor_usuario, justificativa, datetime.now().isoformat()))
    conn.commit()
    conn.close()
    return jsonify({'ok': True, 'de': origem, 'para': destino})


# ─── Histórico ─────────────────────────────────────────────────────────────────

@app.route('/api/historico')
def listar_historico():
    err = require_login()
    if err: return err
    num = request.args.get('tombamento', type=int)
    conn = get_db()
    if num:
        rows = conn.execute(
            "SELECT * FROM historico WHERE tombamento=? ORDER BY id DESC", (num,)
        ).fetchall()
    else:
        rows = conn.execute("SELECT * FROM historico ORDER BY id DESC").fetchall()
    conn.close()
    return jsonify(rows_to_list(rows))


@app.route('/api/historico', methods=['POST'])
def add_historico():
    err = require_login()
    if err: return err
    data = request.get_json()
    conn = get_db()
    conn.execute("""
        INSERT INTO historico (tombamento, tipo, de, para, por, justificativa, data)
        VALUES (?,?,?,?,?,?,?)
    """, (
        data.get('tombamento', 0),
        data.get('tipo', ''),
        data.get('de', ''),
        data.get('para', ''),
        data.get('por', ''),
        data.get('justificativa', ''),
        data.get('data', datetime.now().isoformat())
    ))
    conn.commit()
    conn.close()
    return jsonify({'ok': True}), 201


# ─── Logs ──────────────────────────────────────────────────────────────────────

@app.route('/api/logs')
def listar_logs():
    err = require_ti()
    if err: return err
    conn = get_db()
    rows = conn.execute("SELECT * FROM logs ORDER BY id DESC LIMIT 500").fetchall()
    conn.close()
    return jsonify(rows_to_list(rows))


@app.route('/api/logs', methods=['POST'])
def add_log():
    if 'setor' not in session:
        return jsonify({'ok': True})  # silently ignore se não logado
    data = request.get_json()
    conn = get_db()
    conn.execute("""
        INSERT INTO logs (data, setor, acao, detalhes, nivel)
        VALUES (?,?,?,?,?)
    """, (
        datetime.now().isoformat(),
        session.get('setor', 'Sistema'),
        data.get('acao', ''),
        data.get('detalhes', ''),
        data.get('nivel', 'info')
    ))
    conn.commit()
    conn.close()
    return jsonify({'ok': True}), 201


@app.route('/api/logs', methods=['DELETE'])
def limpar_logs():
    err = require_ti()
    if err: return err
    conn = get_db()
    conn.execute("DELETE FROM logs")
    conn.commit()
    conn.close()
    return jsonify({'ok': True})


# ─── Marcas ────────────────────────────────────────────────────────────────────

@app.route('/api/marcas')
def listar_marcas():
    err = require_login()
    if err: return err
    conn = get_db()
    rows = conn.execute("SELECT nome FROM marcas ORDER BY nome COLLATE NOCASE").fetchall()
    conn.close()
    return jsonify([r['nome'] for r in rows])


@app.route('/api/marcas', methods=['POST'])
def criar_marca():
    err = require_ti()
    if err: return err
    data = request.get_json()
    nome = (data.get('nome') or '').strip()
    if not nome:
        return jsonify({'error': 'Nome obrigatório.'}), 400
    conn = get_db()
    try:
        conn.execute("INSERT INTO marcas VALUES (?)", (nome,))
        conn.commit()
    except sqlite3.IntegrityError:
        conn.close()
        return jsonify({'error': 'Marca já existe.'}), 409
    conn.close()
    return jsonify({'ok': True}), 201


@app.route('/api/marcas/<nome>', methods=['DELETE'])
def remover_marca(nome):
    err = require_ti()
    if err: return err
    conn = get_db()
    conn.execute("DELETE FROM marcas WHERE nome=?", (nome,))
    conn.commit()
    conn.close()
    return jsonify({'ok': True})


# ─── Materiais ─────────────────────────────────────────────────────────────────

@app.route('/api/materiais')
def listar_materiais():
    err = require_login()
    if err: return err
    conn = get_db()
    rows = conn.execute("SELECT nome FROM materiais ORDER BY nome COLLATE NOCASE").fetchall()
    conn.close()
    return jsonify([r['nome'] for r in rows])


@app.route('/api/materiais', methods=['POST'])
def criar_material():
    err = require_ti()
    if err: return err
    data = request.get_json()
    nome = (data.get('nome') or '').strip()
    if not nome:
        return jsonify({'error': 'Nome obrigatório.'}), 400
    conn = get_db()
    try:
        conn.execute("INSERT INTO materiais VALUES (?)", (nome,))
        conn.commit()
    except sqlite3.IntegrityError:
        conn.close()
        return jsonify({'error': 'Material já existe.'}), 409
    conn.close()
    return jsonify({'ok': True}), 201


@app.route('/api/materiais/<nome>', methods=['DELETE'])
def remover_material(nome):
    err = require_ti()
    if err: return err
    conn = get_db()
    conn.execute("DELETE FROM materiais WHERE nome=?", (nome,))
    conn.commit()
    conn.close()
    return jsonify({'ok': True})


# ─── Migração: importar dump completo do localStorage ─────────────────────────

@app.route('/api/migrar', methods=['POST'])
def migrar():
    """
    Recebe dump completo do localStorage e popula o banco SQLite.
    Requer chave de migração para segurança.
    """
    data = request.get_json()
    if data.get('chave') != 'ASTIR_MIGRAR_2026':
        return jsonify({'error': 'Chave de migração inválida.'}), 403

    conn = get_db()

    # ── Setores
    users = data.get('users', {})
    importados_setores = 0
    for nome, senha in users.items():
        conn.execute("INSERT OR REPLACE INTO setores VALUES (?,?)", (nome, str(senha)))
        importados_setores += 1

    # ── Marcas
    marcas = data.get('marcas', [])
    conn.execute("DELETE FROM marcas")
    for m in marcas:
        if m:
            conn.execute("INSERT OR IGNORE INTO marcas VALUES (?)", (str(m),))

    # ── Materiais
    materiais = data.get('materiais', [])
    conn.execute("DELETE FROM materiais")
    for m in materiais:
        if m:
            conn.execute("INSERT OR IGNORE INTO materiais VALUES (?)", (str(m),))

    # ── Tombamentos
    tombamentos = data.get('tombamentos', [])
    importados_tomb = 0
    for t in tombamentos:
        try:
            conn.execute("""
                INSERT OR IGNORE INTO tombamentos
                    (id, numero_tombamento, pat, material, cor, descricao, marca, modelo,
                     numero_serie, status, setor, processo, data_cadastro)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)
            """, (
                t.get('id'),
                t.get('numero_tombamento', 0),
                t.get('pat', ''),
                t.get('material', ''),
                t.get('cor', ''),
                t.get('descricao', ''),
                t.get('marca', ''),
                t.get('modelo', ''),
                t.get('numero_serie', ''),
                t.get('status', 'Ativo'),
                t.get('setor', ''),
                t.get('processo', ''),
                t.get('data_cadastro', datetime.now().isoformat())
            ))
            importados_tomb += 1
        except Exception:
            pass

    # ── Histórico
    historico = data.get('historico', [])
    importados_hist = 0
    for h in historico:
        try:
            conn.execute("""
                INSERT OR IGNORE INTO historico
                    (tombamento, tipo, de, para, por, justificativa, data)
                VALUES (?,?,?,?,?,?,?)
            """, (
                h.get('tombamento', 0),
                h.get('tipo', ''),
                h.get('de', ''),
                h.get('para', ''),
                h.get('por', ''),
                h.get('justificativa', ''),
                h.get('data', datetime.now().isoformat())
            ))
            importados_hist += 1
        except Exception:
            pass

    # ── Logs
    logs = data.get('logs', [])
    for l in logs:
        try:
            conn.execute("""
                INSERT OR IGNORE INTO logs (data, setor, acao, detalhes, nivel)
                VALUES (?,?,?,?,?)
            """, (
                l.get('data', datetime.now().isoformat()),
                l.get('setor', 'Sistema'),
                l.get('acao', ''),
                l.get('detalhes', ''),
                l.get('nivel', 'info')
            ))
        except Exception:
            pass

    conn.commit()
    conn.close()

    return jsonify({
        'ok': True,
        'setores': importados_setores,
        'tombamentos': importados_tomb,
        'historico': importados_hist,
        'marcas': len(marcas),
        'materiais': len(materiais)
    })


# ─── Endpoints de Sincronização em Background ──────────────────────────────────
# Chamados pelo app.js a cada save (fire-and-forget) para manter SQLite atualizado

@app.route('/api/sync/tombamentos', methods=['POST'])
def sync_tombamentos():
    if 'setor' not in session:
        return jsonify({'ok': True})
    lista = request.get_json()
    if not isinstance(lista, list):
        return jsonify({'error': 'Dados inválidos'}), 400
    conn = get_db()
    conn.execute("DELETE FROM tombamentos")
    for t in lista:
        conn.execute("""
            INSERT OR REPLACE INTO tombamentos
                (id, numero_tombamento, pat, material, cor, descricao, marca, modelo,
                 numero_serie, status, setor, processo, data_cadastro)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)
        """, (
            t.get('id'),
            t.get('numero_tombamento', 0),
            t.get('pat', ''),
            t.get('material', ''),
            t.get('cor', ''),
            t.get('descricao', ''),
            t.get('marca', ''),
            t.get('modelo', ''),
            t.get('numero_serie', ''),
            t.get('status', 'Ativo'),
            t.get('setor', ''),
            t.get('processo', ''),
            t.get('data_cadastro', datetime.now().isoformat())
        ))
    conn.commit()
    conn.close()
    return jsonify({'ok': True})


@app.route('/api/sync/marcas', methods=['POST'])
def sync_marcas():
    if 'setor' not in session:
        return jsonify({'ok': True})
    lista = request.get_json()
    if not isinstance(lista, list):
        return jsonify({'error': 'Dados inválidos'}), 400
    conn = get_db()
    conn.execute("DELETE FROM marcas")
    for nome in lista:
        if nome:
            conn.execute("INSERT OR IGNORE INTO marcas VALUES (?)", (str(nome),))
    conn.commit()
    conn.close()
    return jsonify({'ok': True})


@app.route('/api/sync/materiais', methods=['POST'])
def sync_materiais():
    if 'setor' not in session:
        return jsonify({'ok': True})
    lista = request.get_json()
    if not isinstance(lista, list):
        return jsonify({'error': 'Dados inválidos'}), 400
    conn = get_db()
    conn.execute("DELETE FROM materiais")
    for nome in lista:
        if nome:
            conn.execute("INSERT OR IGNORE INTO materiais VALUES (?)", (str(nome),))
    conn.commit()
    conn.close()
    return jsonify({'ok': True})


@app.route('/api/sync/setores', methods=['POST'])
def sync_setores():
    if session.get('setor') != 'TI':
        return jsonify({'ok': True})  # Apenas TI pode sincronizar setores
    users = request.get_json()
    if not isinstance(users, dict):
        return jsonify({'error': 'Dados inválidos'}), 400
    conn = get_db()
    for nome, senha in users.items():
        conn.execute("INSERT OR REPLACE INTO setores VALUES (?,?)", (nome, str(senha)))
    conn.commit()
    conn.close()
    return jsonify({'ok': True})


# ─── Inicializar e rodar ───────────────────────────────────────────────────────

if __name__ == '__main__':
    init_db()
    print("=" * 55)
    print("  ASTIR – Sistema de Tombamento (Flask + SQLite)")
    print("  Acesse: http://localhost:5050")
    print("  Banco:  tombamento.db")
    print("=" * 55)
    app.run(debug=True, port=5050)
