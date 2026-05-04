import sqlite3
conn = sqlite3.connect('tombamento.db')
conn.row_factory = sqlite3.Row
print('=== TI_USUARIOS ===')
for r in conn.execute('SELECT id, login, senha_plain FROM ti_usuarios').fetchall():
    print('  id=%s login=%s senha=%s' % (r['id'], r['login'], r['senha_plain']))
print()
print('=== SETORES ===')
for r in conn.execute('SELECT nome, login, cargo, senha_plain FROM setores ORDER BY nome').fetchall():
    print('  nome=%-30s login=%-30s cargo=%-8s senha=%s' % (r['nome'], r['login'], r['cargo'], r['senha_plain']))
