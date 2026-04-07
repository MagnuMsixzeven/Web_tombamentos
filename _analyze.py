import json
from collections import Counter

with open(r'public\dados_importacao.json', 'r', encoding='utf-8') as f:
    dados = json.load(f)

tombs = dados['tombamentos']
print(f'Total registros: {len(tombs)}')

nums = [t['numero_tombamento'] for t in tombs]
print(f'Menor tombamento: {min(nums)}')
print(f'Maior tombamento: {max(nums)}')

faixa_0 = [n for n in nums if n == 0]
faixa_1_500 = [n for n in nums if 1 <= n <= 500]
faixa_23k = [n for n in nums if 23000 <= n < 24000]
faixa_24k = [n for n in nums if 24000 <= n < 25000]
faixa_25k = [n for n in nums if 25000 <= n < 26000]
print(f'\nTombamento = 0 (sem numero): {len(faixa_0)}')
print(f'Tombamento 1-500: {len(faixa_1_500)}')
print(f'Tombamento 23000-23999: {len(faixa_23k)}')
print(f'Tombamento 24000-24999: {len(faixa_24k)}')
print(f'Tombamento 25000-25999: {len(faixa_25k)}')

print('\n=== Primeiros 15 registros (NOVO PATRIMONIO) ===')
for t in tombs[:15]:
    nome = t.get('nome','')
    print(f'  nome={nome} -> tomb={t["numero_tombamento"]} | {t["produto"]} | {t["setor"]}')

print('\n=== Ultimos 15 registros (ANTIGO PATRIMONIO) ===')
for t in tombs[-15:]:
    nome = t.get('nome','')
    print(f'  nome={nome} -> tomb={t["numero_tombamento"]} | {t["produto"]} | {t["setor"]}')

counter = Counter(nums)
dups = {k: v for k, v in counter.items() if v > 1}
print(f'\nNumeros duplicados: {len(dups)}')
if dups:
    for k, v in sorted(dups.items())[:20]:
        print(f'  Tomb #{k}: aparece {v}x')
