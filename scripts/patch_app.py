from pathlib import Path

p = Path('app.py')
txt = p.read_text(encoding='utf-8', errors='ignore')

start_key = '# Para comparar com a coluna do arquivo (quando existir):'
end_key = 'st.dataframe(table_df[cols_order]'

si = txt.find(start_key)
ei = txt.find(end_key)
if si == -1 or ei == -1:
    raise SystemExit('anchors not found; abort')

before = txt[:si]
after = txt[ei:]

block = (
    "# Para comparar com a coluna do arquivo (quando existir):\n"
    "if 'current_status' not in table_df.columns:\n"
    "    table_df['current_status'] = table_df['current_full']\n"
    "# Rótulos adicionais (análise usa a última fase concluída)\n"
    "table_df['fase_label'] = table_df.get('last_phase_full', table_df['current_full'])\n"
    "table_df['fase_curta'] = table_df['current_short']\n\n"
    "cols_order = [c for c in [\n"
    "    'SITE',\n"
    "    'UF', 'Regional',\n"
    "    'current_status',\n"
    "    'fase_label',\n"
    "    'fase_curta',\n"
    "    'last_date',\n"
    "    'delay_days',\n"
    "    'year',\n"
    "    'Subcon', 'Type', 'Model', 'PO', 'Qty'\n"
    "] if c in table_df.columns]\n\n"
)

new = before + block + after
p.write_text(new, encoding='utf-8')
print('patched app.py')
