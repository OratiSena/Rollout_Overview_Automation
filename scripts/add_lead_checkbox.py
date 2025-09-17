from pathlib import Path
p = Path('app.py')
s = p.read_text(encoding='utf-8', errors='replace')
if 'key="show_lead"' in s:
    print('already_present')
    raise SystemExit(0)
anchor = 'with st.expander("Rollout", expanded=True):'
pos = s.find(anchor)
if pos == -1:
    print('anchor_missing')
    raise SystemExit(1)
# find the checkbox for show_status after anchor
pos2 = s.find('key="show_status"', pos)
if pos2 == -1:
    print('show_status_checkbox_missing')
    raise SystemExit(1)
# find end of the closing parenthesis of this checkbox call
close_idx = s.find('\n', pos2)
# advance until we pass a line that contains just ')' indented
end_call = s.find('\n', close_idx)
# find the first line that equals with `)` when stripped
scan = close_idx
end_call = None
for i in range(5):
    nxt = s.find('\n', scan+1)
    if nxt == -1:
        break
    line = s[scan+1:nxt]
    if line.strip() == ')':
        end_call = nxt
        break
    scan = nxt
if end_call is None:
    # fallback: insert right after current line
    end_call = close_idx
insert_block = "\n        st.checkbox(\"Analise por Site (lead time)\", key=\"show_lead\", value=st.session_state.get(\"show_lead\", True))\n"
s2 = s[:end_call+1] + insert_block + s[end_call+1:]
p.write_text(s2, encoding='utf-8')
print('inserted')
