from pathlib import Path

APP = Path('app.py')
text = APP.read_text(encoding='utf-8', errors='replace')

# 1) Ensure default show_lead state set
needle = 'st.session_state.setdefault("show_status", True)'
if needle in text and 'show_lead' not in text:
    text = text.replace(needle, needle + "\n    st.session_state.setdefault(\"show_lead\", True)")

# 2) Add second checkbox after the first sidebar checkbox
anchor = 'value=st.session_state.get("show_status", True),'
pos = text.find(anchor)
if pos != -1:
    insert_at = text.find('\n', pos)
    if insert_at == -1:
        insert_at = pos + len(anchor)
    block = (
        "\n        st.checkbox(\"Análise por Site (lead time)\","\
        "\n            key=\"show_lead\","\
        "\n            value=st.session_state.get(\"show_lead\", True),\n        )\n"
    )
    text = text[: insert_at + 1] + block + text[insert_at + 1 :]

APP.write_text(text, encoding='utf-8')
print('sidebar_modified')

