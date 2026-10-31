"""Interface local do ensalamento da ABC Aprendiz."""

from io import BytesIO
from pathlib import Path
import hashlib

from openpyxl import load_workbook
import pandas as pd
import streamlit as st

from app import inspect_workbook, make_result, preview_rows, room_choices_for_mode


st.set_page_config(page_title="ABC Aprendiz · Ensalamento", page_icon="🎓", layout="wide")
st.markdown("""
<style>
  .stApp {background: #f7fafc; color: #1b315c;}
  .block-container {max-width: 1180px; padding-top: 2rem; padding-bottom: 4rem;}
  .hero {background: linear-gradient(115deg,#243d87 0%,#2b469a 65%,#0f75bb 100%);
         border-radius: 22px; padding: 2rem 2.4rem; color: white; margin-bottom: 1.6rem;
         border-top: 5px solid #9ccd58; box-shadow: 0 12px 32px #2b469a25;}
  .hero h1 {color: white; margin: .15rem 0 .4rem; font-size: 2.45rem;}
  .hero p {color: #edf7ff; margin: 0; font-size: 1.05rem;}
  .eyebrow {letter-spacing: .14em; text-transform: uppercase; font-size: .76rem; font-weight: 700; color: #d4f3ed;}
  div[data-testid="stMetric"] {background: white; border: 1px solid #e4eaf1; border-radius: 15px;
    padding: 1rem 1.2rem; box-shadow: 0 3px 12px #2b469a0c;}
  .stApp [data-baseweb="tab"], .stApp [data-baseweb="tab"] * {color: #1b315c !important;}
  .stApp [data-baseweb="tab"][aria-selected="true"],
  .stApp [data-baseweb="tab"][aria-selected="true"] * {color: #2b469a !important;}
  .stApp [data-testid="stMetric"] * {color: #1b315c !important;}
  div[data-testid="stFileUploader"], div[data-testid="stDataFrame"] {border-radius: 14px;}
  .small-note {color:#5b6c80; font-size:.92rem;}
</style>
""", unsafe_allow_html=True)


def preview(sheet):
    columns, rows = preview_rows(sheet)
    return pd.DataFrame(rows, columns=columns).dropna(axis=1, how="all")


def output_table(content, sheet_name):
    book = load_workbook(BytesIO(content), read_only=True, data_only=True)
    rows = book[sheet_name].values
    columns = next(rows)
    table = pd.DataFrame(rows, columns=columns)
    book.close()
    return table


brand, symbol = st.columns([5, 1], vertical_alignment="center")
brand.image(str(Path(__file__).with_name("abc1024x279.webp")), width=430)
symbol.image(str(Path(__file__).with_name("abc-logo-1024x798.webp")), width=105)

st.markdown("""
<div class="hero">
  <div class="eyebrow">ABC Aprendiz · Desafio UniSoma 2026</div>
  <h1>Turmas equilibradas, sem retrabalho.</h1>
  <p>Confira os dados, escolha as salas e acompanhe a distribuição antes de baixar o Excel final.</p>
</div>
""", unsafe_allow_html=True)

work_tab, manual_tab = st.tabs(["📊 Ensalamento", "📘 Manual e ajuda"])

with manual_tab:
    manual = Path(__file__).with_name("MANUAL_USUARIO.md").read_text(encoding="utf-8")
    st.markdown(manual)
    st.download_button("Baixar manual (.pdf)",
                       Path(__file__).with_name("MANUAL_USUARIO.pdf").read_bytes(),
                       file_name="Manual_Ensalamento_ABC.pdf", mime="application/pdf")

with work_tab:
    st.subheader("1 · Escolha o arquivo")
    uploaded = st.file_uploader("Planilha Excel com as abas Alunos e Turmas", type="xlsx",
                                help="O arquivo original é preservado. O resultado será baixado como uma cópia.")
    if uploaded is None:
        st.info("Comece selecionando uma planilha .xlsx. O manual contém o formato esperado e exemplos de tratamento de erros.")
        st.stop()

    content = uploaded.getvalue()
    digest = hashlib.sha256(content).hexdigest()
    try:
        book = load_workbook(BytesIO(content), data_only=True)
    except Exception as exc:
        st.error(f"Não foi possível abrir a planilha. Verifique se é um .xlsx válido: {exc}")
        st.stop()

    st.caption(f"Arquivo selecionado: {uploaded.name} · {len(content) / 1024:.0f} KB")
    left, right = st.columns(2, gap="large")
    with left:
        st.markdown("**Prévia · Alunos**")
        if "Alunos" in book:
            st.dataframe(preview(book["Alunos"]), hide_index=True, width="stretch", height=320)
            count = book["Alunos"].max_row
            st.caption("Até 12 linhas exibidas" + (f" · {count - 1} linhas no arquivo" if count else ""))
        else:
            st.warning("A aba Alunos não foi encontrada.")
    with right:
        st.markdown("**Prévia · Turmas**")
        if "Turmas" in book:
            st.dataframe(preview(book["Turmas"]).iloc[:, :2], hide_index=True, width="stretch", height=320)
        else:
            st.warning("A aba Turmas não foi encontrada.")

    students, resources, _, _ = inspect_workbook(book)
    book.close()
    st.subheader("2 · Escolha como usar as salas")
    allow_empty = st.checkbox("Prefiro deixar salas vazias quando possível", value=False,
                              help="Usa o menor número de salas que comporta todos os alunos. Desmarcado, ocupa todas as salas disponíveis.")
    st.caption("As salas ocupadas recebem alunos com diferença de no máximo um entre elas no mesmo turno.")
    choices = room_choices_for_mode(students, resources, allow_empty)
    cols = st.columns(2, gap="large")
    for col, (period, suffix) in zip(cols, (("Manhã", "manha"), ("Tarde", "tarde"))):
        with col:
            available = resources.get(f"turmas_{suffix}")
            capacity = resources.get(f"vagas_{suffix}")
            count = sum(row[2] == period for row in students)
            if available:
                st.metric(f"{period} · salas ocupadas", choices.get(suffix, 0))
                st.caption(f"{count} alunos · {available} {'sala disponível' if available == 1 else 'salas disponíveis'} · {capacity or '—'} vagas por sala")
            else:
                st.warning(f"Quantidade de turmas da {period.lower()} ausente ou inválida.")

    check_book = load_workbook(BytesIO(content), data_only=True)
    _, _, issues, blocking = inspect_workbook(check_book, choices)
    check_book.close()
    if issues:
        with st.expander(f"{'⚠️' if blocking else 'ℹ️'} {len(issues)} aviso(s) e inconsistência(s)", expanded=blocking):
            for issue in issues:
                st.write("• " + issue)
    else:
        st.success("Dados consistentes e capacidade suficiente para a configuração escolhida.")

    signature = (digest, tuple(sorted(choices.items())))
    if st.session_state.get("result_signature") != signature:
        st.session_state.pop("result", None)
    if st.button("Gerar ensalamento" if not blocking else "Gerar relatório de inconsistências",
                 type="primary", width="stretch"):
        try:
            with st.spinner("Validando dados e resolvendo o modelo de otimização…"):
                result_book = load_workbook(BytesIO(content))
                result = make_result(result_book, choices)
                result_book.close()
            st.session_state.result = result
            st.session_state.result_signature = signature
        except Exception as exc:
            st.error(f"Não foi possível gerar o resultado: {exc}")

    if "result" not in st.session_state:
        st.stop()
    result_bytes, blocked, result_issues, stats = st.session_state.result
    st.subheader("3 · Resultado")
    name = Path(uploaded.name).stem
    if blocked:
        st.error("O ensalamento foi bloqueado. Corrija os dados indicados e envie a planilha novamente.")
        st.dataframe(output_table(result_bytes, "Inconsistências"), hide_index=True, width="stretch")
        st.download_button("Baixar planilha com inconsistências", result_bytes,
                           file_name=f"{name}_inconsistencias.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.stop()

    total = sum(row[1] for row in stats)
    metrics = st.columns(4)
    for col, label, number in zip(metrics,
                                  ("Alunos alocados", "Turmas ocupadas", "Meninas", "Meninos"),
                                  (total, len(stats), sum(row[2] for row in stats), sum(row[3] for row in stats))):
        col.metric(label, number)
    room_table = pd.DataFrame(stats, columns=["Turma", "Alunos", "Meninas", "Meninos"])
    room_table["% meninas"] = (100 * room_table["Meninas"] / room_table["Alunos"]).round(1)
    dash_tab, list_tab, warnings_tab = st.tabs(["Visão geral", "Lista de alunos", "Inconsistências"])
    with dash_tab:
        st.markdown("#### Distribuição por turma")
        st.bar_chart(room_table, x="Turma", y=["Meninas", "Meninos"], stack=True,
                     color=["#3dbd9d", "#2b469a"], height=380, sort=False)
        st.dataframe(room_table, hide_index=True, width="stretch")
    with list_tab:
        result_table = output_table(result_bytes, "Ensalamento")
        search = st.text_input("Buscar aluno ou turma", placeholder="Digite um nome ou letra da turma")
        if search:
            result_table = result_table[result_table.astype(str).apply(
                lambda column: column.str.contains(search, case=False, regex=False)).any(axis=1)]
        st.dataframe(result_table, hide_index=True, width="stretch", height=450)
    with warnings_tab:
        st.dataframe(output_table(result_bytes, "Inconsistências"), hide_index=True, width="stretch")

    st.markdown("#### Baixar resultados")
    c1, c2 = st.columns(2)
    c1.download_button("Baixar Excel completo (.xlsx)", result_bytes,
                       file_name=f"{name}_resultado.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       type="primary", width="stretch")
    c2.download_button("Baixar resumo por turma (.csv)",
                       room_table.to_csv(index=False).encode("utf-8-sig"),
                       file_name=f"{name}_resumo.csv", mime="text/csv", width="stretch")
