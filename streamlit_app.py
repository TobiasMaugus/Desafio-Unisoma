"""Interface local do ensalamento da ABC Aprendiz."""

from io import BytesIO
from pathlib import Path
import hashlib

from openpyxl import load_workbook
import pandas as pd
import streamlit as st

from app import inspect_workbook, make_result, preview_rows, room_choices_for_mode


st.set_page_config(page_title="ABC Aprendiz · Ensalamento",
                   page_icon=str(Path(__file__).with_name("abc-logo-1024x798.webp")),
                   layout="wide")
st.html(Path(__file__).parent / "assets" / "interface.css")


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


with st.container(horizontal=True, horizontal_alignment="center"):
    with st.container(key="abc_page", width=1120, gap="medium"):
        with st.container(key="abc_header", horizontal=True, vertical_alignment="center", gap="large"):
            with st.container(key="abc_brand", width=240):
                st.image(str(Path(__file__).with_name("abc1024x279.webp")), width="stretch",
                         alt="ABC Aprendiz — Centro Social de Educação para o Trabalho")
            with st.container(width=700, gap="small"):
                st.title("Ensalamento de alunos")
                st.write("Organize os alunos nas turmas da manhã e da tarde e obtenha a planilha de resultado.")

        work_tab, manual_tab = st.tabs(["Ensalamento", "Manual e ajuda"])

        with manual_tab:
            with st.container(key="abc_manual", width=900):
                manual = Path(__file__).with_name("MANUAL_USUARIO.md").read_text(encoding="utf-8")
                st.markdown(manual)
                st.download_button("Baixar manual (.pdf)",
                                   Path(__file__).with_name("MANUAL_USUARIO.pdf").read_bytes(),
                                   file_name="Manual_Ensalamento_ABC.pdf", mime="application/pdf")

        with work_tab:
            with st.container(key="abc_upload", border=True, gap="small"):
                st.header("1. Selecionar planilha")
                st.write("Envie o arquivo com a lista de alunos e a quantidade de salas e vagas disponíveis.")
                with st.container(horizontal=True, gap="large", vertical_alignment="center"):
                    with st.container(width=660):
                        uploaded = st.file_uploader("Planilha Excel com as abas Alunos e Turmas", type="xlsx",
                                                    help="O arquivo original é preservado. O resultado será baixado como uma cópia.")
                    with st.container(key="abc_format", width=300, gap="small"):
                        st.subheader("O que a planilha precisa ter")
                        st.markdown("**Alunos** · nome, sexo e disponibilidade.  \n"
                                    "**Turmas** · quantidade de salas e vagas por turno.")
                        st.caption("Formato .xlsx · o arquivo original é preservado.")
                if uploaded is None:
                    st.caption("Selecione a planilha para conferir os dados. Consulte o formato completo em Manual e ajuda.")

            if uploaded is None:
                with st.container(key="abc_next", horizontal=True, gap="large"):
                    with st.container(width=450):
                        st.markdown("**2. Organizar as turmas**")
                        st.caption("Depois do envio, confira as salas disponíveis e gere o ensalamento.")
                    with st.container(width=450):
                        st.markdown("**3. Consultar e baixar**")
                        st.caption("Veja a distribuição dos alunos e baixe o Excel com o resultado.")
                st.stop()

            content = uploaded.getvalue()
            digest = hashlib.sha256(content).hexdigest()
            try:
                book = load_workbook(BytesIO(content), data_only=True)
            except Exception as exc:
                st.error(f"Não foi possível abrir a planilha. Verifique se é um .xlsx válido: {exc}")
                st.stop()

            st.caption(f"Arquivo selecionado: {uploaded.name} · {len(content) / 1024:.0f} KB")
            with st.expander("Conferir dados da planilha", expanded=False):
                left, right = st.columns(2, gap="large")
                with left:
                    st.subheader("Alunos")
                    if "Alunos" in book:
                        st.dataframe(preview(book["Alunos"]), hide_index=True, width="stretch", height=320,
                                     alt="Prévia das primeiras linhas da aba Alunos")
                        count = book["Alunos"].max_row
                        st.caption("Até 12 linhas exibidas" + (f" · {count - 1} linhas no arquivo" if count else ""))
                    else:
                        st.warning("A aba Alunos não foi encontrada.")
                with right:
                    st.subheader("Turmas")
                    if "Turmas" in book:
                        st.dataframe(preview(book["Turmas"]).iloc[:, :2], hide_index=True, width="stretch", height=320,
                                     alt="Prévia dos recursos e quantidades da aba Turmas")
                    else:
                        st.warning("A aba Turmas não foi encontrada.")

            students, resources, _, _ = inspect_workbook(book)
            book.close()
            with st.container(key="abc_settings", border=True, gap="medium"):
                st.header("2. Organizar as turmas")
                st.write("Escolha como usar as salas disponíveis antes de gerar o ensalamento.")
                allow_empty = st.checkbox("Prefiro deixar salas vazias quando possível", value=False,
                                          help="Usa o menor número de salas que comporta todos os alunos. Desmarcado, ocupa todas as salas disponíveis.")
                st.caption(
                    "Desmarcado: utiliza todas as salas disponíveis.\n\n"
                    "Marcado: utiliza o menor número de salas necessário. As salas ocupadas "
                    "recebem alunos com diferença de no máximo um entre elas no mesmo turno."
                )
                choices = room_choices_for_mode(students, resources, allow_empty)
                with st.container(key="abc_room_summary"):
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
                    if blocking:
                        st.error("Há dados que precisam de correção. Gere o relatório abaixo para consultar as inconsistências.")
                    else:
                        st.warning("Há avisos para conferir. O ensalamento pode ser gerado e os avisos serão incluídos no Excel.")
                    with st.expander(f"{len(issues)} aviso(s) e inconsistência(s)", expanded=blocking):
                        for issue in issues:
                            st.write("• " + issue)
                else:
                    st.success("Dados consistentes e capacidade suficiente para a configuração escolhida.")

                signature = (digest, tuple(sorted(choices.items())))
                if st.session_state.get("result_signature") != signature:
                    st.session_state.pop("result", None)
                with st.container(key="abc_run", width=400):
                    generate = st.button("Gerar ensalamento" if not blocking else "Gerar relatório de inconsistências",
                                         type="primary", width="stretch")
                if generate:
                    try:
                        with st.status("Preparando o resultado…", expanded=True) as progress:
                            st.write("Abrindo a planilha enviada.")
                            result_book = load_workbook(BytesIO(content))
                            st.write("Validando os dados e gerando o relatório." if blocking else
                                     "Validando os dados e calculando a distribuição das turmas.")
                            result = make_result(result_book, choices)
                            result_book.close()
                            progress.update(label="Relatório de inconsistências gerado." if result[1] else
                                            "Ensalamento concluído.", state="complete", expanded=False)
                        st.session_state.result = result
                        st.session_state.result_signature = signature
                    except Exception as exc:
                        st.error(f"Não foi possível gerar o resultado: {exc}")

            if "result" not in st.session_state:
                st.caption("O resultado aparecerá abaixo após o processamento.")
                st.stop()
            result_bytes, blocked, result_issues, stats = st.session_state.result
            with st.container(key="abc_results", border=True, gap="medium"):
                st.header("3. Consultar e baixar o resultado")
                name = Path(uploaded.name).stem
                if blocked:
                    st.error("O ensalamento foi bloqueado. Corrija os dados indicados e envie a planilha novamente.")
                    st.download_button("Baixar planilha com inconsistências", result_bytes,
                                       file_name=f"{name}_inconsistencias.xlsx",
                                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                    st.dataframe(output_table(result_bytes, "Inconsistências"), hide_index=True, width="stretch",
                                 alt="Inconsistências que precisam ser conferidas e corrigidas na planilha")
                    st.stop()

                st.success("Ensalamento concluído. Confira a distribuição e baixe os arquivos abaixo.")
                if result_issues:
                    st.caption(f"O resultado contém {len(result_issues)} aviso(s). Consulte a aba Inconsistências antes de utilizar a planilha.")
                total = sum(row[1] for row in stats)
                with st.container(key="abc_totals"):
                    metrics = st.columns(4)
                    for col, label, number in zip(metrics,
                                                  ("Alunos alocados", "Turmas ocupadas", "Meninas", "Meninos"),
                                                  (total, len(stats), sum(row[2] for row in stats), sum(row[3] for row in stats))):
                        col.metric(label, number)
                room_table = pd.DataFrame(stats, columns=["Turma", "Alunos", "Meninas", "Meninos"])
                room_table["% meninas"] = (100 * room_table["Meninas"] / room_table["Alunos"]).round(1)

                with st.container(key="abc_downloads"):
                    c1, c2 = st.columns(2)
                    c1.download_button("Baixar Excel completo (.xlsx)", result_bytes,
                                       file_name=f"{name}_resultado.xlsx",
                                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                       type="primary", width="stretch")
                    c2.download_button("Baixar resumo por turma (.csv)",
                                       room_table.to_csv(index=False).encode("utf-8-sig"),
                                       file_name=f"{name}_resumo.csv", mime="text/csv", width="stretch")
                    st.caption("O Excel inclui as abas de entrada, o ensalamento e as inconsistências. O CSV contém o resumo por turma.")

                dash_tab, list_tab, warnings_tab = st.tabs(["Visão geral", "Lista de alunos", "Inconsistências"])
                with dash_tab:
                    st.subheader("Distribuição por turma")
                    st.bar_chart(room_table, x="Turma", y=["Meninas", "Meninos"], stack=True,
                                 color=["#2F9587", "#2B469A"], height=300, sort=False,
                                 alt="Quantidade de meninas e meninos em cada turma; os mesmos valores estão na tabela abaixo")
                    st.dataframe(room_table, hide_index=True, width="stretch",
                                 column_config={"% meninas": st.column_config.NumberColumn("% meninas", format="%.1f%%")},
                                 alt="Resumo por turma com total de alunos, meninas, meninos e proporção feminina")
                with list_tab:
                    result_table = output_table(result_bytes, "Ensalamento")
                    search = st.text_input("Buscar aluno ou turma", placeholder="Digite um nome ou letra da turma")
                    if search:
                        result_table = result_table[result_table.astype(str).apply(
                            lambda column: column.str.contains(search, case=False, regex=False)).any(axis=1)]
                    if result_table.empty:
                        st.info("Nenhum aluno ou turma encontrado para essa busca. Tente outro nome ou letra.")
                    st.dataframe(result_table, hide_index=True, width="stretch", height=450,
                                 alt="Lista de alunos alocados com nome, sexo, disponibilidade e turma")
                with warnings_tab:
                    st.dataframe(output_table(result_bytes, "Inconsistências"), hide_index=True, width="stretch",
                                 alt="Avisos e inconsistências registrados no arquivo de resultado")
