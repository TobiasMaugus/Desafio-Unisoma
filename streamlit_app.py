"""Interface local do ensalamento da ABC Aprendiz."""

from io import BytesIO
from pathlib import Path
import hashlib
import json
import re
import unicodedata

from openpyxl import load_workbook
import pandas as pd
import streamlit as st
from st_aggrid import AgGrid, DataReturnMode, GridOptionsBuilder, JsCode

from app import OPTIONAL_RESOURCES, RESOURCES, format_sheet, inspect_workbook, make_result, room_choices_for_mode


st.set_page_config(page_title="ABC Aprendiz · Ensalamento",
                   page_icon=str(Path(__file__).with_name("abc-logo-1024x798.webp")),
                   layout="wide")
st.html(Path(__file__).parent / "assets" / "interface.css")


def output_table(content, sheet_name):
    book = load_workbook(BytesIO(content), read_only=True, data_only=True)
    rows = book[sheet_name].values
    columns = next(rows)
    table = pd.DataFrame(rows, columns=columns)
    book.close()
    return table


def student_rows(book):
    """Keep the Excel row as a stable registration key, even for duplicate names."""
    if "Alunos" not in book:
        return []
    sheet = book["Alunos"]
    return [[row, sheet.cell(row, 1).value, sheet.cell(row, 2).value, sheet.cell(row, 3).value]
            for row in range(2, sheet.max_row + 1)
            if any(sheet.cell(row, col).value not in (None, "") for col in (1, 2, 3))]


def resource_rows(book):
    if "Turmas" not in book:
        return []
    sheet = book["Turmas"]
    return [[row, sheet.cell(row, 1).value, sheet.cell(row, 2).value]
            for row in range(2, sheet.max_row + 1)
            if any(sheet.cell(row, col).value not in (None, "") for col in (1, 2))]


def workbook_with_rows(content, rows, resources=None):
    book = load_workbook(BytesIO(content))
    if "Alunos" in book:
        sheet = book["Alunos"]
        for row in range(2, sheet.max_row + 1):
            for col in (1, 2, 3):
                sheet.cell(row, col).value = None
        for excel_row, record in enumerate(rows, 2):
            for col, value in enumerate(record[1:4], 1):
                sheet.cell(excel_row, col).value = value
    if resources is not None and "Turmas" in book:
        sheet = book["Turmas"]
        for row in range(2, sheet.max_row + 1):
            sheet.cell(row, 1).value = None
            sheet.cell(row, 2).value = None
        for excel_row, record in enumerate(resources, 2):
            sheet.cell(excel_row, 1).value = record[1]
            sheet.cell(excel_row, 2).value = record[2]
    return book


def issue_guidance(message):
    """Add a concise next step to a validation message without changing exports."""
    prefix, separator, detail = str(message).partition(": ")
    if not separator:
        prefix, detail = "", str(message)
    normalized = normalize_token(detail)
    if "este nome aparece mais de uma vez" in normalized:
        guidance = ("Problema: nome repetido. Solução: confira o cadastro; se forem pessoas diferentes, "
                    "mantenha os dois registros.")
    elif "valor padronizado automaticamente" in normalized:
        guidance = f"Problema: valor fora do padrão. Solução: confira a correção aplicada. ({detail})"
    elif "este registro está vazio" in normalized:
        guidance = ("Problema: registro sem dados. Solução: preencha Nome, Sexo e Disponibilidade "
                    "ou exclua a linha.")
    else:
        guidance_parts = []
        for item in detail.rstrip(". ").split("; "):
            item_normalized = normalize_token(item)
            sex_match = re.search(r"sexo como F ou M\s*\((.*?)\)", item, flags=re.IGNORECASE)
            period_match = re.search(r"disponibilidade como Manhã ou Tarde\s*\((.*?)\)", item,
                                     flags=re.IGNORECASE)
            if sex_match or "sexo como f ou m" in item_normalized:
                value = sex_match.group(1) if sex_match else "valor inválido"
                guidance_parts.append(f"Problema: sexo inválido ({value}). Solução: informe F ou M.")
            elif period_match or "disponibilidade como manha ou tarde" in item_normalized:
                value = period_match.group(1) if period_match else "valor inválido"
                guidance_parts.append(f"Problema: disponibilidade inválida ({value}). Solução: informe Manhã ou Tarde.")
            elif "informe o nome do aluno" in item_normalized:
                guidance_parts.append("Problema: nome ausente ou inválido. Solução: preencha o campo Nome.")
            else:
                guidance_parts.append(f"Problema: {item}. Solução: revise e corrija o valor indicado.")
        guidance = " ".join(guidance_parts)

    return f"{prefix}: {guidance}" if prefix else guidance


def roster_editor(rows, issues, corrections=()):
    attention_by_line = {}
    for issue in issues:
        match = re.match(r"Alunos, linhas? ([\d, ]+):", issue)
        if match:
            level, message = friendly_issue(issue)
            for line in re.findall(r"\d+", match.group(1)):
                attention_by_line.setdefault(int(line), []).append(f"{level}: {issue_guidance(message)}")
    corrections_by_id = {}
    for correction in corrections:
        field = "Sexo" if correction["campo"] == "sexo" else "Disponibilidade"
        corrections_by_id.setdefault(correction["registro"], []).append(
            f"Correção automática em {field}: {correction['original']} → {correction['corrigido']}. "
            "Solução: confira o valor corrigido.")
    values = []
    for position, (record_id, name, sex, period) in enumerate(rows, 2):
        attention = attention_by_line.get(position, []) + corrections_by_id.get(record_id, [])
        values.append({"Registro": record_id, "Nome": name, "Sexo": sex,
                       "Disponibilidade": period, "Atenção": "; ".join(attention) or "—"})
    return pd.DataFrame(values, columns=["Registro", "Nome", "Sexo", "Disponibilidade", "Atenção"])


def styled_rows(frame, theme_type=None):
    theme_type = theme_type or st.context.theme.type
    palette = {
        "dark": {
            "error_bg": "#48232B", "error_text": "#FFD9DE",
            "warning_bg": "#4A3714", "warning_text": "#FFE8B3",
            "corrected_bg": "#183B35", "corrected_text": "#A6E3D5",
        },
        "light": {
            "error_bg": "#FCEEF0", "error_text": "#702A33",
            "warning_bg": "#FFF1CC", "warning_text": "#654B08",
            "corrected_bg": "#EAF5F1", "corrected_text": "#20665B",
        },
    }.get(theme_type, {
        "error_bg": "#FCEEF0", "error_text": "#702A33",
        "warning_bg": "#FFF1CC", "warning_text": "#654B08",
        "corrected_bg": "#EAF5F1", "corrected_text": "#20665B",
    })

    def paint(row):
        label = str(row.get("Status", row.get("Atenção", "")))
        if label.startswith("Erro") or label.startswith("Atenção"):
            background, foreground = palette["error_bg"], palette["error_text"]
        elif label.startswith("Corrigido automaticamente") or label.startswith("Correção automática"):
            background, foreground = palette["corrected_bg"], palette["corrected_text"]
        elif label.startswith("Aviso"):
            background, foreground = palette["warning_bg"], palette["warning_text"]
        else:
            return ["" for _ in row]
        style = f"background-color: {background}; color: {foreground};"
        return [style for _ in row]

    return frame.style.apply(paint, axis=1)


def student_grid(frame, *, key, next_record_id, theme_type=None):
    theme_type = theme_type or st.context.theme.type
    colors = {
        "dark": {
            "error": ("#48232B", "#FFD9DE"),
            "warning": ("#4A3714", "#FFE8B3"),
            "corrected": ("#183B35", "#A6E3D5"),
        },
        "light": {
            "error": ("#FCEEF0", "#702A33"),
            "warning": ("#FFF1CC", "#654B08"),
            "corrected": ("#EAF5F1", "#20665B"),
        },
    }.get(theme_type, {
        "error": ("#FCEEF0", "#702A33"),
        "warning": ("#FFF1CC", "#654B08"),
        "corrected": ("#EAF5F1", "#20665B"),
    })
    add_row_colors = ("#293946", "#E2E8F0") if theme_type == "dark" else ("#EDF2F6", "#35495A")
    base_bg, header_bg, border, text_color, header_text = (
        ("#111A23", "#273542", "#3A4A58", "#E8EEF4", "#E8EEF4")
        if theme_type == "dark" else
        ("#F5F7F9", "#EDF2F6", "#CFD9E2", "#23364B", "#23364B")
    )
    grid_css = {
        ".ag-theme-streamlit": {
            "--ag-background-color": base_bg,
            "--ag-foreground-color": text_color,
            "--ag-header-background-color": header_bg,
            "--ag-header-foreground-color": header_text,
            "--ag-border-color": border,
            "--ag-row-border-color": border,
            "--ag-font-family": "'Segoe UI', Arial, sans-serif",
            "--ag-font-size": "14px",
            "--ag-header-height": "38px",
            "--ag-row-height": "36px",
        },
        ".ag-root-wrapper": {
            "background-color": base_bg,
            "border": f"1px solid {border}",
            "border-radius": "6px",
        },
        ".ag-root, .ag-body, .ag-body-viewport, .ag-center-cols-viewport": {
            "background-color": base_bg,
        },
        ".ag-header, .ag-header-row, .ag-header-cell": {
            "background-color": header_bg,
            "color": header_text,
            "border-color": border,
        },
        ".ag-header-cell-text": {
            "font-family": "'Segoe UI', Arial, sans-serif",
            "font-size": "14px",
            "font-weight": "500",
        },
        ".ag-row": {
            "background-color": base_bg,
            "color": text_color,
            "border-color": border,
            "font-family": "'Segoe UI', Arial, sans-serif",
            "font-size": "14px",
        },
        ".ag-cell": {
            "background-color": "transparent",
            "color": "inherit",
            "border-right": f"1px solid {border}",
            "padding-left": "8px",
        },
        ".ag-row-pinned, .ag-row-pinned .ag-cell": {
            "background-color": add_row_colors[0],
            "color": add_row_colors[1],
        },
    }
    data = frame.copy()
    data["__grid_key"] = [f"aluno-{record_id}" for record_id in data["Registro"]]
    blank_row = {column: None for column in data.columns}
    blank_row.update({"Registro": int(next_record_id), "__grid_key": f"aluno-{int(next_record_id)}"})

    options = GridOptionsBuilder.from_dataframe(data)
    options.configure_default_column(
        editable=True, resizable=True, sortable=False, filter=False,
        suppressMovable=True, wrapText=False,
        valueFormatter=JsCode("function(params) { return params.value == null ? 'None' : params.value; }"),
        cellStyle=JsCode(f"""
function(params) {{
  if (params.node.rowPinned === 'bottom') return {{backgroundColor: {json.dumps(add_row_colors[0])}, color: {json.dumps(add_row_colors[1])}}};
  const label = String((params.data && params.data['Atenção']) || '');
  let pair = null;
  if (label.startsWith('Erro') || label.startsWith('Atenção')) pair = {json.dumps(colors['error'])};
  else if (label.startsWith('Aviso')) pair = {json.dumps(colors['warning'])};
  else if (label.startsWith('Correção automática') || label.startsWith('Corrigido automaticamente')) pair = {json.dumps(colors['corrected'])};
  return pair ? {{backgroundColor: pair[0], color: pair[1]}} : null;
}}
"""),
    )
    options.configure_column("__grid_key", hide=True, editable=False)
    options.configure_column("Registro", editable=False, width=76, minWidth=68)
    options.configure_column("Nome", editable=True, flex=2.7, minWidth=220)
    options.configure_column("Sexo", editable=True, width=72, minWidth=64)
    options.configure_column("Disponibilidade", editable=True, flex=1.3, minWidth=150)
    options.configure_column("Atenção", editable=False, flex=1.8, minWidth=250,
                             cellRenderer=JsCode("""
function(params) {
  return params.node.rowPinned === 'bottom' ? '+' : params.value;
}
"""))
    options.configure_grid_options(
        context={"blankRow": blank_row, "nextRecordId": int(next_record_id)},
        getRowId=JsCode("function(params) { return params.data.__grid_key; }"),
        onCellClicked=JsCode("""
function(params) {
  if (params.colDef.field === 'Atenção' && params.node.rowPinned === 'bottom') {
    const row = {...params.context.blankRow};
    let maxId = Number(params.context.nextRecordId) - 1;
    params.api.forEachNode(node => {
      if (!node.rowPinned && node.data && Number.isFinite(Number(node.data.Registro))) {
        maxId = Math.max(maxId, Number(node.data.Registro));
      }
    });
    row.Registro = maxId + 1;
    row.__grid_key = `aluno-${row.Registro}`;
    const added = params.api.applyTransaction({add: [row]});
    const addedNode = added && added.add && added.add[0];
    if (addedNode) {
      setTimeout(() => {
        params.api.ensureIndexVisible(addedNode.rowIndex, 'bottom');
      }, 0);
    }
  }
}
"""),
        getRowStyle=JsCode(f"""
function(params) {{
  if (params.node.rowPinned === 'bottom') return {{backgroundColor: {json.dumps(add_row_colors[0])}, color: {json.dumps(add_row_colors[1])}}};
  const label = String((params.data && params.data['Atenção']) || '');
  let pair = null;
  if (label.startsWith('Erro') || label.startsWith('Atenção')) pair = {json.dumps(colors['error'])};
  else if (label.startsWith('Aviso')) pair = {json.dumps(colors['warning'])};
  else if (label.startsWith('Correção automática') || label.startsWith('Corrigido automaticamente')) pair = {json.dumps(colors['corrected'])};
  return pair ? {{backgroundColor: pair[0], color: pair[1]}} : null;
}}
"""),
        pinnedBottomRowData=[{"Atenção": "+", "__grid_key": "__add_student_row__"}],
        rowSelection={
            "mode": "multiRow",
            "checkboxes": JsCode("function(params) { return !params.node.rowPinned; }"),
            "headerCheckbox": False,
            "enableClickSelection": False,
        },
        selectionColumnDef={
            "width": 40, "minWidth": 40, "maxWidth": 40,
            "pinned": "left", "lockPosition": "left", "resizable": False,
        },
        isRowSelectable=JsCode("function(params) { return !params.node.rowPinned; }"),
        singleClickEdit=True,
        headerHeight=38,
        rowHeight=36,
        animateRows=False,
        stopEditingWhenCellsLoseFocus=True,
        suppressMovableColumns=True,
        suppressCellFocus=False,
    )
    result = AgGrid(
        data,
        gridOptions=options.build(),
        data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
        update_on=["cellValueChanged", "selectionChanged", "rowDataUpdated"],
        height=400,
        width="100%",
        fit_columns_on_grid_load=False,
        theme="streamlit",
        show_toolbar=False,
        show_search=False,
        show_download_button=False,
        custom_css=grid_css,
        allow_unsafe_jscode=True,
        key=key,
    )
    edited = result["data"]
    if not isinstance(edited, pd.DataFrame):
        edited = pd.DataFrame(edited)
    selected = result.get("selected_rows")
    if isinstance(selected, pd.DataFrame):
        selected_ids = selected.get("Registro", pd.Series(dtype=object)).dropna().tolist()
    elif isinstance(selected, list):
        selected_ids = [row.get("Registro") for row in selected if row.get("Registro") is not None]
    else:
        selected_ids = []
    return edited.drop(columns=["__grid_key"], errors="ignore"), selected_ids


def history_controls(table, *, original, label):
    undo_key, redo_key = f"{table}_undo", f"{table}_redo"
    revision_key = "editor_revision" if table == "roster" else "resources_revision"
    revision = st.session_state[revision_key]
    correction_key = f"{table}_corrections"
    with st.expander(f"Histórico e restauração {label}", expanded=False):
        undo_col, redo_col = st.columns(2, gap="small")
        if undo_col.button("Desfazer", key=f"{table}_undo_button", disabled=not st.session_state[undo_key],
                           width="stretch"):
            previous = st.session_state[undo_key].pop()
            st.session_state[redo_key].append({
                "rows": st.session_state[f"committed_{table}"],
                "corrections": st.session_state[correction_key],
            })
            st.session_state[f"committed_{table}"] = previous["rows"]
            st.session_state[correction_key] = previous["corrections"]
            if table == "roster":
                st.session_state.roster_draft_rows = [list(row) for row in previous["rows"]]
            st.session_state[revision_key] += 1
            invalidate_result()
            st.rerun()
        if redo_col.button("Refazer", key=f"{table}_redo_button", disabled=not st.session_state[redo_key],
                           width="stretch"):
            next_value = st.session_state[redo_key].pop()
            st.session_state[undo_key].append({
                "rows": st.session_state[f"committed_{table}"],
                "corrections": st.session_state[correction_key],
            })
            st.session_state[f"committed_{table}"] = next_value["rows"]
            st.session_state[correction_key] = next_value["corrections"]
            if table == "roster":
                st.session_state.roster_draft_rows = [list(row) for row in next_value["rows"]]
            st.session_state[revision_key] += 1
            invalidate_result()
            st.rerun()
        confirmed = st.checkbox(f"Confirmo restaurar {label} originais",
                                key=f"{table}_confirm_reset_{revision}")
        if st.button("Restaurar dados originais", key=f"{table}_reset_{revision}",
                     disabled=not confirmed):
            st.session_state[f"committed_{table}"] = [list(row) for row in original]
            if table == "roster":
                st.session_state.roster_draft_rows = [list(row) for row in original]
            st.session_state[undo_key] = []
            st.session_state[redo_key] = []
            st.session_state[f"{table}_corrections"] = []
            st.session_state[revision_key] += 1
            invalidate_result()
            st.rerun()


def invalidate_result():
    for key in ("result", "result_signature", "result_room_mode", "result_issue_records", "manual_assignments", "manual_original",
                "manual_history", "manual_signature", "manual_notice"):
        st.session_state.pop(key, None)
    st.session_state.room_mode_choice = "Todas as salas disponíveis"


def issue_records(issues, corrections=(), roster=()):
    records = [friendly_issue(issue) for issue in issues]
    record_lines = {row[0]: position for position, row in enumerate(roster, 2)}
    for item in corrections:
        if item.get("registro") is not None and item["registro"] not in record_lines:
            continue
        field = "sexo" if item["campo"] == "sexo" else "disponibilidade" if item["campo"] == "disponibilidade" else "configuração de turma"
        line = record_lines.get(item.get("registro"), item.get("linha", "?"))
        area = "Turmas" if field.startswith("configuração") else "Alunos"
        records.append(("Correção automática",
                        f"{area}, linha {line}: {field} corrigido automaticamente de “{item['original']}” para “{item['corrigido']}”."))
    return records


def rewrite_result_issues(content, records):
    if not records:
        return content
    book = load_workbook(BytesIO(content))
    sheet = book["Inconsistências"]
    if sheet.max_row > 1:
        sheet.delete_rows(2, sheet.max_row - 1)
    for level, message in records:
        sheet.append((level, message))
    if not records:
        sheet.append(("OK", "Nenhuma inconsistência encontrada."))
    format_sheet(sheet, {"A": 16, "B": 100})
    result = make_bytes(book)
    book.close()
    return result


def inspect_entry(book, roster, choices=None, resource_records=()):
    students, resources, issues, blocking = inspect_workbook(book, choices)
    for position, row in enumerate(roster, 2):
        if all(value is None or str(value).strip() == "" for value in row[1:4]):
            issues.append(f"Alunos, linha {position}: registro vazio; preencha o registro ou exclua-o com confirmação.")
            blocking = True
    for position, row in enumerate(resource_records, 2):
        if all(value is None or str(value).strip() == "" for value in row[1:3]):
            issues.append(f"Turmas, linha {position}: configuração vazia; preencha ou exclua a linha com confirmação.")
            blocking = True
    return students, resources, issues, blocking


def make_blocked_report(content, roster, resources, records):
    book = workbook_with_rows(content, roster, resources)
    for title in ("Ensalamento", "Inconsistências"):
        if title in book:
            del book[title]
    assignments = book.create_sheet("Ensalamento")
    assignments.append(("Nome", "Sexo", "Disponibilidade", "Turma"))
    format_sheet(assignments, {"A": 45, "B": 12, "C": 20, "D": 20})
    diagnostics = book.create_sheet("Inconsistências")
    diagnostics.append(("Situação", "Descrição"))
    for level, message in records:
        diagnostics.append((level, message))
    format_sheet(diagnostics, {"A": 16, "B": 100})
    output = make_bytes(book)
    book.close()
    return output, True, [message for _, message in records], []


def rows_from_editor(frame, existing_rows):
    def clean(value):
        return None if pd.isna(value) else value
    new_id = max((int(row[0]) for row in existing_rows), default=1) + 1
    rows = []
    for _, row in frame.iterrows():
        record_id = clean(row["Registro"])
        name, sex, period = (clean(row[column]) for column in ("Nome", "Sexo", "Disponibilidade"))
        if record_id is None and all(value is None or str(value).strip() == "" for value in (name, sex, period)):
            continue
        if record_id is None:
            record_id, new_id = new_id, new_id + 1
        rows.append([int(record_id), name, sex, period])
    return rows


def merge_roster_editor(frame, existing_rows, deleted_ids=()):
    returned_rows = rows_from_editor(frame, existing_rows)
    deleted_ids = {int(record_id) for record_id in deleted_ids}
    existing_ids = {int(row[0]) for row in existing_rows}
    returned_by_id = {int(row[0]): row for row in returned_rows if row[0] is not None}
    merged = [returned_by_id.pop(int(row[0]), list(row)) for row in existing_rows
              if int(row[0]) not in deleted_ids]
    merged.extend(row for row in returned_rows if row[0] is not None and int(row[0]) not in existing_ids)
    return merged


def resources_from_editor(frame, existing_rows):
    def clean(value):
        return None if pd.isna(value) else value
    new_id = max((int(row[0]) for row in existing_rows), default=1) + 1
    rows = []
    for _, row in frame.iterrows():
        record_id = clean(row["Registro"])
        key, quantity = clean(row["Recursos"]), clean(row["Quantidade"])
        if record_id is None and all(value is None or str(value).strip() == "" for value in (key, quantity)):
            continue
        if record_id is None:
            record_id, new_id = new_id, new_id + 1
        rows.append([int(record_id), key, quantity])
    return rows


def styled_resource_rows(frame, issues, line_positions=None):
    attention_by_line = {}
    for issue in issues:
        match = re.match(r"Turmas, linha (\d+):", issue)
        if match:
            level, message = friendly_issue(issue)
            attention_by_line[int(match.group(1)) - 2] = f"{level}: {issue_guidance(message)}"
    styled = frame.copy()
    positions = line_positions or list(range(len(styled)))
    styled["Atenção"] = [attention_by_line.get(position, "—") for position in positions]
    return styled


def normalize_token(value):
    text = "" if value is None else str(value)
    text = " ".join(text.strip().casefold().split())
    return "".join(char for char in unicodedata.normalize("NFKD", text)
                   if not unicodedata.combining(char))


EQUIVALENCES = {
    "sexo": {
        "f": "F", "feminino": "F", "feminina": "F", "mulher": "F",
        "garota": "F", "menina": "F",
        "m": "M", "masculino": "M", "masculina": "M", "homem": "M",
        "garoto": "M", "menino": "M",
    },
    "disponibilidade": {
        "manha": "Manhã", "matutino": "Manhã", "matutina": "Manhã",
        "periodo da manha": "Manhã", "periodo manha": "Manhã",
        "tarde": "Tarde", "vespertino": "Tarde", "vespertina": "Tarde",
        "periodo da tarde": "Tarde", "periodo tarde": "Tarde",
    },
}
RESOURCE_EQUIVALENCES = {normalize_token(key): key for key in (
    "turmas_manha", "turmas_tarde", "vagas_manha", "vagas_tarde",
    "turmas_usadas_manha", "turmas_usadas_tarde")}


def normalize_roster(rows):
    normalized, corrections = [], []
    for position, row in enumerate(rows, 2):
        record = list(row)
        for column, field in ((2, "sexo"), (3, "disponibilidade")):
            value = record[column]
            replacement = EQUIVALENCES[field].get(normalize_token(value))
            if replacement is not None and value != replacement:
                corrections.append({"registro": record[0], "linha": position, "campo": field,
                                    "original": value, "corrigido": replacement})
                record[column] = replacement
        normalized.append(record)
    return normalized, corrections


def normalize_resources(rows):
    normalized, corrections = [], []
    for position, row in enumerate(rows, 2):
        record = list(row)
        key = record[1]
        replacement = RESOURCE_EQUIVALENCES.get(normalize_token(key))
        if replacement is not None and key != replacement:
            corrections.append({"linha": position, "campo": "recurso", "original": key,
                                "corrigido": replacement})
            record[1] = replacement
        normalized.append(record)
    return normalized, corrections


def friendly_issue(issue):
    text = str(issue)
    lower = normalize_token(text)
    level = "Aviso"
    if "nome duplicado" in lower:
        return level, text.replace("nome duplicado; registros mantidos como pessoas distintas.",
                                   "este nome aparece mais de uma vez; os registros continuam separados.")
    if "texto padronizado" in lower:
        return "Correção automática", text.replace("texto padronizado", "valor padronizado automaticamente")
    level = "Erro" if any(word in lower for word in (
        "invalido", "ausente", "precisa ser", "desconhecido", "repetido", "cabecalhos",
        "faltam", "sao necessarias as abas", "nenhum aluno valido", "bloqueio", "registro vazio",
        "configuracao vazia")) else level
    if "registro vazio" in lower:
        return "Erro", text.replace("registro vazio; preencha o registro ou exclua-o com confirmação.",
                                    "este registro está vazio; preencha nome, sexo e disponibilidade ou exclua-o com confirmação.")
    if "configuração vazia" in text:
        return "Erro", text.replace(
            "configuração vazia; preencha ou exclua a linha com confirmação.",
            "esta linha de configuração está vazia; escolha uma configuração ou exclua-a com confirmação.")
    if "recurso" in lower and "repetido" in lower:
        return "Erro", text.replace("recurso", "configuração").replace(
            "repetido", "está repetida; mantenha apenas uma linha para essa configuração")
    replacements = (
        ("são necessárias as abas Alunos e Turmas.", "A planilha precisa conter as abas “Alunos” e “Turmas”."),
        ("cabeçalhos esperados: Nome, Sexo, Disponibilidade.",
         "Na aba Alunos, use os títulos Nome, Sexo e Disponibilidade nas três primeiras colunas."),
        ("cabeçalhos esperados: Recursos, Quantidade.",
         "Na aba Turmas, use os títulos Recursos e Quantidade nas duas primeiras colunas."),
        ("nenhum aluno válido encontrado.", "Cadastre ao menos um aluno completo na aba Alunos."),
        ("nome ausente ou inválido", "informe o nome do aluno"),
        ("sexo inválido", "informe o sexo como F ou M"),
        ("disponibilidade inválida", "informe a disponibilidade como Manhã ou Tarde"),
        ("recurso desconhecido", "esse recurso não é reconhecido; use um dos nomes indicados no formato da planilha"),
        ("precisa ser inteiro positivo.", "informe uma quantidade inteira maior que zero."),
        ("recurso", "configuração"),
        ("duplicado.", "está repetido; mantenha somente uma linha para cada configuração."),
        ("ausente.", "está faltando; inclua essa configuração na aba Turmas."),
        ("turmas solicitadas, mas apenas", "turmas foram informadas, mas existem apenas"),
        ("turmas não vazias exigem pelo menos", "turmas ocupadas precisam de pelo menos"),
        ("alunos para apenas", "alunos, mas há somente"),
        ("vagas; faltam", "vagas; reduza o número de alunos ou aumente a capacidade"),
        ("rodapé 'Total' ignorado.", "a linha de total foi ignorada; os alunos cadastrados foram mantidos."),
    )
    for old, new in replacements:
        text = text.replace(old, new)
    return level, text


def report_pdf(title, stats, capacity_by_room, issues, students_total, adjusted=False):
    """Create a small branded PDF using only the standard PDF and WinAnsi fonts."""
    lines = [f"ABC Aprendiz · {title}",
             "Distribuição ajustada manualmente" if adjusted else "Solução otimizada pelo modelo",
             f"Total de alunos alocados: {students_total}", "", "Ocupação e composição por turma:"]
    by_period = {"Manhã": [], "Tarde": []}
    for room, count, girls, boys in stats:
        period = "Manhã" if str(room).startswith("MANHÃ") else "Tarde"
        capacity = capacity_by_room.get(room, 0)
        ratio = (100 * girls / count) if count else 0
        lines.append(f"{room} · {period}: {count}/{capacity} alunos · {girls} meninas ({ratio:.1f}%) · {boys} meninos")
        by_period[period].append((count, ratio))
    lines.extend(["", "Indicadores de equilíbrio:"])
    for period, groups in by_period.items():
        if groups:
            sizes = [item[0] for item in groups]
            ratios = [item[1] for item in groups]
            lines.append(f"{period}: diferença de ocupação {max(sizes) - min(sizes)} aluno(s); "
                         f"variação de meninas {max(ratios) - min(ratios):.1f} pontos percentuais.")
    lines.extend(["", "Avisos e inconsistências:"])
    lines.extend([f"{level}: {message}" for level, message in issues] or ["Nenhuma inconsistência encontrada."])

    pages, page_lines = [], []
    for line in lines:
        wrapped = [line[i:i + 88] for i in range(0, max(len(line), 1), 88)] or [""]
        for fragment in wrapped:
            page_lines.append(fragment)
            if len(page_lines) == 48:
                pages.append(page_lines)
                page_lines = []
    if page_lines:
        pages.append(page_lines)

    objects: list[bytes] = []
    def add_object(value: bytes) -> int:
        objects.append(value)
        return len(objects)
    catalog_id = add_object(b"")
    pages_id = add_object(b"")
    font_id = add_object(b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>")
    page_ids = []
    for page_number, page in enumerate(pages, 1):
        commands = ["q 0.055 0.463 0.737 rg 0 780 595 62 re f Q",
                    "BT /F1 16 Tf 42 810 Td (ABC Aprendiz) Tj ET",
                    f"BT /F1 9 Tf 510 810 Td ({page_number}/{len(pages)}) Tj ET"]
        y = 755
        for line in page:
            size = 12 if line in ("Ocupação e composição por turma:", "Indicadores de equilíbrio:",
                                  "Avisos e inconsistências:") else 9
            text = line.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")
            commands.append(f"BT /F1 {size} Tf 42 {y} Td ({text}) Tj ET")
            y -= 14 if size == 12 else 12
        stream = "\n".join(commands).encode("cp1252", errors="replace")
        content_id = add_object(b"<< /Length " + str(len(stream)).encode() + b" >>\nstream\n" + stream + b"\nendstream")
        page_id = add_object(f"<< /Type /Page /Parent {pages_id} 0 R /MediaBox [0 0 595 842] /Resources << /Font << /F1 {font_id} 0 R >> >> /Contents {content_id} 0 R >>".encode())
        page_ids.append(page_id)
    objects[catalog_id - 1] = f"<< /Type /Catalog /Pages {pages_id} 0 R >>".encode()
    objects[pages_id - 1] = f"<< /Type /Pages /Kids [{' '.join(f'{page_id} 0 R' for page_id in page_ids)}] /Count {len(page_ids)} >>".encode()
    output = bytearray(b"%PDF-1.4\n%\xe2\xe3\xcf\xd3\n")
    offsets = [0]
    for i, obj in enumerate(objects, 1):
        offsets.append(len(output))
        output.extend(f"{i} 0 obj\n".encode() + obj + b"\nendobj\n")
    xref = len(output)
    output.extend(f"xref\n0 {len(objects) + 1}\n0000000000 65535 f \n".encode())
    output.extend(b"".join(f"{offset:010d} 00000 n \n".encode() for offset in offsets[1:]))
    output.extend(f"trailer\n<< /Size {len(objects) + 1} /Root {catalog_id} 0 R >>\nstartxref\n{xref}\n%%EOF".encode())
    return bytes(output)


def make_bytes(book):
    buffer = BytesIO()
    book.save(buffer)
    return buffer.getvalue()


def room_stats(assignments, students, rooms):
    stats = []
    for room in rooms:
        members = [i for i, assigned in assignments.items() if assigned == room]
        girls = sum(students[i][1] == "F" for i in members)
        boys = sum(students[i][1] == "M" for i in members)
        stats.append((room, len(members), girls, boys))
    return stats


def validate_adjusted(assignments, students, capacities):
    problems = []
    for room, capacity in capacities.items():
        members = [i for i, assigned in assignments.items() if assigned == room]
        if len(members) > capacity:
            problems.append(f"{room} excede a capacidade ({len(members)}/{capacity}).")
        period = "MANHÃ" if room.startswith("MANHÃ") else "TARDE"
        if any(students[i][2].upper() != period for i in members):
            problems.append(f"{room} contém aluno de outro turno.")
    unassigned = sum(assignments.get(i) is None for i in range(len(students)))
    if unassigned:
        problems.append(f"{unassigned} aluno(s) estão sem turma.")
    return problems


def save_adjusted_result(content, assignments, existing_issues, problems):
    book = load_workbook(BytesIO(content))
    sheet = book["Ensalamento"]
    for i, room in assignments.items():
        sheet.cell(i + 2, 4).value = room
    issues_sheet = book["Inconsistências"]
    if not existing_issues:
        for row in range(2, issues_sheet.max_row + 1):
            if issues_sheet.cell(row, 1).value == "OK":
                issues_sheet.delete_rows(row, 1)
                break
        if not problems:
            issues_sheet.append(("OK", "Nenhuma inconsistência encontrada."))
    else:
        for row in range(2, issues_sheet.max_row + 1):
            if str(issues_sheet.cell(row, 2).value or "").startswith("Ajuste manual:"):
                issues_sheet.delete_rows(row, 1)
                break
    unassigned = sum(value is None for value in assignments.values())
    for problem in problems:
        issues_sheet.append(("Aviso", f"Ajuste manual: {problem}"))
    if unassigned and not any("sem turma" in issue for issue in problems):
        issues_sheet.append(("Aviso", f"Ajuste manual: {unassigned} aluno(s) ficaram sem turma."))
    return make_bytes(book)


def changed(assignments, original):
    return any(assignments.get(i) != original.get(i) for i in original)


def start_manual_action(proposed, label):
    history = st.session_state.get("manual_history", [])
    history.append(dict(st.session_state.manual_assignments))
    st.session_state.manual_history = history
    st.session_state.manual_assignments = proposed
    st.session_state.manual_notice = label


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

        manual_pdf = Path(__file__).with_name("MANUAL_USUARIO.pdf").read_bytes()
        show_manual = st.sidebar.toggle("Abrir manual ao lado", value=False, key="show_manual_pdf")
        if show_manual:
            st.sidebar.subheader("Manual completo")
            st.sidebar.iframe(Path(__file__).with_name("MANUAL_USUARIO.pdf"), height=820,
                              alt="Manual completo de ensalamento da ABC Aprendiz")

        with manual_tab:
            with st.container(key="abc_manual", width=900):
                manual = Path(__file__).with_name("MANUAL_USUARIO.md").read_text(encoding="utf-8")
                st.markdown(manual)
                st.download_button("Baixar manual (.pdf)", manual_pdf,
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

            if st.session_state.get("active_upload_digest") != digest:
                st.session_state.active_upload_digest = digest
                st.session_state.room_mode_choice = "Todas as salas disponíveis"
                raw_students = student_rows(book)
                raw_resources = resource_rows(book)
                normalized_students, initial_student_corrections = normalize_roster(raw_students)
                normalized_resources, initial_resource_corrections = normalize_resources(raw_resources)
                st.session_state.original_roster = raw_students
                st.session_state.original_resources = raw_resources
                st.session_state.committed_roster = normalized_students
                st.session_state.roster_draft_rows = [list(row) for row in normalized_students]
                st.session_state.committed_resources = normalized_resources
                st.session_state.roster_corrections = initial_student_corrections
                st.session_state.resources_corrections = initial_resource_corrections
                st.session_state.editor_revision = 0
                st.session_state.resources_revision = 0
                st.session_state.roster_undo = []
                st.session_state.roster_redo = []
                st.session_state.resources_undo = []
                st.session_state.resources_redo = []
                invalidate_result()
            source_rows = st.session_state.original_roster
            source_resources = st.session_state.original_resources
            committed_rows = st.session_state.committed_roster
            committed_resources = st.session_state.committed_resources
            st.session_state.setdefault("room_mode_choice", "Todas as salas disponíveis")
            effective_rows, effective_corrections = normalize_roster(committed_rows)
            effective_resources, resource_corrections = normalize_resources(committed_resources)
            effective_corrections = [*st.session_state.roster_corrections,
                                     *st.session_state.resources_corrections,
                                     *effective_corrections, *resource_corrections]
            committed_book = workbook_with_rows(content, effective_rows, effective_resources)
            students, resources, _, _ = inspect_entry(
                committed_book, effective_rows, resource_records=effective_resources)
            st.divider()
            with st.container(key="abc_settings", border=True, gap="medium"):
                st.header("2. Conferir dados e gerar o ensalamento")
                st.write("Confira os dados e gere a primeira distribuição. Depois, você poderá comparar as opções de uso das turmas e recalcular.")
                allow_empty = st.session_state.room_mode_choice == "Somente as salas necessárias"
                choices = room_choices_for_mode(students, resources, allow_empty)

                draft_pending = False
                st.caption("Confira e, se necessário, edite os dados diretamente nas tabelas abaixo.")

                draft_book = workbook_with_rows(content, effective_rows, effective_resources)
                _, _, current_issues, _ = inspect_entry(
                    draft_book, effective_rows, choices, effective_resources)
                st.subheader("Turmas e capacidades")
                resource_key = f"abc_resources_{digest[:12]}_{st.session_state.resources_revision}"
                editable_resource_rows = [row for row in committed_resources
                                          if row[1] not in OPTIONAL_RESOURCES]
                optional_resource_rows = [row for row in committed_resources
                                          if row[1] in OPTIONAL_RESOURCES]
                resource_table = pd.DataFrame(editable_resource_rows,
                                              columns=["Registro", "Recursos", "Quantidade"])
                visible_lines = [index for index, row in enumerate(committed_resources)
                                 if row[1] not in OPTIONAL_RESOURCES]
                resource_styler = styled_resource_rows(resource_table, current_issues, visible_lines)
                edited_resources = st.data_editor(
                    styled_rows(resource_styler), key=resource_key, hide_index=True, width="stretch", num_rows="fixed",
                    disabled=["Registro", "Atenção"],
                    column_config={
                        "Registro": st.column_config.NumberColumn("Registro", format="%d", width="small"),
                        "Recursos": st.column_config.SelectboxColumn(
                            "Configuração", options=list(RESOURCES), width="large"),
                        "Quantidade": st.column_config.NumberColumn("Quantidade", min_value=1, step=1),
                        "Atenção": st.column_config.TextColumn("Atenção", width="large"),
                    })
                draft_raw_editable_resources = resources_from_editor(edited_resources, editable_resource_rows)
                draft_raw_resources = [*draft_raw_editable_resources, *optional_resource_rows]
                draft_resources, draft_resource_corrections = normalize_resources(draft_raw_resources)
                resource_widget_state = st.session_state.get(resource_key, {})
                deleted_resource_rows = resource_widget_state.get("deleted_rows", [])
                confirm_resource_delete = True
                if deleted_resource_rows:
                    confirm_resource_delete = st.checkbox(
                        f"Confirmo excluir {len(deleted_resource_rows)} configuração(ões) de turma",
                        key=f"confirm_resource_delete_{st.session_state.resources_revision}")
                changed_resources = draft_raw_resources != committed_resources or bool(deleted_resource_rows)
                if changed_resources:
                    draft_pending = True
                    st.info("Há alterações de turmas ainda não salvas. Resultados anteriores ficam ocultos.")
                with st.container(key="abc_resource_actions"):
                    save_res_col, discard_res_col = st.columns(2, gap="small")
                    if save_res_col.button("Salvar turmas", type="primary",
                                           disabled=not changed_resources or not confirm_resource_delete,
                                           key="save_resources", width="stretch"):
                        st.session_state.resources_undo.append({
                            "rows": [list(row) for row in committed_resources],
                            "corrections": list(st.session_state.resources_corrections),
                        })
                        st.session_state.resources_redo = []
                        st.session_state.committed_resources = draft_resources
                        st.session_state.resources_corrections.extend(draft_resource_corrections)
                        st.session_state.resources_revision += 1
                        invalidate_result()
                        st.rerun()
                    if discard_res_col.button("Descartar alterações", disabled=not changed_resources,
                                              key="discard_resources", width="stretch"):
                        st.session_state.resources_revision += 1
                        st.rerun()
                history_controls("resources", original=source_resources, label="as turmas")
                draft_book.close()

                st.subheader("Alunos")
                st.caption("Atenção: vermelho indica erro, amarelo indica aviso e verde indica correção automática.")
                draft_base_rows = st.session_state.get(
                    "roster_draft_rows", [list(row) for row in committed_rows])
                roster_key = f"abc_roster_{digest[:12]}_{st.session_state.editor_revision}"
                roster_toolbar = st.empty()
                editor_input = roster_editor(draft_base_rows, current_issues, effective_corrections)
                edited_students, selected_student_ids = student_grid(
                    editor_input, key=roster_key,
                    next_record_id=max((int(row[0]) for row in draft_base_rows), default=1) + 1,
                    theme_type=st.context.theme.type)
                draft_raw_students = merge_roster_editor(edited_students, draft_base_rows)
                st.session_state.roster_draft_rows = [list(row) for row in draft_raw_students]
                draft_students, draft_student_corrections = normalize_roster(draft_raw_students)
                changed_roster = draft_raw_students != committed_rows
                if changed_roster:
                    draft_pending = True
                    st.info("Há alterações de alunos ainda não salvas. Resultados anteriores ficam ocultos.")
                with roster_toolbar.container():
                    with st.container(key="abc_roster_grid_toolbar"):
                        if st.button(
                            " ", icon=":material/delete:", type="secondary", width="content",
                            disabled=not selected_student_ids,
                            key=f"abc_delete_selected_students_{digest[:12]}_{st.session_state.editor_revision}",
                            help="Excluir alunos selecionados",
                        ):
                            deleted_ids = {int(record_id) for record_id in selected_student_ids}
                            st.session_state.roster_draft_rows = [
                                list(row) for row in draft_raw_students
                                if int(row[0]) not in deleted_ids
                            ]
                            st.session_state.editor_revision += 1
                            st.rerun()
                with st.container(key="abc_roster_actions"):
                    save_col, discard_col, excel_col = st.columns(3, gap="small")
                    if save_col.button("Salvar alunos", type="primary",
                                       disabled=not changed_roster,
                                       key="save_roster", width="stretch"):
                        st.session_state.roster_undo.append({
                            "rows": [list(row) for row in committed_rows],
                            "corrections": list(st.session_state.roster_corrections),
                        })
                        st.session_state.roster_redo = []
                        st.session_state.committed_roster = draft_students
                        st.session_state.roster_draft_rows = [list(row) for row in draft_students]
                        st.session_state.roster_corrections.extend(draft_student_corrections)
                        st.session_state.editor_revision += 1
                        invalidate_result()
                        st.rerun()
                    if discard_col.button("Descartar alterações", disabled=not changed_roster,
                                          key="discard_roster", width="stretch"):
                        st.session_state.roster_draft_rows = [list(row) for row in committed_rows]
                        st.session_state.editor_revision += 1
                        st.rerun()
                    corrected_download_book = workbook_with_rows(content, effective_rows, effective_resources)
                    excel_col.download_button(
                        "Baixar planilha atualizada", make_bytes(corrected_download_book),
                        file_name=f"{Path(uploaded.name).stem}_atualizada.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        width="stretch")
                    corrected_download_book.close()
                history_controls("roster", original=source_rows, label="os alunos")




                check_book = workbook_with_rows(content, effective_rows, effective_resources)
                _, _, raw_issues, blocking = inspect_entry(
                    check_book, effective_rows, choices, effective_resources)
                check_book.close()
                issues = [message for _, message in issue_records(raw_issues, effective_corrections, effective_rows)]
                if issues and blocking:
                    st.error("O processamento está bloqueado. Corrija os erros indicados nas tabelas e salve as alterações.")
                else:
                    if not issues:
                        st.success("Dados consistentes e capacidade suficiente para a configuração escolhida.")

                roster_signature = digest if committed_rows == source_rows and committed_resources == source_resources else hashlib.sha256(
                    repr((committed_rows, committed_resources)).encode("utf-8")).hexdigest()
                signature = (roster_signature, tuple(sorted(choices.items())))
                requested_generation = st.session_state.pop("requested_room_regeneration", False)
                generate = False
                if requested_generation:
                    generate = True
                elif "result" not in st.session_state:
                    with st.container(key="abc_run", width=400):
                        generate = st.button("Gerar ensalamento" if not blocking else "Gerar relatório de inconsistências",
                                             type="primary", width="stretch")
                if generate or requested_generation:
                    try:
                        with st.status("Preparando o resultado…", expanded=True) as progress:
                            st.write("Preparando a cópia da planilha com os dados confirmados.")
                            result_book = workbook_with_rows(content, effective_rows, effective_resources)
                            st.write("Validando os dados e gerando o relatório." if blocking else
                                     "Validando os dados e calculando a distribuição das turmas.")
                            _, _, model_issues, model_blocking = inspect_workbook(result_book, choices)
                            if blocking and not model_blocking:
                                custom_issues = [issue for issue in raw_issues if issue not in model_issues]
                                result_records = issue_records(custom_issues, effective_corrections, effective_rows)
                                raw_result = make_blocked_report(content, effective_rows, effective_resources,
                                                                 result_records)
                            else:
                                raw_result = make_result(result_book, choices)
                                combined_issues = [*raw_result[2],
                                                   *(issue for issue in raw_issues if issue not in raw_result[2])]
                                result_records = issue_records(combined_issues, effective_corrections, effective_rows)
                            output_bytes = rewrite_result_issues(raw_result[0], result_records)
                            result = (output_bytes, raw_result[1] or blocking, raw_result[2], raw_result[3])
                            result_book.close()
                            progress.update(label="Relatório de inconsistências gerado." if result[1] else
                                            "Ensalamento concluído.", state="complete", expanded=False)
                        st.session_state.result = result
                        st.session_state.result_signature = signature
                        st.session_state.result_room_mode = allow_empty
                        st.session_state.result_issue_records = result_records
                    except Exception as exc:
                        st.error(f"Não foi possível gerar o resultado: {exc}")
                section_result_records = st.session_state.get("result_issue_records", [])
                section_warning_count = sum(level == "Aviso" for level, _ in section_result_records)
                section_result = st.session_state.get("result")
                section_result_visible = section_result is not None and not draft_pending
                section_is_blocked = bool(section_result and section_result[1])
                if section_result_visible and section_result_records and (section_is_blocked or section_warning_count):
                    with st.expander("Relatório geral de erros e avisos", expanded=True):
                        for level, message in section_result_records:
                            (st.error if level == "Erro" else st.warning if level == "Aviso" else st.success)(message)
                if section_result_visible and section_warning_count:
                    st.warning(
                        f"Ensalamento concluído com {section_warning_count} aviso(s) que precisam de atenção. "
                        "A distribuição foi gerada; confira os registros antes de utilizá-la."
                    )
            committed_book.close()

            st.divider()
            if "result" not in st.session_state:
                st.caption("O resultado aparecerá abaixo após o processamento dos dados salvos.")
                st.stop()
            if draft_pending:
                st.warning("O resultado anterior está oculto porque há dados editados ainda não salvos. Salve ou descarte as edições antes de continuar.")
                st.stop()
            result_bytes, blocked, result_issues, stats = st.session_state.result
            result_records = st.session_state.get("result_issue_records", issue_records(result_issues))
            with st.container(key="abc_results", border=True, gap="medium"):
                st.header("3. Consultar e baixar o resultado")
                name = Path(uploaded.name).stem
                if blocked:
                    st.error("O ensalamento não foi gerado porque há erros bloqueantes. Corrija os dados na tabela e salve as alterações.")
                    st.download_button("Baixar planilha com inconsistências", result_bytes,
                                       file_name=f"{name}_inconsistencias.xlsx",
                                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                else:
                    if st.session_state.get("result_signature") != signature:
                        previous_mode = "somente as necessárias" if st.session_state.get("result_room_mode") else "todas as turmas disponíveis"
                        current_mode = st.session_state.room_mode_choice.casefold()
                        st.warning(f"O resultado exibido foi calculado usando {previous_mode}. A opção atual é {current_mode}; recalcule ao final da página para atualizar a distribuição.")
                    warning_count = sum(level == "Aviso" for level, _ in result_records)
                    corrections = [record for record in result_records if record[0] == "Correção automática"]
                    if not warning_count:
                        st.success("Ensalamento concluído normalmente, sem inconsistências.")
                    if corrections:
                        st.success(f"{len(corrections)} correção(ões) automática(s) aplicada(s). Consulte os valores originais e corrigidos na aba Inconsistências.")
                    total = sum(row[1] for row in stats)
                    capacity_by_room = {row[0]: resources.get("vagas_manha" if row[0].startswith("MANHÃ") else "vagas_tarde", 0)
                                        for row in stats}
                    export_room_table = pd.DataFrame(stats, columns=["Turma", "Alunos", "Meninas", "Meninos"])
                    export_denominator = export_room_table["Alunos"].replace(0, pd.NA)
                    export_room_table["% meninas"] = (100 * export_room_table["Meninas"] / export_denominator).fillna(0).round(1)
                    export_room_table["% meninos"] = (100 * export_room_table["Meninos"] / export_denominator).fillna(0).round(1)
                    room_table = pd.DataFrame(stats, columns=["Turma", "Alunos", "Meninas", "Meninos"])
                    room_table["Vagas"] = room_table["Turma"].map(capacity_by_room).fillna(0).astype(int)
                    room_table["Vagas restantes"] = (room_table["Vagas"] - room_table["Alunos"]).clip(lower=0)
                    room_denominator = room_table["Alunos"].replace(0, pd.NA)
                    room_table["% meninas"] = (100 * room_table["Meninas"] / room_denominator).fillna(0).round(1)
                    room_table["% meninos"] = (100 * room_table["Meninos"] / room_denominator).fillna(0).round(1)
                    room_table = room_table[["Turma", "Alunos", "Vagas", "Vagas restantes", "Meninas", "Meninos", "% meninas", "% meninos"]]

                    with st.container(key="abc_totals"):
                        metrics = st.columns(2)
                        metrics[0].metric("Alunos alocados", total)
                        metrics[1].metric("Turmas ocupadas", len(stats))
                        st.caption(f"Distribuição geral: {sum(row[2] for row in stats)} meninas e {sum(row[3] for row in stats)} meninos.")

                    st.subheader("Resumo por turma")
                    st.dataframe(room_table, hide_index=True, width="stretch")
                    with st.expander("Ver gráfico por turma", expanded=False):
                        st.bar_chart(room_table, x="Turma", y=["Meninas", "Meninos"], stack=True,
                                     color=["#2F9587", "#2B469A"], height=300, sort=False,
                                     alt="Quantidade de meninas e meninos em cada turma")

                    with st.container(key="abc_downloads"):
                        st.download_button("Baixar Excel completo (.xlsx)", result_bytes,
                                           file_name=f"{name}_resultado.xlsx",
                                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                           type="primary", width="stretch")
                        st.caption("O Excel inclui as abas de entrada, o ensalamento e as inconsistências.")
                        with st.expander("Outros formatos de download", expanded=False):
                            c1, c2 = st.columns(2)
                            c1.download_button("Baixar resumo por turma (.csv)",
                                               export_room_table.to_csv(index=False).encode("utf-8-sig"),
                                               file_name=f"{name}_resumo.csv", mime="text/csv", width="stretch")
                            original_pdf = report_pdf("Relatório resumo da distribuição", stats, capacity_by_room,
                                                      result_records, total, adjusted=False)
                            c2.download_button("Baixar resumo da distribuição (PDF)", original_pdf,
                                               file_name=f"{name}_relatorio_otimizado.pdf", mime="application/pdf",
                                               width="stretch")

                    st.divider()
                    st.subheader("Quer usar menos salas?")
                    st.write("A distribuição inicial usa todas as salas disponíveis. Se preferir, escolha outra opção abaixo e recalcule.")
                    all_rooms = "Todas as salas disponíveis"
                    minimum_rooms = "Somente as salas necessárias"
                    all_col, minimum_col = st.columns(2, gap="large")
                    with all_col.container(
                            key="abc_mode_all_selected" if st.session_state.room_mode_choice == all_rooms else "abc_mode_all",
                            border=True, gap="small"):
                        st.subheader("Todas as salas disponíveis")
                        st.caption("Mantém todas as salas disponíveis na distribuição.")
                        if st.button("Selecionado" if st.session_state.room_mode_choice == all_rooms else "Usar todas as salas",
                                     key="choose_all_rooms", type="primary" if st.session_state.room_mode_choice == all_rooms else "secondary",
                                     width="stretch", disabled=st.session_state.room_mode_choice == all_rooms):
                            st.session_state.room_mode_choice = all_rooms
                            st.rerun()
                    with minimum_col.container(
                            key="abc_mode_minimum_selected" if st.session_state.room_mode_choice == minimum_rooms else "abc_mode_minimum",
                            border=True, gap="small"):
                        st.subheader("Somente as salas necessárias")
                        st.caption("Usa o menor número de salas que comporta os alunos; a diferença entre turmas fica em até um aluno por turno.")
                        if st.button("Selecionado" if st.session_state.room_mode_choice == minimum_rooms else "Usar somente as salas necessárias",
                                     key="choose_minimum_rooms", type="primary" if st.session_state.room_mode_choice == minimum_rooms else "secondary",
                                     width="stretch", disabled=st.session_state.room_mode_choice == minimum_rooms):
                            st.session_state.room_mode_choice = minimum_rooms
                            st.rerun()

                    st.subheader("Prévia da configuração por turno")
                    room_cards = st.columns(2, gap="large")
                    for col, (period, suffix, card_key) in zip(
                            room_cards,
                            (("Manhã", "manha", "abc_room_morning"),
                             ("Tarde", "tarde", "abc_room_afternoon"))):
                        available = resources.get(f"turmas_{suffix}") or 0
                        capacity = resources.get(f"vagas_{suffix}") or 0
                        count = sum(row[2] == period for row in students)
                        selected_count = choices.get(suffix, 0)
                        with col.container(key=card_key, border=True, gap="small"):
                            if available and capacity:
                                seats = selected_count * capacity
                                st.metric(f"{period} · salas usadas", f"{selected_count} de {available}")
                                st.caption(f"{count} alunos · {capacity} vagas por turma · {max(seats - count, 0)} vagas livres no total")
                            else:
                                st.subheader(period)
                                st.warning("Confira a quantidade de turmas e as vagas por turma na planilha.")

                    if st.session_state.get("result_signature") != signature:
                        if st.button("Recalcular ensalamento", type="primary", key="recalculate_room_mode"):
                            st.session_state.requested_room_regeneration = True
                            st.rerun()


                    list_tab, adjustments_tab = st.tabs(
                        ["Lista de alunos", "Ajustes avançados"])
                    with list_tab:
                        result_table = output_table(result_bytes, "Ensalamento")
                        search_col, period_col = st.columns([2, 1])
                        search = search_col.text_input("Buscar aluno", placeholder="Digite o nome do aluno")
                        selected_period = period_col.selectbox("Turno", ["Todos", "Manhã", "Tarde"])
                        if selected_period != "Todos":
                            result_table = result_table[result_table["Disponibilidade"].astype(str).str.casefold() == selected_period.casefold()]
                        if search:
                            result_table = result_table[result_table["Nome"].astype(str).str.contains(
                                search, case=False, regex=False)]
                        if result_table.empty:
                            st.info("Nenhum aluno encontrado. Tente outro nome ou turno.")
                        else:
                            st.dataframe(result_table, hide_index=True, width="stretch", height=500)
                    with adjustments_tab:
                        st.subheader("Ajustes manuais da distribuição")
                        st.caption("Use esta área somente se precisar alterar a solução otimizada. As mudanças podem afetar o equilíbrio entre turmas.")
                        base_table = output_table(result_bytes, "Ensalamento")
                        original_assignments = {i: row[3] for i, row in enumerate(base_table.itertuples(index=False, name=None))}
                        adjustment_students = [tuple(row[:3]) for row in base_table.itertuples(index=False, name=None)]
                        actual_rooms = list(dict.fromkeys(room for room in original_assignments.values() if room))
                        capacities = {room: resources.get("vagas_manha" if room.startswith("MANHÃ") else "vagas_tarde", 0)
                                      for room in actual_rooms}
                        if st.session_state.get("manual_signature") != st.session_state.result_signature:
                            st.session_state.manual_original = dict(original_assignments)
                            st.session_state.manual_assignments = dict(original_assignments)
                            st.session_state.manual_history = []
                            st.session_state.manual_signature = st.session_state.result_signature
                        adjusted = st.session_state.manual_assignments
                        if st.session_state.get("manual_notice"):
                            st.success(st.session_state.pop("manual_notice"))
                        if changed(adjusted, original_assignments):
                            st.warning("Esta é uma distribuição ajustada manualmente; ela não é a solução otimizada original.")
                        manual_problems = validate_adjusted(adjusted, adjustment_students, capacities)
                        if manual_problems:
                            st.error("A distribuição ajustada não está conforme: " + " ".join(manual_problems))
                        manual_stats = room_stats(adjusted, adjustment_students, actual_rooms)
                        manual_table = pd.DataFrame(manual_stats, columns=["Turma", "Alunos", "Meninas", "Meninos"])
                        manual_table["Capacidade"] = manual_table["Turma"].map(capacities)
                        manual_table["% meninas"] = (100 * manual_table["Meninas"] / manual_table["Alunos"].replace(0, pd.NA)).fillna(0).round(1)
                        morning_col, afternoon_col = st.columns(2)
                        with morning_col:
                            st.subheader("Manhã")
                            st.table(manual_table[manual_table["Turma"].str.startswith("MANHÃ")])
                        with afternoon_col:
                            st.subheader("Tarde")
                            st.table(manual_table[manual_table["Turma"].str.startswith("TARDE")])
                        selected_room = st.selectbox("Consultar turma", actual_rooms, key="manual_view_room")
                        selected_members = [i for i, room in adjusted.items() if room == selected_room]
                        st.caption(f"{len(selected_members)} aluno(s) · capacidade máxima {capacities[selected_room]} · "
                                   f"{sum(adjustment_students[i][1] == 'M' for i in selected_members)} meninos · "
                                   f"{sum(adjustment_students[i][1] == 'F' for i in selected_members)} meninas")
                        st.table(pd.DataFrame(
                            [{"Registro": i + 1, "Aluno": adjustment_students[i][0], "Sexo": adjustment_students[i][1],
                              "Disponibilidade": adjustment_students[i][2]} for i in selected_members],
                            columns=["Registro", "Aluno", "Sexo", "Disponibilidade"]))
                        st.caption(f"Alunos sem turma: {sum(room is None for room in adjusted.values())}.")

                        def person_options(room):
                            return [i for i, assigned in adjusted.items() if assigned == room]

                        def person_label(i):
                            name_i, sex_i, period_i = adjustment_students[i]
                            return f"{name_i} · registro {i + 1} · {sex_i} · {period_i}"

                        with st.expander("Trocar alunos entre turmas", expanded=False):
                            room_a = st.selectbox("Turma de origem", actual_rooms, key="swap_room_a")
                            people_a = person_options(room_a)
                            student_a = st.selectbox("Primeiro aluno", people_a, format_func=person_label,
                                                     key="swap_student_a", disabled=not people_a)
                            room_b = st.selectbox("Outra turma", actual_rooms, key="swap_room_b")
                            people_b = person_options(room_b)
                            student_b = st.selectbox("Segundo aluno", people_b, format_func=person_label,
                                                     key="swap_student_b", disabled=not people_b)
                            cross_turn = room_a.split(" - ")[0] != room_b.split(" - ")[0]
                            if cross_turn:
                                st.error("A troca não é permitida entre turnos diferentes.")
                            elif room_a == room_b:
                                st.error("Selecione duas turmas diferentes para trocar alunos.")
                            elif people_a and people_b:
                                st.caption(f"Prévia: {person_label(student_a)} ↔ {person_label(student_b)}")
                                swap_preview = dict(adjusted)
                                swap_preview[student_a], swap_preview[student_b] = room_b, room_a
                                before = {row[0]: row[1:] for row in room_stats(adjusted, adjustment_students, [room_a, room_b])}
                                after = {row[0]: row[1:] for row in room_stats(swap_preview, adjustment_students, [room_a, room_b])}
                                st.table(pd.DataFrame([
                                    {"Turma": room, "Alunos antes": before[room][0], "Meninas antes": before[room][1],
                                     "Meninos antes": before[room][2], "Alunos depois": after[room][0],
                                     "Meninas depois": after[room][1], "Meninos depois": after[room][2]}
                                    for room in (room_a, room_b)]))
                            different_sex = bool(people_a and people_b and adjustment_students[student_a][1] != adjustment_students[student_b][1])
                            confirm_gender = st.checkbox("Esta troca poderá alterar o equilíbrio de gênero. Desejo continuar.",
                                                          key="confirm_swap_gender") if different_sex else True
                            if st.button("Confirmar troca", key="apply_swap", disabled=(cross_turn or room_a == room_b or
                                                                                         not people_a or not people_b or not confirm_gender)):
                                proposed = dict(adjusted)
                                proposed[student_a], proposed[student_b] = room_b, room_a
                                start_manual_action(proposed, "Troca de alunos aplicada. Confira os novos totais.")
                                st.rerun()

                        with st.expander("Mover aluno", expanded=False):
                            movable = [i for i, room in adjusted.items() if room is not None]
                            move_student = st.selectbox("Aluno", movable, format_func=person_label,
                                                        key="move_student", disabled=not movable)
                            move_room = st.selectbox("Turma de destino", actual_rooms, key="move_room")
                            move_valid = bool(movable) and move_room != adjusted.get(move_student)
                            if movable and move_valid:
                                source_room = adjusted[move_student]
                                same_turn = source_room.split(" - ")[0] == move_room.split(" - ")[0]
                                target_count = sum(room == move_room for room in adjusted.values())
                                has_space = target_count < capacities[move_room]
                                if not same_turn:
                                    st.error("A movimentação só pode ocorrer entre turmas do mesmo turno.")
                                if not has_space:
                                    st.error("A turma de destino está lotada.")
                                if same_turn and has_space:
                                    st.caption(f"Prévia: {person_label(move_student)} · {source_room} → {move_room}")
                                    st.warning("A mudança pode afetar o equilíbrio de quantidade ou gênero.")
                                    proposed_move = dict(adjusted)
                                    proposed_move[move_student] = move_room
                                    before = {row[0]: row[1:] for row in room_stats(adjusted, adjustment_students,
                                                                                  [source_room, move_room])}
                                    after = {row[0]: row[1:] for row in room_stats(proposed_move, adjustment_students,
                                                                                 [source_room, move_room])}
                                    st.table(pd.DataFrame([
                                        {"Turma": room, "Alunos antes": before[room][0], "Meninas antes": before[room][1],
                                         "Meninos antes": before[room][2], "Alunos depois": after[room][0],
                                         "Meninas depois": after[room][1], "Meninos depois": after[room][2]}
                                        for room in (source_room, move_room)]))
                                    confirm_move = st.checkbox("Confirmo esta movimentação manual.", key="confirm_move")
                                else:
                                    confirm_move = False
                            else:
                                same_turn, has_space, confirm_move = False, False, False
                            if st.button("Confirmar movimentação", key="apply_move",
                                         disabled=not (move_valid and same_turn and has_space and confirm_move)):
                                proposed = dict(adjusted)
                                proposed[move_student] = move_room
                                start_manual_action(proposed, "Movimentação aplicada. Confira os novos totais.")
                                st.rerun()

                        with st.expander("Remover aluno do ensalamento", expanded=False):
                            remove_room = st.selectbox("Turma", actual_rooms, key="remove_room")
                            removable = person_options(remove_room)
                            remove_student = st.selectbox("Aluno", removable, format_func=person_label,
                                                          key="remove_student", disabled=not removable)
                            st.warning("Você está prestes a remover este aluno da distribuição. Ele ficará sem turma no resultado ajustado. Deseja continuar?")
                            confirm_remove = st.checkbox("Confirmo que este aluno ficará sem turma.", key="confirm_remove")
                            if st.button("Remover aluno", key="apply_remove", disabled=not removable or not confirm_remove):
                                proposed = dict(adjusted)
                                proposed[remove_student] = None
                                start_manual_action(proposed, "Aluno removido apenas da distribuição ajustada.")
                                st.rerun()

                        undo_col, restore_col = st.columns(2)
                        if undo_col.button("Desfazer última alteração", disabled=not st.session_state.manual_history,
                                           key="undo_manual"):
                            st.session_state.manual_assignments = st.session_state.manual_history.pop()
                            st.session_state.manual_notice = "Última alteração desfeita."
                            st.rerun()
                        restore_confirm = restore_col.checkbox("Confirmo restaurar a solução original.", key="confirm_restore")
                        if restore_col.button("Restaurar ensalamento original", disabled=not changed(adjusted, original_assignments) or not restore_confirm,
                                              key="restore_manual"):
                            st.session_state.manual_assignments = dict(st.session_state.manual_original)
                            st.session_state.manual_history = []
                            st.session_state.manual_notice = "Solução otimizada original restaurada."
                            st.rerun()

                        adjusted_bytes = save_adjusted_result(result_bytes, adjusted, result_issues, manual_problems)
                        if manual_problems:
                            st.warning("A exportação será identificada como não conforme e incluirá os problemas na aba Inconsistências.")
                            adjusted_filename = f"{name}_ensalamento_ajustado_NAO_CONFORME.xlsx"
                        else:
                            adjusted_filename = f"{name}_ensalamento_ajustado.xlsx"
                        st.download_button("Baixar ensalamento ajustado", adjusted_bytes,
                                           file_name=adjusted_filename,
                                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                           disabled=any("excede a capacidade" in problem or "outro turno" in problem
                                                        for problem in manual_problems))
                        if changed(adjusted, original_assignments):
                            adjusted_records = [*result_records,
                                                *(('Aviso', f"Distribuição ajustada: {problem}")
                                                  for problem in manual_problems)]
                            adjusted_pdf = report_pdf("Relatório resumo da distribuição", manual_stats,
                                                      capacities, adjusted_records,
                                                      sum(row[1] for row in manual_stats), adjusted=True)
                            st.download_button("Baixar relatório da distribuição ajustada (PDF)", adjusted_pdf,
                                               file_name=f"{name}_relatorio_ajustado.pdf", mime="application/pdf")
