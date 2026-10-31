"""Ensalamento da ABC Aprendiz por programação inteira."""

from collections import defaultdict
from io import BytesIO
from itertools import islice
import os
from pathlib import Path
import tempfile
import unicodedata

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from pulp import LpInteger, LpMinimize, LpProblem, LpStatus, LpVariable, PULP_CBC_CMD, lpSum, value


RESOURCES = ("turmas_manha", "turmas_tarde", "vagas_manha", "vagas_tarde")
OPTIONAL_RESOURCES = ("turmas_usadas_manha", "turmas_usadas_tarde")


def plain(value):
    return "".join(c for c in unicodedata.normalize("NFKD", str(value).strip().casefold())
                   if not unicodedata.combining(c))


def letter(index):
    result = ""
    while index:
        index, digit = divmod(index - 1, 26)
        result = chr(65 + digit) + result
    return result


def preview_rows(sheet, limit=12):
    rows = list(islice(sheet.values, limit + 1))
    width = max(map(len, rows), default=0)
    if not rows:
        return [], []
    columns = [str(rows[0][i]) if i < len(rows[0]) and rows[0][i] is not None
               else f"Coluna {i + 1}" for i in range(width)]
    return columns, [list(row) + [None] * (width - len(row)) for row in rows[1:]]


def inspect_workbook(book, room_choices=None):
    issues, students, resources = [], [], {}
    blocking = False
    if not {"Alunos", "Turmas"} <= set(book.sheetnames):
        return students, resources, ["Arquivo: são necessárias as abas Alunos e Turmas."], True

    sheet = book["Alunos"]
    if [sheet.cell(1, col).value for col in (1, 2, 3)] != ["Nome", "Sexo", "Disponibilidade"]:
        issues.append("Alunos: cabeçalhos esperados: Nome, Sexo, Disponibilidade.")
        blocking = True
    names = defaultdict(list)
    for row in range(2, sheet.max_row + 1):
        name, sex, period = (sheet.cell(row, col).value for col in (1, 2, 3))
        if all(v is None or str(v).strip() == "" for v in (name, sex, period)):
            continue
        if plain(name) == "total" and sex is None and period is None:
            issues.append(f"Alunos, linha {row}: rodapé 'Total' ignorado.")
            continue
        errors = []
        if not isinstance(name, str) or not name.strip():
            errors.append("nome ausente ou inválido")
        canonical_sex = {"f": "F", "feminino": "F", "m": "M", "masculino": "M"}.get(plain(sex))
        if not canonical_sex:
            errors.append(f"sexo inválido ({sex!s})")
        canonical_period = {"manha": "Manhã", "tarde": "Tarde"}.get(plain(period))
        if not canonical_period:
            errors.append(f"disponibilidade inválida ({period!s})")
        if errors:
            issues.append(f"Alunos, linha {row}: " + "; ".join(errors) + ".")
            blocking = True
            continue
        if sex != canonical_sex or period != canonical_period or name != name.strip():
            issues.append(f"Alunos, linha {row}: texto padronizado para {canonical_sex}, {canonical_period}.")
        students.append((name.strip(), canonical_sex, canonical_period))
        names[plain(name)].append(row)
    for name, rows in names.items():
        if len(rows) > 1:
            issues.append(f"Alunos, linhas {', '.join(map(str, rows))}: nome duplicado; registros mantidos como pessoas distintas.")

    sheet = book["Turmas"]
    if [sheet.cell(1, col).value for col in (1, 2)] != ["Recursos", "Quantidade"]:
        issues.append("Turmas: cabeçalhos esperados: Recursos, Quantidade.")
        blocking = True
    for row in range(2, sheet.max_row + 1):
        key, amount = sheet.cell(row, 1).value, sheet.cell(row, 2).value
        if key is None and amount is None:
            continue
        if key in OPTIONAL_RESOURCES and (amount is None or str(amount).strip() == ""):
            continue
        if key not in RESOURCES + OPTIONAL_RESOURCES:
            issues.append(f"Turmas, linha {row}: recurso desconhecido ({key!s}).")
            blocking = True
        elif key in resources:
            issues.append(f"Turmas, linha {row}: recurso {key} repetido.")
            blocking = True
        elif isinstance(amount, bool) or not isinstance(amount, (int, float)) or not float(amount).is_integer() or amount <= 0:
            issues.append(f"Turmas, linha {row}: {key} precisa ser inteiro positivo.")
            blocking = True
        else:
            resources[key] = int(amount)
    for key in RESOURCES:
        if key not in resources:
            issues.append(f"Turmas: recurso válido {key} ausente.")
            blocking = True
    if room_choices:
        for suffix, used in room_choices.items():
            resources[f"turmas_usadas_{suffix}"] = used
    for period, suffix in (("Manhã", "manha"), ("Tarde", "tarde")):
        if f"turmas_{suffix}" in resources and f"vagas_{suffix}" in resources:
            count = sum(student[2] == period for student in students)
            available = resources[f"turmas_{suffix}"]
            used = resources.get(f"turmas_usadas_{suffix}", min(available, count))
            capacity = used * resources[f"vagas_{suffix}"]
            if used > available:
                issues.append(f"{period}: {used} turmas solicitadas, mas apenas {available} disponíveis.")
                blocking = True
            if used > count:
                issues.append(f"{period}: {used} turmas não vazias exigem pelo menos {used} alunos; há {count}.")
                blocking = True
            if count > capacity:
                issues.append(f"{period}: {count} alunos para apenas {capacity} vagas; faltam {count - capacity}.")
                blocking = True
            resources[f"turmas_usadas_{suffix}"] = used
    return students, resources, issues, blocking


def optimize(students, resources):
    assignments, stats = {}, []
    offset = 0
    for period, suffix in (("Manhã", "manha"), ("Tarde", "tarde")):
        group = [(i, sex) for i, (_, sex, shift) in enumerate(students) if shift == period]
        n = len(group)
        k, capacity = resources[f"turmas_usadas_{suffix}"], resources[f"vagas_{suffix}"]
        if k == 0:
            continue
        female = sum(sex == "F" for _, sex in group)
        model = LpProblem(f"Ensalamento_{suffix}", LpMinimize)
        sizes = [LpVariable(f"size_{j}", lowBound=n // k, upBound=min(capacity, (n + k - 1) // k), cat=LpInteger) for j in range(k)]
        girls = [LpVariable(f"girls_{j}", lowBound=0, cat=LpInteger) for j in range(k)]
        deviations = [LpVariable(f"deviation_{j}", lowBound=0, cat=LpInteger) for j in range(k)]
        worst = LpVariable("worst", lowBound=0, cat=LpInteger)
        model += lpSum(sizes) == n
        model += lpSum(girls) == female
        for j in range(k):
            model += girls[j] <= sizes[j]
            model += deviations[j] >= n * girls[j] - female * sizes[j]
            model += deviations[j] >= female * sizes[j] - n * girls[j]
            model += worst >= deviations[j]
        # Um ponto no pior desvio prevalece sobre qualquer soma dos demais desvios.
        model += (k * n * n + 1) * worst + lpSum(deviations)
        status = model.solve(PULP_CBC_CMD(msg=False))
        if LpStatus[status] != "Optimal":
            raise RuntimeError(f"Não foi encontrada solução ótima para o turno {period} ({LpStatus[status]}).")
        girls_queue = iter(i for i, sex in group if sex == "F")
        boys_queue = iter(i for i, sex in group if sex == "M")
        for j in range(k):
            size, girls_count = int(round(value(sizes[j]))), int(round(value(girls[j])))
            room = f"{period.upper()} - {letter(offset + j + 1)}"
            for _ in range(girls_count):
                assignments[next(girls_queue)] = room
            for _ in range(size - girls_count):
                assignments[next(boys_queue)] = room
            stats.append((room, size, girls_count, size - girls_count))
        offset += k
    assert len(assignments) == len(students)
    return assignments, stats


def format_sheet(sheet, widths):
    for cell in sheet[1]:
        cell.fill = PatternFill("solid", fgColor="17365D")
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(vertical="center")
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    for col, width in widths.items():
        sheet.column_dimensions[col].width = width


def make_result(book, room_choices=None):
    students, resources, issues, blocking = inspect_workbook(book, room_choices)
    if not blocking:
        assignments, stats = optimize(students, resources)
    else:
        assignments, stats = {}, []
    for title in ("Ensalamento", "Inconsistências"):
        if title in book:
            del book[title]
    sheet = book.create_sheet("Ensalamento")
    sheet.append(("Nome", "Sexo", "Disponibilidade", "Turma"))
    if not blocking:
        for i, (name, sex, period) in enumerate(students):
            sheet.append((name, sex, period, assignments[i]))
    format_sheet(sheet, {"A": 45, "B": 12, "C": 20, "D": 20})
    sheet = book.create_sheet("Inconsistências")
    sheet.append(("Situação", "Descrição"))
    for issue in issues:
        sheet.append(("Bloqueio" if blocking and ("inválid" in issue or "ausente" in issue or "precisa" in issue or "desconhecido" in issue or "repetido" in issue or "cabeçalhos" in issue or "faltam" in issue) else "Aviso", issue))
    if not issues:
        sheet.append(("OK", "Nenhuma inconsistência encontrada."))
    format_sheet(sheet, {"A": 16, "B": 100})
    output = BytesIO()
    book.save(output)
    return output.getvalue(), blocking, issues, stats


def process(source, destination):
    book = load_workbook(source)
    output, blocking, issues, stats = make_result(book)
    book.close()
    destination = Path(destination)
    with tempfile.NamedTemporaryFile(suffix=".xlsx", dir=destination.parent, delete=False) as temporary:
        staged = Path(temporary.name)
    try:
        staged.write_bytes(output)
        os.replace(staged, destination)
    finally:
        staged.unlink(missing_ok=True)
    return blocking, issues, stats
