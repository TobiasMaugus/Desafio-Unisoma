r"""Bateria independente do catálogo UniSoma; não altera a aplicação nem as entradas.

Uso: ..\.venv\Scripts\python.exe tests\rodar_instancias.py
Opções: --ids 01,13 --prioridades P0,P1 --extremos --timeout 60 --memory-mb 768
--output CAMINHO substitui o CSV padrão; --preservar-falhas DIR guarda apenas
saídas reprovadas. Todas as demais saídas e mensagens dos subprocessos são temporárias.

Os dois modos da interface são testados. A extração AST carrega as funções reais
de streamlit_app.py sem executar a página. O oráculo não usa a validação/modelo
da aplicação: conta ocorrências, audita células e resolve o objetivo de gênero
por intervalos inteiros e custos marginais convexos, sem chamar um solver.
"""
from __future__ import annotations

import argparse
import ast
from collections import Counter, defaultdict
import csv
import ctypes
from ctypes import wintypes
import hashlib
import heapq
from io import BytesIO
import json
import math
import os
from pathlib import Path
import re
import shutil
import signal
import subprocess
import sys
import tempfile
import time
import traceback
import unicodedata


PROJECT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT))
RESOURCES = ('turmas_manha', 'turmas_tarde', 'vagas_manha', 'vagas_tarde')
PERIODS = (('Manhã', 'manha'), ('Tarde', 'tarde'))
STATUSES = ('APROVADO', 'REPROVADO', 'ERRO_INESPERADO', 'TIMEOUT', 'NAO_EXECUTADO', 'INCONCLUSIVO')
FIELDS = ('id', 'arquivo', 'prioridade', 'categoria', 'resultado_esperado',
          'resultado_observado', 'status', 'tempo_segundos', 'alunos_lidos',
          'alunos_alocados', 'quantidade_inconsistencias', 'restricoes_violadas',
          'descricao_problema')
# Contratos positivos do catálogo: rejeitar estes casos não demonstra aprovação.
MUST_ALLOCATE = {1, 2, 3, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 20, 24, 35, 38}


def token(value):
    value = ' '.join(str(value if value is not None else '').strip().casefold().split())
    return ''.join(c for c in unicodedata.normalize('NFKD', value) if not unicodedata.combining(c))


def fingerprint(path):
    digest = hashlib.sha256()
    with path.open('rb') as handle:
        for block in iter(lambda: handle.read(65536), b''):
            digest.update(block)
    return digest.hexdigest()


def application_functions():
    """Minimal read-only adapter: execute definitions/imports, never Streamlit UI."""
    path = PROJECT / 'streamlit_app.py'
    source = ast.parse(path.read_text(encoding='utf-8'), filename=str(path))
    nodes = [node for node in source.body if isinstance(node, (ast.Import, ast.ImportFrom, ast.FunctionDef))
             or (isinstance(node, ast.Assign) and any(isinstance(target, ast.Name)
                 and target.id in ('EQUIVALENCES', 'RESOURCE_EQUIVALENCES') for target in node.targets))]
    namespace = {'__file__': str(path), '__name__': '_unisoma_test_adapter'}
    exec(compile(ast.Module(body=nodes, type_ignores=[]), str(path), 'exec'), namespace)
    return namespace


def oracle_input(book, cached):
    """Read the contract independently, retaining each Excel registration occurrence."""
    result = {'students': [], 'resources': {}, 'faults': [], 'changes': [], 'duplicates': [],
              'read': 0, 'student_positions': {}, 'resource_positions': {}, 'flexible': False}

    def fault(area, line, field, detail):
        result['faults'].append((area, line, field, detail))

    for area in ('Alunos', 'Turmas'):
        if area not in book:
            fault(area, None, 'aba', f'Aba {area} ausente')
    if 'Alunos' in book:
        sheet, values = book['Alunos'], cached['Alunos']
        if tuple(sheet.cell(1, c).value for c in (1, 2, 3)) != ('Nome', 'Sexo', 'Disponibilidade'):
            fault('Alunos', None, 'cabeçalho', 'Cabeçalho de Alunos incompatível')
        sex_map = {word: sex for sex, words in (
            ('F', ('f', 'feminino', 'feminina', 'mulher', 'garota', 'menina')),
            ('M', ('m', 'masculino', 'masculina', 'homem', 'garoto', 'menino')))
            for word in words}
        period_map = {word: period for period, words in (
            ('Manhã', ('manha', 'matutino', 'matutina', 'periodo da manha', 'periodo manha')),
            ('Tarde', ('tarde', 'vespertino', 'vespertina', 'periodo da tarde', 'periodo tarde')))
            for word in words}
        compact = 2
        for line in range(2, sheet.max_row + 1):
            cells = [sheet.cell(line, c) for c in (1, 2, 3)]
            raw = [cell.value for cell in cells]
            read_values = [values.cell(line, c).value for c in (1, 2, 3)]
            if any(value not in (None, '') for value in read_values):
                result['student_positions'][line] = compact
                compact += 1
            if all(value is None or str(value).strip() == '' for value in raw):
                continue
            name, sex, period = raw
            if token(name) == 'total' and sex is None and period is None:
                continue
            result['read'] += 1
            canonical_sex, canonical_period = sex_map.get(token(sex)), period_map.get(token(period))
            good = True
            for cell, field in zip(cells, ('nome', 'sexo', 'disponibilidade')):
                if cell.data_type == 'f':
                    fault('Alunos', line, field, f'Alunos!{cell.coordinate}: fórmula em {field}')
                    good = False
            if not isinstance(name, str) or not name.strip():
                fault('Alunos', line, 'nome', f'Alunos!A{line}: nome ausente/inválido')
                good = False
            if isinstance(name, str) and any(unicodedata.category(c) == 'Cf' for c in name):
                fault('Alunos', line, 'nome', f'Alunos!A{line}: caracteres invisíveis no nome')
                good = False
            if not canonical_sex:
                fault('Alunos', line, 'sexo', f'Alunos!B{line}: sexo inválido {sex!r}')
                good = False
            if not canonical_period:
                fault('Alunos', line, 'disponibilidade', f'Alunos!C{line}: turno inválido {period!r}')
                good = False
            for field, old, new in (('sexo', sex, canonical_sex), ('disponibilidade', period, canonical_period)):
                if new is not None and old != new:
                    result['changes'].append((line, field, old, new))
            if good:
                result['students'].append((line, name.strip(), canonical_sex, canonical_period))
        grouped = defaultdict(list)
        for line, name, _, _ in result['students']:
            grouped[token(name)].append(line)
        result['duplicates'] = [lines for lines in grouped.values() if len(lines) > 1]
    if 'Turmas' in book:
        sheet = book['Turmas']
        if tuple(sheet.cell(1, c).value for c in (1, 2)) != ('Recursos', 'Quantidade'):
            fault('Turmas', None, 'cabeçalho', 'Cabeçalho de Turmas incompatível')
        seen, compact = set(), 2
        for line in range(2, sheet.max_row + 1):
            key, amount = sheet.cell(line, 1).value, sheet.cell(line, 2).value
            if key is None and amount is None:
                continue
            result['resource_positions'][line] = compact
            compact += 1
            key = token(key)
            if key not in RESOURCES + ('turmas_usadas_manha', 'turmas_usadas_tarde'):
                fault('Turmas', line, str(key), f'Turmas!A{line}: recurso desconhecido {key!r}')
                continue
            if key in seen:
                fault('Turmas', line, key, f'Turmas!A{line}: recurso duplicado {key}')
            seen.add(key)
            if key.startswith('turmas_usadas_') and (amount is None or str(amount).strip() == ''):
                continue
            # Numeric formulas with a valid cached integer are compatible input;
            # missing caches and Excel errors must still be diagnosed. Never evaluate.
            candidate = cached['Turmas'].cell(line, 2).value if sheet.cell(line, 2).data_type == 'f' else amount
            number = None
            if isinstance(candidate, (int, float)) and not isinstance(candidate, bool):
                if math.isfinite(candidate) and candidate == int(candidate):
                    number = int(candidate)
            elif isinstance(candidate, str) and re.fullmatch(r'\s*\d+\s*', candidate):
                number = int(candidate)
                result['flexible'] = True
            empty_period = (key in ('turmas_manha', 'turmas_tarde')
                            and not any(student[3] == ('Manhã' if key == 'turmas_manha' else 'Tarde')
                                        for student in result['students']))
            if number is None or number < 0 or (number == 0 and not empty_period):
                fault('Turmas', line, key, f'Turmas!B{line}: {key} não é inteiro positivo ({amount!r})')
            else:
                result['resources'].setdefault(key, number)
        for key in RESOURCES:
            if key not in seen:
                fault('Turmas', None, key, f'Recurso obrigatório ausente: {key}')
    return result


def alphabet(index):
    """Independent zero-based spreadsheet suffix, including AA, AB, ..."""
    text = ''
    while index >= 0:
        text = chr(65 + index % 26) + text
        index = index // 26 - 1
    return text


def gender_optimum(n, female, k):
    """Exact independent lexicographic (max deviation, sum deviations) optimum.

    Balanced sizes have exactly n % k rooms of size ceil(n/k). For any maximum
    deviation D, each room permits an integer interval of female counts. Their
    sum is attainable iff female lies between the sums of interval endpoints.
    Binary search obtains D; a heap minimizes the remaining convex absolute costs.
    """
    sizes = [n // k + (index < n % k) for index in range(k)]

    def intervals(bound):
        return [(max(0, -((bound - female * size) // n)),
                 min(size, (female * size + bound) // n)) for size in sizes]

    low, high = -1, n * n
    while high - low > 1:
        middle = (low + high) // 2
        bounds = intervals(middle)
        if all(a <= b for a, b in bounds) and sum(a for a, _ in bounds) <= female <= sum(b for _, b in bounds):
            high = middle
        else:
            low = middle
    bounds = intervals(high)
    counts = [a for a, _ in bounds]

    def cost(index, amount):
        return abs(n * amount - female * sizes[index])

    heap = []
    for index, (a, b) in enumerate(bounds):
        if a < b:
            heapq.heappush(heap, (cost(index, a + 1) - cost(index, a), index))
    for _ in range(female - sum(counts)):
        _, index = heapq.heappop(heap)
        counts[index] += 1
        if counts[index] < bounds[index][1]:
            amount = counts[index]
            heapq.heappush(heap, (cost(index, amount + 1) - cost(index, amount), index))
    return high, sum(cost(index, amount) for index, amount in enumerate(counts))


def app_run(functions, content, mode):
    """Same normalization, validation and Excel export order used by the page."""
    from openpyxl import load_workbook
    book = functions['load_input_workbook'](content)
    rows, corrections = functions['normalize_roster'](functions['student_rows'](book))
    resources, resource_corrections = functions['normalize_resources'](functions['resource_rows'](book))
    book.close()
    working = functions['workbook_with_rows'](content, rows, resources)
    students, capacities, _, _ = functions['inspect_entry'](working, rows, resource_records=resources)
    choices = functions['room_choices_for_mode'](students, capacities, mode == 'necessarias')
    _, _, issues, blocking = functions['inspect_entry'](working, rows, choices, resources)
    _, _, model_issues, model_blocking = functions['inspect_workbook'](working, choices)
    if blocking and not model_blocking:
        extra = [issue for issue in issues if issue not in model_issues]
        records = functions['issue_records'](extra, [*corrections, *resource_corrections], rows)
        raw = functions['make_blocked_report'](content, rows, resources, records)
    else:
        raw = functions['make_result'](working, choices)
        combined = [*raw[2], *(issue for issue in issues if issue not in raw[2])]
        records = functions['issue_records'](combined, [*corrections, *resource_corrections], rows)
    output = functions['rewrite_result_issues'](raw[0], records)
    working.close()
    return {'output': output, 'blocking': raw[1] or blocking, 'records': records, 'stats': raw[3],
            'choices': choices, 'valid_count': len(students)}


def audit_output(content, original, oracle, observed, mode, test_id):
    from openpyxl import load_workbook
    book = load_workbook(BytesIO(content), data_only=False)
    violations, uncertain = [], []
    evidence = []
    required = {'Ensalamento': ('Nome', 'Sexo', 'Disponibilidade', 'Turma'),
                'Inconsistências': ('Situação', 'Descrição')}
    if not observed['blocking']:
        required.update(Alunos=('Nome', 'Sexo', 'Disponibilidade'), Turmas=('Recursos', 'Quantidade'))
    for area, headers in required.items():
        if area not in book:
            violations.append(f'Excel: aba {area} ausente')
        elif tuple(book[area].cell(1, c).value for c in range(1, len(headers) + 1)) != headers:
            violations.append(f'Excel: colunas obrigatórias incorretas em {area}')
    for area in original.sheetnames:
        if area not in book:
            violations.append(f'Excel: aba de entrada {area} removida')
    if any(area not in book for area in required):
        book.close()
        return violations, uncertain, 0, 0, ['Excel sem estrutura auditável']
    allocations = [tuple(row[:4]) for row in book['Ensalamento'].iter_rows(min_row=2, values_only=True)
                   if any(value is not None for value in row[:4])]
    diagnostics = [(str(row[0]), str(row[1])) for row in book['Inconsistências'].iter_rows(min_row=2, values_only=True)
                   if row[0] is not None and row[0] != 'OK']
    messages = [token(message) for _, message in diagnostics]
    joined = ' | '.join(messages)
    if len(diagnostics) != len(observed['records']):
        violations.append('Inconsistências: mensagens exibidas e exportadas divergem')
    for level, message in observed['records']:
        if (str(level), str(message)) not in diagnostics:
            violations.append(f'Inconsistência perdida na exportação: {message}')

    # Every independently detected bad field must have a specific diagnostic.
    for area, line, field, detail in oracle['faults']:
        positions = oracle['student_positions'] if area == 'Alunos' else oracle['resource_positions']
        position = positions.get(line)
        relevant = messages
        if line is not None:
            relevant = []
            for message in messages:
                match = re.match(token(area) + r', linhas? ([\d, ]+):', message)
                if match and position in [int(v) for v in re.findall(r'\d+', match.group(1))]:
                    relevant.append(message)
        terms = ('cabecalho', 'colunas', 'coluna', 'titulos') if field == 'cabeçalho' else (token(field),)
        if not any(token(area) in message and any(term in message for term in terms) for message in relevant):
            violations.append(f'Diagnóstico ausente: {detail}')

    for lines in oracle['duplicates']:
        positions = {oracle['student_positions'].get(line) for line in lines}
        matching = []
        for message in messages:
            match = re.match(r'alunos, linhas? ([\d, ]+):', message)
            if match and positions <= {int(v) for v in re.findall(r'\d+', match.group(1))}:
                matching.append(message)
        if not any('nome' in message and any(term in message for term in ('repetid', 'duplicad', 'mais de uma vez'))
                   for message in matching):
            violations.append(f'Duplicidade não informada nas matrículas {lines}')
    for line, field, old, new in oracle['changes']:
        position = oracle['student_positions'].get(line)
        if not any(f'linha {position}:' in message and token(field) in message
                   and token(old) in message and token(new) in message for message in messages):
            violations.append(f'Normalização não documentada: matrícula {line}, {field}, {old!r} -> {new!r}')

    independent_rows = [tuple(row[1:]) for row in oracle['students']]
    resources = oracle['resources']
    expected_rooms, period_rooms = [], {}
    feasible = not oracle['faults'] and bool(independent_rows) and all(key in resources for key in RESOURCES)
    offset = 0
    for period, suffix in PERIODS:
        n = sum(row[2] == period for row in independent_rows)
        available, capacity = resources.get('turmas_' + suffix), resources.get('vagas_' + suffix)
        if available is None or capacity is None:
            feasible = False
            continue
        k = available if mode == 'todas' else min(available, (n + capacity - 1) // capacity)
        if k > n or n > k * capacity:
            feasible = False
        # An infeasible request must not make the independent auditor allocate
        # millions of room names either (notably catalog case 48).
        names = [f'{period.upper()} - {alphabet(offset + index)}' for index in range(k)] if k <= n else []
        period_rooms[period] = names
        expected_rooms.extend(names)
        offset += k

    if observed['blocking']:
        if allocations or observed['stats']:
            violations.append('GRAVE: ensalamento produzido apesar de bloqueio')
        if not diagnostics:
            violations.append('Arquivo bloqueado sem mensagem compreensível')
        if test_id in MUST_ALLOCATE and (feasible or test_id == 3):
            violations.append('Catálogo exige ensalamento, mas a aplicação rejeitou a entrada')
        if not independent_rows and 'Alunos' in original and not any('aluno' in msg for msg in messages):
            violations.append('Ausência de alunos não informada')
        evidence.append(f'bloqueado com {len(diagnostics)} mensagem(ns); sem alocações')
    else:
        if oracle['faults']:
            violations.append('GRAVE: entrada com erros bloqueantes aceita: ' + '; '.join(f[3] for f in oracle['faults']))
        if not feasible:
            violations.append('GRAVE: ensalamento aceito para configuração independentemente inviável')
        actual_rows = [tuple(row[:3]) for row in allocations]
        missing, extra = Counter(independent_rows) - Counter(actual_rows), Counter(actual_rows) - Counter(independent_rows)
        if missing or extra:
            violations.append(f'GRAVE: matrículas perdidas/alteradas={sum(missing.values())}; extras/duplicadas={sum(extra.values())}')
        elif actual_rows != independent_rows:
            uncertain.append('Ordem das matrículas mudou; multiplicidade conferida, identidade individual não rastreável')
        groups = defaultdict(list)
        for name, sex, period, room in allocations:
            if not isinstance(room, str) or not room.startswith(str(period).upper() + ' - '):
                violations.append(f'GRAVE: aluno em turma de outro turno/sem turma: {name!r}, {room!r}')
            groups[room].append((name, sex, period))
        if set(groups) != set(expected_rooms):
            violations.append(f'Turmas vazias, extras ou sequência incorreta: esperadas {expected_rooms}; encontradas {list(groups)}')
        for period, suffix in PERIODS:
            names = period_rooms.get(period, [])
            n = sum(row[2] == period for row in independent_rows)
            female = sum(row[1] == 'F' and row[2] == period for row in independent_rows)
            capacity = resources.get('vagas_' + suffix, 0)
            counts = [len(groups.get(room, [])) for room in names]
            for room, count in zip(names, counts):
                if count > capacity:
                    violations.append(f'GRAVE: capacidade violada em {room}: {count}/{capacity}')
            if names and n > 0:
                if any(not n // len(names) <= count <= (n + len(names) - 1) // len(names) for count in counts):
                    violations.append(f'Desequilíbrio de ocupação em {period}: {counts}')
                deviations = [abs(n * sum(person[1] == 'F' for person in groups.get(room, [])) - female * len(groups.get(room, [])))
                              for room in names]
                actual = (max(deviations, default=0), sum(deviations))
                optimum = gender_optimum(n, female, len(names))
                if actual != optimum:
                    violations.append(f'Objetivo de gênero incorreto em {period}: {actual}; ótimo independente {optimum}')
                evidence.append(f'{period}: {n} alunos/{len(names)} turmas; gênero {actual} = ótimo {optimum}')
        independent_stats = Counter((room, len(members), sum(s == 'F' for _, s, _ in members), sum(s == 'M' for _, s, _ in members))
                                    for room, members in groups.items())
        if independent_stats != Counter(tuple(row) for row in observed['stats']):
            violations.append('Estatísticas por turma divergem dos alunos efetivamente exportados')
        for area, width in (('Alunos', 3), ('Turmas', 2), ('Ensalamento', 4)):
            if area in book:
                for row in book[area].iter_rows(min_row=2, max_col=width):
                    for cell in row:
                        if cell.data_type == 'f':
                            violations.append(f'GRAVE: fórmula reexportada em {area}!{cell.coordinate}')
    book.close()
    return list(dict.fromkeys(violations)), uncertain, len(allocations), len(diagnostics), evidence


def audit_pdf(data):
    """Independent PDF reader: structure, footer count and printable text bounds."""
    if not data.startswith(b'%PDF-') or not data.rstrip().endswith(b'%%EOF'):
        return ['PDF exportado inválido'], []
    try:
        from pypdf import PdfReader
    except ImportError:
        return [], ['PDF não auditado: pypdf indisponível; Excel foi conferido independentemente']
    try:
        reader = PdfReader(BytesIO(data), strict=True)
        if not reader.pages:
            return ['PDF exportado sem páginas'], []
        violations = []
        for index, page in enumerate(reader.pages, 1):
            text = page.extract_text()
            if f'Página {index} de {len(reader.pages)}' not in text:
                violations.append(f'PDF: numeração ausente/incorreta na página {index}')
            y = 0
            for arguments, operator in page.get_contents().operations:
                if operator == b'Tm':
                    y = float(arguments[5])
                elif operator == b'Tj':
                    value = str(arguments[0])
                    footer = value.startswith(('Página ', 'ABC Aprendiz · Organização das turmas'))
                    if value.strip() and not footer and not 62 <= y <= 822:
                        violations.append(f'PDF: conteúdo invade rodapé ou sai da página {index} '
                                          f'(y={y:.1f} pontos): {value[:70]!r}')
                        break
        return violations, []
    except Exception as exc:
        return [f'PDF ilegível no leitor independente: {type(exc).__name__}: {exc}'], []


def worker(request_path):
    # Do not import application/dependencies or read workbooks before containment.
    if sys.stdin.readline().strip() != 'GO':
        return 2
    request = json.loads(Path(request_path).read_text(encoding='utf-8'))
    directory = Path(request_path).parent
    if os.name != 'nt':
        import resource
        memory = request['memory_mb'] * 1024 * 1024
        resource.setrlimit(resource.RLIMIT_AS, (memory, memory))
        resource.setrlimit(resource.RLIMIT_CPU, (math.ceil(request['timeout']), math.ceil(request['timeout']) + 1))
    stage = 'dependências'
    result = {'status': 'INCONCLUSIVO', 'observed': '', 'read': '', 'allocated': '', 'issues': '', 'violations': [], 'problem': ''}
    try:
        from openpyxl import load_workbook
        functions = application_functions()
        content = Path(request['source']).read_bytes()
        stage = 'leitura'
        try:
            cached = load_workbook(BytesIO(content), data_only=True)
        except MemoryError:
            raise
        except Exception as exc:
            # Exact boundary handled by the application's upload try/except.
            message = f'Não foi possível abrir a planilha. Verifique se é um .xlsx válido: {exc}'
            invalid_file = 'arquivo invalido' in token(request['case']['categoria'])
            result.update(status='APROVADO' if invalid_file else 'REPROVADO',
                          observed=f'Rejeição no upload: {type(exc).__name__}: {message}', issues=1,
                          problem='' if invalid_file else 'Arquivo esperado como legível foi rejeitado no upload')
        else:
            stage = 'oráculo de entrada'
            original = load_workbook(BytesIO(content), data_only=False)
            oracle = oracle_input(original, cached)
            cached.close()
            result['read'] = oracle['read']
            findings, doubts, observations, allocated, issues = [], [], [], [], set()
            for mode in ('todas', 'necessarias'):
                stage = f'processamento/{mode}'
                observed = app_run(functions, content, mode)
                path = directory / f'{mode}.xlsx'
                path.write_bytes(observed['output'])
                stage = f'auditoria/{mode}'
                bad, unsure, count, issue_count, evidence = audit_output(
                    observed['output'], original, oracle, observed, mode, int(request['case']['id']))
                findings.extend(f'{mode}: {item}' for item in bad)
                doubts.extend(f'{mode}: {item}' for item in unsure)
                observations.append(f'{mode}: {count} alocados; {issue_count} inconsistências; ' + '; '.join(evidence))
                allocated.append(count)
                issues.update(tuple(record) for record in observed['records'])
                if not observed['blocking']:
                    # Exercise the PDF export actually evaluated by the page too.
                    stage = f'exportação PDF/{mode}'
                    output_book = load_workbook(BytesIO(observed['output']), data_only=True)
                    roster = list(output_book['Ensalamento'].iter_rows(min_row=2, max_col=4, values_only=True))
                    capacities = {room: oracle['resources'].get('vagas_manha' if room.startswith('MANHÃ') else 'vagas_tarde', 0)
                                  for room, *_ in observed['stats']}
                    pdf = functions['report_pdf']('Relatório resumo da distribuição', observed['stats'], capacities,
                                                  observed['records'], count, roster=roster)
                    stage = f'auditoria PDF/{mode}'
                    pdf_bad, pdf_unknown = audit_pdf(pdf)
                    findings.extend(f'{mode}: {item}' for item in pdf_bad)
                    doubts.extend(f'{mode}: {item}' for item in pdf_unknown)
                    output_book.close()
            original.close()
            result.update(status='REPROVADO' if findings else 'INCONCLUSIVO' if doubts else 'APROVADO',
                          observed=' | '.join(observations), allocated=max(allocated, default=0), issues=len(issues),
                          violations=findings, problem=' | '.join(findings + doubts))
    except (ImportError, ModuleNotFoundError) as exc:
        result.update(status='NAO_EXECUTADO', problem=f'{stage}: dependência indisponível: {exc}')
    except Exception as exc:
        status = 'INCONCLUSIVO' if stage.startswith(('auditoria', 'oráculo')) else 'ERRO_INESPERADO'
        result.update(status=status, observed=f'{stage}: {type(exc).__name__}: {exc}',
                      problem=f'{stage}: {type(exc).__name__}: {exc}; ' + traceback.format_exc(limit=5))
    (directory / 'response.json').write_text(json.dumps(result, ensure_ascii=False), encoding='utf-8')
    return 0


class WindowsJob:
    """Hard memory cap for the whole tree; closing the job kills CBC descendants."""
    def __init__(self, memory_mb):
        class Basic(ctypes.Structure):
            _fields_ = [('ProcessTime', ctypes.c_int64), ('JobTime', ctypes.c_int64),
                        ('Flags', wintypes.DWORD), ('MinWorkingSet', ctypes.c_size_t),
                        ('MaxWorkingSet', ctypes.c_size_t), ('ActiveProcessLimit', wintypes.DWORD),
                        ('Affinity', ctypes.c_size_t), ('PriorityClass', wintypes.DWORD),
                        ('SchedulingClass', wintypes.DWORD)]

        class IO(ctypes.Structure):
            _fields_ = [(name, ctypes.c_uint64) for name in ('ReadOp', 'WriteOp', 'OtherOp', 'ReadBytes', 'WriteBytes', 'OtherBytes')]

        class Extended(ctypes.Structure):
            _fields_ = [('Basic', Basic), ('IO', IO), ('ProcessMemory', ctypes.c_size_t),
                        ('JobMemory', ctypes.c_size_t), ('PeakProcessMemory', ctypes.c_size_t),
                        ('PeakJobMemory', ctypes.c_size_t)]

        self.api = ctypes.WinDLL('kernel32', use_last_error=True)
        self.api.CreateJobObjectW.argtypes = [ctypes.c_void_p, wintypes.LPCWSTR]
        self.api.CreateJobObjectW.restype = wintypes.HANDLE
        self.api.SetInformationJobObject.argtypes = [wintypes.HANDLE, ctypes.c_int, ctypes.c_void_p, wintypes.DWORD]
        self.api.AssignProcessToJobObject.argtypes = [wintypes.HANDLE, wintypes.HANDLE]
        self.api.CloseHandle.argtypes = [wintypes.HANDLE]
        self.handle = self.api.CreateJobObjectW(None, None)
        if not self.handle:
            raise ctypes.WinError(ctypes.get_last_error())
        info = Extended()
        info.Basic.Flags = 0x2000 | 0x200  # KILL_ON_JOB_CLOSE | JOB_MEMORY
        info.JobMemory = memory_mb * 1024 * 1024
        if not self.api.SetInformationJobObject(self.handle, 9, ctypes.byref(info), ctypes.sizeof(info)):
            error = ctypes.get_last_error()
            self.close()
            raise ctypes.WinError(error)

    def assign(self, process):
        if not self.api.AssignProcessToJobObject(self.handle, wintypes.HANDLE(int(process._handle))):
            raise ctypes.WinError(ctypes.get_last_error())

    def close(self):
        if self.handle:
            self.api.CloseHandle(self.handle)
            self.handle = None


def isolated_case(case, source, options, parent_directory):
    row = dict(zip(FIELDS, ('',) * len(FIELDS)))
    row.update(id=case['id'], arquivo=case['arquivo'], prioridade=case['prioridade'],
               categoria=case['categoria'], resultado_esperado=case['expectativa'], status='NAO_EXECUTADO', tempo_segundos='0.000')
    if source is None:
        row['descricao_problema'] = 'Arquivo ausente ou nome ambíguo na busca recursiva'
        return row
    if int(case['id']) == 48 and not options.extremos:
        row['descricao_problema'] = 'Caso extremo com 999.999 turmas não executado por segurança; habilitar com --extremos'
        return row
    before = None
    start = time.perf_counter()
    job, process = None, None
    directory = parent_directory / case['id']
    directory.mkdir()
    try:
        before = fingerprint(source)
        if os.name == 'nt':
            job = WindowsJob(options.memory_mb)
        elif sys.platform != 'linux':
            raise OSError('Limites seguros de memória/árvore de processos não implementados nesta plataforma')
        request = {'case': case, 'source': str(source), 'memory_mb': options.memory_mb, 'timeout': options.timeout}
        request_path = directory / 'request.json'
        request_path.write_text(json.dumps(request, ensure_ascii=False), encoding='utf-8')
        environment = dict(os.environ, PYTHONDONTWRITEBYTECODE='1', PYTHONIOENCODING='utf-8',
                           TEMP=str(directory), TMP=str(directory), TMPDIR=str(directory))
        with (directory / 'worker.txt').open('wb') as messages:
            process = subprocess.Popen(
                [sys.executable, '-B', str(Path(__file__).resolve()), '--worker', str(request_path)],
                stdin=subprocess.PIPE, stdout=messages, stderr=subprocess.STDOUT, env=environment,
                cwd=PROJECT, creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0,
                start_new_session=os.name != 'nt')
            if job:
                job.assign(process)
            process.stdin.write(b'GO\n')
            process.stdin.close()
            try:
                process.wait(timeout=options.timeout)
            except subprocess.TimeoutExpired:
                row.update(status='TIMEOUT', descricao_problema=f'Limite de {options.timeout:g}s excedido; árvore inteira encerrada')
            else:
                response_path = directory / 'response.json'
                if response_path.exists():
                    response = json.loads(response_path.read_text(encoding='utf-8'))
                    row.update(status=response['status'], resultado_observado=response['observed'],
                               alunos_lidos=response['read'], alunos_alocados=response['allocated'],
                               quantidade_inconsistencias=response['issues'], restricoes_violadas=' | '.join(response['violations']),
                               descricao_problema=response['problem'])
                else:
                    diagnostic = (directory / 'worker.txt').read_text(encoding='utf-8', errors='replace')[-2000:]
                    row.update(status='ERRO_INESPERADO', descricao_problema=
                               f'Processo terminou sem resposta (código {process.returncode}); limite de memória {options.memory_mb} MiB. {diagnostic}')
    except (OSError, ValueError) as exc:
        row.update(status='NAO_EXECUTADO', descricao_problema=f'Não foi possível garantir execução isolada segura: {exc}')
    except Exception as exc:
        row.update(status='ERRO_INESPERADO', descricao_problema=f'Falha do executor isolado: {type(exc).__name__}: {exc}')
    finally:
        if job:
            job.close()
        elif process and os.name != 'nt':
            try:
                os.killpg(process.pid, signal.SIGKILL)
            except ProcessLookupError:
                pass
        if process:
            if process.poll() is None:
                process.kill()
            process.wait(timeout=10)
            if process.stdin and not process.stdin.closed:
                process.stdin.close()
        row['tempo_segundos'] = f'{time.perf_counter() - start:.3f}'
    try:
        if before is not None and fingerprint(source) != before:
            row.update(status='REPROVADO', descricao_problema='GRAVE: hash da instância original foi alterado')
    except OSError as exc:
        row.update(status='INCONCLUSIVO', descricao_problema=f'Não foi possível conferir integridade final da entrada: {exc}')
    if options.preservar_falhas and row['status'] in ('REPROVADO', 'ERRO_INESPERADO', 'TIMEOUT'):
        destination = options.preservar_falhas / case['id']
        outputs = list(directory.glob('*.xlsx'))
        if outputs:
            destination.mkdir(parents=True, exist_ok=True)
            for output in outputs:
                shutil.copy2(output, destination / output.name)
    return row


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument('--instancias', type=Path, default=PROJECT.parent / 'instancias')
    parser.add_argument('--catalogo', type=Path)
    parser.add_argument('--output', type=Path)
    parser.add_argument('--ids', help='IDs separados por vírgula, por exemplo 01,13,51')
    parser.add_argument('--prioridades', help='P0,P1 ou P2, separados por vírgula')
    parser.add_argument('--extremos', action='store_true', help='Executar também o teste 48, sempre com limites rígidos')
    parser.add_argument('--timeout', type=float, default=60)
    parser.add_argument('--memory-mb', type=int, default=768)
    parser.add_argument('--preservar-falhas', type=Path, metavar='DIR')
    parser.add_argument('--worker', type=Path, help=argparse.SUPPRESS)
    options = parser.parse_args()
    if options.worker:
        return worker(options.worker)
    if not math.isfinite(options.timeout) or options.timeout <= 0 or options.memory_mb < 128:
        parser.error('--timeout deve ser positivo/finito e --memory-mb deve ser pelo menos 128')
    root = options.instancias.resolve()
    catalog = options.catalogo or root / 'CATALOGO_TESTES.csv'
    if not catalog.is_file():
        candidates = sorted(root.rglob('CATALOGO_TESTES.csv'))
        if len(candidates) != 1:
            parser.error('Catálogo não localizado inequivocamente; informe --catalogo')
        catalog = candidates[0]
    with catalog.open(encoding='utf-8-sig', newline='') as handle:
        cases = list(csv.DictReader(handle, delimiter=';'))
    if not cases or not {'id', 'arquivo', 'prioridade', 'categoria', 'expectativa'} <= cases[0].keys():
        parser.error('Catálogo sem as colunas obrigatórias')
    if len({case['id'] for case in cases}) != len(cases):
        parser.error('IDs duplicados no catálogo')
    if options.ids:
        try:
            chosen = {int(value.strip()) for value in options.ids.split(',')}
        except ValueError:
            parser.error('--ids deve conter números separados por vírgula')
        unknown = chosen - {int(case['id']) for case in cases}
        if unknown:
            parser.error(f'IDs não encontrados: {sorted(unknown)}')
        cases = [case for case in cases if int(case['id']) in chosen]
    if options.prioridades:
        priorities = {value.strip().upper() for value in options.prioridades.split(',')}
        if not priorities <= {'P0', 'P1', 'P2'}:
            parser.error('Prioridades aceitas: P0,P1,P2')
        cases = [case for case in cases if case['prioridade'] in priorities]
    cases.sort(key=lambda case: (case['prioridade'], int(case['id'])))
    paths = list(root.rglob('*.xlsx'))
    output = (options.output or root / 'RESULTADO_TESTES.csv').resolve()
    originals = {path.resolve() for path in paths} | {catalog.resolve()}
    if output in originals:
        parser.error('O relatório não pode sobrescrever uma instância nem o catálogo')
    if options.preservar_falhas and options.preservar_falhas.resolve().is_relative_to(root):
        parser.error('Preserve saídas fora da pasta das instâncias para não contaminar futuras buscas')
    start = time.perf_counter()
    rows = []
    with tempfile.TemporaryDirectory(prefix='unisoma_robustez_') as temporary:
        for case in cases:
            relative = case['arquivo'].replace('\\', '/')
            matches = [path for path in paths if path.relative_to(root).as_posix() == relative
                       or path.relative_to(root).as_posix().endswith('/' + relative)]
            source = matches[0] if len(matches) == 1 else None
            row = isolated_case(case, source, options, Path(temporary))
            rows.append(row)
            print(f"{case['id']} {row['status']:<16} {row['tempo_segundos']:>8}s {case['arquivo']}", flush=True)
        output.parent.mkdir(parents=True, exist_ok=True)
        # Stage then atomically replace the single consolidated report.
        with tempfile.NamedTemporaryFile(mode='w', encoding='utf-8-sig', newline='',
                                         suffix='.csv', dir=output.parent, delete=False) as handle:
            staged = Path(handle.name)
            writer = csv.DictWriter(handle, fieldnames=FIELDS, delimiter=';')
            writer.writeheader()
            writer.writerows(rows)
        try:
            os.replace(staged, output)
        finally:
            staged.unlink(missing_ok=True)
    counts = Counter(row['status'] for row in rows)
    print(f'\nTotal: {len(rows)} | Aprovados: {counts["APROVADO"]} | Reprovados: {counts["REPROVADO"]} | '
          f'Exceções: {counts["ERRO_INESPERADO"]} | Timeouts: {counts["TIMEOUT"]} | '
          f'Inconclusivos: {counts["INCONCLUSIVO"]} | Não executados: {counts["NAO_EXECUTADO"]} | '
          f'Tempo: {time.perf_counter() - start:.2f}s')
    failed = [f'{row["id"]} ({row["status"]})' for row in rows if row['status'] not in ('APROVADO', 'NAO_EXECUTADO')]
    print('Testes com problemas: ' + (', '.join(failed) if failed else 'nenhum'))
    print(f'Relatório: {output}')
    return 1 if any(counts[status] for status in ('REPROVADO', 'ERRO_INESPERADO', 'TIMEOUT', 'INCONCLUSIVO')) else 0


if __name__ == '__main__':
    raise SystemExit(main())
