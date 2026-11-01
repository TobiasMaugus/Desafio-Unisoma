"""Read generated PDFs to verify rosters, pagination and adjusted allocations."""
from io import BytesIO
from pathlib import Path
import sys

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from pdf_report import build_distribution_report

PdfReader = pytest.importorskip('pypdf').PdfReader


def report(stats, roster, issues=(), adjusted=False):
    data = build_distribution_report(
        'Relatório resumo da distribuição', stats,
        {row[0]: 80 for row in stats}, issues,
        sum(row[1] for row in stats), adjusted=adjusted, roster=roster)
    return PdfReader(BytesIO(data), strict=True)


def test_summary_and_separate_class_rosters():
    reader = report(
        [('MANHÃ A', 2, 1, 1), ('TARDE A', 1, 1, 0), ('TARDE B', 0, 0, 0)],
        [('Ana (Silva) \\ Souza', 'F', 'Manhã', 'MANHÃ A'),
         ('José', 'M', 'Manhã', 'MANHÃ A'),
         ('Lúcia', 'F', 'Tarde', 'TARDE A')])
    assert len(reader.pages) == 4
    summary, morning, afternoon, empty = [page.extract_text() for page in reader.pages]
    assert 'ALUNOS ALOCADOS' in summary
    assert 'Equilíbrio entre as turmas' in summary
    assert 'Nenhuma inconsistência encontrada.' in summary
    assert 'Ana (Silva) \\ Souza' in morning
    assert 'José' in morning and 'Lúcia' not in morning
    assert 'Lúcia' in afternoon and 'José' not in afternoon
    assert 'Nenhum aluno alocado nesta turma.' in empty
    for index, page in enumerate(reader.pages, 1):
        assert f'Página {index} de 4' in page.extract_text()
        assert 'Solução otimizada pelo modelo' in page.extract_text()
        assert '/Helvetica-Bold' in str(page['/Resources']['/Font']['/F2'].get_object())


def test_long_rosters_repeat_headers_without_losing_or_mixing_students():
    roster = [(f'Pessoa {index:03} de Oliveira e Souza', 'F', 'Manhã', 'MANHÃ A')
              for index in range(1, 76)]
    roster[1] = ('Nome muito longo ' + 'Sobrenome ' * 18, 'F', 'Manhã', 'MANHÃ A')
    reader = report([('MANHÃ A', 75, 75, 0)], roster)
    assert len(reader.pages) > 3
    all_text = '\n'.join(page.extract_text() for page in reader.pages[1:])
    for index in range(1, 76):
        if index != 2:
            assert all_text.count(f'Pessoa {index:03} de Oliveira e Souza') == 1
    assert all_text.count('Sobrenome') == 18
    for page in reader.pages[1:]:
        text = page.extract_text()
        assert 'MANHÃ A' in text
        assert 'Nome do aluno' in text
        assert 'Disponibilidade' in text
        coordinates = [args[5] for args, operator in page.get_contents().operations if operator == b'Tm']
        assert all(20 <= y <= 822 for y in coordinates)


def test_adjusted_report_uses_updated_rooms_and_keeps_unassigned_students():
    reader = report(
        [('MANHÃ A', 0, 0, 0), ('MANHÃ B', 1, 1, 0)],
        [('Aluna transferida', 'F', 'Manhã', 'MANHÃ B'),
         ('Aluno removido', 'M', 'Manhã', None)],
        [('Aviso', '1 aluno sem turma.')], adjusted=True)
    text = [page.extract_text() for page in reader.pages]
    assert 'Aluna transferida' not in text[1]
    assert 'Aluna transferida' in text[2]
    assert 'Alunos sem turma' in text[3] and 'Aluno removido' in text[3]
    assert all('Distribuição ajustada manualmente' in page for page in text)


def test_many_issues_and_classes_paginate_without_losing_messages():
    stats = [(f'MANHÃ {index:02}', 0, 0, 0) for index in range(40)]
    issues = [('Aviso', f'Diagnóstico {index:03}: ' + 'Revise o cadastro. ' * 10)
              for index in range(30)]
    reader = report(stats, [], issues)
    text = '\n'.join(page.extract_text() for page in reader.pages)
    assert all(text.count(f'Diagnóstico {index:03}:') == 1 for index in range(30))
    assert len(reader.pages) > len(stats) + 1


def test_streamlit_exports_original_and_adjusted_rosters(monkeypatch):
    import pdf_report
    from test_interface import launch, synthetic, upload

    exports = []
    original = pdf_report.build_distribution_report

    def capture(*args, **kwargs):
        kwargs['roster'] = list(kwargs['roster'])
        data = original(*args, **kwargs)
        exports.append((kwargs, PdfReader(BytesIO(data))))
        return data

    monkeypatch.setattr(pdf_report, 'build_distribution_report', capture)
    instance = upload(launch(), synthetic())
    next(button for button in instance.button if button.label == 'Gerar ensalamento').click().run()
    assert not instance.exception
    optimized, reader = exports[-1]
    assert not optimized['adjusted']
    assert len(optimized['roster']) == 6
    assert len(reader.pages) == 5
    for name, _, _, room in optimized['roster']:
        class_texts = [page.extract_text() for page in reader.pages[1:]]
        assert any(room in text and name in text for text in class_texts)

    instance.checkbox(key='confirm_remove').check().run()
    removed = instance.selectbox(key='remove_student').value
    instance.button(key='apply_remove').click().run()
    assert not instance.exception
    adjusted, reader = exports[-1]
    assert adjusted['adjusted']
    assert adjusted['roster'][removed][3] is None
    assert optimized['roster'][removed][3] is not None
    assert 'Alunos sem turma' in reader.pages[-1].extract_text()
    assert adjusted['roster'][removed][0] in reader.pages[-1].extract_text()
