"""Regression checks for the real Streamlit flow, without changing processing."""
from hashlib import sha256
from io import BytesIO
import json
from pathlib import Path
import sys

from openpyxl import Workbook, load_workbook
import pytest
from streamlit.testing.v1 import AppTest

PROJECT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT))
import app

MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
FIXTURE = Path(__file__).parent / 'fixtures' / 'before_refactor.json'


def workbook_digest(content):
    """Hash worksheet values and styles, excluding changing ZIP timestamps."""
    book = load_workbook(BytesIO(content))
    payload = []
    for sheet in book:
        payload.append((sheet.title, sheet.freeze_panes, sheet.auto_filter.ref,
                        [(c.coordinate, c.value, c.number_format, str(c._style))
                         for row in sheet for c in row]))
    book.close()
    return sha256(json.dumps(payload, ensure_ascii=False, default=str).encode()).hexdigest()


def launch():
    instance = AppTest.from_file(str(PROJECT / 'streamlit_app.py'), default_timeout=30).run()
    assert not instance.exception
    return instance


def upload(instance, content, name='entrada.xlsx'):
    instance.file_uploader[0].set_value((name, content, MIME)).run()
    assert not instance.exception
    return instance


def synthetic():
    book = Workbook()
    sheet = book.active
    sheet.title = 'Alunos'
    sheet.append(['Nome', 'Sexo', 'Disponibilidade'])
    for row in [('Pessoa A','F','Manhã'), ('Pessoa B','M','Manhã'),
                ('Pessoa C','F','Manhã'), ('Pessoa D','M','Manhã'),
                ('Pessoa E','F','Tarde'), ('Pessoa F','M','Tarde')]:
        sheet.append(row)
    sheet = book.create_sheet('Turmas')
    for row in [('Recursos','Quantidade'), ('turmas_manha',3),
                ('turmas_tarde',1), ('vagas_manha',2), ('vagas_tarde',2)]:
        sheet.append(row)
    buffer = BytesIO()
    book.save(buffer)
    return buffer.getvalue()


@pytest.mark.parametrize('number', (1,2,3,4))
@pytest.mark.parametrize('allow_empty', (False,True))
def test_official_output_matches_before_refactor(number, allow_empty, monkeypatch):
    path = PROJECT.parent / 'instancias' / f'instancia_{number}.xlsx'
    content = path.read_bytes()
    expected = json.loads(FIXTURE.read_text(encoding='utf-8'))['cases'][f'{number}_{allow_empty}']
    calls = []
    original = app.make_result
    def tracked(book, choices=None):
        calls.append(choices.copy())
        return original(book, choices)
    monkeypatch.setattr(app, 'make_result', tracked)
    instance = upload(launch(), content, path.name)
    instance.checkbox[0].set_value(allow_empty).run()
    assert not calls, 'Uploading and choosing rooms must not run optimization'
    instance.button[0].click().run()
    assert not instance.exception
    result, blocked, issues, stats = instance.session_state['result']
    assert len(calls) == 1
    assert calls[0] == expected['choices']
    assert blocked == expected['blocked']
    assert issues == expected['issues']
    assert [list(row) for row in stats] == expected['stats']
    assert workbook_digest(result) == expected['workbook_digest']
    assert instance.session_state['result_signature'] == (
        sha256(content).hexdigest(), tuple(sorted(expected['choices'].items())))
    instance.run()
    assert len(calls) == 1, 'Ordinary reruns must not optimize again'
    assert instance.session_state['result'][0] == result
    assert len(instance.get('download_button')) == (2 if blocked else 3)
    assert path.read_bytes() == content


def test_empty_and_invalid_upload():
    instance = launch()
    assert len(instance.file_uploader) == 1
    assert len(instance.get('download_button')) == 1  # Manual available before upload.
    assert not instance.button
    upload(instance, b'not an Excel workbook')
    assert instance.error
    assert not instance.button


def test_session_search_room_change_and_file_change(monkeypatch):
    calls = []
    original = app.make_result
    def tracked(book, choices=None):
        calls.append(choices.copy())
        return original(book, choices)
    monkeypatch.setattr(app, 'make_result', tracked)
    instance = upload(launch(), synthetic())
    instance.button[0].click().run()
    result = instance.session_state['result'][0]
    instance.text_input[0].set_value('Pessoa A').run()
    assert len(calls) == 1
    assert instance.session_state['result'][0] == result
    tables = [d.value for d in instance.dataframe if list(d.value.columns) == ['Nome','Sexo','Disponibilidade','Turma']]
    assert len(tables[-1]) == 1 and tables[-1].iloc[0]['Nome'] == 'Pessoa A'
    instance.checkbox[0].check().run()
    assert 'result' not in instance.session_state
    instance.button[0].click().run()
    assert len(calls) == 2 and calls[-1] == {'manha':2, 'tarde':1}
    upload(instance, (PROJECT.parent/'instancias'/'instancia_4.xlsx').read_bytes())
    assert 'result' not in instance.session_state


def test_solver_error_stays_visible(monkeypatch):
    def fail(*args, **kwargs):
        raise RuntimeError('Diagnóstico de teste: solver indisponível')
    monkeypatch.setattr(app, 'make_result', fail)
    instance = upload(launch(), synthetic())
    instance.button[0].click().run()
    assert not instance.exception
    assert any('solver indisponível' in alert.value for alert in instance.error)
    assert 'result' not in instance.session_state


def test_processing_and_manual_are_byte_identical():
    before = json.loads(FIXTURE.read_text(encoding='utf-8'))
    for filename, digest in before['protected_hashes'].items():
        assert sha256((PROJECT/filename).read_bytes()).hexdigest() == digest


if __name__ == '__main__':
    # One-time capture, run BEFORE touching the UI. Contains hashes, not student names.
    assert not FIXTURE.exists(), 'Do not overwrite the pre-refactor reference'
    reference = {'protected_hashes': {}, 'cases': {}}
    for filename in ('app.py','requirements.txt','MANUAL_USUARIO.md','MANUAL_USUARIO.pdf'):
        reference['protected_hashes'][filename] = sha256((PROJECT/filename).read_bytes()).hexdigest()
    for number in (1,2,3,4):
        for allow_empty in (False,True):
            content = (PROJECT.parent/'instancias'/f'instancia_{number}.xlsx').read_bytes()
            instance = upload(launch(), content, f'instancia_{number}.xlsx')
            instance.checkbox[0].set_value(allow_empty).run()
            instance.button[0].click().run()
            assert not instance.exception
            result, blocked, issues, stats = instance.session_state['result']
            reference['cases'][f'{number}_{allow_empty}'] = {
                'choices': dict(instance.session_state['result_signature'][1]),
                'blocked': blocked, 'issues': issues, 'stats': stats,
                'workbook_digest': workbook_digest(result)}
    FIXTURE.parent.mkdir(parents=True, exist_ok=True)
    FIXTURE.write_text(json.dumps(reference, ensure_ascii=False, indent=2), encoding='utf-8')
    print('Captured 8 reference scenarios from the original Streamlit UI.')
