"""Regression checks for the real Streamlit flow, without changing processing."""
from hashlib import sha256
from io import BytesIO
import json
from pathlib import Path
import sys

from openpyxl import Workbook, load_workbook
import pytest
import streamlit as st
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
    instance.button[-1].click().run()
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
    labels = [item.label for item in instance.get('download_button')]
    assert 'Baixar planilha atualizada' in labels
    assert len(labels) == (3 if blocked else 5)
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
    instance.button[-1].click().run()
    result = instance.session_state['result'][0]
    instance.text_input[0].set_value('Pessoa A').run()
    assert len(calls) == 1
    assert instance.session_state['result'][0] == result
    tables = [d.value for d in instance.dataframe if list(d.value.columns) == ['Nome','Sexo','Disponibilidade','Turma']]
    assert len(tables[-1]) == 1 and tables[-1].iloc[0]['Nome'] == 'Pessoa A'
    instance.checkbox[0].check().run()
    assert 'result' not in instance.session_state
    instance.button[-1].click().run()
    assert len(calls) == 2 and calls[-1] == {'manha':2, 'tarde':1}
    upload(instance, (PROJECT.parent/'instancias'/'instancia_4.xlsx').read_bytes())
    assert 'result' not in instance.session_state


def test_solver_error_stays_visible(monkeypatch):
    def fail(*args, **kwargs):
        raise RuntimeError('Diagnóstico de teste: solver indisponível')
    monkeypatch.setattr(app, 'make_result', fail)
    instance = upload(launch(), synthetic())
    instance.button[-1].click().run()
    assert not instance.exception
    assert any('solver indisponível' in alert.value for alert in instance.error)
    assert 'result' not in instance.session_state


def test_nonblocking_duplicate_names_are_reported_as_warnings():
    book = load_workbook(BytesIO(synthetic()))
    book['Alunos']['A3'] = book['Alunos']['A2'].value
    buffer = BytesIO()
    book.save(buffer)
    instance = upload(launch(), buffer.getvalue())
    instance.button[-1].click().run()
    assert instance.session_state['result'][1] is False
    assert any('Ensalamento concluído com 1 aviso' in item.value for item in instance.warning)
    assert any('nome duplicado' in issue for issue in instance.session_state['result'][2])


def test_processing_and_manual_are_byte_identical():
    before = json.loads(FIXTURE.read_text(encoding='utf-8'))
    for filename, digest in before['protected_hashes'].items():
        assert sha256((PROJECT/filename).read_bytes()).hexdigest() == digest


def test_corrected_session_data_runs_without_reupload_and_preserves_original():
    book = load_workbook(BytesIO(synthetic()))
    book['Alunos']['B2'] = 'X'
    invalid = BytesIO()
    book.save(invalid)
    original = invalid.getvalue()
    corrected_book = load_workbook(BytesIO(original))
    corrected_book['Alunos']['B2'] = 'F'
    students, resources, issues, blocked = app.inspect_workbook(corrected_book, {'manha': 3, 'tarde': 1})
    assert not blocked
    assert not any('sexo inválido' in issue for issue in issues)
    output, result_blocked, _, stats = app.make_result(corrected_book, {'manha': 3, 'tarde': 1})
    assert not result_blocked and len(stats) == 4
    assert load_workbook(BytesIO(output))['Alunos']['B2'].value == 'F'
    assert load_workbook(BytesIO(original))['Alunos']['B2'].value == 'X'


def test_editor_save_invalidates_old_result_and_reuses_corrected_roster():
    book = load_workbook(BytesIO(synthetic()))
    book['Alunos']['B2'] = 'X'
    buffer = BytesIO()
    book.save(buffer)
    original = buffer.getvalue()
    instance = upload(launch(), original)
    instance.button[-1].click().run()
    assert instance.session_state['result'][1] is True

    editor_key = next(key for key in instance.session_state if key.startswith('abc_roster_'))
    instance.session_state[editor_key] = {
        'edited_rows': {0: {'Sexo': 'F'}}, 'added_rows': [], 'deleted_rows': []}
    instance.run()
    assert any('não salvas' in item.value for item in instance.info)
    instance.button(key='save_roster').click().run()
    assert 'result' not in instance.session_state
    assert instance.session_state['committed_roster'][0][2] == 'F'
    assert 'Baixar planilha atualizada' in [item.label for item in instance.get('download_button')]
    instance.button[-1].click().run()
    assert instance.session_state['result'][1] is False
    result_book = load_workbook(BytesIO(instance.session_state['result'][0]))
    assert result_book['Alunos']['B2'].value == 'F'
    assert load_workbook(BytesIO(original))['Alunos']['B2'].value == 'X'


def test_editor_discard_keeps_confirmed_roster_unchanged():
    instance = upload(launch(), synthetic())
    original_roster = list(instance.session_state['committed_roster'])
    editor_key = next(key for key in instance.session_state if key.startswith('abc_roster_'))
    instance.session_state[editor_key] = {
        'edited_rows': {0: {'Sexo': 'X'}}, 'added_rows': [], 'deleted_rows': []}
    instance.run()
    assert any('não salvas' in item.value for item in instance.info)
    instance.button(key='discard_roster').click().run()
    assert instance.session_state['committed_roster'] == original_roster
    assert not any('não salvas' in item.value for item in instance.info)


def test_manual_swap_move_remove_undo_and_restore(monkeypatch):
    instance = upload(launch(), synthetic())
    instance.button[-1].click().run()
    original = dict(instance.session_state['manual_assignments'])

    instance.selectbox(key='move_student').set_value(2).run()
    instance.selectbox(key='move_room').set_value('MANHÃ - A').run()
    assert any('turma de destino está lotada' in item.value for item in instance.error)
    assert instance.button(key='apply_move').disabled

    instance.selectbox(key='swap_room_b').set_value('TARDE - D').run()
    assert any('turnos diferentes' in item.value for item in instance.error)
    assert instance.button(key='apply_swap').disabled

    instance.selectbox(key='swap_room_b').set_value('MANHÃ - B').run()
    result_table = load_workbook(BytesIO(instance.session_state['result'][0]))['Ensalamento']
    student_sexes = {i: result_table.cell(i + 2, 2).value for i in range(len(original))}
    room_a_member = next(i for i, room in original.items() if room == 'MANHÃ - A' and student_sexes[i] == 'M')
    room_b_member = next(i for i, room in original.items() if room == 'MANHÃ - B' and student_sexes[i] == 'F')
    instance.selectbox(key='swap_student_a').set_value(room_a_member).run()
    instance.selectbox(key='swap_student_b').set_value(room_b_member).run()
    assert any(item.key == 'confirm_swap_gender' for item in instance.checkbox)
    assert instance.button(key='apply_swap').disabled
    instance.checkbox(key='confirm_swap_gender').set_value(True).run()
    before_swap = dict(instance.session_state['manual_assignments'])
    instance.button(key='apply_swap').click().run()
    assert instance.session_state['manual_assignments'] != before_swap

    assigned_to_b = next(i for i, room in instance.session_state['manual_assignments'].items()
                         if room == 'MANHÃ - B')
    instance.selectbox(key='move_student').set_value(assigned_to_b).run()
    instance.selectbox(key='move_room').set_value('MANHÃ - C').run()
    instance.checkbox(key='confirm_move').set_value(True).run()
    instance.button(key='apply_move').click().run()
    assert instance.session_state['manual_assignments'][assigned_to_b] == 'MANHÃ - C'

    instance.selectbox(key='remove_room').set_value('MANHÃ - C').run()
    instance.selectbox(key='remove_student').set_value(assigned_to_b).run()
    instance.checkbox(key='confirm_remove').set_value(True).run()
    instance.button(key='apply_remove').click().run()
    assert instance.session_state['manual_assignments'][assigned_to_b] is None
    assert 'sem turma' in ' '.join(
        item.value for item in instance.error if hasattr(item, 'value')).lower()
    downloads = {}
    original_download = st.download_button
    def capture_download(label, data, *args, **kwargs):
        if label == 'Baixar ensalamento ajustado':
            downloads[label] = (data, kwargs.get('file_name'))
        return original_download(label, data, *args, **kwargs)
    monkeypatch.setattr(st, 'download_button', capture_download)
    instance.run()
    adjusted_bytes, adjusted_name = downloads['Baixar ensalamento ajustado']
    adjusted_book = load_workbook(BytesIO(adjusted_bytes))
    assert 'NAO_CONFORME' in adjusted_name
    assert adjusted_book['Ensalamento'].cell(assigned_to_b + 2, 4).value is None
    assert any('sem turma' in str(row[1]).lower()
               for row in adjusted_book['Inconsistências'].iter_rows(min_row=2, values_only=True))

    instance.button(key='undo_manual').click().run()
    assert instance.session_state['manual_assignments'][assigned_to_b] == 'MANHÃ - C'
    instance.checkbox(key='confirm_restore').set_value(True).run()
    instance.button(key='restore_manual').click().run()
    assert instance.session_state['manual_assignments'] == original


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
