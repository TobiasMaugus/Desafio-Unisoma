"""Browser regression: run Streamlit on port 8517, then run this with the project .venv."""
from io import BytesIO
from pathlib import Path

from openpyxl import load_workbook
from playwright.sync_api import expect, sync_playwright

from test_interface import synthetic


def main():
    book = load_workbook(BytesIO(synthetic()))
    book['Alunos']['B4'] = 'X'
    book['Alunos']['A5'] = 'Pessoa E'
    buffer = BytesIO()
    book.save(buffer)
    book.close()
    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(channel='msedge', headless=True)
        page = browser.new_page(viewport={'width': 1600, 'height': 1100})
        errors = []
        page.on('pageerror', lambda error: errors.append(str(error)))
        page.goto('http://127.0.0.1:8517')
        page.locator('input[type=file]').set_input_files({
            'name': 'avisos.xlsx',
            'mimeType': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            'buffer': buffer.getvalue(),
        })
        grid = page.frame_locator('iframe[title="st_aggrid.AgGrid.agGrid"]')

        def cell(record, column):
            return grid.locator(f'[row-id="aluno-{record}"] [col-id="{column}"]')

        def edit(record, column, value):
            expect(page.locator('[data-testid="stStatusWidget"]')).not_to_be_visible()
            page.wait_for_timeout(500)
            target = cell(record, column)
            target.click()
            target.locator('input').fill(value)
            target.locator('input').press('Enter')

        expect(cell(4, 'Atenção')).to_contain_text('registro 4:')
        expect(cell(5, 'Atenção')).to_contain_text('registros 5, 6:')
        grid.locator('[row-id="aluno-2"] input[type=checkbox]').check()
        page.get_by_role('button', name='delete', exact=False).click()
        expect(cell(2, 'Nome')).to_have_count(0)
        expect(cell(3, 'Atenção')).to_have_text('—')
        expect(cell(4, 'Atenção')).to_contain_text('registro 4:')
        expect(cell(4, 'Atenção')).to_contain_text('sexo inválido')
        expect(cell(5, 'Atenção')).to_contain_text('registros 5, 6:')

        # Edits must refresh diagnostics before saving, including related duplicate rows.
        edit(4, 'Sexo', 'F')
        expect(cell(4, 'Atenção')).to_have_text('—')
        edit(5, 'Nome', 'Pessoa D')
        expect(cell(5, 'Atenção')).to_have_text('—')
        expect(cell(6, 'Atenção')).to_have_text('—')
        edit(4, 'Sexo', 'feminino')
        expect(cell(4, 'Atenção')).to_contain_text('Correção automática em Sexo')
        edit(4, 'Sexo', 'X')
        expect(cell(4, 'Atenção')).to_contain_text('sexo inválido')
        expect(cell(4, 'Atenção')).not_to_contain_text('Correção automática')

        page.get_by_role('button', name='Salvar alunos', exact=True).click()
        expect(page.get_by_role('button', name='Salvar alunos', exact=True)).to_be_disabled()
        expect(cell(4, 'Atenção')).to_contain_text('registro 4:')
        with page.expect_download() as pending:
            page.get_by_role('button', name='Baixar planilha atualizada', exact=True).click()
        downloaded = load_workbook(BytesIO(Path(pending.value.path()).read_bytes()))
        assert downloaded['Alunos']['A2'].value == 'Pessoa B'
        assert downloaded['Alunos']['B3'].value == 'X'
        downloaded.close()

        page.get_by_text('Histórico e restauração os alunos', exact=True).click()
        page.locator('.st-key-roster_undo_button button').click()
        expect(cell(2, 'Nome')).to_have_text('Pessoa A')
        expect(cell(4, 'Atenção')).to_contain_text('registro 4:')
        expect(cell(5, 'Atenção')).to_contain_text('registros 5, 6:')
        # Removing one duplicate must clear the warning on its remaining counterpart.
        grid.locator('[row-id="aluno-5"] input[type=checkbox]').check()
        page.get_by_role('button', name='delete', exact=False).click()
        expect(cell(5, 'Nome')).to_have_count(0)
        expect(cell(6, 'Atenção')).to_have_text('—')
        page.locator('.st-key-abc_roster_actions').get_by_role(
            'button', name='Descartar alterações', exact=True).click()
        expect(cell(5, 'Atenção')).to_contain_text('registros 5, 6:')
        expect(cell(6, 'Atenção')).to_contain_text('registros 5, 6:')

        # New rows retain editing and validation support.
        grid.locator('.ag-row-pinned [col-id="Atenção"]').click()
        expect(cell(8, 'Nome')).to_be_visible()
        expect(cell(8, 'Atenção')).to_contain_text('registro 8:')
        edit(8, 'Nome', 'Pessoa nova')
        edit(8, 'Sexo', 'F')
        edit(8, 'Disponibilidade', 'Tarde')
        expect(cell(8, 'Atenção')).to_have_text('—')
        page.locator('.st-key-abc_roster_actions').get_by_role(
            'button', name='Descartar alterações', exact=True).click()
        expect(cell(8, 'Nome')).to_have_count(0)
        assert not errors, errors
        expect(page.locator('[data-testid=stException]')).to_have_count(0)
        browser.close()
    print('PASS: deletion, edits, duplicate warnings, corrections, save, download, undo, discard and addition')


if __name__ == '__main__':
    main()
