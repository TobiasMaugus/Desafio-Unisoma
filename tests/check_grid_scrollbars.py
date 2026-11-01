"""Run against Streamlit on port 8520 with Playwright and installed Edge."""
from io import BytesIO
from pathlib import Path

from openpyxl import load_workbook
from playwright.sync_api import expect, sync_playwright

from test_interface import synthetic


with sync_playwright() as pw:
    # Playwright hides native scrollbars by default in headless screenshots.
    browser = pw.chromium.launch(channel='msedge', headless=True, ignore_default_args=['--hide-scrollbars'])
    page = browser.new_page(viewport={'width': 1600, 'height': 1000}, accept_downloads=True)
    errors = []
    page.on('pageerror', lambda error: errors.append(str(error)))
    page.goto('http://127.0.0.1:8520')
    book = load_workbook(BytesIO(synthetic()))
    for number in range(70):
        book['Alunos'].append([f'Aluno {number}', 'F' if number % 2 else 'M', 'Manhã'])
    book['Turmas']['B4'] = 40
    content = BytesIO()
    book.save(content)
    page.locator('input[type=file]').set_input_files({
        'name': 'rolagem.xlsx', 'mimeType': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'buffer': content.getvalue()})
    grid = page.frame_locator('iframe[title="st_aggrid.AgGrid.agGrid"]')
    horizontal = grid.locator('.ag-body-horizontal-scroll-viewport')
    vertical = grid.locator('.ag-body-vertical-scroll-viewport')
    for width in (1600, 1366, 390):
        page.set_viewport_size({'width': width, 'height': 1000})
        expect(horizontal).to_be_visible()
        expect(vertical).to_be_visible()
        page.wait_for_timeout(350)
        assert horizontal.evaluate('e => e.scrollWidth > e.clientWidth')
        assert vertical.evaluate('e => e.scrollHeight > e.clientHeight')
        assert horizontal.evaluate('e => e.offsetHeight > e.clientHeight')
        assert vertical.evaluate('e => e.offsetWidth > e.clientWidth')
        horizontal.evaluate('e => { e.scrollLeft = 300; }')
        vertical.evaluate('e => { e.scrollTop = 400; }')
        page.wait_for_timeout(250)
        assert horizontal.evaluate('e => e.scrollLeft') > 0
        assert grid.locator('.ag-center-cols-viewport').evaluate('e => e.scrollLeft') > 0
        assert grid.locator('.ag-body-viewport').evaluate('e => e.scrollTop') > 0
    page.set_viewport_size({'width': 1366, 'height': 1000})
    horizontal.evaluate('e => { e.scrollLeft = 0; }')
    vertical.evaluate('e => { e.scrollTop = 0; }')
    page.wait_for_timeout(300)
    page.locator('iframe[title="st_aggrid.AgGrid.agGrid"]').screenshot(path='tests/artifacts/interface_scrollbars.png')
    page.get_by_role('button', name='Gerar ensalamento', exact=True).click()
    excel = page.get_by_role('button', name='Baixar Excel completo (.xlsx)', exact=True)
    pdf = page.get_by_role('button', name='Baixar resumo da distribuição (PDF)', exact=True)
    expect(excel).to_be_enabled()
    expect(pdf).to_be_enabled()
    assert abs(excel.bounding_box()['y'] - pdf.bounding_box()['y']) < 5
    assert excel.bounding_box()['width'] < page.locator('.st-key-abc_downloads').bounding_box()['width'] / 2
    expect(page.get_by_text('Outros formatos de download', exact=True)).to_have_count(0)
    expect(page.get_by_role('button', name='Baixar resumo por turma (.csv)', exact=True)).to_have_count(0)
    for button, suffix in ((excel, '.xlsx'), (pdf, '.pdf')):
        with page.expect_download() as pending:
            button.click()
        assert pending.value.suggested_filename.endswith(suffix)
        data = Path(pending.value.path()).read_bytes()
        if suffix == '.xlsx':
            result = load_workbook(BytesIO(data))
            assert result['Ensalamento'].max_row == 77
            result.close()
        else:
            assert data.startswith(b'%PDF')
    assert not errors, errors
    expect(page.locator('[data-testid="stException"]')).to_have_count(0)
    browser.close()
    print('PASS: horizontal and vertical scrollbars visible and functional at 1600, 1366 and 390 px; Excel/PDF side by side and downloaded; CSV removed.')
